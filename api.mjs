import { getStore } from "@netlify/blobs";
import ExcelJS from "exceljs";

export const config = { path: "/api" };

const store = () => getStore("coders-cup");
const key = (d) => `${d}_nov_competitions.xlsx`;
const str = (v) => String(v ?? "").trim();
const json = (o, s = 200) => Response.json(o, { status: s });
const validDay = (d) => Number.isInteger(d) && d >= 12 && d <= 25;

async function load(d) {
  const wb = new ExcelJS.Workbook();
  const buf = await store().get(key(d), { type: "arrayBuffer" });
  if (buf) await wb.xlsx.load(buf);
  return wb;
}
async function save(d, wb) {
  const b = await wb.xlsx.writeBuffer();
  await store().set(key(d), b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength));
}
function read(wb) {
  return wb.worksheets.map((ws) => {
    const teams = [];
    for (let r = 6; r <= ws.rowCount; r++) {
      const team = str(ws.getCell(r, 1).value);
      if (!team) continue;
      teams.push({
        team,
        members: [2, 3, 4].map((c) => str(ws.getCell(r, c).value)).filter(Boolean),
        at: str(ws.getCell(r, 5).value),
      });
    }
    return { id: ws.name, name: str(ws.getCell("B1").value), venue: str(ws.getCell("B2").value), time: str(ws.getCell("B3").value), teams };
  });
}
function sheetName(wb, name) {
  const base = name.replace(/[\\/*?:\[\]]/g, " ").trim().slice(0, 28) || "Competition";
  let n = base, i = 2;
  while (wb.getWorksheet(n)) n = `${base} ${i++}`;
  return n;
}

export default async (req) => {
  try {
    const url = new URL(req.url);
    if (req.method === "GET") {
      const day = Number(url.searchParams.get("day"));
      if (!validDay(day)) return json({ error: "Invalid day" }, 400);
      const wb = await load(day);
      if (url.searchParams.get("download")) {
        if (!wb.worksheets.length) return json({ error: "No data for this day yet" }, 404);
        const buf = await wb.xlsx.writeBuffer();
        return new Response(buf, { headers: {
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Content-Disposition": `attachment; filename="${key(day)}"` } });
      }
      return json({ competitions: read(wb) });
    }

    const b = await req.json();
    const day = Number(b.day);
    if (!validDay(day)) return json({ error: "Invalid day" }, 400);
    const wb = await load(day);

    if (b.action === "addComp") {
      const name = str(b.name), venue = str(b.venue), time = str(b.time);
      if (!name || !venue || !time) return json({ error: "Name, venue and time are required" }, 400);
      const ws = wb.addWorksheet(sheetName(wb, name));
      ws.addRows([["Competition", name], ["Venue", venue], ["Time slot", time], [], ["Team", "Member 1", "Member 2", "Member 3", "Marked present at"]]);
      ws.getRow(5).font = { bold: true };
      ws.getColumn(1).width = 26; [2, 3, 4].forEach((c) => (ws.getColumn(c).width = 24)); ws.getColumn(5).width = 24;
      [1, 2, 3].forEach((r) => (ws.getCell(r, 1).font = { bold: true }));
    } else if (b.action === "addTeam") {
      const ws = wb.getWorksheet(str(b.comp));
      const team = str(b.team);
      const members = (Array.isArray(b.members) ? b.members : []).map(str).filter(Boolean).slice(0, 3);
      if (!ws) return json({ error: "Competition not found" }, 404);
      if (!team || !members.length) return json({ error: "Team name and at least 1 member required" }, 400);
      if (read(wb).find((c) => c.id === ws.name).teams.some((t) => t.team.toLowerCase() === team.toLowerCase()))
        return json({ error: "This team is already added" }, 409);
      const at = new Date().toLocaleString("en-GB", { timeZone: "Asia/Karachi" });
      ws.addRow([team, members[0] || "", members[1] || "", members[2] || "", at]);
    } else return json({ error: "Unknown action" }, 400);

    await save(day, wb);
    return json({ competitions: read(wb) });
  } catch (e) {
    return json({ error: e.message || "Server error" }, 500);
  }
};
