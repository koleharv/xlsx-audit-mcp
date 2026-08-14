// Self-contained smoke test: builds a synthetic workbook in a temp dir, then
// drives the built server over stdio.
// Usage: npm run build && node scripts/smoke.mjs
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { strToU8, zipSync } from "fflate";

const sheet = (body) =>
  `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;

const xlsx = zipSync({
  "xl/workbook.xml": strToU8(
    `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Inputs" sheetId="1" r:id="rId1"/><sheet name="Model" sheetId="2" r:id="rId2"/></sheets></workbook>`,
  ),
  "xl/_rels/workbook.xml.rels": strToU8(
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://x/ws" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://x/ws" Target="worksheets/sheet2.xml"/></Relationships>`,
  ),
  "xl/worksheets/sheet1.xml": strToU8(sheet(`<row r="1"><c r="B1"><v>0.13</v></c></row>`)),
  "xl/worksheets/sheet2.xml": strToU8(
    sheet(
      `<row r="1"><c r="A1"><f>Inputs!B1*1200</f><v>156</v></c><c r="B1"><f>INDIRECT("A1")</f><v>156</v></c><c r="C1"><f>C2+1</f><v>0</v></c></row><row r="2"><c r="A2"><f>SUM(A1:A1)</f><v>156</v></c><c r="C2"><f>C1+1</f><v>1</v></c></row>`,
    ),
  ),
});

const dir = await mkdtemp(join(tmpdir(), "xlsx-audit-smoke-"));
const path = join(dir, "model.xlsx");
await writeFile(path, xlsx);

const client = new Client({ name: "smoke", version: "0.0.0" });
await client.connect(new StdioClientTransport({ command: process.execPath, args: ["dist/index.js"] }));

const { tools } = await client.listTools();
console.log("tools:", tools.map((t) => t.name).join(", "));

const calls = [
  ["workbook_overview", { path }],
  ["list_formulas", { path, contains: "INDIRECT" }],
  ["trace_cell", { path, cell: "Inputs!B1" }],
  ["audit_workbook", { path }],
  ["trace_cell", { path, cell: "Nope!Z9" }], // must be isError
];

for (const [name, args] of calls) {
  const res = await client.callTool({ name, arguments: args });
  console.log(`\n== ${name} ${res.isError ? "(isError)" : ""}\n${res.content[0].text.slice(0, 600)}`);
}

await client.close();
await rm(dir, { recursive: true, force: true });
console.log("\nSMOKE OK");
