// Builds a synthetic-but-valid .xlsx in memory (fflate zip + hand-written XML).
// The workbook deliberately contains every case the analyzer must handle:
// cross-sheet refs, a quoted sheet name, a LOG10 collision, INDIRECT with a
// ref-looking string, a circular pair, shared formulas, hardcoded constants,
// merged cells, a defined name, a shared string, and an external link.

import { strToU8, zipSync } from "fflate";

const WORKBOOK = `<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Inputs" sheetId="1" r:id="rId1"/>
    <sheet name="Model" sheetId="2" r:id="rId2"/>
    <sheet name="My Data" sheetId="3" r:id="rId3"/>
  </sheets>
  <definedNames>
    <definedName name="TaxRate">Inputs!$B$1</definedName>
  </definedNames>
</workbook>`;

const WORKBOOK_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://x/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://x/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://x/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`;

const SHARED_STRINGS = `<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Tax rate</t></si>
</sst>`;

const SHEET_INPUTS = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>0.13</v></c></row>
    <row r="2"><c r="B2"><v>100</v></c></row>
  </sheetData>
</worksheet>`;

const SHEET_MODEL = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:F6"/>
  <sheetData>
    <row r="1">
      <c r="A1"><f>Inputs!B1*Inputs!B2</f><v>13</v></c>
      <c r="B1"><f>LOG10(100)</f><v>2</v></c>
      <c r="C1"><f>INDIRECT("A"&amp;1)</f><v>13</v></c>
      <c r="D1"><f>D2+1</f><v>0</v></c>
      <c r="E1"><f>'My Data'!A1+5</f><v>12</v></c>
      <c r="F1"><f t="shared" ref="F1:F3" si="0">A1*2</f><v>26</v></c>
    </row>
    <row r="2">
      <c r="A2"><f>A1*1.05</f><v>13.65</v></c>
      <c r="D2"><f>D1+1</f><v>1</v></c>
      <c r="F2"><f t="shared" si="0"/><v>27.3</v></c>
    </row>
    <row r="3">
      <c r="A3"><f>SUM(A1:A2)</f><v>26.65</v></c>
      <c r="F3"><f t="shared" si="0"/><v>53.3</v></c>
    </row>
  </sheetData>
  <mergeCells count="1"><mergeCell ref="A5:B6"/></mergeCells>
</worksheet>`;

const SHEET_MYDATA = `<?xml version="1.0"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>7</v></c></row>
  </sheetData>
</worksheet>`;

const EXTERNAL_LINK_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://x/externalLinkPath" Target="file:///C:/old/model2019.xlsx"/>
</Relationships>`;

export function buildTestXlsx(): Uint8Array {
  return zipSync({
    "xl/workbook.xml": strToU8(WORKBOOK),
    "xl/_rels/workbook.xml.rels": strToU8(WORKBOOK_RELS),
    "xl/sharedStrings.xml": strToU8(SHARED_STRINGS),
    "xl/worksheets/sheet1.xml": strToU8(SHEET_INPUTS),
    "xl/worksheets/sheet2.xml": strToU8(SHEET_MODEL),
    "xl/worksheets/sheet3.xml": strToU8(SHEET_MYDATA),
    "xl/externalLinks/externalLink1.xml": strToU8("<externalLink/>"),
    "xl/externalLinks/_rels/externalLink1.xml.rels": strToU8(EXTERNAL_LINK_RELS),
  });
}
