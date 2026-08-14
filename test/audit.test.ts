import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import {
  auditWorkbook,
  colToNum,
  extractRefs,
  findCycles,
  findDependents,
  numToCol,
  parseCellAddress,
  translateFormula,
} from "../src/formulas.js";
import { loadWorkbook, type Workbook } from "../src/xlsx.js";
import { buildTestXlsx } from "./fixtures.js";

let dir: string;
let wb: Workbook;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "xlsx-audit-test-"));
  const path = join(dir, "model.xlsx");
  await writeFile(path, buildTestXlsx());
  wb = await loadWorkbook(path);
});

afterAll(() => rm(dir, { recursive: true, force: true }));

describe("column math", () => {
  it("round-trips column letters", () => {
    expect(colToNum("A")).toBe(1);
    expect(colToNum("AA")).toBe(27);
    expect(numToCol(27)).toBe("AA");
    expect(numToCol(colToNum("XFD"))).toBe("XFD");
  });
});

describe("extractRefs", () => {
  it("extracts plain, absolute, and range refs", () => {
    expect(extractRefs("=A1+$B$2").map((t) => t.text)).toEqual(["A1", "$B$2"]);
    const range = extractRefs("=SUM(A1:B10)")[0];
    expect(range).toMatchObject({ startCol: 1, startRow: 1, endCol: 2, endRow: 10, isRange: true });
  });

  it("extracts sheet-qualified refs, including quoted names", () => {
    expect(extractRefs("=Inputs!B1*2")[0].sheet).toBe("Inputs");
    expect(extractRefs("='My Data'!A1+5")[0].sheet).toBe("My Data");
  });

  it("does not mistake function names for cell refs", () => {
    expect(extractRefs("=LOG10(5)")).toEqual([]);
    expect(extractRefs("=ATAN2(1,2)")).toEqual([]);
  });

  it("ignores ref-looking text inside string literals", () => {
    expect(extractRefs('=INDIRECT("A1")')).toEqual([]);
    expect(extractRefs('=IF(B2>0,"see C3",D4)').map((t) => t.text)).toEqual(["B2", "D4"]);
  });
});

describe("translateFormula", () => {
  it("shifts relative refs and preserves absolutes", () => {
    expect(translateFormula("A1*$B$1", 1, 0)).toBe("A2*$B$1");
    expect(translateFormula("SUM(A1:A10)", 2, 1)).toBe("SUM(B3:B12)");
    expect(translateFormula("$A1+B$2", 3, 3)).toBe("$A4+E$2");
  });

  it("leaves string literals untouched", () => {
    expect(translateFormula('IF(A1>0,"A1 is fine",B1)', 1, 0)).toBe('IF(A2>0,"A1 is fine",B2)');
  });
});

describe("loadWorkbook", () => {
  it("loads sheets, values, shared strings, defined names, external links", () => {
    expect(wb.sheets.map((s) => s.name)).toEqual(["Inputs", "Model", "My Data"]);
    expect(wb.sheets[0].cells.get("A1")?.value).toBe("Tax rate");
    expect(wb.sheets[0].cells.get("B1")?.value).toBe("0.13");
    expect(wb.definedNames).toEqual([{ name: "TaxRate", ref: "Inputs!$B$1" }]);
    expect(wb.externalLinks).toEqual(["file:///C:/old/model2019.xlsx"]);
    expect(wb.sheets[1].merged).toEqual(["A5:B6"]);
  });

  it("materializes shared formulas with translated refs", () => {
    const model = wb.sheets[1];
    expect(model.cells.get("F1")?.formula).toBe("A1*2");
    expect(model.cells.get("F2")?.formula).toBe("A2*2");
    expect(model.cells.get("F3")?.formula).toBe("A3*2");
  });
});

describe("dependency queries", () => {
  it("finds cross-sheet dependents", () => {
    const target = parseCellAddress("Inputs!B1", wb.sheets);
    const deps = findDependents(wb, target, 10);
    expect(deps.map((d) => d.cell)).toEqual(["Model!A1"]);
  });

  it("finds dependents via range containment", () => {
    const target = parseCellAddress("Model!A2", wb.sheets);
    const deps = findDependents(wb, target, 10);
    const cells = deps.map((d) => d.cell).sort();
    expect(cells).toEqual(["Model!A3", "Model!F2"]);
    expect(deps.find((d) => d.cell === "Model!A3")?.viaRange).toBe("A1:A2");
  });

  it("parses quoted sheet addresses and rejects garbage", () => {
    expect(parseCellAddress("'My Data'!A1", wb.sheets).sheet).toBe("My Data");
    expect(() => parseCellAddress("NotASheet!A1", wb.sheets)).toThrow(/No sheet named/);
    expect(() => parseCellAddress("!!!", wb.sheets)).toThrow(/Not a cell address/);
  });
});

describe("findCycles", () => {
  it("detects the circular pair exactly once", () => {
    const { cycles } = findCycles(wb);
    expect(cycles).toHaveLength(1);
    const flat = cycles[0].join("→");
    expect(flat).toContain("Model!D1");
    expect(flat).toContain("Model!D2");
  });
});

describe("auditWorkbook", () => {
  it("produces the full risk report", () => {
    const report = auditWorkbook(wb);
    expect(report.circularReferences.count).toBe(1);
    expect(report.volatileFunctions.byFunction).toEqual({ INDIRECT: 1 });
    // 100 (LOG10 arg), 5, 1.05, and the *2 in each of the three shared-formula cells; 0/1 excluded:
    const constants = report.hardcodedConstants.examples.map((e) => e.detail);
    expect(report.hardcodedConstants.count).toBe(6);
    expect(constants.some((c) => c.startsWith("1.05"))).toBe(true);
    expect(constants.some((c) => c.startsWith("100"))).toBe(true);
    expect(constants.some((c) => c.startsWith("5"))).toBe(true);
    expect(report.externalWorkbookLinks).toEqual(["file:///C:/old/model2019.xlsx"]);
    expect(report.mergedCells).toBe(1);
  });
});
