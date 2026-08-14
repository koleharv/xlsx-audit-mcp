#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  auditWorkbook,
  cellIdToString,
  extractRefs,
  findDependents,
  numToCol,
  parseCellAddress,
} from "./formulas.js";
import { loadWorkbook } from "./xlsx.js";

const server = new McpServer({ name: "xlsx-audit", version: "0.1.0" });

type ToolResult = { content: { type: "text"; text: string }[]; isError?: boolean };

function jsonResult(data: unknown): ToolResult {
  return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
}

function guarded<A extends unknown[]>(fn: (...args: A) => Promise<ToolResult>) {
  return async (...args: A): Promise<ToolResult> => {
    try {
      return await fn(...args);
    } catch (err) {
      return { content: [{ type: "text", text: err instanceof Error ? err.message : String(err) }], isError: true };
    }
  };
}

const pathArg = z.string().describe("Path to an .xlsx workbook");

server.registerTool(
  "workbook_overview",
  {
    title: "Workbook overview",
    description:
      "Structure of an .xlsx workbook: sheets with dimensions and formula counts, defined names, external workbook links, merged cells. The map you need before tracing or auditing.",
    inputSchema: { path: pathArg },
  },
  guarded(async ({ path }) => {
    const wb = await loadWorkbook(path);
    return jsonResult({
      sheets: wb.sheets.map((s) => ({
        name: s.name,
        dimension: s.dimension,
        cells: s.cells.size,
        formulas: s.formulaCells.length,
        mergedRanges: s.merged.length,
      })),
      definedNames: wb.definedNames,
      externalWorkbookLinks: wb.externalLinks,
    });
  }),
);

server.registerTool(
  "list_formulas",
  {
    title: "List formulas",
    description:
      "Formulas with addresses and cached values. Filter with `contains` (e.g. 'INDIRECT', 'VLOOKUP') and/or `sheet`.",
    inputSchema: {
      path: pathArg,
      sheet: z.string().optional().describe("Only this sheet"),
      contains: z.string().optional().describe("Case-insensitive substring of the formula text"),
      limit: z.number().int().min(1).max(200).default(50).describe("Max results"),
    },
  },
  guarded(async ({ path, sheet, contains, limit }) => {
    const wb = await loadWorkbook(path);
    const needle = contains?.toLowerCase();
    const results: { cell: string; formula: string; value?: string }[] = [];
    let total = 0;
    for (const s of wb.sheets) {
      if (sheet && s.name.toLowerCase() !== sheet.toLowerCase()) continue;
      for (const cell of s.formulaCells) {
        if (needle && !cell.formula!.toLowerCase().includes(needle)) continue;
        total++;
        if (results.length < limit) {
          results.push({ cell: `${s.name}!${cell.ref}`, formula: cell.formula!, value: cell.value });
        }
      }
    }
    return jsonResult({ totalMatches: total, showing: results.length, results });
  }),
);

server.registerTool(
  "trace_cell",
  {
    title: "Trace a cell",
    description:
      "Audit one cell: its formula and cached value, precedents (cells and ranges it reads), and dependents (formulas that read it, directly or via a range). Address like 'Summary!B14'; unqualified addresses use the first sheet.",
    inputSchema: {
      path: pathArg,
      cell: z.string().describe("Cell address, e.g. B14 or 'Summary!B14'"),
      max_dependents: z.number().int().min(1).max(200).default(30),
    },
  },
  guarded(async ({ path, cell, max_dependents }) => {
    const wb = await loadWorkbook(path);
    const id = parseCellAddress(cell, wb.sheets);
    const sheet = wb.sheets.find((s) => s.name === id.sheet)!;
    const ref = `${numToCol(id.col)}${id.row}`;
    const target = sheet.cells.get(ref);
    const precedents = target?.formula
      ? extractRefs(target.formula).map((t) => ({ ref: t.text, sheet: t.sheet ?? id.sheet }))
      : [];
    return jsonResult({
      cell: cellIdToString(id),
      formula: target?.formula,
      value: target?.value,
      precedents,
      dependents: findDependents(wb, id, max_dependents),
    });
  }),
);

server.registerTool(
  "audit_workbook",
  {
    title: "Audit the workbook",
    description:
      "Risk report across the whole model: circular references (with example chains), volatile functions (INDIRECT/OFFSET/NOW/RAND...), hardcoded constants inside formulas, external workbook links, merged cells, very long formulas.",
    inputSchema: { path: pathArg },
  },
  guarded(async ({ path }) => jsonResult(auditWorkbook(await loadWorkbook(path)))),
);

await server.connect(new StdioServerTransport());
