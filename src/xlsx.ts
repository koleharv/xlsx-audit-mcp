// Workbook loader: .xlsx (a zip of XML parts) → an in-memory model of sheets,
// cells-with-formulas, defined names, merged ranges, and external links.
// Whole file is held in memory — spreadsheets are tens of MB at worst.

import { readFile, stat } from "node:fs/promises";
import { resolve } from "node:path";
import { XMLParser } from "fast-xml-parser";
import { unzipSync, strFromU8 } from "fflate";
import { translateFormula } from "./formulas.js";

export interface Cell {
  ref: string;
  formula?: string;
  value?: string;
}

export interface Sheet {
  name: string;
  cells: Map<string, Cell>;
  formulaCells: Cell[];
  merged: string[];
  dimension?: string;
}

export interface Workbook {
  path: string;
  sheets: Sheet[];
  definedNames: { name: string; ref: string }[];
  externalLinks: string[];
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  parseTagValue: false,
  parseAttributeValue: false,
  isArray: (name) => ["sheet", "row", "c", "definedName", "mergeCell", "si", "Relationship"].includes(name),
});

function asArray<T>(v: T | T[] | undefined): T[] {
  return v === undefined ? [] : Array.isArray(v) ? v : [v];
}

// Shared-string table: cell values with t="s" are indexes into this.
function parseSharedStrings(xml: string | undefined): string[] {
  if (!xml) return [];
  const doc = parser.parse(xml);
  return asArray(doc.sst?.si).map((si: any) => {
    if (typeof si.t === "string") return si.t;
    if (si.t?.["#text"] !== undefined) return String(si.t["#text"]);
    // rich text runs: <si><r><t>a</t></r><r><t>b</t></r></si>
    const runs = asArray(si.r).map((r: any) => (typeof r.t === "string" ? r.t : (r.t?.["#text"] ?? "")));
    return runs.join("");
  });
}

function cellValue(c: any, shared: string[]): string | undefined {
  const v = c.v;
  const raw = v === undefined ? undefined : typeof v === "object" ? v["#text"] : String(v);
  if (raw === undefined) return undefined;
  if (c["@_t"] === "s") return shared[Number(raw)] ?? raw;
  if (c["@_t"] === "b") return raw === "1" ? "TRUE" : "FALSE";
  return raw;
}

function parseRef(ref: string): { col: number; row: number } {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  if (!m) throw new Error(`Bad cell ref: ${ref}`);
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
}

function parseSheet(name: string, xml: string, shared: string[]): Sheet {
  const doc = parser.parse(xml);
  const ws = doc.worksheet ?? {};
  const cells = new Map<string, Cell>();
  const formulaCells: Cell[] = [];
  // Shared formulas: the master <f t="shared" ref si> carries the text; the
  // other cells carry only si and get the master translated by their offset.
  const sharedFormulas = new Map<string, { text: string; anchor: { col: number; row: number } }>();

  for (const row of asArray(ws.sheetData?.row)) {
    for (const c of asArray((row as any).c)) {
      const ref = c["@_r"];
      if (!ref) continue;
      let formula: string | undefined;
      const f = c.f;
      if (f !== undefined) {
        const text = typeof f === "object" ? (f["#text"] !== undefined ? String(f["#text"]) : "") : String(f);
        const si = typeof f === "object" ? f["@_si"] : undefined;
        if (si !== undefined && text) {
          sharedFormulas.set(String(si), { text, anchor: parseRef(ref) });
          formula = text;
        } else if (si !== undefined) {
          const master = sharedFormulas.get(String(si));
          if (master) {
            const here = parseRef(ref);
            formula = translateFormula(master.text, here.row - master.anchor.row, here.col - master.anchor.col);
          }
        } else {
          formula = text || undefined;
        }
      }
      const cell: Cell = { ref, ...(formula ? { formula } : {}), value: cellValue(c, shared) };
      cells.set(ref, cell);
      if (formula) formulaCells.push(cell);
    }
  }

  const merged = asArray(ws.mergeCells?.mergeCell).map((m: any) => m["@_ref"]).filter(Boolean);
  return { name, cells, formulaCells, merged, dimension: ws.dimension?.["@_ref"] };
}

function loadFromBuffer(path: string, buf: Buffer): Workbook {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(new Uint8Array(buf));
  } catch {
    throw new Error(`Not an .xlsx file (failed to read as a zip archive): ${path}`);
  }
  const text = (name: string): string | undefined => (files[name] ? strFromU8(files[name]) : undefined);

  const workbookXml = text("xl/workbook.xml");
  if (!workbookXml) throw new Error(`Not an .xlsx workbook (xl/workbook.xml missing): ${path}`);
  const wbDoc = parser.parse(workbookXml);

  // Relationship id → part path, for locating each sheet's XML.
  const rels = new Map<string, string>();
  for (const rel of asArray(parser.parse(text("xl/_rels/workbook.xml.rels") ?? "").Relationships?.Relationship)) {
    rels.set((rel as any)["@_Id"], (rel as any)["@_Target"]);
  }

  const shared = parseSharedStrings(text("xl/sharedStrings.xml"));

  const sheets: Sheet[] = [];
  for (const s of asArray(wbDoc.workbook?.sheets?.sheet)) {
    const target = rels.get((s as any)["@_r:id"]) ?? `worksheets/sheet${(s as any)["@_sheetId"]}.xml`;
    const partPath = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
    const xml = text(partPath);
    if (xml) sheets.push(parseSheet((s as any)["@_name"], xml, shared));
  }

  const definedNames = asArray(wbDoc.workbook?.definedNames?.definedName).map((d: any) => ({
    name: d["@_name"],
    ref: typeof d === "object" && d["#text"] !== undefined ? String(d["#text"]) : "",
  }));

  const externalLinks: string[] = [];
  for (const [name, data] of Object.entries(files)) {
    if (/^xl\/externalLinks\/_rels\/.*\.rels$/.test(name)) {
      for (const rel of asArray(parser.parse(strFromU8(data)).Relationships?.Relationship)) {
        const target = (rel as any)["@_Target"];
        if (target) externalLinks.push(decodeURIComponent(target));
      }
    }
  }

  return { path, sheets, definedNames, externalLinks };
}

const cache = new Map<string, { size: number; mtimeMs: number; workbook: Workbook }>();

export async function loadWorkbook(rawPath: string): Promise<Workbook> {
  const path = resolve(rawPath);
  const s = await stat(path).catch(() => {
    throw new Error(`Path not found: ${path}`);
  });
  const hit = cache.get(path);
  if (hit && hit.size === s.size && hit.mtimeMs === s.mtimeMs) return hit.workbook;
  const workbook = loadFromBuffer(path, await readFile(path));
  cache.set(path, { size: s.size, mtimeMs: s.mtimeMs, workbook });
  return workbook;
}
