// Formula analysis: reference extraction, shared-formula translation,
// dependency queries, cycle detection, and audit smells. Pure logic.

import type { Sheet, Workbook } from "./xlsx.js";

export function colToNum(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n: number): string {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export interface RefToken {
  /** Sheet name if the ref was qualified; undefined = same sheet. */
  sheet?: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
  isRange: boolean;
  text: string;
}

// Sheet prefix: 'Quoted Name'! or PlainName! — then A1 or A1:B2.
// Guards: not preceded by identifier chars (the G10 in LOG10 is not a cell),
// not followed by "(" (that would make it a function call like LOG10(...)).
const REF_RE =
  /(?<![A-Za-z0-9_.$!])(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!)?(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?(?![\w(])/g;

/** Replace string literals ("..." with "" escapes) by spaces so their contents are never parsed. */
export function stripStrings(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/g, (m) => " ".repeat(m.length));
}

function parseEndpoint(text: string): { col: number; row: number; colAbs: boolean; rowAbs: boolean } {
  const m = text.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/)!;
  return { col: colToNum(m[2]), row: Number(m[4]), colAbs: m[1] === "$", rowAbs: m[3] === "$" };
}

export function extractRefs(formula: string): RefToken[] {
  const clean = stripStrings(formula);
  const tokens: RefToken[] = [];
  for (const m of clean.matchAll(REF_RE)) {
    const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
    const a = parseEndpoint(m[3]);
    const b = m[4] !== undefined ? parseEndpoint(m[4]) : a;
    tokens.push({
      sheet,
      startCol: Math.min(a.col, b.col),
      startRow: Math.min(a.row, b.row),
      endCol: Math.max(a.col, b.col),
      endRow: Math.max(a.row, b.row),
      isRange: m[4] !== undefined,
      text: m[0],
    });
  }
  return tokens;
}

/** Shift relative refs by (dr, dc) — how Excel materializes shared/filled formulas. */
export function translateFormula(formula: string, dr: number, dc: number): string {
  // Strings must survive untouched, so operate on the original but skip
  // matches that fall inside string literals (found via the stripped copy).
  const clean = stripStrings(formula);
  let out = "";
  let last = 0;
  for (const m of clean.matchAll(REF_RE)) {
    const idx = m.index!;
    out += formula.slice(last, idx);
    const prefixLen = m[0].length - m[3].length - (m[4] !== undefined ? m[4].length + 1 : 0);
    out += formula.slice(idx, idx + prefixLen);
    const shift = (endpoint: string): string => {
      const e = endpoint.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/)!;
      const col = e[1] === "$" ? colToNum(e[2]) : colToNum(e[2]) + dc;
      const row = e[3] === "$" ? Number(e[4]) : Number(e[4]) + dr;
      return `${e[1]}${numToCol(Math.max(1, col))}${e[3]}${Math.max(1, row)}`;
    };
    out += shift(m[3]);
    if (m[4] !== undefined) out += `:${shift(m[4])}`;
    last = idx + m[0].length;
  }
  return out + formula.slice(last);
}

// ---------------------------------------------------------------------------
// Addresses and graph queries

export interface CellId {
  sheet: string;
  col: number;
  row: number;
}

export function parseCellAddress(text: string, sheets: Sheet[]): CellId {
  const m = text.trim().match(/^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})$/);
  if (!m) throw new Error(`Not a cell address: "${text}" (expected e.g. B14 or Summary!B14)`);
  const sheetName = m[1]?.replace(/''/g, "'") ?? m[2];
  const sheet = sheetName
    ? sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase())
    : sheets[0];
  if (!sheet) throw new Error(`No sheet named "${sheetName}". Sheets: ${sheets.map((s) => s.name).join(", ")}`);
  return { sheet: sheet.name, col: colToNum(m[3]), row: Number(m[4]) };
}

export function cellIdToString(id: CellId): string {
  return `${id.sheet}!${numToCol(id.col)}${id.row}`;
}

function tokenSheet(token: RefToken, defaultSheet: string): string {
  return token.sheet ?? defaultSheet;
}

function tokenContains(token: RefToken, defaultSheet: string, target: CellId): boolean {
  return (
    tokenSheet(token, defaultSheet).toLowerCase() === target.sheet.toLowerCase() &&
    target.col >= token.startCol &&
    target.col <= token.endCol &&
    target.row >= token.startRow &&
    target.row <= token.endRow
  );
}

export interface Dependent {
  cell: string;
  formula: string;
  viaRange?: string;
}

export function findDependents(workbook: Workbook, target: CellId, limit: number): Dependent[] {
  const out: Dependent[] = [];
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.formulaCells) {
      for (const token of extractRefs(cell.formula!)) {
        if (tokenContains(token, sheet.name, target)) {
          out.push({
            cell: `${sheet.name}!${cell.ref}`,
            formula: cell.formula!,
            ...(token.isRange ? { viaRange: token.text } : {}),
          });
          break;
        }
      }
      if (out.length >= limit) return out;
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Cycle detection: three-color DFS over formula cells. Range references
// contribute edges only up to an area cap, so SUM(A:A) can't explode the
// graph — capped ranges are counted and reported instead of silently dropped.

const RANGE_AREA_CAP = 10_000;

export interface CycleReport {
  cycles: string[][];
  cappedRanges: number;
}

export function findCycles(workbook: Workbook, maxCycles = 10): CycleReport {
  const formulaMap = new Map<string, { formula: string; sheet: string; ref: string }>();
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.formulaCells) {
      formulaMap.set(`${sheet.name.toLowerCase()}!${cell.ref}`, { formula: cell.formula!, sheet: sheet.name, ref: cell.ref });
    }
  }

  let cappedRanges = 0;
  const edges = (key: string): string[] => {
    const node = formulaMap.get(key)!;
    const out: string[] = [];
    for (const token of extractRefs(node.formula)) {
      const sheetName = (token.sheet ?? node.sheet).toLowerCase();
      const area = (token.endCol - token.startCol + 1) * (token.endRow - token.startRow + 1);
      if (area > RANGE_AREA_CAP) {
        cappedRanges++;
        continue;
      }
      for (let r = token.startRow; r <= token.endRow; r++) {
        for (let c = token.startCol; c <= token.endCol; c++) {
          const k = `${sheetName}!${numToCol(c)}${r}`;
          if (formulaMap.has(k)) out.push(k);
        }
      }
    }
    return out;
  };

  const color = new Map<string, 1 | 2>(); // 1 = in progress, 2 = done
  const cycles: string[][] = [];

  for (const start of formulaMap.keys()) {
    if (color.get(start) === 2 || cycles.length >= maxCycles) continue;
    const stack: { key: string; edges: string[]; i: number }[] = [{ key: start, edges: edges(start), i: 0 }];
    color.set(start, 1);
    while (stack.length) {
      const top = stack[stack.length - 1];
      if (top.i >= top.edges.length || cycles.length >= maxCycles) {
        color.set(top.key, 2);
        stack.pop();
        continue;
      }
      const next = top.edges[top.i++];
      if (color.get(next) === 1) {
        const at = stack.findIndex((f) => f.key === next);
        cycles.push([...stack.slice(at).map((f) => formulaMap.get(f.key)!.sheet + "!" + formulaMap.get(f.key)!.ref), formulaMap.get(next)!.sheet + "!" + formulaMap.get(next)!.ref]);
      } else if (!color.has(next)) {
        color.set(next, 1);
        stack.push({ key: next, edges: edges(next), i: 0 });
      }
    }
  }
  return { cycles, cappedRanges };
}

// ---------------------------------------------------------------------------
// Smells

const VOLATILE_RE = /\b(NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY|OFFSET|INDIRECT|CELL|INFO)\s*\(/gi;

export interface SmellHit {
  cell: string;
  detail: string;
}

export interface AuditReport {
  circularReferences: { count: number; examples: string[][]; cappedRanges: number };
  volatileFunctions: { count: number; byFunction: Record<string, number>; examples: SmellHit[] };
  hardcodedConstants: { count: number; examples: SmellHit[] };
  externalWorkbookLinks: string[];
  structuredTableRefs: number;
  mergedCells: number;
  longFormulas: { count: number; examples: SmellHit[] };
}

export function auditWorkbook(workbook: Workbook): AuditReport {
  const volatile: SmellHit[] = [];
  const byFunction: Record<string, number> = {};
  const constants: SmellHit[] = [];
  const long: SmellHit[] = [];
  let structuredTableRefs = 0;
  let mergedCells = 0;

  for (const sheet of workbook.sheets) {
    mergedCells += sheet.merged.length;
    for (const cell of sheet.formulaCells) {
      const addr = `${sheet.name}!${cell.ref}`;
      const clean = stripStrings(cell.formula!);

      for (const m of clean.matchAll(VOLATILE_RE)) {
        const fn = m[1].toUpperCase();
        byFunction[fn] = (byFunction[fn] ?? 0) + 1;
        if (volatile.length < 20) volatile.push({ cell: addr, detail: fn });
      }

      structuredTableRefs += [...clean.matchAll(/\[[^\]]*\]/g)].length;

      // Hardcoded constants: numeric literals left over once refs are removed.
      // ponytail: heuristic — 0 and 1 are ignored as structural (IF(x,1,0)),
      // and scientific-notation tails can slip through; good enough to flag models.
      const noRefs = clean.replace(REF_RE, " ");
      for (const m of noRefs.matchAll(/(?<![\w.$])\d+(?:\.\d+)?(?![\w.])/g)) {
        const value = Number(m[0]);
        if (value === 0 || value === 1) continue;
        if (constants.length < 20) constants.push({ cell: addr, detail: `${m[0]} in ${truncate(cell.formula!, 80)}` });
        else constants.push({ cell: "", detail: "" }); // count only
      }

      if (cell.formula!.length > 200 && long.length < 10) {
        long.push({ cell: addr, detail: `${cell.formula!.length} chars` });
      }
    }
  }

  const { cycles, cappedRanges } = findCycles(workbook);
  return {
    circularReferences: { count: cycles.length, examples: cycles.slice(0, 5), cappedRanges },
    volatileFunctions: {
      count: Object.values(byFunction).reduce((a, b) => a + b, 0),
      byFunction,
      examples: volatile,
    },
    hardcodedConstants: { count: constants.length, examples: constants.filter((c) => c.cell).slice(0, 20) },
    externalWorkbookLinks: workbook.externalLinks,
    structuredTableRefs,
    mergedCells,
    longFormulas: { count: long.length, examples: long },
  };
}

function truncate(s: string, n: number): string {
  return s.length > n ? `${s.slice(0, n)}…` : s;
}
