# xlsx-audit-mcp — Design

**Date:** 2026-08-14
**Goal:** Fifth portfolio MCP server; new engineering genre: static analysis.

## Idea selection

Calendar (.ics) analytics was checked first and found crowded (CalDAV/feed
servers). Excel MCP servers also exist in numbers — but every one of them is a
**data read/write tool**. None audits the *model*: no dependency tracing, no
circular-reference detection, no risk report. Spreadsheet errors are a
famously expensive problem (financial models, planning sheets), and "trace
precedents" is what auditors actually do in Excel by hand.

Positioning: *Excel MCP servers read your data. This one reviews your model.*

## What it does

- *"What feeds the Total cell on the Summary sheet?"* (precedents)
- *"If I change this assumption, what breaks?"* (dependents, incl. via ranges)
- *"Audit this workbook"* — circular references, volatile functions
  (INDIRECT/OFFSET/NOW/RAND), hardcoded constants buried in formulas,
  external-workbook links, merged cells, extra-long formulas.

## Architecture

- `src/xlsx.ts` — workbook loader: unzip (fflate) + XML (fast-xml-parser) →
  sheets, cells with formulas and cached values, defined names, merged
  ranges, external links. Whole-file in memory (xlsx are typically tens of
  MB; ceiling documented). Cache keyed by path+mtime.
- `src/formulas.ts` — the analysis core, pure and heavily tested:
  - **Reference tokenizer:** extracts A1-style refs from formula text after
    stripping quoted strings; rejects function-name collisions (the `G10` in
    `LOG10(` is not a cell), handles `$` absolutes, sheet-qualified refs
    (incl. quoted sheet names), and ranges.
  - **Dependency graph:** edges formula-cell → referenced cells. Ranges are
    kept as ranges (never expanded for storage); dependents queries use
    range-containment tests instead.
  - **Cycle detection:** iterative three-color DFS over formula cells; a
    range edge contributes only when its area is ≤ a cap (10k cells), so a
    `SUM(A:A)` can't explode the graph — capped ranges are reported, not
    silently dropped.
  - **Smells:** volatile functions, hardcoded numeric literals (heuristic:
    numeric tokens outside refs/strings, excluding 0/1/exponent-like small
    ints), external links, merged cells, formulas > 200 chars.
- `src/index.ts` — MCP wiring, four tools.

## Tools

| Tool | Purpose |
|------|---------|
| `workbook_overview(path)` | Sheets with dimensions and formula counts, defined names, external links, merged-cell totals |
| `list_formulas(path, sheet?, contains?, limit?)` | Formulas with addresses and cached values — filter e.g. for "INDIRECT" or "VLOOKUP" |
| `trace_cell(path, cell)` | One cell's formula, value, precedents, and dependents (direct + via ranges, capped) |
| `audit_workbook(path)` | Ranked risk report across all smells with example locations |

## Limitations (documented)

R1C1 notation and structured table references (`[@Column]`) are detected and
counted but not resolved into the graph. Formula *evaluation* is out of scope
(cached values from the file are shown instead — no HyperFormula dependency).

## Testing

Synthetic .xlsx files are **built in the test suite** (fflate `zipSync` +
hand-written sheet XML), covering: cross-sheet refs, quoted sheet names,
absolute refs, ranges, a deliberate circular chain, LOG10-style collisions,
volatile functions, and hardcoded constants. Smoke generates a workbook in a
temp dir and drives the server over stdio.
