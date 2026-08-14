# xlsx-audit-mcp

An [MCP](https://modelcontextprotocol.io) server that **audits Excel workbooks**. Other Excel MCP servers read and write your data — this one reviews your *model*:

- *"What feeds the Total cell on the Summary sheet?"* — precedent tracing
- *"If I change this assumption, what breaks?"* — dependent tracing, including cells that consume it through ranges like `SUM(A1:A40)`
- *"Audit this workbook"* — circular references with example chains, volatile functions (`INDIRECT`, `OFFSET`, `NOW`, `RAND`...), hardcoded constants buried inside formulas, external workbook links, merged cells, extra-long formulas

Spreadsheet mistakes are famously expensive. This is the "trace precedents" discipline auditors apply by hand, exposed to an LLM for a whole workbook at once. Local files only; nothing leaves your machine.

## Quick start

**Claude Code**

```bash
claude mcp add xlsx-audit -- npx -y xlsx-audit-mcp
```

**Claude Desktop** — add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "xlsx-audit": {
      "command": "npx",
      "args": ["-y", "xlsx-audit-mcp"]
    }
  }
}
```

Then: *"Audit C:\\models\\budget-2026.xlsx and tell me what looks fragile."*

## Tools

| Tool | What it does |
|------|--------------|
| `workbook_overview` | Sheets, dimensions, formula counts, defined names, external links |
| `list_formulas` | Formulas with addresses and cached values, filterable (`INDIRECT`, `VLOOKUP`, ...) |
| `trace_cell` | One cell's formula, value, precedents, and dependents (direct + via ranges) |
| `audit_workbook` | Ranked risk report across the whole model |

## How it works

- **Reference tokenizer** that understands real formulas: string literals are stripped first (the `"A1"` in `INDIRECT("A1")` is not a reference), function names can't collide (the `G10` in `LOG10(...)` is not a cell), `$` absolutes, quoted sheet names (`'My Data'!A1`), and ranges are handled.
- **Shared formulas are materialized.** Excel stores filled formulas once with an offset scheme; the loader translates them per-cell (relative refs shifted, absolutes preserved), so dependency queries see what each cell actually computes.
- **Ranges are never expanded for storage** — dependents queries use range-containment tests, and cycle detection caps range fan-out (a `SUM(A:A)` can't explode the graph; capped ranges are reported, not silently dropped).
- **No formula evaluation.** Cached values from the file are shown instead — no spreadsheet engine dependency.

Known limitations: R1C1 notation and structured table references (`[@Column]`) are counted but not resolved into the graph.

## Development

```bash
npm install
npm test                 # offline tests — synthetic workbooks built in-suite
npm run build            # tsc → dist/
node scripts/smoke.mjs   # end-to-end: generates a workbook, drives the server over stdio
```

Architecture: [`src/xlsx.ts`](src/xlsx.ts) (zip + XML → workbook model) and [`src/formulas.ts`](src/formulas.ts) (tokenizer, graph, smells) are pure logic; [`src/index.ts`](src/index.ts) is the MCP wiring.

## License

MIT
