# spreadsheet-audit-mcp

An [MCP](https://modelcontextprotocol.io) server that lets Claude (or any MCP client) audit an `.xlsx` workbook for the errors that quietly break spreadsheet models. Local and read-only: it opens the file you point it at, never writes to it, never uses the network.

## What it checks
| Rule | Severity | What it catches |
|---|---|---|
| `CALC-ERROR` | HIGH | `#DIV/0!`, `#REF!`, `#N/A`, `#VALUE!`, `#NAME?` after an **independent recalculation** (pure-Python [`formulas`](https://pypi.org/project/formulas/) engine) |
| `CACHED-ERROR` | HIGH | error values stored in the file by the last app that saved it |
| `TYPED-OVER` | HIGH | a hard-coded number sitting inside a filled-down formula run |
| `PATTERN-BREAK` | HIGH | the one formula that differs from consistent neighbours above and below |
| `EXT-LINK` | HIGH | formulas that read another file |
| `NAME-REF` | HIGH | named ranges that point to `#REF!` |
| `MAGIC-NUMBER` | MEDIUM | rates and thresholds buried in formulas instead of labelled inputs (rounding tolerances like `<=0.005` are ignored) |
| `LOOKUP-APPROX` | MEDIUM | `VLOOKUP` with approximate match (omitted or `TRUE`) |
| `HIDDEN-SHEET` | MEDIUM | hidden or very-hidden sheets |
| `HIDDEN-ROWCOL`, `VOLATILE` | LOW | hidden rows/columns; `INDIRECT`, `OFFSET`, `TODAY`, `NOW`, `RAND` |

Repeated `MAGIC-NUMBER`, `VOLATILE` and `LOOKUP-APPROX` hits are grouped per sheet with a count, so 700 repeats do not bury the real problems.

## Tools
| Tool | What it does |
|---|---|
| `audit_workbook` | Full audit (`path`, `recalc`, `min_severity`, `limit`). |
| `inspect_cell` | One cell's formula, stored result and the cells it reads (`path`, `sheet`, `cell`). |

## Install and use
Requires Node 18+, Python 3.9+ with `openpyxl` (and optionally `formulas` for the recalculation check): `pip install openpyxl formulas`.
```bash
npm install
npm test        # 14 Python tests + 8 end-to-end MCP client checks
```
Claude Desktop / Claude Code config:
```json
{ "mcpServers": { "spreadsheet-audit": { "command": "node", "args": ["/absolute/path/to/server.mjs"] } } }
```
Then ask: "Audit ./budget.xlsx and tell me which findings could change the bottom line."

## Honest limits
Heuristic. It points at where models usually go wrong; it cannot prove a model is right, and some findings are intentional (for example a sensitivity grid with fixed steps). If the recalculation engine cannot handle a workbook it says so in `recalculation_note` and the static checks still run.

## Need a spreadsheet model built or reviewed, or a custom MCP server for your tools?
Open an issue or message the author.

License: MIT
