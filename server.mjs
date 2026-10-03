#!/usr/bin/env node
// spreadsheet-audit-mcp: lets Claude audit an .xlsx workbook for the classic model errors.
// Local only: reads the file you point it at, never writes to it, never uses the network.
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { spawnSync } from "node:child_process";
import { existsSync, statSync } from "node:fs";
import { join, dirname, resolve, extname } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const SCRIPT = join(HERE, "xlsx_audit.py");
const PYTHON = process.env.PYTHON || (process.platform === "win32" ? "python" : "python3");
const SEV = ["HIGH", "MEDIUM", "LOW", "INFO"];
const MAX_MB = 25;

function run(args, timeoutMs) {
  const r = spawnSync(PYTHON, [SCRIPT, ...args], { encoding: "utf8", timeout: timeoutMs, maxBuffer: 64 * 1024 * 1024 });
  if (r.error) throw new Error(`Could not run ${PYTHON}: ${r.error.message}`);
  if (r.status !== 0 && r.status !== 1) throw new Error(`Auditor failed (${r.status}): ${(r.stderr || "").slice(-400)}`);
  const out = (r.stdout || "").trim();
  return JSON.parse(out.slice(out.indexOf("{")));
}

function checkFile(p) {
  const abs = resolve(p);
  if (!existsSync(abs) || !statSync(abs).isFile()) throw new Error(`Not a file: ${abs}`);
  if (![".xlsx", ".xlsm"].includes(extname(abs).toLowerCase())) throw new Error("Only .xlsx / .xlsm files are supported");
  if (statSync(abs).size > MAX_MB * 1024 * 1024) throw new Error(`File is larger than ${MAX_MB} MB`);
  return abs;
}

const server = new McpServer({ name: "spreadsheet-audit", version: "0.1.0" });
const ro = { readOnlyHint: true, idempotentHint: true, openWorldHint: false };

server.registerTool(
  "audit_workbook",
  {
    title: "Audit a spreadsheet for formula and model errors",
    description:
      "Audits an .xlsx workbook: error values after an independent recalculation, typed-over formulas, formulas that break the pattern of their neighbours, " +
      "hard-coded numbers inside formulas, approximate-match lookups, volatile functions, external links, hidden sheets/rows/columns and broken named ranges. " +
      "Heuristic: it points at where models usually go wrong, it cannot prove a model is right.",
    inputSchema: {
      path: z.string().describe("Path to the .xlsx file"),
      recalc: z.boolean().default(true).describe("Recalculate with an independent engine to find error values (slower on big files)"),
      min_severity: z.enum(SEV).default("LOW"),
      limit: z.number().int().min(1).max(200).default(50),
    },
    annotations: ro,
  },
  async ({ path, recalc, min_severity, limit }) => {
    try {
      const abs = checkFile(path);
      const res = run([abs, ...(recalc ? [] : ["--no-recalc"])], 180_000);
      const cutoff = SEV.indexOf(min_severity);
      const shown = res.findings.filter((f) => SEV.indexOf(f.severity) <= cutoff).slice(0, limit);
      const payload = { workbook: res.workbook, sheets: res.sheets, cells: res.cells, formulas: res.formulas, counts: res.counts, recalculation_note: res.recalculation_note, shown: shown.length, findings: shown,
        note: "Heuristic audit. Review each finding; a clean result does not prove the model is correct." };
      return { content: [{ type: "text", text: JSON.stringify(payload, null, 1) }], structuredContent: payload };
    } catch (e) {
      return { isError: true, content: [{ type: "text", text: String(e.message || e) }] };
    }
  }
);

server.registerTool(
  "inspect_cell",
  {
    title: "Show one cell's formula and what it reads",
    description: "Returns the formula (or value), the stored result if the file has one, and the cells the formula reads. Use it to follow a suspicious number back to its inputs.",
    inputSchema: { path: z.string(), sheet: z.string(), cell: z.string().regex(/^[A-Za-z]{1,3}\d{1,7}$/, "e.g. B12") },
    annotations: ro,
  },
  async ({ path, sheet, cell }) => {
    try {
      const abs = checkFile(path);
      const res = run([abs, "--cell", `${sheet}!${cell.toUpperCase()}`], 60_000);
      return { content: [{ type: "text", text: JSON.stringify(res, null, 1) }], structuredContent: res };
    } catch (e) {
      return { isError: true, content: [{ type: "text", text: String(e.message || e) }] };
    }
  }
);

await server.connect(new StdioServerTransport());
