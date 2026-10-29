import { appendFileSync, existsSync, mkdirSync, readdirSync, renameSync, statSync, unlinkSync } from "node:fs";
import { readdir, readFile, stat } from "node:fs/promises";
import { join } from "node:path";
import { randomUUID } from "node:crypto";
import { release as systemRelease } from "node:os";
import { AsyncLocalStorage } from "node:async_hooks";
import type { McpDiagnosticEvent, McpDiagnosticRole, McpDiagnosticSource, McpDiagnosticsSnapshot } from "../shared/mcpDiagnostics.js";

export const MCP_LOG_LIMIT = 1024 * 1024;
export interface McpTrace { connectionId: string; traceId: string; source: McpDiagnosticSource }
export const mcpTraceContext = new AsyncLocalStorage<McpTrace>();
type Details = Partial<Omit<McpDiagnosticEvent, "time" | "sessionId" | "pid" | "role" | "event" | "stage">>;
let current: McpDiagnosticLogger | undefined;
export function currentMcpDiagnostics(): McpDiagnosticLogger | undefined { return current; }
export function initializeMcpDiagnostics(directory: string, role: McpDiagnosticRole, version: string, source: McpDiagnosticSource = "external") {
  current = new McpDiagnosticLogger(directory, role, version, source);
  return current;
}
export function isProcessAlive(pid: number): boolean {
  try { process.kill(pid, 0); return true; } catch (error) { return (error as NodeJS.ErrnoException).code === "EPERM"; }
}
function safe(value: string): string {
  return value.replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/giu, "[email]")
    .replace(/(?:bearer\s+|token[=:]\s*)\S+/giu, "[credential]").replace(/[\r\n\x00-\x1f]/g, " ").slice(0, 128);
}
export function diagnosticErrorCode(error: unknown): string {
  const text = error instanceof Error ? error.message : "";
  if (/未运行或未启用/.test(text)) return "WORKBENCH_DISABLED_OR_MISSING";
  if (/认证失败/.test(text)) return "BRIDGE_AUTH_FAILED";
  if (/超时|timed out|timeout/i.test(text)) return "TIMEOUT";
  if (/LOG_WRITE_FAILED/.test(text)) return "LOG_WRITE_FAILED";
  if (/connection closed/i.test(text)) return "MCP_CONNECTION_CLOSED";
  if (/上线年月/.test(text)) return "INVALID_TARGET_MONTH";
  const code = (error as NodeJS.ErrnoException | undefined)?.code;
  if (typeof code === "string" && /^E[A-Z0-9_]{1,30}$/.test(code)) return code;
  if (/断开|响应无效/.test(text)) return "BRIDGE_DISCONNECTED";
  return "OPERATION_FAILED";
}
export class McpDiagnosticLogger {
  readonly sessionId = randomUUID();
  readonly directory: string;
  readonly path: string;
  available = true;
  private ended = false;
  private sequence = 0;
  private clientName?: string;
  private clientVersion?: string;
  private protocolVersion?: string;
  constructor(userData: string, readonly role: McpDiagnosticRole, private readonly version: string, readonly source: McpDiagnosticSource = "external") {
    this.directory = join(userData, "logs", "mcp");
    this.path = join(this.directory, `mcp-${role}-${process.pid}-${this.sessionId}.log`);
    try { mkdirSync(this.directory, { recursive: true }); this.prune(); } catch { this.fail(); }
    this.record("process-start", "process", { version, system: `${process.platform}/${systemRelease()}/${process.arch}`, ppid: process.ppid,
      runtime: process.versions.electron ? "electron" : "node", runtimeVersion: process.versions.node });
  }
  record(event: string, stage: string, details: Details = {}): void {
    const entry: McpDiagnosticEvent = { time: new Date().toISOString(), sequence: ++this.sequence, sessionId: this.sessionId,
      pid: process.pid, role: this.role, source: details.source ?? this.source, event: safe(event), stage: safe(stage), level: details.level ?? "info" };
    const strings = ["connectionId", "traceId", "method", "tool", "clientName", "clientVersion", "protocolVersion", "version", "system", "errorCode", "runtime", "runtimeVersion"] as const;
    for (const key of strings) if (typeof details[key] === "string") entry[key] = safe(details[key]);
    this.clientName = entry.clientName ?? this.clientName;
    this.clientVersion = entry.clientVersion ?? this.clientVersion;
    this.protocolVersion = entry.protocolVersion ?? this.protocolVersion;
    entry.clientName = this.clientName; entry.clientVersion = this.clientVersion; entry.protocolVersion = this.protocolVersion; entry.version = this.version;
    for (const key of ["ppid", "durationMs", "count"] as const) if (Number.isFinite(details[key])) entry[key] = details[key];
    const line = JSON.stringify(entry) + "\n";
    try {
      if (existsSync(this.path) && statSync(this.path).size + Buffer.byteLength(line) > MCP_LOG_LIMIT) {
        const previous = this.path.replace(/\.log$/, ".previous.log");
        if (existsSync(previous)) unlinkSync(previous);
        renameSync(this.path, previous);
      }
      appendFileSync(this.path, line, { encoding: "utf8", mode: 0o600 });
      this.available = true;
    } catch { this.fail(); }
  }
  end(event = "process-exit"): void {
    if (this.ended) return;
    this.record(event, "closed"); this.ended = true;
    try { this.prune(); } catch { /* diagnostic retention is best effort */ }
  }
  private fail(): void {
    if (this.available) process.stderr.write("MCP diagnostics unavailable: LOG_WRITE_FAILED\n");
    this.available = false;
  }
  private prune(): void {
    const groups = new Map<string, { pid: number; files: string[]; modified: number }>();
    for (const name of readdirSync(this.directory)) {
      const match = name.match(/^(mcp-(?:stdio|workbench)-(\d+)-[a-f0-9-]{36})(?:\.previous)?\.log$/);
      if (!match) continue;
      const group = groups.get(match[1]) ?? { pid: Number(match[2]), files: [], modified: 0 };
      group.files.push(name); group.modified = Math.max(group.modified, statSync(join(this.directory, name)).mtimeMs); groups.set(match[1], group);
    }
    const inactive = [...groups.values()].filter(group => !isProcessAlive(group.pid)).sort((a,b) => b.modified-a.modified);
    for (const group of inactive.slice(20)) for (const name of group.files) unlinkSync(join(this.directory, name));
  }
}
const parsedLogCache = new Map<string, { signature: string; events: McpDiagnosticEvent[] }>();
export async function readMcpDiagnostics(userData: string): Promise<McpDiagnosticsSnapshot> {
  const directory = join(userData, "logs", "mcp");
  const events: McpDiagnosticEvent[] = [];
  try {
    const names = await readdir(directory);
    for (const path of parsedLogCache.keys()) if (path.startsWith(directory) && !names.includes(path.slice(directory.length+1))) parsedLogCache.delete(path);
    for (const name of names) {
      if (!/^mcp-(?:stdio|workbench)-\d+-[a-f0-9-]{36}(?:\.previous)?\.log$/.test(name)) continue;
      const path = join(directory, name);
      const metadata = await stat(path);
      if (metadata.size > MCP_LOG_LIMIT) continue;
      const signature = `${metadata.mtimeMs}/${metadata.size}/${metadata.ino}`;
      const cached = parsedLogCache.get(path);
      if (cached?.signature === signature) { events.push(...cached.events); continue; }
      const fileEvents: McpDiagnosticEvent[] = [];
      for (const line of (await readFile(path, "utf8")).split("\n")) {
        try { const item = JSON.parse(line) as McpDiagnosticEvent;
          if (typeof item.sessionId === "string" && typeof item.time === "string" && ["stdio", "workbench"].includes(item.role)) fileEvents.push(item);
        } catch { /* ignore a partial trailing write */ }
      }
      parsedLogCache.set(path, { signature, events: fileEvents }); events.push(...fileEvents);
    }
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "ENOENT") return { directory, available: false, warning: "MCP 日志无法读取", sessions: [], events: [] };
  }
  events.sort((a,b) => a.time.localeCompare(b.time) || (a.sessionId === b.sessionId ? a.sequence-b.sequence : 0));
  const sessions = new Map<string, McpDiagnosticsSnapshot["sessions"][number]>();
  for (const event of events) {
    const old = sessions.get(event.sessionId);
    sessions.set(event.sessionId, { sessionId: event.sessionId, role: event.role, source: old?.source ?? event.source, pid: event.pid,
      lastActivity: event.time, stage: event.stage, event: event.event,
      active: !["process-exit", "connection-closed", "process-error"].includes(event.event) && isProcessAlive(event.pid),
      clientName: event.clientName ?? old?.clientName, clientVersion: event.clientVersion ?? old?.clientVersion,
      lastFailure: event.level === "error" || event.level === "warning" ? event : old?.lastFailure });
  }
  const available = current?.available !== false;
  return { directory, available, warning: available ? undefined : "MCP 日志写入失败，请检查目录权限和客户端 stderr",
    sessions: [...sessions.values()].sort((a,b)=>b.lastActivity.localeCompare(a.lastActivity)), events: events.slice(-100).reverse() };
}
export async function createMcpDiagnosticArchive(userData: string, version: string): Promise<Buffer> {
  const { default: JSZip } = await import("jszip");
  const info = await readMcpDiagnostics(userData);
  const zip = new JSZip();
  zip.file("summary.json", JSON.stringify({ version, exportedAt: new Date().toISOString(), platform: process.platform,
    arch: process.arch, node: process.versions.node, electron: process.versions.electron, available: info.available, sessions: info.sessions }, null, 2));
  const names = await readdir(info.directory).catch(() => [] as string[]);
  for (const name of names) {
    if (!/^mcp-(?:stdio|workbench)-\d+-[a-f0-9-]{36}(?:\.previous)?\.log$/.test(name)) continue;
    const path = join(info.directory, name);
    try { if ((await stat(path)).size <= MCP_LOG_LIMIT) zip.file(name, await readFile(path)); } catch { /* concurrent cleanup */ }
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
