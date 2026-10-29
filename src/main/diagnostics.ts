import { appendFileSync, copyFileSync, existsSync, mkdirSync, rmSync, statSync, truncateSync } from "node:fs";
import { join } from "node:path";
import { release as systemRelease } from "node:os";

const MAX_LOG_BYTES = 1024 * 1024;
const MAX_DETAIL_LENGTH = 2000;

let paths: { directory: string; currentLogPath: string; previousLogPath: string } | null = null;
let sessionId = "not-initialized";

export function initializeStartupDiagnostics(userDataDirectory: string, appVersion: string): ReturnType<typeof diagnosticInfo> {
  const directory = join(userDataDirectory, "logs");
  mkdirSync(directory, { recursive: true });
  paths = {
    directory,
    currentLogPath: join(directory, "startup.log"),
    previousLogPath: join(directory, "startup.previous.log"),
  };
  sessionId = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  rotateIfNeeded(0);
  writeStartupDiagnostic("session-start", [
    `version=${appVersion}`,
    `platform=${process.platform}`,
    `system=${systemRelease()}`,
    `arch=${process.arch}`,
    `node=${process.versions.node}`,
    `electron=${process.versions.electron ?? "unknown"}`,
    `chrome=${process.versions.chrome ?? "unknown"}`,
  ].join(" "));
  return diagnosticInfo();
}

export function diagnosticInfo(): { directory: string; currentLogPath: string } {
  if (!paths) throw new Error("诊断日志尚未初始化");
  return { directory: paths.directory, currentLogPath: paths.currentLogPath };
}

export function sanitizeDiagnosticDetail(value: unknown): string {
  const raw = value instanceof Error ? `${value.name}: ${value.message}` : String(value ?? "");
  return raw
    .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/giu, "[redacted-email]")
    .replace(/需求“[^”]*”/gu, "某条需求")
    .replace(/[\r\n]+/g, " ")
    .slice(0, MAX_DETAIL_LENGTH);
}

export function writeStartupDiagnostic(event: string, detail: unknown = ""): string | undefined {
  if (!paths) return undefined;
  try {
    const entry = `[${new Date().toISOString()}] session=${sessionId} ${safeEvent(event)}${detail === "" ? "" : `: ${sanitizeDiagnosticDetail(detail)}`}\n`;
    rotateIfNeeded(Buffer.byteLength(entry, "utf8"));
    appendFileSync(paths.currentLogPath, entry, "utf8");
    return paths.currentLogPath;
  } catch {
    return undefined;
  }
}

export function beginDiagnosticStage(stage: string): () => void {
  const startedAt = Date.now();
  writeStartupDiagnostic("stage-start", stage);
  return () => writeStartupDiagnostic("stage-complete", `${stage} durationMs=${Date.now() - startedAt}`);
}

function rotateIfNeeded(incomingBytes: number): void {
  if (!paths || !existsSync(paths.currentLogPath)) return;
  if (statSync(paths.currentLogPath).size + incomingBytes <= MAX_LOG_BYTES) return;
  rmSync(paths.previousLogPath, { force: true });
  copyFileSync(paths.currentLogPath, paths.previousLogPath);
  truncateSync(paths.currentLogPath, 0);
}

function safeEvent(value: string): string {
  return value.replace(/[^a-z0-9_.-]/giu, "-").slice(0, 80) || "event";
}

export const STARTUP_LOG_MAX_BYTES = MAX_LOG_BYTES;
