import { afterEach, describe, expect, it } from "vitest";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { initializeStartupDiagnostics, sanitizeDiagnosticDetail, STARTUP_LOG_MAX_BYTES, writeStartupDiagnostic } from "../src/main/diagnostics";

const directories: string[] = [];
afterEach(async () => {
  for (const directory of directories.splice(0)) await rm(directory, { recursive: true, force: true });
});

describe("启动诊断日志", () => {
  it("记录会话环境并脱敏邮箱和需求标题", async () => {
    const directory = await mkdtemp(join(tmpdir(), "startup-log-"));
    directories.push(directory);
    const info = initializeStartupDiagnostics(directory, "0.8.12");
    writeStartupDiagnostic("renderer-error", "需求“隐私需求”由 owner@company.com 触发");
    const text = await readFile(info.currentLogPath, "utf8");
    expect(text).toContain("version=0.8.12");
    expect(text).toContain("某条需求");
    expect(text).toContain("[redacted-email]");
    expect(text).not.toContain("隐私需求");
    expect(text).not.toContain("owner@company.com");
  });

  it("达到 1MB 后只轮转一个 previous 文件", async () => {
    const directory = await mkdtemp(join(tmpdir(), "startup-log-"));
    directories.push(directory);
    const info = initializeStartupDiagnostics(directory, "0.8.12");
    await writeFile(info.currentLogPath, "x".repeat(STARTUP_LOG_MAX_BYTES - 64));
    writeStartupDiagnostic("large-event", "x".repeat(200));
    expect(existsSync(join(info.directory, "startup.previous.log"))).toBe(true);
    expect((await stat(info.currentLogPath)).size).toBeLessThanOrEqual(STARTUP_LOG_MAX_BYTES);
  });

  it("精简异常文本限制单行和长度", () => {
    const value = sanitizeDiagnosticDetail(`a@b.com\n${"x".repeat(3000)}`);
    expect(value).not.toContain("\n");
    expect(value.length).toBeLessThanOrEqual(2000);
  });
});
