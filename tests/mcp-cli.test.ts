import { it, expect } from "vitest";
import { build } from "esbuild";
import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { readMcpDiagnostics } from "../src/main/mcpDiagnostics";

it("真实独立进程握手及工具发现成功，工作台未启用有明确错误，stdout仅协议", async () => {
  const directory = await mkdtemp(join(tmpdir(), "mcp-cli-test-"));
  const script = join(directory, "server.cjs");
  await build({ entryPoints: ["src/main/mcpCli.ts"], outfile: script, bundle: true, platform: "node", format: "cjs", target: "node22" });
  const client = new Client({ name: "standalone-test", version: "1" });
  const transport = new StdioClientTransport({ command: process.execPath, args: [script], env: { ...process.env, SOLUTION_REQUIREMENTS_USER_DATA_DIR: directory }, stderr: "pipe" });
  try {
    await client.connect(transport);
    expect((await client.listTools()).tools).toHaveLength(7);
    expect((await client.callTool({ name: "list_requirements", arguments: { limit: 1 } })).isError).toBe(true);
    const info = await readMcpDiagnostics(directory);
    expect(info.events.some(event => event.event === "process-start" && event.runtime === "node")).toBe(true);
    expect(info.events.some(event => event.event === "handshake-complete")).toBe(true);
  } finally { await client.close(); await rm(directory, { recursive: true, force: true }); }
}, 15000);

it("初始化前 stdin 关闭有准确诊断且进程自然退出", async () => {
  const directory = await mkdtemp(join(tmpdir(), "mcp-eof-test-"));
  const script = join(directory, "server.cjs");
  await build({ entryPoints: ["src/main/mcpCli.ts"], outfile: script, bundle: true, platform: "node", format: "cjs", target: "node22" });
  const child = spawn(process.execPath, [script], { env: { ...process.env, SOLUTION_REQUIREMENTS_USER_DATA_DIR: directory }, stdio: ["pipe", "pipe", "pipe"] });
  let stdout = ""; child.stdout.on("data", value => stdout += value);
  try {
    const completed = new Promise<number | null>((resolve, reject) => { child.once("exit", resolve); child.once("error", reject); });
    child.stdin.end();
    expect(await completed).toBe(0);
    expect(stdout).toBe("");
    expect((await readMcpDiagnostics(directory)).events.some(event => event.errorCode === "STDIN_CLOSED_BEFORE_INITIALIZE")).toBe(true);
  } finally { child.kill(); await rm(directory, { recursive: true, force: true }); }
}, 15000);
