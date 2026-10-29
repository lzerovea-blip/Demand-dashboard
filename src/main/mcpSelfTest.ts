import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import type { McpSelfTestResult } from "../shared/mcpDiagnostics.js";
import { diagnosticErrorCode, type McpDiagnosticLogger } from "./mcpDiagnostics.js";
import type { McpLaunchConfiguration } from "../shared/mcpLaunch.js";

export async function runMcpSelfTest(launch: McpLaunchConfiguration, userData: string, version: string, logger?: McpDiagnosticLogger): Promise<McpSelfTestResult> {
  if (!launch.available) {
    logger?.record("self-test-failed", "process", { source: "self-test", level: "error", errorCode: "MCP_RUNTIME_MISSING" });
    return { success: false, stage: "process", message: "独立 MCP 运行文件不完整，请重新解压新版完整文件夹。", checkedAt: new Date().toISOString() };
  }
  const client = new Client({ name: "solution-workbench-self-test", version });
  const transport = new StdioClientTransport({ command: launch.command, args: launch.args,
    env: { ...process.env, SOLUTION_REQUIREMENTS_USER_DATA_DIR: userData, SOLUTION_MCP_SELF_TEST: "1" }, stderr: "pipe" });
  let loggingUnavailable = false;
  transport.stderr?.on("data", chunk => { if (chunk.toString().includes("LOG_WRITE_FAILED")) loggingUnavailable = true; });
  let stage = "handshake";
  logger?.record("self-test-start", stage, { source: "self-test" });
  const timeout = (promise: Promise<unknown>) => {
    let timer: ReturnType<typeof setTimeout>;
    return Promise.race([promise, new Promise((_, reject) => { timer = setTimeout(() => reject(new Error("timeout")), 15000); })]).finally(() => clearTimeout(timer));
  };
  try {
    await timeout(client.connect(transport));
    stage = "tools";
    const tools = await client.listTools({}, { timeout: 15000 });
    const expected = ["list_requirements", "get_requirement", "list_requirement_references", "propose_requirement_changes", "get_proposal_status", "propose_dictionary_additions", "propose_export"];
    if (tools.tools.length !== expected.length || !expected.every(name => tools.tools.some(tool => tool.name === name))) throw new Error("Unexpected tool inventory");
    stage = "tool";
    const result = await client.callTool({ name: "list_requirements", arguments: { query: `__local_self_test_${crypto.randomUUID()}__`, limit: 1 } }, undefined, { timeout: 15000 });
    if (result.isError) throw new Error("工作台未运行或未启用");
    const content = result.content as { type: string; text?: string }[];
    const data = JSON.parse(content.find(item => item.type === "text")?.text ?? "null");
    if (!data || !Array.isArray(data.items) || data.items.length !== 0 || data.total !== 0) throw new Error("Unexpected query response");
    if (loggingUnavailable) { stage = "diagnostics"; throw new Error("LOG_WRITE_FAILED"); }
    logger?.record("self-test-complete", stage, { source: "self-test" });
    return { success: true, stage, message: "本机握手、工具发现和只读查询通过；外部客户端配置需单独验证。", checkedAt: new Date().toISOString() };
  } catch (error) {
    const code = diagnosticErrorCode(error);
    logger?.record("self-test-failed", stage, { source: "self-test", level: "error", errorCode: code });
    return { success: false, stage, message: `本机自检失败：${code}。请查看对应阶段日志。`, checkedAt: new Date().toISOString() };
  } finally { await client.close().catch(() => undefined); }
}
