import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { CATEGORIES, OVERSEAS_REGIONS, REQUIREMENT_SCOPES } from "../shared/types.js";
import { callLocalMcpBridge } from "./mcpBridge.js";
import { currentMcpDiagnostics, diagnosticErrorCode, mcpTraceContext, type McpDiagnosticLogger } from "./mcpDiagnostics.js";
import { McpDiagnosticTransport } from "./mcpDiagnosticTransport.js";
import { MCP_APP_VERSION } from "../shared/mcpVersion.js";
import { mcpDictionaryBatchSchema } from "../shared/mcpDictionaries.js";
import { MCP_TARGET_MONTH_PATTERN } from "../shared/mcp.js";
import { mcpExportRequestSchema } from "../shared/mcpExports.js";

const fields = z.object({
  title: z.string(), description: z.string(), domainL0Id: z.string(), domainId: z.string(),
  scope: z.enum(REQUIREMENT_SCOPES), allowLatestVersionUpgrade: z.boolean(),
  overseasRegions: z.array(z.enum(OVERSEAS_REGIONS)), category: z.enum(CATEGORIES),
  targetMonth: z.string().regex(MCP_TARGET_MONTH_PATTERN, "上线年月必须使用 YYYY-MM，例如 2027-03；月份为 01–12").describe("必填上线年月，格式 YYYY-MM，例如 2027-03。不接受 2027年3月、2027/03、时间戳或无效月份。"), productIds: z.array(z.string()),
  planningMemberId: z.string(), productManagerMemberId: z.string(), seMemberId: z.string(),
  algorithmMemberId: z.string().optional(), deviceWorkloadPm: z.number(), appWorkloadPm: z.number(),
  cloudWorkloadPm: z.number(), researchAppWorkloadPm: z.number(), algorithmWorkloadPm: z.number(),
}).strict();

const operation = z.discriminatedUnion("type", [
  z.object({ type: z.literal("create"), fields }),
  z.object({ type: z.literal("update"), id: z.string(), fields: fields.partial().strict() }),
  z.object({ type: z.literal("duplicate"), id: z.string() }),
  z.object({ type: z.literal("delete"), id: z.string() }),
]);

export function buildMcpServer(userData: string, logger?: McpDiagnosticLogger): McpServer {
  const server = new McpServer({ name: "solution-requirements-manager", version: MCP_APP_VERSION });
  const run = async (action: string, args: unknown) => {
    try { return { content: [{ type: "text" as const, text: JSON.stringify(await callLocalMcpBridge(userData, action, args, mcpTraceContext.getStore(), logger)) }] }; }
    catch (error) { logger?.record("request-failed", "tool", { ...mcpTraceContext.getStore(), tool: action, level: "error", errorCode: diagnosticErrorCode(error) });
      return { isError: true, content: [{ type: "text" as const, text: error instanceof Error ? error.message : "MCP 请求失败" }] }; }
  };
  server.registerTool("list_requirements", { description: "分页查询本机需求池；不返回图片内容、成员邮箱或 PPT 权限。",
    inputSchema: { query: z.string().optional(), offset: z.number().int().min(0).optional(), limit: z.number().int().min(1).max(100).optional() } },
    (args) => run("list_requirements", args));
  server.registerTool("get_requirement", { description: "按需求 ID 读取需求业务字段。", inputSchema: { id: z.string().min(1) } },
    (args) => run("get_requirement", args));
  server.registerTool("list_requirement_references", { description: "读取领域、产品、成员 ID/姓名/角色、岗位选项和路标/泳道字典。新增 L1 或泳道须先使用已有的父级 ID；新建父级需先经工作台批准，再查询 ID。不返回成员邮箱。" },
    () => run("list_requirement_references", {}));
  server.registerTool("propose_requirement_changes", { description: "提出最多 25 项新增、修改、复制或删除需求；仅进入工作台人工确认，不直接写入。",
    inputSchema: { operations: z.array(operation).min(1).max(25) } },
    (args) => run("propose_requirement_changes", args));
  server.registerTool("get_proposal_status", { description: "查询提案待确认、导出中、完成、取消、失效或冲突状态。导出完成后 exportResult 返回本机文件路径，不返回文件内容，也不会自动上传。", inputSchema: { id: z.string().min(1) } },
    (args) => run("get_proposal_status", args));
  server.registerTool("propose_dictionary_additions", { description: "提出 1–25 项字典新增，必须由工作台人工确认后整批写入。支持领域 L0/L1、产品、成员、路标、泳道。只新增，不改名、不删除、不替换、不授权 PPT。成员须提供姓名、唯一工号、完整企业邮箱；负责领域填写已有 L1 ID。路标类型 standard=普通卡片、medical=医疗认证。", inputSchema: { entries: mcpDictionaryBatchSchema } },
    (args) => run("propose_dictionary_additions", args));
  server.registerTool("propose_export", { description: "提出文件导出请求，必须在工作台确认并选择本机保存位置，不直接写文件。format: ppt=受企业IRM保护的PPT；requirements_excel=需求池协作Excel；tiering_excel=只读分层Excel；workspace=完整需求包（包含图片、人员账号和PPT白名单，全部数据不按时间过滤）。除完整包外须 start/end=YYYYH1或YYYYH2，例如2027H1。PPT最多4个半年；includeSummary默认true，roadmapIds省略=全部路标，[]则不选路标，仅封面/汇总，ID可查询字典。PPT必须已有白名单及Windows企业IRM环境，禁止无保护降级。不接受路径、账号或上传地址。返回提案ID后通过get_proposal_status查询结果。", inputSchema: { request: mcpExportRequestSchema } },
    (args) => run("propose_export", args));
  return server;
}

export async function serveMcpStdio(userData: string, onClosed?: () => void): Promise<void> {
  const logger = currentMcpDiagnostics();
  const server = buildMcpServer(userData, logger);
  const transport = logger ? new McpDiagnosticTransport(new StdioServerTransport(), logger) : new StdioServerTransport();
  let initialized = false;
  server.server.oninitialized = () => { initialized = true; if (transport instanceof McpDiagnosticTransport) transport.initialized(server.server.getClientVersion()); };
  process.stdin.once("end", () => {
    logger?.record("stdin-ended", "stdio", { level: initialized ? "info" : "warning", errorCode: initialized ? "STDIN_CLOSED" : "STDIN_CLOSED_BEFORE_INITIALIZE" });
    void server.close().finally(() => onClosed?.());
  });
  await server.connect(transport);
}
