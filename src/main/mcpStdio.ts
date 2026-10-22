import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { CATEGORIES, OVERSEAS_REGIONS, REQUIREMENT_SCOPES } from "../shared/types.js";
import { callLocalMcpBridge } from "./mcpBridge.js";

const fields = z.object({
  title: z.string(), description: z.string(), domainL0Id: z.string(), domainId: z.string(),
  scope: z.enum(REQUIREMENT_SCOPES), allowLatestVersionUpgrade: z.boolean(),
  overseasRegions: z.array(z.enum(OVERSEAS_REGIONS)), category: z.enum(CATEGORIES),
  targetMonth: z.string(), productIds: z.array(z.string()),
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

export function buildMcpServer(userData: string): McpServer {
  const server = new McpServer({ name: "solution-requirements-manager", version: "0.9.8" });
  const run = async (action: string, args: unknown) => {
    try { return { content: [{ type: "text" as const, text: JSON.stringify(await callLocalMcpBridge(userData, action, args)) }] }; }
    catch (error) { return { isError: true, content: [{ type: "text" as const, text: error instanceof Error ? error.message : "MCP 请求失败" }] }; }
  };
  server.registerTool("list_requirements", { description: "分页查询本机需求池；不返回图片内容、成员邮箱或 PPT 权限。",
    inputSchema: { query: z.string().optional(), offset: z.number().int().min(0).optional(), limit: z.number().int().min(1).max(100).optional() } },
    (args) => run("list_requirements", args));
  server.registerTool("get_requirement", { description: "按需求 ID 读取需求业务字段。", inputSchema: { id: z.string().min(1) } },
    (args) => run("get_requirement", args));
  server.registerTool("list_requirement_references", { description: "读取需求录入所需领域、产品及责任人成员 ID、姓名与角色，不返回邮箱。" },
    () => run("list_requirement_references", {}));
  server.registerTool("propose_requirement_changes", { description: "提出最多 25 项新增、修改、复制或删除需求；仅进入工作台人工确认，不直接写入。",
    inputSchema: { operations: z.array(operation).min(1).max(25) } },
    (args) => run("propose_requirement_changes", args));
  server.registerTool("get_proposal_status", { description: "查询提案待确认、已确认、取消、失效或冲突状态。", inputSchema: { id: z.string().min(1) } },
    (args) => run("get_proposal_status", args));
  return server;
}

export async function serveMcpStdio(userData: string): Promise<void> {
  const server = buildMcpServer(userData);
  await server.connect(new StdioServerTransport());
  console.error("solution-requirements-manager MCP stdio ready");
}
