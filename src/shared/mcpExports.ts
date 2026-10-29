import { z } from "zod";
import type { AppSnapshot, PptExportRequest } from "./types.js";
import { halfYearSequence, requirementsInRange } from "./roadmap.js";
import { buildPptExportPlan } from "./pptExport.js";
import { isTieringEligibleRequirement } from "./tiering.js";

const halfYear = z.string().regex(/^\d{4}H[12]$/, "半年格式须为 YYYYH1 或 YYYYH2，例如 2027H1");
const range = { start: halfYear, end: halfYear };
export const mcpExportRequestSchema = z.discriminatedUnion("format", [
  z.object({ format: z.literal("workspace") }).strict(),
  z.object({ format: z.literal("requirements_excel"), ...range }).strict(),
  z.object({ format: z.literal("tiering_excel"), ...range }).strict(),
  z.object({ format: z.literal("ppt"), ...range, materialName: z.string().trim().min(1).max(200).optional(), includeSummary: z.boolean().default(true), roadmapIds: z.array(z.string().min(1)).max(100).optional() }).strict(),
]);
export type McpExportRequest = z.infer<typeof mcpExportRequestSchema>;
export interface McpExportResult { canceled: boolean; path?: string; slideCount?: number }
export const MCP_EXPORT_NAMES = { workspace: "完整需求包", requirements_excel: "需求池 Excel", tiering_excel: "分层分级 Excel", ppt: "受保护 PPT" };

export function mcpPptRequest(snapshot: AppSnapshot, request: Extract<McpExportRequest, { format: "ppt" }>): PptExportRequest {
  const ids = request.roadmapIds ?? (snapshot.roadmaps ?? []).map(item => item.id);
  if (ids.some(id => !(snapshot.roadmaps ?? []).some(item => item.id === id))) throw new Error("所选路标 ID 已不存在，请重新读取字典");
  return { start: request.start, end: request.end, materialName: request.materialName ?? "解决方案需求路标", sections: [...(request.includeSummary ? ["summary"] : []), ...new Set(ids)] };
}

export function validateMcpExport(snapshot: AppSnapshot, input: unknown): McpExportRequest {
  const request = mcpExportRequestSchema.parse(input);
  if (request.format !== "workspace") halfYearSequence(request.start, request.end);
  if (request.format === "ppt") buildPptExportPlan(snapshot, mcpPptRequest(snapshot, request));
  return request;
}

export function mcpExportPreview(snapshot: AppSnapshot, request: McpExportRequest): Record<string, unknown> {
  if (request.format === "workspace") return { exportFormat: MCP_EXPORT_NAMES.workspace, exportRange: "全部数据（不按时间过滤）", exportCount: snapshot.requirements.length, exportDraftCount: snapshot.draftRequirements?.length ?? 0, exportWarning: "完整包包含需求和草稿、图片、人员工号与邮箱、PPT白名单、字典、路标配置、分层数据及删除记录，请仅交给可信接收人。不会自动发送或上传。" };
  const selected = requirementsInRange(snapshot.requirements, request.start, request.end);
  const preview: Record<string, unknown> = { exportFormat: MCP_EXPORT_NAMES[request.format], exportRange: `${request.start} — ${request.end}`, exportCount: request.format === "tiering_excel" ? selected.filter(isTieringEligibleRequirement).length : selected.length };
  if (request.format === "requirements_excel") preview.exportDraftCount = requirementsInRange(snapshot.draftRequirements ?? [], request.start, request.end).length;
  if (request.format === "ppt") {
    const ppt = mcpPptRequest(snapshot, request);
    preview.materialName = ppt.materialName;
    preview.exportSections = ppt.sections.map(id => id === "summary" ? "需求工作量汇总分析" : snapshot.roadmaps!.find(item => item.id === id)!.name).join("、") || "仅封面";
    preview.exportWarning = "沿用现有国内/海外分页、续页及医疗独立时间轴规则。只允许企业 IRM 保护导出，使用当前已保存白名单及半年有效期；不接受模型指定账号或降低保护。";
  } else preview.exportWarning = "Excel 不设置密码保护，按现有表格结构导出；不包含图片。";
  return preview;
}
