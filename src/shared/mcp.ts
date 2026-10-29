import type { Requirement, SaveRequirementInput } from "./types.js";
import type { McpDictionaryAddition } from "./mcpDictionaries.js";
import type { McpExportRequest, McpExportResult } from "./mcpExports.js";

/** MCP never accepts image bytes, roadmap assignments or tiering edits. */
export type McpRequirementFields = Omit<SaveRequirementInput, "id" | "images" | "tieringFeatureId" | "source" | "medicalSubmissionMilestones">;
export type McpRequirementPatch = Partial<McpRequirementFields>;
export type McpRequirementOperation =
  | { type: "create"; fields: McpRequirementFields }
  | { type: "update"; id: string; fields: McpRequirementPatch }
  | { type: "duplicate"; id: string }
  | { type: "delete"; id: string };

export interface McpFieldDifference { field: string; before: unknown; after: unknown }
export interface McpProposalItem {
  operation: McpRequirementOperation | { type: "create_dictionary"; entry: McpDictionaryAddition } | { type: "export"; request: McpExportRequest };
  title: string;
  differences: McpFieldDifference[];
  baselineUpdatedAt?: string;
}
export type McpProposalState = "pending" | "approved" | "cancelled" | "expired" | "conflict" | "failed" | "exporting";
export interface McpProposal {
  id: string;
  state: McpProposalState;
  createdAt: string;
  expiresAt: string;
  items: McpProposalItem[];
  error?: string;
  exportResult?: McpExportResult;
}
export interface McpConnectionStatus { enabled: boolean; pendingCount: number; }

export const MCP_PROPOSAL_TTL_MS = 30 * 60 * 1000;
export const MCP_TARGET_MONTH_PATTERN = /^\d{4}-(0[1-9]|1[0-2])$/;
export function validateMcpTargetMonth(fields: { targetMonth?: unknown }): void {
  if (fields.targetMonth !== undefined && (typeof fields.targetMonth !== "string" || !MCP_TARGET_MONTH_PATTERN.test(fields.targetMonth))) {
    throw new Error("上线年月格式错误：请填写 YYYY-MM，例如 2027-03；月份必须为 01–12，不接受中文日期、斜杠或时间戳");
  }
}
export const MCP_EDITABLE_FIELDS = [
  "title", "description", "domainL0Id", "domainId", "scope", "allowLatestVersionUpgrade",
  "overseasRegions", "category", "targetMonth", "productIds", "planningMemberId",
  "productManagerMemberId", "seMemberId", "algorithmMemberId", "deviceWorkloadPm",
  "appWorkloadPm", "cloudWorkloadPm", "researchAppWorkloadPm", "algorithmWorkloadPm",
] as const;

export function mcpRequirementView(requirement: Requirement): Record<string, unknown> {
  return {
    id: requirement.id,
    ...Object.fromEntries(MCP_EDITABLE_FIELDS.map((field) => [field, requirement[field]])),
    workloadPm: requirement.workloadPm,
    imageCount: requirement.images.length,
    createdAt: requirement.createdAt,
    updatedAt: requirement.updatedAt,
  };
}

export function mcpSaveInput(requirement: Requirement, patch: McpRequirementPatch): SaveRequirementInput {
  return {
    ...requirement,
    ...patch,
    id: requirement.id,
    planningMemberId: patch.planningMemberId ?? requirement.planningMemberId ?? "",
    productManagerMemberId: patch.productManagerMemberId ?? requirement.productManagerMemberId ?? "",
    seMemberId: patch.seMemberId ?? requirement.seMemberId ?? "",
    researchAppWorkloadPm: patch.researchAppWorkloadPm ?? requirement.researchAppWorkloadPm ?? 0,
    algorithmWorkloadPm: patch.algorithmWorkloadPm ?? requirement.algorithmWorkloadPm ?? 0,
    images: requirement.images,
    tieringFeatureId: requirement.tieringFeatureId,
    source: undefined,
    medicalSubmissionMilestones: undefined,
  };
}

export function mcpDifferences(before: Record<string, unknown>, after: Record<string, unknown>): McpFieldDifference[] {
  return Object.keys(after).filter((field) => JSON.stringify(before[field]) !== JSON.stringify(after[field]))
    .map((field) => ({ field, before: before[field], after: after[field] }));
}
