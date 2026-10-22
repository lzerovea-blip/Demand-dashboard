import type { Requirement, SaveRequirementInput } from "./types.js";

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
  operation: McpRequirementOperation;
  title: string;
  differences: McpFieldDifference[];
  baselineUpdatedAt?: string;
}
export type McpProposalState = "pending" | "approved" | "cancelled" | "expired" | "conflict" | "failed";
export interface McpProposal {
  id: string;
  state: McpProposalState;
  createdAt: string;
  expiresAt: string;
  items: McpProposalItem[];
  error?: string;
}
export interface McpConnectionStatus { enabled: boolean; pendingCount: number; }

export const MCP_PROPOSAL_TTL_MS = 30 * 60 * 1000;
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
