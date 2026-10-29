import { z } from "zod";
import { MEMBER_ROLES, type AppSnapshot } from "./types.js";

const name = z.string().trim().min(1).max(200);
export const mcpDictionaryAdditionSchema = z.discriminatedUnion("kind", [
  z.object({ kind: z.literal("domain_l0"), name }).strict(),
  z.object({ kind: z.literal("domain_l1"), name, parentId: z.string().min(1) }).strict(),
  z.object({ kind: z.literal("product"), name }).strict(),
  z.object({ kind: z.literal("member"), name, employeeNumber: name, microsoftAccount: z.string().trim().email(), role: z.enum(MEMBER_ROLES).optional(), domainL1Ids: z.array(z.string().min(1)).default([]) }).strict(),
  z.object({ kind: z.literal("roadmap"), name, roadmapKind: z.enum(["standard", "medical"]) }).strict(),
  z.object({ kind: z.literal("lane"), name, roadmapId: z.string().min(1) }).strict(),
]);
export const mcpDictionaryBatchSchema = z.array(mcpDictionaryAdditionSchema).min(1).max(25);
export type McpDictionaryAddition = z.infer<typeof mcpDictionaryAdditionSchema>;
export const MCP_DICTIONARY_NAMES = { domain_l0: "领域 L0", domain_l1: "领域 L1", product: "产品", member: "成员", roadmap: "路标", lane: "泳道" };

/** Add-only: never accept IDs to replace existing records or infer parents by name. */
export function validateMcpDictionaryAdditions(snapshot: AppSnapshot, input: unknown): McpDictionaryAddition[] {
  const additions = mcpDictionaryBatchSchema.parse(input);
  const seen = new Set<string>();
  const key = (text: string) => text.trim().toLowerCase();
  for (const item of additions) {
    const identity = `${item.kind}:${item.kind === "lane" ? item.roadmapId : ""}:${key(item.kind === "member" ? item.employeeNumber : item.name)}`;
    if (seen.has(identity)) throw new Error("同一批次含重复字典条目");
    seen.add(identity);
    if (item.kind === "domain_l0" || item.kind === "domain_l1") {
      const level = item.kind === "domain_l0" ? "L0" : "L1";
      if (snapshot.domains.some(value => (value.level ?? "L1") === level && key(value.name) === key(item.name))) throw new Error("领域名称已存在，不允许通过新增覆盖");
      if (item.kind === "domain_l1" && !snapshot.domains.some(value => value.id === item.parentId && value.level === "L0")) throw new Error("请选择已有的领域 L0 ID");
    } else if (item.kind === "product") {
      if (snapshot.products.some(value => key(value.name) === key(item.name))) throw new Error("产品名称已存在");
    } else if (item.kind === "member") {
      const accountKey = `account:${key(item.microsoftAccount)}`;
      if (seen.has(accountKey) || snapshot.members.some(value => key(value.employeeNumber) === key(item.employeeNumber) || key(value.microsoftAccount) === key(item.microsoftAccount))) throw new Error("成员工号或企业邮箱已存在");
      seen.add(accountKey);
      if (item.domainL1Ids.some(id => !snapshot.domains.some(value => value.id === id && (value.level ?? "L1") === "L1"))) throw new Error("负责领域须为已有的领域 L1 ID");
    } else if (item.kind === "roadmap") {
      if ((snapshot.roadmaps ?? []).some(value => key(value.name) === key(item.name))) throw new Error("路标名称已存在");
    } else {
      if (!(snapshot.roadmaps ?? []).some(value => value.id === item.roadmapId)) throw new Error("请选择已有的路标 ID");
      if ((snapshot.roadmapLanes ?? []).some(value => value.roadmapId === item.roadmapId && key(value.name) === key(item.name))) throw new Error("该路标下泳道名称已存在");
    }
  }
  return additions;
}
