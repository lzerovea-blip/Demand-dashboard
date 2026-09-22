import type {
  DictionaryItem,
  GroupOverride,
  Member,
  Requirement,
  WorkspaceData,
  WorkspaceFieldDifference,
  WorkspaceImportDifference,
  TieringChangeLog,
  TieringEvidenceSource,
  TieringFeature,
  TieringProductSeries,
  TieringSupportRecord,
  TieringVerificationIssue,
} from "./types.js";
import { normalizeTieringData } from "./tiering.js";
import { formatHistoricalLaunch, historicalLaunchStatusLabel } from "./tieringLaunches.js";
import { formatMedicalSubmissionMonth, medicalSubmissionTimeline } from "./medicalSubmission.js";

export function diffWorkspaceData(before: WorkspaceData, after: WorkspaceData): WorkspaceImportDifference[] {
  const differences: WorkspaceImportDifference[] = [];
  compareCollection(differences, "需求", before.requirements, after.requirements, (item) => item.id, requirementTitle, (left, right) => requirementFields(left, right, before.members ?? [], after.members ?? []));
  compareCollection(differences, "领域", before.domains, after.domains, (item) => item.id, (item) => item.name, dictionaryFields);
  compareCollection(differences, "产品", before.products, after.products, (item) => item.id, (item) => item.name, dictionaryFields);
  compareCollection(differences, "路标卡片", before.groupOverrides, after.groupOverrides, (item) => item.groupKey, (item) => item.cardTitle, overrideFields);
  compareCollection(differences, "成员", before.members ?? [], after.members ?? [], (item) => item.id, (item) => item.name, memberFields);
  const beforeTiering = normalizeTieringData(before.tiering);
  const afterTiering = normalizeTieringData(after.tiering);
  compareCollection(differences, "分层功能", beforeTiering.features, afterTiering.features, (item) => item.id, (item) => item.requirementTitle, tieringFeatureFields);
  compareCollection(differences, "产品系列", beforeTiering.productSeries, afterTiering.productSeries, (item) => item.id, (item) => item.name, tieringSeriesFields);
  compareCollection(differences, "系列支持", beforeTiering.supportRecords, afterTiering.supportRecords, (item) => item.id, (item) => item.id, tieringSupportFields);
  compareCollection(differences, "证据来源", beforeTiering.evidenceSources, afterTiering.evidenceSources, (item) => item.id, (item) => item.title, tieringSourceFields);
  compareCollection(differences, "待核验项", beforeTiering.verificationIssues, afterTiering.verificationIssues, (item) => item.id, (item) => item.description, tieringIssueFields);
  compareCollection(differences, "变更记录", beforeTiering.changeLogs, afterTiering.changeLogs, (item) => item.id, (item) => item.id, tieringLogFields);

  const beforePolicy = normalizePolicy(before);
  const afterPolicy = normalizePolicy(after);
  if (stable(beforePolicy) !== stable(afterPolicy)) {
    differences.push({
      entity: "PPT白名单",
      key: "ppt-access-policy",
      title: "PPT 成员白名单与权限",
      change: "修改",
      fields: changedFields([
        ["白名单", policyText(before), policyText(after)],
        ["更新时间", before.pptAccessPolicy?.updatedAt ?? "—", after.pptAccessPolicy?.updatedAt ?? "—"],
      ]),
    });
  }
  return differences;
}

export function workspaceDifferenceDigest(differences: readonly WorkspaceImportDifference[]): string {
  const input = stable(differences);
  let hash = 0x811c9dc5;
  for (let index = 0; index < input.length; index += 1) {
    hash ^= input.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return `diff-${(hash >>> 0).toString(16).padStart(8, "0")}-${differences.length}`;
}

function compareCollection<T>(
  target: WorkspaceImportDifference[],
  entity: WorkspaceImportDifference["entity"],
  beforeItems: readonly T[],
  afterItems: readonly T[],
  keyOf: (item: T) => string,
  titleOf: (item: T) => string,
  fieldsOf: (before: T, after: T) => WorkspaceFieldDifference[],
): void {
  const before = new Map(beforeItems.map((item) => [keyOf(item), item]));
  const after = new Map(afterItems.map((item) => [keyOf(item), item]));
  const keys = [...new Set([...before.keys(), ...after.keys()])].sort();
  for (const key of keys) {
    const current = before.get(key);
    const incoming = after.get(key);
    if (!current && incoming) {
      target.push({ entity, key, title: titleOf(incoming), change: "新增", fields: [{ field: "记录", currentValue: "—", incomingValue: summary(incoming) }] });
      continue;
    }
    if (current && !incoming) {
      target.push({ entity, key, title: titleOf(current), change: "删除", fields: [{ field: "记录", currentValue: summary(current), incomingValue: "—" }] });
      continue;
    }
    if (!current || !incoming) continue;
    const fields = fieldsOf(current, incoming);
    if (fields.length) target.push({ entity, key, title: titleOf(incoming), change: "修改", fields });
  }
}

function requirementTitle(item: Requirement): string {
  return item.title || item.id;
}

function requirementFields(before: Requirement, after: Requirement, beforeMembers: Member[], afterMembers: Member[]): WorkspaceFieldDifference[] {
  const memberName = (id: string | undefined, members: Member[]) => {
    const member = members.find((item) => item.id === id);
    return member ? `${member.name}（${member.employeeNumber}）` : (id || "未配置");
  };
  return changedFields([
    ["需求标题", before.title, after.title],
    ["需求描述", before.description, after.description],
    ["领域 L0", before.domainL0Id, after.domainL0Id],
    ["领域 L1", before.domainId, after.domainId],
    ["产品经理", memberName(before.productManagerMemberId, beforeMembers), memberName(after.productManagerMemberId, afterMembers)],
    ["规划", memberName(before.planningMemberId, beforeMembers), memberName(after.planningMemberId, afterMembers)],
    ["SE", memberName(before.seMemberId, beforeMembers), memberName(after.seMemberId, afterMembers)],
    ["算法", memberName(before.algorithmMemberId, beforeMembers), memberName(after.algorithmMemberId, afterMembers)],
    ["来源", before.source, after.source],
    ["适用范围", before.scope, after.scope],
    ["允许最近版本升级", yesNo(before.allowLatestVersionUpgrade), yesNo(after.allowLatestVersionUpgrade)],
    ["分类", before.category, after.category],
    ["上线月份", before.targetMonth, after.targetMonth],
    ["医疗认证时间节点", medicalTimelineText(before), medicalTimelineText(after)],
    ["医疗泳道顺序", before.medicalRoadmapSortOrder === undefined ? "默认顺序" : String(before.medicalRoadmapSortOrder + 1), after.medicalRoadmapSortOrder === undefined ? "默认顺序" : String(after.medicalRoadmapSortOrder + 1)],
    ["历史建议时间", launchText(before), launchText(after)],
    ["历史时间核验", historicalLaunchStatusLabel(before.historicalLaunch) || "—", historicalLaunchStatusLabel(after.historicalLaunch) || "—"],
    ["历史证据", before.historicalLaunch?.evidenceUrls.join("、") || "—", after.historicalLaunch?.evidenceUrls.join("、") || "—"],
    ["匹配产品", before.productIds.join("、") || "—", after.productIds.join("、") || "—"],
    ["设备工作量", String(before.deviceWorkloadPm), String(after.deviceWorkloadPm)],
    ["运动健康 App 工作量", String(before.appWorkloadPm), String(after.appWorkloadPm)],
    ["创新研究 App 工作量", String(before.researchAppWorkloadPm ?? 0), String(after.researchAppWorkloadPm ?? 0)],
    ["算法工作量", String(before.algorithmWorkloadPm ?? 0), String(after.algorithmWorkloadPm ?? 0)],
    ["云侧工作量", String(before.cloudWorkloadPm), String(after.cloudWorkloadPm)],
    ["待拆分工作量", String(before.unallocatedWorkloadPm), String(after.unallocatedWorkloadPm)],
    ["图片", imageText(before), imageText(after)],
    ["更新时间", before.updatedAt, after.updatedAt],
  ]);
}

function dictionaryFields(before: DictionaryItem, after: DictionaryItem): WorkspaceFieldDifference[] {
  return changedFields([
    ["名称", before.name, after.name],
    ["层级", before.level ?? "L1", after.level ?? "L1"],
    ["上游 L0", before.parentId ?? "—", after.parentId ?? "—"],
    ["状态", before.active ? "启用" : "停用", after.active ? "启用" : "停用"],
    ["顺序", String(before.sortOrder), String(after.sortOrder)],
  ]);
}

function overrideFields(before: GroupOverride, after: GroupOverride): WorkspaceFieldDifference[] {
  return changedFields([
    ["领域标签", before.cardTitle, after.cardTitle],
    ["摘要", before.cardSummary, after.cardSummary],
    ["顺序", String(before.sortOrder ?? "—"), String(after.sortOrder ?? "—")],
    ["更新时间", before.updatedAt, after.updatedAt],
  ]);
}

function memberFields(before: Member, after: Member): WorkspaceFieldDifference[] {
  return changedFields([
    ["姓名", before.name, after.name],
    ["工号", before.employeeNumber, after.employeeNumber],
    ["Microsoft 365 账号", before.microsoftAccount, after.microsoftAccount],
    ["岗位角色", before.role ?? "—", after.role ?? "—"],
    ["负责领域 L1", before.domainL1Ids.join("、") || "—", after.domainL1Ids.join("、") || "—"],
    ["更新时间", before.updatedAt, after.updatedAt],
  ]);
}

function tieringFeatureFields(before: TieringFeature, after: TieringFeature): WorkspaceFieldDifference[] {
  return changedFields([["L0", before.l0, after.l0], ["L1", before.l1, after.l1], ["需求", before.requirementTitle, after.requirementTitle], ["建议上线时间", featureLaunchText(before), featureLaunchText(after)], ["关联需求", before.requirementId ?? "—", after.requirementId ?? "—"]]);
}

function launchText(item: Requirement): string {
  return item.historicalLaunch ? `${formatHistoricalLaunch(item.historicalLaunch, item.targetMonth)} · 置信度 ${item.historicalLaunch.confidence}` : "—";
}

function featureLaunchText(item: TieringFeature): string {
  return item.historicalLaunch ? `${formatHistoricalLaunch(item.historicalLaunch)} · 置信度 ${item.historicalLaunch.confidence}` : "—";
}

function tieringSeriesFields(before: TieringProductSeries, after: TieringProductSeries): WorkspaceFieldDifference[] {
  return changedFields([["名称", before.name, after.name], ["系列边界", before.boundary, after.boundary], ["代表产品", before.representativeProducts, after.representativeProducts], ["生命周期", before.lifecycle, after.lifecycle]]);
}

function tieringSupportFields(before: TieringSupportRecord, after: TieringSupportRecord): WorkspaceFieldDifference[] {
  return changedFields([["支持状态", before.status, after.status]]);
}

function tieringSourceFields(before: TieringEvidenceSource, after: TieringEvidenceSource): WorkspaceFieldDifference[] {
  return changedFields([["来源标题", before.title, after.title], ["URL", before.url, after.url], ["适用范围", before.scope, after.scope], ["证据等级", before.evidenceLevel, after.evidenceLevel], ["维护状态", before.maintenanceStatus, after.maintenanceStatus]]);
}

function tieringIssueFields(before: TieringVerificationIssue, after: TieringVerificationIssue): WorkspaceFieldDifference[] {
  return changedFields([["问题描述", before.description, after.description], ["涉及系列", before.seriesIds.join("、"), after.seriesIds.join("、")], ["优先级", before.priority, after.priority], ["状态", before.status, after.status], ["负责人", before.owner, after.owner], ["处理说明", before.resolution, after.resolution]]);
}

function tieringLogFields(before: TieringChangeLog, after: TieringChangeLog): WorkspaceFieldDifference[] {
  return changedFields([
    ["变更时间", before.changedAt, after.changedAt],
    ["触发类型", before.triggerType, after.triggerType],
    ["变更前", before.before, after.before],
    ["变更后", before.after, after.after],
    ["变更原因", before.reason, after.reason],
    ["审核人", before.reviewer, after.reviewer],
    ["记录状态", before.recordStatus, after.recordStatus],
    ["备注", before.note, after.note],
    ["更新时间", before.updatedAt, after.updatedAt],
  ]);
}

function medicalTimelineText(item: Requirement): string {
  if (!item.medicalSubmissionMilestones) return "—";
  return medicalSubmissionTimeline(item.targetMonth, item.medicalSubmissionMilestones)
    .map((node) => `${node.label}:${node.month ? formatMedicalSubmissionMonth(node.month) : "未填写"}`)
    .join("；");
}

function changedFields(rows: Array<[string, string, string]>): WorkspaceFieldDifference[] {
  return rows.filter(([, currentValue, incomingValue]) => currentValue !== incomingValue)
    .map(([field, currentValue, incomingValue]) => ({ field, currentValue, incomingValue }));
}

function imageText(item: Requirement): string {
  return item.images.map((image) => `${image.name}(${image.sizeBytes})`).join("、") || "—";
}

function policyText(data: WorkspaceData): string {
  return normalizePolicy(data).map((item) => `${item.memberId}:${item.permission}`).join("、") || "—";
}

function normalizePolicy(data: WorkspaceData): Array<{ memberId: string; permission: string }> {
  return [...(data.pptAccessPolicy?.entries ?? [])].map((item) => ({ memberId: item.memberId, permission: item.permission })).sort((a, b) => a.memberId.localeCompare(b.memberId));
}

function yesNo(value: boolean): string {
  return value ? "是" : "否";
}

function summary(value: unknown): string {
  if (value && typeof value === "object" && "title" in value) return String((value as { title: unknown }).title);
  if (value && typeof value === "object" && "name" in value) return String((value as { name: unknown }).name);
  return stable(value);
}

function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(",")}]`;
  if (value && typeof value === "object") return `{${Object.entries(value as Record<string, unknown>).sort(([a], [b]) => a.localeCompare(b)).map(([key, item]) => `${JSON.stringify(key)}:${stable(item)}`).join(",")}}`;
  return JSON.stringify(value) ?? "";
}
