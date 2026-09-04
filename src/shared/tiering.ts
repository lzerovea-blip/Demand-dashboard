import type {
  DictionaryItem,
  Member,
  Requirement,
  RequirementSource,
  TieringData,
  TieringEvidenceSource,
  TieringFeature,
  TieringProductSeries,
  TieringSupportRecord,
  TieringVerificationIssue,
} from "./types.js";
import { TIERING_AUDIT_STATUSES, TIERING_SUPPORT_STATUSES } from "./types.js";
import {
  formatHistoricalLaunch,
  historicalLaunchForFeatureId,
  normalizeHistoricalLaunchMetadata,
} from "./tieringLaunches.js";

export const TIERING_BASELINE_DATE = "2026-08-05";
export const TIERING_EXPECTED_COUNTS = {
  features: 156,
  productSeries: 9,
  supportRecords: 1404,
  evidenceSources: 44,
  verificationIssues: 96,
} as const;

/**
 * 产品系列属于应用基础字典。它们必须随程序存在，不能依赖某台电脑
 * 曾经导入过分层分级数据包。
 */
const SERIES_ORDER = ["PS01", "PS02", "PS11", "PS03", "PS04", "PS10", "PS05", "PS06", "PS07", "PS08", "PS09", "PS12", "PS13"];
export const BUILT_IN_TIERING_PRODUCT_SERIES: readonly TieringProductSeries[] = [
  { id: "PS01", name: "WATCH数字系列", boundary: "WATCH数字命名系列", representativeProducts: "WATCH 3/4/5、WATCH十周年等", managementEntry: "华为运动健康", sourceIds: ["S043"], lifecycle: "现网/历史并存", note: "系列内部代际差异较大；以适用产品清单为准", sortOrder: 0, updatedAt: "2026-08-05T00:00:00.000Z" },
  { id: "PS02", name: "GT系列", boundary: "长续航时尚与专业运动系列", representativeProducts: "GT、GT 2-6、GT Runner、GT Cyber等", managementEntry: "华为运动健康", sourceIds: ["S029"], lifecycle: "现网/历史并存", note: "Pro、Runner与标准款能力有差异", sortOrder: 1, updatedAt: "2026-08-05T00:00:00.000Z" },
  { id: "PS03", name: "ULTIMATE系列", boundary: "高端户外与潜水系列", representativeProducts: "WATCH Ultimate、Ultimate 2、Ultimate Design", managementEntry: "华为运动健康", sourceIds: ["S037"], lifecycle: "现网", note: "高尔夫、潜水和户外能力有明显型号差异", sortOrder: 2, updatedAt: "2026-08-05T00:00:00.000Z" },
  { id: "PS04", name: "FIT系列", boundary: "轻薄运动手表系列", representativeProducts: "WATCH FIT、FIT 2-5、Pro/New/Mini", managementEntry: "华为运动健康", sourceIds: ["S025"], lifecycle: "现网/历史并存", note: "Pro与标准款传感器、专项运动能力不同", sortOrder: 3, updatedAt: "2026-08-05T00:00:00.000Z" },
  { id: "PS05", name: "D系列", boundary: "腕部血压健康手表系列", representativeProducts: "WATCH D、WATCH D2", managementEntry: "华为运动健康", sourceIds: ["S014"], lifecycle: "现网/历史并存", note: "D与D2动态血压能力不同", sortOrder: 4, updatedAt: "2026-08-05T00:00:00.000Z" },
  { id: "PS06", name: "WATCH Buds系列", boundary: "耳机与手表融合系列", representativeProducts: "WATCH Buds、WATCH Buds 2", managementEntry: "华为运动健康", sourceIds: ["S006"], lifecycle: "现网/历史并存", note: "Buds 2新增能力需逐项核验", sortOrder: 5, updatedAt: "2026-08-05T00:00:00.000Z" },
  { id: "PS07", name: "Band系列", boundary: "智能手环及通话手环系列", representativeProducts: "Band 6-11、B5/B6/B7等", managementEntry: "华为运动健康", sourceIds: ["S003"], lifecycle: "现网/历史并存", note: "手环代际和标准/NFC/Pro款差异显著", sortOrder: 6, updatedAt: "2026-08-05T00:00:00.000Z" },
  { id: "PS08", name: "儿童手表", boundary: "儿童与青少年手表系列", representativeProducts: "儿童手表4/5系列、超新星手表X1等", managementEntry: "智能关怀为主", sourceIds: ["S042"], lifecycle: "现网/历史并存", note: "非华为运动健康主入口；矩阵仅标设备数据能力关联", sortOrder: 7, updatedAt: "2026-08-05T00:00:00.000Z" },
  { id: "PS09", name: "体脂秤", boundary: "华为体重与身体成分设备", representativeProducts: "智能体脂秤、2 Pro、3、3 Pro", managementEntry: "华为运动健康", sourceIds: ["S015"], lifecycle: "现网/历史并存", note: "普通版、蓝牙/WLAN版和Pro版指标不同", sortOrder: 8, updatedAt: "2026-08-05T00:00:00.000Z" },
  ...([
    ["PS10", "FIT Pro系列"], ["PS11", "GT Pro系列"], ["PS12", "戒指系列"], ["PS13", "无屏系列"],
  ] as const).map(([id, name]): TieringProductSeries => ({ id, name, boundary: "", representativeProducts: "", managementEntry: "", sourceIds: [], lifecycle: "", note: "新增系列，支持规格独立判断，不继承其他系列结论", sortOrder: 0, updatedAt: "2026-09-04T00:00:00.000Z" })),
].map((item) => ({ ...item, sortOrder: SERIES_ORDER.indexOf(item.id) })).sort((a, b) => a.sortOrder - b.sortOrder);

export function missingBuiltInTieringProductSeries(existing: readonly Pick<TieringProductSeries, "id">[]): TieringProductSeries[] {
  const existingIds = new Set(existing.map((item) => item.id));
  return BUILT_IN_TIERING_PRODUCT_SERIES
    .filter((item) => !existingIds.has(item.id))
    .map((item) => ({ ...item, sourceIds: [...item.sourceIds] }));
}

export const EMPTY_TIERING_DATA: TieringData = {
  features: [],
  productSeries: [],
  supportRecords: [],
  evidenceSources: [],
  verificationIssues: [],
  changeLogs: [],
};

const SOURCE_BY_L0: Record<string, RequirementSource> = {
  基础体征测量: "健康基础",
  休息放松: "健康基础",
  健康报告: "健康进阶",
  女性健康: "健康进阶",
  智能关爱: "健康进阶",
  华为创新研究app: "健康高阶",
  慢病管理: "健康高阶",
  保持活力: "运动基础",
  专业运动: "运动高阶",
};

export function tieringSourceForL0(l0: string): RequirementSource {
  const source = SOURCE_BY_L0[l0.trim()];
  if (!source) throw new Error(`分层功能 L0“${l0}”缺少需求来源映射`);
  return source;
}

export function tieringFeatureIdForRequirement(requirementId: string): string {
  return `REQ-${requirementId}`;
}

export function resolveTieringRequirementDomains(
  requirement: Pick<Requirement, "domainL0Id" | "domainId">,
  domains: readonly DictionaryItem[],
): { domainL0: DictionaryItem; domainL1: DictionaryItem } | undefined {
  const domainsById = new Map(domains.map((item) => [item.id, item]));
  const domainL1 = domainsById.get(requirement.domainId);
  const domainL0 = domainsById.get(domainL1?.parentId ?? requirement.domainL0Id);
  if (!domainL1 || domainL0?.level !== "L0" || (domainL1.level ?? "L1") !== "L1" || domainL1.parentId !== domainL0.id) {
    return undefined;
  }
  return { domainL0, domainL1 };
}

/** A missing or ambiguous link is retained for repair, never matched by title/name. */
export function resolveTieringFeature(feature: TieringFeature, requirements: readonly Requirement[], domains: readonly DictionaryItem[]) {
  const candidates = requirements.filter((item) => item.id === feature.requirementId || item.tieringFeatureId === feature.id);
  const requirement = candidates.length === 1 ? candidates[0] : undefined;
  const domain = requirement && resolveTieringRequirementDomains(requirement, domains);
  const invalidLink = candidates.length > 1 || Boolean(feature.requirementId && !requirement) || Boolean(requirement && !domain)
    || Boolean(requirement?.tieringFeatureId && requirement.tieringFeatureId !== feature.id);
  return {
    requirement,
    invalidLink,
    feature: !invalidLink && requirement && domain ? { ...feature, l0: domain.domainL0.name, l1: domain.domainL1.name, requirementTitle: requirement.title } : feature,
  };
}

/** Only unlinked specifications use a name path; linked rows are protected by demand IDs. */
export function domainReferenceMessages(domain: DictionaryItem, snapshot: { domains: readonly DictionaryItem[]; requirements: readonly Requirement[]; members: readonly Member[]; features: readonly TieringFeature[] }): string[] {
  const { domains, requirements, members, features } = snapshot;
  const children = domains.filter((item) => item.parentId === domain.id);
  const demands = requirements.filter((item) => domain.level === "L0"
    ? (resolveTieringRequirementDomains(item, domains)?.domainL0.id ?? item.domainL0Id) === domain.id
    : item.domainId === domain.id);
  const parentName = domains.find((item) => item.id === domain.parentId)?.name;
  const independent = features.filter((item) => {
    const resolved = resolveTieringFeature(item, requirements, domains);
    if (resolved.requirement && !resolved.invalidLink) return false;
    return domain.level === "L0" ? item.l0 === domain.name : item.l1 === domain.name && item.l0 === parentName;
  });
  const owners = members.filter((item) => item.domainL1Ids.includes(domain.id));
  const messages: string[] = [];
  const add = (names: string[], label: string, action: string) => {
    if (names.length) messages.push(`${label} ${names.length} 项：${names.join("、")}。${action}`);
  };
  add(children.map((item) => item.name), "子级领域 L1", "请在字典设置迁移或删除子级领域");
  add(demands.map((item) => item.title), "需求引用", "请在需求池调整领域");
  add(independent.map((item) => `${item.requirementTitle}（${item.id}）`), "独立规格或异常关联", "请在分层分级恢复需求关联后调整领域");
  add(owners.map((item) => item.name), "负责成员", "请在人员管理解除负责领域关联");
  return messages;
}

export function renameIndependentTieringDomain(feature: TieringFeature, previous: DictionaryItem, next: DictionaryItem, domains: readonly DictionaryItem[], requirements: readonly Requirement[]): TieringFeature {
  // A dangling link is not a confirmed independent historical specification.
  if (feature.requirementId || resolveTieringFeature(feature, requirements, domains).requirement) return feature;
  if (previous.level === "L0" && feature.l0 === previous.name) return { ...feature, l0: next.name };
  const oldParent = domains.find((item) => item.id === previous.parentId);
  const newParent = domains.find((item) => item.id === next.parentId);
  if ((previous.level ?? "L1") === "L1" && oldParent && newParent && feature.l0 === oldParent.name && feature.l1 === previous.name) {
    return { ...feature, l0: newParent.name, l1: next.name };
  }
  return feature;
}

export function isHistoricalTieringFeature(feature: Pick<TieringFeature, "id">): boolean {
  return /^F\d{3}$/.test(feature.id);
}

export function tieringSupportId(featureId: string, seriesId: string): string {
  return `${featureId}::${seriesId}`;
}

export function historicalRequirementId(featureId: string): string {
  return `historical-huawei-${featureId}`;
}

export function normalizeTieringData(raw?: Partial<TieringData> | null): TieringData {
  if (!raw) return cloneTieringData(EMPTY_TIERING_DATA);
  return {
    features: (raw.features ?? []).map(normalizeFeature),
    productSeries: (raw.productSeries ?? []).map(normalizeSeries),
    supportRecords: (raw.supportRecords ?? []).map(normalizeSupport),
    evidenceSources: (raw.evidenceSources ?? []).map(normalizeSource),
    verificationIssues: (raw.verificationIssues ?? []).map(normalizeIssue),
    changeLogs: (raw.changeLogs ?? []).map((item) => ({
      ...item,
      sourceIds: [...new Set(item.sourceIds ?? [])],
      note: item.note ?? "",
    })),
  };
}

export function validateTieringData(raw?: Partial<TieringData> | null): TieringData {
  const data = normalizeTieringData(raw);
  unique(data.features, "分层功能");
  unique(data.productSeries, "产品系列");
  unique(data.supportRecords, "系列支持");
  unique(data.evidenceSources, "证据来源");
  unique(data.verificationIssues, "待核验项");
  unique(data.changeLogs, "变更记录");
  const featureIds = new Set(data.features.map((item) => item.id));
  const seriesIds = new Set(data.productSeries.map((item) => item.id));
  const sourceIds = new Set(data.evidenceSources.map((item) => item.id));
  for (const feature of data.features) {
    if (!feature.id || !feature.l0 || !feature.l1 || !feature.requirementTitle) throw new Error("分层功能存在空层级或无效 ID");
  }
  for (const record of data.supportRecords) {
    if (!featureIds.has(record.featureId) || !seriesIds.has(record.seriesId)) throw new Error(`系列支持“${record.id}”引用了不存在的功能或系列`);
    if (record.id !== tieringSupportId(record.featureId, record.seriesId)) throw new Error(`系列支持“${record.id}”的稳定 ID 无效`);
    validateTieringSupportRecord(record);
    if (record.sourceIds.some((id) => !sourceIds.has(id))) throw new Error(`系列支持“${record.id}”引用了不存在的证据来源`);
  }
  for (const issue of data.verificationIssues) {
    if (!featureIds.has(issue.featureId) || issue.seriesIds.some((id) => !seriesIds.has(id))) throw new Error(`待核验项“${issue.id}”引用了不存在的功能或系列`);
    if (issue.sourceIds.some((id) => !sourceIds.has(id))) throw new Error(`待核验项“${issue.id}”引用了不存在的证据来源`);
  }
  return data;
}

export function validateTieringSupportRecord(record: TieringSupportRecord): void {
  if (!TIERING_SUPPORT_STATUSES.includes(record.status)) throw new Error("请选择有效的系列支持状态");
  if (!TIERING_AUDIT_STATUSES.includes(record.auditStatus)) throw new Error("请选择有效的审核状态");
}

export function createTieringSupportRecord(featureId: string, seriesId: string, updatedAt: string): TieringSupportRecord {
  return {
    id: tieringSupportId(featureId, seriesId),
    featureId,
    seriesId,
    status: "未判断",
    supportedModels: "",
    excludedModels: "",
    phoneOsMarket: "",
    minHealthAppVersion: "",
    minDeviceVersion: "",
    prerequisites: "",
    sourceIds: [],
    sourceUrl: "",
    verifiedAt: "",
    evidenceLevel: "",
    auditStatus: "待内审",
    note: "",
    updatedAt,
  };
}

export function buildTieringFeatureForRequirement(
  requirement: Requirement,
  domainL0Name: string,
  domainL1Name: string,
  sortOrder: number,
  updatedAt: string,
  existing?: TieringFeature,
): TieringFeature {
  return {
    id: existing?.id ?? requirement.tieringFeatureId ?? tieringFeatureIdForRequirement(requirement.id),
    l0: domainL0Name.trim(),
    l1: domainL1Name.trim(),
    requirementTitle: requirement.title.trim(),
    requirementId: requirement.id,
    ...(existing?.historicalLaunch ? { historicalLaunch: existing.historicalLaunch } : {}),
    sortOrder: existing?.sortOrder ?? sortOrder,
    createdAt: existing?.createdAt ?? updatedAt,
    updatedAt,
  };
}

export function buildHistoricalRequirement(
  feature: TieringFeature,
  supportRecords: readonly TieringSupportRecord[],
  productSeries: readonly TieringProductSeries[],
  domainL0: DictionaryItem,
  domainL1: DictionaryItem,
  existing?: Requirement,
): Requirement {
  const now = existing?.updatedAt ?? `${TIERING_BASELINE_DATE}T00:00:00.000Z`;
  const historicalLaunch = normalizeHistoricalLaunchMetadata(feature.historicalLaunch)
    ?? historicalLaunchForFeatureId(feature.id);
  const statusCounts = new Map<string, number>();
  for (const record of supportRecords.filter((item) => item.featureId === feature.id)) statusCounts.set(record.status, (statusCounts.get(record.status) ?? 0) + 1);
  const summary = [...statusCounts.entries()].map(([status, count]) => `${status}${count}个系列`).join("、") || `${productSeries.length}个系列待核验`;
  return {
    id: historicalRequirementId(feature.id),
    title: feature.requirementTitle,
    description: [
      `分层路径：${feature.l0} → ${feature.l1} → ${feature.requirementTitle}`,
      `历史现网基线：${TIERING_BASELINE_DATE} 中国大陆现网。`,
      `系列支持摘要：${summary}；完整规格请在“分层分级”查看。`,
      `建议上线时间：${formatHistoricalLaunch(historicalLaunch)}。`,
      `时间证据：置信度 ${historicalLaunch.confidence}；${historicalLaunch.note}`,
      historicalLaunch.precision === "month"
        ? "该月份按中国大陆首次公开可用时间记录。"
        : `系统占位月份为 ${historicalLaunch.suggestedTargetMonth}；仅用于半年时间轴归档，不代表真实首发月份。`,
    ].join("\n"),
    images: [],
    domainL0Id: domainL0.id,
    domainId: domainL1.id,
    source: tieringSourceForL0(feature.l0),
    scope: "仅国内",
    allowLatestVersionUpgrade: true,
    overseasRegions: [],
    category: "平台演进",
    targetMonth: historicalLaunch.suggestedTargetMonth,
    productIds: [],
    deviceWorkloadPm: 0,
    appWorkloadPm: 0,
    cloudWorkloadPm: 0,
    unallocatedWorkloadPm: 0,
    workloadPm: 0,
    createdAt: existing?.createdAt ?? now,
    updatedAt: now,
    tieringFeatureId: feature.id,
    historicalLaunch,
  };
}

export function migrateArchivedHistoricalRequirement(requirement: Requirement, feature: TieringFeature): Requirement {
  if (!isHistoricalTieringFeature(feature) || requirement.historicalLaunch || requirement.targetMonth !== "2026-09") return requirement;
  const historicalLaunch = normalizeHistoricalLaunchMetadata(feature.historicalLaunch)
    ?? historicalLaunchForFeatureId(feature.id);
  const descriptionLines = requirement.description
    .split("\n")
    .filter((line) => !line.startsWith("归档月份：") && !line.startsWith("建议上线时间：") && !line.startsWith("时间证据：") && !line.startsWith("系统占位月份为 ") && line !== "该月份按中国大陆首次公开可用时间记录。");
  return {
    ...requirement,
    targetMonth: historicalLaunch.suggestedTargetMonth,
    historicalLaunch,
    description: [
      ...descriptionLines,
      `建议上线时间：${formatHistoricalLaunch(historicalLaunch)}。`,
      `时间证据：置信度 ${historicalLaunch.confidence}；${historicalLaunch.note}`,
      historicalLaunch.precision === "month"
        ? "该月份按中国大陆首次公开可用时间记录。"
        : `系统占位月份为 ${historicalLaunch.suggestedTargetMonth}；仅用于半年时间轴归档，不代表真实首发月份。`,
    ].join("\n"),
  };
}

export function migrateTieringRequirementDescription(requirement: Requirement, feature: TieringFeature): Requirement {
  if (!requirement.historicalLaunch && !isHistoricalTieringFeature(feature)) return requirement;
  const path = `分层路径：${feature.l0} → ${feature.l1} → ${feature.requirementTitle}`;
  const lines = requirement.description.split("\n");
  const pathIndex = lines.findIndex((line) => line.startsWith("分层路径："));
  if (pathIndex >= 0) lines[pathIndex] = path;
  else lines.unshift(path);
  const description = lines.join("\n").replaceAll("需求分层分级管理", "分层分级");
  return description === requirement.description ? requirement : { ...requirement, description };
}

export function mergeTieringData(localRaw: TieringData, incomingRaw: TieringData): TieringData {
  const local = normalizeTieringData(localRaw);
  const incoming = validateTieringData(incomingRaw);
  return {
    features: mergeByTimestamp(local.features, incoming.features),
    productSeries: mergeByTimestamp(local.productSeries, incoming.productSeries),
    supportRecords: mergeByTimestamp(local.supportRecords, incoming.supportRecords),
    evidenceSources: mergeByTimestamp(local.evidenceSources, incoming.evidenceSources),
    verificationIssues: mergeByTimestamp(local.verificationIssues, incoming.verificationIssues),
    changeLogs: mergeByTimestamp(local.changeLogs, incoming.changeLogs, "changedAt"),
  };
}

export function cloneTieringData(data: TieringData): TieringData {
  return JSON.parse(JSON.stringify(data)) as TieringData;
}

function normalizeFeature(item: TieringFeature): TieringFeature {
  const legacy = item as TieringFeature & { l2?: string; l3?: string; l4?: string };
  const usesLegacyShape = !item.l0 && Boolean(legacy.l2 || legacy.l3 || legacy.l4);
  const historicalLaunch = normalizeHistoricalLaunchMetadata(item.historicalLaunch)
    ?? (/^F\d{3}$/.test(item.id) ? historicalLaunchForFeatureId(item.id) : undefined);
  return {
    id: item.id,
    l0: (usesLegacyShape ? legacy.l1 : item.l0) ?? "",
    l1: (usesLegacyShape ? legacy.l2 : item.l1) ?? "",
    requirementTitle: (usesLegacyShape ? legacy.l4 : item.requirementTitle) ?? "",
    requirementId: item.requirementId || undefined,
    ...(historicalLaunch ? { historicalLaunch } : {}),
    sortOrder: Number.isInteger(item.sortOrder) ? item.sortOrder : 0,
    createdAt: item.createdAt || item.updatedAt || new Date(0).toISOString(),
    updatedAt: item.updatedAt || item.createdAt || new Date(0).toISOString(),
  };
}

function normalizeSeries(item: TieringProductSeries): TieringProductSeries {
  return { ...item, sourceIds: [...new Set(item.sourceIds ?? [])], sortOrder: Number.isInteger(item.sortOrder) ? item.sortOrder : 0 };
}

function normalizeSupport(item: TieringSupportRecord): TieringSupportRecord {
  return {
    ...item,
    id: tieringSupportId(item.featureId, item.seriesId),
    status: normalizeTieringSupportStatus(String(item.status)),
    sourceIds: [...new Set(item.sourceIds ?? [])],
    supportedModels: item.supportedModels ?? "",
    excludedModels: item.excludedModels ?? "",
    phoneOsMarket: item.phoneOsMarket ?? "",
    minHealthAppVersion: item.minHealthAppVersion ?? "",
    minDeviceVersion: item.minDeviceVersion ?? "",
    prerequisites: item.prerequisites ?? "",
    sourceUrl: item.sourceUrl ?? "",
    verifiedAt: item.verifiedAt ?? "",
    evidenceLevel: item.evidenceLevel ?? "",
    auditStatus: item.auditStatus ?? "待内审",
    note: item.note ?? "",
  };
}

export function normalizeTieringSupportStatus(status: string): TieringSupportRecord["status"] {
  if (status === "支持" || status === "部分支持" || status === "条件支持" || status === "部分且条件") return "支持";
  if (status === "不支持" || status === "不适用") return "不支持";
  return "未判断";
}

function normalizeSource(item: TieringEvidenceSource): TieringEvidenceSource {
  return { ...item, owner: item.owner ?? "", note: item.note ?? "", updatedAt: item.updatedAt || item.verifiedAt || item.collectedAt || new Date(0).toISOString() };
}

function normalizeIssue(item: TieringVerificationIssue): TieringVerificationIssue {
  return { ...item, seriesIds: [...new Set(item.seriesIds ?? [])], sourceIds: [...new Set(item.sourceIds ?? [])], owner: item.owner ?? "", targetDate: item.targetDate ?? "", resolution: item.resolution ?? "" };
}

function unique(items: readonly { id: string }[], label: string): void {
  if (items.some((item) => !item.id) || new Set(items.map((item) => item.id)).size !== items.length) throw new Error(`${label}包含空 ID 或重复 ID`);
}

function mergeByTimestamp<T extends { id: string }>(local: T[], incoming: T[], timestampField: keyof T = "updatedAt" as keyof T): T[] {
  const result = new Map(local.map((item) => [item.id, item]));
  for (const item of incoming) {
    const existing = result.get(item.id);
    const incomingTime = Date.parse(String(item[timestampField] ?? ""));
    const existingTime = Date.parse(String(existing?.[timestampField] ?? ""));
    if (!existing || !Number.isFinite(existingTime) || incomingTime >= existingTime) result.set(item.id, item);
  }
  return [...result.values()];
}
