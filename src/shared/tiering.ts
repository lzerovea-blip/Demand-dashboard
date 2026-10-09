import type {
  DictionaryItem,
  Member,
  Requirement,
  RequirementTombstone,
  RequirementSource,
  TieringData,
  TieringEvidenceSource,
  TieringFeature,
  TieringProductSeries,
  TieringSupportRecord,
  TieringVerificationIssue,
  WorkspaceData,
} from "./types.js";
import { TIERING_AUDIT_STATUSES, TIERING_SUPPORT_STATUSES } from "./types.js";
import {
  formatHistoricalLaunch,
  historicalLaunchForFeatureId,
  normalizeHistoricalLaunchMetadata,
} from "./tieringLaunches.js";

export const TIERING_BASELINE_DATE = "2026-08-05";
export const TIERING_ELIGIBLE_SOURCES = [
  "运动基础",
  "运动进阶",
  "运动高阶",
  "健康基础",
  "健康进阶",
  "健康高阶",
  "AI运动",
  "AI健康",
  "AI综合",
] as const satisfies readonly RequirementSource[];

const TIERING_ELIGIBLE_SOURCE_SET = new Set<RequirementSource>(TIERING_ELIGIBLE_SOURCES);

/**
 * 仅用于旧协作表格的九来源布局兼容；新版分层分级管理全部需求。
 */
export function isTieringEligibleSource(source: RequirementSource | undefined): boolean {
  return Boolean(source && TIERING_ELIGIBLE_SOURCE_SET.has(source));
}

export function isTieringEligibleRequirement(_requirement: Pick<Requirement, "source">): boolean {
  return true;
}

export const TIERING_EXPECTED_COUNTS = {
  features: 156,
  productSeries: 15,
  supportRecords: 2340,
  evidenceSources: 44,
  verificationIssues: 96,
} as const;

/**
 * 产品系列属于应用基础字典。它们必须随程序存在，不能依赖某台电脑
 * 曾经导入过分层分级数据包。
 */
const SERIES_ORDER = ["PS01", "PS14", "PS02", "PS11", "PS03", "PS04", "PS10", "PS05", "PS06", "PS07", "PS15", "PS08", "PS09", "PS12", "PS13"];
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
    ["PS14", "WATCH Pro系列"], ["PS15", "Band Pro系列"],
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

/** A confirmed deletion unlinks its historical specification without discarding support evidence. */
export function unlinkDeletedHistoricalRequirement(
  feature: TieringFeature,
  requirements: readonly Pick<Requirement, "id">[],
  tombstones: readonly RequirementTombstone[],
): TieringFeature {
  if (!isHistoricalTieringFeature(feature) || !feature.requirementId
    || requirements.some((item) => item.id === feature.requirementId)) return feature;
  const deletion = tombstones.find((item) => item.id === feature.requirementId);
  if (!deletion) return feature;
  return {
    ...feature,
    requirementId: undefined,
    updatedAt: feature.updatedAt >= deletion.deletedAt ? feature.updatedAt : deletion.deletedAt,
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
      reviewer: item.reviewer ?? "",
      recordStatus: item.recordStatus ?? "有效",
      updatedAt: item.updatedAt || item.changedAt,
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

/**
 * 将工作区中的分层分级数据收敛为当前业务范围，不修改传入对象。
 *
 * 该函数刻意不读取当前时间，确保导入检查、差异摘要与真正写入前使用
 * 同一份稳定语义：只保留历史基线或运动/健康/AI 需求关联的功能，补齐
 * 内置系列及完整支持矩阵，并清理被移除功能的从属记录。
 */
export function normalizeWorkspaceTieringScope(raw: WorkspaceData): WorkspaceData {
  const tiering = normalizeTieringData(raw.tiering);
  const domains = raw.domains.map((item) => ({ ...item }));
  const requirements = raw.requirements.map((item) => ({ ...item }));
  const originalFeatures = tiering.features.map((item) => ({ ...item }));
  const originalById = new Map(originalFeatures.map((item) => [item.id, item]));
  const originalByRequirementId = new Map(
    originalFeatures
      .filter((item) => item.requirementId)
      .map((item) => [item.requirementId!, item]),
  );
  const claimedFeatureIds = new Set<string>();
  const nextFeatures: TieringFeature[] = [];
  let nextSortOrder = originalFeatures.reduce((max, item) => Math.max(max, item.sortOrder), -1) + 1;

  for (let index = 0; index < requirements.length; index += 1) {
    const requirement = requirements[index];
    if (!isTieringEligibleRequirement(requirement)) {
      if (requirement.tieringFeatureId) requirements[index] = { ...requirement, tieringFeatureId: undefined };
      continue;
    }
    const resolved = resolveTieringRequirementDomains(requirement, domains);
    const candidate = (requirement.tieringFeatureId ? originalById.get(requirement.tieringFeatureId) : undefined)
      ?? originalByRequirementId.get(requirement.id);
    // 旧 JSON 可能没有完整的 L0/L1 关系；保留既有规格供导入阶段给出原有校验，
    // 但不依据名称猜测或创建错误关联。
    if (!resolved) {
      if (candidate && !claimedFeatureIds.has(candidate.id)) {
        claimedFeatureIds.add(candidate.id);
        nextFeatures.push(candidate);
      }
      continue;
    }
    // 完整包可能只在需求侧保留稳定的历史功能 ID（F001–F156），而没有同时
    // 携带对应的功能行。此时必须沿用需求上的关联 ID，后续再补建功能与内置系列的
    // 系列支持记录；若直接退化为 REQ-*，会丢失历史基线的恢复语义。
    const fallbackId = requirement.tieringFeatureId ?? tieringFeatureIdForRequirement(requirement.id);
    const existing = candidate && !claimedFeatureIds.has(candidate.id) ? candidate : originalById.get(fallbackId);
    const featureId = existing?.id ?? fallbackId;
    if (claimedFeatureIds.has(featureId)) throw new Error(`分层功能“${featureId}”关联了多条需求`);
    claimedFeatureIds.add(featureId);
    const linkedRequirement = {
      ...requirement,
      domainL0Id: resolved.domainL0.id,
      tieringFeatureId: featureId,
    };
    requirements[index] = linkedRequirement;
    nextFeatures.push(buildTieringFeatureForRequirement(
      linkedRequirement,
      resolved.domainL0.name,
      resolved.domainL1.name,
      existing?.sortOrder ?? nextSortOrder++,
      existing?.updatedAt ?? linkedRequirement.updatedAt,
      existing,
    ));
  }

  for (const feature of originalFeatures) {
    if (!isHistoricalTieringFeature(feature) || claimedFeatureIds.has(feature.id)) continue;
    const linkedRequirement = feature.requirementId
      ? requirements.find((item) => item.id === feature.requirementId)
      : undefined;
    const unlinked = linkedRequirement && !isTieringEligibleRequirement(linkedRequirement)
      ? { ...feature, requirementId: undefined }
      : unlinkDeletedHistoricalRequirement(feature, requirements, raw.requirementTombstones ?? []);
    nextFeatures.push(unlinked);
  }

  const builtInById = new Map(BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => [item.id, item]));
  const productSeries = [
    ...tiering.productSeries.map((item) => {
      const builtIn = builtInById.get(item.id);
      return builtIn ? { ...item, sourceIds: [...item.sourceIds], sortOrder: builtIn.sortOrder } : { ...item, sourceIds: [...item.sourceIds] };
    }),
    ...missingBuiltInTieringProductSeries(tiering.productSeries),
  ].sort((left, right) => left.sortOrder - right.sortOrder || left.id.localeCompare(right.id));
  const featureIds = new Set(nextFeatures.map((item) => item.id));
  const seriesIds = new Set(productSeries.map((item) => item.id));
  const supportById = new Map(
    tiering.supportRecords
      .filter((item) => featureIds.has(item.featureId) && seriesIds.has(item.seriesId))
      .map((item) => [item.id, { ...item, sourceIds: [...item.sourceIds] }]),
  );
  for (const feature of nextFeatures) for (const series of productSeries) {
    const id = tieringSupportId(feature.id, series.id);
    if (!supportById.has(id)) supportById.set(id, createTieringSupportRecord(feature.id, series.id, feature.updatedAt));
  }

  return {
    ...raw,
    requirements,
    domains,
    tiering: {
      features: nextFeatures.sort((left, right) => left.sortOrder - right.sortOrder || left.id.localeCompare(right.id)),
      productSeries,
      supportRecords: [...supportById.values()].sort((left, right) => left.id.localeCompare(right.id)),
      evidenceSources: tiering.evidenceSources.map((item) => ({ ...item })),
      verificationIssues: tiering.verificationIssues
        .filter((item) => featureIds.has(item.featureId))
        .map((item) => ({ ...item, seriesIds: [...item.seriesIds], sourceIds: [...item.sourceIds] })),
      changeLogs: tiering.changeLogs
        .filter((item) => featureIds.has(item.featureId))
        .map((item) => ({ ...item, sourceIds: [...item.sourceIds] })),
    },
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
    source: undefined,
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
    changeLogs: mergeByTimestamp(local.changeLogs, incoming.changeLogs),
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
