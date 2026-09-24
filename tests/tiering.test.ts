import { describe, expect, it } from "vitest";
import { tieringFeatureHalfYear, tieringMatchesOwnersAndVersion } from "../src/shared/tieringFilters";
import {
  buildHistoricalRequirement,
  buildTieringFeatureForRequirement,
  BUILT_IN_TIERING_PRODUCT_SERIES,
  createTieringSupportRecord,
  EMPTY_TIERING_DATA,
  historicalRequirementId,
  isTieringEligibleSource,
  migrateArchivedHistoricalRequirement,
  migrateTieringRequirementDescription,
  mergeTieringData,
  missingBuiltInTieringProductSeries,
  normalizeTieringData,
  normalizeTieringSupportStatus,
  normalizeWorkspaceTieringScope,
  resolveTieringRequirementDomains,
  TIERING_EXPECTED_COUNTS,
  tieringFeatureIdForRequirement,
  tieringSourceForL0,
  tieringSupportId,
  unlinkDeletedHistoricalRequirement,
  validateTieringData,
  validateTieringSupportRecord,
} from "../src/shared/tiering";
import { buildHistoricalLaunchLedger, formatHistoricalLaunch, historicalLaunchForFeatureId } from "../src/shared/tieringLaunches";
import type { DictionaryItem, Requirement, TieringFeature, TieringSupportRecord, WorkspaceData } from "../src/shared/types";

it("未定送检版本在分层全部版本可见，不回退到旧历史日期匹配具体版本", () => {
  const feature = { id: "F001", historicalLaunch: { suggestedTargetMonth: "2027-03" } } as TieringFeature;
  const requirement = { targetMonth: "", source: "CE认证" } as Requirement;
  expect(tieringFeatureHalfYear(feature, requirement)).toBe("");
  expect(tieringMatchesOwnersAndVersion(feature, requirement, [], { planningMemberId: "", productManagerMemberId: "", halfYear: "" })).toBe(true);
  expect(tieringMatchesOwnersAndVersion(feature, requirement, [], { planningMemberId: "", productManagerMemberId: "", halfYear: "2027H1" })).toBe(false);
});

const feature: TieringFeature = { id: "F001", l0: "基础体征测量", l1: "心率", requirementTitle: "单次心率测量", requirementId: "historical-huawei-F001", sortOrder: 0, createdAt: "2026-08-05T00:00:00.000Z", updatedAt: "2026-08-05T00:00:00.000Z" };
const support: TieringSupportRecord = { id: "F001::PS01", featureId: "F001", seriesId: "PS01", status: "支持", supportedModels: "", excludedModels: "", phoneOsMarket: "中国大陆", minHealthAppVersion: "", minDeviceVersion: "", prerequisites: "", sourceIds: ["S001"], sourceUrl: "", verifiedAt: "2026-08-05", evidenceLevel: "A", auditStatus: "已公开核验", note: "", updatedAt: "2026-08-05T00:00:00.000Z" };

describe("需求分层分级", () => {
  it("内置十五个稳定产品系列供全新电脑初始化", () => {
    expect(BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => item.id)).toEqual([
      "PS01", "PS14", "PS02", "PS11", "PS03", "PS04", "PS10", "PS05", "PS06", "PS07", "PS15", "PS08", "PS09", "PS12", "PS13",
    ]);
    expect(BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => item.name)).toEqual([
      "WATCH数字系列", "WATCH Pro系列", "GT系列", "GT Pro系列", "ULTIMATE系列", "FIT系列", "FIT Pro系列", "D系列", "WATCH Buds系列", "Band系列", "Band Pro系列", "儿童手表", "体脂秤", "戒指系列", "无屏系列",
    ]);
    expect(TIERING_EXPECTED_COUNTS).toMatchObject({ features: 156, productSeries: 15, supportRecords: 2340 });
  });

  it("只补齐缺失系列且不覆盖本机已有系列资料", () => {
    const existing = [{ ...BUILT_IN_TIERING_PRODUCT_SERIES[0], name: "本机自定义名称" }];
    const missing = missingBuiltInTieringProductSeries(existing);
    expect(missing).toHaveLength(14);
    expect(missing.some((item) => item.id === "PS01")).toBe(false);
    expect(existing[0].name).toBe("本机自定义名称");
  });

  it("使用稳定功能、系列支持和历史需求 ID", () => {
    expect(tieringSupportId("F001", "PS01")).toBe("F001::PS01");
    expect(historicalRequirementId("F156")).toBe("historical-huawei-F156");
  });

  it("九类 L0 按确认规则映射到需求来源", () => {
    expect(["基础体征测量", "休息放松"].map(tieringSourceForL0)).toEqual(["健康基础", "健康基础"]);
    expect(["健康报告", "女性健康", "智能关爱"].map(tieringSourceForL0)).toEqual(["健康进阶", "健康进阶", "健康进阶"]);
    expect(["华为创新研究app", "慢病管理"].map(tieringSourceForL0)).toEqual(["健康高阶", "健康高阶"]);
    expect(tieringSourceForL0("保持活力")).toBe("运动基础");
    expect(tieringSourceForL0("专业运动")).toBe("运动高阶");
  });

  it("只将运动、健康和 AI 的九个来源纳入分层分级", () => {
    expect([
      "运动基础", "运动进阶", "运动高阶",
      "健康基础", "健康进阶", "健康高阶",
      "AI运动", "AI健康", "AI综合",
    ].every((source) => isTieringEligibleSource(source as Requirement["source"]))).toBe(true);
    expect(["行业", "全场景", "health平台", "NMPA认证", "CE认证"]
      .every((source) => !isTieringEligibleSource(source as Requirement["source"]))).toBe(true);
  });

  it("历史需求使用逐功能建议上线时间并带证据标识", () => {
    const l0: DictionaryItem = { id: "l0", name: "基础体征测量", sortOrder: 0, active: true, level: "L0" };
    const l1: DictionaryItem = { id: "l1", name: "心率", sortOrder: 0, active: true, level: "L1", parentId: "l0" };
    const requirement = buildHistoricalRequirement(feature, [support], [{ id: "PS01", name: "WATCH数字系列", boundary: "", representativeProducts: "", managementEntry: "", sourceIds: [], lifecycle: "", note: "", sortOrder: 0, updatedAt: support.updatedAt }], l0, l1);
    expect(requirement).toMatchObject({ id: "historical-huawei-F001", title: "单次心率测量", targetMonth: "2015-09", category: "平台演进", scope: "仅国内", allowLatestVersionUpgrade: true, workloadPm: 0, source: "健康基础", tieringFeatureId: "F001" });
    expect(requirement.historicalLaunch).toMatchObject({ precision: "year", verificationStatus: "pending", confidence: "C" });
    expect(requirement.description).toContain("仅用于半年时间轴归档，不代表真实首发月份");
  });

  it("156 个稳定功能 ID 均有历史上线证据且覆盖 2014—2025", () => {
    const ledger = buildHistoricalLaunchLedger(Array.from({ length: 156 }, (_, index) => ({ id: `F${String(index + 1).padStart(3, "0")}` })));
    expect(ledger).toHaveLength(156);
    expect(new Set(ledger.map((item) => item.featureId)).size).toBe(156);
    const years = ledger.map((item) => Number(item.suggestedTargetMonth.slice(0, 4)));
    expect(Math.min(...years)).toBe(2014);
    expect(Math.max(...years)).toBe(2025);
  });

  it("区分明确月份、半年占位和年份待核验", () => {
    expect(historicalLaunchForFeatureId("F054")).toMatchObject({ suggestedTargetMonth: "2018-09", precision: "month", verificationStatus: "confirmed" });
    expect(historicalLaunchForFeatureId("F002")).toMatchObject({ suggestedTargetMonth: "2018-09", precision: "half-year", verificationStatus: "pending" });
    expect(historicalLaunchForFeatureId("F146")).toMatchObject({ suggestedTargetMonth: "2014-04", precision: "year", verificationStatus: "pending" });
    expect(formatHistoricalLaunch(historicalLaunchForFeatureId("F002"))).toBe("2018H2 · 月份待核验");
    expect(formatHistoricalLaunch(historicalLaunchForFeatureId("F146"))).toBe("2014年 · 半年及月份待核验");
  });

  it("把旧 2026-09 归档需求幂等迁移到真实时间台账", () => {
    const legacy = {
      ...buildHistoricalRequirement(feature, [support], [{ id: "PS01", name: "WATCH数字系列", boundary: "", representativeProducts: "", managementEntry: "", sourceIds: [], lifecycle: "", note: "", sortOrder: 0, updatedAt: support.updatedAt }], { id: "l0", name: "基础体征测量", sortOrder: 0, active: true, level: "L0" }, { id: "l1", name: "心率", sortOrder: 0, active: true, level: "L1", parentId: "l0" }),
      targetMonth: "2026-09",
      historicalLaunch: undefined,
      description: "路径\n归档月份：2026年9月，仅用于历史建档，不代表真实首发月份。",
    };
    const migrated = migrateArchivedHistoricalRequirement(legacy, feature);
    expect(migrated.targetMonth).toBe("2015-09");
    expect(migrated.historicalLaunch?.precision).toBe("year");
    expect(migrated.description).not.toContain("归档月份：2026年9月");
    expect(migrateArchivedHistoricalRequirement(migrated, feature)).toEqual(migrated);
  });

  it("普通 REQ 分层功能不进入历史上线时间迁移", () => {
    const ordinaryFeature = {
      ...feature,
      id: "REQ-450151ce-e9df-4111-afee-03e240805c9f",
      requirementId: "450151ce-e9df-4111-afee-03e240805c9f",
    };
    const ordinaryRequirement = {
      ...buildHistoricalRequirement(feature, [support], [], { id: "l0", name: "基础体征测量", sortOrder: 0, active: true, level: "L0" }, { id: "l1", name: "心率", sortOrder: 0, active: true, level: "L1", parentId: "l0" }),
      id: ordinaryFeature.requirementId,
      tieringFeatureId: ordinaryFeature.id,
      targetMonth: "2026-09",
      historicalLaunch: undefined,
    };
    expect(migrateArchivedHistoricalRequirement(ordinaryRequirement, ordinaryFeature)).toEqual(ordinaryRequirement);
  });

  it("历史需求描述移除 L3 并使用分层分级新名称", () => {
    const requirement = buildHistoricalRequirement(feature, [support], [], { id: "l0", name: "基础体征测量", sortOrder: 0, active: true, level: "L0" }, { id: "l1", name: "心率", sortOrder: 0, active: true, level: "L1", parentId: "l0" });
    const legacy = { ...requirement, description: requirement.description.replace("基础体征测量 → 心率 → 单次心率测量", "基础体征测量 → 心率 → 日常心率 → 单次心率测量").replace("分层分级", "需求分层分级管理") };
    const migrated = migrateTieringRequirementDescription(legacy, feature);
    expect(migrated.description).toContain("分层路径：基础体征测量 → 心率 → 单次心率测量");
    expect(migrated.description).not.toContain("日常心率");
    expect(migrated.description).not.toContain("需求分层分级管理");
  });

  it("旧七状态按确认规则收敛为支持、不支持、未判断", () => {
    expect(["支持", "部分支持", "条件支持", "部分且条件"].map(normalizeTieringSupportStatus)).toEqual(["支持", "支持", "支持", "支持"]);
    expect(["不支持", "不适用"].map(normalizeTieringSupportStatus)).toEqual(["不支持", "不支持"]);
    expect(normalizeTieringSupportStatus("待核验")).toBe("未判断");
    expect(() => validateTieringSupportRecord({ ...support, status: "未判断" })).not.toThrow();
  });

  it("兼容旧 L1-L4 功能并永久丢弃 L3", () => {
    const legacy = { id: "F001", l1: "基础体征测量", l2: "心率", l3: "日常心率", l4: "单次心率测量", sortOrder: 0, createdAt: support.updatedAt, updatedAt: support.updatedAt };
    const normalized = normalizeTieringData({ features: [legacy as unknown as TieringFeature] });
    expect(normalized.features[0]).toMatchObject({ l0: "基础体征测量", l1: "心率", requirementTitle: "单次心率测量" });
    expect(normalized.features[0]).not.toHaveProperty("l3");
    expect(normalized.features[0]).not.toHaveProperty("l4");
  });

  it("普通需求使用稳定 ID 建立分层行并初始化未判断状态", () => {
    const requirement = {
      id: "r-normal",
      title: "新的健康需求",
      tieringFeatureId: tieringFeatureIdForRequirement("r-normal"),
    } as Requirement;
    const next = buildTieringFeatureForRequirement(requirement, "健康报告", "健康摘要", 156, support.updatedAt);
    expect(next).toMatchObject({ id: "REQ-r-normal", l0: "健康报告", l1: "健康摘要", requirementTitle: "新的健康需求", requirementId: "r-normal", sortOrder: 156 });
    expect(createTieringSupportRecord(next.id, "PS01", support.updatedAt)).toMatchObject({ id: "REQ-r-normal::PS01", status: "未判断" });
  });

  it("仅需求侧保留 F001 时仍恢复历史功能，移出范围后只解除关联", () => {
    const domainL0: DictionaryItem = { id: "l0", name: "基础体征测量", sortOrder: 0, active: true, level: "L0" };
    const domainL1: DictionaryItem = { id: "l1", name: "心率", sortOrder: 0, active: true, level: "L1", parentId: "l0" };
    const historicalRequirement = {
      ...buildHistoricalRequirement(feature, [], [], domainL0, domainL1),
      id: "r-f001",
      tieringFeatureId: "F001",
    };
    const workspace: WorkspaceData = {
      requirements: [historicalRequirement],
      domains: [domainL0, domainL1],
      products: [],
      members: [],
      groupOverrides: [],
      pptAccessPolicy: { entries: [], updatedAt: support.updatedAt },
      tiering: structuredClone(EMPTY_TIERING_DATA),
    };

    const linked = normalizeWorkspaceTieringScope(workspace);
    expect(linked.requirements[0].tieringFeatureId).toBe("F001");
    expect(linked.tiering?.features).toContainEqual(expect.objectContaining({ id: "F001", requirementId: "r-f001" }));
    expect(linked.tiering?.supportRecords.filter((item) => item.featureId === "F001")).toHaveLength(15);

    const unlinked = normalizeWorkspaceTieringScope({
      ...linked,
      requirements: linked.requirements.map((item) => ({ ...item, source: "行业" as const })),
    });
    expect(unlinked.requirements[0].tieringFeatureId).toBeUndefined();
    expect(unlinked.tiering?.features).toContainEqual(expect.objectContaining({ id: "F001", requirementId: undefined }));
    expect(unlinked.tiering?.supportRecords.filter((item) => item.featureId === "F001")).toHaveLength(15);
  });

  it("历史需求已确认删除时解除旧关联并保留规格，缺少删除记录时仍提示真正异常", () => {
    const deletedFeature = { ...feature, requirementId: "r-f001" };
    const tombstone = { id: "r-f001", title: feature.requirementTitle, deletedAt: "2026-09-24T00:00:00.000Z" };
    expect(unlinkDeletedHistoricalRequirement(deletedFeature, [], [])).toBe(deletedFeature);
    expect(unlinkDeletedHistoricalRequirement(deletedFeature, [], [tombstone])).toMatchObject({ id: "F001", requirementId: undefined, updatedAt: tombstone.deletedAt });
    const normalized = normalizeWorkspaceTieringScope({
      requirements: [], requirementTombstones: [tombstone], domains: [], products: [], members: [], groupOverrides: [],
      pptAccessPolicy: { entries: [], updatedAt: tombstone.deletedAt },
      tiering: { ...structuredClone(EMPTY_TIERING_DATA), features: [deletedFeature], productSeries: [BUILT_IN_TIERING_PRODUCT_SERIES[0]], supportRecords: [support] },
    });
    expect(normalized.tiering?.features[0].requirementId).toBeUndefined();
    expect(normalized.tiering?.supportRecords.find((item) => item.id === support.id)?.status).toBe("支持");
    expect(normalizeWorkspaceTieringScope({ ...normalized, requirementTombstones: [] }).tiering?.features[0].requirementId).toBeUndefined();
  });

  it("旧需求领域关联失效时可识别并跳过启动同步", () => {
    const domains: DictionaryItem[] = [
      { id: "l0", name: "健康", sortOrder: 0, active: true, level: "L0" },
      { id: "l1", name: "睡眠", sortOrder: 0, active: true, level: "L1", parentId: "l0" },
    ];
    expect(resolveTieringRequirementDomains({ domainL0Id: "legacy-l0", domainId: "l1" }, domains)).toMatchObject({
      domainL0: { id: "l0" },
      domainL1: { id: "l1" },
    });
    expect(resolveTieringRequirementDomains({ domainL0Id: "l0", domainId: "missing-l1" }, domains)).toBeUndefined();
  });

  it("分层数据按稳定 ID 幂等合并并保留较新版本", () => {
    const base = { ...EMPTY_TIERING_DATA, features: [feature] };
    const newer = { ...EMPTY_TIERING_DATA, features: [{ ...feature, requirementTitle: "新版名称", updatedAt: "2026-09-01T00:00:00.000Z" }] };
    const merged = mergeTieringData(base, newer);
    expect(merged.features).toHaveLength(1);
    expect(merged.features[0].requirementTitle).toBe("新版名称");
  });

  it("拒绝支持记录引用不存在的功能或系列", () => {
    expect(() => validateTieringData({ ...EMPTY_TIERING_DATA, supportRecords: [support] })).toThrow("不存在的功能或系列");
  });
});
