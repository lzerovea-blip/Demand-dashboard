import { describe, expect, it } from "vitest";
import { cloneRequirementAsCopy, filterRequirements, filterRequirementsForRegion, listReleaseVersionOptions, nextRequirementCopyTitle, normalizeLegacyGroupOverride, normalizeRequirement, requirementVisibleInRegion } from "../src/shared/requirements";
import type { Requirement } from "../src/shared/types";

describe("需求复制", () => {
  it("连续副本使用首个空闲序号", () => {
    expect(nextRequirementCopyTitle("睡眠监测", ["睡眠监测", "睡眠监测 副本（1）", "睡眠监测 副本（3）"])).toBe("睡眠监测 副本（2）");
    expect(nextRequirementCopyTitle("睡眠监测 副本（1）", ["睡眠监测 副本（1）"])).toBe("睡眠监测 副本（2）");
  });

  it("复制全部业务字段并重新生成需求和图片 ID 与时间", () => {
    const source = requirement("r1", "睡眠监测", "d1", "健康基础", "产品专属", "2027-03");
    source.description = "完整描述";
    source.productIds = ["p1"];
    source.images = [{ id: "old-image", name: "a.png", mimeType: "image/png", sizeBytes: 1, dataUrl: "data:image/png;base64,YQ==" }];
    source.tieringFeatureId = "F012";
    source.historicalLaunch = { suggestedTargetMonth: "2014-04", precision: "year", verificationStatus: "pending", confidence: "C", evidenceUrls: ["https://example.com"], note: "待核验" };
    const ids = ["new-requirement", "new-image"];
    const copy = cloneRequirementAsCopy(source, [source.title], () => ids.shift()!, "2026-08-28T10:00:00.000Z");
    expect(copy).toMatchObject({ id: "new-requirement", title: "睡眠监测 副本（1）", description: "完整描述", productIds: ["p1"], createdAt: "2026-08-28T10:00:00.000Z", updatedAt: "2026-08-28T10:00:00.000Z" });
    expect(copy.images[0].id).toBe("new-image");
    expect(copy.images[0].dataUrl).toBe(source.images[0].dataUrl);
    expect(copy.images).not.toBe(source.images);
    expect(copy.productIds).not.toBe(source.productIds);
    expect(copy.tieringFeatureId).toBeUndefined();
    expect(copy.historicalLaunch).toBeUndefined();
  });
});

describe("历史需求兼容", () => {
  it("为旧数据补齐描述和图片字段", () => {
    const normalized = normalizeRequirement({
      id: "r1",
      title: "体动",
      domainId: "d1",
      source: "健康基础",
      category: "体验优化",
      targetMonth: "2027-03",
      productIds: [],
      workloadPm: 1,
      createdAt: "2026-08-19",
      updatedAt: "2026-08-19",
    });
    expect(normalized.description).toBe("");
    expect(normalized.images).toEqual([]);
    expect(normalized.overseasRegions).toEqual([]);
    expect(normalized.scope).toBe("全球");
    expect(normalized.allowLatestVersionUpgrade).toBe(false);
    expect(normalized).toMatchObject({
      deviceWorkloadPm: 0,
      appWorkloadPm: 0,
      cloudWorkloadPm: 0,
      unallocatedWorkloadPm: 1,
      workloadPm: 1,
    });
  });

  it("旧分层历史需求缺少元数据时从稳定功能 ID 自动修正归档月份", () => {
    const normalized = normalizeRequirement({
      ...requirement("historical-huawei-F146", "步数统计", "d1", "运动基础", "平台演进", "2026-09"),
      tieringFeatureId: "F146",
    });
    expect(normalized.targetMonth).toBe("2014-04");
    expect(normalized.historicalLaunch).toMatchObject({ precision: "year", verificationStatus: "pending", confidence: "C" });
  });

  it("保留最近版本升级设置并兼容旧数据默认否", () => {
    const source = requirement("upgrade", "允许升级", "d1", "健康基础", "体验优化", "2027-03");
    expect(normalizeRequirement({ ...source, allowLatestVersionUpgrade: true }).allowLatestVersionUpgrade).toBe(true);
    expect(normalizeRequirement({ ...source, allowLatestVersionUpgrade: undefined }).allowLatestVersionUpgrade).toBe(false);
  });

  it("由设备、App、云侧和待拆分工作量重新计算合计", () => {
    const normalized = normalizeRequirement({
      id: "r2",
      title: "运动负荷",
      description: "",
      images: [],
      domainId: "d1",
      source: "运动进阶",
      category: "体验优化",
      targetMonth: "2027-04",
      productIds: [],
      deviceWorkloadPm: 0.5,
      appWorkloadPm: 1,
      cloudWorkloadPm: 0.25,
      unallocatedWorkloadPm: 0,
      workloadPm: 99,
      createdAt: "2026-08-19",
      updatedAt: "2026-08-19",
    });
    expect(normalized.workloadPm).toBe(1.75);
  });

  it("把旧版海外研究来源迁移为行业并移除旧区域", () => {
    const normalized = normalizeRequirement({
      id: "legacy-overseas",
      title: "旧版海外研究需求",
      description: "",
      images: [],
      domainId: "d1",
      source: "海外研究",
      overseasRegions: ["欧州"],
      category: "体验优化",
      targetMonth: "2027-04",
      productIds: [],
      workloadPm: 1,
      createdAt: "2026-08-19",
      updatedAt: "2026-08-19",
    });
    expect(normalized.source).toBe("行业");
    expect(normalized.overseasRegions).toEqual([]);
  });

  it("把旧版 AI 和小写 ai 来源迁移为 AI综合", () => {
    const normalizedLowercase = normalizeRequirement({
      id: "legacy-ai",
      title: "AI 健康助手",
      description: "",
      images: [],
      domainId: "d1",
      source: "ai",
      overseasRegions: [],
      category: "体验优化",
      targetMonth: "2027-05",
      productIds: [],
      workloadPm: 1,
      createdAt: "2026-08-19",
      updatedAt: "2026-08-19",
    });
    const normalizedUppercase = normalizeRequirement({
      id: "legacy-AI",
      title: "AI 综合助手",
      description: "",
      images: [],
      domainId: "d1",
      source: "AI",
      overseasRegions: [],
      category: "平台演进",
      targetMonth: "2027-05",
      productIds: [],
      workloadPm: 1,
      createdAt: "2026-08-19",
      updatedAt: "2026-08-19",
    });
    expect(normalizedLowercase.source).toBe("AI综合");
    expect(normalizedUppercase.source).toBe("AI综合");
    expect(normalizeLegacyGroupOverride({
      groupKey: "d1::ai::2027-05",
      cardTitle: "AI 健康助手",
      cardSummary: "",
      updatedAt: "2026-08-19",
    }).groupKey).toBe("d1::AI综合::2027-05");
    expect(normalizeLegacyGroupOverride({
      groupKey: "d1::AI::2027-05",
      cardTitle: "AI 综合助手",
      cardSummary: "",
      updatedAt: "2026-08-19",
    }).groupKey).toBe("d1::AI综合::2027-05");
  });

  it("把旧版医疗送检来源和卡片键迁移为 NMPA送检", () => {
    const normalized = normalizeRequirement({
      id: "legacy-medical",
      title: "旧医疗送检需求",
      description: "",
      images: [],
      domainId: "d1",
      source: "医疗送检",
      category: "平台演进",
      targetMonth: "2027-06",
      productIds: [],
      workloadPm: 1,
      createdAt: "2026-08-19",
      updatedAt: "2026-08-19",
    });
    expect(normalized.source).toBe("NMPA送检");
    expect(normalizeLegacyGroupOverride({
      groupKey: "d1::医疗送检::2027-06",
      cardTitle: "送检",
      cardSummary: "",
      updatedAt: "2026-08-19",
    }).groupKey).toBe("d1::NMPA送检::2027-06");
  });

  it("清除体验优化的旧产品关联，并保留平台演进和产品专属的单选产品", () => {
    const exclusive = normalizeRequirement({
      ...requirement("legacy-exclusive", "旧多产品", "d1", "健康基础", "产品专属", "2027-05"),
      productIds: ["p1", "p2", "p1"],
    });
    const experience = normalizeRequirement({
      ...requirement("legacy-experience", "旧体验优化", "d1", "健康基础", "体验优化", "2027-05"),
      productIds: ["p1"],
    });
    const platform = normalizeRequirement({
      ...requirement("legacy-platform", "平台演进", "d1", "健康基础", "平台演进", "2027-05"),
      productIds: ["p1"],
    });
    expect(exclusive.productIds).toEqual(["p1"]);
    expect(experience.productIds).toEqual([]);
    expect(platform.productIds).toEqual(["p1"]);
  });
});

describe("适用范围", () => {
  const domestic = { ...requirement("domestic", "国内", "d1", "健康基础", "体验优化", "2027-03"), scope: "仅国内" as const };
  const overseas = { ...requirement("overseas", "海外", "d1", "健康基础", "体验优化", "2027-03"), scope: "仅海外" as const };
  const global = { ...requirement("global", "全球", "d1", "健康基础", "体验优化", "2027-03"), scope: "全球" as const };

  it("国内和海外视图均包含全球需求", () => {
    expect(filterRequirementsForRegion([domestic, overseas, global], "domestic").map((item) => item.id)).toEqual(["domestic", "global"]);
    expect(filterRequirementsForRegion([domestic, overseas, global], "overseas").map((item) => item.id)).toEqual(["overseas", "global"]);
    expect(requirementVisibleInRegion(global, "domestic")).toBe(true);
    expect(requirementVisibleInRegion(global, "overseas")).toBe(true);
  });

  it("保留三种有效范围，缺失或异常历史值迁移为全球", () => {
    expect(normalizeRequirement(domestic).scope).toBe("仅国内");
    expect(normalizeRequirement(overseas).scope).toBe("仅海外");
    expect(normalizeRequirement({ ...global, scope: undefined }).scope).toBe("全球");
    expect(normalizeRequirement({ ...global, scope: "未知" as never }).scope).toBe("全球");
  });
});

describe("需求池筛选", () => {
  const requirements: Requirement[] = [
    requirement("r1", "社交时差", "d1", "健康基础", "体验优化", "2027-03"),
    requirement("r2", "体动", "d1", "健康基础", "产品专属", "2027-04"),
    requirement("r3", "跑姿", "d2", "运动进阶", "体验优化", "2026-12"),
  ];
  const domains = new Map([["d1", "睡眠"], ["d2", "跑步"]]);

  it("按实际需求提取、去重并按半年排列上线版本", () => {
    expect(listReleaseVersionOptions([...requirements, { ...requirements[0], id: "r4" }, { ...requirements[0], id: "historical", targetMonth: "2014-04" }])).toEqual([
      { halfYear: "2027H1", version: "HMOS 7.1" },
      { halfYear: "2026H2", version: "HMOS 7.0" },
      { halfYear: "2014H1", version: "历史版本" },
    ]);
  });

  it("版本按半年倒序而非 HMOS 字符串排序", () => {
    expect(listReleaseVersionOptions(["2028-09", "2029-09", "2016-04", "2014-04", "invalid"].map((targetMonth) => ({ ...requirements[0], targetMonth }))).map((item) => item.halfYear)).toEqual(["2029H2", "2028H2", "2016H1", "2014H1"]);
  });

  it("版本筛选包含同半年不同月份并与关键词、L0、L1 组合，不再按来源或分类过滤", () => {
    expect(filterRequirements(requirements, domains, {
      query: "睡眠",
      domainL0Id: "l0",
      domainL1Id: "d1",
      targetHalfYear: "2027H1",
    }).map((item) => item.id)).toEqual(["r1", "r2"]);
    expect(filterRequirements(requirements, domains, {
      query: "睡眠",
      domainL0Id: "another-l0",
      domainL1Id: "d1",
      targetHalfYear: "2027H1",
    }).map((item) => item.id)).toEqual([]);
  });
});

function requirement(id: string, title: string, domainId: string, source: Requirement["source"], category: Requirement["category"], targetMonth: string): Requirement {
  return {
    id,
    title,
    description: "",
    images: [],
    domainL0Id: "l0",
    domainId,
    source,
    scope: "全球",
    allowLatestVersionUpgrade: false,
    overseasRegions: [],
    category,
    targetMonth,
    productIds: [],
    deviceWorkloadPm: 0,
    appWorkloadPm: 1,
    cloudWorkloadPm: 0,
    unallocatedWorkloadPm: 0,
    workloadPm: 1,
    createdAt: "2026-01-01",
    updatedAt: "2026-01-01",
  };
}
