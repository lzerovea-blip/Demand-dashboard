import { describe, expect, it } from "vitest";
import { filterRequirements, filterRequirementsForRegion, listTargetMonths, normalizeLegacyGroupOverride, normalizeRequirement, requirementVisibleInRegion } from "../src/shared/requirements";
import { splitTextIntoLines } from "../src/shared/text";
import type { Requirement } from "../src/shared/types";

describe("需求池标题分行", () => {
  it("每20个Unicode字符换行且不限制行数", () => {
    expect(splitTextIntoLines("字".repeat(20))).toEqual(["字".repeat(20)]);
    expect(splitTextIntoLines("字".repeat(21))).toEqual(["字".repeat(20), "字"]);
    expect(splitTextIntoLines("字".repeat(41))).toEqual(["字".repeat(20), "字".repeat(20), "字"]);
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
    expect(normalized).toMatchObject({
      deviceWorkloadPm: 0,
      appWorkloadPm: 0,
      cloudWorkloadPm: 0,
      unallocatedWorkloadPm: 1,
      workloadPm: 1,
    });
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

  it("把旧版小写 ai 来源迁移为大写 AI", () => {
    const normalized = normalizeRequirement({
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
    expect(normalized.source).toBe("AI");
    expect(normalizeLegacyGroupOverride({
      groupKey: "d1::ai::2027-05",
      cardTitle: "AI 健康助手",
      cardSummary: "",
      updatedAt: "2026-08-19",
    }).groupKey).toBe("d1::AI::2027-05");
  });

  it("把旧版产品匹配收敛为产品专属单选，体验优化不保留产品", () => {
    const exclusive = normalizeRequirement({
      ...requirement("legacy-exclusive", "旧多产品", "d1", "健康基础", "产品专属", "2027-05"),
      productIds: ["p1", "p2", "p1"],
    });
    const experience = normalizeRequirement({
      ...requirement("legacy-experience", "旧体验优化", "d1", "健康基础", "体验优化", "2027-05"),
      productIds: ["p1"],
    });
    expect(exclusive.productIds).toEqual(["p1"]);
    expect(experience.productIds).toEqual([]);
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

  it("提取、去重并按时间排列完整年月", () => {
    expect(listTargetMonths([...requirements, { ...requirements[0], id: "r4" }])).toEqual(["2026-12", "2027-03", "2027-04"]);
  });

  it("月份与关键词、来源、分类组合筛选", () => {
    expect(filterRequirements(requirements, domains, {
      query: "睡眠",
      source: "健康基础",
      category: "体验优化",
      targetMonth: "2027-03",
    }).map((item) => item.id)).toEqual(["r1"]);
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
