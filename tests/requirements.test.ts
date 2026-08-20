import { describe, expect, it } from "vitest";
import { filterRequirements, listTargetMonths, normalizeRequirement } from "../src/shared/requirements";
import type { Requirement } from "../src/shared/types";

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
    domainId,
    source,
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
