import { describe, expect, it } from "vitest";
import { latestRequirementSelections } from "../src/shared/entryDefaults";
import type { Requirement } from "../src/shared/types";

describe("快速录入默认项", () => {
  it("沿用最近一次保存的来源、分类和上线年月", () => {
    const requirements = [
      requirement("旧", "运动基础", "体验优化", "2026-08", "2026-08-01"),
      requirement("新", "健康进阶", "产品专属", "2027-03", "2026-08-02"),
    ];
    expect(latestRequirementSelections(requirements)).toEqual({
      source: "健康进阶",
      overseasRegions: [],
      category: "产品专属",
      targetMonth: "2027-03",
    });
  });

  it("没有历史需求时不提供覆盖值", () => {
    expect(latestRequirementSelections([])).toBeUndefined();
  });
});

function requirement(
  id: string,
  source: Requirement["source"],
  category: Requirement["category"],
  targetMonth: string,
  updatedAt: string,
): Requirement {
  return {
    id,
    title: id,
    description: "",
    images: [],
    domainId: "d1",
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
    createdAt: updatedAt,
    updatedAt,
  };
}
