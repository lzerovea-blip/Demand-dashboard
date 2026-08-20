import { describe, expect, it } from "vitest";
import {
  buildHalfYearSummaries,
  buildRoadmapGroups,
  groupKeyFor,
  halfYearOf,
  halfYearSequence,
  monthsForHalfYear,
  roadmapCardLabel,
  roadmapMonthColumnTemplate,
  roadmapTrackOrder,
} from "../src/shared/roadmap";
import type { AppSnapshot, Requirement } from "../src/shared/types";

const baseRequirements: Requirement[] = [
  requirement("r1", "早间血压测量", "d1", "健康基础", "体验优化", "2026-05", [], 1.5, { device: 1.5 }),
  requirement("r2", "午间血压测量", "d1", "健康基础", "产品专属", "2026-05", ["p1"], 2, { app: 2 }),
  requirement("r3", "夜间血压测量", "d1", "健康基础", "产品专属", "2026-06", ["p1", "p2"], 2.5, { cloud: 2.5 }),
  requirement("r4", "跑姿纠正", "d2", "运动进阶", "体验优化", "2026-07", [], 3),
];

const snapshot: AppSnapshot = {
  requirements: baseRequirements,
  domains: [
    { id: "d1", name: "血压", active: true, sortOrder: 0 },
    { id: "d2", name: "跑步", active: true, sortOrder: 1 },
  ],
  products: [
    { id: "p1", name: "产品A", active: true, sortOrder: 0 },
    { id: "p2", name: "产品B", active: true, sortOrder: 1 },
  ],
  groupOverrides: [{ groupKey: "d1::健康基础::2026-05", cardTitle: "分时血压", cardSummary: "早间与午间测量", updatedAt: "2026-01-01" }],
  templates: [],
};

describe("半年与月份", () => {
  it("按六月/七月正确切分半年", () => {
    expect(halfYearOf("2026-06")).toBe("2026H1");
    expect(halfYearOf("2026-07")).toBe("2026H2");
  });

  it("生成连续半年和固定六个月", () => {
    expect(halfYearSequence("2026H2", "2027H2")).toEqual(["2026H2", "2027H1", "2027H2"]);
    expect(monthsForHalfYear("2026H2")).toEqual(["2026-07", "2026-08", "2026-09", "2026-10", "2026-11", "2026-12"]);
  });
});

describe("领域合并", () => {
  it("分类不同但领域、来源、月份相同仍合并", () => {
    const groups = buildRoadmapGroups(snapshot, "2026H1", "2026H2");
    const may = groups.find((item) => item.targetMonth === "2026-05");
    expect(may?.requirements).toHaveLength(2);
    expect(may?.categories).toEqual(["体验优化", "产品专属"]);
    expect(may?.cardTitle).toBe("分时血压");
    expect(may?.productNames).toEqual(["产品A"]);
    expect(may?.totalWorkloadPm).toBe(3.5);
  });

  it("月份不同不合并", () => {
    const may = groupKeyFor(baseRequirements[0]);
    const june = groupKeyFor(baseRequirements[2]);
    expect(may).not.toBe(june);
  });
});

describe("工作量汇总", () => {
  it("按半年、来源、运动健康和分类准确汇总", () => {
    const summaries = buildHalfYearSummaries(baseRequirements, "2026H1", "2026H2");
    expect(summaries[0]).toMatchObject({
      halfYear: "2026H1",
      healthWorkload: 6,
      sportsWorkload: 0,
      experienceWorkload: 1.5,
      exclusiveWorkload: 4.5,
      totalWorkload: 6,
      deviceWorkload: 1.5,
      appWorkload: 2,
      cloudWorkload: 2.5,
      unallocatedWorkload: 0,
      bySide: { device: 1.5, app: 2, cloud: 2.5 },
      experienceBySide: { device: 1.5, app: 0, cloud: 0 },
      requirementCount: 3,
    });
    expect(summaries[1]).toMatchObject({
      halfYear: "2026H2",
      sportsWorkload: 3,
      totalWorkload: 3,
      requirementCount: 1,
    });
  });

  it("历史待拆分工作量不混入三端投入与占比", () => {
    const legacy = {
      ...requirement("legacy", "历史需求", "d1", "健康基础", "体验优化", "2026-05", [], 1),
      appWorkloadPm: 0,
      unallocatedWorkloadPm: 1,
    };
    const [summary] = buildHalfYearSummaries([legacy], "2026H1", "2026H1");
    expect(summary).toMatchObject({
      totalWorkload: 0,
      deviceWorkload: 0,
      appWorkload: 0,
      cloudWorkload: 0,
      unallocatedWorkload: 1,
      sportsWorkload: 0,
      healthWorkload: 0,
    });
  });
});

describe("路标首屏排序", () => {
  it("只有健康需求时优先展示健康路标", () => {
    const groups = buildRoadmapGroups(snapshot, "2026H1", "2026H1");
    expect(roadmapTrackOrder(groups)).toEqual(["健康", "运动"]);
  });

  it("存在运动需求时保持运动路标在前", () => {
    const groups = buildRoadmapGroups(snapshot, "2026H1", "2026H2");
    expect(roadmapTrackOrder(groups)).toEqual(["运动", "健康"]);
  });
});

describe("路标月份列宽", () => {
  it("空月份固定为72px，有需求月份保持宽列", () => {
    const months = ["2026-05", "2026-06", "2026-07"];
    expect(roadmapMonthColumnTemplate(months, [{ targetMonth: "2026-06" }])).toBe(
      "72px minmax(280px, 1fr) 72px",
    );
  });

  it("按传入的运动健康全局卡片统一判断月份占用", () => {
    const groups = buildRoadmapGroups(snapshot, "2026H1", "2026H2");
    const months = ["2026-05", "2026-06", "2026-07", "2026-08"];
    expect(roadmapMonthColumnTemplate(months, groups)).toBe(
      "minmax(280px, 1fr) minmax(280px, 1fr) minmax(280px, 1fr) 72px",
    );
  });

  it("整段没有需求时全部压缩", () => {
    expect(roadmapMonthColumnTemplate(["2027-01", "2027-02", "2027-03"], [])).toBe("72px 72px 72px");
  });
});

describe("路标长条卡片文案", () => {
  it("一条需求时显示标题和需求标题", () => {
    expect(roadmapCardLabel({ cardTitle: "健康摘要", requirements: [baseRequirements[0]] })).toBe("健康摘要：早间血压测量");
  });

  it("最多显示三条需求，超过时追加等", () => {
    expect(roadmapCardLabel({ cardTitle: "健康摘要", requirements: baseRequirements.slice(0, 3) })).toBe("健康摘要：早间血压测量、午间血压测量、夜间血压测量");
    expect(roadmapCardLabel({ cardTitle: "健康摘要", requirements: baseRequirements })).toBe("健康摘要：早间血压测量、午间血压测量、夜间血压测量等");
  });
});

function requirement(
  id: string,
  title: string,
  domainId: string,
  source: Requirement["source"],
  category: Requirement["category"],
  targetMonth: string,
  productIds: string[],
  workloadPm: number,
  breakdown: Partial<{ device: number; app: number; cloud: number }> = { app: workloadPm },
): Requirement {
  return { id, title, description: "", images: [], domainId, source, category, targetMonth, productIds, deviceWorkloadPm: breakdown.device ?? 0, appWorkloadPm: breakdown.app ?? 0, cloudWorkloadPm: breakdown.cloud ?? 0, unallocatedWorkloadPm: 0, workloadPm, createdAt: "2026-01-01", updatedAt: "2026-01-01" };
}
