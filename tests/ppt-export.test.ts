import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { createRoadmapPresentation, layoutRoadmapCell, measureTextWidth } from "../src/main/pptxExport";
import {
  DEFAULT_PPT_EXPORT_SECTIONS,
  MAX_PPT_HALF_YEARS,
  buildPptExportPlan,
  roadmapBadgeKind,
  visibleRequirementTitles,
} from "../src/shared/pptExport";
import { buildRoadmapGroups } from "../src/shared/roadmap";
import type { AppSnapshot, PptExportRequest, Requirement, RequirementCategory, RequirementSource } from "../src/shared/types";

function requirement(
  id: string,
  title: string,
  source: RequirementSource,
  targetMonth: string,
  category: RequirementCategory = "体验优化",
  productIds: string[] = [],
  domainId = "d1",
): Requirement {
  return {
    id,
    title,
    description: `${title}的描述`,
    images: [],
    domainL0Id: "l0",
    domainId,
    source,
    scope: "全球",
    allowLatestVersionUpgrade: false,
    overseasRegions: [],
    category,
    targetMonth,
    productIds,
    deviceWorkloadPm: 1,
    appWorkloadPm: 0.5,
    cloudWorkloadPm: 0.25,
    unallocatedWorkloadPm: 0,
    workloadPm: 1.75,
    createdAt: "2026-08-20T00:00:00.000Z",
    updatedAt: `2026-08-20T00:00:${id.padStart(2, "0")}.000Z`,
  };
}

const snapshot: AppSnapshot = {
  requirements: [
    requirement("1", "社交时差、体动", "健康基础", "2026-03"),
    requirement("2", "睡眠监测", "健康基础", "2026-03"),
    requirement("3", "夜间恢复", "健康基础", "2026-03"),
    requirement("4", "异常趋势", "健康基础", "2026-03", "平台演进"),
    requirement("5", "另一个月的独立卡片", "健康基础", "2026-05", "产品专属", ["p1"]),
    requirement("6", "训练建议", "AI运动", "2026-08"),
    requirement("7", "NMPA资料", "NMPA送检", "2027-02", "平台演进", [], "d2"),
  ],
  domains: [
    { id: "l0", name: "健康", sortOrder: 0, active: true, level: "L0" },
    { id: "d1", name: "睡眠", sortOrder: 0, active: true, level: "L1", parentId: "l0" },
    { id: "d2", name: "心电算法", sortOrder: 1, active: true, level: "L1", parentId: "l0" },
  ],
  products: [{ id: "p1", name: "手表A", sortOrder: 0, active: true }],
  groupOverrides: [],
  members: [],
  pptAccessPolicy: { entries: [], updatedAt: "1970-01-01T00:00:00.000Z" },
  templates: [],
};

const fullRequest: PptExportRequest = {
  start: "2026H1",
  end: "2027H2",
  materialName: "解决方案需求路标",
  sections: [...DEFAULT_PPT_EXPORT_SECTIONS],
};

describe("PPT导出计划", () => {
  it("全选严格输出固定封面和八个单页", () => {
    const plan = buildPptExportPlan(snapshot, fullRequest);
    expect(plan.slideCount).toBe(9);
    expect(plan.roadmapPages.map((item) => item.track)).toEqual(["运动", "健康", "行业", "全场景", "health平台", "AI", "医疗送检"]);
    expect(plan.roadmapPages.find((item) => item.track === "行业")?.groups).toEqual([]);
  });

  it("部分选择页数等于固定封面加选中项且保持固定顺序", () => {
    const plan = buildPptExportPlan(snapshot, { ...fullRequest, sections: ["医疗送检", "summary", "运动"] });
    expect(plan.selectedSections).toEqual(["summary", "运动", "医疗送检"]);
    expect(plan.slideCount).toBe(4);
  });

  it("超过四个半年明确拒绝而不是截断", () => {
    expect(() => buildPptExportPlan(snapshot, { ...fullRequest, end: "2028H1" })).toThrow(`${MAX_PPT_HALF_YEARS} 个连续半年`);
  });

  it("同一领域同一半年不同月份保持两张独立卡片", () => {
    const health = buildPptExportPlan(snapshot, fullRequest).roadmapPages.find((item) => item.track === "健康")!;
    expect(health.groups.map((item) => item.targetMonth)).toEqual(["2026-03", "2026-05"]);
  });
});

describe("路标卡片内容和标识", () => {
  it("单条组合标题不按顿号拆分，四条时仅第三条增加省略号", () => {
    expect(visibleRequirementTitles(snapshot.requirements.slice(0, 4))).toEqual(["社交时差、体动", "睡眠监测", "夜间恢复…"]);
  });

  it("产品标识优先于体验优化，平台演进无产品不显示", () => {
    const productGroup = buildRoadmapGroups({ ...snapshot, requirements: [
      requirement("10", "优化", "健康基础", "2026-03"),
      requirement("11", "有产品", "健康基础", "2026-03", "平台演进", ["p1"]),
    ] }, "2026H1", "2026H1")[0];
    const platformGroup = buildRoadmapGroups({ ...snapshot, requirements: [requirement("12", "平台", "健康基础", "2026-03", "平台演进")] }, "2026H1", "2026H1")[0];
    expect(roadmapBadgeKind(productGroup)).toBe("product");
    expect(roadmapBadgeKind(platformGroup)).toBeNull();
  });

  it("字体按0.5pt递减并在极限情况下压缩边距和需求数", () => {
    const groups = buildRoadmapGroups({ ...snapshot, requirements: Array.from({ length: 12 }, (_, index) => requirement(
      String(20 + index),
      `连续英文VeryLongUnbrokenRequirementTitle${index}`,
      "健康基础",
      `2026-${String(index < 6 ? 3 : 5).padStart(2, "0")}`,
      "体验优化",
      [],
      index < 6 ? "d1" : "d2",
    )) }, "2026H1", "2026H1");
    const layout = layoutRoadmapCell(groups, 273.77, 170.95);
    expect([10, 9.5, 9, 8.5, 8]).toContain(layout.fontSize);
    expect(layout.cards.every((card) => card.top + card.height <= 170.95)).toBe(true);
    expect(measureTextWidth("连续英文ABC123", 10, true)).toBeGreaterThan(0);
  });
});

describe("PPTX输出", () => {
  it("生成原生可编辑的完整9页文件且无详情页、月份标题或内部制作文字", async () => {
    const bytes = await createRoadmapPresentation(snapshot, fullRequest);
    const zip = await JSZip.loadAsync(bytes);
    const slideFiles = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
    expect(slideFiles).toHaveLength(9);
    expect(zip.file("ppt/theme/theme1.xml")).toBeTruthy();
    const xml = (await Promise.all(slideFiles.map((name) => zip.file(name)!.async("string")))).join("\n");
    expect(xml).toContain("需求工作量汇总分析");
    expect(xml).toContain("三端需求结构");
    expect(xml).not.toContain("七类路标投入");
    expect(xml).not.toContain("领域详情");
    expect(xml).not.toContain("制作说明");
    expect(xml).not.toMatch(/>\d{1,2}月</);
    expect(xml).toContain("社交时差、体动");
    expect(xml).toContain("产品专属");
    expect(xml).toContain("↑");
    expect(xml).toContain("当前区间暂无需求");
  }, 30_000);

  it("空状态路标仍输出两行半年表头和泳道", async () => {
    const bytes = await createRoadmapPresentation(snapshot, { ...fullRequest, sections: ["行业"] });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("ppt/slides/slide2.xml")!.async("string");
    expect(xml).toContain("行业路标");
    expect(xml).toContain("26H1");
    expect(xml).toContain("产品标识");
    expect(xml).toContain("当前区间暂无需求");
  }, 30_000);
});
