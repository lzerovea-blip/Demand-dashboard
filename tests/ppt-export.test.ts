import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { createRoadmapPresentation } from "../src/main/pptxExport";
import { buildPptExportPlan, dynamicRoadmapMonthWidths } from "../src/shared/pptExport";
import type { AppSnapshot, Requirement } from "../src/shared/types";

const requirements: Requirement[] = ["需求A", "需求B", "需求C"].map((title, index) => ({
  id: `r${index + 1}`,
  title,
  description: `${title}的需求描述`,
  images: [],
  domainId: "d1",
  source: "健康基础",
  overseasRegions: [],
  category: index === 2 ? "产品专属" : "体验优化",
  targetMonth: "2027-03",
  productIds: index === 2 ? ["p1"] : [],
  deviceWorkloadPm: 1,
  appWorkloadPm: 0.5,
  cloudWorkloadPm: 0.25,
  unallocatedWorkloadPm: 0,
  workloadPm: 1.75,
  createdAt: "2026-08-20T00:00:00.000Z",
  updatedAt: `2026-08-20T00:00:0${index}.000Z`,
}));

const snapshot: AppSnapshot = {
  requirements,
  domains: [{ id: "d1", name: "睡眠", sortOrder: 0, active: true }],
  products: [{ id: "p1", name: "手表A", sortOrder: 0, active: true }],
  groupOverrides: [],
  templates: [],
};

describe("PPT导出计划", () => {
  it("按半年生成运动、健康、海外研究三路标并对详情分页", () => {
    const plan = buildPptExportPlan(snapshot, "2027H1", "2027H1");
    expect(plan.roadmapPages).toHaveLength(3);
    expect(plan.detailPages).toHaveLength(2);
    expect(plan.firstDetailSlideByGroup[plan.groups[0].key]).toBe(7);
    expect(plan.slideCount).toBe(8);
  });

  it("空月份固定缩窄并把剩余宽度分配给有需求月份", () => {
    const months = ["2027-01", "2027-02", "2027-03", "2027-04", "2027-05", "2027-06"];
    const widths = dynamicRoadmapMonthWidths(months, ["2027-03"], 10.04);
    expect(widths[0]).toBe(0.62);
    expect(widths[2]).toBeGreaterThan(6);
    expect(widths.reduce((sum, value) => sum + value, 0)).toBeCloseTo(10.04);
  });

  it("生成结构完整且可解压的PPTX", async () => {
    const bytes = await createRoadmapPresentation(snapshot, "2027H1", "2027H1");
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file("ppt/slides/slide1.xml")).toBeTruthy();
    expect(zip.file("ppt/slides/slide8.xml")).toBeTruthy();
    expect(zip.file("ppt/presentation.xml")).toBeTruthy();
  }, 20_000);
});
