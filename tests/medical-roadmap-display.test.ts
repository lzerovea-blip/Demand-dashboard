import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { MedicalSubmissionRoadmapBoard, RequirementEditor, SummaryPage } from "../src/renderer/App";
import { analysisWorkloadTotal, buildHalfYearSummaries, buildRoadmapGroups, medicalRequirementInRange, medicalRoadmapDemandRows, medicalRoadmapTimelineMonths, requirementsInRange } from "../src/shared/roadmap";
import { filterRequirementsForRegion } from "../src/shared/requirements";
import { MEDICAL_SUBMISSION_MILESTONE_FIELDS } from "../src/shared/medicalSubmission";
import { MEMBER_ROLES, type AppSnapshot, type Requirement } from "../src/shared/types";

const requirement: Requirement = {
  id: "medical", title: "独立需求标题", description: "这段详细描述不得出现在泳道左侧", source: "NMPA认证",
  targetMonth: "2028-02", medicalSubmissionMilestones: { registrationStartMonth: "2026-01", clinicalStartMonth: "2027-03", clinicalCompleteMonth: "2027-03", certificateMonth: "2029-05" },
  domainL0Id: "l0", domainId: "l1", scope: "全球", allowLatestVersionUpgrade: true, overseasRegions: [], category: "平台演进", productIds: [], images: [],
  deviceWorkloadPm: 1, appWorkloadPm: 2, researchAppWorkloadPm: 3, cloudWorkloadPm: 4, algorithmWorkloadPm: 5, unallocatedWorkloadPm: 0, workloadPm: 15,
  createdAt: "2026-09-09T00:00:00Z", updatedAt: "2026-09-09T00:00:00Z",
};
const snapshot: AppSnapshot = {
  requirements: [requirement], domains: [{ id: "l0", name: "健康", level: "L0", active: true, sortOrder: 0 }, { id: "l1", name: "心率", parentId: "l0", level: "L1", active: true, sortOrder: 0 }],
  products: [], members: [], groupOverrides: [], templates: [], pptAccessPolicy: { entries: [], updatedAt: "2026-09-09T00:00:00Z" },
};

describe("医疗路标五节点时间窗", () => {
  for (const source of ["NMPA认证", "CE认证"] as const) {
    it.each(["submissionMonth", ...MEDICAL_SUBMISSION_MILESTONE_FIELDS.map(item => item.key)])(`${source} 任一 %s 在窗内即显示`, (key) => {
      const row = { ...requirement, source, targetMonth: key === "submissionMonth" ? "2027-03" : "2028-02", medicalSubmissionMilestones: key === "submissionMonth" ? undefined : { [key]: "2027-03" } };
      expect(medicalRequirementInRange(row, "2027H1", "2027H1")).toBe(true);
      expect(buildRoadmapGroups({ ...snapshot, requirements: [row] }, "2027H1", "2027H1", { medicalDemandRows: true, medicalRange: "selected" })).toHaveLength(1);
    });
  }
  it("首末月包含、窗外排除；送检空但其他节点全在窗外不再误入，五节点全空保留", () => {
    for (const month of ["2027-01", "2027-06"]) expect(medicalRequirementInRange({ ...requirement, targetMonth: "", medicalSubmissionMilestones: { certificateMonth: month } }, "2027H1", "2027H1")).toBe(true);
    for (const month of ["2026-12", "2027-07", "2027-13"]) expect(medicalRequirementInRange({ ...requirement, targetMonth: "", medicalSubmissionMilestones: { certificateMonth: month } }, "2027H1", "2027H1")).toBe(false);
    expect(medicalRequirementInRange({ ...requirement, targetMonth: "", medicalSubmissionMilestones: undefined }, "2027H1", "2027H1")).toBe(true);
  });
  it("多个节点命中只显示一次、保持稳定 ID 和共享轴，统计/导出仍使用送检月份", () => {
    const groups = buildRoadmapGroups(snapshot, "2027H1", "2027H1", { medicalDemandRows: true, medicalRange: "selected" });
    expect(groups.map(i => i.key)).toEqual(["medical-requirement:medical"]);
    expect(medicalRoadmapTimelineMonths(groups, "2027H1", "2027H1")[0]).toBe("2027-01");
    expect(medicalRoadmapTimelineMonths(groups, "2027H1", "2027H1").at(-1)).toBe("2029-05");
    expect(requirementsInRange(snapshot.requirements, "2027H1", "2027H1")).toEqual([]);
    expect(buildRoadmapGroups(snapshot, "2027H1", "2027H1")).toEqual([]);
    expect(buildRoadmapGroups({ ...snapshot, requirements: [{ ...requirement, source: "健康基础" }] }, "2027H1", "2027H1", { medicalDemandRows: true })).toEqual([]);
  });
  it("国内外与时间窗 AND 筛选，排序和详情仍按需求 ID", () => {
    const rows: Requirement[] = [
      { ...requirement, id: "a", scope: "仅国内", medicalRoadmapSortOrder: 2 },
      { ...requirement, id: "b", scope: "仅海外", medicalRoadmapSortOrder: 1 },
      { ...requirement, id: "c", scope: "全球", medicalRoadmapSortOrder: 0 },
    ];
    for (const [region, expected] of [["domestic", ["c", "a"]], ["overseas", ["c", "b"]]] as const) {
      const groups = buildRoadmapGroups({ ...snapshot, requirements: filterRequirementsForRegion(rows, region) }, "2027H1", "2027H1", { medicalDemandRows: true, medicalRange: "selected" });
      expect(medicalRoadmapDemandRows(groups).map(i => i.requirement.id)).toEqual(expected);
    }
  });
});

describe("医疗路标真实组件输出及五项录入回归", () => {
  it("汇总页面排除算法金额，需求与导出既有总量不变", () => {
    const summaries = buildHalfYearSummaries([{ ...requirement, targetMonth: "2027-03", algorithmWorkloadPm: 9876, workloadPm: 9886 }], "2027H1", "2027H1");
    expect(analysisWorkloadTotal(summaries)).toBe(10);
    expect(summaries[0].totalWorkload).toBe(9886);
    const html = renderToStaticMarkup(createElement(SummaryPage, { summaries }));
    expect(html).toContain("总工作量（四项合计）");
    expect(html).toContain("算法工作量不纳入本页汇总分析");
    expect(html).not.toContain("9876");
    expect(html).not.toContain("9,876");
    expect(html).not.toContain("五项合计");
  });
  it("仅算法投入时分析为零，多半年小数四项相加且不包含待拆分", () => {
    expect(analysisWorkloadTotal(buildHalfYearSummaries([{ ...requirement, targetMonth: "2027-03", deviceWorkloadPm: 0, appWorkloadPm: 0, researchAppWorkloadPm: 0, cloudWorkloadPm: 0 }], "2027H1", "2027H1"))).toBe(0);
    const rows = ["2027-03", "2027-09"].map(targetMonth => ({ ...requirement, targetMonth, deviceWorkloadPm: 0.1, appWorkloadPm: 0.2, researchAppWorkloadPm: 0.3, cloudWorkloadPm: 0.4, unallocatedWorkloadPm: 99 }));
    expect(analysisWorkloadTotal(buildHalfYearSummaries(rows, "2027H1", "2027H2"))).toBe(2);
  });
  it.each([false, true])("全屏=%s：左侧只含标题，节点无重复年月，完整日期仍在无障碍描述", (focusMode) => {
    const html = renderToStaticMarkup(createElement(MedicalSubmissionRoadmapBoard, {
      title: "医疗认证路标", track: "医疗认证", groups: buildRoadmapGroups(snapshot, "2027H1", "2027H1", { medicalDemandRows: true }), start: "2027H1", end: "2027H1",
      focusMode, focusActive: true, region: "domestic", onRegionChange: () => {}, onOpenGroup: () => {}, onReorderGroups: async () => {}, onReorderMedicalRows: async () => {},
    }));
    expect(html).toContain("<strong>独立需求标题</strong>");
    expect(html).not.toContain(requirement.description);
    expect(html).toContain("临床启动 / 临床完成");
    expect(html).not.toMatch(/<small>\d{4}年/);
    expect(html).toContain("临床启动时间2027年3月");
    expect(html).toContain('data-roadmap-month-header="2027-03"');
    expect(html).not.toContain("is-l1-related");
  });
  it("编辑器渲染五项工作量及正确总量，并提供算法责任人", () => {
    const html = renderToStaticMarkup(createElement(RequirementEditor, { value: requirement, snapshot, onClose: () => {}, onSave: async () => {}, onDelete: async () => {}, requestConfirmation: async () => false }));
    for (const name of ["设备工作量", "运动健康 App 工作量", "创新研究 App 工作量", "云侧工作量", "算法工作量"]) expect(html).toContain(name);
    expect(html).toContain("合计 15 人月");
    expect(html).toContain('aria-label="需求算法"');
    expect(MEMBER_ROLES).toContain("算法");
  });
});
