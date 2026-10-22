import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { createRoadmapPresentation, groupRoadmapCardShapes, layoutRoadmapCell, layoutRoadmapLanes, measureTextWidth, paginateRoadmapPage, presentationSlideCount } from "../src/main/pptxExport";
import { validateGeneratedPptx } from "../src/main/pptxValidation";
import {
  DEFAULT_PPT_EXPORT_SECTIONS,
  MAX_PPT_HALF_YEARS,
  buildPptExportPlan,
  pptBaseSlideCount,
  roadmapBadgeKind,
  visibleRequirementTitles,
  type PptRoadmapPagePlan,
} from "../src/shared/pptExport";
import { buildRoadmapGroups, groupMatchesRoadmapLane, halfYearOf, roadmapLanesForTrack } from "../src/shared/roadmap";
import type { AppSnapshot, PptExportRequest, Requirement, RequirementCategory, RequirementSource, Track } from "../src/shared/types";

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
    requirement("7", "NMPA资料", "NMPA认证", "2027-02", "平台演进", [], "d2"),
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

it("新增工作量仅影响汇总总量，不改变简化封面、三端图表或路标页样式", async () => {
  const extended: AppSnapshot = { ...snapshot, requirements: snapshot.requirements.map((item) => ({ ...item, researchAppWorkloadPm: 2, algorithmWorkloadPm: 3, workloadPm: item.workloadPm + 5 })) };
  const basePlan = buildPptExportPlan(snapshot, fullRequest);
  const newPlan = buildPptExportPlan(extended, fullRequest);
  expect(newPlan.summaries.reduce((sum, item) => sum + item.totalWorkload, 0)).toBe(basePlan.summaries.reduce((sum, item) => sum + item.totalWorkload, 0) + newPlan.requirements.length * 5);
  expect(newPlan.summaries.map((item) => item.bySide)).toEqual(basePlan.summaries.map((item) => item.bySide));
  const baseline = await JSZip.loadAsync(await createRoadmapPresentation(snapshot, fullRequest));
  const output = await JSZip.loadAsync(await createRoadmapPresentation(extended, fullRequest));
  const slide2 = await output.file("ppt/slides/slide2.xml")!.async("string");
  expect(slide2).toContain("创新研究 App");
  expect(slide2).toContain("未计入下方三端图表");
  expect(slide2).toContain("运动健康 App");
  expect(await output.file("ppt/slides/slide1.xml")!.async("string")).toBe(await baseline.file("ppt/slides/slide1.xml")!.async("string"));
  for (let page = 3; page <= presentationSlideCount(newPlan); page++) expect(await output.file(`ppt/slides/slide${page}.xml`)!.async("string")).toBe(await baseline.file(`ppt/slides/slide${page}.xml`)!.async("string"));
});

function denseSnapshot(source: RequirementSource = "健康基础", count = 18): AppSnapshot {
  return {
    ...snapshot,
    domains: [snapshot.domains[0], ...Array.from({ length: count }, (_, index) => ({
      id: `dense-${index}`, name: `密集领域${index + 1}`, sortOrder: index, active: true,
      level: "L1" as const, parentId: "l0",
    }))],
    requirements: Array.from({ length: count }, (_, index) => requirement(
      String(100 + index), `需求${index + 1}：社交时差、体动与连续英文LongUnbrokenTitle`, source,
      "2026-03", index % 3 === 0 ? "平台演进" : "体验优化", index % 3 === 1 ? ["p1"] : [], `dense-${index}`,
    )),
  };
}

function cellLayouts(page: PptRoadmapPagePlan) {
  const table = layoutRoadmapLanes(page);
  return table.lanes.flatMap((lane) => page.halfYears.map((half) => ({
    height: lane.height,
    layout: layoutRoadmapCell(page.groups.filter((group) => groupMatchesRoadmapLane(group, lane.id) && halfYearOf(group.targetMonth) === half), table.halfWidth, lane.height),
  })));
}

describe("PPT导出计划", () => {
  it("医疗认证按区域独立全量范围、十二个月及泳道高度续页", async () => {
    const data = denseSnapshot("NMPA认证", 10);
    data.requirements.forEach((item, index) => {
      item.targetMonth = "2030-05";
      item.medicalSubmissionMilestones = { registrationStartMonth: "2028-11", certificateMonth: "2030-06" };
      item.category = "平台演进";
      item.productIds = ["p1"];
      item.title = `独立认证需求${index}`;
    });
    const plan = buildPptExportPlan(data, { ...fullRequest, sections: ["医疗认证"] });
    expect(plan.requirements).toHaveLength(0);
    expect(plan.roadmapPages[0].groups).toHaveLength(10);
    expect(plan.roadmapPages[0].medicalMonths?.[0]).toBe("2028-11");
    expect(plan.roadmapPages[0].medicalMonths?.at(-1)).toBe("2030-06");
    const pages = paginateRoadmapPage(plan.roadmapPages[0]);
    expect(pages).toHaveLength(4);
    expect(pages.map((page) => page.medicalMonths?.length)).toEqual([12, 8, 12, 8]);
    expect(pages[0].groups).toEqual(pages[1].groups);
    const zip = await JSZip.loadAsync(await createRoadmapPresentation(data, { ...fullRequest, sections: ["医疗认证"] }));
    expect(Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))).toHaveLength(9);
    const xml = await zip.file("ppt/slides/slide2.xml")!.async("string");
    expect(xml).toContain("独立认证需求");
    expect(xml).not.toContain(data.products[0].name);
    expect(xml).toContain("国内");
    expect(xml).toContain("注册启动");
    expect(xml).not.toContain("的描述");
    expect(xml).not.toMatch(/\b(?:NaN|Infinity)\b/);
  });
  it("未定送检版本不进入封面数量、汇总或 PPT 卡片，其他输出保持一致", () => {
    const original = buildPptExportPlan(snapshot, fullRequest);
    const undated = requirement("undated-medical", "待定送检", "NMPA认证", "");
    const plan = buildPptExportPlan({ ...snapshot, requirements: [...snapshot.requirements, undated] }, fullRequest);
    expect(plan.requirements).toEqual(original.requirements);
    expect(plan.summaries).toEqual(original.summaries);
    expect(plan.roadmapPages.find((page) => page.track === "医疗认证")!.groups.some((group) => group.requirements[0].id === undated.id)).toBe(true);
    expect(presentationSlideCount(plan)).toBe(presentationSlideCount(original));
  });

  it("全选为封面、汇总及七张路标各国内海外两页", () => {
    const plan = buildPptExportPlan(snapshot, fullRequest);
    expect(plan.slideCount).toBe(16);
    expect(presentationSlideCount(plan)).toBe(16);
    expect(plan.roadmapPages.map((item) => item.track)).toEqual(["运动", "健康", "行业", "全场景", "health平台", "AI", "医疗认证"].flatMap((track) => [track, track]));
    expect(plan.roadmapPages.map((item) => item.region)).toEqual(Array.from({ length: 7 }, () => ["domestic", "overseas"]).flat());
    expect(plan.roadmapPages.find((item) => item.track === "行业")?.groups).toEqual([]);
  });

  it("同领域同月份先按区域筛选再分组，全球两页可见但汇总只计一次", async () => {
    const data: AppSnapshot = { ...snapshot, requirements: [
      { ...requirement("21", "仅国内条目", "健康基础", "2026-03"), scope: "仅国内" },
      { ...requirement("22", "仅海外条目", "健康基础", "2026-03"), scope: "仅海外" },
      requirement("23", "全球条目", "健康基础", "2026-03"),
    ] };
    const request = { ...fullRequest, sections: ["summary", "健康"] };
    const plan = buildPptExportPlan(data, request);
    expect(plan.roadmapPages.map((page) => page.groups.flatMap((group) => group.requirements.map((item) => item.id)))).toEqual([["21", "23"], ["22", "23"]]);
    expect(plan.requirements).toHaveLength(3);
    expect(plan.summaries.reduce((total, item) => total + item.totalWorkload, 0)).toBe(5.25);
    const zip = await JSZip.loadAsync(await createRoadmapPresentation(data, request));
    const domestic = await zip.file("ppt/slides/slide3.xml")!.async("string");
    const overseas = await zip.file("ppt/slides/slide4.xml")!.async("string");
    expect(domestic).toContain("仅国内条目");
    expect(domestic).not.toContain("仅海外条目");
    expect(overseas).toContain("仅海外条目");
    expect(overseas).not.toContain("仅国内条目");
    expect([domestic, overseas].every((xml) => xml.includes("全球条目"))).toBe(true);
  });

  it("没有区域需求时仍保留该区域空状态页", async () => {
    const data: AppSnapshot = { ...snapshot, requirements: [{ ...requirement("24", "国内需求", "健康基础", "2026-03"), scope: "仅国内" }] };
    const request = { ...fullRequest, sections: ["健康"] };
    const plan = buildPptExportPlan(data, request);
    expect(plan.roadmapPages[1].groups).toEqual([]);
    expect(presentationSlideCount(plan)).toBe(3);
    const zip = await JSZip.loadAsync(await createRoadmapPresentation(data, request));
    const overseas = await zip.file("ppt/slides/slide3.xml")!.async("string");
    expect(overseas).toContain("海外");
    expect(overseas).toContain("当前区间暂无需求");
    expect(overseas).not.toContain("国内需求");
  });

  it("医疗认证分别使用国内海外需求的独立节点范围", () => {
    const data: AppSnapshot = { ...snapshot, requirements: [
      { ...requirement("25", "国内认证", "NMPA认证", "2026-03"), scope: "仅国内", medicalSubmissionMilestones: { registrationStartMonth: "2025-01", certificateMonth: "2026-04" } },
      { ...requirement("26", "海外认证", "CE认证", "2028-05"), scope: "仅海外", medicalSubmissionMilestones: { registrationStartMonth: "2028-02", certificateMonth: "2029-06" } },
    ] };
    const plan = buildPptExportPlan(data, { ...fullRequest, sections: ["医疗认证"] });
    expect(plan.roadmapPages.map((page) => [page.medicalMonths?.[0], page.medicalMonths?.at(-1)])).toEqual([["2025-01", "2026-04"], ["2028-02", "2029-06"]]);
    expect(plan.roadmapPages.map((page) => page.groups.map((group) => group.requirements[0].id))).toEqual([["25"], ["26"]]);
  });

  it("部分选择页数等于固定封面加选中项且保持固定顺序", () => {
    const plan = buildPptExportPlan(snapshot, { ...fullRequest, sections: ["医疗认证", "summary", "运动"] });
    expect(plan.selectedSections).toEqual(["summary", "运动", "医疗认证"]);
    expect(plan.slideCount).toBe(6);
    expect(pptBaseSlideCount([])).toBe(1);
    expect(pptBaseSlideCount(["summary"])).toBe(2);
    expect(pptBaseSlideCount(plan.selectedSections)).toBe(6);
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

  it("正文固定7pt并在极限情况下压缩需求数", () => {
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
    expect(layout.fontSize).toBe(7);
    expect(layout.cards.every((card) => card.top + card.height <= 170.95)).toBe(true);
    expect(measureTextWidth("连续英文ABC123", 10, true)).toBeGreaterThan(0);
  });

  it.each([1, 2, 3, 4])("%s 个半年时卡片限制宽度、长标题不侵入正文且产品不占空间", (halves) => {
    const data = { ...snapshot, products: [
      { id: "p1", name: "WATCHVeryLongUnbrokenProductName", active: true, sortOrder: 0 },
      { id: "p2", name: "手表二", active: true, sortOrder: 1 },
      { id: "p3", name: "手表三", active: true, sortOrder: 2 },
    ], requirements: [requirement("99", "混排标题WWW连续健康监测、趋势与恢复", "健康基础", "2026-03", "平台演进", ["p1", "p2", "p3"])] };
    const group = buildRoadmapGroups(data, "2026H1", "2026H1")[0];
    group.cardTitle = "很长的领域标题LongDomain";
    const cellWidth = (1223.07 - 128) / halves;
    const card = layoutRoadmapCell([group], cellWidth, 512).cards[0];
    expect(card.width).toBeLessThanOrEqual(260);
    expect(card.width + card.margin * 2).toBeLessThanOrEqual(cellWidth);
    expect(card.titleFontSize).toBe(8);
    expect(card.titleLines.join("")).toBe(group.cardTitle);
    expect(card.titleTop + card.titleLines.length * card.titleLineHeight).toBeLessThanOrEqual(card.height - 4);
    expect(card.bodyX).toBeGreaterThan(card.infoX + card.infoWidth);
    expect(card.bodyX + card.bodyWidth).toBeLessThanOrEqual(card.width - 6);
    expect(card).not.toHaveProperty("productPills");
    expect(card.bodyLines.every((line) => measureTextWidth(line, 7) <= card.bodyWidth)).toBe(true);
  });

  it("左右分栏保留摘要省略规则，领域名计入高度，产品不再影响高度", () => {
    const group = buildRoadmapGroups(snapshot, "2026H1", "2026H1")[0];
    const normal = layoutRoadmapCell([group], 273.77, 512);
    expect(normal.cards[0].visibleRequirements).toEqual(["社交时差、体动", "睡眠监测", "夜间恢复…"]);
    const tall = { ...group, cardTitle: "很长的领域名称".repeat(5), productNames: ["手表A", "手表B", "手表C"] };
    const tallCard = layoutRoadmapCell([tall], 273.77, 512).cards[0];
    expect(tallCard.height).toBeGreaterThan(normal.cards[0].height);
    expect(tallCard.height).toBe(layoutRoadmapCell([{ ...tall, productNames: [] }], 273.77, 512).cards[0].height);
  });

  it("同一半年三泳道布局中的单格容纳四类大需求和八条子需求，字号一致且无摘要遗漏", () => {
    const data = denseSnapshot("健康基础", 4);
    data.requirements.forEach((item, index) => { item.title = `子需求${index + 1}之一`; });
    data.requirements.push(...data.requirements.map((item, index) => ({ ...item, id: `child-${item.id}`, title: `子需求${index + 1}之二` })));
    const page = buildPptExportPlan(data, { ...fullRequest, sections: ["健康"] }).roadmapPages[0];
    const parts = paginateRoadmapPage(page);
    expect(parts).toHaveLength(1);
    const layout = layoutRoadmapCell(page.groups, 273.7675, (580.87 - 68) / 3);
    expect(layout.cards).toHaveLength(4);
    expect(layout.cards.flatMap((card) => card.visibleRequirements)).toHaveLength(8);
    expect(new Set(layout.cards.map((card) => card.fontSize))).toEqual(new Set([7]));
    expect(new Set(layout.cards.map((card) => card.titleFontSize))).toEqual(new Set([8]));
    for (const card of layout.cards) {
      expect(card.visibleRequirements.every((value) => !value.endsWith("…"))).toBe(true);
      expect(card.top + card.height).toBeLessThanOrEqual(170.96);
    }
  });
});

describe("密集路标自动续页", () => {
  it.each<[Track, RequirementSource]>([
    ["运动", "运动高阶"], ["健康", "健康基础"], ["行业", "行业"], ["全场景", "全场景"],
    ["health平台", "health平台"], ["AI", "AI健康"], ["医疗认证", "CE认证"],
  ])("%s 在不同泳道高度下保留全部卡片、顺序、标识和表头", (track, source) => {
    const plan = buildPptExportPlan(denseSnapshot(source, 30), { ...fullRequest, sections: [track] });
    const page = plan.roadmapPages[0];
    const pages = paginateRoadmapPage(page);
    expect(pages.length).toBeGreaterThan(1);
    expect(pages.flatMap((part) => part.groups).map((group) => group.key).sort()).toEqual(page.groups.map((group) => group.key).sort());
    expect(presentationSlideCount(plan)).toBe(1 + pages.length * 2);
    pages.forEach((part, index) => {
      expect(part.continuationIndex).toBe(index);
      expect(part.continuationCount).toBe(pages.length);
      expect(part.halfYears).toEqual(page.halfYears);
      expect(part.track).toBe(track);
      if (track === "医疗认证") { expect(part.medicalMonths).toEqual(["2026-03"]); return; }
      for (const { height, layout } of cellLayouts(part)) {
      expect(layout.fontSize).toBe(7);
      layout.cards.forEach((card, cardIndex) => {
        expect(card.top).toBeGreaterThanOrEqual(layout.margin);
        expect(card.top + card.height).toBeLessThanOrEqual(height - layout.margin + 0.01);
        expect(card.fontSize).toBe(7);
        expect(card).not.toHaveProperty("badge");
        expect(card).not.toHaveProperty("badgeTop");
        card.bodyLines.forEach((line) => expect(measureTextWidth(line, card.fontSize)).toBeLessThanOrEqual(card.bodyWidth));
        if (cardIndex) expect(card.top).toBeGreaterThanOrEqual(layout.cards[cardIndex - 1].top + layout.cards[cardIndex - 1].height + 3);
      });
      }
    });
  });

  it.each(["2026H1", "2026H2", "2027H2"])("%s 为结束半年时按实际列宽分页", (end) => {
    const plan = buildPptExportPlan(denseSnapshot(), { ...fullRequest, end, sections: ["健康"] });
    const pages = paginateRoadmapPage(plan.roadmapPages[0]);
    expect(pages.length).toBeGreaterThan(1);
    expect(pages.flatMap((page) => page.groups)).toEqual(plan.groups);
    pages.forEach((page) => expect(() => cellLayouts(page)).not.toThrow());
  });

  it("各半年和泳道独立分页，不重复卡片，仍保留同名单元格内的排序", () => {
    const dense = denseSnapshot();
    dense.domains[2].name = dense.domains[1].name;
    dense.requirements.push(
      ...dense.requirements.slice(0, 9).map((item) => ({ ...item, id: `next-${item.id}`, targetMonth: "2027-04" })),
      ...dense.requirements.slice(0, 7).map((item) => ({ ...item, id: `advanced-${item.id}`, source: "健康高阶" as const })),
    );
    const plan = buildPptExportPlan(dense, { ...fullRequest, sections: ["健康"] });
    const page = plan.roadmapPages[0];
    const orderedKey = page.groups.find((group) => group.domainId === "dense-8" && group.source === "健康基础" && group.targetMonth === "2026-03")!.key;
    dense.groupOverrides = [{ groupKey: orderedKey, cardTitle: "密集领域9", cardSummary: "", sortOrder: 0, updatedAt: "2026-09-03T00:00:00.000Z" }];
    const orderedPage = buildPptExportPlan(dense, { ...fullRequest, sections: ["健康"] }).roadmapPages[0];
    const parts = paginateRoadmapPage(orderedPage);
    const output = parts.flatMap((part) => part.groups);
    expect(new Set(output.map((group) => group.key)).size).toBe(orderedPage.groups.length);
    for (const lane of roadmapLanesForTrack("健康")) {
      for (const halfYear of orderedPage.halfYears) {
        const inCell = (group: typeof page.groups[number]) => groupMatchesRoadmapLane(group, lane) && halfYearOf(group.targetMonth) === halfYear;
        expect(output.filter(inCell)).toEqual(orderedPage.groups.filter(inCell));
      }
    }
    expect(parts.every((part) => part.groups.length > 0)).toBe(true);
    expect(output.filter((group) => group.source === "健康基础" && group.targetMonth === "2026-03")[0].key).toBe(orderedKey);
  });

  it("空路标保留一页，单张超长卡片准确报出位置且不静默丢失", () => {
    const empty = buildPptExportPlan(snapshot, { ...fullRequest, sections: ["行业"] }).roadmapPages[0];
    expect(paginateRoadmapPage(empty)).toEqual([{ ...empty, continuationIndex: 0, continuationCount: 1 }]);
    const oversized = denseSnapshot("健康基础", 1);
    oversized.requirements[0].title = "超长需求".repeat(500);
    const plan = buildPptExportPlan(oversized, { ...fullRequest, sections: ["健康"] });
    expect(() => paginateRoadmapPage(plan.roadmapPages[0])).toThrow(/单张卡片.*健康路标.*基础.*2026H1.*密集领域1/);
  });

  it("密集长中文、全角标点和连续英文的实绘行与容量计算一致", () => {
    const dense = denseSnapshot();
    dense.requirements[2].title = "连续健康监测与长标题换行验证：社交时差、体动与趋势分析";
    dense.requirements[3].title = "WWWMMMMVeryLongUnbrokenRequirementTitle".repeat(3);
    const page = buildPptExportPlan(dense, { ...fullRequest, sections: ["健康"] }).roadmapPages[0];
    const cards = paginateRoadmapPage(page).flatMap((part) => cellLayouts(part).flatMap((cell) => cell.layout.cards));
    for (const card of cards) {
      expect(card.bodyOffset + card.bodyLines.length * card.bodyLineHeight).toBeLessThanOrEqual(card.height - 4 + 0.01);
      expect(card.bodyLines[0]).toMatch(/^· \S/);
      expect(card.bodyLines.every((line) => measureTextWidth(line, card.fontSize) <= card.bodyWidth)).toBe(true);
    }
  });
});

describe("PPTX输出", () => {
  it("密集完整材料页数、页码与校验一致，续页重复表头且每张卡片仅出现一次", async () => {
    const dense = denseSnapshot();
    const plan = buildPptExportPlan(dense, fullRequest);
    const expectedCount = presentationSlideCount(plan);
    const bytes = await createRoadmapPresentation(dense, fullRequest);
    expect(expectedCount).toBeGreaterThan(16);
    expect((await validateGeneratedPptx(bytes, expectedCount)).slides).toBe(expectedCount);
    const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
    const normalZip = await JSZip.loadAsync(await createRoadmapPresentation(snapshot, fullRequest));
    expect(await zip.file("ppt/theme/theme1.xml")!.async("string")).toEqual(await normalZip.file("ppt/theme/theme1.xml")!.async("string"));
    const xmls = await Promise.all(Array.from({ length: expectedCount }, (_, index) => zip.file(`ppt/slides/slide${index + 1}.xml`)!.async("string")));
    const roadmaps = xmls.slice(2);
    roadmaps.forEach((xml) => {
      if (!xml.includes("医疗认证路标")) {
        for (const header of ["26H1", "26H2", "27H1", "27H2", "HMOS 6.1", "HMOS 8.0"]) expect(xml).toContain(header);
        expect(xml).toContain("<a:tbl>");
      }
    });
    const xml = xmls.join("\n");
    const cardBodies = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map((match) => match[0]).filter((shape) => shape.includes('name="roadmap-body-'));
    expect(cardBodies.length).toBeGreaterThan(0);
    cardBodies.forEach((body) => {
      expect(body).toContain('wrap="none"');
      expect(body).toContain("<a:lnSpc>");
    });
    expect(xmls.filter((slide) => slide.includes("健康路标"))[0].split("<a:t>基础</a:t>")).toHaveLength(2);
    const titleShapes = xml.match(/<p:sp>.*?<\/p:sp>/gs)!.filter((shape) => shape.includes('name="roadmap-title-'));
    const renderedTitles = titleShapes.map((shape) => [...shape.matchAll(/<a:t>(.*?)<\/a:t>/gs)].map((match) => match[1]).join(""));
    for (const domain of dense.domains.slice(1)) expect(renderedTitles.filter((title) => title === domain.name)).toHaveLength(2);
    expect(xml).toContain("国内（续 2/");
    expect(xml).toContain("产品专属");
    expect(xml.includes("↑")).toBe(false);
    expect(xml).not.toContain('name="roadmap-product-');
    expect(xml).not.toContain('name="roadmap-experience-');
    const roadmapShapes = xml.match(/<p:sp>.*?<\/p:sp>/gs)!.filter((shape) => /name="roadmap-(?:title|body)-/.test(shape));
    expect(roadmapShapes.length).toBeGreaterThan(0);
    for (const shape of roadmapShapes) {
      const size = shape.includes('name="roadmap-title-') ? "800" : "700";
      expect([...shape.matchAll(/\bsz="(\d+)"/g)].every((match) => match[1] === size)).toBe(true);
      expect(shape.includes('name="roadmap-title-') ? shape.includes('b="1"') : !shape.includes('b="1"')).toBe(true);
    }
    expect(xml).not.toMatch(/(?:NaN|Infinity|七类路标投入|领域详情|>\d{1,2}月<)/);
  }, 30_000);

  it("生成原生可编辑的完整16页文件且无详情页、普通路标月份标题或内部制作文字", async () => {
    const bytes = await createRoadmapPresentation(snapshot, fullRequest);
    const zip = await JSZip.loadAsync(bytes);
    const slideFiles = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
    expect(slideFiles).toHaveLength(16);
    expect(zip.file("ppt/theme/theme1.xml")).toBeTruthy();
    const xml = (await Promise.all(slideFiles.map((name) => zip.file(name)!.async("string")))).join("\n");
    const allXml = (await Promise.all(Object.keys(zip.files)
      .filter((name) => /(?:\.xml|\.rels)$/.test(name))
      .map((name) => zip.file(name)!.async("string")))).join("\n");
    expect(xml).toContain("需求工作量汇总分析");
    expect(xml).toContain("三端需求结构");
    expect(xml).not.toContain("七类路标投入");
    expect(xml).not.toContain("领域详情");
    expect(xml).not.toContain("制作说明");
    expect(allXml).not.toMatch(/\b(?:NaN|Infinity)\b/);
    for (const name of slideFiles) {
      const slideXml = await zip.file(name)!.async("string");
      if (!slideXml.includes("医疗认证路标")) expect(slideXml).not.toMatch(/>\d{1,2}月</);
    }
    expect(xml).toContain("社交时差、体动");
    expect(xml).toContain("HMOS 6.1");
    expect(xml).toContain("HMOS 8.0");
    expect(xml).toContain("产品专属");
    expect(xml).not.toContain('name="roadmap-product-');
    expect(xml).not.toContain('name="roadmap-experience-');
    expect(xml).toContain("当前区间暂无需求");
  }, 30_000);

  it("封面仅含标题，全部内容页及续页均无权限日期页脚和右上角页码", async () => {
    const data = denseSnapshot();
    const security = { exportedAt: new Date("2026-10-20T03:00:00Z"), expiration: new Date("2027-04-20T03:00:00Z") };
    const plan = buildPptExportPlan(data, fullRequest);
    const zip = await JSZip.loadAsync(await createRoadmapPresentation(data, fullRequest, security));
    const cover = await zip.file("ppt/slides/slide1.xml")!.async("string");
    expect([...cover.matchAll(/<a:t>(.*?)<\/a:t>/g)].map((match) => match[1])).toEqual([fullRequest.materialName]);
    expect(cover).toContain('<a:srgbClr val="FFFFFF"/>');
    for (let index = 2; index <= presentationSlideCount(plan); index++) {
      const slide = await zip.file(`ppt/slides/slide${index}.xml`)!.async("string");
      expect(slide).not.toContain("irm-expiration-footer");
      expect(slide).not.toContain("企业 IRM 权限有效期至");
      const shapes = slide.match(/<p:sp>.*?<\/p:sp>/gs) ?? [];
      expect(shapes.some((shape) => shape.includes('x="11144250" y="247650"'))).toBe(false);
    }
    const coverOnly = await createRoadmapPresentation(data, { ...fullRequest, sections: [] }, security);
    expect((await validateGeneratedPptx(coverOnly, 1)).slides).toBe(1);
  });

  it("医疗同月节点和认证标识保持原生且7pt一致，产品及重复年月不再出现", async () => {
    const data = { ...snapshot, requirements: [{ ...snapshot.requirements[6], productIds: ["p1"], medicalSubmissionMilestones: {
      registrationStartMonth: "2026-12", clinicalStartMonth: "2027-02", clinicalCompleteMonth: "2027-02", certificateMonth: "2027-04",
    } }] };
    const bytes = await createRoadmapPresentation(data, { ...fullRequest, sections: ["医疗认证"] });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("ppt/slides/slide2.xml")!.async("string");
    expect(xml).not.toContain('val="F5F7F9"');
    expect(xml).not.toContain('val="C00000"');
    expect(xml).toContain('val="DB5A4C"');
    expect(xml).toContain('val="3989DC"');
    expect(xml).not.toContain('name="medical-product-');
    expect(xml).not.toContain("手表A");
    const nodeShapes = xml.match(/<p:sp>.*?<\/p:sp>/gs)!.filter((shape) => shape.includes('name="medical-node-'));
    expect(nodeShapes).toHaveLength(3);
    expect(nodeShapes.join("")).toContain("送检");
    expect(nodeShapes.join("")).toContain("临床启动");
    const nodeText = nodeShapes.flatMap((shape) => [...shape.matchAll(/<a:t>(.*?)<\/a:t>/g)].map((match) => match[1])).join(" ");
    expect(nodeText).not.toMatch(/2026|2027|\d月/);
    for (const shape of nodeShapes) expect([...shape.matchAll(/\bsz="(\d+)"/g)].every((match) => match[1] === "700")).toBe(true);
    expect(xml).not.toContain("NMPA资料的描述");
  });

  it("空状态路标输出紧凑深红底白字单行半年表头和居中泳道", async () => {
    const bytes = await createRoadmapPresentation(snapshot, { ...fullRequest, sections: ["行业"] });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("ppt/slides/slide2.xml")!.async("string");
    expect(xml).toContain("行业路标");
    expect(xml).toContain("26H1");
    expect(xml).toContain("HMOS 6.1");
    expect(xml).not.toContain("产品标识");
    expect(xml).toContain("当前区间暂无需求");
    const headers = (xml.match(/<p:sp>.*?<\/p:sp>/gs) ?? []).filter((shape) => shape.includes('name="roadmap-half-'));
    expect(headers).toHaveLength(4);
    for (const header of headers) {
      expect([...header.matchAll(/<a:t>(.*?)<\/a:t>/g)]).toHaveLength(1);
      expect(header).toMatch(/<a:t>\d{2}H[12]（HMOS [\d.]+）<\/a:t>/);
      expect(header).toContain('val="FFFFFF"');
      expect(header).toContain('cy="228600"');
    }
    expect(xml).toContain('val="8B1E2D"');
    expect(xml).toContain('y="590550"');
    const lane = (xml.match(/<p:sp>.*?<\/p:sp>/gs) ?? []).find((shape) => shape.includes('name="roadmap-lane-'))!;
    expect(lane).toContain('anchor="ctr"');
    expect(lane).toContain('algn="ctr"');
    expect(xml).not.toContain('y="609600"');
    expect(xml).not.toContain('val="F5F7F9"');
  }, 30_000);

  it("每个需求块是独立可编辑组合，五个子对象位置保持不变且ID唯一", async () => {
    const plan = buildPptExportPlan(snapshot, { ...fullRequest, sections: ["健康"] });
    const zip = await JSZip.loadAsync(await createRoadmapPresentation(snapshot, { ...fullRequest, sections: ["健康"] }));
    const xml = await zip.file("ppt/slides/slide2.xml")!.async("string");
    const groups = xml.match(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g) ?? [];
    expect(groups).toHaveLength(plan.roadmapPages[0].groups.length);
    for (const group of groups) {
      expect(group).toContain('name="roadmap-card-');
      expect(group.match(/<p:sp>/g)).toHaveLength(5);
      expect(group).toContain('name="roadmap-title-');
      expect(group).toContain('name="roadmap-body-');
      const bounds = group.match(/<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/)!;
      expect(group).toContain(`<a:chOff x="${bounds[1]}" y="${bounds[2]}"/><a:chExt cx="${bounds[3]}" cy="${bounds[4]}"/>`);
    }
    const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((match) => match[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(groupRoadmapCardShapes(xml)).toBe(xml);
  });
});

describe("泳道自适应高度", () => {
  it("整表高度保持不变，密集泳道借用均分高度且不用续页", () => {
    const data = denseSnapshot("健康高阶", 10);
    data.requirements.forEach((item) => { item.title = "短需求"; });
    data.requirements.push(requirement("small", "基础需求", "健康基础", "2027-03"));
    const page = buildPptExportPlan(data, { ...fullRequest, sections: ["健康"] }).roadmapPages[0];
    expect(paginateRoadmapPage(page)).toHaveLength(1);
    const table = layoutRoadmapLanes(page);
    expect(table.lanes[0].height).toBeGreaterThan(300);
    expect(table.lanes[1].height).toBeLessThan(table.bodyCapacity / 3);
    expect(table.lanes[1].height).toBeGreaterThanOrEqual(48);
    expect(table.lanes[2].height).toBeCloseTo(table.lanes[1].height);
    expect(table.lanes[1].top).toBe(table.lanes[0].top + table.lanes[0].height);
    expect(table.height).toBe(618);
    expect(table.lanes.reduce((sum, lane) => sum + lane.height, 0)).toBeCloseTo(table.bodyCapacity);
    expect(() => cellLayouts(page)).not.toThrow();
  });

  it("各半年取同泳道的最大内容高度，长领域、长正文保留统一字体并准确续页", () => {
    const data = denseSnapshot("健康高阶", 26);
    data.requirements.push(...data.requirements.slice(0, 8).map((item) => ({ ...item, id: `basic-${item.id}`, source: "健康基础" as const, targetMonth: "2027-09" })));
    const page = buildPptExportPlan(data, { ...fullRequest, sections: ["健康"] }).roadmapPages[0];
    const pages = paginateRoadmapPage(page);
    expect(pages.length).toBeGreaterThan(1);
    expect(pages.flatMap((part) => part.groups).map((group) => group.key).sort()).toEqual(page.groups.map((group) => group.key).sort());
    for (const part of pages) {
      const table = layoutRoadmapLanes(part);
      expect(table.height).toBe(618);
      expect(table.lanes.at(-1)!.top + table.lanes.at(-1)!.height).toBeCloseTo(618);
      for (const { height, layout } of cellLayouts(part)) {
        for (const card of layout.cards) {
          expect(card.fontSize).toBe(7);
          expect(card.titleFontSize).toBe(8);
          expect(card.top + card.height).toBeLessThanOrEqual(height - 4 + 0.01);
        }
      }
    }
  });

  it("空页和长泳道名称均保持整表高度，空单泳道使用全部正文区域", () => {
    const page = buildPptExportPlan(snapshot, { ...fullRequest, sections: ["行业"] }).roadmapPages[0];
    expect(layoutRoadmapLanes(page).height).toBe(618);
    expect(layoutRoadmapLanes(page).lanes[0].height).toBe(586);
    const long = { ...page, lanes: [{ id: "lane-long", name: "特别长的自定义泳道名称用于核验居中显示与高度", roadmapId: page.track, sortOrder: 0, updatedAt: "2026-10-21T00:00:00.000Z" }] };
    expect(layoutRoadmapLanes(long).lanes[0].height).toBeGreaterThan(48);
  });

  it.each([1, 2, 3, 4])("%s个半年下，稀疏三泳道均衡占满正文，不将整表收短", (halves) => {
    const page = buildPptExportPlan(snapshot, { ...fullRequest, sections: ["运动"] }).roadmapPages[0];
    const table = layoutRoadmapLanes({ ...page, halfYears: page.halfYears.slice(0, halves) });
    expect(table.height).toBe(618);
    expect(table.lanes).toHaveLength(3);
    table.lanes.forEach((lane) => expect(lane.height).toBeCloseTo(586 / 3));
    expect(table.lanes.at(-1)!.top + table.lanes.at(-1)!.height).toBeCloseTo(618);
  });

  it("两个密集泳道同时借高，仍为第三泳道保留完整内容及最低高度", () => {
    const data = denseSnapshot("健康高阶", 7);
    data.requirements.forEach((item) => { item.title = "短需求"; });
    data.requirements.push(...data.requirements.map((item) => ({ ...item, id: `advanced-${item.id}`, source: "健康进阶" as const })));
    const page = buildPptExportPlan(data, { ...fullRequest, sections: ["健康"] }).roadmapPages[0];
    expect(paginateRoadmapPage(page)).toHaveLength(1);
    const table = layoutRoadmapLanes(page);
    expect(table.lanes[0].height).toBeGreaterThan(586 / 3);
    expect(table.lanes[1].height).toBeCloseTo(table.lanes[0].height);
    expect(table.lanes[2].height).toBeGreaterThanOrEqual(48);
    expect(table.lanes[2].height).toBeLessThan(586 / 3);
    expect(table.height).toBe(618);
    expect(() => cellLayouts(page)).not.toThrow();
  });

  it("实际输出表格及所有续页的行高合计固定为618px，组合和字号保持不变", async () => {
    const data = denseSnapshot("健康高阶", 30);
    const request = { ...fullRequest, sections: ["健康"] };
    const plan = buildPptExportPlan(data, request);
    const zip = await JSZip.loadAsync(await createRoadmapPresentation(data, request));
    for (let page = 2; page <= presentationSlideCount(plan); page++) {
      const xml = await zip.file(`ppt/slides/slide${page}.xml`)!.async("string");
      const table = xml.match(/<a:tbl>[\s\S]*?<\/a:tbl>/)![0];
      const heights = [...table.matchAll(/<a:tr h="(\d+)"/g)].map((match) => Number(match[1]) / 9525);
      expect(heights[0]).toBe(32);
      expect(heights.reduce((sum, value) => sum + value, 0)).toBeCloseTo(618, 3);
      expect(xml).toContain('<p:grpSp>');
      expect(xml).toContain('val="8B1E2D"');
      const textShapes = (xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? []).filter((shape) => /name="roadmap-(?:title|body)-/.test(shape));
      for (const shape of textShapes) {
        const size = shape.includes('name="roadmap-title-') ? "800" : "700";
        expect([...shape.matchAll(/\bsz="(\d+)"/g)].every((match) => match[1] === size)).toBe(true);
      }
    }
  });
});
