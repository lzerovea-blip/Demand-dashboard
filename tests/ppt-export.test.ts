import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { createRoadmapPresentation, layoutRoadmapCell, measureTextWidth, paginateRoadmapPage, presentationSlideCount } from "../src/main/pptxExport";
import { validateGeneratedPptx } from "../src/main/pptxValidation";
import {
  DEFAULT_PPT_EXPORT_SECTIONS,
  MAX_PPT_HALF_YEARS,
  buildPptExportPlan,
  roadmapBadgeKind,
  visibleRequirementTitles,
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

it("新增工作量仅影响封面和汇总总量，不改变三端图表或路标页样式", async () => {
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
  for (let page = 3; page <= 9; page++) expect(await output.file(`ppt/slides/slide${page}.xml`)!.async("string")).toBe(await baseline.file(`ppt/slides/slide${page}.xml`)!.async("string"));
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

describe("PPT导出计划", () => {
  it("医疗认证独立全量范围按十二个月和泳道高度续页，产品跟随需求", async () => {
    const data = denseSnapshot("NMPA认证", 6);
    data.requirements.forEach((item, index) => {
      item.targetMonth = "2030-05";
      item.medicalSubmissionMilestones = { registrationStartMonth: "2028-11", certificateMonth: "2030-06" };
      item.category = "平台演进";
      item.productIds = ["p1"];
      item.title = `独立认证需求${index}`;
    });
    const plan = buildPptExportPlan(data, { ...fullRequest, sections: ["医疗认证"] });
    expect(plan.requirements).toHaveLength(0);
    expect(plan.roadmapPages[0].groups).toHaveLength(6);
    expect(plan.roadmapPages[0].medicalMonths?.[0]).toBe("2028-11");
    expect(plan.roadmapPages[0].medicalMonths?.at(-1)).toBe("2030-06");
    const pages = paginateRoadmapPage(plan.roadmapPages[0]);
    expect(pages).toHaveLength(4);
    expect(pages.map((page) => page.medicalMonths?.length)).toEqual([12, 8, 12, 8]);
    expect(pages[0].groups).toEqual(pages[1].groups);
    const zip = await JSZip.loadAsync(await createRoadmapPresentation(data, { ...fullRequest, sections: ["医疗认证"] }));
    expect(Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))).toHaveLength(5);
    const xml = await zip.file("ppt/slides/slide2.xml")!.async("string");
    expect(xml).toContain("独立认证需求");
    expect(xml).toContain(data.products[0].name);
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

  it("全选基础为封面加八页，非密集内容实际仍为九页", () => {
    const plan = buildPptExportPlan(snapshot, fullRequest);
    expect(plan.slideCount).toBe(9);
    expect(presentationSlideCount(plan)).toBe(9);
    expect(plan.roadmapPages.map((item) => item.track)).toEqual(["运动", "健康", "行业", "全场景", "health平台", "AI", "医疗认证"]);
    expect(plan.roadmapPages.find((item) => item.track === "行业")?.groups).toEqual([]);
  });

  it("部分选择页数等于固定封面加选中项且保持固定顺序", () => {
    const plan = buildPptExportPlan(snapshot, { ...fullRequest, sections: ["医疗认证", "summary", "运动"] });
    expect(plan.selectedSections).toEqual(["summary", "运动", "医疗认证"]);
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

  it("字体固定8pt并在极限情况下压缩边距和需求数", () => {
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
    expect(layout.fontSize).toBe(8);
    expect(layout.cards.every((card) => card.top + card.height <= 170.95)).toBe(true);
    expect(measureTextWidth("连续英文ABC123", 10, true)).toBeGreaterThan(0);
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
    expect(presentationSlideCount(plan)).toBe(1 + pages.length);
    pages.forEach((part, index) => {
      expect(part.continuationIndex).toBe(index);
      expect(part.continuationCount).toBe(pages.length);
      expect(part.halfYears).toEqual(page.halfYears);
      expect(part.track).toBe(track);
      if (track === "医疗认证") { expect(part.medicalMonths).toEqual(["2026-03"]); return; }
      const height = (580.87 - 68) / roadmapLanesForTrack(track).length;
      const layout = layoutRoadmapCell(part.groups, 273.7675, height);
      expect(layout.fontSize).toBeGreaterThanOrEqual(8);
      layout.cards.forEach((card, cardIndex) => {
        expect(card.top).toBeGreaterThanOrEqual(layout.margin);
        expect(card.top + card.height).toBeLessThanOrEqual(height - layout.margin + 0.01);
        expect(card.fontSize).toBe(8);
        expect(card).not.toHaveProperty("badge");
        expect(card).not.toHaveProperty("badgeTop");
        card.bodyLines.forEach((line) => expect(measureTextWidth(line, card.fontSize)).toBeLessThanOrEqual(273.7675 - card.margin * 2));
        if (cardIndex) expect(card.top).toBeGreaterThanOrEqual(layout.cards[cardIndex - 1].top + layout.cards[cardIndex - 1].height + 4);
      });
    });
  });

  it.each(["2026H1", "2026H2", "2027H2"])("%s 为结束半年时按实际列宽分页", (end) => {
    const plan = buildPptExportPlan(denseSnapshot(), { ...fullRequest, end, sections: ["健康"] });
    const pages = paginateRoadmapPage(plan.roadmapPages[0]);
    expect(pages.length).toBeGreaterThan(1);
    expect(pages.flatMap((page) => page.groups)).toEqual(plan.groups);
    pages.forEach((page) => expect(() => layoutRoadmapCell(page.groups, (1223.07 - 128) / page.halfYears.length, (580.87 - 68) / 3)).not.toThrow());
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
    const cards = paginateRoadmapPage(page).flatMap((part) => layoutRoadmapCell(part.groups, 273.7675, 170.95).cards);
    for (const card of cards) {
      expect(card.height).toBeCloseTo(card.bodyOffset + card.bodyLines.length * card.bodyLineHeight);
      expect(card.bodyLines[0]).toMatch(/^· \S/);
      expect(card.bodyLines.every((line) => measureTextWidth(line, card.fontSize) <= 273.7675 - card.margin * 2)).toBe(true);
    }
  });
});

describe("PPTX输出", () => {
  it("密集完整材料页数、页码与校验一致，续页重复表头且每张卡片仅出现一次", async () => {
    const dense = denseSnapshot();
    const plan = buildPptExportPlan(dense, fullRequest);
    const expectedCount = presentationSlideCount(plan);
    const bytes = await createRoadmapPresentation(dense, fullRequest);
    expect(expectedCount).toBeGreaterThan(9);
    expect((await validateGeneratedPptx(bytes, expectedCount)).slides).toBe(expectedCount);
    const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
    const normalZip = await JSZip.loadAsync(await createRoadmapPresentation(snapshot, fullRequest));
    expect(await zip.file("ppt/theme/theme1.xml")!.async("string")).toEqual(await normalZip.file("ppt/theme/theme1.xml")!.async("string"));
    const xmls = await Promise.all(Array.from({ length: expectedCount }, (_, index) => zip.file(`ppt/slides/slide${index + 1}.xml`)!.async("string")));
    const roadmaps = xmls.slice(2);
    roadmaps.forEach((xml, index) => {
      expect(xml).toContain(`<a:t>${String(index + 3).padStart(2, "0")}</a:t>`);
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
    for (const domain of dense.domains.slice(1)) expect(xml.split(`<a:t>${domain.name}</a:t>`)).toHaveLength(2);
    expect(xml).toContain("健康路标（续 2/");
    expect(xml).toContain("产品专属");
    expect(xml.includes("↑")).toBe(false);
    expect(/roadmap-(?:product|experience)-/.test(xml)).toBe(false);
    const roadmapShapes = xml.match(/<p:sp>.*?<\/p:sp>/gs)!.filter((shape) => /name="roadmap-(?:title|body)-/.test(shape));
    expect(roadmapShapes.length).toBeGreaterThan(0);
    for (const shape of roadmapShapes) {
      expect([...shape.matchAll(/\bsz="(\d+)"/g)].every((match) => match[1] === "800")).toBe(true);
      expect(shape.includes('name="roadmap-title-') ? shape.includes('b="1"') : !shape.includes('b="1"')).toBe(true);
    }
    expect(xml).not.toMatch(/(?:NaN|Infinity|七类路标投入|领域详情|>\d{1,2}月<)/);
  }, 30_000);

  it("生成原生可编辑的完整9页文件且无详情页、月份标题或内部制作文字", async () => {
    const bytes = await createRoadmapPresentation(snapshot, fullRequest);
    const zip = await JSZip.loadAsync(bytes);
    const slideFiles = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
    expect(slideFiles).toHaveLength(9);
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
    expect(xml).not.toMatch(/name="roadmap-(?:product|experience)-/);
    expect(xml).toContain("当前区间暂无需求");
  }, 30_000);

  it("空状态路标仍输出两行半年表头和泳道", async () => {
    const bytes = await createRoadmapPresentation(snapshot, { ...fullRequest, sections: ["行业"] });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("ppt/slides/slide2.xml")!.async("string");
    expect(xml).toContain("行业路标");
    expect(xml).toContain("26H1");
    expect(xml).toContain("HMOS 6.1");
    expect(xml).not.toContain("产品标识");
    expect(xml).toContain("当前区间暂无需求");
  }, 30_000);
});
