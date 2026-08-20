import PptxGenJSImport from "pptxgenjs";
import { imageSize } from "image-size";
import { buildPptExportPlan, dynamicRoadmapMonthWidths, type PptDetailPagePlan, type PptExportPlan, type PptRoadmapPagePlan } from "../shared/pptExport.js";
import { roadmapCardLabel, type HalfYearSummary, type RoadmapGroup } from "../shared/roadmap.js";
import type { AppSnapshot, Requirement, RequirementCategory, RequirementImage, Track } from "../shared/types.js";
import { sumWorkloadBreakdown, type WorkloadSide } from "../shared/workload.js";

const FONT = "HarmonyOS Sans SC";
const C = {
  ink: "101820",
  navy: "17232E",
  navy2: "21313D",
  paper: "F3F6F7",
  white: "FFFFFF",
  muted: "70808B",
  faint: "DCE4E8",
  mint: "85E8B8",
  mintDeep: "30A978",
  blue: "287EF0",
  cyan: "12A99A",
  orange: "EE6B56",
  violet: "7357D9",
  lilac: "F0EEFF",
  paleMint: "E9FAF2",
  paleBlue: "EAF2FF",
  paleOrange: "FFF0EC",
} as const;

const SIDE_META: Record<WorkloadSide, { label: string; color: string }> = {
  device: { label: "设备侧", color: C.cyan },
  app: { label: "App 侧", color: C.blue },
  cloud: { label: "云侧", color: C.violet },
};

type PptxPresentation = any;
type PptxSlide = any;
const PptxGenJS = PptxGenJSImport as unknown as { new(): PptxPresentation };

export async function createRoadmapPresentation(
  snapshot: AppSnapshot,
  start: string,
  end: string,
): Promise<Buffer> {
  const plan = buildPptExportPlan(snapshot, start, end);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "需求路标工作台";
  pptx.company = "运动健康方案团队";
  pptx.subject = `${formatHalfYear(start)} 至 ${formatHalfYear(end)}需求路标`;
  pptx.title = "运动健康需求路标";
  pptx.lang = "zh-CN";
  pptx.theme = {
    headFontFace: FONT,
    bodyFontFace: FONT,
    lang: "zh-CN",
  };
  pptx.defineSlideMaster({
    title: "ROADMAP_LIGHT",
    background: { color: C.paper },
    objects: [],
    slideNumber: { x: 12.3, y: 0.26, w: 0.5, h: 0.2, color: C.ink, fontFace: FONT, fontSize: 9, bold: true, align: "right", margin: 0 },
  });
  pptx.defineSlideMaster({
    title: "ROADMAP_WHITE",
    background: { color: C.white },
    objects: [],
    slideNumber: { x: 12.3, y: 0.26, w: 0.5, h: 0.2, color: C.ink, fontFace: FONT, fontSize: 9, bold: true, align: "right", margin: 0 },
  });
  pptx.defineSlideMaster({
    title: "ROADMAP_DARK",
    background: { color: C.ink },
    objects: [],
    slideNumber: { x: 12.3, y: 0.26, w: 0.5, h: 0.2, color: C.white, fontFace: FONT, fontSize: 9, bold: true, align: "right", margin: 0 },
  });

  addCoverSlide(pptx, plan);
  addSummaryOverviewSlide(pptx, plan);
  addSummaryTrendSlide(pptx, plan);
  for (const page of plan.roadmapPages) addRoadmapSlide(pptx, plan, page);
  for (const page of plan.detailPages) addDetailSlide(pptx, page);

  const output = await pptx.write({ outputType: "nodebuffer", compression: true });
  return Buffer.from(output as Uint8Array);
}

function addCoverSlide(pptx: PptxPresentation, plan: PptExportPlan): void {
  const slide = pptx.addSlide({ masterName: "ROADMAP_DARK" });
  addTopBar(slide, "需求路标输出", true);
  addShape(slide, "rect", 0.67, 0.98, 0.94, 0.07, C.mint, C.mint);
  addText(slide, "运动健康\n需求路标", 0.67, 1.5, 6.8, 1.72, 38, C.white, true);
  addText(slide, `${formatHalfYear(plan.start)} — ${formatHalfYear(plan.end)}`, 0.69, 3.43, 5.8, 0.38, 17, "B8C6CE", false);
  addPill(slide, `${plan.requirements.length} 条需求`, 0.7, 4.2, 1.45, C.paleMint, C.mintDeep);
  addPill(slide, `${formatNumber(totalWorkload(plan.summaries))} 人月`, 2.27, 4.2, 1.52, C.paleBlue, C.blue);
  addPill(slide, `${plan.groups.length} 张路标卡片`, 3.92, 4.2, 1.85, C.paleOrange, C.orange);
  addText(slide, `本地离线生成 · ${formatDate(new Date())}`, 0.7, 6.62, 3.4, 0.25, 10, "71828D", false);

  addShape(slide, "roundRect", 8.46, 1.36, 3.9, 4.62, C.navy2, "334654", 1);
  addText(slide, `${plan.halfYears.length} 个半年`, 8.83, 1.77, 1.7, 0.27, 14, C.mint, true);
  addText(slide, "路标概览", 8.83, 2.14, 2.4, 0.42, 24, C.white, true);
  const tracks: Track[] = ["运动", "健康"];
  tracks.forEach((track, index) => {
    const scoped = plan.groups.filter((group) => group.track === track);
    const y = 3.0 + index * 1.2;
    addShape(slide, "ellipse", 8.86, y + 0.12, 0.09, 0.09, index ? C.orange : C.blue, index ? C.orange : C.blue);
    addText(slide, `${track}路标`, 9.12, y, 1.0, 0.28, 12, "B6C5CF", true);
    addShape(slide, "line", 10.05, y + 0.17, 1.72, 0, "40515E", "40515E", 1);
    addText(slide, `${scoped.length} 张卡片`, 10.1, y + 0.25, 1.6, 0.28, 12, C.white, true, "right");
  });
  addNotes(slide, "封面数据来自本地工作区当前选定时间范围。");
}

function addSummaryOverviewSlide(pptx: PptxPresentation, plan: PptExportPlan): void {
  const slide = pptx.addSlide({ masterName: "ROADMAP_LIGHT" });
  addTopBar(slide, "汇总分析");
  const totals = summaryTotals(plan.summaries);
  addTitle(slide, "SUMMARY", "三端投入与业务结构一页看清", `${formatHalfYear(plan.start)} 至 ${formatHalfYear(plan.end)} · ${plan.requirements.length} 条需求`);

  addMetric(slide, 0.67, "总工作量（三端合计）", `${formatNumber(totals.total)} 人月`, C.ink);
  addMetric(slide, 3.33, "设备侧", `${formatNumber(totals.device)} 人月`, C.cyan);
  addMetric(slide, 6.0, "App 侧", `${formatNumber(totals.app)} 人月`, C.blue);
  addMetric(slide, 8.67, "云侧", `${formatNumber(totals.cloud)} 人月`, C.violet);

  addDonut(slide, "三端投入占比", 0.68, 3.7, 3.65, 2.65, [
    { label: "设备侧", value: totals.device, color: C.cyan },
    { label: "App 侧", value: totals.app, color: C.blue },
    { label: "云侧", value: totals.cloud, color: C.violet },
  ]);
  addDonut(slide, "运动 / 健康占比", 4.84, 3.7, 3.65, 2.65, [
    { label: "运动", value: totals.sports, color: C.blue },
    { label: "健康", value: totals.health, color: C.orange },
  ]);
  addDonut(slide, "体验优化 / 产品专属", 9.0, 3.7, 3.65, 2.65, [
    { label: "体验优化", value: totals.experience, color: C.violet },
    { label: "产品专属", value: totals.exclusive, color: C.ink },
  ]);
  if (totals.unallocated > 0) addText(slide, `另有 ${formatNumber(totals.unallocated)} 人月历史工作量待拆分，未计入三端。`, 0.72, 6.77, 7.8, 0.25, 10, C.orange, true);
  addNotes(slide, "指标按设备侧、App 侧、云侧工作量累加；业务与分类占比使用同一工作量口径。");
}

function addSummaryTrendSlide(pptx: PptxPresentation, plan: PptExportPlan): void {
  const slide = pptx.addSlide({ masterName: "ROADMAP_LIGHT" });
  addTopBar(slide, "投入趋势");
  addTitle(slide, "WORKLOAD TREND", "三端投入节奏与体验优化占比", "柱状图显示各端工作量；右侧折线显示各端体验优化工作量占比。");
  const categories = plan.summaries.map((item) => formatHalfYear(item.halfYear));
  const safeCategories = categories.length ? categories : ["暂无数据"];
  const summaries = plan.summaries.length ? plan.summaries : [emptySummary("暂无数据")];

  addText(slide, "三端投入工作量", 0.72, 2.42, 5.0, 0.3, 15, C.ink, true);
  slide.addChart("bar", (["device", "app", "cloud"] as WorkloadSide[]).map((side) => ({
    name: SIDE_META[side].label,
    labels: safeCategories,
    values: summaries.map((item) => item.bySide[side]),
  })), {
    x: 0.7, y: 2.78, w: 6.2, h: 3.55,
    barDir: "col", grouping: "clustered", gapWidthPct: 55,
    chartColors: [C.cyan, C.blue, C.violet],
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 10,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 10,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 9,
    showValue: false, showTitle: false,
    showCatName: false,
    showPercent: false,
    showLeaderLines: false,
    showSerName: false,
    valGridLine: { color: C.faint, width: 1 },
    showBorder: false,
  });

  addText(slide, "体验优化占比（按端）", 7.28, 2.42, 4.8, 0.3, 15, C.ink, true);
  slide.addChart("line", (["device", "app", "cloud"] as WorkloadSide[]).map((side) => ({
    name: SIDE_META[side].label,
    labels: safeCategories,
    values: summaries.map((item) => item.bySide[side] ? Math.round((item.experienceBySide[side] / item.bySide[side]) * 100) : 0),
  })), {
    x: 7.25, y: 2.78, w: 5.35, h: 3.55,
    chartColors: [C.cyan, C.blue, C.violet],
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 10,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 10,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 9,
    valAxisMinVal: 0, valAxisMaxVal: 100, valAxisMajorUnit: 25,
    valAxisLabelFormatCode: "0\"%\"", showTitle: false, showValue: false,
    lineSize: 3, showMarker: true, markerSize: 6,
    valGridLine: { color: C.faint, width: 1 },
    showBorder: false,
  });
  addNotes(slide, "体验优化占比 = 该端体验优化工作量 / 该端总工作量；无投入时按 0% 处理。");
}

function addRoadmapSlide(pptx: PptxPresentation, plan: PptExportPlan, page: PptRoadmapPagePlan): void {
  const slide = pptx.addSlide({ masterName: "ROADMAP_LIGHT" });
  addTopBar(slide, `${page.track}路标`);
  const pageWorkload = page.groups.reduce((sum, group) => sum + group.totalWorkloadPm, 0);
  addTitle(slide, `${page.track === "运动" ? "SPORT" : "HEALTH"} ROADMAP`, `${formatHalfYear(page.halfYear)} ${page.track}需求路标`, `${page.groups.length} 张卡片 · ${formatNumber(pageWorkload)} 人月 · 点击卡片跳转详情`);
  addRoadmapGrid(slide, page, plan.firstDetailSlideByGroup);
  addNotes(slide, "同领域、同来源、同上线年月自动合并为一张卡片；产品专属卡片使用深色边框。");
}

function addDetailSlide(pptx: PptxPresentation, page: PptDetailPagePlan): void {
  const slide = pptx.addSlide({ masterName: "ROADMAP_WHITE" });
  addTopBar(slide, "领域详情");
  const partText = page.partCount > 1 ? ` · ${page.part}/${page.partCount}` : "";
  addTitle(slide, "DOMAIN DETAIL", `${page.group.cardTitle}：${page.group.requirements.length} 条需求${partText}`, `${page.group.source} · ${formatMonth(page.group.targetMonth)} · ${formatNumber(page.group.totalWorkloadPm)} 人月`);
  const cards = page.requirements.length ? page.requirements : page.group.requirements.slice(0, 1);
  cards.forEach((requirement, index) => addRequirementCard(slide, requirement, page.group, index, cards.length));
  addNotes(slide, page.group.cardSummary || "领域详情由工作台当前数据自动生成。");
}

function addRoadmapGrid(slide: PptxSlide, page: PptRoadmapPagePlan, detailLinks: Record<string, number>): void {
  const gridX = 1.66;
  const gridY = 2.67;
  const gridWidth = 10.04;
  const rowHeight = 1.03;
  const monthWidths = dynamicRoadmapMonthWidths(page.months, page.occupiedMonths, gridWidth);
  const monthX: number[] = [];
  monthWidths.reduce((x, width) => { monthX.push(x); return x + width; }, gridX);
  addText(slide, page.halfYear.slice(0, 4), 0.7, 2.33, 0.7, 0.28, 14, C.ink, true);
  page.months.forEach((month, index) => addText(slide, `${Number(month.slice(5))}月`, monthX[index], 2.33, monthWidths[index], 0.25, 11, C.muted, true, "center"));
  const levels = ["基础", "进阶", "高阶"] as const;
  levels.forEach((level, row) => {
    const accent = row === 0 ? C.orange : row === 1 ? C.blue : C.violet;
    addShape(slide, "ellipse", 0.76, gridY + row * rowHeight + 0.43, 0.1, 0.1, accent, accent);
    addText(slide, level, 0.96, gridY + row * rowHeight + 0.32, 0.55, 0.3, 13, C.ink, true);
    page.months.forEach((month, col) => {
      const x = monthX[col];
      const w = monthWidths[col];
      addShape(slide, "roundRect", x, gridY + row * rowHeight, Math.max(0.5, w - 0.06), rowHeight - 0.07, C.white, C.faint, 1);
      const groups = page.groups.filter((group) => group.level === level && group.targetMonth === month);
      if (groups.length) addRoadmapCellCards(slide, groups, x + 0.06, gridY + row * rowHeight + 0.07, Math.max(0.42, w - 0.18), rowHeight - 0.21, detailLinks);
    });
  });
  if (!page.groups.length) {
    addText(slide, `当前半年暂无${page.track}需求`, 4.35, 4.02, 4.65, 0.46, 19, C.muted, true, "center");
    addText(slide, "空状态保留月份与三级泳道，避免误解为漏数。", 4.15, 4.48, 5.05, 0.35, 12, C.muted, false, "center");
  }
}

function addRoadmapCellCards(
  slide: PptxSlide,
  groups: RoadmapGroup[],
  x: number,
  y: number,
  w: number,
  h: number,
  detailLinks: Record<string, number>,
): void {
  const shown = groups.slice(0, 3);
  const gap = 0.05;
  const cardHeight = Math.max(0.22, (h - gap * (shown.length - 1)) / shown.length);
  shown.forEach((group, index) => {
    const hasExclusive = group.categories.includes("产品专属");
    const cardY = y + index * (cardHeight + gap);
    addShape(slide, "roundRect", x, cardY, w, cardHeight, hasExclusive ? C.navy : C.paper, hasExclusive ? C.navy : C.faint, 1);
    addText(slide, roadmapCardLabel(group), x + 0.07, cardY + 0.03, Math.max(0.1, w - 0.14), Math.max(0.12, cardHeight - 0.06), w < 1.1 ? 8 : 10, hasExclusive ? C.white : C.ink, true, "left", {
      slide: detailLinks[group.key],
      tooltip: "查看领域详情",
    });
  });
  if (groups.length > shown.length) addText(slide, `+${groups.length - shown.length}`, x + w - 0.3, y + h - 0.2, 0.25, 0.16, 8, C.muted, true, "right");
}

function addRequirementCard(slide: PptxSlide, requirement: Requirement, group: RoadmapGroup, index: number, count: number): void {
  const cardY = count === 1 ? 2.48 : 2.4 + index * 2.12;
  const cardH = count === 1 ? 3.85 : 1.92;
  const image = requirement.images[0];
  const cardW = image ? 8.1 : 11.95;
  addShape(slide, "roundRect", 0.66, cardY, cardW, cardH, C.paper, C.paper, 1);
  const accent = group.level === "基础" ? C.orange : group.level === "进阶" ? C.blue : C.violet;
  addShape(slide, "ellipse", 0.92, cardY + 0.26, 0.13, 0.13, accent, accent);
  addText(slide, String(index + 1).padStart(2, "0"), 1.18, cardY + 0.17, 0.42, 0.24, 10, C.muted, true);
  addText(slide, requirement.title, 1.68, cardY + 0.12, image ? 4.25 : 6.6, 0.38, count === 1 ? 23 : 18, C.ink, true);
  addCategoryPill(slide, requirement.category, image ? 6.25 : 10.36, cardY + 0.16);
  const products = group.requirementProductNames[requirement.id] ?? [];
  if (products.length) addText(slide, products.join("、"), image ? 6.0 : 9.1, cardY + 0.62, image ? 1.65 : 2.7, 0.22, 9, C.muted, false, "right");
  addText(slide, "需求描述", 0.95, cardY + 0.72, 1.1, 0.24, 10, C.muted, true);
  const description = requirement.description.trim() || "描述待补充";
  addText(slide, truncate(description, count === 1 ? 290 : 135), 0.95, cardY + 1.03, cardW - 0.7, count === 1 ? 1.25 : 0.46, count === 1 ? 15 : 12, requirement.description.trim() ? C.ink : C.muted, false);
  const workload = sumWorkloadBreakdown([requirement]);
  const workloadLabel = `设备 ${formatNumber(workload.device)}  ·  App ${formatNumber(workload.app)}  ·  云 ${formatNumber(workload.cloud)}  ·  合计 ${formatNumber(workload.total)} 人月`;
  addText(slide, workloadLabel, 0.95, cardY + cardH - 0.52, cardW - 2.65, 0.25, 10, C.muted, true);
  addText(slide, formatMonth(requirement.targetMonth), cardW - 0.95, cardY + cardH - 0.52, 1.35, 0.25, 10, C.ink, true, "right");
  if (image) addRequirementImage(slide, image, 9.04, cardY, 3.62, cardH, requirement.images.length);
}

function addRequirementImage(slide: PptxSlide, image: RequirementImage, x: number, y: number, w: number, h: number, imageCount: number): void {
  addShape(slide, "roundRect", x, y, w, h, C.navy, C.navy, 1);
  const frame = containImage(image.dataUrl, x + 0.08, y + 0.08, w - 0.16, h - 0.42);
  slide.addImage({ data: image.dataUrl, ...frame });
  addText(slide, imageCount > 1 ? `首图 · 共 ${imageCount} 张` : image.name, x + 0.16, y + h - 0.3, w - 0.32, 0.18, 8, "B7C4CD", false, "center");
}

function addDonut(
  slide: PptxSlide,
  title: string,
  x: number,
  y: number,
  w: number,
  h: number,
  items: Array<{ label: string; value: number; color: string }>,
): void {
  addText(slide, title, x, y - 0.12, w, 0.28, 13, C.ink, true);
  const total = items.reduce((sum, item) => sum + item.value, 0);
  if (!total) {
    slide.addChart("doughnut", [{ name: title, labels: ["暂无数据"], values: [1] }], {
      x, y: y + 0.24, w, h: h - 0.24,
      chartColors: ["E3E9EC"], holeSize: 64,
      showLegend: false, showTitle: false, showValue: false, showPercent: false, showCatName: false,
      showBorder: false,
    });
    addText(slide, "暂无可计算数据", x + 0.78, y + 1.28, w - 1.56, 0.3, 11, C.muted, true, "center");
    addText(slide, items.map((item) => item.label).join(" / "), x + 0.3, y + h - 0.23, w - 0.6, 0.18, 8, C.muted, false, "center");
    return;
  }
  const normalized = items;
  slide.addChart("doughnut", [{ name: title, labels: normalized.map((item) => item.label), values: normalized.map((item) => item.value) }], {
    x, y: y + 0.24, w, h: h - 0.24,
    chartColors: normalized.map((item) => item.color), holeSize: 64,
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 9,
    showTitle: false, showValue: false, showPercent: true, showCatName: false,
    dataLabelColor: C.ink, dataLabelFormatCode: "0%",
    showBorder: false,
  });
}

function addTopBar(slide: PptxSlide, section: string, dark = false): void {
  const muted = dark ? "AFC0CB" : C.muted;
  addText(slide, "需求路标工作台", 0.67, 0.27, 2.5, 0.22, 9, muted, true);
  addText(slide, section, 10.25, 0.27, 1.65, 0.22, 9, muted, true, "right");
}

function addTitle(slide: PptxSlide, eyebrow: string, title: string, subtitle: string): void {
  addText(slide, eyebrow.toUpperCase(), 0.68, 0.86, 4.3, 0.24, 10, C.mintDeep, true);
  addText(slide, title, 0.68, 1.16, 11.8, 0.55, 26, C.ink, true);
  addText(slide, subtitle, 0.68, 1.87, 11.1, 0.3, 12, C.muted, false);
}

function addMetric(slide: PptxSlide, x: number, label: string, value: string, accent: string): void {
  addText(slide, label, x, 2.45, 2.25, 0.25, 10, C.muted, true);
  addText(slide, value, x, 2.77, 2.35, 0.5, 25, C.ink, true);
  addShape(slide, "rect", x, 3.34, 2.22, 0.06, accent, accent);
}

function addPill(slide: PptxSlide, text: string, x: number, y: number, w: number, fill: string, color: string): void {
  addText(slide, text, x, y, w, 0.33, 10, color, true, "center", undefined, fill, "roundRect");
}

function addCategoryPill(slide: PptxSlide, category: RequirementCategory, x: number, y: number): void {
  addPill(slide, category, x, y, 1.18, category === "产品专属" ? C.navy : C.lilac, category === "产品专属" ? C.white : C.violet);
}

function addText(
  slide: PptxSlide,
  text: string,
  x: number,
  y: number,
  w: number,
  h: number,
  fontSize: number,
  color: string,
  bold = false,
  align: "left" | "center" | "right" = "left",
  hyperlink?: { slide?: number; url?: string; tooltip?: string },
  fill?: string,
  shape: "rect" | "roundRect" = "rect",
): void {
  slide.addText(text, {
    x, y, w, h, fontFace: FONT, fontSize, color, bold, align, valign: "mid",
    margin: 0, breakLine: false, fit: "shrink", hyperlink,
    ...(fill ? { shape, fill: { color: fill }, line: { color: fill, transparency: 100 }, radius: 0.1 } : {}),
  });
}

function addShape(
  slide: PptxSlide,
  shape: "rect" | "roundRect" | "ellipse" | "line",
  x: number,
  y: number,
  w: number,
  h: number,
  fill: string,
  line: string,
  lineWidth = 0,
): void {
  slide.addShape(shape, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: line, width: lineWidth, transparency: lineWidth ? 0 : 100 },
    radius: shape === "roundRect" ? 0.1 : undefined,
  });
}

function addNotes(slide: PptxSlide, note: string): void {
  slide.addNotes(`[Sources]\n- 本地需求路标工作台数据快照（离线生成）\n\n${note}`);
}

function summaryTotals(summaries: HalfYearSummary[]) {
  const device = summaries.reduce((sum, item) => sum + item.deviceWorkload, 0);
  const app = summaries.reduce((sum, item) => sum + item.appWorkload, 0);
  const cloud = summaries.reduce((sum, item) => sum + item.cloudWorkload, 0);
  return {
    device,
    app,
    cloud,
    total: device + app + cloud,
    sports: summaries.reduce((sum, item) => sum + item.sportsWorkload, 0),
    health: summaries.reduce((sum, item) => sum + item.healthWorkload, 0),
    experience: summaries.reduce((sum, item) => sum + item.experienceWorkload, 0),
    exclusive: summaries.reduce((sum, item) => sum + item.exclusiveWorkload, 0),
    unallocated: summaries.reduce((sum, item) => sum + item.unallocatedWorkload, 0),
  };
}

function totalWorkload(summaries: HalfYearSummary[]): number {
  return summaries.reduce((sum, item) => sum + item.totalWorkload, 0);
}

function emptySummary(halfYear: string): HalfYearSummary {
  return {
    halfYear,
    bySource: { 运动基础: 0, 运动进阶: 0, 运动高阶: 0, 健康基础: 0, 健康进阶: 0, 健康高阶: 0 },
    sportsWorkload: 0, healthWorkload: 0, experienceWorkload: 0, exclusiveWorkload: 0,
    totalWorkload: 0, deviceWorkload: 0, appWorkload: 0, cloudWorkload: 0, unallocatedWorkload: 0,
    bySide: { device: 0, app: 0, cloud: 0 },
    experienceBySide: { device: 0, app: 0, cloud: 0 },
    requirementCount: 0,
  };
}

function formatHalfYear(value: string): string {
  const match = /^(\d{4})H([12])$/.exec(value);
  return match ? `${match[1]} H${match[2]}` : value;
}

function formatMonth(value: string): string {
  const [year, month] = value.split("-");
  return `${year}年${Number(month)}月`;
}

function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(2).replace(/0+$/, "").replace(/\.$/, "");
}

function formatDate(date: Date): string {
  return `${date.getFullYear()}.${String(date.getMonth() + 1).padStart(2, "0")}.${String(date.getDate()).padStart(2, "0")}`;
}

function truncate(value: string, max: number): string {
  return value.length <= max ? value : `${value.slice(0, Math.max(0, max - 1))}…`;
}

function containImage(dataUrl: string, x: number, y: number, w: number, h: number): { x: number; y: number; w: number; h: number } {
  try {
    const base64 = dataUrl.slice(dataUrl.indexOf(",") + 1);
    const dimensions = imageSize(Buffer.from(base64, "base64"));
    if (!dimensions.width || !dimensions.height) return { x, y, w, h };
    const scale = Math.min(w / dimensions.width, h / dimensions.height);
    const fittedW = dimensions.width * scale;
    const fittedH = dimensions.height * scale;
    return { x: x + (w - fittedW) / 2, y: y + (h - fittedH) / 2, w: fittedW, h: fittedH };
  } catch {
    return { x, y, w, h };
  }
}
