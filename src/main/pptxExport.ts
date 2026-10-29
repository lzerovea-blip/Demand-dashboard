import PptxGenJSImport from "pptxgenjs";
import JSZip from "jszip";
import {
  buildPptExportPlan,
  groupsForHalfYear,
  visibleRequirementTitles,
  type PptExportPlan,
  type PptRoadmapPagePlan,
} from "../shared/pptExport.js";
import { groupMatchesRoadmapLane, halfYearOf, medicalRoadmapDemandRows, hmosVersionForHalfYear, roadmapLanesForTrack, type HalfYearSummary, type RoadmapGroup } from "../shared/roadmap.js";
import { medicalSubmissionTimeline } from "../shared/medicalSubmission.js";
import { CATEGORIES } from "../shared/types.js";
import type { AppSnapshot, PptExportRequest, RequirementCategory, Track } from "../shared/types.js";
import type { WorkloadSide } from "../shared/workload.js";
import { pptPermissionExpiration } from "../shared/pptAccess.js";

const FONT = "HarmonyOS Sans SC";
const PX = 1 / 96;
const C = {
  ink: "101820",
  navy: "17232E",
  navy2: "21313D",
  paper: "F3F6F7",
  white: "FFFFFF",
  muted: "70808B",
  muted2: "98A4AC",
  faint: "DCE4E8",
  faint2: "E9EEF1",
  mint: "85E8B8",
  mintDeep: "30A978",
  blue: "287EF0",
  coral: "EE6B56",
  coralPale: "FFF0EC",
  violet: "7357D9",
  amber: "C96A08",
  red: "C00000",
} as const;

const CATEGORY_COLORS: Record<RequirementCategory, string> = {
  体验优化: C.violet,
  平台演进: C.amber,
  产品专属: C.ink,
};

const SIDE_META: Record<WorkloadSide, { label: string; color: string }> = {
  device: { label: "设备侧", color: C.mintDeep },
  app: { label: "运动健康 App", color: C.blue },
  cloud: { label: "云侧", color: C.coral },
};

const ROADMAP_TABLE = {
  x: 27.6,
  y: 62,
  w: 1223.07,
  h: 618,
  laneW: 128,
  headerH: 32,
} as const;
const MEDICAL_HEADER_HEIGHT = 56;

const ROADMAP_STYLE = {
  header: "8B1E2D",
  lane: "F8FAFB",
  border: "DFE5E9",
  stripe: "9AA4AB",
  divider: "CFD7DC",
  ink: "17232C",
  blue: "3989DC",
  submission: "DB5A4C",
  connector: "CFDBE4",
} as const;

type PptxPresentation = any;
type PptxSlide = any;
const PptxGenJS = PptxGenJSImport as unknown as { new(): PptxPresentation };

export interface RoadmapCardLayout {
  group: RoadmapGroup;
  fontSize: number;
  margin: number;
  gapAfter: number;
  visibleRequirements: string[];
  titleLines: string[];
  bodyLines: string[];
  titleLineHeight: number;
  bodyLineHeight: number;
  bodyOffset: number;
  width: number;
  infoX: number;
  infoWidth: number;
  titleTop: number;
  titleFontSize: number;
  dividerX: number;
  bodyX: number;
  bodyWidth: number;
  height: number;
  top: number;
}

export interface RoadmapCellLayout {
  fontSize: number;
  margin: number;
  cardGap: number;
  visibleLimit: number;
  cards: RoadmapCardLayout[];
}

export interface PaginatedRoadmapPage extends PptRoadmapPagePlan {
  continuationIndex: number;
  continuationCount: number;
}

export interface RoadmapLaneLayout {
  id: string;
  name: string;
  top: number;
  height: number;
}

export interface RoadmapTableLayout {
  lanes: RoadmapLaneLayout[];
  height: number;
  bodyCapacity: number;
  halfWidth: number;
}

class RoadmapCellCapacityError extends Error {}
export interface PptExportSecurityContext {
  exportedAt: Date;
  expiration: Date;
}

function defaultSecurityContext(): PptExportSecurityContext {
  const exportedAt = new Date();
  return { exportedAt, expiration: pptPermissionExpiration(exportedAt) };
}

export async function createRoadmapPresentation(
  snapshot: AppSnapshot,
  request: PptExportRequest,
  _security: PptExportSecurityContext = defaultSecurityContext(),
): Promise<Buffer> {
  const plan = buildPptExportPlan(snapshot, request);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "林忠维";
  pptx.company = "运动健康方案团队";
  pptx.subject = `${formatHalfYear(plan.start)} 至 ${formatHalfYear(plan.end)}解决方案需求路标`;
  pptx.title = plan.materialName;
  pptx.lang = "zh-CN";
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT, lang: "zh-CN" };

  addCoverSlide(pptx, plan);
  for (const section of plan.selectedSections) {
    if (section === "summary") addSummarySlide(pptx, plan, pptx._slides.length + 1);
    else {
      for (const page of plan.roadmapPages.filter((item) => item.track === section)) {
        paginateRoadmapPage(page).forEach((part) => addRoadmapSlide(pptx, part));
      }
    }
  }

  const output = await pptx.write({ outputType: "nodebuffer", compression: true });
  // PptxGenJS 4 declares a master per slide but writes only the shared master.
  // Remove only those unreferenced declarations; keep masters and themes intact.
  const archive = await JSZip.loadAsync(output as Uint8Array);
  const contentTypes = await archive.file("[Content_Types].xml")!.async("string");
  const normalizedTypes = contentTypes.replace(/<Override\b[^>]*\bPartName="(\/ppt\/slideMasters\/slideMaster\d+\.xml)"[^>]*\/>/g,
    (declaration, part: string) => archive.file(part.slice(1)) ? declaration : "");
  archive.file("[Content_Types].xml", normalizedTypes);
  for (const entry of Object.values(archive.files).filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.name))) {
    const original = await entry.async("string");
    const grouped = groupRoadmapCardShapes(original);
    if (grouped !== original) archive.file(entry.name, grouped);
  }
  return archive.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

/** Group only the five explicitly named objects of a card; child geometry stays editable. */
export function groupRoadmapCardShapes(xml: string): string {
  if (/<p:cNvPr\b[^>]*name="roadmap-card-/.test(xml)) return xml;
  const shapes = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)];
  let nextId = Math.max(1, ...[...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((match) => Number(match[1]))) + 1;
  const replacements: Array<{ start: number; end: number; value: string }> = [];
  for (let index = 0; index < shapes.length; index++) {
    const frame = shapes[index];
    const name = frame[0].match(/name="roadmap-frame-([^"]+)"/)?.[1];
    if (!name) continue;
    const block = shapes.slice(index, index + 5);
    if (block.length !== 5 || !block[4][0].includes(`name="roadmap-body-${name}"`)) throw new Error("PPT 需求块组合对象不完整");
    const bounds = frame[0].match(/<a:xfrm[^>]*>\s*<a:off x="(\d+)" y="(\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/);
    if (!bounds) throw new Error("PPT 需求块组合坐标缺失");
    const [, x, y, w, h] = bounds;
    const start = frame.index!;
    const end = block[4].index! + block[4][0].length;
    const children = xml.slice(start, end);
    const value = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${nextId++}" name="roadmap-card-${name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/><a:chOff x="${x}" y="${y}"/><a:chExt cx="${w}" cy="${h}"/></a:xfrm></p:grpSpPr>${children}</p:grpSp>`;
    replacements.push({ start, end, value });
    index += 4;
  }
  for (const replacement of replacements.reverse()) xml = xml.slice(0, replacement.start) + replacement.value + xml.slice(replacement.end);
  return xml;
}

/** Actual count, including overflow pages; shared by validation and the export result. */
export function presentationSlideCount(plan: PptExportPlan): number {
  return 1 + (plan.selectedSections.includes("summary") ? 1 : 0)
    + plan.roadmapPages.reduce((count, page) => count + paginateRoadmapPage(page).length, 0);
}

function addCoverSlide(pptx: PptxPresentation, plan: PptExportPlan): void {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  text(slide, plan.materialName, 90, 260, 1100, 200, 44, ROADMAP_STYLE.ink, true, "center", "mid", { name: "cover-title" });
}

function addSummarySlide(pptx: PptxPresentation, plan: PptExportPlan, slideNumber: number): void {
  const slide = pptx.addSlide();
  slide.background = { color: C.paper };
  addSlideHeader(slide, slideNumber, "WORKLOAD SUMMARY", "需求工作量汇总分析", `${formatHalfYear(plan.start)}—${formatHalfYear(plan.end)} · ${plan.requirements.length} 条需求`);
  const total = totals(plan.summaries);
  const metrics: Array<[string, number, string]> = [
    ["总工作量", total.total + total.unallocated, C.ink],
    ["设备侧", total.device, C.mintDeep],
    ["运动健康 App", total.app, C.blue],
    ["云侧", total.cloud, C.coral],
  ];
  metrics.forEach(([label, value, color], index) => metricCard(slide, 64 + index * 288, 194, 266, 101, label, value, color));
  text(slide, `创新研究 App ${formatNumber(total.researchApp)} PM、算法 ${formatNumber(total.algorithm)} PM 已计入总工作量，未计入下方三端图表。`, 64, 299, 1152, 18, 9, C.muted);

  roundRect(slide, 64, 320, 700, 326, C.white, C.faint, 1, 12);
  text(slide, "三端工作量趋势", 88, 342, 230, 24, 16, C.ink, true);
  text(slide, "单位：PM", 620, 346, 116, 18, 10, C.muted, false, "right");
  drawSummaryTrend(slide, plan.summaries);

  roundRect(slide, 784, 320, 434, 326, C.white, C.faint, 1, 12);
  text(slide, "三端需求结构", 808, 342, 230, 24, 16, C.ink, true);
  text(slide, "体验优化／平台演进／产品专属", 808, 370, 360, 18, 10, C.muted);
  drawSideCategoryStructure(slide, plan.summaries);
  if (total.unallocated > 0) {
    roundRect(slide, 808, 613, 386, 22, C.coralPale, C.coralPale, 0, 6);
    text(slide, `提示：另有 ${formatNumber(total.unallocated)} PM 历史工作量未拆分至三端`, 820, 616, 362, 16, 9, C.coral, true);
  }
  notes(slide, "总工作量包含五项投入及历史待拆分。图表仅统计设备、运动健康 App、云侧的实际工作量，不包含创新研究 App 和算法。本页不展示七类路标投入。");
}

function addRoadmapSlide(pptx: PptxPresentation, page: PaginatedRoadmapPage): void {
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  const continuation = page.continuationCount > 1
    ? `（${page.continuationIndex ? "续 " : ""}${page.continuationIndex + 1}/${page.continuationCount}）`
    : "";
  const region = page.region ? ` · ${page.region === "domestic" ? "国内" : "海外"}` : "";
  text(slide, `${page.name ?? `${page.track}路标`}${region}${continuation}`, 50, 20, 1100, 38, 24, C.ink, true, "left", "mid");
  if (page.kind === "medical" || page.track === "医疗认证") drawMedicalCertificationTable(slide, page);
  else drawRoadmapTable(slide, page);
  notes(slide, page.kind === "medical" || page.track === "医疗认证" ? "当前区域医疗认证需求按独立泳道展示，共享五节点月份轴；超过十二个月或纵向容量时自动续页。" : `${page.name ?? page.track}按半年及区域展示；左侧为领域，右侧每条圆点文本对应一条独立需求记录。`);
}

export function paginateRoadmapPage(page: PptRoadmapPagePlan): PaginatedRoadmapPage[] {
  if (page.kind === "medical" || page.track === "医疗认证") return paginateMedicalCertificationPage(page);
  if (page.lanes && page.lanes.length > 3) {
    const parts: PaginatedRoadmapPage[] = [];
    for (let index = 0; index < page.lanes.length; index += 3) {
      const lanes = page.lanes.slice(index, index + 3);
      parts.push(...paginateRoadmapPage({ ...page, lanes, groups: page.groups.filter((group) => lanes.some((lane) => groupMatchesRoadmapLane(group, lane.id))) }));
    }
    return parts.map((part, continuationIndex) => ({ ...part, continuationIndex, continuationCount: parts.length }));
  }
  const lanes = roadmapLaneIds(page);
  const queues = lanes.flatMap((lane) => page.halfYears.map((halfYear) =>
    groupsForHalfYear(page.groups, halfYear).filter((group) => groupMatchesRoadmapLane(group, lane))));
  const parts: PptRoadmapPagePlan[] = [];
  // Round-robin preserves each cell's order without letting the first lane consume
  // every page. The same adaptive height calculation is used here and for drawing.
  do {
    const selected: RoadmapGroup[] = [];
    let progress: boolean;
    do {
      progress = false;
      for (const queue of queues) {
        if (!queue.length) continue;
        const candidate = [...selected, queue[0]];
        try {
          layoutRoadmapLanes({ ...page, groups: candidate });
          selected.push(queue.shift()!);
          progress = true;
        } catch (error) {
          if (!(error instanceof RoadmapCellCapacityError)) throw error;
        }
      }
    } while (progress);
    if (!selected.length && queues.some((queue) => queue.length)) {
      const group = queues.find((queue) => queue.length)![0];
      const lane = lanes.find((lane) => groupMatchesRoadmapLane(group, lane));
      throw new RoadmapCellCapacityError(`PPT 路标单张卡片超过可排版容量：${page.track}路标 / ${lane} / ${halfYearOf(group.targetMonth)} / ${group.cardTitle}。请缩短该领域标题或需求标题后重试`);
    }
    const keys = new Set(selected.map((group) => group.key));
    const part = { ...page, groups: page.groups.filter((group) => keys.has(group.key)) };
    layoutRoadmapLanes(part);
    parts.push(part);
  } while (queues.some((queue) => queue.length));
  return parts.map((part, continuationIndex) => ({ ...part, continuationIndex, continuationCount: parts.length }));
}

function roadmapLaneIds(page: PptRoadmapPagePlan): string[] {
  return page.lanes ? (page.lanes.length ? page.lanes.map((item) => item.id) : ["__empty_lane"]) : [...roadmapLanesForTrack(page.track)];
}

function roadmapLaneName(page: PptRoadmapPagePlan, lane: string): string {
  return page.lanes?.find((item) => item.id === lane)?.name ?? (lane === "__empty_lane" ? "请添加泳道" : lane);
}

function requiredCellHeight(groups: RoadmapGroup[], width: number, visibleLimit: number): number {
  const cards = calculateCardLayouts(groups, width, 7, 4, 3, visibleLimit);
  return cards.reduce((sum, card) => sum + card.height, 0) + Math.max(0, cards.length - 1) * 3 + 8;
}

/** Keep the table full height; dense lanes borrow from the balanced row baseline. */
export function layoutRoadmapLanes(page: PptRoadmapPagePlan): RoadmapTableLayout {
  const bodyCapacity = ROADMAP_TABLE.h - ROADMAP_TABLE.headerH;
  const halfWidth = (ROADMAP_TABLE.w - ROADMAP_TABLE.laneW) / page.halfYears.length;
  const needs = roadmapLaneIds(page).map((id) => {
    const name = roadmapLaneName(page, id);
    const minimumLabel = Math.max(48, wrapText(name, ROADMAP_TABLE.laneW - 24, 12, true).length * pointsToPixels(12) * 1.18 + 16);
    const cells = page.halfYears.map((half) => groupsForHalfYear(page.groups, half).filter((group) => groupMatchesRoadmapLane(group, id)));
    return { id, name,
      minimum: Math.max(minimumLabel, ...cells.map((groups) => requiredCellHeight(groups, halfWidth, 1))),
      preferred: Math.max(minimumLabel, ...cells.map((groups) => requiredCellHeight(groups, halfWidth, 3))),
    };
  });
  const minimum = needs.reduce((sum, lane) => sum + lane.minimum, 0);
  if (minimum > bodyCapacity + 0.01) throw new RoadmapCellCapacityError("PPT 泳道内容需要拆分续页");
  const preferred = needs.reduce((sum, lane) => sum + lane.preferred, 0);
  let heights: number[];
  if (preferred <= bodyCapacity) {
    // Start from an equal-height baseline. Reserve complete content height for
    // any lane exceeding its share, then rebalance the remaining lanes.
    heights = needs.map(() => 0);
    let remaining = bodyCapacity;
    let pending = needs.map((_, index) => index);
    while (pending.length) {
      const share = remaining / pending.length;
      const dense = pending.filter((index) => needs[index].preferred > share);
      if (!dense.length) {
        pending.forEach((index) => { heights[index] = share; });
        break;
      }
      for (const index of dense) {
        heights[index] = needs[index].preferred;
        remaining -= heights[index];
      }
      pending = pending.filter((index) => !dense.includes(index));
    }
  } else {
    // All lanes still keep their measured one-item minimum. Use the remaining
    // capacity for longer summaries without shrinking their font sizes.
    const desiredExtra = preferred - minimum;
    const ratio = desiredExtra ? Math.max(0, bodyCapacity - minimum) / desiredExtra : 0;
    heights = needs.map((lane) => lane.minimum + (lane.preferred - lane.minimum) * ratio);
  }
  // Absorb floating-point remainder in the final row, keeping the bottom edge
  // identical on empty, sparse, dense and continuation pages.
  heights[heights.length - 1] = bodyCapacity - heights.slice(0, -1).reduce((sum, height) => sum + height, 0);
  let top = ROADMAP_TABLE.headerH;
  const lanes = needs.map((lane, index) => {
    const height = heights[index];
    const placed = { id: lane.id, name: lane.name, top, height };
    top += height;
    return placed;
  });
  return { lanes, height: ROADMAP_TABLE.h, bodyCapacity, halfWidth };
}

function splitRoadmapCellGroups(
  groups: RoadmapGroup[],
  width: number,
  height: number,
  track: Track,
  lane: string,
  halfYear: string,
): RoadmapGroup[][] {
  const chunks: RoadmapGroup[][] = [];
  let current: RoadmapGroup[] = [];
  for (const group of groups) {
    const candidate = [...current, group];
    try {
      layoutRoadmapCell(candidate, width, height);
      current = candidate;
    } catch {
      if (current.length) chunks.push(current);
      try {
        layoutRoadmapCell([group], width, height);
      } catch {
        throw new Error(`PPT 路标单张卡片超过可排版容量：${track}路标 / ${lane} / ${halfYear} / ${group.cardTitle}。请缩短该领域标题或需求标题`);
      }
      current = [group];
    }
  }
  if (current.length) chunks.push(current);
  return chunks;
}

function drawRoadmapTable(slide: PptxSlide, page: PptRoadmapPagePlan): void {
  const { x, y, w, laneW, headerH } = ROADMAP_TABLE;
  const layout = layoutRoadmapLanes(page);
  const { lanes, height: h, halfWidth: halfW } = layout;
  // Labels are editable text boxes below; the native table only supplies the grid.
  // Drawing labels in both places produces overlapping text on every continuation.
  const rows = Array.from({ length: lanes.length + 1 }, () => Array(page.halfYears.length + 1).fill(""));
  const colW = [laneW * PX, ...page.halfYears.map(() => halfW * PX)];
  const rowH = [headerH * PX, ...lanes.map((lane) => lane.height * PX)];
  slide.addTable(rows, {
    x: x * PX,
    y: y * PX,
    w: w * PX,
    h: h * PX,
    colW,
    rowH,
    border: { type: "solid", color: ROADMAP_STYLE.border, pt: 0.75 },
    fill: C.white,
    color: C.ink,
    fontFace: FONT,
    fontSize: 1,
    margin: 0,
    autoFit: false,
    valign: "middle",
    align: "center",
  });
  rect(slide, x, y, w, headerH, ROADMAP_STYLE.header, ROADMAP_STYLE.border, 0.75);
  text(slide, "路标泳道", x, y, laneW, headerH, 12, C.white, true, "center", "mid", { wrap: false });
  page.halfYears.forEach((halfYear, index) => {
    const left = x + laneW + index * halfW;
    line(slide, left, y, 0, headerH, ROADMAP_STYLE.border, 0.75);
    text(slide, `${shortHalf(halfYear)}（${hmosVersionForHalfYear(halfYear)}）`, left + 4, y + 4, halfW - 8, headerH - 8, 12, C.white, true, "center", "mid", { wrap: false, name: `roadmap-half-${halfYear}` });
  });
  lanes.forEach((lane) => {
    const rowY = y + lane.top;
    const laneH = lane.height;
    rect(slide, x, rowY, laneW, laneH, ROADMAP_STYLE.lane, ROADMAP_STYLE.border, 0.75);
    const labelLines = wrapText(lane.name, laneW - 24, 12, true);
    text(slide, labelLines.join("\n"), x + 12, rowY + 8, laneW - 24, laneH - 16, 12, ROADMAP_STYLE.ink, true, "center", "mid", { wrap: false, name: `roadmap-lane-${lane.id}` });
    page.halfYears.forEach((halfYear, halfIndex) => {
      const cellX = x + laneW + halfIndex * halfW;
      const groups = groupsForHalfYear(page.groups, halfYear).filter((group) => groupMatchesRoadmapLane(group, lane.id));
      if (groups.length) drawRoadmapCell(slide, groups, cellX, rowY, halfW, laneH, page.track, lane.id, halfYear);
      else if (page.groups.length === 0) text(slide, "当前区间暂无需求", cellX + 8, rowY + 8, halfW - 16, laneH - 16, 10.5, C.muted2, true, "center", "mid");
    });
  });
}

function drawRoadmapCell(
  slide: PptxSlide,
  groups: RoadmapGroup[],
  cellX: number,
  cellY: number,
  cellW: number,
  cellH: number,
  track: Track,
  lane: string,
  halfYear: string,
): void {
  const layout = layoutRoadmapCell(groups, cellW, cellH);
  for (const [index, card] of layout.cards.entries()) {
    const left = cellX + card.margin;
    const top = cellY + card.top;
    const titleHeight = card.titleLines.length * card.titleLineHeight;
    const name = `${track}-${lane}-${halfYear}-${index}`;
    rect(slide, left, top, card.width, card.height, C.white, ROADMAP_STYLE.border, 0.75, 7, `roadmap-frame-${name}`);
    line(slide, left + 1.5, top + 5, 0, card.height - 10, ROADMAP_STYLE.stripe, 2.25);
    line(slide, left + card.dividerX, top + 4, 0, card.height - 8, ROADMAP_STYLE.divider, 0.75);
    // Use exactly the measured lines. A second Office wrap can add an unmeasured
    // line (or strand the bullet), invalidating the pagination capacity check.
    text(slide, card.titleLines.join("\n"), left + card.infoX, top + card.titleTop, card.infoWidth, titleHeight, card.titleFontSize, ROADMAP_STYLE.ink, true, "left", "top", { margin: 0, wrap: false, lineSpacing: card.titleLineHeight * 72 / 96, name: `roadmap-title-${name}` });
    const bodyTop = top + card.bodyOffset;
    text(slide, card.bodyLines.join("\n"), left + card.bodyX, bodyTop, card.bodyWidth, card.bodyLines.length * card.bodyLineHeight, card.fontSize, "000000", false, "left", "top", { margin: 0, wrap: false, lineSpacing: card.bodyLineHeight * 72 / 96, name: `roadmap-body-${name}` });
  }
}

export function layoutRoadmapCell(groups: RoadmapGroup[], width: number, height: number): RoadmapCellLayout {
  const attempts: Array<{ fontSize: number; margin: number; gap: number; visibleLimit: number }> = [
    { fontSize: 7, margin: 4, gap: 3, visibleLimit: 3 },
    { fontSize: 7, margin: 4, gap: 3, visibleLimit: 2 },
    { fontSize: 7, margin: 4, gap: 3, visibleLimit: 1 },
  ];
  for (const attempt of attempts) {
    const cards = calculateCardLayouts(groups, width, attempt.fontSize, attempt.margin, attempt.gap, attempt.visibleLimit);
    const used = cards.reduce((sum, card) => sum + card.height, 0) + Math.max(0, cards.length - 1) * attempt.gap + attempt.margin * 2;
    if (used <= height + 0.01) {
      let top = attempt.margin;
      return {
        fontSize: attempt.fontSize,
        margin: attempt.margin,
        cardGap: attempt.gap,
        visibleLimit: attempt.visibleLimit,
        cards: cards.map((card) => {
          const placed = { ...card, top };
          top += card.height + attempt.gap;
          return placed;
        }),
      };
    }
  }
  throw new RoadmapCellCapacityError("PPT 路标单元格内容超过可排版容量：需要拆分续页");
}

function calculateCardLayouts(groups: RoadmapGroup[], width: number, fontSize: number, margin: number, gap: number, visibleLimit: number): RoadmapCardLayout[] {
  const cardWidth = Math.min(260, width - margin * 2);
  const infoX = 9;
  const infoWidth = 64;
  const dividerX = infoX + infoWidth + 4;
  const bodyX = dividerX + 1 + 4;
  const bodyWidth = cardWidth - bodyX - 6;
  const titleFontSize = 8;
  const bodyLineHeight = pointsToPixels(fontSize) * 1.18;
  return groups.map((group, index) => {
    const titleLineHeight = pointsToPixels(titleFontSize) * 1.18;
    const titleLines = wrapText(group.cardTitle, infoWidth - 2, titleFontSize, true);
    const visibleRequirements = visibleRequirementTitlesForLayout(group, visibleLimit);
    // Reserve indentation and a small font-substitution tolerance before wrapping.
    const bodyLines = visibleRequirements.flatMap((item) => wrapText(item, bodyWidth - measureTextWidth("· ", fontSize) - 4, fontSize, false)
      .map((part, partIndex) => `${partIndex ? "  " : "· "}${part}`));
    const titleHeight = titleLines.length * titleLineHeight;
    const infoHeight = titleHeight;
    const bodyHeight = Math.max(1, bodyLines.length) * bodyLineHeight;
    const height = Math.max(32, 8 + Math.max(infoHeight, bodyHeight));
    const titleTop = 4 + (height - 8 - infoHeight) / 2;
    const bodyOffset = 4 + (height - 8 - bodyHeight) / 2;
    return {
      group,
      fontSize,
      margin,
      gapAfter: index === groups.length - 1 ? 0 : gap,
      visibleRequirements,
      titleLines,
      bodyLines,
      titleLineHeight,
      bodyLineHeight,
      bodyOffset,
      width: cardWidth,
      infoX,
      infoWidth,
      titleTop,
      titleFontSize,
      dividerX,
      bodyX,
      bodyWidth,
      height,
      top: 0,
    };
  });
}

function visibleRequirementTitlesForLayout(group: RoadmapGroup, limit: number): string[] {
  const visible = visibleRequirementTitles(group.requirements, Math.min(3, limit));
  if (group.requirements.length > limit && visible.length && !visible.at(-1)?.endsWith("…")) {
    visible[visible.length - 1] = `${visible[visible.length - 1]}…`;
  }
  return visible;
}

function medicalRowLayout(group: RoadmapGroup) {
  const item = group.requirements[0];
  const titleLines = wrapText(item.title, 210, 7, true);
  const titleHeight = titleLines.length * 12;
  const laneName = group.laneName ?? (item.source === "NMPA认证" ? "NMPA认证" : "CE认证");
  const badgeLines = wrapText(laneName, 204, 7, true);
  const badgeHeight = Math.max(16, badgeLines.length * 12 + 4);
  const badgeWidth = Math.min(222, Math.max(...badgeLines.map((part) => measureTextWidth(part, 7, true))) + 14);
  const nodes = medicalSubmissionTimeline(item.targetMonth, item.medicalSubmissionMilestones).filter((node) => node.month);
  const nodeLines = Math.max(1, ...nodes.map((node) => nodes.filter((other) => other.month === node.month).length));
  return { titleLines, titleHeight, badgeLines, badgeHeight, badgeWidth,
    height: Math.max(64, 12 + titleHeight + 8 + badgeHeight + 8, 24 + 12 + nodeLines * 12 + 8) };
}

function medicalRowHeight(group: RoadmapGroup): number {
  return medicalRowLayout(group).height;
}

function paginateMedicalCertificationPage(page: PptRoadmapPagePlan): PaginatedRoadmapPage[] {
  const groups = medicalRoadmapDemandRows(page.groups).map(({ group }) => group);
  const batches: RoadmapGroup[][] = [];
  let batch: RoadmapGroup[] = [];
  let used = 0;
  for (const group of groups) {
    const height = medicalRowHeight(group);
    if (height > ROADMAP_TABLE.h - MEDICAL_HEADER_HEIGHT) throw new RoadmapCellCapacityError(`医疗认证需求标题过长，单泳道无法容纳：${group.requirements[0].title}`);
    if (batch.length && used + height > ROADMAP_TABLE.h - MEDICAL_HEADER_HEIGHT) { batches.push(batch); batch = []; used = 0; }
    batch.push(group); used += height;
  }
  if (batch.length || !batches.length) batches.push(batch);
  const months = page.medicalMonths ?? [];
  const segments: string[][] = [];
  for (let index = 0; index < months.length; index += 12) segments.push(months.slice(index, index + 12));
  if (!segments.length) segments.push([]);
  const pages = batches.flatMap((groups) => segments.map((medicalMonths) => ({ ...page, groups, medicalMonths })));
  return pages.map((part, continuationIndex) => ({ ...part, continuationIndex, continuationCount: pages.length }));
}

function drawMedicalCertificationTable(slide: PptxSlide, page: PaginatedRoadmapPage): void {
  const { x, y, w, h } = ROADMAP_TABLE;
  const headerH = MEDICAL_HEADER_HEIGHT;
  const laneW = 240;
  const months = page.medicalMonths ?? [];
  const columnW = (w - laneW) / Math.max(1, months.length);
  rect(slide, x, y, w, h, C.white, ROADMAP_STYLE.border, 0.75);
  rect(slide, x, y, w, headerH, ROADMAP_STYLE.header, ROADMAP_STYLE.border, 0.75);
  line(slide, x + laneW, y, 0, h, ROADMAP_STYLE.border, 0.75);
  text(slide, "认证需求", x + 8, y, laneW - 16, headerH, 12, C.white, true, "center", "mid");
  if (!months.length) text(slide, "时间节点未填写", x + laneW + 20, y, w - laneW - 40, headerH, 12, C.white, false, "center", "mid");
  let offset = 0;
  for (const half of [...new Set(months.map(halfYearOf))]) {
    const count = months.filter((month) => halfYearOf(month) === half).length;
    rect(slide, x + laneW + offset * columnW, y, count * columnW, 28, ROADMAP_STYLE.header, ROADMAP_STYLE.border, 0.75);
    text(slide, `${shortHalf(half)}（${hmosVersionForHalfYear(half)}）`, x + laneW + offset * columnW + 2, y + 2, count * columnW - 4, 24, 8, C.white, true, "center", "mid", { wrap: false, fit: "shrink", name: `medical-half-${half}` });
    offset += count;
  }
  months.forEach((month, index) => {
    const left = x + laneW + index * columnW;
    text(slide, `${Number(month.slice(5))}月`, left, y + 30, columnW, 22, 10, C.white, false, "center");
    line(slide, left, y + 28, 0, h - 28, ROADMAP_STYLE.border, 0.6);
  });
  let rowY = y + headerH;
  if (!page.groups.length) text(slide, "暂无医疗认证需求", x + laneW + 20, rowY + 28, w - laneW - 40, 30, 12, C.muted, false, "center");
  for (const { group, requirement } of medicalRoadmapDemandRows(page.groups)) {
    const layout = medicalRowLayout(group);
    const { height } = layout;
    rect(slide, x, rowY, laneW, height, ROADMAP_STYLE.lane, ROADMAP_STYLE.border, 0.75);
    text(slide, layout.titleLines.join("\n"), x + 14, rowY + 12, 212, layout.titleHeight, 7, ROADMAP_STYLE.ink, true, "left", "top", { margin: 0, wrap: false, lineSpacing: 9, name: `medical-title-${requirement.id}` });
    const badgeY = rowY + 12 + layout.titleHeight + 8;
    roundRect(slide, x + 14, badgeY, layout.badgeWidth, layout.badgeHeight, requirement.source === "NMPA认证" ? "D64D85" : C.blue, C.white, 0, 7);
    text(slide, layout.badgeLines.join("\n"), x + 21, badgeY + 2, layout.badgeWidth - 14, layout.badgeHeight - 4, 7, C.white, true, "left", "top", { wrap: false, lineSpacing: 9, name: `medical-certification-${requirement.id}` });
    const nodes = medicalSubmissionTimeline(requirement.targetMonth, requirement.medicalSubmissionMilestones).filter((node) => node.month);
    const sorted = nodes.map((node) => node.month!).sort();
    const nodeY = rowY + 24;
    const visible = months.filter((month) => month >= sorted[0] && month <= sorted.at(-1)!);
    if (visible.length) {
      const first = sorted[0] < months[0] ? 0 : months.indexOf(visible[0]) + 0.5;
      const last = sorted.at(-1)! > months.at(-1)! ? months.length : months.indexOf(visible.at(-1)!) + 0.5;
      line(slide, x + laneW + first * columnW, nodeY, (last - first) * columnW, 0, ROADMAP_STYLE.connector, 1.5);
      if (first === 0) text(slide, "‹", x + laneW, nodeY - 9, 12, 18, 8, C.muted);
      if (last === months.length) text(slide, "›", x + w - 12, nodeY - 9, 12, 18, 8, C.muted);
    }
    months.forEach((month, index) => {
      const values = nodes.filter((node) => node.month === month);
      if (!values.length) return;
      const left = x + laneW + index * columnW;
      const isSubmission = values.some((node) => node.derived);
      const nodeColor = isSubmission ? ROADMAP_STYLE.submission : ROADMAP_STYLE.blue;
      ellipse(slide, left + columnW / 2 - 8, nodeY - 8, 16, 16, "E8F1FB", "E8F1FB", 0);
      ellipse(slide, left + columnW / 2 - 5, nodeY - 5, 10, 10, nodeColor, nodeColor, 0);
      text(slide, values.map((node) => node.label.replace("时间", "")).join("\n"), left + 2, nodeY + 12, columnW - 4, values.length * 12, 7, isSubmission ? "B74438" : "325976", true, "center", "top", { margin: 0, wrap: false, lineSpacing: 9, name: `medical-node-${requirement.id}-${month}` });
    });
    if (!nodes.length) text(slide, "时间节点未填写", x + laneW + 16, rowY + 22, w - laneW - 32, 25, 8, C.muted);
    line(slide, x, rowY + height, w, 0, C.faint, 0.7);
    rowY += height;
  }
  text(slide, months.length ? `已匹配医疗需求，独立自动范围；本段 ${months[0]} 至 ${months.at(-1)}` : "已匹配医疗需求，暂无已填写日期", x, y + h + 9, w, 20, 9, C.muted);
}

function wrapText(value: string, maxWidth: number, fontSize: number, bold: boolean): string[] {
  const chars = Array.from(value || " ");
  const lines: string[] = [];
  let lineValue = "";
  let lineWidth = 0;
  for (const char of chars) {
    if (char === "\r") continue;
    if (char === "\n") {
      lines.push(lineValue);
      lineValue = "";
      lineWidth = 0;
      continue;
    }
    const charWidth = measureTextWidth(char, fontSize, bold);
    if (lineValue && lineWidth + charWidth > maxWidth) {
      lines.push(lineValue);
      lineValue = char;
      lineWidth = charWidth;
    } else {
      lineValue += char;
      lineWidth += charWidth;
    }
  }
  if (lineValue || !lines.length) lines.push(lineValue);
  return lines;
}

export function measureTextWidth(value: string, fontSize: number, bold = false): number {
  const pixelSize = pointsToPixels(fontSize);
  const factor = bold ? 1.04 : 1;
  return Array.from(value).reduce((sum, char) => {
    if (/\s/u.test(char)) return sum + pixelSize * 0.3;
    if (/[\u3000-\u9fff\uf900-\ufaff\uff01-\uff60]/u.test(char)) return sum + pixelSize * factor;
    if (/[MWmw]/u.test(char)) return sum + pixelSize * 0.9 * factor;
    if (/[A-Z0-9]/u.test(char)) return sum + pixelSize * 0.62 * factor;
    return sum + pixelSize * 0.54 * factor;
  }, 0);
}

function drawSummaryTrend(slide: PptxSlide, summaries: HalfYearSummary[]): void {
  const plot = { x: 112, y: 389, w: 610, h: 172 };
  const maximum = Math.max(1, ...summaries.flatMap((item) => [item.bySide.device, item.bySide.app, item.bySide.cloud]));
  [0, 0.33, 0.66, 1].forEach((ratio, index) => {
    const y = plot.y + plot.h - ratio * plot.h;
    line(slide, plot.x, y, plot.w, 0, C.faint2, 1);
    text(slide, formatNumber((maximum * ratio)), 60, y - 8, 44, 16, 9, C.muted, false, "right");
  });
  const groupWidth = plot.w / Math.max(1, summaries.length);
  const barWidth = Math.min(18, groupWidth / 5);
  summaries.forEach((summary, halfIndex) => {
    const center = plot.x + groupWidth * (halfIndex + 0.5);
    (["device", "app", "cloud"] as WorkloadSide[]).forEach((side, sideIndex) => {
      const value = summary.bySide[side];
      const height = value / maximum * plot.h;
      const x = center - barWidth * 1.7 + sideIndex * barWidth * 1.2;
      if (height > 0) rect(slide, x, plot.y + plot.h - height, barWidth, height, SIDE_META[side].color, SIDE_META[side].color, 0, 3);
    });
    text(slide, shortHalf(summary.halfYear), center - groupWidth / 2, 570, groupWidth, 20, 10, C.muted, true, "center");
  });
  (["device", "app", "cloud"] as WorkloadSide[]).forEach((side, index) => {
    const x = 180 + index * 160;
    ellipse(slide, x, 611, 8, 8, SIDE_META[side].color, SIDE_META[side].color, 0);
    text(slide, SIDE_META[side].label, x + 14, 605, 140, 18, 10, C.muted);
  });
}

function drawSideCategoryStructure(slide: PptxSlide, summaries: HalfYearSummary[]): void {
  (["device", "app", "cloud"] as WorkloadSide[]).forEach((side, row) => {
    const y = 410 + row * 60;
    const values = CATEGORIES.map((category) => summaries.reduce((sum, item) => sum + item.categoryBySide[category][side], 0));
    const total = values.reduce((sum, value) => sum + value, 0);
    text(slide, SIDE_META[side].label, 808, y + 8, 98, 18, 10, C.ink, true);
    let cursor = 910;
    values.forEach((value, index) => {
      const width = total ? value / total * 240 : index === 0 ? 240 : 0;
      if (width > 0) rect(slide, cursor, y, width, 34, total ? CATEGORY_COLORS[CATEGORIES[index]] : C.faint2, total ? CATEGORY_COLORS[CATEGORIES[index]] : C.faint2, 0, 4);
      if (total && width >= 48) text(slide, formatNumber(value), cursor, y + 8, width, 18, 9, C.white, true, "center");
      cursor += width;
    });
    text(slide, formatNumber(total), 1160, y + 8, 42, 18, 10, C.ink, true, "right");
  });
  CATEGORIES.forEach((category, index) => {
    const x = 812 + index * 126;
    ellipse(slide, x, 588, 8, 8, CATEGORY_COLORS[category], CATEGORY_COLORS[category], 0);
    text(slide, category, x + 14, 582, 102, 18, 10, C.muted);
  });
}

function metricCard(slide: PptxSlide, x: number, y: number, w: number, h: number, label: string, value: number, color: string): void {
  roundRect(slide, x, y, w, h, C.white, C.faint, 1, 12);
  text(slide, label, x + 20, y + 17, w - 40, 20, 11, C.muted, true);
  text(slide, formatNumber(value), x + 20, y + 43, w - 78, 38, 25, color, true);
  text(slide, "PM", x + w - 58, y + 55, 34, 20, 10, C.muted, true, "right");
}

function totals(summaries: HalfYearSummary[]) {
  return summaries.reduce((result, item) => ({
    total: result.total + item.totalWorkload,
    device: result.device + item.bySide.device,
    app: result.app + item.bySide.app,
    cloud: result.cloud + item.bySide.cloud,
    researchApp: result.researchApp + item.researchAppWorkload,
    algorithm: result.algorithm + item.algorithmWorkload,
    unallocated: result.unallocated + item.unallocatedWorkload,
  }), { total: 0, device: 0, app: 0, cloud: 0, researchApp: 0, algorithm: 0, unallocated: 0 });
}

function addSlideHeader(slide: PptxSlide, page: number, eyebrow: string, titleValue: string, subtitle: string): void {
  addChrome(slide, page, eyebrow, false);
  text(slide, titleValue, 64, 105, 760, 42, 27, C.ink, true);
  text(slide, subtitle, 64, 153, 850, 24, 12, C.muted);
}

function addChrome(slide: PptxSlide, page: number, section: string, dark: boolean): void {
  text(slide, section, 64, 26, 360, 18, 9, dark ? "AFC0CB" : C.muted, true);
  line(slide, 64, 58, 1152, 0, dark ? C.navy2 : C.faint, 1);
}

function rect(slide: PptxSlide, x: number, y: number, w: number, h: number, fill: string, border: string, borderWidth = 1, radius = 0, name?: string): void {
  // PptxGenJS divides by the smallest round-rectangle dimension. Zero-sized
  // rounded shapes therefore emit `fmla="val Infinity"`, which PowerPoint
  // treats as corrupt OpenXML and repeatedly offers to repair.
  if (w <= 0 || h <= 0) return;
  slide.addShape(radius ? "roundRect" : "rect", { x: x * PX, y: y * PX, w: w * PX, h: h * PX, rectRadius: radius * PX, fill: { color: fill }, line: { color: border, width: borderWidth }, objectName: name });
}

function roundRect(slide: PptxSlide, x: number, y: number, w: number, h: number, fill: string, border: string, borderWidth = 1, radius = 8): void {
  rect(slide, x, y, w, h, fill, border, borderWidth, radius);
}

function ellipse(slide: PptxSlide, x: number, y: number, w: number, h: number, fill: string, border: string, borderWidth = 1): void {
  slide.addShape("ellipse", { x: x * PX, y: y * PX, w: w * PX, h: h * PX, fill: { color: fill }, line: { color: border, width: borderWidth } });
}

function line(slide: PptxSlide, x: number, y: number, w: number, h: number, color: string, width = 1): void {
  slide.addShape("line", { x: x * PX, y: y * PX, w: w * PX, h: h * PX, line: { color, width } });
}

function text(
  slide: PptxSlide,
  value: string,
  x: number,
  y: number,
  w: number,
  h: number,
  fontSize: number,
  color: string,
  bold = false,
  align: "left" | "center" | "right" = "left",
  valign: "top" | "mid" | "bottom" = "top",
  extra: { margin?: number; breakLine?: boolean; name?: string; wrap?: boolean; lineSpacing?: number; fit?: "none" | "shrink" } = {},
): void {
  slide.addText(value, {
    x: x * PX,
    y: y * PX,
    w: w * PX,
    h: h * PX,
    fontFace: FONT,
    fontSize,
    color,
    bold,
    margin: extra.margin ?? 0,
    breakLine: extra.breakLine,
    fit: extra.fit ?? (extra.wrap === false ? "none" : "shrink"),
    wrap: extra.wrap,
    lineSpacing: extra.lineSpacing,
    paraSpaceAfter: extra.lineSpacing ? 0 : undefined,
    paraSpaceBefore: extra.lineSpacing ? 0 : undefined,
    valign,
    align,
    objectName: extra.name,
  });
}

function notes(slide: PptxSlide, value: string): void {
  slide.addNotes(`[Sources]\n- Data: local offline workspace\n- Purpose: ${value}`);
}

function pointsToPixels(points: number): number {
  return points * 96 / 72;
}

function shortHalf(value: string): string {
  return value.replace(/^(\d{2})(\d{2})/, "$2");
}

function formatHalfYear(value: string): string {
  return value.replace("H", " H");
}

function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(2).replace(/0+$/, "").replace(/\.$/, "");
}
