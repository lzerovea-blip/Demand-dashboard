import PptxGenJSImport from "pptxgenjs";
import {
  buildPptExportPlan,
  groupsForHalfYear,
  roadmapBadgeKind,
  visibleRequirementTitles,
  type PptExportPlan,
  type PptRoadmapPagePlan,
  type RoadmapBadgeKind,
} from "../shared/pptExport.js";
import { groupMatchesRoadmapLane, hmosVersionForHalfYear, roadmapLanesForTrack, type HalfYearSummary, type RoadmapGroup } from "../shared/roadmap.js";
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
  app: { label: "App 侧", color: C.blue },
  cloud: { label: "云侧", color: C.coral },
};

const ROADMAP_TABLE = {
  x: 27.6,
  y: 76.93,
  w: 1223.07,
  h: 580.87,
  laneW: 128,
  headerH: 68,
} as const;

type PptxPresentation = any;
type PptxSlide = any;
const PptxGenJS = PptxGenJSImport as unknown as { new(): PptxPresentation };

export interface RoadmapCardLayout {
  group: RoadmapGroup;
  badge: RoadmapBadgeKind;
  fontSize: number;
  margin: number;
  gapAfter: number;
  visibleRequirements: string[];
  titleLines: string[];
  bodyLines: string[];
  titleLineHeight: number;
  bodyLineHeight: number;
  bodyOffset: number;
  height: number;
  top: number;
  badgeLeft: number;
  badgeTop: number;
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
  security: PptExportSecurityContext = defaultSecurityContext(),
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

  addCoverSlide(pptx, plan, security);
  for (const section of plan.selectedSections) {
    if (section === "summary") addSummarySlide(pptx, plan);
    else {
      const page = plan.roadmapPages.find((item) => item.track === section);
      if (page) paginateRoadmapPage(page).forEach((part) => addRoadmapSlide(pptx, part));
    }
  }

  const output = await pptx.write({ outputType: "nodebuffer", compression: true });
  return Buffer.from(output as Uint8Array);
}

/** Actual count, including overflow pages; shared by validation and the export result. */
export function presentationSlideCount(plan: PptExportPlan): number {
  return 1 + plan.selectedSections.reduce((count, section) => {
    if (section === "summary") return count + 1;
    const page = plan.roadmapPages.find((item) => item.track === section);
    return count + (page ? paginateRoadmapPage(page).length : 0);
  }, 0);
}

function addCoverSlide(pptx: PptxPresentation, plan: PptExportPlan, security: PptExportSecurityContext): void {
  const slide = pptx.addSlide();
  slide.background = { color: C.ink };
  addChrome(slide, 1, "ROADMAP REVIEW", true);
  rect(slide, 64, 116, 40, 4, C.mint, C.mint, 0, 2);
  text(slide, "解决方案版本 · 半年度视图", 64, 137, 520, 26, 13, C.mint, true);
  text(slide, plan.materialName, 64, 178, 620, 140, 48, C.white, true, "left", "mid", { breakLine: true });
  text(slide, `${formatHalfYear(plan.start)} — ${formatHalfYear(plan.end)}`, 64, 338, 520, 34, 20, "AFC0CB", true);

  const total = totals(plan.summaries);
  const stats: Array<[string, string]> = [
    ["时间范围", `${plan.halfYears.length} 个半年`],
    ["需求数量", `${plan.requirements.length} 条`],
    ["总工作量", `${formatNumber(total.total + total.unallocated)} PM`],
    ["生成日期", formatDate(security.exportedAt)],
  ];
  stats.forEach(([label, value], index) => {
    const x = 64 + index * 166;
    if (index > 0) line(slide, x - 22, 448, 0, 76, C.navy2, 1);
    text(slide, label, x, 444, 135, 20, 11, "AFC0CB", true);
    text(slide, value, x, 476, 145, 32, 20, C.white, true);
  });
  text(slide, `企业 IRM 权限有效期至 ${formatDate(security.expiration)}`, 64, 614, 520, 20, 11, C.mint, true);

  roundRect(slide, 760, 116, 457, 471, C.navy, C.navy2, 1, 18);
  text(slide, "MATERIAL WINDOW", 796, 149, 220, 20, 10, C.mint, true);
  text(slide, "统一的半年度路标视图", 796, 177, 340, 36, 24, C.white, true);
  text(slide, "汇总分析与七类路标采用一致页眉、间距、卡片和标识规则。", 796, 221, 350, 48, 13, "AFC0CB");
  line(slide, 830, 299, 0, 205, "60717D", 2);
  plan.halfYears.forEach((halfYear, index) => {
    const y = 289 + index * 68;
    ellipse(slide, 820, y, 20, 20, index === 0 ? C.mint : C.navy2, index === 0 ? C.mint : "60717D", 2);
    text(slide, formatHalfYear(halfYear), 862, y - 2, 150, 24, 15, C.white, true, "left", "mid");
  });
  text(slide, "解决方案需求路标材料", 64, 645, 400, 22, 11, "AFC0CB");
  notes(slide, "封面数据来自本地工作区当前导出时间范围。所有元素均为可编辑原生对象。");
}

function addSummarySlide(pptx: PptxPresentation, plan: PptExportPlan): void {
  const slide = pptx.addSlide();
  slide.background = { color: C.paper };
  addSlideHeader(slide, 2, "WORKLOAD SUMMARY", "需求工作量汇总分析", `${formatHalfYear(plan.start)}—${formatHalfYear(plan.end)} · ${plan.requirements.length} 条需求`);
  const total = totals(plan.summaries);
  const metrics: Array<[string, number, string]> = [
    ["总工作量", total.total + total.unallocated, C.ink],
    ["设备侧", total.device, C.mintDeep],
    ["App 侧", total.app, C.blue],
    ["云侧", total.cloud, C.coral],
  ];
  metrics.forEach(([label, value, color], index) => metricCard(slide, 64 + index * 288, 194, 266, 101, label, value, color));

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
  notes(slide, "三端和分类均按需求实际工作量计算；本页不展示七类路标投入。");
}

function addRoadmapSlide(pptx: PptxPresentation, page: PaginatedRoadmapPage): void {
  const slide = pptx.addSlide();
  slide.background = { color: C.paper };
  const slideNumber = pptx._slides.length;
  const continuation = page.continuationCount > 1
    ? `（${page.continuationIndex ? "续 " : ""}${page.continuationIndex + 1}/${page.continuationCount}）`
    : "";
  text(slide, `${page.track}路标${continuation}`, 50, 20, 650, 38, 24, C.ink, true, "left", "mid");
  text(slide, String(slideNumber).padStart(2, "0"), 1170, 26, 46, 18, 9, C.ink, true, "right");
  line(slide, 50, 64, 1170, 0, C.faint, 1);
  drawRoadmapTable(slide, page);
  notes(slide, `${page.track}路标按半年展示；每条圆点文本对应一条独立需求记录，不展示产品名称。`);
}

export function paginateRoadmapPage(page: PptRoadmapPagePlan): PaginatedRoadmapPage[] {
  const lanes = roadmapLanesForTrack(page.track);
  const cellWidth = (ROADMAP_TABLE.w - ROADMAP_TABLE.laneW) / page.halfYears.length;
  const cellHeight = (ROADMAP_TABLE.h - ROADMAP_TABLE.headerH) / lanes.length;
  const chunksByCell: RoadmapGroup[][][] = [];

  // Split whole cards independently per cell. A continuation retains every half-year and lane.
  for (const lane of lanes) {
    for (const halfYear of page.halfYears) {
      const groups = groupsForHalfYear(page.groups, halfYear).filter((group) => groupMatchesRoadmapLane(group, lane));
      const chunks: RoadmapGroup[][] = [];
      let current: RoadmapGroup[] = [];
      for (const group of groups) {
        const candidate = [...current, group];
        try {
          layoutRoadmapCell(candidate, cellWidth, cellHeight);
          current = candidate;
        } catch (error) {
          if (!(error instanceof RoadmapCellCapacityError)) throw error;
          // Pagination cannot fix an individual card that is taller than an entire cell.
          try {
            layoutRoadmapCell([group], cellWidth, cellHeight);
          } catch (singleCardError) {
            if (!(singleCardError instanceof RoadmapCellCapacityError)) throw singleCardError;
            throw new RoadmapCellCapacityError(`PPT 路标单张卡片超过可排版容量：${page.track}路标 / ${lane} / ${halfYear} / ${group.cardTitle}。请缩短该领域标题或需求标题后重试`);
          }
          if (current.length) chunks.push(current);
          current = [group];
        }
      }
      if (current.length) chunks.push(current);
      chunksByCell.push(chunks);
    }
  }

  const continuationCount = Math.max(1, ...chunksByCell.map((chunks) => chunks.length));
  return Array.from({ length: continuationCount }, (_, continuationIndex) => {
    const keys = new Set(chunksByCell.flatMap((chunks) => chunks[continuationIndex] ?? []).map((group) => group.key));
    return {
      ...page,
      groups: page.groups.filter((group) => keys.has(group.key)),
      continuationIndex,
      continuationCount,
    };
  });
}

function drawRoadmapTable(slide: PptxSlide, page: PptRoadmapPagePlan): void {
  const { x, y, w, h, laneW, headerH } = ROADMAP_TABLE;
  const lanes = roadmapLanesForTrack(page.track);
  const halfW = (w - laneW) / page.halfYears.length;
  const bodyH = h - headerH;
  const laneH = bodyH / lanes.length;
  // Labels are editable text boxes below; the native table only supplies the grid.
  // Drawing labels in both places produces overlapping text on every continuation.
  const rows = Array.from({ length: lanes.length + 1 }, () => Array(page.halfYears.length + 1).fill(""));
  const colW = [laneW * PX, ...page.halfYears.map(() => halfW * PX)];
  const rowH = [headerH * PX, ...lanes.map(() => laneH * PX)];
  slide.addTable(rows, {
    x: x * PX,
    y: y * PX,
    w: w * PX,
    h: h * PX,
    colW,
    rowH,
    border: { type: "solid", color: "D2D2D2", pt: 1 },
    fill: C.white,
    color: C.ink,
    fontFace: FONT,
    fontSize: 12,
    margin: 0,
    autoFit: false,
    valign: "middle",
    align: "center",
  });
  rect(slide, x, y, w, headerH, C.red, C.red, 0);
  text(slide, "类型", x, y, laneW, headerH, 16, C.white, true, "center", "mid");
  page.halfYears.forEach((halfYear, index) => {
    text(slide, `${shortHalf(halfYear)}\n（${hmosVersionForHalfYear(halfYear)}）`, x + laneW + index * halfW, y + 6, halfW, headerH - 12, 12, C.white, true, "center", "mid");
  });
  lanes.forEach((lane, laneIndex) => {
    const rowY = y + headerH + laneIndex * laneH;
    text(slide, lane, x + 4, rowY, laneW - 8, laneH, lane.length > 4 ? 12 : 16, C.ink, true, "center", "mid");
    page.halfYears.forEach((halfYear, halfIndex) => {
      const cellX = x + laneW + halfIndex * halfW;
      const groups = groupsForHalfYear(page.groups, halfYear).filter((group) => groupMatchesRoadmapLane(group, lane));
      if (groups.length) drawRoadmapCell(slide, groups, cellX, rowY, halfW, laneH, page.track, lane, halfYear);
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
    const width = cellW - card.margin * 2;
    const titleHeight = card.titleLines.length * card.titleLineHeight;
    // Use exactly the measured lines. A second Office wrap can add an unmeasured
    // line (or strand the bullet), invalidating the pagination capacity check.
    text(slide, card.titleLines.join("\n"), left, top, width, titleHeight, card.fontSize, C.ink, true, "left", "top", { margin: 0, wrap: false, lineSpacing: card.titleLineHeight * 72 / 96, name: `roadmap-title-${track}-${lane}-${halfYear}-${index}` });
    const bodyTop = top + card.bodyOffset;
    text(slide, card.bodyLines.join("\n"), left, bodyTop, width, card.height - card.bodyOffset, card.fontSize, C.ink, false, "left", "top", { margin: 0, wrap: false, lineSpacing: card.bodyLineHeight * 72 / 96, name: `roadmap-body-${track}-${lane}-${halfYear}-${index}` });
    if (card.badge === "product") {
      roundRect(slide, cellX + card.badgeLeft, cellY + card.badgeTop, 58, 17, C.navy, C.navy, 0, 4);
      text(slide, "产品专属", cellX + card.badgeLeft, cellY + card.badgeTop, 58, 17, 8.5, C.white, true, "center", "mid", { margin: 0, name: `roadmap-product-${track}-${lane}-${halfYear}-${index}` });
    } else if (card.badge === "experience") {
      ellipse(slide, cellX + card.badgeLeft, cellY + card.badgeTop, 17, 17, C.navy, C.navy, 0);
      text(slide, "↑", cellX + card.badgeLeft, cellY + card.badgeTop - 1, 17, 18, 11, C.white, true, "center", "mid", { margin: 0, name: `roadmap-experience-${track}-${lane}-${halfYear}-${index}` });
    }
  }
}

export function layoutRoadmapCell(groups: RoadmapGroup[], width: number, height: number): RoadmapCellLayout {
  const attempts: Array<{ fontSize: number; margin: number; gap: number; visibleLimit: number }> = [
    ...[10, 9.5, 9, 8.5, 8].map((fontSize) => ({ fontSize, margin: 8, gap: 7, visibleLimit: 3 })),
    { fontSize: 8, margin: 4, gap: 4, visibleLimit: 3 },
    { fontSize: 8, margin: 4, gap: 4, visibleLimit: 2 },
    { fontSize: 8, margin: 4, gap: 4, visibleLimit: 1 },
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
          const placed = { ...card, top, badgeTop: top + card.badgeTop };
          top += card.height + attempt.gap;
          return placed;
        }),
      };
    }
  }
  throw new RoadmapCellCapacityError("PPT 路标单元格内容超过可排版容量：需要拆分续页");
}

function calculateCardLayouts(groups: RoadmapGroup[], width: number, fontSize: number, margin: number, gap: number, visibleLimit: number): RoadmapCardLayout[] {
  const textWidth = width - margin * 2;
  const bodyLineHeight = pointsToPixels(fontSize) * 1.25;
  return groups.map((group, index) => {
    const badge = roadmapBadgeKind(group);
    const titleLineHeight = Math.max(pointsToPixels(fontSize) * 1.18, badge ? 17 : 0);
    const badgeWidth = badge === "product" ? 58 : badge === "experience" ? 17 : 0;
    const badgeGap = badge ? 7 : 0;
    let titleLines = wrapText(group.cardTitle, textWidth, fontSize, true);
    const lastLineWidth = measureTextWidth(titleLines.at(-1) ?? "", fontSize, true);
    if (badge && lastLineWidth + badgeGap + badgeWidth > textWidth) {
      titleLines = wrapText(group.cardTitle, Math.max(20, textWidth - badgeWidth - badgeGap), fontSize, true);
    }
    const visibleRequirements = visibleRequirementTitlesForLayout(group, visibleLimit);
    // Reserve indentation and a small font-substitution tolerance before wrapping.
    const bodyLines = visibleRequirements.flatMap((item) => wrapText(item, textWidth - measureTextWidth("· ", fontSize) - 4, fontSize, false)
      .map((part, partIndex) => `${partIndex ? "  " : "· "}${part}`));
    const titleHeight = titleLines.length * titleLineHeight;
    const bodyOffset = titleHeight + 2;
    const height = bodyOffset + Math.max(1, bodyLines.length) * bodyLineHeight;
    const dynamicLastWidth = measureTextWidth(titleLines.at(-1) ?? "", fontSize, true);
    const badgeTop = (titleLines.length - 1) * titleLineHeight + Math.max(0, (titleLineHeight - 17) / 2);
    return {
      group,
      badge,
      fontSize,
      margin,
      gapAfter: index === groups.length - 1 ? 0 : gap,
      visibleRequirements,
      titleLines,
      bodyLines,
      titleLineHeight,
      bodyLineHeight,
      bodyOffset,
      height,
      top: 0,
      badgeLeft: Math.min(width - margin - badgeWidth, margin + dynamicLastWidth + badgeGap),
      badgeTop,
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
    const x = 240 + index * 125;
    ellipse(slide, x, 611, 8, 8, SIDE_META[side].color, SIDE_META[side].color, 0);
    text(slide, SIDE_META[side].label, x + 14, 605, 90, 18, 10, C.muted);
  });
}

function drawSideCategoryStructure(slide: PptxSlide, summaries: HalfYearSummary[]): void {
  (["device", "app", "cloud"] as WorkloadSide[]).forEach((side, row) => {
    const y = 410 + row * 60;
    const values = CATEGORIES.map((category) => summaries.reduce((sum, item) => sum + item.categoryBySide[category][side], 0));
    const total = values.reduce((sum, value) => sum + value, 0);
    text(slide, SIDE_META[side].label, 808, y + 8, 58, 18, 11, C.ink, true);
    let cursor = 870;
    values.forEach((value, index) => {
      const width = total ? value / total * 280 : index === 0 ? 280 : 0;
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
    unallocated: result.unallocated + item.unallocatedWorkload,
  }), { total: 0, device: 0, app: 0, cloud: 0, unallocated: 0 });
}

function addSlideHeader(slide: PptxSlide, page: number, eyebrow: string, titleValue: string, subtitle: string): void {
  addChrome(slide, page, eyebrow, false);
  text(slide, titleValue, 64, 105, 760, 42, 27, C.ink, true);
  text(slide, subtitle, 64, 153, 850, 24, 12, C.muted);
}

function addChrome(slide: PptxSlide, page: number, section: string, dark: boolean): void {
  text(slide, section, 64, 26, 360, 18, 9, dark ? "AFC0CB" : C.muted, true);
  text(slide, String(page).padStart(2, "0"), 1170, 26, 46, 18, 9, dark ? C.white : C.ink, true, "right");
  line(slide, 64, 58, 1152, 0, dark ? C.navy2 : C.faint, 1);
}

function rect(slide: PptxSlide, x: number, y: number, w: number, h: number, fill: string, border: string, borderWidth = 1, radius = 0): void {
  // PptxGenJS divides by the smallest round-rectangle dimension. Zero-sized
  // rounded shapes therefore emit `fmla="val Infinity"`, which PowerPoint
  // treats as corrupt OpenXML and repeatedly offers to repair.
  if (w <= 0 || h <= 0) return;
  slide.addShape(radius ? "roundRect" : "rect", { x: x * PX, y: y * PX, w: w * PX, h: h * PX, rectRadius: radius * PX, fill: { color: fill }, line: { color: border, width: borderWidth } });
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
  extra: { margin?: number; breakLine?: boolean; name?: string; wrap?: boolean; lineSpacing?: number } = {},
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
    fit: extra.wrap === false ? "none" : "shrink",
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

function formatDate(date: Date): string {
  return `${date.getFullYear()}.${String(date.getMonth() + 1).padStart(2, "0")}.${String(date.getDate()).padStart(2, "0")}`;
}

function formatNumber(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(2).replace(/0+$/, "").replace(/\.$/, "");
}
