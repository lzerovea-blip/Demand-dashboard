import { CATEGORIES, ROADMAP_TRACKS, SOURCES } from "./types.js";
import type {
  AppSnapshot,
  GroupOverride,
  Level,
  Requirement,
  RequirementCategory,
  RequirementSource,
  Track,
} from "./types.js";
import { sumWorkloadBreakdown, type WorkloadBreakdown, type WorkloadSide } from "./workload.js";
import { splitTextIntoLines } from "./text.js";

export interface RoadmapGroup {
  key: string;
  domainId: string;
  domainName: string;
  domainSortOrder: number;
  source: RequirementSource;
  track: Track;
  level: Level | null;
  overseasRegions: Requirement["overseasRegions"];
  targetMonth: string;
  cardTitle: string;
  cardSummary: string;
  sortOrder?: number;
  requirements: Requirement[];
  productIds: string[];
  productNames: string[];
  requirementProductNames: Record<string, string[]>;
  totalWorkloadPm: number;
  workloadBreakdown: WorkloadBreakdown;
  categories: RequirementCategory[];
}

export interface HalfYearSummary {
  halfYear: string;
  bySource: Record<RequirementSource, number>;
  byTrack: Record<Track, number>;
  sportsWorkload: number;
  healthWorkload: number;
  byCategory: Record<RequirementCategory, number>;
  totalWorkload: number;
  deviceWorkload: number;
  appWorkload: number;
  cloudWorkload: number;
  unallocatedWorkload: number;
  bySide: Record<WorkloadSide, number>;
  categoryBySide: Record<RequirementCategory, Record<WorkloadSide, number>>;
  requirementCount: number;
}

export function trackOf(source: RequirementSource): Track {
  if (source.startsWith("运动")) return "运动";
  if (source.startsWith("健康")) return "健康";
  if (source.startsWith("AI")) return "AI";
  if (source === "NMPA送检" || source === "CE送检") return "医疗送检";
  return source as Track;
}

export function levelOf(source: RequirementSource): Level | null {
  if (!source.startsWith("运动") && !source.startsWith("健康")) return null;
  return source.replace(/^运动|^健康/, "") as Level;
}

export function groupKeyFor(requirement: Pick<Requirement, "domainId" | "source" | "targetMonth">): string {
  return `${requirement.domainId}::${requirement.source}::${requirement.targetMonth}`;
}

export function halfYearOf(targetMonth: string): string {
  const [yearRaw, monthRaw] = targetMonth.split("-");
  const year = Number(yearRaw);
  const month = Number(monthRaw);
  if (!Number.isInteger(year) || month < 1 || month > 12) {
    throw new Error(`无效上线年月：${targetMonth}`);
  }
  return `${year}H${month <= 6 ? 1 : 2}`;
}

function parseHalfYear(value: string): { year: number; half: 1 | 2 } {
  const match = /^(\d{4})H([12])$/.exec(value);
  if (!match) throw new Error(`无效半年：${value}`);
  return { year: Number(match[1]), half: Number(match[2]) as 1 | 2 };
}

export function shortHalfYear(value: string): string {
  const { year, half } = parseHalfYear(value);
  return `${String(year).slice(-2)}H${half}`;
}

export function hmosVersionForHalfYear(value: string): string {
  const { year, half } = parseHalfYear(value);
  if (year < 2024) return "历史版本";
  const h1Major = 6 + (year - 2026);
  return `HMOS ${half === 1 ? `${h1Major}.1` : `${h1Major + 1}.0`}`;
}

export function hmosVersionForMonth(targetMonth: string): string {
  return hmosVersionForHalfYear(halfYearOf(targetMonth));
}

export function formatHalfYearWithHmos(value: string): string {
  return `${shortHalfYear(value)}（${hmosVersionForHalfYear(value)}）`;
}

export function halfYearSequence(start: string, end: string): string[] {
  const serialOf = (value: string) => {
    const { year, half } = parseHalfYear(value);
    return year * 2 + half - 1;
  };
  const from = serialOf(start);
  const to = serialOf(end);
  if (from > to) throw new Error("起始半年不能晚于结束半年");
  return Array.from({ length: to - from + 1 }, (_, index) => {
    const serial = from + index;
    const year = Math.floor(serial / 2);
    const half = (serial % 2) + 1;
    return `${year}H${half}`;
  });
}

export function monthsForHalfYear(halfYear: string): string[] {
  const { year, half } = parseHalfYear(halfYear);
  const startMonth = half === 1 ? 1 : 7;
  return Array.from({ length: 6 }, (_, index) => `${year}-${String(startMonth + index).padStart(2, "0")}`);
}

export function requirementsInRange(requirements: Requirement[], start: string, end: string): Requirement[] {
  const periods = new Set(halfYearSequence(start, end));
  return requirements.filter((requirement) => periods.has(halfYearOf(requirement.targetMonth)));
}

export function buildRoadmapGroups(snapshot: AppSnapshot, start: string, end: string): RoadmapGroup[] {
  const domains = new Map(snapshot.domains.map((item) => [item.id, item]));
  const products = new Map(snapshot.products.map((item) => [item.id, item]));
  const overrides = new Map(snapshot.groupOverrides.map((item) => [item.groupKey, item]));
  const groups = new Map<string, Requirement[]>();

  for (const requirement of requirementsInRange(snapshot.requirements, start, end)) {
    const key = groupKeyFor(requirement);
    groups.set(key, [...(groups.get(key) ?? []), requirement]);
  }

  return [...groups.entries()]
    .map(([key, requirements]) => {
      const seed = requirements[0];
      const domainName = domains.get(seed.domainId)?.name ?? "未命名领域";
      const override = overrides.get(key);
      const productIds = [...new Set(requirements.flatMap((item) => item.productIds))];
      const overseasRegions = [...new Set(requirements.flatMap((item) => item.overseasRegions))];
      const categories = [...new Set(requirements.map((item) => item.category))];
      const requirementProductNames = Object.fromEntries(
        requirements.map((item) => [
          item.id,
          item.productIds.map((id) => products.get(id)?.name).filter(Boolean) as string[],
        ]),
      );
      const workloadBreakdown = sumWorkloadBreakdown(requirements);
      return {
        key,
        domainId: seed.domainId,
        domainName,
        domainSortOrder: domains.get(seed.domainId)?.sortOrder ?? Number.MAX_SAFE_INTEGER,
        source: seed.source,
        track: trackOf(seed.source),
        level: levelOf(seed.source),
        overseasRegions,
        targetMonth: seed.targetMonth,
        cardTitle: domainName,
        cardSummary: requirements.map((item) => item.title).join("、"),
        ...(override?.sortOrder === undefined ? {} : { sortOrder: override.sortOrder }),
        requirements: [...requirements].sort((a, b) => a.title.localeCompare(b.title, "zh-CN")),
        productIds,
        productNames: productIds.map((id) => products.get(id)?.name).filter(Boolean) as string[],
        requirementProductNames,
        totalWorkloadPm: workloadBreakdown.total,
        workloadBreakdown,
        categories,
      } satisfies RoadmapGroup;
    })
    .sort((a, b) => a.targetMonth.localeCompare(b.targetMonth)
      || a.source.localeCompare(b.source, "zh-CN")
      || compareRoadmapGroupOrder(a, b));
}

export function compareRoadmapGroupOrder(
  a: Pick<RoadmapGroup, "sortOrder" | "domainSortOrder" | "domainName" | "key">,
  b: Pick<RoadmapGroup, "sortOrder" | "domainSortOrder" | "domainName" | "key">,
): number {
  const aOrdered = Number.isInteger(a.sortOrder) && (a.sortOrder ?? -1) >= 0;
  const bOrdered = Number.isInteger(b.sortOrder) && (b.sortOrder ?? -1) >= 0;
  if (aOrdered !== bOrdered) return aOrdered ? -1 : 1;
  if (aOrdered && bOrdered && a.sortOrder !== b.sortOrder) return (a.sortOrder ?? 0) - (b.sortOrder ?? 0);
  return a.domainSortOrder - b.domainSortOrder
    || a.domainName.localeCompare(b.domainName, "zh-CN")
    || a.key.localeCompare(b.key);
}

export function moveRoadmapGroupKey(
  groupKeys: string[],
  draggedKey: string,
  targetKey?: string,
  position: "before" | "after" | "end" = "end",
): string[] {
  if (!groupKeys.includes(draggedKey)) return groupKeys;
  const next = groupKeys.filter((key) => key !== draggedKey);
  if (position === "end" || !targetKey || !next.includes(targetKey)) return [...next, draggedKey];
  const targetIndex = next.indexOf(targetKey);
  next.splice(targetIndex + (position === "after" ? 1 : 0), 0, draggedKey);
  return next;
}

export function roadmapTrackOrder(groups: Pick<RoadmapGroup, "track">[]): Track[] {
  const hasSports = groups.some((item) => item.track === "运动");
  const hasHealth = groups.some((item) => item.track === "健康");
  const core: Track[] = !hasSports && hasHealth ? ["健康", "运动"] : ["运动", "健康"];
  return [...core, ...ROADMAP_TRACKS.filter((track) => track !== "运动" && track !== "健康")];
}

export type AiRoadmapLane = "AI运动" | "AI健康" | "AI综合";
export type MedicalSubmissionRoadmapLane = "NMPA送检" | "CE送检";
export type RoadmapLane = Level | AiRoadmapLane | MedicalSubmissionRoadmapLane | "需求";

export function roadmapLanesForTrack(track: Track): readonly RoadmapLane[] {
  if (track === "运动" || track === "健康") return ["高阶", "进阶", "基础"];
  if (track === "AI") return ["AI运动", "AI健康", "AI综合"];
  if (track === "医疗送检") return ["NMPA送检", "CE送检"];
  return ["需求"];
}

export function groupMatchesRoadmapLane(group: Pick<RoadmapGroup, "track" | "level" | "source">, lane: RoadmapLane): boolean {
  if (group.track === "运动" || group.track === "健康") return group.level === lane;
  if (group.track === "AI") return group.source === lane;
  if (group.track === "医疗送检") return group.source === lane;
  return lane === "需求";
}

export function roadmapPrimaryLane(group: Pick<RoadmapGroup, "track" | "level" | "source">): RoadmapLane {
  if (group.track === "运动" || group.track === "健康") return group.level ?? "基础";
  if (group.track === "AI") return group.source as AiRoadmapLane;
  if (group.track === "医疗送检") return group.source as MedicalSubmissionRoadmapLane;
  return "需求";
}

export function roadmapMonthColumnTemplate(
  months: string[],
  groups: Array<Pick<RoadmapGroup, "targetMonth">>,
): string {
  const occupiedMonths = new Set(groups.map((item) => item.targetMonth));
  return months
    .map((month) => occupiedMonths.has(month) ? "minmax(280px, 1fr)" : "56px")
    .join(" ");
}

export function roadmapMonthColumnTemplateForTrack(
  months: string[],
  groups: Array<Pick<RoadmapGroup, "targetMonth" | "track">>,
  track: Track,
): string {
  return roadmapMonthColumnTemplate(months, groups.filter((item) => item.track === track));
}

export function buildHalfYearSummaries(requirements: Requirement[], start: string, end: string): HalfYearSummary[] {
  return halfYearSequence(start, end).map((halfYear) => {
    const items = requirements.filter((item) => halfYearOf(item.targetMonth) === halfYear);
    const threeSideTotal = (item: Requirement) => item.deviceWorkloadPm + item.appWorkloadPm + item.cloudWorkloadPm;
    const bySource = Object.fromEntries(
      SOURCES.map((source) => [
        source,
        round(items.filter((item) => item.source === source).reduce((sum, item) => sum + threeSideTotal(item), 0)),
      ]),
    ) as Record<RequirementSource, number>;
    const byTrack = Object.fromEntries(
      ROADMAP_TRACKS.map((track) => [
        track,
        round(items.filter((item) => trackOf(item.source) === track).reduce((sum, item) => sum + threeSideTotal(item), 0)),
      ]),
    ) as Record<Track, number>;
    const sportsWorkload = byTrack.运动;
    const healthWorkload = byTrack.健康;
    const byCategory = Object.fromEntries(CATEGORIES.map((category) => [
      category,
      round(items.filter((item) => item.category === category).reduce((sum, item) => sum + threeSideTotal(item), 0)),
    ])) as Record<RequirementCategory, number>;
    const workloadBreakdown = sumWorkloadBreakdown(items);
    const bySide: Record<WorkloadSide, number> = {
      device: workloadBreakdown.device,
      app: workloadBreakdown.app,
      cloud: workloadBreakdown.cloud,
    };
    const categoryBySide = Object.fromEntries(CATEGORIES.map((category) => {
      const breakdown = sumWorkloadBreakdown(items.filter((item) => item.category === category));
      return [category, { device: breakdown.device, app: breakdown.app, cloud: breakdown.cloud }];
    })) as Record<RequirementCategory, Record<WorkloadSide, number>>;
    return {
      halfYear,
      bySource,
      byTrack,
      sportsWorkload,
      healthWorkload,
      byCategory,
      totalWorkload: round(workloadBreakdown.device + workloadBreakdown.app + workloadBreakdown.cloud),
      deviceWorkload: workloadBreakdown.device,
      appWorkload: workloadBreakdown.app,
      cloudWorkload: workloadBreakdown.cloud,
      unallocatedWorkload: workloadBreakdown.unallocated,
      bySide,
      categoryBySide,
      requirementCount: items.length,
    };
  });
}

export function summarizeCategoryWorkloadForSide(
  summaries: Array<Pick<HalfYearSummary, "categoryBySide">>,
  side: WorkloadSide,
): Record<RequirementCategory, number> {
  return Object.fromEntries(CATEGORIES.map((category) => [
    category,
    round(summaries.reduce((sum, item) => sum + item.categoryBySide[category][side], 0)),
  ])) as Record<RequirementCategory, number>;
}

export function categoryPercentagesForSide(
  summary: Pick<HalfYearSummary, "bySide" | "categoryBySide">,
  side: WorkloadSide,
): Record<RequirementCategory, number> {
  const sideTotal = summary.bySide[side];
  return Object.fromEntries(CATEGORIES.map((category) => [
    category,
    sideTotal ? round((summary.categoryBySide[category][side] / sideTotal) * 100) : 0,
  ])) as Record<RequirementCategory, number>;
}

function round(value: number): number {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

export function findGroupOverride(overrides: GroupOverride[], groupKey: string): GroupOverride | undefined {
  return overrides.find((item) => item.groupKey === groupKey);
}

export function roadmapCardLabel(group: Pick<RoadmapGroup, "cardTitle" | "requirements">): string {
  const titles = group.requirements.slice(0, 3).map((item) => item.title.trim()).filter(Boolean);
  const suffix = group.requirements.length > 3 ? "等" : "";
  return `${group.cardTitle}：${titles.join("、")}${suffix}`;
}

export function roadmapCardDisplayLines(text: string, charactersPerLine = 20, maximumLines = 2): string[] {
  if (charactersPerLine <= 0 || maximumLines <= 0) return [];
  const characters = Array.from(text);
  const maximumCharacters = charactersPerLine * maximumLines;
  const truncated = characters.length > maximumCharacters;
  const visible = truncated
    ? [...characters.slice(0, Math.max(0, maximumCharacters - 1)), "…"]
    : characters.slice(0, maximumCharacters);
  return splitTextIntoLines(visible.join(""), charactersPerLine);
}

export interface RoadmapCardLineParts {
  emphasis: string;
  content: string;
}

export function roadmapCardDisplayParts(
  text: string,
  emphasizedCharacters: number,
  charactersPerLine = 20,
  maximumLines = 2,
): RoadmapCardLineParts[] {
  let characterOffset = 0;
  return roadmapCardDisplayLines(text, charactersPerLine, maximumLines).map((line) => {
    const characters = Array.from(line);
    const emphasisLength = Math.max(0, Math.min(characters.length, emphasizedCharacters - characterOffset));
    characterOffset += characters.length;
    return {
      emphasis: characters.slice(0, emphasisLength).join(""),
      content: characters.slice(emphasisLength).join(""),
    };
  });
}
