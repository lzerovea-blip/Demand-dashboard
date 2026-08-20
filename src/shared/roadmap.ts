import { OVERSEAS_REGIONS, SOURCES } from "./types.js";
import type {
  AppSnapshot,
  GroupOverride,
  Level,
  OverseasRegion,
  Requirement,
  RequirementCategory,
  RequirementSource,
  Track,
} from "./types.js";
import { sumWorkloadBreakdown, type WorkloadBreakdown, type WorkloadSide } from "./workload.js";

export interface RoadmapGroup {
  key: string;
  domainId: string;
  domainName: string;
  source: RequirementSource;
  track: Track;
  level: Level | null;
  overseasRegions: OverseasRegion[];
  targetMonth: string;
  cardTitle: string;
  cardSummary: string;
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
  sportsWorkload: number;
  healthWorkload: number;
  overseasWorkload: number;
  experienceWorkload: number;
  exclusiveWorkload: number;
  totalWorkload: number;
  deviceWorkload: number;
  appWorkload: number;
  cloudWorkload: number;
  unallocatedWorkload: number;
  bySide: Record<WorkloadSide, number>;
  experienceBySide: Record<WorkloadSide, number>;
  requirementCount: number;
}

export function trackOf(source: RequirementSource): Track {
  if (source === "海外研究") return "海外研究";
  return source.startsWith("运动") ? "运动" : "健康";
}

export function levelOf(source: RequirementSource): Level | null {
  if (source === "海外研究") return null;
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

export function halfYearSequence(start: string, end: string): string[] {
  const parse = (value: string) => {
    const match = /^(\d{4})H([12])$/.exec(value);
    if (!match) throw new Error(`无效半年：${value}`);
    return Number(match[1]) * 2 + Number(match[2]) - 1;
  };
  const from = parse(start);
  const to = parse(end);
  if (from > to) throw new Error("起始半年不能晚于结束半年");
  return Array.from({ length: to - from + 1 }, (_, index) => {
    const serial = from + index;
    const year = Math.floor(serial / 2);
    const half = (serial % 2) + 1;
    return `${year}H${half}`;
  });
}

export function monthsForHalfYear(halfYear: string): string[] {
  const match = /^(\d{4})H([12])$/.exec(halfYear);
  if (!match) throw new Error(`无效半年：${halfYear}`);
  const year = Number(match[1]);
  const startMonth = match[2] === "1" ? 1 : 7;
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
        source: seed.source,
        track: trackOf(seed.source),
        level: levelOf(seed.source),
        overseasRegions,
        targetMonth: seed.targetMonth,
        cardTitle: override?.cardTitle.trim() || domainName,
        cardSummary: override?.cardSummary.trim() || requirements.map((item) => item.title).join("、"),
        requirements: [...requirements].sort((a, b) => a.title.localeCompare(b.title, "zh-CN")),
        productIds,
        productNames: productIds.map((id) => products.get(id)?.name).filter(Boolean) as string[],
        requirementProductNames,
        totalWorkloadPm: workloadBreakdown.total,
        workloadBreakdown,
        categories,
      } satisfies RoadmapGroup;
    })
    .sort((a, b) => a.targetMonth.localeCompare(b.targetMonth) || a.source.localeCompare(b.source, "zh-CN"));
}

export function roadmapTrackOrder(groups: Pick<RoadmapGroup, "track">[]): Track[] {
  const hasSports = groups.some((item) => item.track === "运动");
  const hasHealth = groups.some((item) => item.track === "健康");
  return !hasSports && hasHealth ? ["健康", "运动", "海外研究"] : ["运动", "健康", "海外研究"];
}

export type RoadmapLane = Level | OverseasRegion;

export function roadmapLanesForTrack(track: Track): readonly RoadmapLane[] {
  return track === "海外研究" ? OVERSEAS_REGIONS : ["基础", "进阶", "高阶"];
}

export function groupMatchesRoadmapLane(group: Pick<RoadmapGroup, "track" | "level" | "overseasRegions">, lane: RoadmapLane): boolean {
  return group.track === "海外研究"
    ? group.overseasRegions.includes(lane as OverseasRegion)
    : group.level === lane;
}

export function roadmapMonthColumnTemplate(
  months: string[],
  groups: Array<Pick<RoadmapGroup, "targetMonth">>,
): string {
  const occupiedMonths = new Set(groups.map((item) => item.targetMonth));
  return months
    .map((month) => occupiedMonths.has(month) ? "minmax(280px, 1fr)" : "72px")
    .join(" ");
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
    const sportsWorkload = round(items.filter((item) => trackOf(item.source) === "运动").reduce((sum, item) => sum + threeSideTotal(item), 0));
    const healthWorkload = round(items.filter((item) => trackOf(item.source) === "健康").reduce((sum, item) => sum + threeSideTotal(item), 0));
    const overseasWorkload = round(items.filter((item) => trackOf(item.source) === "海外研究").reduce((sum, item) => sum + threeSideTotal(item), 0));
    const experienceItems = items.filter((item) => item.category === "体验优化");
    const experienceWorkload = round(experienceItems.reduce((sum, item) => sum + threeSideTotal(item), 0));
    const exclusiveWorkload = round(items.filter((item) => item.category === "产品专属").reduce((sum, item) => sum + threeSideTotal(item), 0));
    const workloadBreakdown = sumWorkloadBreakdown(items);
    const bySide: Record<WorkloadSide, number> = {
      device: workloadBreakdown.device,
      app: workloadBreakdown.app,
      cloud: workloadBreakdown.cloud,
    };
    const experienceBreakdown = sumWorkloadBreakdown(experienceItems);
    const experienceBySide: Record<WorkloadSide, number> = {
      device: experienceBreakdown.device,
      app: experienceBreakdown.app,
      cloud: experienceBreakdown.cloud,
    };
    return {
      halfYear,
      bySource,
      sportsWorkload,
      healthWorkload,
      overseasWorkload,
      experienceWorkload,
      exclusiveWorkload,
      totalWorkload: round(workloadBreakdown.device + workloadBreakdown.app + workloadBreakdown.cloud),
      deviceWorkload: workloadBreakdown.device,
      appWorkload: workloadBreakdown.app,
      cloudWorkload: workloadBreakdown.cloud,
      unallocatedWorkload: workloadBreakdown.unallocated,
      bySide,
      experienceBySide,
      requirementCount: items.length,
    };
  });
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
