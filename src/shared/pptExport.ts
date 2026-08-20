import {
  buildHalfYearSummaries,
  buildRoadmapGroups,
  halfYearSequence,
  monthsForHalfYear,
  requirementsInRange,
  type HalfYearSummary,
  type RoadmapGroup,
} from "./roadmap.js";
import type { AppSnapshot, Requirement, Track } from "./types.js";

export interface PptRoadmapPagePlan {
  halfYear: string;
  track: Track;
  months: string[];
  groups: RoadmapGroup[];
  occupiedMonths: string[];
}

export interface PptDetailPagePlan {
  group: RoadmapGroup;
  requirements: Requirement[];
  part: number;
  partCount: number;
  slideNumber: number;
}

export interface PptExportPlan {
  start: string;
  end: string;
  halfYears: string[];
  requirements: Requirement[];
  summaries: HalfYearSummary[];
  groups: RoadmapGroup[];
  roadmapPages: PptRoadmapPagePlan[];
  detailPages: PptDetailPagePlan[];
  firstDetailSlideByGroup: Record<string, number>;
  slideCount: number;
}

const FIXED_OPENING_SLIDES = 3;
const REQUIREMENTS_PER_DETAIL_PAGE = 2;

export function buildPptExportPlan(snapshot: AppSnapshot, start: string, end: string): PptExportPlan {
  const halfYears = halfYearSequence(start, end);
  const requirements = requirementsInRange(snapshot.requirements, start, end);
  const summaries = buildHalfYearSummaries(snapshot.requirements, start, end);
  const groups = buildRoadmapGroups(snapshot, start, end);
  const roadmapPages = halfYears.flatMap((halfYear) => {
    const months = monthsForHalfYear(halfYear);
    const halfGroups = groups.filter((group) => months.includes(group.targetMonth));
    const occupiedMonths = [...new Set(halfGroups.map((group) => group.targetMonth))];
    return (["运动", "健康"] as Track[]).map((track) => ({
      halfYear,
      track,
      months,
      groups: halfGroups.filter((group) => group.track === track),
      occupiedMonths,
    }));
  });

  const detailStartSlide = FIXED_OPENING_SLIDES + roadmapPages.length + 1;
  const detailPages: PptDetailPagePlan[] = [];
  const firstDetailSlideByGroup: Record<string, number> = {};

  for (const group of groups) {
    const chunks = chunk(group.requirements, REQUIREMENTS_PER_DETAIL_PAGE);
    const partCount = Math.max(1, chunks.length);
    const firstSlide = detailStartSlide + detailPages.length;
    firstDetailSlideByGroup[group.key] = firstSlide;
    (chunks.length ? chunks : [[]]).forEach((items, index) => {
      detailPages.push({
        group,
        requirements: items,
        part: index + 1,
        partCount,
        slideNumber: detailStartSlide + detailPages.length,
      });
    });
  }

  return {
    start,
    end,
    halfYears,
    requirements,
    summaries,
    groups,
    roadmapPages,
    detailPages,
    firstDetailSlideByGroup,
    slideCount: FIXED_OPENING_SLIDES + roadmapPages.length + detailPages.length,
  };
}

export function dynamicRoadmapMonthWidths(
  months: string[],
  occupiedMonths: string[],
  totalWidth: number,
  emptyWidth = 0.62,
): number[] {
  const occupied = new Set(occupiedMonths);
  const occupiedCount = months.filter((month) => occupied.has(month)).length;
  if (occupiedCount === 0) return months.map(() => totalWidth / months.length);
  const emptyCount = months.length - occupiedCount;
  const occupiedWidth = Math.max(emptyWidth, (totalWidth - emptyCount * emptyWidth) / occupiedCount);
  return months.map((month) => occupied.has(month) ? occupiedWidth : emptyWidth);
}

function chunk<T>(items: T[], size: number): T[][] {
  return Array.from({ length: Math.ceil(items.length / size) }, (_, index) => items.slice(index * size, (index + 1) * size));
}
