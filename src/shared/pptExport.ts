import {
  buildHalfYearSummaries,
  buildRoadmapGroups,
  halfYearOf,
  halfYearSequence,
  requirementsInRange,
  type HalfYearSummary,
  type RoadmapGroup,
} from "./roadmap.js";
import { ROADMAP_TRACKS } from "./types.js";
import type { AppSnapshot, PptExportRequest, PptExportSectionId, Requirement, Track } from "./types.js";

export const MAX_PPT_HALF_YEARS = 4;

export const PPT_EXPORT_SECTIONS: ReadonlyArray<{ id: PptExportSectionId; label: string }> = [
  { id: "summary", label: "需求工作量汇总分析" },
  ...ROADMAP_TRACKS.map((track) => ({ id: track, label: `${track}路标` as const })),
];

export const DEFAULT_PPT_EXPORT_SECTIONS = PPT_EXPORT_SECTIONS.map((item) => item.id);

export interface PptRoadmapPagePlan {
  track: Track;
  halfYears: string[];
  groups: RoadmapGroup[];
}

export interface PptExportPlan {
  start: string;
  end: string;
  materialName: string;
  selectedSections: PptExportSectionId[];
  halfYears: string[];
  requirements: Requirement[];
  summaries: HalfYearSummary[];
  groups: RoadmapGroup[];
  roadmapPages: PptRoadmapPagePlan[];
  /** Base count before the main-process layout adds roadmap continuation pages. */
  slideCount: number;
}

export function normalizePptExportRequest(input: PptExportRequest): PptExportRequest {
  const valid = new Set(PPT_EXPORT_SECTIONS.map((item) => item.id));
  const selected = [...new Set(input.sections ?? DEFAULT_PPT_EXPORT_SECTIONS)]
    .filter((section): section is PptExportSectionId => valid.has(section));
  return {
    start: input.start,
    end: input.end,
    materialName: input.materialName?.trim() || "解决方案需求路标",
    sections: PPT_EXPORT_SECTIONS.map((item) => item.id).filter((id) => selected.includes(id)),
  };
}

export function buildPptExportPlan(snapshot: AppSnapshot, startOrRequest: string | PptExportRequest, end?: string): PptExportPlan {
  const request = normalizePptExportRequest(typeof startOrRequest === "string"
    ? { start: startOrRequest, end: end ?? startOrRequest, materialName: "解决方案需求路标", sections: DEFAULT_PPT_EXPORT_SECTIONS }
    : startOrRequest);
  const halfYears = halfYearSequence(request.start, request.end);
  if (halfYears.length > MAX_PPT_HALF_YEARS) {
    throw new Error(`PPT 一次最多支持 ${MAX_PPT_HALF_YEARS} 个连续半年，请缩短导出区间后重试`);
  }
  const requirements = requirementsInRange(snapshot.requirements, request.start, request.end);
  const summaries = buildHalfYearSummaries(snapshot.requirements, request.start, request.end);
  const groups = buildRoadmapGroups(snapshot, request.start, request.end);
  const roadmapPages = ROADMAP_TRACKS.filter((track) => request.sections.includes(track)).map((track) => ({
    track,
    halfYears,
    groups: groups.filter((group) => group.track === track),
  }));
  return {
    start: request.start,
    end: request.end,
    materialName: request.materialName,
    selectedSections: request.sections,
    halfYears,
    requirements,
    summaries,
    groups,
    roadmapPages,
    slideCount: 1 + request.sections.length,
  };
}

export function groupsForHalfYear(groups: RoadmapGroup[], halfYear: string): RoadmapGroup[] {
  return groups.filter((group) => halfYearOf(group.targetMonth) === halfYear);
}

export type RoadmapBadgeKind = "product" | "experience" | null;

export function roadmapBadgeKind(group: Pick<RoadmapGroup, "requirements">): RoadmapBadgeKind {
  if (group.requirements.some((item) => item.productIds.length > 0)) return "product";
  if (group.requirements.some((item) => item.category === "体验优化")) return "experience";
  return null;
}

export function visibleRequirementTitles(requirements: Requirement[], limit = 3): string[] {
  const visible = requirements.slice(0, limit).map((item) => item.title.trim()).filter(Boolean);
  if (requirements.length > limit && visible.length) visible[visible.length - 1] = `${visible[visible.length - 1]}…`;
  return visible;
}
