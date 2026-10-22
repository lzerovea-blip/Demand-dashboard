import {
  buildHalfYearSummaries,
  buildRoadmapGroups,
  halfYearOf,
  halfYearSequence,
  requirementsInRange,
  medicalRoadmapTimelineMonths,
  type HalfYearSummary,
  type RoadmapGroup,
} from "./roadmap.js";
import { ROADMAP_TRACKS } from "./types.js";
import type { AppSnapshot, PptExportRequest, PptExportSectionId, Requirement, Track } from "./types.js";
import { roadmapDefinitionsFor, roadmapLanesFor } from "./roadmapManagement.js";
import type { RoadmapLaneDefinition } from "./types.js";
import { filterRequirementsForRegion, type RoadmapRegion } from "./requirements.js";

export const MAX_PPT_HALF_YEARS = 4;

export const PPT_EXPORT_SECTIONS: ReadonlyArray<{ id: PptExportSectionId; label: string }> = [
  { id: "summary", label: "需求工作量汇总分析" },
  ...ROADMAP_TRACKS.map((track) => ({ id: track, label: `${track}路标` as const })),
];

export const DEFAULT_PPT_EXPORT_SECTIONS = PPT_EXPORT_SECTIONS.map((item) => item.id);
export function pptBaseSlideCount(sections: readonly string[]): number {
  return 1 + sections.reduce((total, section) => total + (section === "summary" ? 1 : 2), 0);
}
export function pptExportSections(snapshot: Partial<AppSnapshot>): Array<{ id: string; label: string }> {
  return [{ id: "summary", label: "需求工作量汇总分析" }, ...roadmapDefinitionsFor(snapshot).map((item) => ({ id: item.id, label: item.name }))];
}

export interface PptRoadmapPagePlan {
  name?: string;
  kind?: "standard" | "medical";
  lanes?: RoadmapLaneDefinition[];
  track: Track;
  region?: RoadmapRegion;
  halfYears: string[];
  groups: RoadmapGroup[];
  medicalMonths?: string[];
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

export function normalizePptExportRequest(input: PptExportRequest, snapshot?: Partial<AppSnapshot>): PptExportRequest {
  const sections = snapshot ? pptExportSections(snapshot) : PPT_EXPORT_SECTIONS;
  const valid = new Set(sections.map((item) => item.id));
  const selected = [...new Set(input.sections ?? sections.map((item) => item.id))]
    .filter((section): section is PptExportSectionId => valid.has(section));
  return {
    start: input.start,
    end: input.end,
    materialName: input.materialName?.trim() || "解决方案需求路标",
    sections: sections.map((item) => item.id).filter((id) => selected.includes(id)),
  };
}

export function buildPptExportPlan(snapshot: AppSnapshot, startOrRequest: string | PptExportRequest, end?: string): PptExportPlan {
  const request = normalizePptExportRequest(typeof startOrRequest === "string"
    ? { start: startOrRequest, end: end ?? startOrRequest, materialName: "解决方案需求路标", sections: pptExportSections(snapshot).map((item) => item.id) }
    : startOrRequest, snapshot);
  const halfYears = halfYearSequence(request.start, request.end);
  if (halfYears.length > MAX_PPT_HALF_YEARS) {
    throw new Error(`PPT 一次最多支持 ${MAX_PPT_HALF_YEARS} 个连续半年，请缩短导出区间后重试`);
  }
  const requirements = requirementsInRange(snapshot.requirements, request.start, request.end);
  const summaries = buildHalfYearSummaries(snapshot.requirements, request.start, request.end, snapshot);
  const groups = buildRoadmapGroups(snapshot, request.start, request.end);
  const regions = (["domestic", "overseas"] as const).map((region) => {
    const regionalSnapshot = { ...snapshot, requirements: filterRequirementsForRegion(snapshot.requirements, region) };
    return { region,
      groups: buildRoadmapGroups(regionalSnapshot, request.start, request.end),
      medicalGroups: buildRoadmapGroups(regionalSnapshot, request.start, request.end, { medicalDemandRows: true }),
    };
  });
  const roadmapPages: PptRoadmapPagePlan[] = roadmapDefinitionsFor(snapshot).filter((definition) => request.sections.includes(definition.id)).flatMap((definition) => regions.map((regional) => {
    const pageGroups = (definition.kind === "medical" ? regional.medicalGroups : regional.groups).filter((group) => group.track === definition.id);
    return {
      track: definition.id, name: definition.name, kind: definition.kind, region: regional.region,
      ...(snapshot.roadmapAssignments !== undefined ? { lanes: roadmapLanesFor(snapshot, definition.id) } : {}),
      halfYears, groups: pageGroups,
      ...(definition.kind === "medical" ? { medicalMonths: medicalRoadmapTimelineMonths(pageGroups) } : {}),
    };
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
    slideCount: pptBaseSlideCount(request.sections),
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
