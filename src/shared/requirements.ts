import { OVERSEAS_REGIONS } from "./types.js";
import type { OverseasRegion, Requirement, RequirementCategory, RequirementSource } from "./types.js";
import { roundWorkload } from "./workload.js";

type LegacyRequirement = Omit<Requirement, "description" | "images" | "overseasRegions" | "deviceWorkloadPm" | "appWorkloadPm" | "cloudWorkloadPm" | "unallocatedWorkloadPm"> &
  Partial<Pick<Requirement, "description" | "images" | "overseasRegions" | "deviceWorkloadPm" | "appWorkloadPm" | "cloudWorkloadPm" | "unallocatedWorkloadPm">>;

export function normalizeRequirement(input: LegacyRequirement): Requirement {
  const hasBreakdown = ["deviceWorkloadPm", "appWorkloadPm", "cloudWorkloadPm", "unallocatedWorkloadPm"]
    .some((key) => Object.prototype.hasOwnProperty.call(input, key));
  const deviceWorkloadPm = nonNegative(input.deviceWorkloadPm);
  const appWorkloadPm = nonNegative(input.appWorkloadPm);
  const cloudWorkloadPm = nonNegative(input.cloudWorkloadPm);
  const categorized = roundWorkload(deviceWorkloadPm + appWorkloadPm + cloudWorkloadPm);
  const legacyTotal = nonNegative(input.workloadPm);
  const overseasRegions = input.source === "海外研究" && Array.isArray(input.overseasRegions)
    ? [...new Set(input.overseasRegions.filter((item): item is OverseasRegion => OVERSEAS_REGIONS.includes(item as OverseasRegion)))]
    : [];
  const unallocatedWorkloadPm = hasBreakdown
    ? (Number.isFinite(input.unallocatedWorkloadPm) ? nonNegative(input.unallocatedWorkloadPm) : roundWorkload(Math.max(0, legacyTotal - categorized)))
    : legacyTotal;
  return {
    ...input,
    description: typeof input.description === "string" ? input.description : "",
    images: Array.isArray(input.images) ? input.images : [],
    overseasRegions,
    deviceWorkloadPm,
    appWorkloadPm,
    cloudWorkloadPm,
    unallocatedWorkloadPm,
    workloadPm: roundWorkload(categorized + unallocatedWorkloadPm),
  };
}

function nonNegative(value: unknown): number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 ? roundWorkload(value) : 0;
}

export interface RequirementFilters {
  query: string;
  source: RequirementSource | "";
  category: RequirementCategory | "";
  targetMonth: string;
}

export function listTargetMonths(requirements: Requirement[]): string[] {
  return [...new Set(requirements.map((item) => item.targetMonth).filter((item) => /^\d{4}-(0[1-9]|1[0-2])$/.test(item)))].sort();
}

export function filterRequirements(
  requirements: Requirement[],
  domainNames: ReadonlyMap<string, string>,
  filters: RequirementFilters,
): Requirement[] {
  const normalizedQuery = filters.query.trim().toLocaleLowerCase("zh-CN");
  return requirements.filter((item) => {
    const searchText = `${item.title} ${domainNames.get(item.domainId) ?? ""} ${item.overseasRegions.join(" ")}`.toLocaleLowerCase("zh-CN");
    return (!normalizedQuery || searchText.includes(normalizedQuery))
      && (!filters.source || item.source === filters.source)
      && (!filters.category || item.category === filters.category)
      && (!filters.targetMonth || item.targetMonth === filters.targetMonth);
  });
}
