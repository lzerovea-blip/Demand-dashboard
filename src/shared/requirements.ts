import type { GroupOverride, Requirement, RequirementCategory, RequirementSource } from "./types.js";
import { roundWorkload } from "./workload.js";

export const LEGACY_OVERSEAS_SOURCE = "海外研究" as const;
export const LEGACY_AI_SOURCE = "ai" as const;

type LegacyRequirement = Omit<Requirement, "source" | "description" | "images" | "overseasRegions" | "domainL0Id" | "deviceWorkloadPm" | "appWorkloadPm" | "cloudWorkloadPm" | "unallocatedWorkloadPm"> &
  { source: RequirementSource | typeof LEGACY_OVERSEAS_SOURCE | typeof LEGACY_AI_SOURCE } &
  Partial<Pick<Requirement, "description" | "images" | "overseasRegions" | "domainL0Id" | "deviceWorkloadPm" | "appWorkloadPm" | "cloudWorkloadPm" | "unallocatedWorkloadPm">>;

export function normalizeRequirement(input: LegacyRequirement): Requirement {
  const hasBreakdown = ["deviceWorkloadPm", "appWorkloadPm", "cloudWorkloadPm", "unallocatedWorkloadPm"]
    .some((key) => Object.prototype.hasOwnProperty.call(input, key));
  const deviceWorkloadPm = nonNegative(input.deviceWorkloadPm);
  const appWorkloadPm = nonNegative(input.appWorkloadPm);
  const cloudWorkloadPm = nonNegative(input.cloudWorkloadPm);
  const categorized = roundWorkload(deviceWorkloadPm + appWorkloadPm + cloudWorkloadPm);
  const legacyTotal = nonNegative(input.workloadPm);
  const source = normalizeLegacySource(input.source) as RequirementSource;
  const productIds = input.category === "产品专属" && Array.isArray(input.productIds)
    ? [...new Set(input.productIds.filter((id): id is string => typeof id === "string" && Boolean(id)))].slice(0, 1)
    : [];
  const unallocatedWorkloadPm = hasBreakdown
    ? (Number.isFinite(input.unallocatedWorkloadPm) ? nonNegative(input.unallocatedWorkloadPm) : roundWorkload(Math.max(0, legacyTotal - categorized)))
    : legacyTotal;
  return {
    ...input,
    source,
    description: typeof input.description === "string" ? input.description : "",
    images: Array.isArray(input.images) ? input.images : [],
    domainL0Id: typeof input.domainL0Id === "string" ? input.domainL0Id : "",
    productIds,
    overseasRegions: [],
    deviceWorkloadPm,
    appWorkloadPm,
    cloudWorkloadPm,
    unallocatedWorkloadPm,
    workloadPm: roundWorkload(categorized + unallocatedWorkloadPm),
  };
}

export function normalizeLegacyGroupOverride(item: GroupOverride): GroupOverride {
  const [domainId, source, targetMonth] = item.groupKey.split("::");
  if (!domainId || !source || !targetMonth) return item;
  const normalizedSource = normalizeLegacySource(source);
  return normalizedSource === source ? item : { ...item, groupKey: `${domainId}::${normalizedSource}::${targetMonth}` };
}

export function normalizeLegacySource(source: string): string {
  if (source === LEGACY_OVERSEAS_SOURCE) return "行业";
  if (source === LEGACY_AI_SOURCE) return "AI";
  return source;
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
    const searchText = `${item.title} ${domainNames.get(item.domainL0Id) ?? ""} ${domainNames.get(item.domainId) ?? ""} ${item.overseasRegions.join(" ")}`.toLocaleLowerCase("zh-CN");
    return (!normalizedQuery || searchText.includes(normalizedQuery))
      && (!filters.source || item.source === filters.source)
      && (!filters.category || item.category === filters.category)
      && (!filters.targetMonth || item.targetMonth === filters.targetMonth);
  });
}
