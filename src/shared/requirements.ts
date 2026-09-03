import { REQUIREMENT_SCOPES } from "./types.js";
import type { GroupOverride, Requirement, RequirementCategory, RequirementScope, RequirementSource } from "./types.js";
import { halfYearOf, hmosVersionForHalfYear } from "./roadmap.js";
import { historicalLaunchForFeatureId, normalizeHistoricalLaunchMetadata } from "./tieringLaunches.js";
import { roundWorkload } from "./workload.js";

export const LEGACY_OVERSEAS_SOURCE = "海外研究" as const;
export const LEGACY_AI_SOURCE = "ai" as const;
export const LEGACY_UPPERCASE_AI_SOURCE = "AI" as const;
export const LEGACY_MEDICAL_SUBMISSION_SOURCE = "医疗送检" as const;

type LegacyRequirement = Omit<Requirement, "source" | "scope" | "allowLatestVersionUpgrade" | "description" | "images" | "overseasRegions" | "domainL0Id" | "deviceWorkloadPm" | "appWorkloadPm" | "cloudWorkloadPm" | "unallocatedWorkloadPm"> &
  { source: RequirementSource | typeof LEGACY_OVERSEAS_SOURCE | typeof LEGACY_AI_SOURCE | typeof LEGACY_UPPERCASE_AI_SOURCE | typeof LEGACY_MEDICAL_SUBMISSION_SOURCE } &
  Partial<Pick<Requirement, "scope" | "allowLatestVersionUpgrade" | "description" | "images" | "overseasRegions" | "domainL0Id" | "deviceWorkloadPm" | "appWorkloadPm" | "cloudWorkloadPm" | "unallocatedWorkloadPm">>;

export function normalizeRequirement(input: LegacyRequirement): Requirement {
  const hasBreakdown = ["deviceWorkloadPm", "appWorkloadPm", "cloudWorkloadPm", "unallocatedWorkloadPm"]
    .some((key) => Object.prototype.hasOwnProperty.call(input, key));
  const deviceWorkloadPm = nonNegative(input.deviceWorkloadPm);
  const appWorkloadPm = nonNegative(input.appWorkloadPm);
  const cloudWorkloadPm = nonNegative(input.cloudWorkloadPm);
  const categorized = roundWorkload(deviceWorkloadPm + appWorkloadPm + cloudWorkloadPm);
  const legacyTotal = nonNegative(input.workloadPm);
  const source = normalizeLegacySource(input.source) as RequirementSource;
  const productIds = normalizeRequirementProductIds(input.category, input.productIds);
  let historicalLaunch = normalizeHistoricalLaunchMetadata(input.historicalLaunch);
  if (!historicalLaunch && input.tieringFeatureId && input.targetMonth === "2026-09") {
    try {
      historicalLaunch = historicalLaunchForFeatureId(input.tieringFeatureId);
    } catch {
      // 非标准扩展功能保持旧数据，不阻断普通工作区读取。
    }
  }
  const unallocatedWorkloadPm = hasBreakdown
    ? (Number.isFinite(input.unallocatedWorkloadPm) ? nonNegative(input.unallocatedWorkloadPm) : roundWorkload(Math.max(0, legacyTotal - categorized)))
    : legacyTotal;
  return {
    ...input,
    source,
    scope: normalizeRequirementScope(input.scope),
    allowLatestVersionUpgrade: input.allowLatestVersionUpgrade === true,
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
    targetMonth: historicalLaunch?.suggestedTargetMonth ?? input.targetMonth,
    historicalLaunch,
  };
}

export function normalizeRequirementProductIds(category: RequirementCategory, productIds: unknown): string[] {
  if (category === "体验优化") return [];
  return Array.isArray(productIds)
    ? [...new Set(productIds.filter((id): id is string => typeof id === "string" && Boolean(id)))].slice(0, 1)
    : [];
}

export type RoadmapRegion = "domestic" | "overseas";

export function normalizeRequirementScope(scope: unknown): RequirementScope {
  return REQUIREMENT_SCOPES.includes(scope as RequirementScope) ? scope as RequirementScope : "全球";
}

export function requirementVisibleInRegion(
  requirement: Pick<Requirement, "scope">,
  region: RoadmapRegion,
): boolean {
  return requirement.scope === "全球" || requirement.scope === (region === "domestic" ? "仅国内" : "仅海外");
}

export function filterRequirementsForRegion(requirements: Requirement[], region: RoadmapRegion): Requirement[] {
  return requirements.filter((item) => requirementVisibleInRegion(item, region));
}

export function normalizeLegacyGroupOverride(item: GroupOverride): GroupOverride {
  const [domainId, source, targetMonth] = item.groupKey.split("::");
  if (!domainId || !source || !targetMonth) return item;
  const normalizedSource = normalizeLegacySource(source);
  return normalizedSource === source ? item : { ...item, groupKey: `${domainId}::${normalizedSource}::${targetMonth}` };
}

export function normalizeLegacySource(source: string): string {
  if (source === LEGACY_OVERSEAS_SOURCE) return "行业";
  if (source === LEGACY_AI_SOURCE || source === LEGACY_UPPERCASE_AI_SOURCE) return "AI综合";
  if (source === LEGACY_MEDICAL_SUBMISSION_SOURCE) return "NMPA送检";
  return source;
}

function nonNegative(value: unknown): number {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 ? roundWorkload(value) : 0;
}

export interface RequirementFilters {
  query: string;
  domainL0Id: string;
  domainL1Id: string;
  targetHalfYear: string;
}

export interface ReleaseVersionOption {
  halfYear: string;
  version: string;
}

export function listReleaseVersionOptions(requirements: Requirement[]): ReleaseVersionOption[] {
  const halfYears = new Set<string>();
  for (const requirement of requirements) {
    try {
      halfYears.add(halfYearOf(requirement.targetMonth));
    } catch {
      // Invalid historical months remain excluded from filter options and validation handles them elsewhere.
    }
  }
  return [...halfYears].sort((a, b) => Number(b.slice(0, 4)) - Number(a.slice(0, 4)) || Number(b.slice(-1)) - Number(a.slice(-1)))
    .map((halfYear) => ({ halfYear, version: hmosVersionForHalfYear(halfYear) }));
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
      && (!filters.domainL0Id || item.domainL0Id === filters.domainL0Id)
      && (!filters.domainL1Id || item.domainId === filters.domainL1Id)
      && (!filters.targetHalfYear || halfYearOf(item.targetMonth) === filters.targetHalfYear);
  });
}

const COPY_TITLE_PATTERN = /^(.*?)\s*副本（([1-9]\d*)）$/u;

export function requirementCopyBaseTitle(title: string): string {
  return COPY_TITLE_PATTERN.exec(title.trim())?.[1]?.trim() || title.trim();
}

export function nextRequirementCopyTitle(title: string, existingTitles: Iterable<string>): string {
  const base = requirementCopyBaseTitle(title);
  const used = new Set<number>();
  for (const candidate of existingTitles) {
    const match = COPY_TITLE_PATTERN.exec(candidate.trim());
    if (match?.[1]?.trim() === base) used.add(Number(match[2]));
  }
  let index = 1;
  while (used.has(index)) index += 1;
  return `${base} 副本（${index}）`;
}

export function cloneRequirementAsCopy(
  source: Requirement,
  existingTitles: Iterable<string>,
  createId: () => string = () => crypto.randomUUID(),
  now = new Date().toISOString(),
): Requirement {
  return {
    ...source,
    id: createId(),
    title: nextRequirementCopyTitle(source.title, existingTitles),
    images: source.images.map((image) => ({ ...image, id: createId() })),
    productIds: [...source.productIds],
    overseasRegions: [...source.overseasRegions],
    createdAt: now,
    updatedAt: now,
    tieringFeatureId: undefined,
    historicalLaunch: undefined,
  };
}
