import type { Requirement } from "./types.js";

export interface RequirementEntrySelections {
  domainL0Id: string;
  domainId: string;
  source: Requirement["source"];
  overseasRegions: Requirement["overseasRegions"];
  category: Requirement["category"];
  targetMonth: string;
}

export function latestRequirementSelections(requirements: Requirement[]): RequirementEntrySelections | undefined {
  const latest = requirements.reduce<Requirement | undefined>((current, item) => {
    if (!current) return item;
    return item.updatedAt.localeCompare(current.updatedAt) > 0 ? item : current;
  }, undefined);

  return latest
    ? { domainL0Id: latest.domainL0Id, domainId: latest.domainId, source: latest.source, overseasRegions: latest.overseasRegions, category: latest.category, targetMonth: latest.targetMonth }
    : undefined;
}
