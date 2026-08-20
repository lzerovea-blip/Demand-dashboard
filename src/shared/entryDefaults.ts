import type { Requirement } from "./types.js";

export interface RequirementEntrySelections {
  source: Requirement["source"];
  category: Requirement["category"];
  targetMonth: string;
}

export function latestRequirementSelections(requirements: Requirement[]): RequirementEntrySelections | undefined {
  const latest = requirements.reduce<Requirement | undefined>((current, item) => {
    if (!current) return item;
    return item.updatedAt.localeCompare(current.updatedAt) > 0 ? item : current;
  }, undefined);

  return latest
    ? { source: latest.source, category: latest.category, targetMonth: latest.targetMonth }
    : undefined;
}
