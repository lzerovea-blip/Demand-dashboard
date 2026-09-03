import type { DictionaryItem } from "./types.js";

export const ALL_DICTIONARY_PARENTS = "all";

/** Resolve both dependent selectors together, including deleted/reparented dictionaries. */
export function domainFilterSelection(domains: readonly DictionaryItem[], l0Id: string, l1Id = "") {
  const l0Options = domains.filter((item) => item.level === "L0");
  const domainL0Id = l0Options.some((item) => item.id === l0Id) ? l0Id : "";
  const l1Options = domains.filter((item) => (item.level ?? "L1") === "L1" && (!domainL0Id || item.parentId === domainL0Id));
  const domainL1Id = l1Options.some((item) => item.id === l1Id) ? l1Id : "";
  return { domainL0Id, domainL1Id, l0Options, l1Options };
}

export function toggleDomainSelection(selected: readonly string[], id: string): string[] {
  return selected.includes(id) ? selected.filter((item) => item !== id) : [...selected, id];
}

export function filterDictionaryItemsByParent(
  items: DictionaryItem[],
  parentFilter: string,
): DictionaryItem[] {
  if (!parentFilter || parentFilter === ALL_DICTIONARY_PARENTS) return items;
  return items.filter((item) => item.parentId === parentFilter);
}
