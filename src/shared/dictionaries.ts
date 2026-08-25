import type { DictionaryItem } from "./types.js";

export const ALL_DICTIONARY_PARENTS = "all";

export function filterDictionaryItemsByParent(
  items: DictionaryItem[],
  parentFilter: string,
): DictionaryItem[] {
  if (!parentFilter || parentFilter === ALL_DICTIONARY_PARENTS) return items;
  return items.filter((item) => item.parentId === parentFilter);
}
