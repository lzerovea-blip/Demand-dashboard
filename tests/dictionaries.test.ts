import { describe, expect, it } from "vitest";
import { ALL_DICTIONARY_PARENTS, filterDictionaryItemsByParent } from "../src/shared/dictionaries";
import type { DictionaryItem } from "../src/shared/types";

const items: DictionaryItem[] = [
  { id: "sleep", name: "睡眠", sortOrder: 0, active: true, level: "L1", parentId: "health" },
  { id: "run", name: "跑步", sortOrder: 1, active: true, level: "L1", parentId: "sports" },
  { id: "heart", name: "心脏健康", sortOrder: 2, active: true, level: "L1", parentId: "health" },
];

describe("领域 L1 的 L0 标签筛选", () => {
  it("全部标签保留所有 L1", () => {
    expect(filterDictionaryItemsByParent(items, ALL_DICTIONARY_PARENTS)).toEqual(items);
  });

  it("选中 L0 后只返回其关联 L1", () => {
    expect(filterDictionaryItemsByParent(items, "health").map((item) => item.id)).toEqual(["sleep", "heart"]);
    expect(filterDictionaryItemsByParent(items, "missing")).toEqual([]);
  });
});
