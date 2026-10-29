import { describe, expect, it } from "vitest";
import { ALL_DICTIONARY_PARENTS, domainFilterSelection, filterDictionaryItemsByParent, toggleDomainSelection } from "../src/shared/dictionaries";
import type { DictionaryItem } from "../src/shared/types";

const items: DictionaryItem[] = [
  { id: "sleep", name: "睡眠", sortOrder: 0, active: true, level: "L1", parentId: "health" },
  { id: "run", name: "跑步", sortOrder: 1, active: true, level: "L1", parentId: "sports" },
  { id: "heart", name: "心脏健康", sortOrder: 2, active: true, level: "L1", parentId: "health" },
];

describe("领域 L1 的 L0 标签筛选", () => {
  const domains: DictionaryItem[] = [...items, ...["health", "sports", "empty"].map((id) => ({ id, name: id, level: "L0" as const, sortOrder: 0, active: true }))];
  it("L0 联动子级且切换后清空不匹配的 L1", () => {
    expect(domainFilterSelection(domains, "health", "sleep").domainL1Id).toBe("sleep");
    expect(domainFilterSelection(domains, "sports", "sleep")).toMatchObject({ domainL0Id: "sports", domainL1Id: "", l1Options: [items[1]] });
  });
  it("空子级和被删除选项复位", () => {
    expect(domainFilterSelection(domains, "empty", "sleep")).toMatchObject({ domainL0Id: "empty", domainL1Id: "", l1Options: [] });
    expect(domainFilterSelection(domains, "deleted", "deleted")).toMatchObject({ domainL0Id: "", domainL1Id: "", l1Options: items });
  });
  it("人员多选跨 L0 保留，仅取消指定项且不修改输入数组", () => {
    const selected = ["sleep"];
    const next = toggleDomainSelection(selected, "run");
    expect(next).toEqual(["sleep", "run"]);
    expect(toggleDomainSelection(next, "sleep")).toEqual(["run"]);
    expect(selected).toEqual(["sleep"]);
  });
  it("全部标签保留所有 L1", () => {
    expect(filterDictionaryItemsByParent(items, ALL_DICTIONARY_PARENTS)).toEqual(items);
  });

  it("选中 L0 后只返回其关联 L1", () => {
    expect(filterDictionaryItemsByParent(items, "health").map((item) => item.id)).toEqual(["sleep", "heart"]);
    expect(filterDictionaryItemsByParent(items, "missing")).toEqual([]);
  });
});
