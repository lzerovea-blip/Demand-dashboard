import { describe, expect, it } from "vitest";
import type { WorkspaceWorkbookFieldChange } from "../src/shared/types";
import { adoptedWorkbookDeleteCount, allWorkbookChangeIds } from "../src/shared/workbookSelection";

function change(requirementId: string, kind: WorkspaceWorkbookFieldChange["kind"], field: string): WorkspaceWorkbookFieldChange {
  return {
    id: `${requirementId}:${kind}:${field}`, requirementId, row: 2, title: requirementId, kind, field,
    baselineValue: "", currentValue: "", excelValue: "", conflict: false,
  };
}

describe("协作 Excel 差异全选", () => {
  it("包含删除，但同一需求的删除优先于字段和系列修改", () => {
    const changes = [change("r1", "field", "title"), change("r1", "series", "PS01"), change("r1", "delete", "delete"), change("r2", "field", "title")];
    const selected = allWorkbookChangeIds(changes);
    expect(selected).toEqual(["r1:delete:delete", "r2:field:title"]);
    expect(adoptedWorkbookDeleteCount(changes, selected, {})).toBe(1);
    expect(adoptedWorkbookDeleteCount(changes, [], {})).toBe(0);
  });

  it("空预览不产生勾选项，正常字段和冲突项目均保留可选", () => {
    expect(allWorkbookChangeIds([])).toEqual([]);
    const conflict = { ...change("r3", "field", "scope"), conflict: true };
    expect(allWorkbookChangeIds([conflict])).toEqual([conflict.id]);
    const deleteConflict = { ...change("r3", "delete", "delete"), conflict: true };
    expect(adoptedWorkbookDeleteCount([deleteConflict], [deleteConflict.id], { [deleteConflict.id]: "local" })).toBe(0);
    expect(adoptedWorkbookDeleteCount([deleteConflict], [deleteConflict.id], { [deleteConflict.id]: "excel" })).toBe(1);
  });
});
