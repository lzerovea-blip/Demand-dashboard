import { describe, expect, it } from "vitest";
import { restoreTieringSeriesSelection } from "../src/renderer/TieringPage";
import { BUILT_IN_TIERING_PRODUCT_SERIES } from "../src/shared/tiering";

const allIds = BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => item.id);
const previousIds = allIds.filter((id) => id !== "PS14" && id !== "PS15");

describe("分层分级可见系列选择", () => {
  it("旧会话曾全选十三项时默认显示新增两项", () => {
    expect(restoreTieringSeriesSelection(allIds, previousIds)).toEqual(allIds);
  });

  it("旧会话自定义选择及清空选择保持不变", () => {
    expect(restoreTieringSeriesSelection(allIds, ["PS01", "PS07"])).toEqual(["PS01", "PS07"]);
    expect(restoreTieringSeriesSelection(allIds, [])).toEqual([]);
  });

  it("新会话明确取消 Pro 系列后不会在重开页面时重新勾选", () => {
    expect(restoreTieringSeriesSelection(allIds, { ids: previousIds, knownIds: allIds })).toEqual(previousIds);
  });
});
