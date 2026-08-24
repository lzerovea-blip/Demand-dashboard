import { describe, expect, it } from "vitest";
import { captureRoadmapMonthAnchor, roadmapCardIsL1Related, roadmapScrollLeftForAnchor } from "../src/renderer/roadmapInteraction";

describe("全屏相同 L1 联动", () => {
  it("只在启用时选中相同领域 L1", () => {
    expect(roadmapCardIsL1Related("sleep", "sleep", true)).toBe(true);
    expect(roadmapCardIsL1Related("heart", "sleep", true)).toBe(false);
    expect(roadmapCardIsL1Related("sleep", "sleep", false)).toBe(false);
    expect(roadmapCardIsL1Related("sleep", null, true)).toBe(false);
  });
});

describe("全屏路标月份位置保持", () => {
  it("记录冻结泳道右侧当前可见月份及相对位置", () => {
    const columns = [
      { month: "2027-01", offsetLeft: 116, width: 72 },
      { month: "2027-02", offsetLeft: 188, width: 280 },
      { month: "2027-03", offsetLeft: 468, width: 72 },
    ];
    expect(captureRoadmapMonthAnchor(100, 116, columns)).toEqual({ month: "2027-02", viewportOffset: 88 });
  });

  it("列宽变化后仍按同一月份恢复并限制在可滚动范围", () => {
    const anchor = { month: "2027-02", viewportOffset: 88 };
    const changedColumns = [
      { month: "2027-01", offsetLeft: 116, width: 280 },
      { month: "2027-02", offsetLeft: 396, width: 72 },
      { month: "2027-03", offsetLeft: 468, width: 72 },
    ];
    expect(roadmapScrollLeftForAnchor(anchor, changedColumns, 600)).toBe(308);
    expect(roadmapScrollLeftForAnchor(anchor, changedColumns, 250)).toBe(250);
  });

  it("没有月份列时不创建锚点", () => {
    expect(captureRoadmapMonthAnchor(0, 116, [])).toBeUndefined();
  });
});
