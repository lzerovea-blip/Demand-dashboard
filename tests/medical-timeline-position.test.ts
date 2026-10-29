import { describe, expect, it } from "vitest";
import { medicalTimelinePopoverPosition } from "../src/renderer/medicalTimeline";

describe("医疗时间轴浮层定位", () => {
  const panel = { width: 500, height: 220 };
  it("下方空间足够时显示在卡片下方并限制左右边界", () => {
    expect(medicalTimelinePopoverPosition({ left: 900, right: 1100, top: 100, bottom: 160, width: 200, height: 60 }, panel, 1024, 768)).toEqual({ left: 512, top: 168, placement: "below" });
  });
  it("底部空间不足而顶部足够时翻转到上方", () => {
    expect(medicalTimelinePopoverPosition({ left: 100, right: 300, top: 650, bottom: 710, width: 200, height: 60 }, panel, 1024, 768)).toEqual({ left: 100, top: 422, placement: "above" });
  });
});
