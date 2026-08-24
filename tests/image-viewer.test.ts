import { describe, expect, it } from "vitest";
import { moveImageViewerIndex } from "../src/renderer/imageViewer";

describe("moveImageViewerIndex", () => {
  it("cycles forward and backward through multiple images", () => {
    expect(moveImageViewerIndex(0, 3, 1)).toBe(1);
    expect(moveImageViewerIndex(2, 3, 1)).toBe(0);
    expect(moveImageViewerIndex(0, 3, -1)).toBe(2);
  });

  it("does not move when there is only one image", () => {
    expect(moveImageViewerIndex(0, 1, 1)).toBe(0);
    expect(moveImageViewerIndex(0, 1, -1)).toBe(0);
  });

  it("normalizes an out-of-range current index before moving", () => {
    expect(moveImageViewerIndex(4, 3, 1)).toBe(2);
    expect(moveImageViewerIndex(-1, 3, -1)).toBe(1);
  });
});
