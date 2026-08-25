import { describe, expect, it } from "vitest";
import { roadmapCategoryMarks } from "../src/shared/roadmapCategory";

describe("路标分类标识", () => {
  it("平台演进使用红色演进标识", () => {
    expect(roadmapCategoryMarks(["平台演进"])).toEqual([
      { category: "平台演进", label: "演进", className: "platform", color: "#c63f3f" },
    ]);
  });

  it("合并卡片保留全部分类并使用固定顺序", () => {
    expect(roadmapCategoryMarks(["产品专属", "体验优化", "平台演进", "产品专属"])
      .map((item) => item.category))
      .toEqual(["体验优化", "平台演进", "产品专属"]);
  });
});
