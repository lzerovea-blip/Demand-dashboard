import { describe, expect, it } from "vitest";
import { splitTextIntoLines } from "../src/shared/text";

describe("通用固定字符分行（供路标卡片使用）", () => {
  it("按20个Unicode字符分段且不自行截断内容", () => {
    for (const length of [20, 21, 40, 41]) {
      const text = "字".repeat(length);
      const lines = splitTextIntoLines(text);
      expect(lines.join("")).toBe(text);
      expect(lines.map((line) => Array.from(line).length)).toEqual(
        length === 20 ? [20] : length === 21 ? [20, 1] : length === 40 ? [20, 20] : [20, 20, 1],
      );
    }
  });

  it("保留中文、英文、标点、空格和完整Unicode码点", () => {
    expect(splitTextIntoLines("心率A1 →😀（测试）", 3)).toEqual(["心率A", "1 →", "😀（测", "试）"]);
  });

  it("空文本或非正的每行字符数不产生分段", () => {
    expect(splitTextIntoLines("")).toEqual([]);
    expect(splitTextIntoLines("心率", 0)).toEqual([]);
    expect(splitTextIntoLines("心率", -1)).toEqual([]);
  });
});
