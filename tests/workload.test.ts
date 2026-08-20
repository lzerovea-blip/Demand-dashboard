import { describe, expect, it } from "vitest";
import { normalizeWorkloadEdit, parseWorkloadPart, parseWorkloadPm, sumWorkloadBreakdown } from "../src/shared/workload";

describe("工作量输入", () => {
  it("同时接受正整数和小数", () => {
    expect(parseWorkloadPm("1")).toBe(1);
    expect(parseWorkloadPm("2")).toBe(2);
    expect(parseWorkloadPm("0.5")).toBe(0.5);
    expect(parseWorkloadPm("1.25")).toBe(1.25);
    expect(parseWorkloadPm(".5")).toBe(0.5);
  });

  it("支持中文输入法常见的小数分隔符", () => {
    expect(normalizeWorkloadEdit("1。5")).toBe("1.5");
    expect(normalizeWorkloadEdit("1，5")).toBe("1.5");
  });

  it("拒绝零、负数、不完整小数和非数字", () => {
    expect(parseWorkloadPm("0")).toBeNull();
    expect(parseWorkloadPm("-1")).toBeNull();
    expect(parseWorkloadPm("1.")).toBeNull();
    expect(parseWorkloadPm("abc")).toBeNull();
  });

  it("三侧明细允许单项为零并自动汇总", () => {
    expect(parseWorkloadPart("")).toBe(0);
    expect(parseWorkloadPart("0")).toBe(0);
    expect(parseWorkloadPart("0.5")).toBe(0.5);
    expect(parseWorkloadPart("-1")).toBeNull();
    expect(sumWorkloadBreakdown([{
      deviceWorkloadPm: 0.5,
      appWorkloadPm: 1,
      cloudWorkloadPm: 0.25,
      unallocatedWorkloadPm: 0,
    }])).toEqual({ device: 0.5, app: 1, cloud: 0.25, unallocated: 0, total: 1.75 });
  });
});
