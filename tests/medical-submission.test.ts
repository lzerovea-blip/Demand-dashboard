import { describe, expect, it } from "vitest";
import {
  formatMedicalSubmissionMonth,
  initialRequirementTargetMonth,
  medicalSubmissionTimeline,
  normalizeMedicalSubmissionMilestones,
  validateMedicalSubmissionMilestones,
  validateRequirementTargetMonth,
} from "../src/shared/medicalSubmission";

describe("医疗认证时间节点", () => {
  it("允许全部留空、部分填写和相邻节点同月", () => {
    expect(validateMedicalSubmissionMilestones("NMPA认证", "2027-03", undefined)).toBeUndefined();
    expect(validateMedicalSubmissionMilestones("CE认证", "2027-03", {
      registrationStartMonth: "2027-01",
      clinicalStartMonth: "2027-03",
      clinicalCompleteMonth: "2027-03",
      certificateMonth: "2027-09",
    })).toEqual({ registrationStartMonth: "2027-01", clinicalStartMonth: "2027-03", clinicalCompleteMonth: "2027-03", certificateMonth: "2027-09" });
  });

  it("允许任意倒序，仅拒绝非法月份，并将送检时间始终派生自送检版本", () => {
    expect(validateMedicalSubmissionMilestones("NMPA认证", "2027-03", { registrationStartMonth: "2027-04", clinicalStartMonth: "2025-01", clinicalCompleteMonth: "2024-10", certificateMonth: "2023-02" })).toEqual({ registrationStartMonth: "2027-04", clinicalStartMonth: "2025-01", clinicalCompleteMonth: "2024-10", certificateMonth: "2023-02" });
    expect(() => validateMedicalSubmissionMilestones("CE认证", "2027-03", { clinicalStartMonth: "2027-13" })).toThrow(/临床启动时间必须是有效月份/);
    expect(medicalSubmissionTimeline("2028-02", { certificateMonth: "2028-08" })).toMatchObject([
      { label: "注册启动时间", month: undefined },
      { label: "送检时间", month: "2028-02", derived: true },
      { label: "临床启动时间", month: undefined },
      { label: "临床完成时间", month: undefined },
      { label: "获证时间", month: "2028-08" },
    ]);
  });

  it("仅医疗来源允许空送检版本，五节点全空不生成伪造月份", () => {
    for (const source of ["NMPA认证", "CE认证"] as const) {
      expect(validateMedicalSubmissionMilestones(source, "", undefined)).toBeUndefined();
      expect(() => validateRequirementTargetMonth(source, "")).not.toThrow();
      expect(() => validateRequirementTargetMonth(source, "2027-13")).toThrow(/有效年月/);
      expect(validateMedicalSubmissionMilestones(source, "", { certificateMonth: "2025-01" })).toEqual({ certificateMonth: "2025-01" });
    }
    expect(medicalSubmissionTimeline("", undefined).every((node) => node.month === undefined)).toBe(true);
    expect(() => validateRequirementTargetMonth("健康基础", "")).toThrow(/上线年月/);
    expect(() => validateRequirementTargetMonth("健康基础", "2027-03")).not.toThrow();
  });

  it("新建医疗默认空；普通需求不继承上条医疗需求的空月份", () => {
    expect(initialRequirementTargetMonth("NMPA认证", "2027-03", "2026-09")).toBe("");
    expect(initialRequirementTargetMonth("CE认证", undefined, "2026-09")).toBe("");
    expect(initialRequirementTargetMonth("健康基础", "", "2026-09")).toBe("2026-09");
    expect(initialRequirementTargetMonth("健康基础", "2027-03", "2026-09")).toBe("2027-03");
  });

  it("非医疗来源保存时清理节点，月份格式化不伪造空值", () => {
    expect(normalizeMedicalSubmissionMilestones("健康基础", { registrationStartMonth: "2027-01" })).toBeUndefined();
    expect(formatMedicalSubmissionMonth("2027-03")).toBe("2027年3月");
    expect(formatMedicalSubmissionMonth()).toBe("未填写");
  });
});
