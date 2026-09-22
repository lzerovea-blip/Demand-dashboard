import type { MedicalSubmissionMilestones, RequirementSource } from "./types.js";

export const MEDICAL_SUBMISSION_SOURCES = ["NMPA认证", "CE认证"] as const;

export const MEDICAL_SUBMISSION_MILESTONE_FIELDS = [
  { key: "registrationStartMonth", label: "注册启动时间" },
  { key: "clinicalStartMonth", label: "临床启动时间" },
  { key: "clinicalCompleteMonth", label: "临床完成时间" },
  { key: "certificateMonth", label: "获证时间" },
] as const;

export type MedicalSubmissionMilestoneKey = (typeof MEDICAL_SUBMISSION_MILESTONE_FIELDS)[number]["key"];

export interface MedicalSubmissionTimelineNode {
  key: MedicalSubmissionMilestoneKey | "submissionMonth";
  label: string;
  month?: string;
  derived: boolean;
}

export function isMedicalSubmissionSource(source: string): source is (typeof MEDICAL_SUBMISSION_SOURCES)[number] {
  return MEDICAL_SUBMISSION_SOURCES.includes(source as (typeof MEDICAL_SUBMISSION_SOURCES)[number]);
}

export function isValidRequirementMonth(value: unknown): value is string {
  return typeof value === "string" && /^\d{4}-(0[1-9]|1[0-2])$/.test(value);
}

export function validateRequirementTargetMonth(source: string, targetMonth: string): void {
  if (isMedicalSubmissionSource(source) && targetMonth === "") return;
  if (!isValidRequirementMonth(targetMonth)) {
    throw new Error(isMedicalSubmissionSource(source) ? "送检版本应为有效年月或留空" : "上线年月格式应为 YYYY-MM");
  }
}

export function initialRequirementTargetMonth(source: string, previousMonth: string | undefined, currentMonth: string): string {
  return isMedicalSubmissionSource(source) ? "" : previousMonth || currentMonth;
}

export function normalizeMedicalSubmissionMilestones(
  source: RequirementSource,
  input: MedicalSubmissionMilestones | null | undefined,
): MedicalSubmissionMilestones | undefined {
  if (!isMedicalSubmissionSource(source) || !input || typeof input !== "object") return undefined;
  const normalized: MedicalSubmissionMilestones = {};
  for (const { key } of MEDICAL_SUBMISSION_MILESTONE_FIELDS) {
    const value = typeof input[key] === "string" ? input[key]?.trim() : "";
    if (value) normalized[key] = value;
  }
  return Object.keys(normalized).length ? normalized : undefined;
}

export function medicalSubmissionTimeline(
  targetMonth: string,
  milestones?: MedicalSubmissionMilestones,
): MedicalSubmissionTimelineNode[] {
  return [
    { key: "registrationStartMonth", label: "注册启动时间", month: milestones?.registrationStartMonth, derived: false },
    { key: "submissionMonth", label: "送检时间", month: targetMonth || undefined, derived: true },
    { key: "clinicalStartMonth", label: "临床启动时间", month: milestones?.clinicalStartMonth, derived: false },
    { key: "clinicalCompleteMonth", label: "临床完成时间", month: milestones?.clinicalCompleteMonth, derived: false },
    { key: "certificateMonth", label: "获证时间", month: milestones?.certificateMonth, derived: false },
  ];
}

export function validateMedicalSubmissionMilestones(
  source: RequirementSource,
  targetMonth: string,
  input?: MedicalSubmissionMilestones,
): MedicalSubmissionMilestones | undefined {
  const normalized = normalizeMedicalSubmissionMilestones(source, input);
  if (!isMedicalSubmissionSource(source)) return undefined;
  validateRequirementTargetMonth(source, targetMonth);
  for (const node of medicalSubmissionTimeline(targetMonth, normalized)) {
    if (node.month && !isValidRequirementMonth(node.month)) throw new Error(`${node.label}必须是有效月份`);
  }
  return normalized;
}

export function formatMedicalSubmissionMonth(month?: string): string {
  if (!month) return "未填写";
  const [year, rawMonth] = month.split("-");
  return `${year}年${Number(rawMonth)}月`;
}
