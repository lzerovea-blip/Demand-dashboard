export interface PptDiagnosticTextInput {
  createdAt: string;
  appVersion: string;
  stage: string;
  powerPointStatus: string;
  irmStatus: string;
  signedInAccount: string;
  editableFileSize: string;
  protectedFileSize: string;
  completedSteps: string[];
  errorMessage: string;
}

export function compactPptDiagnosticValue(value: unknown, limit = 360): string {
  const normalized = String(value ?? "").replace(/\s+/g, " ").trim();
  return normalized.length > limit ? `${normalized.slice(0, limit - 1)}…` : normalized;
}

export function formatPptDiagnosticText(input: PptDiagnosticTextInput): string {
  const steps = input.completedSteps.slice(-6).map((item) => compactPptDiagnosticValue(item, 80));
  return [
    "解决方案需求管理 - PPT 导出诊断",
    `时间：${compactPptDiagnosticValue(input.createdAt, 80)}`,
    `应用版本：${compactPptDiagnosticValue(input.appVersion, 40)}`,
    `失败阶段：${compactPptDiagnosticValue(input.stage, 80)}`,
    `PowerPoint：${compactPptDiagnosticValue(input.powerPointStatus, 40)}`,
    `企业 IRM：${compactPptDiagnosticValue(input.irmStatus, 40)}`,
    `当前账号：${compactPptDiagnosticValue(input.signedInAccount, 120)}`,
    `原始 PPT：${compactPptDiagnosticValue(input.editableFileSize, 40)}`,
    `受保护 PPT：${compactPptDiagnosticValue(input.protectedFileSize, 40)}`,
    `已完成步骤：${steps.length ? steps.join(" → ") : "无"}`,
    `错误摘要：${compactPptDiagnosticValue(input.errorMessage)}`,
    "说明：日志不包含需求内容、图片或白名单账号。",
  ].join("\n");
}
