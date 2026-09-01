export interface PptDiagnosticTextInput {
  createdAt: string;
  appVersion: string;
  stage: string;
  powerPointAndIrmStatus: string;
  officeIdentity: string;
  preferredDocumentOwner: string;
  effectiveDocumentOwner: string;
  fallbackUsed: boolean;
  standardizationStatus: string;
  editableFileSize: string;
  normalizedFileSize: string;
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
    `PowerPoint / IRM：${compactPptDiagnosticValue(input.powerPointAndIrmStatus, 80)}`,
    `身份：Office=${compactPptDiagnosticValue(input.officeIdentity, 80)}；首选=${compactPptDiagnosticValue(input.preferredDocumentOwner, 80)}；最终=${compactPptDiagnosticValue(input.effectiveDocumentOwner, 80)}；回退=${input.fallbackUsed ? "是" : "否"}`,
    `标准化：${compactPptDiagnosticValue(input.standardizationStatus, 100)}`,
    `文件：原始 ${compactPptDiagnosticValue(input.editableFileSize, 40)}；标准化 ${compactPptDiagnosticValue(input.normalizedFileSize, 40)}；受保护 ${compactPptDiagnosticValue(input.protectedFileSize, 40)}`,
    `已完成步骤：${steps.length ? steps.join(" → ") : "无"}`,
    `错误摘要：${compactPptDiagnosticValue(input.errorMessage)}`,
    "诊断文件：已自动保留在本机，仅保留最近 3 次失败记录。",
    "说明：诊断 PPT 包含业务数据；文字日志不包含需求内容、图片或白名单账号。",
  ].join("\n");
}
