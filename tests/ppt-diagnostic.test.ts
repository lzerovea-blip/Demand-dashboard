import { describe, expect, it } from "vitest";
import { compactPptDiagnosticValue, formatPptDiagnosticText } from "../src/shared/pptDiagnostic";

describe("PPT 简洁诊断日志", () => {
  it("固定为十二行纯文本并只保留最近六个步骤", () => {
    const content = formatPptDiagnosticText({
      createdAt: "2026-08-31T12:34:56.000Z",
      appVersion: "v0.7.40",
      stage: "重新打开受保护 PPT",
      powerPointAndIrmStatus: "PowerPoint 已检测；企业 IRM 可用",
      officeIdentity: "employee@china.example.com",
      preferredDocumentOwner: "owner@example.com",
      effectiveDocumentOwner: "employee@china.example.com",
      fallbackUsed: true,
      standardizationStatus: "已完成修复/标准化并由新 PowerPoint 实例常规重开",
      editableFileSize: "1000 字节",
      normalizedFileSize: "1100 字节",
      protectedFileSize: "1200 字节",
      completedSteps: ["1", "2", "3", "4", "5", "6", "7"],
      errorMessage: "PowerPoint could not open the file.",
    });
    expect(content.split("\n")).toHaveLength(12);
    expect(content).toContain("已完成步骤：2 → 3 → 4 → 5 → 6 → 7");
    expect(content).toContain("Office=employee@china.example.com；首选=owner@example.com；最终=employee@china.example.com；回退=是");
    expect(content).toContain("标准化：已完成修复/标准化并由新 PowerPoint 实例常规重开");
    expect(content).not.toContain("需求标题");
    expect(content).not.toContain("白名单账号列表");
  });

  it("把多行系统错误压缩成单行并限制长度", () => {
    expect(compactPptDiagnosticValue("第一行\r\n第二行\t第三行")).toBe("第一行 第二行 第三行");
    expect(compactPptDiagnosticValue("123456789", 6)).toBe("12345…");
  });
});
