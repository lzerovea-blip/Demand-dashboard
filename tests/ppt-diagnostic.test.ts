import { describe, expect, it } from "vitest";
import { compactPptDiagnosticValue, formatPptDiagnosticText } from "../src/shared/pptDiagnostic";

describe("PPT 简洁诊断日志", () => {
  it("固定为十二行纯文本并只保留最近六个步骤", () => {
    const content = formatPptDiagnosticText({
      createdAt: "2026-08-31T12:34:56.000Z",
      appVersion: "v0.7.38",
      stage: "重新打开受保护 PPT",
      powerPointStatus: "已检测",
      irmStatus: "可用",
      documentOwnerAccount: "owner@example.com",
      editableFileSize: "1000 字节",
      protectedFileSize: "1200 字节",
      completedSteps: ["1", "2", "3", "4", "5", "6", "7"],
      errorMessage: "PowerPoint could not open the file.",
    });
    expect(content.split("\n")).toHaveLength(12);
    expect(content).toContain("已完成步骤：2 → 3 → 4 → 5 → 6 → 7");
    expect(content).toContain("文档所有者：owner@example.com");
    expect(content).not.toContain("需求标题");
    expect(content).not.toContain("白名单账号列表");
  });

  it("把多行系统错误压缩成单行并限制长度", () => {
    expect(compactPptDiagnosticValue("第一行\r\n第二行\t第三行")).toBe("第一行 第二行 第三行");
    expect(compactPptDiagnosticValue("123456789", 6)).toBe("12345…");
  });
});
