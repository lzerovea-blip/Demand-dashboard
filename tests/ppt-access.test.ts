import { describe, expect, it } from "vitest";
import {
  POWERPOINT_IRM_SCRIPT,
  checkPowerPointIrmEnvironment,
  encodeWindowsPowerShellScript,
  powerShellExecutableCandidates,
  verifyPowerPointIrmGrants,
} from "../src/main/pptIrm";
import { addCalendarMonthsClamped, formatPptExpirationDate, pptPermissionExpiration } from "../src/shared/pptAccess";

describe("PPT 企业 IRM 有效期", () => {
  it("按六个日历月计算并在月末自动收敛", () => {
    expect(formatPptExpirationDate(addCalendarMonthsClamped(new Date(2025, 7, 31, 9), 6))).toBe("2026-02-28");
    expect(formatPptExpirationDate(addCalendarMonthsClamped(new Date(2023, 7, 31, 9), 6))).toBe("2024-02-29");
    expect(formatPptExpirationDate(pptPermissionExpiration(new Date(2026, 9, 31, 9)))).toBe("2027-04-30");
  });

  it("非 Windows 环境明确拒绝无保护降级", async () => {
    const status = await checkPowerPointIrmEnvironment("darwin");
    expect(status).toMatchObject({ supported: false, powerPointInstalled: false, irmReady: false });
    expect(status.message).toContain("Windows");
  });

  it("PowerShell 流程包含清空旧权限、两种权限、设置到期和回读", () => {
    expect(POWERPOINT_IRM_SCRIPT).toContain("$permission.Enabled = $true");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$permission.RemoveAll() ".trim());
    expect(POWERPOINT_IRM_SCRIPT).toContain("$permission.Add");
    expect(POWERPOINT_IRM_SCRIPT).toContain("ExpirationDate");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$permission.Item($index)");
  });

  it("便携版优先检查 Windows 系统 PowerShell 并兼容 PowerShell 7 与 PATH", () => {
    expect(powerShellExecutableCandidates({ SystemRoot: "C:\\Windows", ProgramFiles: "C:\\Program Files" })).toEqual([
      "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe",
      "C:\\Windows\\Sysnative\\WindowsPowerShell\\v1.0\\powershell.exe",
      "C:\\Program Files\\PowerShell\\7\\pwsh.exe",
      "powershell.exe",
      "pwsh.exe",
    ]);
    expect(POWERPOINT_IRM_SCRIPT).toContain("未找到桌面版 Microsoft PowerPoint");
  });

  it("Windows PowerShell 5.1 脚本使用 UTF-8 BOM 保护中文提示", () => {
    const encoded = encodeWindowsPowerShellScript(POWERPOINT_IRM_SCRIPT);
    expect(encoded.charCodeAt(0)).toBe(0xFEFF);
    expect(encoded.slice(1)).toBe(POWERPOINT_IRM_SCRIPT);
    expect(encodeWindowsPowerShellScript(encoded)).toBe(encoded);
  });

  it("以实际启用 IRM 为准，不要求新文档预先存在 DocumentAuthor", () => {
    expect(POWERPOINT_IRM_SCRIPT).not.toContain("IsNullOrWhiteSpace($author)");
    const enabledAt = POWERPOINT_IRM_SCRIPT.indexOf("$permission.Enabled = $true");
    const authorAt = POWERPOINT_IRM_SCRIPT.indexOf("$result.signedInAccount = [string]$permission.DocumentAuthor");
    expect(enabledAt).toBeGreaterThan(-1);
    expect(authorAt).toBeGreaterThan(enabledAt);
  });

  it("权限回读允许 PowerPoint 自动保留额外的文档所有者", () => {
    const expirationUtc = "2027-02-28T00:00:00.000Z";
    expect(() => verifyPowerPointIrmGrants([
      { account: "owner@example.com", permission: 64, expirationUtc: "9999-12-31T23:59:59.000Z" },
      { account: "reader@example.com", permission: 1, expirationUtc },
    ], [
      { account: "reader@example.com", permission: 1, expirationUtc },
    ])).not.toThrow();
    expect(() => verifyPowerPointIrmGrants([], [
      { account: "reader@example.com", permission: 1, expirationUtc },
    ])).toThrow("缺少账号");
  });
});
