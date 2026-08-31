import { describe, expect, it } from "vitest";
import { POWERPOINT_IRM_SCRIPT, checkPowerPointIrmEnvironment } from "../src/main/pptIrm";
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
});
