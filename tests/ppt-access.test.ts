import { describe, expect, it } from "vitest";
import {
  POWERPOINT_IRM_SCRIPT,
  buildPowerPointIrmAccess,
  checkPowerPointIrmEnvironment,
  encodeWindowsPowerShellScript,
  powerShellExecutableCandidates,
  shouldRetryWithOfficeIdentity,
  verifyPowerPointIrmGrants,
} from "../src/main/pptIrm";
import { addCalendarMonthsClamped, formatPptExpirationDate, pptPermissionExpiration } from "../src/shared/pptAccess";
import type { Member, PptAccessPolicy } from "../src/shared/types";

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
    expect(POWERPOINT_IRM_SCRIPT).toContain("[Microsoft.Office.Core.MsoPermission][int]$grant.permission");
    expect(POWERPOINT_IRM_SCRIPT).toContain("白名单第 $grantIndex 项权限添加失败");
    expect(POWERPOINT_IRM_SCRIPT).toContain("白名单第 $grantIndex 项有效期设置失败");
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

  it("探测允许新文档没有作者，保护时优先使用人员字典映射的文档所有者", () => {
    expect(POWERPOINT_IRM_SCRIPT).toContain("Find-OfficeIdentity");
    const enabledAt = POWERPOINT_IRM_SCRIPT.indexOf("$permission.Enabled = $true");
    const authorAt = POWERPOINT_IRM_SCRIPT.indexOf("$result.officeIdentity = Find-OfficeIdentity");
    expect(enabledAt).toBeGreaterThan(-1);
    expect(authorAt).toBeGreaterThan(enabledAt);
    expect(POWERPOINT_IRM_SCRIPT).toContain("@china.huawei.com");
    expect(POWERPOINT_IRM_SCRIPT).toContain("Where-Object { $_.ToLowerInvariant().EndsWith");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$permission.DocumentAuthor = $author");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$payload.ownerAccount");
  });

  it("文档所有者保留完全控制且不重复写入白名单", () => {
    const owner = pptMember("owner", "l00860412", "zw.lin@huawei.com");
    const reader = pptMember("reader", "l10000000", "reader@huawei.com");
    const policy: PptAccessPolicy = {
      entries: [
        { memberId: owner.id, permission: "read" },
        { memberId: reader.id, permission: "change" },
      ],
      updatedAt: "2026-09-01T00:00:00.000Z",
    };
    const access = buildPowerPointIrmAccess({
      members: [owner, reader],
      policy,
      expiration: new Date("2027-03-01T00:00:00.000Z"),
      officeIdentity: "l00860412@china.huawei.com",
    });
    expect(access.officeIdentity).toBe("l00860412@china.huawei.com");
    expect(access.preferredDocumentOwner).toBe("zw.lin@huawei.com");
    expect(access.grants).toEqual([{
      account: "reader@huawei.com",
      permission: 15,
      expirationUtc: "2027-03-01T00:00:00.000Z",
    }]);
    expect(buildPowerPointIrmAccess({
      members: [owner],
      policy: { ...policy, entries: [{ memberId: owner.id, permission: "read" }] },
      expiration: new Date("2027-03-01T00:00:00.000Z"),
      officeIdentity: "l00860412@china.huawei.com",
    }).grants).toEqual([]);
  });

  it("导出者的 Office 身份和显示账号都不会重复加入普通白名单", () => {
    const office = pptMember("office", "other", "l00860412@china.huawei.com");
    const alias = pptMember("alias", "l00860412", "zw.lin@huawei.com");
    const reader = pptMember("reader", "l10000000", "reader@huawei.com");
    const access = buildPowerPointIrmAccess({
      members: [office, alias, reader],
      policy: {
        entries: [
          { memberId: office.id, permission: "read" },
          { memberId: alias.id, permission: "change" },
          { memberId: reader.id, permission: "read" },
        ],
        updatedAt: "2026-09-01T00:00:00.000Z",
      },
      expiration: new Date("2027-03-01T00:00:00.000Z"),
      officeIdentity: "l00860412@china.huawei.com",
    });
    expect(access.grants.map((item) => item.account)).toEqual(["reader@huawei.com"]);
  });

  it("以 OpenXML 格式保存并等待文件稳定，回读失败时尝试可见窗口", () => {
    expect(POWERPOINT_IRM_SCRIPT).toContain("SaveAs([string]$payload.outputPath, 24)");
    expect(POWERPOINT_IRM_SCRIPT).toContain("Wait-PowerPointFileReady ([string]$payload.outputPath) \"受保护 PPT\"");
    expect(POWERPOINT_IRM_SCRIPT).toContain("隐藏回读失败，转为可见窗口回读");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$powerPoint.Visible = -1");
    expect(POWERPOINT_IRM_SCRIPT).not.toContain("$powerPoint.Visible = $true");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$powerPoint.Presentations.Open([string]$payload.outputPath, -1, 0, 0)");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$powerPoint.Presentations.Open([string]$payload.outputPath, -1, 0, -1)");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$result.stage = $stage");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$result.diagnostics = $diagnostics");
  });

  it("原始 PPT 修复后另存为标准 PPT，并由新的 PowerPoint 实例常规重开", () => {
    expect(POWERPOINT_IRM_SCRIPT).toContain("$powerPoint.Visible = -1");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$powerPoint.Presentations.Open([string]$payload.inputPath, 0, 0, -1)");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$powerPoint.Presentations.Open2007([string]$payload.inputPath, 0, -1, -1, -1)");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$presentation.SaveAs([string]$payload.normalizedPath, 24)");
    expect(POWERPOINT_IRM_SCRIPT).not.toContain("ProtectedViewWindows.Open");
    expect(POWERPOINT_IRM_SCRIPT).not.toContain("$presentation.SaveCopyAs");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$result.sourceSlideCount = [int]$presentation.Slides.Count");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$result.normalizedSlideCount = [int]$presentation.Slides.Count");
    expect(POWERPOINT_IRM_SCRIPT).toContain("$stage = \"常规重开标准化 PPT\"");
    expect(POWERPOINT_IRM_SCRIPT).toContain("标准化 PPT 常规重开成功");
    expect(POWERPOINT_IRM_SCRIPT).toContain("Unblock-File");
    expect(POWERPOINT_IRM_SCRIPT).toContain("Wait-PowerPointFileReady ([string]$payload.inputPath) \"原始 PPT\"");
  });

  it("只对所有者认证、保存和回读类失败回退 Office 登录身份", () => {
    for (const stage of ["设置文档所有者", "保存受保护 PPT", "重新打开受保护 PPT", "校验 IRM 权限"]) {
      expect(shouldRetryWithOfficeIdentity(stage)).toBe(true);
    }
    expect(shouldRetryWithOfficeIdentity("添加白名单权限（第 1 项）")).toBe(false);
    expect(shouldRetryWithOfficeIdentity("设置白名单有效期（第 1 项）")).toBe(false);
    expect(shouldRetryWithOfficeIdentity("常规重开标准化 PPT")).toBe(false);
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
    ])).toThrow("缺少白名单第 1 项");
  });
});

function pptMember(id: string, employeeNumber: string, microsoftAccount: string): Member {
  return {
    id,
    name: id,
    employeeNumber,
    microsoftAccount,
    role: "规划",
    domainL1Ids: [],
    sortOrder: 0,
    createdAt: "2026-09-01T00:00:00.000Z",
    updatedAt: "2026-09-01T00:00:00.000Z",
  };
}
