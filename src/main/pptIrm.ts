import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, win32 } from "node:path";
import { promisify } from "node:util";
import type { IrmEnvironmentStatus, Member, PptAccessPolicy, PptPermissionLevel } from "../shared/types.js";

const execFileAsync = promisify(execFile);
const MSO_PERMISSION: Record<PptPermissionLevel, number> = { read: 1, change: 15 };

export interface PowerShellGrant {
  account: string;
  permission: number;
  expirationUtc: string;
}

interface PowerShellResult {
  ok: boolean;
  powerPointInstalled?: boolean;
  irmReady?: boolean;
  signedInAccount?: string;
  grants?: PowerShellGrant[];
  stage?: string;
  diagnostics?: string[];
  error?: string;
}

export interface PowerPointIrmErrorDetails {
  stage?: string;
  diagnostics: string[];
}

function firstEnvironmentValue(environment: NodeJS.ProcessEnv, keys: string[]): string | undefined {
  for (const key of keys) {
    const value = environment[key];
    if (value?.trim()) return value.trim();
  }
  return undefined;
}

export function powerShellExecutableCandidates(environment: NodeJS.ProcessEnv = process.env): string[] {
  const candidates: string[] = [];
  const windowsRoot = firstEnvironmentValue(environment, ["SystemRoot", "SYSTEMROOT", "windir", "WINDIR"]);
  if (windowsRoot) {
    candidates.push(
      win32.join(windowsRoot, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
      win32.join(windowsRoot, "Sysnative", "WindowsPowerShell", "v1.0", "powershell.exe"),
    );
  }
  const programFiles = firstEnvironmentValue(environment, ["ProgramW6432", "PROGRAMW6432", "ProgramFiles", "PROGRAMFILES"]);
  if (programFiles) candidates.push(win32.join(programFiles, "PowerShell", "7", "pwsh.exe"));
  candidates.push("powershell.exe", "pwsh.exe");
  return [...new Set(candidates)];
}

export function encodeWindowsPowerShellScript(script: string): string {
  // Windows PowerShell 5.1 treats UTF-8 files without a BOM as the active ANSI
  // code page. The generated script contains Chinese diagnostics, so a BOM is
  // required to prevent corrupted quotes and parser errors on Chinese Windows.
  return script.startsWith("\uFEFF") ? script : `\uFEFF${script}`;
}

function isMissingExecutableError(error: unknown): boolean {
  return typeof error === "object" && error !== null && "code" in error && error.code === "ENOENT";
}

export interface IrmProtectionRequest {
  inputPath: string;
  outputPath: string;
  members: Member[];
  policy: PptAccessPolicy;
  expiration: Date;
}

export function verifyPowerPointIrmGrants(actual: PowerShellGrant[], expected: PowerShellGrant[]): void {
  for (const expectedGrant of expected) {
    const accountMatches = actual.filter(
      (item) => item.account.toLocaleLowerCase("en-US") === expectedGrant.account.toLocaleLowerCase("en-US"),
    );
    if (!accountMatches.length) throw new Error(`IRM 权限回读缺少账号：${expectedGrant.account}`);
    const permissionMatch = accountMatches.find((item) => item.permission === expectedGrant.permission);
    if (!permissionMatch) throw new Error(`IRM 权限回读等级不一致：${expectedGrant.account}`);
    if (Math.abs(Date.parse(permissionMatch.expirationUtc) - Date.parse(expectedGrant.expirationUtc)) > 60_000) {
      throw new Error(`IRM 到期时间回读不一致：${expectedGrant.account}`);
    }
  }
}

export async function checkPowerPointIrmEnvironment(platform = process.platform): Promise<IrmEnvironmentStatus> {
  if (platform !== "win32") {
    return {
      supported: false,
      platform,
      powerPointInstalled: false,
      irmReady: false,
      message: "企业 IRM 保护仅支持已安装桌面版 PowerPoint 的 Windows 环境。",
    };
  }
  try {
    const result = await runPowerPointIrmScript({ action: "probe" });
    return {
      supported: true,
      platform,
      powerPointInstalled: Boolean(result.powerPointInstalled),
      irmReady: Boolean(result.irmReady),
      signedInAccount: result.signedInAccount || undefined,
      message: result.irmReady
        ? `PowerPoint 企业 IRM 可用${result.signedInAccount ? `，当前作者：${result.signedInAccount}` : ""}`
        : (result.error || "PowerPoint 未登录可使用 IRM 的企业账号。"),
    };
  } catch (error) {
    return {
      supported: true,
      platform,
      powerPointInstalled: false,
      irmReady: false,
      message: error instanceof Error ? error.message : String(error),
    };
  }
}

export async function protectPowerPointWithIrm(request: IrmProtectionRequest): Promise<void> {
  if (process.platform !== "win32") throw new Error("受保护 PPT 只能在 Windows 桌面版 PowerPoint 企业 IRM 环境中导出");
  const memberById = new Map(request.members.map((member) => [member.id, member]));
  const grants = request.policy.entries.map((entry) => {
    const member = memberById.get(entry.memberId);
    if (!member) throw new Error("PPT 白名单包含不存在的成员，请重新保存白名单");
    return {
      account: member.microsoftAccount,
      permission: MSO_PERMISSION[entry.permission],
      expirationUtc: request.expiration.toISOString(),
    };
  });
  if (!grants.length) throw new Error("请至少添加一名成员到 PPT 白名单");

  const result = await runPowerPointIrmScript({
    action: "protect",
    inputPath: request.inputPath,
    outputPath: request.outputPath,
    grants,
  }, 120_000);
  const actual = result.grants ?? [];
  verifyPowerPointIrmGrants(actual, grants);
}

async function runPowerPointIrmScript(payload: object, timeout = 45_000): Promise<PowerShellResult> {
  const directory = await mkdtemp(join(tmpdir(), "solution-ppt-irm-"));
  const scriptPath = join(directory, "protect-ppt.ps1");
  const payloadPath = join(directory, "payload.json");
  const resultPath = join(directory, "result.json");
  try {
    await Promise.all([
      writeFile(scriptPath, encodeWindowsPowerShellScript(POWERPOINT_IRM_SCRIPT), "utf8"),
      writeFile(payloadPath, JSON.stringify(payload), "utf8"),
    ]);
    const arguments_ = [
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy", "Bypass",
        "-File", scriptPath,
        "-PayloadPath", payloadPath,
        "-ResultPath", resultPath,
      ];
    let launchError: unknown;
    let commandFound = false;
    for (const executable of powerShellExecutableCandidates()) {
      try {
        await execFileAsync(executable, arguments_, { timeout, windowsHide: true, maxBuffer: 1024 * 1024 });
        commandFound = true;
        launchError = undefined;
        break;
      } catch (error) {
        launchError = error;
        if (isMissingExecutableError(error)) continue;
        // The script writes a structured result before returning a failure code.
        commandFound = true;
        break;
      }
    }
    if (!commandFound) {
      throw new Error("未找到 Windows PowerShell（已检查系统目录、PowerShell 7 和 PATH）");
    }
    let rawResult: string;
    try {
      rawResult = await readFile(resultPath, "utf8");
    } catch {
      const detail = launchError instanceof Error ? launchError.message : "PowerShell 未返回检测结果";
      throw new Error(`PowerShell 未能完成 PowerPoint 检测：${detail}`);
    }
    const result = JSON.parse(rawResult.replace(/^\uFEFF/u, "")) as PowerShellResult;
    if (!result.ok) {
      const error = new Error(result.error || "PowerPoint 企业 IRM 操作失败") as Error & PowerPointIrmErrorDetails;
      error.stage = result.stage;
      error.diagnostics = result.diagnostics ?? [];
      throw error;
    }
    return result;
  } catch (error) {
    throw error;
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
}

export const POWERPOINT_IRM_SCRIPT = String.raw`param(
  [Parameter(Mandatory=$true)][string]$PayloadPath,
  [Parameter(Mandatory=$true)][string]$ResultPath
)
$ErrorActionPreference = "Stop"
$powerPoint = $null
$presentation = $null
$protectedViewWindow = $null
$result = @{ ok = $false; powerPointInstalled = $false; irmReady = $false }
$stage = "初始化"
$diagnostics = @()
function Find-OfficeIdentity([string]$DocumentAuthor) {
  $candidates = @()
  if (-not [string]::IsNullOrWhiteSpace($DocumentAuthor)) { $candidates += $DocumentAuthor }
  foreach ($root in @(
    "HKCU:\Software\Microsoft\Office\16.0\Common\Identity",
    "HKCU:\Software\Microsoft\Office\Common\Identity"
  )) {
    if (-not (Test-Path -LiteralPath $root)) { continue }
    $rootProperties = Get-ItemProperty -LiteralPath $root -ErrorAction SilentlyContinue
    foreach ($name in @("IdentityName", "EmailAddress", "UserEmail", "ADUserName")) {
      if ($rootProperties -and $rootProperties.$name) { $candidates += [string]$rootProperties.$name }
    }
    $identitiesPath = Join-Path $root "Identities"
    if (Test-Path -LiteralPath $identitiesPath) {
      foreach ($identityKey in Get-ChildItem -LiteralPath $identitiesPath -ErrorAction SilentlyContinue) {
        $identity = Get-ItemProperty -LiteralPath $identityKey.PSPath -ErrorAction SilentlyContinue
        foreach ($name in @("EmailAddress", "IdentityName", "UserEmail", "ADUserName")) {
          if ($identity -and $identity.$name) { $candidates += [string]$identity.$name }
        }
      }
    }
  }
  foreach ($candidate in ($candidates | Select-Object -Unique)) {
    $normalized = $candidate.Trim()
    if ($normalized -match "^[^@\s]+@[^@\s]+\.[^@\s]+$") { return $normalized }
  }
  return ""
}
function Wait-PowerPointFileReady([string]$Path, [string]$Label) {
  $lastLength = -1
  $stableChecks = 0
  for ($attempt = 0; $attempt -lt 40; $attempt++) {
    if (Test-Path -LiteralPath $Path) {
      $length = (Get-Item -LiteralPath $Path).Length
      if ($length -gt 0 -and $length -eq $lastLength) { $stableChecks++ } else { $stableChecks = 0 }
      $lastLength = $length
      if ($stableChecks -ge 2) { return }
    }
    Start-Sleep -Milliseconds 250
  }
  throw "$Label 文件未完成写入"
}
try {
  $stage = "读取请求"
  $payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $diagnostics += "请求已读取"
  $stage = "启动 PowerPoint"
  try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
  } catch {
    throw "未找到桌面版 Microsoft PowerPoint，或 PowerPoint COM 组件未正确注册：$($_.Exception.Message)"
  }
  $result.powerPointInstalled = $true
  $diagnostics += "PowerPoint COM 已启动"
  if ($payload.action -eq "probe") {
    $stage = "检测 IRM"
    $presentation = $powerPoint.Presentations.Add()
    $permission = $presentation.Permission
    $permission.Enabled = $true
    $result.irmReady = [bool]$permission.Enabled
    if (-not $result.irmReady) { throw "PowerPoint 企业 IRM 服务不可用" }
    $result.signedInAccount = Find-OfficeIdentity ([string]$permission.DocumentAuthor)
    $diagnostics += "IRM 已启用"
    $result.ok = $true
  } elseif ($payload.action -eq "protect") {
    $stage = "打开原始 PPT"
    Wait-PowerPointFileReady ([string]$payload.inputPath) "原始 PPT"
    $diagnostics += "原始 PPT ZIP/XML 已校验"
    try { Unblock-File -LiteralPath ([string]$payload.inputPath) -ErrorAction SilentlyContinue } catch {}
    $regularOpenError = ""
    try {
      $powerPoint.Visible = $true
      $presentation = $powerPoint.Presentations.Open([string]$payload.inputPath, $false, $false, $true)
      $diagnostics += "原始 PPT 常规打开成功"
    } catch {
      $regularOpenError = $_.Exception.Message
      $diagnostics += "常规打开失败，尝试修复模式"
      try {
        $protectedViewWindow = $powerPoint.ProtectedViewWindows.Open([string]$payload.inputPath, "", $true)
        $presentation = $protectedViewWindow.Edit()
        $diagnostics += "原始 PPT 修复模式打开成功"
      } catch {
        throw "原始 PPT 已通过 ZIP/XML 校验，但 PowerPoint 仍无法打开。常规打开：$regularOpenError；修复模式：$($_.Exception.Message)"
      }
    }
    $permission = $presentation.Permission
    $stage = "识别 Office 账号"
    $author = Find-OfficeIdentity ([string]$permission.DocumentAuthor)
    $permission.RemoveAll()
    $permission.Enabled = $true
    if (-not $permission.Enabled) { throw "PowerPoint 企业 IRM 服务不可用" }
    if ([string]::IsNullOrWhiteSpace($author)) {
      $author = Find-OfficeIdentity ([string]$permission.DocumentAuthor)
    }
    if ([string]::IsNullOrWhiteSpace($author)) {
      throw "无法识别当前 PowerPoint 企业账号，不能为导出者保留所有者权限。请先在 PowerPoint 的账户页面确认已登录企业邮箱"
    }
    try {
      $permission.DocumentAuthor = $author
    } catch {
      throw "无法将当前 PowerPoint 企业账号设置为文档所有者：$author。$($_.Exception.Message)"
    }
    $diagnostics += "文档所有者已设置"
    $stage = "写入白名单权限"
    foreach ($grant in $payload.grants) {
      $userPermission = $permission.Add([string]$grant.account, [int]$grant.permission)
      $expirationUtc = [DateTime]::Parse([string]$grant.expirationUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
      $userPermission.ExpirationDate = $expirationUtc.ToLocalTime()
    }
    $diagnostics += "白名单权限已写入"
    $stage = "保存受保护 PPT"
    $presentation.SaveAs([string]$payload.outputPath, 24)
    $presentation.Close()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) | Out-Null
    $presentation = $null
    try { $powerPoint.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint) | Out-Null } catch {}
    $powerPoint = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $stage = "等待文件写入"
    Wait-PowerPointFileReady ([string]$payload.outputPath) "受保护 PPT"
    $diagnostics += "受保护文件已稳定"

    $stage = "重新打开受保护 PPT"
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $hiddenOpenError = ""
    try {
      $presentation = $powerPoint.Presentations.Open([string]$payload.outputPath, $true, $false, $false)
      $diagnostics += "隐藏回读成功"
    } catch {
      $hiddenOpenError = $_.Exception.Message
      $diagnostics += "隐藏回读失败，转为可见窗口回读"
      try {
        $powerPoint.Visible = $true
        $presentation = $powerPoint.Presentations.Open([string]$payload.outputPath, $true, $false, $true)
        $diagnostics += "可见窗口回读成功"
      } catch {
        throw "保护后的 PPT 已保存，但 PowerPoint 无法重新打开。隐藏验证：$hiddenOpenError；可见窗口验证：$($_.Exception.Message)。请确认当前 Office 账号 $author 有权使用企业 IRM"
      }
    }
    $stage = "校验 IRM 权限"
    $permission = $presentation.Permission
    if (-not $permission.Enabled) { throw "保存后的 PPT 未启用 IRM 权限" }
    $verified = @()
    for ($index = 1; $index -le $permission.Count; $index++) {
      $item = $permission.Item($index)
      $verified += @{
        account = [string]$item.UserId
        permission = [int]$item.Permission
        expirationUtc = ([DateTime]$item.ExpirationDate).ToUniversalTime().ToString("o")
      }
    }
    $result.grants = $verified
    $result.signedInAccount = $author
    $result.irmReady = $true
    $diagnostics += "IRM 权限回读完成"
    $result.ok = $true
  } else {
    throw "不支持的 IRM 操作"
  }
} catch {
  $result.error = $_.Exception.Message
} finally {
  $result.stage = $stage
  $result.diagnostics = $diagnostics
  if ($presentation -ne $null) {
    try { $presentation.Close() } catch {}
    try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) | Out-Null } catch {}
  }
  if ($protectedViewWindow -ne $null) {
    try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($protectedViewWindow) | Out-Null } catch {}
  }
  if ($powerPoint -ne $null) {
    try { $powerPoint.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint) | Out-Null } catch {}
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  $result | ConvertTo-Json -Depth 6 -Compress | Set-Content -LiteralPath $ResultPath -Encoding UTF8
}
if (-not $result.ok) { exit 1 }
`;
