import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, win32 } from "node:path";
import { promisify } from "node:util";
import { resolvePptDocumentOwnerAccount } from "../shared/members.js";
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
  officeIdentity?: string;
  normalized?: boolean;
  normalizationMethod?: string;
  sourceSlideCount?: number;
  normalizedSlideCount?: number;
  effectiveDocumentOwner?: string;
  verifiedGrantCount?: number;
  grants?: PowerShellGrant[];
  stage?: string;
  diagnostics?: string[];
  error?: string;
}

export interface PowerPointIrmErrorDetails {
  stage?: string;
  diagnostics: string[];
  powerPointInstalled?: boolean;
  irmReady?: boolean;
  officeIdentity?: string;
  preferredDocumentOwner?: string;
  effectiveDocumentOwner?: string;
  fallbackUsed?: boolean;
  normalized?: boolean;
  normalizationMethod?: string;
  sourceSlideCount?: number;
  normalizedSlideCount?: number;
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
  normalizedPath: string;
  preferredOutputPath: string;
  fallbackOutputPath: string;
  members: Member[];
  policy: PptAccessPolicy;
  expiration: Date;
  officeIdentity?: string;
}

export interface IrmProtectionResult {
  outputPath: string;
  officeIdentity?: string;
  preferredDocumentOwner?: string;
  effectiveDocumentOwner?: string;
  fallbackUsed: boolean;
  normalized: boolean;
  verifiedGrantCount: number;
}

export function buildPowerPointIrmAccess(
  request: Pick<IrmProtectionRequest, "members" | "policy" | "expiration" | "officeIdentity">,
): { officeIdentity?: string; preferredDocumentOwner?: string; grants: PowerShellGrant[] } {
  const memberById = new Map(request.members.map((member) => [member.id, member]));
  const requestedGrants = request.policy.entries.map((entry) => {
    const member = memberById.get(entry.memberId);
    if (!member) throw new Error("PPT 白名单包含不存在的成员，请重新保存白名单");
    return {
      account: member.microsoftAccount,
      permission: MSO_PERMISSION[entry.permission],
      expirationUtc: request.expiration.toISOString(),
    };
  });
  if (!requestedGrants.length) throw new Error("请至少添加一名成员到 PPT 白名单");

  const officeIdentity = request.officeIdentity?.trim().toLocaleLowerCase("en-US");
  const preferredDocumentOwner = resolvePptDocumentOwnerAccount(officeIdentity, request.members);
  const ownerIdentities = new Set(
    [officeIdentity, preferredDocumentOwner].filter((value): value is string => Boolean(value)),
  );
  return {
    officeIdentity,
    preferredDocumentOwner,
    // PowerPoint already grants DocumentAuthor full control. Adding that same
    // identity again through Permission.Add raises E_INVALIDARG in some Office
    // enterprise builds, so the owner must not be duplicated as a whitelist grant.
    grants: requestedGrants.filter((grant) => !ownerIdentities.has(grant.account.toLocaleLowerCase("en-US"))),
  };
}

export function verifyPowerPointIrmGrants(actual: PowerShellGrant[], expected: PowerShellGrant[]): void {
  for (const [index, expectedGrant] of expected.entries()) {
    const accountMatches = actual.filter(
      (item) => item.account.toLocaleLowerCase("en-US") === expectedGrant.account.toLocaleLowerCase("en-US"),
    );
    if (!accountMatches.length) throw new Error(`IRM 权限回读缺少白名单第 ${index + 1} 项`);
    const permissionMatch = accountMatches.find((item) => item.permission === expectedGrant.permission);
    if (!permissionMatch) throw new Error(`IRM 权限回读等级不一致（白名单第 ${index + 1} 项）`);
    if (Math.abs(Date.parse(permissionMatch.expirationUtc) - Date.parse(expectedGrant.expirationUtc)) > 60_000) {
      throw new Error(`IRM 到期时间回读不一致（白名单第 ${index + 1} 项）`);
    }
  }
}

export function resolveIrmEnvironmentStatus(
  status: IrmEnvironmentStatus,
  members: readonly Member[],
): IrmEnvironmentStatus {
  if (!status.supported || !status.powerPointInstalled || !status.irmReady) return status;
  const preferredDocumentOwner = resolvePptDocumentOwnerAccount(status.officeIdentity, members);
  if (!status.officeIdentity || !preferredDocumentOwner) {
    return {
      ...status,
      irmReady: false,
      preferredDocumentOwner,
      message: "PowerPoint 企业 IRM 服务可用，但无法识别当前 Office 登录身份。请在 PowerPoint 的“文件 → 账户”中确认企业账号已登录后重试。",
    };
  }
  return {
    ...status,
    preferredDocumentOwner,
    message: `PowerPoint 企业 IRM 可用；Office 登录身份：${status.officeIdentity}；文档所有者：${preferredDocumentOwner}。`,
  };
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
      officeIdentity: result.officeIdentity || undefined,
      message: result.irmReady
        ? `PowerPoint 企业 IRM 可用${result.officeIdentity ? `，Office 登录身份：${result.officeIdentity}` : ""}`
        : (result.error || "PowerPoint 未登录可使用 IRM 的企业账号。"),
    };
  } catch (error) {
    const details = error as Partial<PowerPointIrmErrorDetails>;
    return {
      supported: true,
      platform,
      powerPointInstalled: Boolean(details.powerPointInstalled),
      irmReady: Boolean(details.irmReady),
      officeIdentity: details.officeIdentity || undefined,
      message: error instanceof Error ? error.message : String(error),
    };
  }
}

export async function protectPowerPointWithIrm(request: IrmProtectionRequest): Promise<IrmProtectionResult> {
  if (process.platform !== "win32") throw new Error("受保护 PPT 只能在 Windows 桌面版 PowerPoint 企业 IRM 环境中导出");
  const { officeIdentity, preferredDocumentOwner, grants } = buildPowerPointIrmAccess(request);
  const primaryOwner = preferredDocumentOwner || officeIdentity;
  if (!primaryOwner) throw new Error("无法识别 Office 登录身份，也无法从人员字典确定文档所有者");

  let normalization: PowerShellResult;
  try {
    normalization = await runPowerPointIrmScript({
      action: "normalize",
      inputPath: request.inputPath,
      normalizedPath: request.normalizedPath,
    }, 120_000);
  } catch (error) {
    const details = asPowerPointIrmError(error);
    details.officeIdentity = officeIdentity;
    details.preferredDocumentOwner = preferredDocumentOwner;
    details.normalized = false;
    throw error;
  }

  const protectAttempt = (ownerAccount: string, outputPath: string) => runPowerPointIrmScript({
    action: "protect",
    inputPath: request.normalizedPath,
    outputPath,
    grants,
    ownerAccount,
    officeIdentity,
  }, 120_000);

  try {
    const result = await protectAttempt(primaryOwner, request.preferredOutputPath);
    const actual = result.grants ?? [];
    verifyPowerPointIrmGrants(actual, grants);
    return {
      outputPath: request.preferredOutputPath,
      officeIdentity,
      preferredDocumentOwner,
      effectiveDocumentOwner: result.effectiveDocumentOwner || primaryOwner,
      fallbackUsed: false,
      normalized: Boolean(normalization.normalized),
      verifiedGrantCount: result.verifiedGrantCount ?? actual.length,
    };
  } catch (primaryError) {
    const primaryDetails = asPowerPointIrmError(primaryError);
    primaryDetails.diagnostics = [...(normalization.diagnostics ?? []), ...primaryDetails.diagnostics];
    const canFallback = Boolean(
      officeIdentity
      && officeIdentity !== primaryOwner.toLocaleLowerCase("en-US")
      && shouldRetryWithOfficeIdentity(primaryDetails.stage),
    );
    if (!canFallback || !officeIdentity) {
      primaryDetails.officeIdentity = officeIdentity;
      primaryDetails.preferredDocumentOwner = preferredDocumentOwner;
      primaryDetails.effectiveDocumentOwner = primaryDetails.effectiveDocumentOwner || primaryOwner;
      primaryDetails.fallbackUsed = false;
      primaryDetails.normalized = Boolean(normalization.normalized);
      throw primaryError;
    }
    try {
      const fallback = await protectAttempt(officeIdentity, request.fallbackOutputPath);
      const actual = fallback.grants ?? [];
      verifyPowerPointIrmGrants(actual, grants);
      return {
        outputPath: request.fallbackOutputPath,
        officeIdentity,
        preferredDocumentOwner,
        effectiveDocumentOwner: fallback.effectiveDocumentOwner || officeIdentity,
        fallbackUsed: true,
        normalized: Boolean(normalization.normalized),
        verifiedGrantCount: fallback.verifiedGrantCount ?? actual.length,
      };
    } catch (fallbackError) {
      const fallbackDetails = asPowerPointIrmError(fallbackError);
      fallbackDetails.diagnostics = [
        ...primaryDetails.diagnostics,
        "首选文档所有者保护失败，已使用 Office 登录身份重试",
        ...fallbackDetails.diagnostics,
      ];
      fallbackDetails.officeIdentity = officeIdentity;
      fallbackDetails.preferredDocumentOwner = preferredDocumentOwner;
      fallbackDetails.effectiveDocumentOwner = fallbackDetails.effectiveDocumentOwner || officeIdentity;
      fallbackDetails.fallbackUsed = true;
      fallbackDetails.normalized = Boolean(normalization.normalized);
      const primaryMessage = primaryError instanceof Error ? primaryError.message : String(primaryError);
      const fallbackMessage = fallbackError instanceof Error ? fallbackError.message : String(fallbackError);
      const combined = new Error(`首选文档所有者保护失败：${primaryMessage}；Office 登录身份回退仍失败：${fallbackMessage}`) as Error & PowerPointIrmErrorDetails;
      Object.assign(combined, fallbackDetails);
      throw combined;
    }
  }
}

export function shouldRetryWithOfficeIdentity(stage: string | undefined): boolean {
  return Boolean(stage && /设置文档所有者|保存受保护 PPT|重新打开受保护 PPT|校验 IRM 权限/u.test(stage));
}

function asPowerPointIrmError(error: unknown): PowerPointIrmErrorDetails {
  const candidate = error as Partial<PowerPointIrmErrorDetails>;
  if (!Array.isArray(candidate.diagnostics)) candidate.diagnostics = [];
  return candidate as PowerPointIrmErrorDetails;
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
      const detail = result.error || "PowerPoint 企业 IRM 操作失败";
      const error = new Error(result.stage ? `${result.stage}失败：${detail}` : detail) as Error & PowerPointIrmErrorDetails;
      error.stage = result.stage;
      error.diagnostics = result.diagnostics ?? [];
      error.powerPointInstalled = result.powerPointInstalled;
      error.irmReady = result.irmReady;
      error.officeIdentity = result.officeIdentity;
      error.effectiveDocumentOwner = result.effectiveDocumentOwner;
      error.normalized = result.normalized;
      error.normalizationMethod = result.normalizationMethod;
      error.sourceSlideCount = result.sourceSlideCount;
      error.normalizedSlideCount = result.normalizedSlideCount;
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
$result = @{ ok = $false; powerPointInstalled = $false; irmReady = $false }
$stage = "初始化"
$diagnostics = @()
function Find-OfficeIdentity([string]$DocumentAuthor) {
  $candidates = @()
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
  if ($env:USERNAME -match "^[Ll][0-9]+$") { $candidates += "$($env:USERNAME)@china.huawei.com" }
  if (-not [string]::IsNullOrWhiteSpace($DocumentAuthor)) { $candidates += $DocumentAuthor }
  $validCandidates = @()
  foreach ($candidate in ($candidates | Select-Object -Unique)) {
    $normalized = $candidate.Trim()
    if ($normalized -match "^[^@\s]+@[^@\s]+\.[^@\s]+$") { $validCandidates += $normalized }
  }
  $officeLogin = $validCandidates | Where-Object { $_.ToLowerInvariant().EndsWith("@china.huawei.com") } | Select-Object -First 1
  if ($officeLogin) { return $officeLogin }
  if ($validCandidates.Count -gt 0) { return $validCandidates[0] }
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
    $result.officeIdentity = Find-OfficeIdentity ([string]$permission.DocumentAuthor)
    $diagnostics += "IRM 已启用"
    $result.ok = $true
  } elseif ($payload.action -eq "normalize") {
    $stage = "打开原始 PPT"
    Wait-PowerPointFileReady ([string]$payload.inputPath) "原始 PPT"
    $diagnostics += "原始 PPT ZIP/XML 已校验"
    try { Unblock-File -LiteralPath ([string]$payload.inputPath) -ErrorAction SilentlyContinue } catch {}
    $regularOpenError = ""
    try {
      $powerPoint.Visible = -1
      $presentation = $powerPoint.Presentations.Open([string]$payload.inputPath, 0, 0, -1)
      $diagnostics += "原始 PPT 常规打开成功"
      $result.normalizationMethod = "常规打开 + SaveAs OpenXML"
    } catch {
      $regularOpenError = $_.Exception.Message
      $diagnostics += "常规打开失败，尝试 Open2007 修复模式"
      try {
        # Open2007 的第 5 个参数是 OpenAndRepair。Untitled=-1 可避免
        # PowerPoint 将修复结果继续绑定到原始临时文件。
        $presentation = $powerPoint.Presentations.Open2007([string]$payload.inputPath, 0, -1, -1, -1)
        $diagnostics += "原始 PPT Open2007 修复模式打开成功"
        $result.normalizationMethod = "Open2007 修复 + SaveAs OpenXML"
      } catch {
        throw "原始 PPT 已通过 ZIP/XML 校验，但 PowerPoint 仍无法打开。常规打开：$regularOpenError；Open2007 修复：$($_.Exception.Message)"
      }
    }
    $result.sourceSlideCount = [int]$presentation.Slides.Count
    $stage = "保存标准化 PPT"
    try {
      # SaveAs 会让 PowerPoint 按当前对象模型完整重写 OpenXML 包；
      # SaveCopyAs 只复制当前文件状态，不能可靠提交修复结果。
      $presentation.SaveAs([string]$payload.normalizedPath, 24)
    } catch {
      throw "PowerPoint 无法保存标准化 PPT：$($_.Exception.Message)"
    }
    $diagnostics += "标准化 PPT 已按 OpenXML SaveAs 保存"
    $presentation.Close()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) | Out-Null
    $presentation = $null
    try { $powerPoint.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint) | Out-Null } catch {}
    $powerPoint = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Wait-PowerPointFileReady ([string]$payload.normalizedPath) "标准化 PPT"
    $stage = "常规重开标准化 PPT"
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    try {
      $presentation = $powerPoint.Presentations.Open([string]$payload.normalizedPath, 0, 0, -1)
    } catch {
      throw "标准化 PPT 仍无法由新的 PowerPoint 实例常规打开：$($_.Exception.Message)"
    }
    $result.normalizedSlideCount = [int]$presentation.Slides.Count
    if ($result.normalizedSlideCount -ne $result.sourceSlideCount) {
      throw "标准化 PPT 页数不一致：保存前 $($result.sourceSlideCount) 页，重开后 $($result.normalizedSlideCount) 页"
    }
    $diagnostics += "标准化 PPT 常规重开成功（$($result.normalizedSlideCount) 页）"
    $result.normalized = $true
    $result.ok = $true
  } elseif ($payload.action -eq "protect") {
    $result.officeIdentity = [string]$payload.officeIdentity
    $result.normalized = $true
    $stage = "打开标准化 PPT"
    Wait-PowerPointFileReady ([string]$payload.inputPath) "标准化 PPT"
    $powerPoint.Visible = -1
    try {
      $presentation = $powerPoint.Presentations.Open([string]$payload.inputPath, 0, 0, -1)
      $diagnostics += "标准化 PPT 常规打开成功"
    } catch {
      throw "标准化 PPT 在写入 IRM 前无法常规打开：$($_.Exception.Message)"
    }
    $permission = $presentation.Permission
    $stage = "设置文档所有者"
    $author = ([string]$payload.ownerAccount).Trim().ToLowerInvariant()
    $permission.RemoveAll()
    $permission.Enabled = $true
    if (-not $permission.Enabled) { throw "PowerPoint 企业 IRM 服务不可用" }
    if ([string]::IsNullOrWhiteSpace($author)) { throw "未提供有效的文档所有者账号" }
    try {
      $permission.DocumentAuthor = $author
    } catch {
      throw "无法设置文档所有者：$($_.Exception.Message)"
    }
    $effectiveAuthor = ([string]$permission.DocumentAuthor).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($effectiveAuthor)) { throw "PowerPoint 未返回有效的文档所有者" }
    $result.effectiveDocumentOwner = $effectiveAuthor
    $diagnostics += "文档所有者已设置"
    $stage = "写入白名单权限"
    $grantIndex = 0
    foreach ($grant in $payload.grants) {
      $grantIndex++
      $stage = "添加白名单权限（第 $grantIndex 项）"
      try {
        $permissionValue = [Microsoft.Office.Core.MsoPermission][int]$grant.permission
        $userPermission = $permission.Add([string]$grant.account, $permissionValue)
      } catch {
        throw "白名单第 $grantIndex 项权限添加失败：$($_.Exception.Message)"
      }
      $diagnostics += "白名单第 $grantIndex 项权限已添加"
      $stage = "设置白名单有效期（第 $grantIndex 项）"
      try {
        $expirationUtc = [DateTime]::Parse([string]$grant.expirationUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
        $userPermission.ExpirationDate = $expirationUtc.ToLocalTime()
      } catch {
        throw "白名单第 $grantIndex 项有效期设置失败：$($_.Exception.Message)"
      }
      $diagnostics += "白名单第 $grantIndex 项有效期已设置"
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
      $presentation = $powerPoint.Presentations.Open([string]$payload.outputPath, -1, 0, 0)
      $diagnostics += "隐藏回读成功"
    } catch {
      $hiddenOpenError = $_.Exception.Message
      $diagnostics += "隐藏回读失败，转为可见窗口回读"
      try {
        $powerPoint.Visible = -1
        $presentation = $powerPoint.Presentations.Open([string]$payload.outputPath, -1, 0, -1)
        $diagnostics += "可见窗口回读成功"
      } catch {
        throw "保护后的 PPT 已保存，但 PowerPoint 无法重新打开。隐藏验证：$hiddenOpenError；可见窗口验证：$($_.Exception.Message)"
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
    $result.verifiedGrantCount = $verified.Count
    $result.effectiveDocumentOwner = ([string]$permission.DocumentAuthor).Trim().ToLowerInvariant()
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
