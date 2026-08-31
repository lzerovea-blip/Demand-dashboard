import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, win32 } from "node:path";
import { promisify } from "node:util";
import type { IrmEnvironmentStatus, Member, PptAccessPolicy, PptPermissionLevel } from "../shared/types.js";

const execFileAsync = promisify(execFile);
const MSO_PERMISSION: Record<PptPermissionLevel, number> = { read: 1, change: 15 };

interface PowerShellGrant {
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
  error?: string;
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
  if (actual.length !== grants.length) throw new Error("IRM 权限回读数量不一致，已终止导出");
  for (const expected of grants) {
    const matched = actual.find((item) => item.account.toLocaleLowerCase("en-US") === expected.account.toLocaleLowerCase("en-US"));
    if (!matched || matched.permission !== expected.permission) throw new Error(`IRM 权限回读不一致：${expected.account}`);
    if (Math.abs(Date.parse(matched.expirationUtc) - Date.parse(expected.expirationUtc)) > 60_000) throw new Error(`IRM 到期时间回读不一致：${expected.account}`);
  }
}

async function runPowerPointIrmScript(payload: object, timeout = 45_000): Promise<PowerShellResult> {
  const directory = await mkdtemp(join(tmpdir(), "solution-ppt-irm-"));
  const scriptPath = join(directory, "protect-ppt.ps1");
  const payloadPath = join(directory, "payload.json");
  const resultPath = join(directory, "result.json");
  try {
    await Promise.all([
      writeFile(scriptPath, POWERPOINT_IRM_SCRIPT, "utf8"),
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
    if (!result.ok) throw new Error(result.error || "PowerPoint 企业 IRM 操作失败");
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
try {
  $payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
  try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
  } catch {
    throw "未找到桌面版 Microsoft PowerPoint，或 PowerPoint COM 组件未正确注册：$($_.Exception.Message)"
  }
  $result.powerPointInstalled = $true
  if ($payload.action -eq "probe") {
    $presentation = $powerPoint.Presentations.Add()
    $permission = $presentation.Permission
    $author = [string]$permission.DocumentAuthor
    $result.signedInAccount = $author
    if ([string]::IsNullOrWhiteSpace($author)) { throw "PowerPoint 未登录可使用 IRM 的企业账号" }
    $permission.Enabled = $true
    $result.irmReady = [bool]$permission.Enabled
    if (-not $result.irmReady) { throw "PowerPoint 企业 IRM 服务不可用" }
    $result.ok = $true
  } elseif ($payload.action -eq "protect") {
    $presentation = $powerPoint.Presentations.Open([string]$payload.inputPath, $false, $false, $false)
    $permission = $presentation.Permission
    $author = [string]$permission.DocumentAuthor
    if ([string]::IsNullOrWhiteSpace($author)) { throw "PowerPoint 未登录可使用 IRM 的企业账号" }
    $permission.RemoveAll()
    $permission.Enabled = $true
    foreach ($grant in $payload.grants) {
      $userPermission = $permission.Add([string]$grant.account, [int]$grant.permission)
      $expirationUtc = [DateTime]::Parse([string]$grant.expirationUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
      $userPermission.ExpirationDate = $expirationUtc.ToLocalTime()
    }
    $presentation.SaveAs([string]$payload.outputPath)
    $presentation.Close()
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation) | Out-Null
    $presentation = $null

    $presentation = $powerPoint.Presentations.Open([string]$payload.outputPath, $true, $false, $false)
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
    $result.signedInAccount = [string]$permission.DocumentAuthor
    $result.irmReady = $true
    $result.ok = $true
  } else {
    throw "不支持的 IRM 操作"
  }
} catch {
  $result.error = $_.Exception.Message
} finally {
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
