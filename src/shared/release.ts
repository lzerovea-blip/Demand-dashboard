const RELEASE_REPOSITORY_URL = "https://github.com/lzerovea-blip/Demand-dashboard";

export function windowsSetupDownloadUrl(version: string): string {
  const normalizedVersion = version.trim().replace(/^v/i, "");
  const assetName = `Solution-Requirements-Manager-Setup-${normalizedVersion}-x64.exe`;
  return `${RELEASE_REPOSITORY_URL}/releases/download/v${normalizedVersion}/${assetName}`;
}
