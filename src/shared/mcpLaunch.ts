import { posix, win32 } from "node:path";
export interface McpLaunchConfiguration { command: string; args: string[]; available: boolean; error?: string }
export function resolveMcpUserData(platform: string, home: string, env: Record<string, string | undefined>): string {
  if (env.SOLUTION_REQUIREMENTS_USER_DATA_DIR) return env.SOLUTION_REQUIREMENTS_USER_DATA_DIR;
  const path = platform === "win32" ? win32 : posix;
  const base = platform === "win32" ? env.APPDATA || path.join(home, "AppData", "Roaming")
    : platform === "darwin" ? path.join(home, "Library", "Application Support") : env.XDG_CONFIG_HOME || path.join(home, ".config");
  return path.join(base, "需求路标工作台");
}
export function mcpLaunchPaths(platform: string, arch: string, resources: string, appRoot?: string): McpLaunchConfiguration {
  const path = platform === "win32" ? win32 : posix;
  const runtime = platform === "win32" ? "node.exe" : "node";
  return appRoot ? { command: path.join(appRoot, "build-resources", "mcp-runtime", `${platform}-${arch}`, runtime),
    args: [path.join(appRoot, "dist-mcp", "server.cjs")], available: true }
    : { command: path.join(resources, "mcp", runtime), args: [path.join(resources, "mcp", "server.cjs")], available: true };
}
