import { describe, it, expect } from "vitest";
import { mcpLaunchPaths, resolveMcpUserData } from "../src/shared/mcpLaunch";
import { diagnosticErrorCode } from "../src/main/mcpDiagnostics";
describe("独立 MCP 入口", () => {
  it("Windows 使用独立 Node 和独立脚本，不启动 Electron", () => {
    expect(mcpLaunchPaths("win32", "x64", "C:\\工具\\resources")).toEqual({ command: "C:\\工具\\resources\\mcp\\node.exe", args: ["C:\\工具\\resources\\mcp\\server.cjs"], available: true });
  });
  it("macOS 包内与开发路径", () => {
    expect(mcpLaunchPaths("darwin", "arm64", "/App/Resources").command).toBe("/App/Resources/mcp/node");
    expect(mcpLaunchPaths("darwin", "arm64", "/unused", "/project").args).toEqual(["/project/dist-mcp/server.cjs"]);
  });
  it("用户目录与工作台一致，隔离测试支持覆盖", () => {
    expect(resolveMcpUserData("win32", "C:\\Users\\a", { APPDATA: "C:\\Users\\a\\AppData\\Roaming" })).toBe("C:\\Users\\a\\AppData\\Roaming\\需求路标工作台");
    expect(resolveMcpUserData("win32", "C:\\Users\\a", {})).toBe("C:\\Users\\a\\AppData\\Roaming\\需求路标工作台");
    expect(resolveMcpUserData("darwin", "/Users/a", {})).toBe("/Users/a/Library/Application Support/需求路标工作台");
    expect(resolveMcpUserData("linux", "/users/a", { XDG_CONFIG_HOME: "/config" })).toBe("/config/需求路标工作台");
    expect(resolveMcpUserData("darwin", "/Users/a", { SOLUTION_REQUIREMENTS_USER_DATA_DIR: "/isolated" })).toBe("/isolated");
  });
  it("连接提前关闭有独立错误码", () => { expect(diagnosticErrorCode(new Error("MCP error -32000: Connection closed"))).toBe("MCP_CONNECTION_CLOSED"); });
});
