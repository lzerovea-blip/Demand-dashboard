import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { LocalMcpBridge, callLocalMcpBridge } from "../src/main/mcpBridge";
import { McpProposalStore } from "../src/main/mcpProposals";
import type { AppSnapshot } from "../src/shared/types";

let directory: string;
let bridge: LocalMcpBridge;
const snapshot = { requirements: [], domains: [], products: [], members: [], groupOverrides: [],
  pptAccessPolicy: { entries: [], updatedAt: "2026-01-01T00:00:00.000Z" }, templates: [] } as AppSnapshot;
beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), "mcp-bridge-"));
  const store = new McpProposalStore({ getSnapshot: () => snapshot, applyMcpRequirementBatch: async () => snapshot });
  bridge = new LocalMcpBridge(directory, () => snapshot, store, () => undefined);
});
afterEach(async () => { await bridge.stop(); await rm(directory, { recursive: true, force: true }); });

describe("local MCP bridge", () => {
  it("requires the workbench to be running and enabled", async () => {
    await expect(callLocalMcpBridge(directory, "list_requirements", {})).rejects.toThrow("未运行或未启用");
    await bridge.start();
    if (process.platform !== "win32") {
      const sessionFile = join(directory, "mcp-session.json");
      const descriptor = JSON.parse(await readFile(sessionFile, "utf8")) as { pipe: string };
      expect((await stat(sessionFile)).mode & 0o777).toBe(0o600);
      expect((await stat(descriptor.pipe)).mode & 0o777).toBe(0o600);
    }
    expect(await callLocalMcpBridge(directory, "list_requirements", { offset: 0, limit: 30 })).toEqual({ total: 0, offset: 0, limit: 30, items: [] });
    await bridge.stop();
    await expect(callLocalMcpBridge(directory, "list_requirements", {})).rejects.toThrow("未运行或未启用");
  });
  it("does not expose forbidden mutation actions", async () => {
    await bridge.start();
    await expect(callLocalMcpBridge(directory, "save_tiering_support", {})).rejects.toThrow("不支持的 MCP 操作");
    expect(await callLocalMcpBridge(directory, "list_requirement_references", {})).toEqual({ domains: [], products: [], members: [] });
  });
});
