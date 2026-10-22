import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { Database } from "sql.js";
import { LocalDatabase } from "../src/main/database";
import type { Member, Requirement, SaveRequirementInput, WorkspaceData } from "../src/shared/types";

const environment = vi.hoisted(() => ({ directory: "" }));
vi.mock("electron", () => ({ app: { getPath: () => environment.directory, getAppPath: () => process.cwd() } }));
const timestamp = "2026-01-01T00:00:00.000Z";
const requirement: Requirement = {
  id: "r1", title: "原需求", description: "描述", images: [], domainL0Id: "l0", domainId: "l1",
  scope: "全球", allowLatestVersionUpgrade: true, overseasRegions: [], category: "平台演进", targetMonth: "2027-03",
  productIds: [], planningMemberId: "plan", productManagerMemberId: "pm", seMemberId: "se",
  deviceWorkloadPm: 1, appWorkloadPm: 0, cloudWorkloadPm: 0, researchAppWorkloadPm: 0, algorithmWorkloadPm: 0,
  unallocatedWorkloadPm: 0, workloadPm: 1, createdAt: timestamp, updatedAt: timestamp,
};
const members: Member[] = (["规划", "产品经理", "SE"] as const).map((role, index) => ({ id: ["plan", "pm", "se"][index],
  name: role, role, employeeNumber: `E${index}`, microsoftAccount: `owner${index}@example.com`, domainL1Ids: [], sortOrder: index,
  createdAt: timestamp, updatedAt: timestamp }));
let database: LocalDatabase;
beforeEach(async () => {
  environment.directory = await mkdtemp(join(tmpdir(), "mcp-batch-"));
  Object.defineProperty(process, "resourcesPath", { configurable: true, value: environment.directory });
  database = new LocalDatabase();
  await database.initialize();
  await database.importWorkspace({ requirements: [requirement], members, products: [], groupOverrides: [],
    domains: [{ id: "l0", name: "领域", level: "L0", active: true, sortOrder: 0 },
      { id: "l1", name: "子领域", level: "L1", parentId: "l0", active: true, sortOrder: 1 }],
    pptAccessPolicy: { entries: [], updatedAt: timestamp } } as WorkspaceData, "replace");
});
afterEach(async () => {
  (database as unknown as { db: Database }).db.close();
  await rm(environment.directory, { recursive: true, force: true });
});

describe("approved MCP batches", () => {
  it("creates, copies and deletes through existing requirement and tiering rules", async () => {
    const fields = { ...requirement, title: "新需求", images: undefined, id: undefined } as unknown as Omit<SaveRequirementInput, "images" | "id">;
    await database.applyMcpRequirementBatch([{ type: "create", fields }, { type: "duplicate", id: "r1" }, { type: "delete", id: "r1" }]);
    const state = database.getSnapshot();
    expect(state.requirements).toHaveLength(2);
    expect(state.requirements.map((item) => item.title)).toContain("新需求");
    expect(state.requirements.map((item) => item.title)).toContain("原需求 副本（1）");
    expect(state.tiering?.features.some((item) => item.requirementId === "r1")).toBe(false);
    expect(database.getWorkspaceData().requirementTombstones?.some((item) => item.id === "r1")).toBe(true);
  });
  it("rolls back all operations and file bytes when later validation fails", async () => {
    const before = await readFile(join(environment.directory, "roadmap.db"));
    await expect(database.applyMcpRequirementBatch([
      { type: "update", id: "r1", fields: { title: "先改标题" } },
      { type: "update", id: "r1", fields: { deviceWorkloadPm: -1 } },
    ])).rejects.toThrow("工作量");
    expect(database.getSnapshot().requirements.find((item) => item.id === "r1")?.title).toBe("原需求");
    expect(await readFile(join(environment.directory, "roadmap.db"))).toEqual(before);
  });
});
