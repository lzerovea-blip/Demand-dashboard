import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { Database } from "sql.js";
import { LocalDatabase } from "../src/main/database";
import { McpProposalStore } from "../src/main/mcpProposals";
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
  it("previews all dictionary types without writing, then persists only after approval and restart", async () => {
    const before = database.getSnapshot();
    const roadmapId = before.roadmaps![0].id;
    const policy = JSON.stringify(before.pptAccessPolicy);
    const bytes = await readFile(join(environment.directory, "roadmap.db"));
    const store = new McpProposalStore(database);
    const proposal = store.proposeDictionaries([
      { kind: "domain_l0", name: "新一级领域" },
      { kind: "domain_l1", name: "新二级领域", parentId: "l0" },
      { kind: "product", name: "新产品" },
      { kind: "member", name: "新算法成员", employeeNumber: "NEW1", microsoftAccount: "new1@example.com", role: "算法", domainL1Ids: ["l1"] },
      { kind: "roadmap", name: "新医疗路标", roadmapKind: "medical" },
      { kind: "lane", name: "新泳道", roadmapId },
    ]);
    expect(proposal.items).toHaveLength(6);
    expect(await readFile(join(environment.directory, "roadmap.db"))).toEqual(bytes);
    expect(database.getSnapshot().products).toHaveLength(0);
    await store.decide(proposal.id, true);
    const state = database.getSnapshot();
    expect(state.domains.some(item => item.name === "新二级领域" && item.parentId === "l0")).toBe(true);
    expect(state.members.some(item => item.name === "新算法成员" && item.role === "算法")).toBe(true);
    expect(state.roadmaps!.some(item => item.name === "新医疗路标" && item.kind === "medical")).toBe(true);
    expect(state.roadmapLanes!.some(item => item.name === "新泳道" && item.roadmapId === roadmapId)).toBe(true);
    expect(state.requirements).toEqual(before.requirements);
    expect(JSON.stringify(state.pptAccessPolicy)).toBe(policy);
    const reopened = new LocalDatabase();
    try { await reopened.initialize(); expect(reopened.getSnapshot().products.some(item => item.name === "新产品")).toBe(true); }
    finally { (reopened as unknown as { db: Database }).db.close(); }
  });
  it("rejects duplicates, unknown parents, IDs and unsupported roles without changing anything", () => {
    const store = new McpProposalStore(database);
    for (const entries of [
      [{ kind: "domain_l0", name: "领域" }],
      [{ kind: "domain_l1", name: "新", parentId: "missing" }],
      [{ kind: "lane", name: "新", roadmapId: "missing" }],
      [{ kind: "product", name: "新", id: "overwrite" }],
      [{ kind: "product", name: "新" }, { kind: "product", name: " 新 " }],
      [{ kind: "member", name: "新", employeeNumber: "E0", microsoftAccount: "new@example.com", domainL1Ids: [] }],
      [{ kind: "member", name: "新", employeeNumber: "N", microsoftAccount: "owner0@example.com", domainL1Ids: [] }],
      [{ kind: "member", name: "新", employeeNumber: "N", microsoftAccount: "bad", domainL1Ids: [] }],
      [{ kind: "member", name: "新", employeeNumber: "N", microsoftAccount: "n@example.com", domainL1Ids: ["l0"] }],
      [{ kind: "member", name: "新", employeeNumber: "N", microsoftAccount: "n@example.com", role: "非法角色", domainL1Ids: [] }],
    ]) expect(() => store.proposeDictionaries(entries)).toThrow();
    expect(store.listPending()).toHaveLength(0);
  });
  it("rechecks duplicates and parent existence at approval, so no overwrite occurs", async () => {
    const store = new McpProposalStore(database);
    const proposal = store.proposeDictionaries([{ kind: "product", name: "并发产品" }]);
    await database.saveProduct({ name: "并发产品" });
    await expect(store.decide(proposal.id, true)).rejects.toThrow("已存在");
    expect(database.getSnapshot().products.filter(item => item.name === "并发产品")).toHaveLength(1);
    const parent = store.proposeDictionaries([{ kind: "domain_l1", name: "待新增", parentId: "l0" }]);
    parent.items[0].operation = { type: "create_dictionary", entry: { kind: "domain_l1", name: "待新增", parentId: "missing" } };
    await expect(store.decide(parent.id, true)).rejects.toThrow("已有的领域 L0");
  });
  it("rolls back all staged dictionary types if a later operation fails", async () => {
    const before = database.getSnapshot();
    const bytes = await readFile(join(environment.directory, "roadmap.db"));
    vi.spyOn(database, "saveRoadmap").mockRejectedValueOnce(new Error("模拟后续写入失败"));
    await expect(database.applyMcpDictionaryBatch([
      { kind: "product", name: "不可部分写入" },
      { kind: "roadmap", name: "失败路标", roadmapKind: "standard" },
    ])).rejects.toThrow("模拟后续写入失败");
    expect(database.getSnapshot()).toEqual(before);
    expect(await readFile(join(environment.directory, "roadmap.db"))).toEqual(bytes);
  });
  it("rejecting or expiring a dictionary proposal never writes", async () => {
    let now = 0;
    const store = new McpProposalStore(database, undefined, () => now);
    const declined = store.proposeDictionaries([{ kind: "product", name: "拒绝" }]);
    await store.decide(declined.id, false);
    const expired = store.proposeDictionaries([{ kind: "product", name: "过期" }]);
    now = 31 * 60 * 1000;
    await expect(store.decide(expired.id, true)).rejects.toThrow("不可处理");
    expect(database.getSnapshot().products).toHaveLength(0);
  });
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
