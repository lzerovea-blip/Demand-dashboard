import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { Database } from "sql.js";
import JSZip from "jszip";
import { LocalDatabase } from "../src/main/database";
import { createWorkspaceWorkbook, inspectWorkspaceWorkbook } from "../src/main/workspaceWorkbook";
import { EMPTY_TIERING_DATA, domainReferenceMessages, renameIndependentTieringDomain, resolveTieringFeature } from "../src/shared/tiering";
import type { DictionaryItem, Requirement, TieringFeature, WorkspaceData } from "../src/shared/types";

const environment = vi.hoisted(() => ({ directory: "" }));
vi.mock("electron", () => ({ app: { getPath: () => environment.directory, getAppPath: () => process.cwd() } }));
const domains: DictionaryItem[] = [
  { id: "a", name: "原领域", level: "L0", sortOrder: 0, active: true },
  { id: "b", name: "新领域", level: "L0", sortOrder: 1, active: true },
  { id: "a1", name: "原标签", level: "L1", parentId: "a", sortOrder: 2, active: true },
  { id: "b1", name: "新标签", level: "L1", parentId: "b", sortOrder: 3, active: true },
];
const timestamp = "2026-01-01T00:00:00.000Z";
const requirement: Requirement = {
  id: "r1", title: "健康需求", description: "保留用户完整描述 → 旧路径", domainL0Id: "a", domainId: "a1", source: "健康基础", scope: "全球", allowLatestVersionUpgrade: true, overseasRegions: [], category: "平台演进", targetMonth: "2027-03", productIds: [], images: [], deviceWorkloadPm: 1, appWorkloadPm: 2, cloudWorkloadPm: 3, unallocatedWorkloadPm: 0, workloadPm: 6, createdAt: timestamp, updatedAt: timestamp,
};
const independent: TieringFeature = { id: "F002", l0: "原领域", l1: "原标签", requirementTitle: "独立规格", sortOrder: 0, createdAt: timestamp, updatedAt: timestamp };
function data(requirements = [requirement]): WorkspaceData {
  return { requirements, domains, products: [], members: [], groupOverrides: [], pptAccessPolicy: { entries: [], updatedAt: timestamp }, tiering: structuredClone(EMPTY_TIERING_DATA) };
}
let database: LocalDatabase;
const opened: LocalDatabase[] = [];
function internals(db = database) {
  return db as unknown as { db: Database; persist: () => Promise<void>; upsertTieringForRequirement: (...args: unknown[]) => void };
}
async function open() {
  const db = new LocalDatabase();
  opened.push(db);
  await db.initialize();
  return db;
}
beforeEach(async () => {
  environment.directory = await mkdtemp(join(tmpdir(), "roadmap-domain-sync-"));
  Object.defineProperty(process, "resourcesPath", { configurable: true, value: environment.directory });
  database = await open();
  await database.importWorkspace(data(), "replace");
});
afterEach(async () => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
  for (const db of opened.splice(0)) internals(db).db.close();
  await rm(environment.directory, { recursive: true, force: true });
});

describe("领域与分层同步（真实 SQLite 事务）", () => {
  it.each(["REQ-r1", "F001"])("%s 需求改选领域立即同步，原标签及 L0 可删除", async (featureId) => {
    await database.importWorkspace(data([{ ...requirement, tieringFeatureId: featureId }]), "replace");
    await database.saveTieringSupport({ recordId: `${featureId}::PS01`, status: "支持" });
    const before = database.getSnapshot();
    await database.saveRequirement({ ...before.requirements[0], domainL0Id: "b", domainId: "b1", title: "修改后需求" });
    const after = database.getSnapshot();
    expect(after.tiering?.features[0]).toMatchObject({ id: featureId, l0: "新领域", l1: "新标签", requirementTitle: "修改后需求" });
    expect(after.tiering?.supportRecords).toEqual(before.tiering?.supportRecords);
    expect(after.tiering?.changeLogs).toEqual(before.tiering?.changeLogs);
    await database.deleteDomain("a1");
    await database.deleteDomain("a");
    expect(database.getSnapshot().domains.map((item) => item.id)).toEqual(["b", "b1"]);
  });

  it("字典改名与 L1 迁移原子更新关联需求及历史独立规格", async () => {
    const workspace = database.getWorkspaceData();
    workspace.tiering!.features.push(independent);
    await database.importWorkspace(workspace, "replace");
    const before = database.getSnapshot();
    await database.saveDomain({ id: "a", name: "原领域改名" });
    expect(database.getSnapshot().tiering?.features.every((item) => item.l0 === "原领域改名")).toBe(true);
    await database.saveDomain({ id: "a1", name: "迁移后标签", parentId: "b" });
    const after = database.getSnapshot();
    expect(after.requirements[0]).toEqual({ ...before.requirements[0], domainL0Id: "b" });
    expect(after.tiering?.features.every((item) => item.l0 === "新领域" && item.l1 === "迁移后标签")).toBe(true);
    expect(after.tiering?.supportRecords).toEqual(before.tiering?.supportRecords);
    expect(after.tiering?.changeLogs).toEqual(before.tiering?.changeLogs);
    await database.deleteDomain("a");
  });

  it("真实引用列举需求、成员、子级与独立规格名称和处理位置", async () => {
    await database.saveMember({ name: "规划甲", employeeNumber: "E01", microsoftAccount: "test@example.com", role: "规划", domainL1Ids: ["a1"] });
    const workspace = database.getWorkspaceData();
    workspace.tiering!.features.push(independent);
    await database.importWorkspace(workspace, "replace");
    await expect(database.deleteDomain("a1")).rejects.toThrow(/健康需求[\s\S]*分层分级[\s\S]*规划甲[\s\S]*人员管理/);
    await expect(database.deleteDomain("a")).rejects.toThrow(/子级领域 L1 1 项：原标签/);
    await expect(database.saveMember({ name: "无效", employeeNumber: "E02", microsoftAccount: "test2@example.com", domainL1Ids: ["missing"] })).rejects.toThrow("无效的领域 L1");
  });

  it("旧分层名称不作为有效关联需求的删除引用；按完整路径匹配独立规格", async () => {
    const moved = { ...requirement, domainL0Id: "b", domainId: "b1", tieringFeatureId: "F001" };
    const stale = { ...independent, id: "F001", requirementId: "r1" };
    expect(resolveTieringFeature(stale, [moved], domains).feature).toMatchObject({ id: "F001", l0: "新领域", l1: "新标签", requirementTitle: requirement.title });
    expect(domainReferenceMessages(domains[2], { domains, requirements: [moved], members: [], features: [stale] })).toEqual([]);
    expect(renameIndependentTieringDomain({ ...independent, l0: "另一条路径" }, domains[2], { ...domains[2], name: "更名" }, domains, [])).toMatchObject({ l1: "原标签" });
    const conflicting = resolveTieringFeature(stale, [{ ...moved, tieringFeatureId: "F002" }], domains);
    expect(conflicting.invalidLink).toBe(true);
    expect(conflicting.feature).toEqual(stale);
  });

  it("启动幂等修复旧领域缓存及 L0 ID，保留描述、日期、工作量、支持状态和日志", async () => {
    await database.saveTieringSupport({ recordId: "REQ-r1::PS01", status: "支持" });
    const before = database.getSnapshot();
    const stale = { ...before.tiering!.features[0], l0: "旧名", l1: "旧标签", requirementTitle: "旧标题" };
    internals().db.run("UPDATE tiering_features SET payload = ? WHERE id = ?", [JSON.stringify(stale), stale.id]);
    const movedDomain = { ...domains[2], parentId: "b" };
    internals().db.run("UPDATE domains SET payload = ? WHERE id = ?", [JSON.stringify(movedDomain), "a1"]);
    await internals().persist();
    const reopened = await open();
    const after = reopened.getSnapshot();
    expect(after.requirements[0]).toEqual({ ...before.requirements[0], domainL0Id: "b" });
    expect(after.tiering!.features[0]).toMatchObject({ l0: "新领域", l1: "原标签", requirementTitle: requirement.title });
    expect(after.tiering!.supportRecords).toEqual(before.tiering!.supportRecords);
    expect(after.tiering!.changeLogs).toEqual(before.tiering!.changeLogs);
    expect((await open()).getSnapshot()).toEqual(after);
  });

  it("无效关联保留原规格和支持状态，并返回警告数量", async () => {
    const before = database.getSnapshot();
    internals().db.run("DELETE FROM domains WHERE id = 'a1'");
    await internals().persist();
    const reopened = new LocalDatabase();
    opened.push(reopened);
    expect((await reopened.initialize()).skippedTieringRequirements).toBeGreaterThan(0);
    expect(reopened.getSnapshot().tiering).toEqual(before.tiering);
  });

  it("同步中途失败整笔回滚，磁盘和内存均不发生部分迁移", async () => {
    const before = database.getSnapshot();
    const disk = await readFile(join(environment.directory, "roadmap.db"));
    vi.spyOn(internals(), "upsertTieringForRequirement").mockImplementation(() => { throw new Error("模拟同步失败"); });
    await expect(database.saveDomain({ id: "a1", name: "不应生效", parentId: "b" })).rejects.toThrow("模拟同步失败");
    expect(database.getSnapshot()).toEqual(before);
    expect(await readFile(join(environment.directory, "roadmap.db"))).toEqual(disk);
  });

  it("持久化失败恢复内存与原磁盘数据", async () => {
    const before = database.getSnapshot();
    const disk = await readFile(join(environment.directory, "roadmap.db"));
    vi.spyOn(internals(), "persist").mockRejectedValue(new Error("模拟磁盘失败"));
    await expect(database.saveDomain({ id: "a1", name: "不应保存", parentId: "b" })).rejects.toThrow("模拟磁盘失败");
    expect(database.getSnapshot()).toEqual(before);
    expect(await readFile(join(environment.directory, "roadmap.db"))).toEqual(disk);
  });

  it("数据包导入同步失败回滚，不清空既有人员、支持状态或需求", async () => {
    const before = database.getSnapshot();
    const incoming = database.getWorkspaceData();
    incoming.requirements[0].title = "导入的新标题";
    vi.spyOn(internals(), "upsertTieringForRequirement").mockImplementation(() => { throw new Error("导入同步失败"); });
    await expect(database.importWorkspace(incoming, "replace")).rejects.toThrow("导入同步失败");
    expect(database.getSnapshot()).toEqual(before);
  });

  it.each(["REQ-r1", "F001"])("Excel 显式删除 %s：普通规格删除，历史规格只解除关联", async (featureId) => {
    await database.importWorkspace(data([{ ...requirement, tieringFeatureId: featureId }]), "replace");
    const workspace = database.getWorkspaceData();
    const zip = await JSZip.loadAsync(await createWorkspaceWorkbook(workspace, "0.8.9", timestamp));
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    zip.file("xl/worksheets/sheet1.xml", sheet.replace(/<c r="A2"[^>]*>[\s\S]*?<\/c>/, '<c r="A2" t="inlineStr"><is><t>删除</t></is></c>'));
    const prepared = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "删除.xlsx", workspace);
    await database.importWorkspaceWorkbook(prepared, "excel-wins");
    const after = database.getSnapshot();
    expect(after.requirements).toHaveLength(0);
    expect(after.tiering?.features).toHaveLength(featureId === "F001" ? 1 : 0);
    expect(after.tiering?.supportRecords).toHaveLength(featureId === "F001" ? 9 : 0);
    expect(after.tiering?.features[0]?.requirementId).toBeUndefined();
  });

  it("完整包合并与 Excel v6 回导按当前需求刷新分层，保留支持历史", async () => {
    await database.saveTieringSupport({ recordId: "REQ-r1::PS01", status: "支持" });
    const before = database.getSnapshot();
    const incoming = database.getWorkspaceData();
    incoming.requirements[0] = { ...incoming.requirements[0], domainL0Id: "b", domainId: "b1", updatedAt: "2099-01-01T00:00:00.000Z" };
    await database.importWorkspace(incoming, "merge");
    expect(database.getSnapshot().tiering?.features[0]).toMatchObject({ l0: "新领域", l1: "新标签" });
    const workspace = database.getWorkspaceData();
    const zip = await JSZip.loadAsync(await createWorkspaceWorkbook(workspace, "0.8.9", timestamp, { start: "2027H1", end: "2027H2" }));
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    // The workbook writer uses inline business strings in this export; replace only D2/E2.
    const editCell = (xml: string, address: string, value: string) => xml.replace(new RegExp(`<c r="${address}"[^>]*>[\\s\\S]*?</c>`), `<c r="${address}" t="inlineStr"><is><t>${value}</t></is></c>`);
    zip.file("xl/worksheets/sheet1.xml", editCell(editCell(sheet, "D2", "原领域"), "E2", "原标签"));
    const prepared = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "领域修改.xlsx", workspace);
    expect(prepared.preview.errors).toEqual([]);
    await database.importWorkspaceWorkbook(prepared, "excel-wins");
    expect(database.getSnapshot().tiering?.features[0]).toMatchObject({ l0: "原领域", l1: "原标签" });
    expect(database.getSnapshot().tiering?.supportRecords).toEqual(before.tiering?.supportRecords);
    expect(database.getSnapshot().tiering?.changeLogs).toEqual(before.tiering?.changeLogs);
  });
});
