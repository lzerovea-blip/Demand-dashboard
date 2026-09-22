import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, readFile, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { Database } from "sql.js";
import JSZip from "jszip";
import { LocalDatabase } from "../src/main/database";
import { createWorkspaceWorkbook as createCurrentWorkspaceWorkbook, inspectWorkspaceWorkbook } from "../src/main/workspaceWorkbook";
import { EMPTY_TIERING_DATA, domainReferenceMessages, renameIndependentTieringDomain, resolveTieringFeature } from "../src/shared/tiering";
import type { DictionaryItem, Requirement, SaveRequirementInput, Member, TieringFeature, WorkspaceData } from "../src/shared/types";

const createWorkspaceWorkbook = (workspace: WorkspaceData, appVersion: string, exportedAt?: string, range?: { start: string; end: string }) =>
  createCurrentWorkspaceWorkbook(workspace, appVersion, exportedAt, range, "v7");

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
  researchAppWorkloadPm: 0, algorithmWorkloadPm: 0,
  planningMemberId: "plan", productManagerMemberId: "pm", seMemberId: "se",
  id: "r1", title: "健康需求", description: "保留用户完整描述 → 旧路径", domainL0Id: "a", domainId: "a1", source: "健康基础", scope: "全球", allowLatestVersionUpgrade: true, overseasRegions: [], category: "平台演进", targetMonth: "2027-03", productIds: [], images: [], deviceWorkloadPm: 1, appWorkloadPm: 2, cloudWorkloadPm: 3, unallocatedWorkloadPm: 0, workloadPm: 6, createdAt: timestamp, updatedAt: timestamp,
};
const independent: TieringFeature = { id: "F002", l0: "原领域", l1: "原标签", requirementTitle: "独立规格", sortOrder: 0, createdAt: timestamp, updatedAt: timestamp };
function data(requirements = [requirement]): WorkspaceData {
  const members: Member[] = (["规划", "产品经理", "SE"] as const).map((role, index) => ({ id: ["plan", "pm", "se"][index], name: role, role, employeeNumber: `OWN-${index}`, microsoftAccount: `owner${index}@example.com`, domainL1Ids: [], sortOrder: index, createdAt: timestamp, updatedAt: timestamp }));
  return { requirements, domains, products: [], members, groupOverrides: [], pptAccessPolicy: { entries: [], updatedAt: timestamp }, tiering: structuredClone(EMPTY_TIERING_DATA) };
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
  it("v8 逐项勾选在同一事务写入需求标题与系列状态，未选字段保持本机", async () => {
    const bytes = await createCurrentWorkspaceWorkbook(database.getWorkspaceData(), "0.8.26", timestamp);
    const zip = await JSZip.loadAsync(bytes);
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    zip.file("xl/worksheets/sheet1.xml", sheet
      .replace(/<c r="B2"[^>]*>[\s\S]*?<\/c>/, '<c r="B2" t="inlineStr"><is><t>表格标题</t></is></c>')
      .replace(/<c r="Q2"[^>]*>[\s\S]*?<\/c>/, '<c r="Q2" t="inlineStr"><is><t>支持</t></is></c>'));
    const prepared = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "v8.xlsx", database.getWorkspaceData());
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((item) => item.id)).toEqual(["r1:field:title", "r1:series:PS01"]);
    await database.importWorkspaceWorkbook(prepared, "local-wins", { selectedChangeIds: ["r1:series:PS01"], conflictDecisions: {} });
    expect(database.getSnapshot().requirements[0].title).toBe("健康需求");
    expect(database.getSnapshot().tiering!.supportRecords.find((item) => item.seriesId === "PS01")?.status).toBe("支持");
    expect(database.getSnapshot().tiering!.changeLogs).toHaveLength(1);
    const reopened = await open();
    expect(reopened.getSnapshot().tiering!.supportRecords.find((item) => item.seriesId === "PS01")?.status).toBe("支持");
  });
  it("删除记录跨重启持久化，旧包不复活需求，较新修改可以恢复", async () => {
    await database.deleteRequirement("r1");
    const deleted = database.getWorkspaceData();
    expect(deleted.requirements).toEqual([]);
    expect(deleted.requirementTombstones).toEqual([expect.objectContaining({ id: "r1", title: "健康需求" })]);

    const reopened = await open();
    expect(reopened.getWorkspaceData().requirementTombstones).toEqual(deleted.requirementTombstones);
    await reopened.importWorkspace(data([{ ...requirement, updatedAt: timestamp }]), "merge");
    expect(reopened.getSnapshot().requirements).toEqual([]);

    await reopened.importWorkspace(data([{ ...requirement, title: "对方后续修改", updatedAt: "2099-01-01T00:00:00.000Z" }]), "merge");
    expect(reopened.getSnapshot().requirements).toContainEqual(expect.objectContaining({ id: "r1", title: "对方后续修改" }));
    expect(reopened.getWorkspaceData().requirementTombstones).toEqual([]);
  });

  it("整体替换和 Excel 显式删除均生成删除记录", async () => {
    await database.importWorkspace({ ...data(), requirements: [] }, "replace");
    expect(database.getWorkspaceData().requirementTombstones).toContainEqual(expect.objectContaining({ id: "r1", title: "健康需求" }));

    await database.importWorkspace(data(), "replace");
    const workbook = await JSZip.loadAsync(await createWorkspaceWorkbook(database.getWorkspaceData(), "0.8.23", timestamp));
    const sheet = await workbook.file("xl/worksheets/sheet1.xml")!.async("string");
    workbook.file("xl/worksheets/sheet1.xml", sheet.replace(/<c r="A2"[^>]*>[\s\S]*?<\/c>/, '<c r="A2" t="inlineStr"><is><t>删除</t></is></c>'));
    const prepared = await inspectWorkspaceWorkbook(await workbook.generateAsync({ type: "nodebuffer" }), "显式删除.xlsx", database.getWorkspaceData());
    await database.importWorkspaceWorkbook(prepared, "excel-wins");
    expect(database.getWorkspaceData().requirementTombstones).toContainEqual(expect.objectContaining({ id: "r1", title: "健康需求" }));
  });

  it("算法责任保存、复制、重启和 Excel 更新不丢失；已引用人员不可删除或改岗", async () => {
    await database.saveMember({ name: "算法甲", employeeNumber: "ALG-1", microsoftAccount: "alg1@example.com", role: "算法", domainL1Ids: [] });
    const member = database.getSnapshot().members.find(item => item.employeeNumber === "ALG-1")!;
    await database.saveRequirement({ ...requirement, algorithmMemberId: member.id } as SaveRequirementInput);
    await expect(database.deleteMember(member.id)).rejects.toThrow();
    await expect(database.saveMember({ ...member, role: "开发" })).rejects.toThrow();
    await expect(database.saveRequirement({ ...requirement, algorithmMemberId: "pm" } as SaveRequirementInput)).rejects.toThrow("岗位角色不匹配");
    expect(database.getSnapshot().requirements[0].algorithmMemberId).toBe(member.id);
    const workspace = database.getWorkspaceData();
    const zip = await JSZip.loadAsync(await createWorkspaceWorkbook(workspace, "0.8.21", timestamp));
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    zip.file("xl/worksheets/sheet1.xml", sheet.replace(/<c r="B2"[^>]*>[\s\S]*?<\/c>/, '<c r="B2" t="inlineStr"><is><t>表格修改标题</t></is></c>'));
    const prepared = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "算法保留.xlsx", workspace);
    await database.importWorkspaceWorkbook(prepared, "excel-wins");
    expect(database.getSnapshot().requirements[0]).toMatchObject({ title: "表格修改标题", algorithmMemberId: member.id });
    await database.duplicateRequirement("r1");
    const reopened = await open();
    expect(reopened.getSnapshot().requirements.every(item => item.algorithmMemberId === member.id)).toBe(true);
    for (const row of reopened.getSnapshot().requirements) await reopened.saveRequirement({ ...row, algorithmMemberId: "" } as SaveRequirementInput);
    await expect(reopened.deleteMember(member.id)).resolves.toBeDefined();
  });
  it("五项工作量保存、复制、重启，编辑清零及非法值回滚", async () => {
    const input = { ...requirement, researchAppWorkloadPm: 1.235, algorithmWorkloadPm: 2.345 } as SaveRequirementInput;
    await database.saveRequirement(input);
    expect(database.getSnapshot().requirements[0]).toMatchObject({ researchAppWorkloadPm: 1.24, algorithmWorkloadPm: 2.35, workloadPm: 9.59 });
    await database.duplicateRequirement(requirement.id);
    const reopened = await open();
    expect(reopened.getSnapshot().requirements.every((item) => item.researchAppWorkloadPm === 1.24 && item.algorithmWorkloadPm === 2.35 && item.workloadPm === 9.59)).toBe(true);
    const before = reopened.getSnapshot();
    for (const bad of [-1, NaN, Infinity]) await expect(reopened.saveRequirement({ ...input, algorithmWorkloadPm: bad })).rejects.toThrow(/工作量/);
    expect(reopened.getSnapshot()).toEqual(before);
    await reopened.saveRequirement({ ...input, researchAppWorkloadPm: 0, algorithmWorkloadPm: 0 });
    expect(reopened.getSnapshot().requirements.find((item) => item.id === requirement.id)?.workloadPm).toBe(6);
  });

  it("算法成员支持负责领域、岗位筛选所需数据及白名单，不新增需求责任人", async () => {
    const next = await database.saveMember({ name: "算法甲", employeeNumber: "ALG-01", microsoftAccount: "algorithm@example.com", role: "算法", domainL1Ids: ["a1"] });
    const member = next.members.find((item) => item.role === "算法")!;
    expect(member).toMatchObject({ name: "算法甲", domainL1Ids: ["a1"] });
    await database.savePptAccessPolicy({ entries: [{ memberId: member.id, permission: "read" }] });
    expect((await open()).getSnapshot().pptAccessPolicy.entries).toContainEqual({ memberId: member.id, permission: "read" });
  });

  it("医疗泳道排序跨来源、空日期持久化，编辑保留、复制追加且保存失败回滚", async () => {
    await database.importWorkspace(data(["a", "b", "c"].map((id, index) => ({ ...requirement, id, source: index === 1 ? "CE认证" : "NMPA认证", targetMonth: index === 2 ? "" : "2027-03" }))), "replace");
    await database.saveMedicalRoadmapOrder({ requirementIds: ["c", "b", "a"] });
    const reopened = await open();
    expect([...reopened.getSnapshot().requirements].sort((a, b) => a.medicalRoadmapSortOrder! - b.medicalRoadmapSortOrder!).map((item) => item.id)).toEqual(["c", "b", "a"]);
    const edited = reopened.getSnapshot().requirements.find((item) => item.id === "c")!;
    await reopened.saveRequirement({ ...edited, title: "改名仍保留顺序" } as SaveRequirementInput);
    expect(reopened.getSnapshot().requirements.find((item) => item.id === "c")?.medicalRoadmapSortOrder).toBe(0);
    await reopened.duplicateRequirement("c");
    expect(reopened.getSnapshot().requirements.find((item) => !["a", "b", "c"].includes(item.id))?.medicalRoadmapSortOrder).toBeUndefined();
    const before = reopened.getSnapshot();
    await expect(reopened.saveMedicalRoadmapOrder({ requirementIds: ["a", "a"] })).rejects.toThrow(/重复/);
    await expect(reopened.saveMedicalRoadmapOrder({ requirementIds: ["a", "missing"] })).rejects.toThrow(/刷新/);
    const persist = vi.spyOn(internals(reopened), "persist").mockRejectedValueOnce(new Error("模拟落盘失败"));
    await expect(reopened.saveMedicalRoadmapOrder({ requirementIds: ["a", "b", "c"] })).rejects.toThrow("模拟落盘失败");
    persist.mockRestore();
    expect(reopened.getSnapshot()).toEqual(before);
    // Exercise SQL transaction rollback separately from a persistence failure.
    const run = internals(reopened).db.run.bind(internals(reopened).db);
    let writes = 0;
    const spy = vi.spyOn(internals(reopened).db, "run").mockImplementation((sql, params) => {
      if (String(sql).startsWith("UPDATE requirements SET payload") && ++writes === 2) throw new Error("模拟事务失败");
      return run(sql, params);
    });
    await expect(reopened.saveMedicalRoadmapOrder({ requirementIds: ["a", "b", "c"] })).rejects.toThrow("模拟事务失败");
    spy.mockRestore();
    expect(reopened.getSnapshot()).toEqual(before);
  });

  it("仅九个运动、健康和 AI 来源生成分层记录，其他来源复制也不生成", async () => {
    const included = [
      "运动基础", "运动进阶", "运动高阶",
      "健康基础", "健康进阶", "健康高阶",
      "AI运动", "AI健康", "AI综合",
    ] as const;
    const excluded = ["行业", "全场景", "health平台", "NMPA认证", "CE认证"] as const;
    const requirements = [...included, ...excluded].map((source, index) => ({
      ...requirement,
      id: `scope-${index}`,
      title: `${source}需求`,
      source,
    }));
    await database.importWorkspace(data(requirements), "replace");
    const snapshot = database.getSnapshot();
    expect(snapshot.tiering?.features).toHaveLength(included.length);
    expect(snapshot.tiering?.supportRecords).toHaveLength(included.length * 13);
    expect(snapshot.requirements.filter((item) => included.includes(item.source as typeof included[number])).every((item) => Boolean(item.tieringFeatureId))).toBe(true);
    expect(snapshot.requirements.filter((item) => excluded.includes(item.source as typeof excluded[number])).every((item) => item.tieringFeatureId === undefined)).toBe(true);
    const excludedRequirement = snapshot.requirements.find((item) => item.source === "行业")!;
    const duplicated = await database.duplicateRequirement(excludedRequirement.id);
    expect(duplicated.tiering?.features).toHaveLength(included.length);
    expect(duplicated.requirements.find((item) => item.id !== excludedRequirement.id && item.title.startsWith("行业需求 副本"))?.tieringFeatureId).toBeUndefined();
  });

  it("来源跨范围修改原子清理普通规格、支持明细、待核验和变更记录，再改回时重建未判断矩阵", async () => {
    await database.saveTieringSupport({ recordId: "REQ-r1::PS01", status: "支持", reason: "确认支持" });
    const issue = {
      id: "issue-r1", featureId: "REQ-r1", seriesIds: ["PS01"], issueType: "待核验", description: "核验",
      priority: "P1", status: "待处理", sourceIds: [], owner: "", createdAt: timestamp, targetDate: "", updatedAt: timestamp, resolution: "",
    };
    internals().db.run("INSERT INTO tiering_issues (id, updated_at, payload) VALUES (?, ?, ?)", [issue.id, issue.updatedAt, JSON.stringify(issue)]);
    const excluded = await database.saveRequirement({ ...database.getSnapshot().requirements[0], source: "行业" } as SaveRequirementInput);
    expect(excluded.requirements[0].source).toBe("行业");
    expect(excluded.requirements[0]).not.toHaveProperty("tieringFeatureId");
    expect(excluded.tiering).toMatchObject({ features: [], supportRecords: [], verificationIssues: [], changeLogs: [] });

    const restored = await database.saveRequirement({ ...excluded.requirements[0], source: "AI综合" } as SaveRequirementInput);
    expect(restored.requirements[0].tieringFeatureId).toBe("REQ-r1");
    expect(restored.tiering?.features).toHaveLength(1);
    expect(restored.tiering?.supportRecords).toHaveLength(13);
    expect(restored.tiering?.supportRecords.every((item) => item.status === "未判断")).toBe(true);
    expect(restored.tiering?.verificationIssues).toEqual([]);
    expect(restored.tiering?.changeLogs).toEqual([]);
  });

  it("历史 F001–F156 在关联需求移出范围时只解除关联，保留历史规格和支持矩阵", async () => {
    await database.importWorkspace(data([{ ...requirement, tieringFeatureId: "F001" }]), "replace");
    await database.saveTieringSupport({ recordId: "F001::PS01", status: "支持", reason: "历史核验" });
    const moved = await database.saveRequirement({ ...database.getSnapshot().requirements[0], source: "health平台" } as SaveRequirementInput);
    expect(moved.requirements[0].tieringFeatureId).toBeUndefined();
    expect(moved.tiering?.features).toContainEqual(expect.objectContaining({ id: "F001", requirementId: undefined }));
    expect(moved.tiering?.supportRecords).toHaveLength(13);
    expect(moved.tiering?.supportRecords.find((item) => item.id === "F001::PS01")?.status).toBe("支持");
    expect(moved.tiering?.changeLogs).toHaveLength(1);
  });

  it("v11 启动迁移清理旧库中非目标路标的普通分层数据，保留需求本身", async () => {
    const current = database.getSnapshot().requirements[0];
    const legacy = { ...current, source: "全场景" as const, tieringFeatureId: "REQ-r1" };
    internals().db.run("UPDATE requirements SET payload = ? WHERE id = ?", [JSON.stringify(legacy), legacy.id]);
    internals().db.run("UPDATE meta SET value = '10' WHERE key = 'schema_version'");
    await internals().persist();
    const upgraded = await open();
    expect(upgraded.getSnapshot().requirements).toContainEqual(expect.objectContaining({ id: "r1", source: "全场景" }));
    expect(upgraded.getSnapshot().requirements[0]).not.toHaveProperty("tieringFeatureId");
    expect(upgraded.getSnapshot().tiering).toMatchObject({ features: [], supportRecords: [], verificationIssues: [], changeLogs: [] });
    expect(internals(upgraded).db.exec("SELECT value FROM meta WHERE key = 'schema_version'")[0].values[0][0]).toBe("11");
  });

  it("旧包中的普通孤立规格会清理，F001–F156 独立历史规格继续保留", async () => {
    const workspace = data();
    workspace.tiering = {
      ...structuredClone(EMPTY_TIERING_DATA),
      features: [
        { ...independent, id: "REQ-orphan", requirementTitle: "普通孤立规格" },
        { ...independent, id: "F002", requirementTitle: "历史独立规格" },
      ],
    };
    const imported = await database.importWorkspace(workspace, "replace");
    expect(imported.tiering?.features.map((item) => item.id).sort()).toEqual(["F002", "REQ-r1"]);
    expect(imported.tiering?.supportRecords.filter((item) => item.featureId === "REQ-orphan")).toEqual([]);
    expect(imported.tiering?.supportRecords.filter((item) => item.featureId === "F002")).toHaveLength(13);
  });

  it("Excel 回导将来源移出范围时，需求更新与分层清理同一事务生效", async () => {
    await database.saveTieringSupport({ recordId: "REQ-r1::PS01", status: "支持", reason: "回导前状态" });
    const workspace = database.getWorkspaceData();
    const zip = await JSZip.loadAsync(await createWorkspaceWorkbook(workspace, "0.8.24", timestamp));
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    zip.file("xl/worksheets/sheet1.xml", sheet.replace(/<c r="F2"[^>]*>[\s\S]*?<\/c>/, '<c r="F2" t="inlineStr"><is><t>行业</t></is></c>'));
    const prepared = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "来源移出范围.xlsx", workspace);
    expect(prepared.preview.errors).toEqual([]);
    const imported = await database.importWorkspaceWorkbook(prepared, "excel-wins");
    expect(imported.requirements[0].source).toBe("行业");
    expect(imported.requirements[0]).not.toHaveProperty("tieringFeatureId");
    expect(imported.tiering).toMatchObject({ features: [], supportRecords: [], verificationIssues: [], changeLogs: [] });
  });

  it("v7 启动升级前备份，唯一责任人补齐并为全部规格增加四个未判断系列", async () => {
    const raw = JSON.parse(internals().db.exec("SELECT payload FROM requirements WHERE id='r1'")[0].values[0][0] as string);
    delete raw.planningMemberId; delete raw.productManagerMemberId; delete raw.seMemberId;
    internals().db.run("UPDATE requirements SET payload=? WHERE id='r1'", [JSON.stringify(raw)]);
    for (const row of internals().db.exec("SELECT id,payload FROM members")[0].values) {
      const member = { ...JSON.parse(row[1] as string), domainL1Ids: ["a1"] };
      internals().db.run("UPDATE members SET payload=? WHERE id=?", [JSON.stringify(member), row[0]]);
    }
    internals().db.run("DELETE FROM tiering_support_records WHERE id LIKE '%::PS10' OR id LIKE '%::PS11' OR id LIKE '%::PS12' OR id LIKE '%::PS13'");
    internals().db.run("DELETE FROM tiering_product_series WHERE id IN ('PS10','PS11','PS12','PS13')");
    internals().db.run("UPDATE meta SET value='7' WHERE key='schema_version'");
    await internals().persist();
    const upgraded = await open();
    const snapshot = upgraded.getSnapshot();
    expect(snapshot.requirements[0]).toMatchObject({ planningMemberId: "plan", productManagerMemberId: "pm", seMemberId: "se" });
    expect(snapshot.tiering?.productSeries.map((item) => item.name)).toEqual(["WATCH数字系列", "GT系列", "GT Pro系列", "ULTIMATE系列", "FIT系列", "FIT Pro系列", "D系列", "WATCH Buds系列", "Band系列", "儿童手表", "体脂秤", "戒指系列", "无屏系列"]);
    expect(snapshot.tiering?.supportRecords).toHaveLength(13);
    expect(snapshot.tiering?.supportRecords.filter((item) => ["PS10", "PS11", "PS12", "PS13"].includes(item.seriesId)).every((item) => item.status === "未判断")).toBe(true);
    expect((await readdir(join(environment.directory, "backups"))).some((name) => name.startsWith("roadmap-before-v0.8.24-"))).toBe(true);
  });

  it("变更记录只编辑管理信息，并可永久删除且不改变系列状态", async () => {
    await database.saveTieringSupport({ recordId: "REQ-r1::PS01", status: "支持", reason: "首次确认" });
    const created = database.getSnapshot().tiering!.changeLogs[0];
    const updated = await database.updateTieringChangeLog({ logId: created.id, reason: "复核确认", reviewer: "审核甲", recordStatus: "已归档", note: "仅更新管理信息" });
    expect(updated.tiering!.changeLogs[0]).toMatchObject({ id: created.id, changedAt: created.changedAt, before: created.before, after: created.after, reason: "复核确认", reviewer: "审核甲", recordStatus: "已归档", note: "仅更新管理信息" });
    expect(updated.tiering!.changeLogs[0].updatedAt).not.toBe(created.updatedAt);
    expect(updated.tiering!.supportRecords.find((item) => item.id === "REQ-r1::PS01")?.status).toBe("支持");
    const deleted = await database.deleteTieringChangeLog(created.id);
    expect(deleted.tiering!.changeLogs).toHaveLength(0);
    expect(deleted.tiering!.supportRecords.find((item) => item.id === "REQ-r1::PS01")?.status).toBe("支持");
  });

  it("医疗认证需求保存、复制和切换普通来源时正确维护节点", async () => {
    const medical = await database.saveRequirement({
      ...database.getSnapshot().requirements[0],
      source: "NMPA认证",
      targetMonth: "2027-03",
      medicalSubmissionMilestones: { registrationStartMonth: "2027-01", clinicalStartMonth: "2027-04", certificateMonth: "2027-08" },
    } as SaveRequirementInput);
    expect(medical.requirements[0].medicalSubmissionMilestones).toEqual({ registrationStartMonth: "2027-01", clinicalStartMonth: "2027-04", certificateMonth: "2027-08" });
    const copied = await database.duplicateRequirement(medical.requirements[0].id);
    expect(copied.requirements.find((item) => item.id !== medical.requirements[0].id)?.medicalSubmissionMilestones).toEqual(medical.requirements[0].medicalSubmissionMilestones);
    await expect(database.saveRequirement({ ...medical.requirements[0], targetMonth: "2026-12" } as SaveRequirementInput)).resolves.toBeDefined();
    const cleared = await database.saveRequirement({ ...medical.requirements[0], source: "健康基础" } as SaveRequirementInput);
    expect(cleared.requirements.find((item) => item.id === medical.requirements[0].id)?.medicalSubmissionMilestones).toBeUndefined();
  });

  it("空送检版本保存、复制、重启不补日期；非法日期或切回普通来源失败不部分更新", async () => {
    const input = { ...requirement, source: "CE认证", targetMonth: "", medicalSubmissionMilestones: { registrationStartMonth: "2027-12", certificateMonth: "2025-01" } } as SaveRequirementInput;
    await database.saveRequirement(input);
    await database.duplicateRequirement(requirement.id);
    const reopened = await open();
    expect(reopened.getSnapshot().requirements).toHaveLength(2);
    expect(reopened.getSnapshot().requirements.every((item) => item.targetMonth === "" && item.medicalSubmissionMilestones?.certificateMonth === "2025-01")).toBe(true);
    expect(reopened.getSnapshot().tiering?.features).toHaveLength(0);
    expect(reopened.getSnapshot().requirements.every((item) => item.tieringFeatureId === undefined)).toBe(true);
    const before = reopened.getSnapshot();
    await expect(reopened.saveRequirement({ ...input, source: "健康基础" })).rejects.toThrow(/上线年月/);
    await expect(reopened.saveRequirement({ ...input, targetMonth: "2027-00" })).rejects.toThrow(/送检版本/);
    expect(reopened.getSnapshot()).toEqual(before);
    await reopened.saveRequirement({ ...input, targetMonth: "2028-03" });
    await reopened.saveRequirement(input);
    expect(reopened.getSnapshot().requirements.find((item) => item.id === requirement.id)?.targetMonth).toBe("");
  });

  it("需求引用成员时禁止删除或更改岗位，改名不破坏责任人引用", async () => {
    const planning = database.getSnapshot().members.find((item) => item.id === "plan")!;
    await expect(database.deleteMember("plan")).rejects.toThrow(/仍被 1 条需求引用.*健康需求/);
    await expect(database.saveMember({ ...planning, role: "UX" })).rejects.toThrow(/更改岗位角色/);
    const renamed = await database.saveMember({ ...planning, name: "规划新姓名" });
    expect(renamed.requirements[0].planningMemberId).toBe("plan");
    expect(renamed.members.find((item) => item.id === "plan")?.name).toBe("规划新姓名");
  });

  it("旧包缺少责任人字段时，替换导入也保留本机分配和所需成员", async () => {
    const incoming = data().requirements.map((item) => {
      const legacy = { ...item };
      delete legacy.planningMemberId; delete legacy.productManagerMemberId; delete legacy.seMemberId;
      return legacy;
    });
    const legacyWorkspace = { ...data(incoming), members: [] };
    const snapshot = await database.importWorkspace(legacyWorkspace, "replace");
    expect(snapshot.requirements[0]).toMatchObject({ planningMemberId: "plan", productManagerMemberId: "pm", seMemberId: "se" });
    expect(snapshot.members.map((item) => item.id).sort()).toEqual(["plan", "pm", "se"]);
  });
  it.each(["REQ-r1", "F001"])("%s 需求改选领域立即同步，原标签及 L0 可删除", async (featureId) => {
    await database.importWorkspace(data([{ ...requirement, tieringFeatureId: featureId }]), "replace");
    await database.saveTieringSupport({ recordId: `${featureId}::PS01`, status: "支持" });
    const before = database.getSnapshot();
    await database.saveRequirement({ ...before.requirements[0], domainL0Id: "b", domainId: "b1", title: "修改后需求" } as SaveRequirementInput);
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

  it("需求来源移出分层范围但落盘失败时，需求与分层清理均恢复", async () => {
    const before = database.getSnapshot();
    const disk = await readFile(join(environment.directory, "roadmap.db"));
    const persist = vi.spyOn(internals(), "persist").mockRejectedValueOnce(new Error("模拟需求落盘失败"));
    await expect(database.saveRequirement({ ...before.requirements[0], source: "行业" } as SaveRequirementInput)).rejects.toThrow("模拟需求落盘失败");
    persist.mockRestore();
    expect(database.getSnapshot()).toEqual(before);
    expect(await readFile(join(environment.directory, "roadmap.db"))).toEqual(disk);
  });

  it("完整包范围规范化完成但落盘失败时，恢复导入前内存与磁盘", async () => {
    const before = database.getSnapshot();
    const disk = await readFile(join(environment.directory, "roadmap.db"));
    const incoming = database.getWorkspaceData();
    incoming.requirements[0] = { ...incoming.requirements[0], source: "全场景", updatedAt: "2099-01-01T00:00:00.000Z" };
    const persist = vi.spyOn(internals(), "persist").mockRejectedValueOnce(new Error("模拟导入落盘失败"));
    await expect(database.importWorkspace(incoming, "replace")).rejects.toThrow("模拟导入落盘失败");
    persist.mockRestore();
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
    expect(after.tiering?.supportRecords).toHaveLength(featureId === "F001" ? 13 : 0);
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
