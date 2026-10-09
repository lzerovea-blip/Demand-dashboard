import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, readFile, readdir, rm, mkdir, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { Database } from "sql.js";
import JSZip from "jszip";
import { LocalDatabase } from "../src/main/database";
import { createWorkspacePackage, inspectWorkspacePackage } from "../src/main/workspacePackage";
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
  it("v0.9.0 新需求、独立医疗计划及解除匹配持久化；可生成隔离工作台验收数据", async () => {
    const input = { ...requirement, id: undefined, source: undefined } as SaveRequirementInput;
    await database.saveRequirement({ ...input, title: "批量匹配演示需求" });
    const normal = database.getSnapshot().requirements.find((r)=>r.title==="批量匹配演示需求")!;
    expect(database.getSnapshot().roadmapAssignments!.some((a)=>a.requirementId===normal.id)).toBe(false);
    await database.saveRoadmapAssignments({ requirementIds:[normal.id], roadmapId:"运动", laneId:"运动高阶" });
    await database.saveRequirement({ ...input, title:"医疗认证独立排期演示", targetMonth:"2027-06" });
    const medical = database.getSnapshot().requirements.find((r)=>r.title==="医疗认证独立排期演示")!;
    await database.saveRoadmapAssignments({ requirementIds:[medical.id], roadmapId:"医疗认证", laneId:"CE认证" });
    await database.saveMedicalRoadmapPlan({ requirementId:medical.id, roadmapId:"医疗认证", registrationStartMonth:"2027-01", submissionMonth:"2027-02", clinicalStartMonth:"2027-04", clinicalCompleteMonth:"2027-07", certificateMonth:"2028-03", updatedAt:timestamp });
    await database.saveRequirement({ ...input, title:"新录入尚未匹配的需求" });
    const before = database.getSnapshot();
    const reopened = await open(); expect(reopened.getSnapshot()).toEqual(before);
    const qa = process.env.ROADMAP_V090_QA_DIR;
    if (qa) { await mkdir(qa,{recursive:true}); await writeFile(join(qa,"roadmap.db"),Buffer.from(internals().db.export())); }
  });
  it("v8 逐项勾选在同一事务写入需求标题与系列状态，未选字段保持本机", async () => {
    const bytes = await createCurrentWorkspaceWorkbook(database.getWorkspaceData(), "0.8.26", timestamp, undefined, "v8");
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
  it("升级前的 v10 删除操作不会因为来源字段移除而产生假冲突，回导同步清理匹配", async () => {
    const original = await createCurrentWorkspaceWorkbook(data(), "0.8.31", timestamp, undefined, "v10");
    const zip = await JSZip.loadAsync(original); const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    zip.file("xl/worksheets/sheet1.xml", xml.replace(/<c r="B2"([^>]*)>[\s\S]*?<\/c>/, (_all, attrs: string) => `<c r="B2" ${/s="\d+"/.exec(attrs)?.[0] ?? ""} t="inlineStr"><is><t>删除</t></is></c>`));
    const prepared = await inspectWorkspaceWorkbook(await zip.generateAsync({type:"nodebuffer"}),"升级前v10.xlsx",database.getWorkspaceData());
    expect(prepared.preview.errors).toEqual([]); expect(prepared.preview.fieldChanges?.[0]).toMatchObject({ kind:"delete", conflict:false });
    await database.importWorkspaceWorkbook(prepared,"local-wins",{selectedChangeIds:["r1:delete"],conflictDecisions:{}});
    expect(database.getSnapshot().roadmapAssignments).toEqual([]); expect(database.getSnapshot().requirements).toEqual([]);
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
    await database.saveMedicalRoadmapPlanOrder({ roadmapId: "医疗认证", requirementIds: ["c", "b", "a"] });
    const reopened = await open();
    expect([...reopened.getSnapshot().medicalRoadmapPlans!].sort((a, b) => a.sortOrder! - b.sortOrder!).map((item) => item.requirementId)).toEqual(["c", "b", "a"]);
    const edited = reopened.getSnapshot().requirements.find((item) => item.id === "c")!;
    await reopened.saveRequirement({ ...edited, title: "改名仍保留顺序" } as SaveRequirementInput);
    expect(reopened.getSnapshot().medicalRoadmapPlans!.find((item) => item.requirementId === "c")?.sortOrder).toBe(0);
    await reopened.duplicateRequirement("c");
    expect(reopened.getSnapshot().requirements.find((item) => !["a", "b", "c"].includes(item.id))?.medicalRoadmapSortOrder).toBeUndefined();
    const before = reopened.getSnapshot();
    await expect(reopened.saveMedicalRoadmapPlanOrder({ roadmapId: "医疗认证", requirementIds: ["a", "a"] })).rejects.toThrow(/集合已变化/);
    await expect(reopened.saveMedicalRoadmapPlanOrder({ roadmapId: "医疗认证", requirementIds: ["a", "missing"] })).rejects.toThrow(/集合已变化/);
    const persist = vi.spyOn(internals(reopened), "persist").mockRejectedValueOnce(new Error("模拟落盘失败"));
    await expect(reopened.saveMedicalRoadmapPlanOrder({ roadmapId: "医疗认证", requirementIds: ["a", "b", "c"] })).rejects.toThrow("模拟落盘失败");
    persist.mockRestore();
    expect(reopened.getSnapshot()).toEqual(before);
    // Exercise SQL transaction rollback separately from a persistence failure.
    const run = internals(reopened).db.run.bind(internals(reopened).db);
    let writes = 0;
    const spy = vi.spyOn(internals(reopened).db, "run").mockImplementation((sql, params) => {
      if (String(sql).startsWith("INSERT INTO medical_roadmap_plans") && ++writes === 2) throw new Error("模拟事务失败");
      return run(sql, params);
    });
    await expect(reopened.saveMedicalRoadmapPlanOrder({ roadmapId: "医疗认证", requirementIds: ["a", "b", "c"] })).rejects.toThrow("模拟事务失败");
    spy.mockRestore();
    expect(reopened.getSnapshot()).toEqual(before);
  });

  it("旧十四来源需求均进入分层；复制生成分层但不复制路标归属", async () => {
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
    expect(snapshot.tiering?.features).toHaveLength(requirements.length);
    expect(snapshot.tiering?.supportRecords).toHaveLength(requirements.length * 15);
    expect(snapshot.requirements.every((item) => Boolean(item.tieringFeatureId))).toBe(true);
    const excludedRequirement = snapshot.requirements.find((item) => item.title === "行业需求")!;
    const duplicated = await database.duplicateRequirement(excludedRequirement.id);
    expect(duplicated.tiering?.features).toHaveLength(requirements.length + 1);
    const copy = duplicated.requirements.find((item) => item.title.startsWith("行业需求 副本"))!;
    expect(copy.tieringFeatureId).toBeDefined();
    expect(duplicated.roadmapAssignments!.some((item) => item.requirementId === copy.id)).toBe(false);
  });

  it("路标归属调整不改变需求分层、支持明细、待核验或变更记录", async () => {
    await database.saveTieringSupport({ recordId: "REQ-r1::PS01", status: "支持", reason: "确认支持" });
    const issue = {
      id: "issue-r1", featureId: "REQ-r1", seriesIds: ["PS01"], issueType: "待核验", description: "核验",
      priority: "P1", status: "待处理", sourceIds: [], owner: "", createdAt: timestamp, targetDate: "", updatedAt: timestamp, resolution: "",
    };
    internals().db.run("INSERT INTO tiering_issues (id, updated_at, payload) VALUES (?, ?, ?)", [issue.id, issue.updatedAt, JSON.stringify(issue)]);
    const before = database.getSnapshot();
    const excluded = await database.saveRoadmapAssignments({ requirementIds: ["r1"], roadmapId: "行业", laneId: "行业" });
    expect(excluded.requirements).toEqual(before.requirements);
    expect(excluded.tiering).toEqual(before.tiering);
    const restored = await database.saveRoadmapAssignments({ requirementIds: ["r1"], roadmapId: "AI", laneId: "AI综合" });
    expect(restored.requirements[0].tieringFeatureId).toBe("REQ-r1");
    expect(restored.tiering?.features).toHaveLength(1);
    expect(restored.tiering?.supportRecords).toHaveLength(15);
    expect(restored.tiering?.supportRecords.find((item) => item.id === "REQ-r1::PS01")?.status).toBe("支持");
    expect(restored.tiering?.verificationIssues).toHaveLength(1);
    expect(restored.tiering?.changeLogs).toHaveLength(1);
  });

  it("历史 F001–F156 在关联需求移出范围时只解除关联，保留历史规格和支持矩阵", async () => {
    await database.importWorkspace(data([{ ...requirement, tieringFeatureId: "F001" }]), "replace");
    await database.saveTieringSupport({ recordId: "F001::PS01", status: "支持", reason: "历史核验" });
    const moved = await database.saveRoadmapAssignments({ requirementIds: ["r1"], roadmapId: "health平台", laneId: "health平台" });
    expect(moved.requirements[0].tieringFeatureId).toBe("F001");
    expect(moved.tiering?.features).toContainEqual(expect.objectContaining({ id: "F001", requirementId: "r1" }));
    expect(moved.tiering?.supportRecords).toHaveLength(15);
    expect(moved.tiering?.supportRecords.find((item) => item.id === "F001::PS01")?.status).toBe("支持");
    expect(moved.tiering?.changeLogs).toHaveLength(1);
  });

  it("v12 启动为全部旧需求保留分层并迁移来源，重复运行不重建归属", async () => {
    const current = database.getSnapshot().requirements[0];
    const legacy = { ...current, source: "全场景" as const, tieringFeatureId: "REQ-r1" };
    internals().db.run("UPDATE requirements SET payload = ? WHERE id = ?", [JSON.stringify(legacy), legacy.id]);
    internals().db.run("UPDATE meta SET value = '10' WHERE key = 'schema_version'");
    internals().db.run("DELETE FROM settings WHERE key='roadmap_management_initialized'");
    await internals().persist();
    const upgraded = await open();
    expect(upgraded.getSnapshot().requirements[0]).toMatchObject({ id: "r1", tieringFeatureId: "REQ-r1" });
    expect(upgraded.getSnapshot().requirements[0].source).toBeUndefined();
    expect(upgraded.getSnapshot().roadmapAssignments).toContainEqual(expect.objectContaining({ roadmapId: "全场景", requirementId: "r1" }));
    expect(internals(upgraded).db.exec("SELECT value FROM meta WHERE key = 'schema_version'")[0].values[0][0]).toBe("12");
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
    expect(imported.tiering?.supportRecords.filter((item) => item.featureId === "F002")).toHaveLength(15);
  });

  it("旧 Excel 的来源修改被忽略，分层与路标匹配均保持", async () => {
    await database.saveTieringSupport({ recordId: "REQ-r1::PS01", status: "支持", reason: "回导前状态" });
    const workspace = database.getWorkspaceData();
    const zip = await JSZip.loadAsync(await createWorkspaceWorkbook(workspace, "0.8.24", timestamp));
    const sheet = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    zip.file("xl/worksheets/sheet1.xml", sheet.replace(/<c r="F2"[^>]*>[\s\S]*?<\/c>/, '<c r="F2" t="inlineStr"><is><t>行业</t></is></c>'));
    const prepared = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "来源移出范围.xlsx", workspace);
    expect(prepared.preview.errors).toEqual([]);
    const imported = await database.importWorkspaceWorkbook(prepared, "excel-wins");
    expect(imported.requirements[0].source).toBeUndefined();
    expect(imported.requirements[0].tieringFeatureId).toBe("REQ-r1");
    expect(imported.roadmapAssignments).toEqual(workspace.roadmapAssignments);
    expect(imported.tiering?.supportRecords.find((item)=>item.id === "REQ-r1::PS01")?.status).toBe("支持");
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
    expect(snapshot.tiering?.productSeries.map((item) => item.name)).toEqual(["WATCH数字系列", "WATCH Pro系列", "GT系列", "GT Pro系列", "ULTIMATE系列", "FIT系列", "FIT Pro系列", "D系列", "WATCH Buds系列", "Band系列", "Band Pro系列", "儿童手表", "体脂秤", "戒指系列", "无屏系列"]);
    expect(snapshot.tiering?.supportRecords).toHaveLength(15);
    expect(snapshot.tiering?.supportRecords.filter((item) => ["PS10", "PS11", "PS12", "PS13"].includes(item.seriesId)).every((item) => item.status === "未判断")).toBe(true);
    expect((await readdir(join(environment.directory, "backups"))).some((name) => name.startsWith("roadmap-before-v0.9.0-"))).toBe(true);
  });

  it("旧库启动补齐 WATCH Pro 与 Band Pro，不继承原系列结论且重复启动不产生副本", async () => {
    await database.saveTieringSupport({ recordId: "REQ-r1::PS01", status: "支持", reason: "已有 WATCH 结论" });
    await database.saveTieringSupport({ recordId: "REQ-r1::PS07", status: "不支持", reason: "已有 Band 结论" });
    const originalLogs = database.getSnapshot().tiering!.changeLogs.length;
    internals().db.run("DELETE FROM tiering_support_records WHERE id LIKE '%::PS14' OR id LIKE '%::PS15'");
    internals().db.run("DELETE FROM tiering_product_series WHERE id IN ('PS14','PS15')");
    await internals().persist();

    const upgraded = await open();
    const tiering = upgraded.getSnapshot().tiering!;
    expect(tiering.productSeries).toHaveLength(15);
    expect(tiering.supportRecords).toHaveLength(15);
    expect(tiering.supportRecords.find((item) => item.id === "REQ-r1::PS01")?.status).toBe("支持");
    expect(tiering.supportRecords.find((item) => item.id === "REQ-r1::PS07")?.status).toBe("不支持");
    expect(tiering.supportRecords.filter((item) => ["PS14", "PS15"].includes(item.seriesId)).map((item) => item.status)).toEqual(["未判断", "未判断"]);
    expect(tiering.changeLogs).toHaveLength(originalLogs);

    const reopened = await open();
    expect(reopened.getSnapshot().tiering?.productSeries).toHaveLength(15);
    expect(reopened.getSnapshot().tiering?.supportRecords).toHaveLength(15);
  });

  it("历史需求删除及旧库残留关联修复后，不再误报关联无效", async () => {
    await database.importWorkspace(data([{ ...requirement, tieringFeatureId: "F001" }]), "replace");
    await database.saveTieringSupport({ recordId: "F001::PS01", status: "支持", reason: "保留历史结论" });
    const before = database.getSnapshot().tiering!;
    const oldFeature = before.features.find((item) => item.id === "F001")!;
    await database.deleteRequirement(requirement.id);
    expect(database.getSnapshot().tiering?.features.find((item) => item.id === "F001")?.requirementId).toBeUndefined();
    internals().db.run("UPDATE tiering_features SET payload = ? WHERE id = ?", [JSON.stringify(oldFeature), oldFeature.id]);
    await internals().persist();

    const repaired = await open();
    const tiering = repaired.getSnapshot().tiering!;
    expect(tiering.features.find((item) => item.id === oldFeature.id)?.requirementId).toBeUndefined();
    expect(tiering.supportRecords.find((item) => item.id === "F001::PS01")?.status).toBe("支持");
    expect(tiering.changeLogs).toEqual(before.changeLogs);
    expect(resolveTieringFeature(tiering.features[0], repaired.getSnapshot().requirements, repaired.getSnapshot().domains).invalidLink).toBe(false);
    const restarted = await open();
    expect(restarted.getSnapshot().tiering?.features.find((item) => item.id === "F001")?.requirementId).toBeUndefined();
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

  it("医疗节点独立保存，需求编辑、复制和解除匹配均不篡改节点", async () => {
    await database.saveRoadmapAssignments({ requirementIds: ["r1"], roadmapId: "医疗认证", laneId: "CE认证" });
    const plan = { requirementId: "r1", roadmapId: "医疗认证", submissionMonth: "2028-05", registrationStartMonth: "2027-01", clinicalStartMonth: "2026-04", certificateMonth: "2025-08", updatedAt: timestamp };
    await database.saveMedicalRoadmapPlan(plan);
    const before = database.getSnapshot().medicalRoadmapPlans;
    await database.saveRequirement({ ...database.getSnapshot().requirements[0], targetMonth: "2026-12" } as SaveRequirementInput);
    expect(database.getSnapshot().medicalRoadmapPlans).toEqual(before);
    const copied = await database.duplicateRequirement("r1");
    const copy = copied.requirements.find((item)=>item.id!=="r1")!;
    expect(copied.roadmapAssignments!.some((a)=>a.requirementId===copy.id)).toBe(false);
    expect(copied.medicalRoadmapPlans).toEqual(before);
    await database.saveRoadmapAssignments({ requirementIds: ["r1"] });
    expect((await open()).getSnapshot().medicalRoadmapPlans).toEqual(before);
    await database.saveRoadmapAssignments({ requirementIds: ["r1"], roadmapId: "医疗认证", laneId: "NMPA认证" });
    expect(database.getSnapshot().medicalRoadmapPlans).toEqual(before);
  });

  it("历史空日期编辑和复制兼容，新建需求必填日期，非法医疗节点事务不写入", async () => {
    await database.importWorkspace(data([{ ...requirement, source: "CE认证", targetMonth: "", medicalSubmissionMilestones: { certificateMonth: "2025-01" } }]), "replace");
    await database.saveRequirement({ ...database.getSnapshot().requirements[0], title: "旧需求编辑" } as SaveRequirementInput);
    await database.duplicateRequirement("r1");
    const reopened = await open();
    expect(reopened.getSnapshot().requirements.every((item) => item.targetMonth === "")).toBe(true);
    expect(reopened.getSnapshot().tiering?.features).toHaveLength(2);
    expect(reopened.getSnapshot().medicalRoadmapPlans).toHaveLength(1);
    const before = reopened.getSnapshot();
    await expect(reopened.saveRequirement({ ...requirement, id: undefined, targetMonth: "" } as SaveRequirementInput)).rejects.toThrow("新需求必须填写上线年月");
    await expect(reopened.saveMedicalRoadmapPlan({ requirementId: "r1", roadmapId: "医疗认证", submissionMonth: "2027-00", updatedAt: timestamp })).rejects.toThrow("有效年月");
    expect(reopened.getSnapshot()).toEqual(before);
  });
  it("遗留来源参数不能清空普通需求日期或改变路标匹配", async () => {
    const before = database.getSnapshot();
    await expect(database.saveRequirement({ ...before.requirements[0], source:"CE认证", targetMonth:"" } as SaveRequirementInput)).rejects.toThrow("上线年月");
    expect(database.getSnapshot()).toEqual(before);
  });

  it("路标字典和批量匹配保护引用，保存失败及无效批次完整回滚", async () => {
    await database.saveRoadmap({ name: "试验路标", kind: "standard", sortOrder: 0 });
    const map = database.getSnapshot().roadmaps!.find((r)=>r.name==="试验路标")!;
    await database.saveRoadmapLane({ roadmapId: map.id, name: "规划泳道", sortOrder: 0 });
    const lane = database.getSnapshot().roadmapLanes!.find((r)=>r.roadmapId===map.id)!;
    const beforeRequirement = database.getSnapshot().requirements[0];
    await database.saveRoadmapAssignments({ requirementIds: ["r1"], roadmapId: map.id, laneId: lane.id });
    expect(database.getSnapshot().requirements[0]).toEqual(beforeRequirement);
    await expect(database.deleteRoadmap(map.id)).rejects.toThrow("仍有需求匹配");
    await expect(database.deleteRoadmapLane(lane.id)).rejects.toThrow("仍被需求使用");
    await expect(database.saveRoadmap({ ...map, kind: "medical" })).rejects.toThrow("不能更改");
    const before = database.getSnapshot();
    await expect(database.saveRoadmapAssignments({ requirementIds: ["r1","missing"], roadmapId: "健康", laneId: "健康基础" })).rejects.toThrow("集合已变化");
    const persist = vi.spyOn(internals(),"persist").mockRejectedValueOnce(new Error("模拟落盘失败"));
    await expect(database.saveRoadmapAssignments({ requirementIds:["r1"] })).rejects.toThrow("模拟落盘失败"); persist.mockRestore();
    expect(database.getSnapshot()).toEqual(before);
    await database.saveRoadmapAssignments({ requirementIds:["r1"] });
    await expect(database.saveRoadmap({ ...map, kind:"medical" })).rejects.toThrow("不能更改");
    await database.deleteRoadmapLane(lane.id); await database.deleteRoadmap(map.id);
    expect((await open()).getSnapshot().roadmaps!.some((r)=>r.id===map.id)).toBe(false);
  });
  it("新分组键的卡片排序可以随完整包往返，导出不会出现空卡片标题", async () => {
    const key = "健康::健康基础::a1::2027-03";
    await database.saveGroupOrder({ source:"健康基础",roadmapId:"健康",laneId:"健康基础",targetMonth:"2027-03",groupKeys:[key] });
    expect(database.getSnapshot().groupOverrides[0]).toMatchObject({ groupKey:key,cardTitle:"原标签",sortOrder:0 });
    const archive = await createWorkspacePackage(database.getWorkspaceData(),"0.9.0");
    const imported = await inspectWorkspacePackage(archive,"排序.roadmap");
    expect(imported.data.groupOverrides[0]).toMatchObject({ groupKey:key,cardTitle:"原标签",sortOrder:0 });
  });

  it("删除无引用路标或泳道时一并清理旧卡片排序，完整包仍可导出", async () => {
    await database.saveRoadmap({ name: "临时路标", kind: "standard" });
    const roadmap = database.getSnapshot().roadmaps!.find((item) => item.name === "临时路标")!;
    await database.saveRoadmapLane({ roadmapId: roadmap.id, name: "临时泳道" });
    const lane = database.getSnapshot().roadmapLanes!.find((item) => item.roadmapId === roadmap.id)!;
    await database.saveRoadmapAssignments({ requirementIds: ["r1"], roadmapId: roadmap.id, laneId: lane.id });
    const groupKey = `${roadmap.id}::${lane.id}::a1::2027-03`;
    await database.saveGroupOrder({ source: "", roadmapId: roadmap.id, laneId: lane.id, targetMonth: "2027-03", groupKeys: [groupKey] });
    expect(database.getSnapshot().groupOverrides.some((item) => item.groupKey === groupKey)).toBe(true);
    await database.saveRoadmapAssignments({ requirementIds: ["r1"] });
    await database.deleteRoadmapLane(lane.id);
    expect(database.getSnapshot().groupOverrides.some((item) => item.groupKey === groupKey)).toBe(false);
    await database.deleteRoadmap(roadmap.id);
    const archive = await createWorkspacePackage(database.getWorkspaceData(), "0.9.0");
    expect((await inspectWorkspacePackage(archive, "删除后.roadmap")).data.groupOverrides).toEqual([]);
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
    expect(after.tiering?.supportRecords).toHaveLength(featureId === "F001" ? 15 : 0);
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
