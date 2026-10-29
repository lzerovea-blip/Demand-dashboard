import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { createHash } from "node:crypto";
import type { AppSnapshot, Requirement, WorkspaceData } from "../src/shared/types";
import { normalizeRoadmapManagement, validateRoadmapManagement, assignmentGroupKey } from "../src/shared/roadmapManagement";
import { normalizeWorkspaceTieringScope } from "../src/shared/tiering";
import { buildHalfYearSummaries, buildRoadmapGroups, medicalRoadmapTimelineMonths } from "../src/shared/roadmap";
import { mergeWorkspaceData } from "../src/shared/workspace";
import { diffWorkspaceData } from "../src/shared/workspaceDiff";
import { createWorkspacePackage, inspectWorkspacePackage } from "../src/main/workspacePackage";
import { createWorkspaceWorkbook, inspectWorkspaceWorkbook, applyPreparedWorkspaceWorkbook } from "../src/main/workspaceWorkbook";
import { buildPptExportPlan } from "../src/shared/pptExport";
import { createRoadmapPresentation, presentationSlideCount } from "../src/main/pptxExport";
import { validateGeneratedPptx } from "../src/main/pptxValidation";

const time = "2026-10-09T00:00:00.000Z";
const requirement: Requirement = { id: "one", title: "完整需求标题、组合不拆分", description: "需求内容", images: [], domainL0Id: "l0", domainId: "l1", scope: "全球", allowLatestVersionUpgrade: true, overseasRegions: [], category: "平台演进", targetMonth: "2027-03", productIds: [], deviceWorkloadPm: 1, appWorkloadPm: 2, cloudWorkloadPm: 3, researchAppWorkloadPm: 4, algorithmWorkloadPm: 5, unallocatedWorkloadPm: 0, workloadPm: 15, createdAt: time, updatedAt: time };
function data(): WorkspaceData {
  return { requirements: [requirement, { ...requirement, id: "two", title: "未匹配需求" }], domains: [{ id: "l0", name: "健康领域", level: "L0", sortOrder: 0, active: true }, { id: "l1", name: "睡眠", level: "L1", parentId: "l0", sortOrder: 0, active: true }], products: [], members: [], pptAccessPolicy: { entries: [], updatedAt: time }, groupOverrides: [], roadmaps: [{ id: "custom", name: "自定义健康规划", kind: "standard", sortOrder: 0, updatedAt: time }, { id: "clinical", name: "国际认证规划", kind: "medical", sortOrder: 1, updatedAt: time }], roadmapLanes: [{ id: "lane", roadmapId: "custom", name: "重点迭代", sortOrder: 0, updatedAt: time }, { id: "ce", roadmapId: "clinical", name: "CE认证", sortOrder: 0, updatedAt: time }], roadmapAssignments: [{ requirementId: "one", roadmapId: "custom", laneId: "lane", updatedAt: time }], roadmapAssignmentTombstones: [], medicalRoadmapPlans: [] };
}
describe("v0.9.0 独立路标模型", () => {
  it("旧来源迁移七路标十四泳道及医疗独立日期，保留旧卡片顺序且幂等", () => {
    const legacy = { ...data(), roadmaps: undefined, roadmapLanes: undefined, roadmapAssignments: undefined, medicalRoadmapPlans: undefined, requirements: [{ ...requirement, source: "健康高阶" as const }, { ...requirement, id: "two", source: "CE认证" as const, targetMonth: "", medicalSubmissionMilestones: { certificateMonth: "2025-01", registrationStartMonth: "2029-01" }, medicalRoadmapSortOrder: 4 }], groupOverrides: [{ groupKey: "l1::健康高阶::2027-03", cardTitle: "睡眠", cardSummary: "", sortOrder: 2, updatedAt: time }] };
    const migrated = normalizeRoadmapManagement(legacy);
    expect(migrated.roadmaps).toHaveLength(7); expect(migrated.roadmapLanes).toHaveLength(14);
    expect(migrated.roadmapAssignments[0]).toMatchObject({ requirementId: "one", roadmapId: "健康", laneId: "健康高阶" });
    expect(migrated.medicalRoadmapPlans[0]).toMatchObject({ requirementId: "two", submissionMonth: undefined, certificateMonth: "2025-01", registrationStartMonth: "2029-01", sortOrder: 4 });
    expect(migrated.groupOverrides[0].groupKey).toBe("健康::健康高阶::l1::2027-03");
    expect(normalizeRoadmapManagement(migrated)).toEqual(migrated);
  });
  it("显式未匹配不会被遗留来源重新匹配；一条需求只能一个泳道", () => {
    const state = { ...data(), requirements: [{ ...requirement, source: "健康基础" as const }], roadmapAssignments: [] };
    expect(normalizeRoadmapManagement(state).roadmapAssignments).toEqual([]);
    const invalid = { ...data(), roadmapAssignments: [...data().roadmapAssignments!, { requirementId: "one", roadmapId: "clinical", laneId: "ce", updatedAt: time }] };
    expect(() => validateRoadmapManagement(invalid)).toThrow("重复");
    expect(() => validateRoadmapManagement({ ...data(), roadmapAssignments: [{ requirementId: "one", roadmapId: "clinical", laneId: "lane", updatedAt: time }] })).toThrow("匹配无效");
  });
  it("路标分组依据独立匹配；全部需求分层且汇总包含未匹配", () => {
    const state = normalizeWorkspaceTieringScope(data());
    expect(state.tiering!.features).toHaveLength(2); expect(state.tiering!.supportRecords).toHaveLength(30);
    const groups = buildRoadmapGroups({ ...state, templates: [] }, "2027H1", "2027H1");
    expect(groups).toHaveLength(1); expect(groups[0]).toMatchObject({ roadmapName: "自定义健康规划", laneName: "重点迭代", key: assignmentGroupKey("custom", "lane", "l1", "2027-03") });
    const summary = buildHalfYearSummaries(state.requirements, "2027H1", "2027H1", state)[0];
    expect(summary.requirementCount).toBe(2); expect(summary.byTrack).toMatchObject({ custom: 6, 未匹配: 6 });
  });
  it("解除记录阻止旧匹配恢复；新匹配能恢复；旧来源不能覆盖本机手动归属", () => {
    const original = data();
    const local = { ...original, roadmapAssignments: [], roadmapAssignmentTombstones: [{ requirementId: "one", removedAt: "2026-10-10T00:00:00.000Z" }] };
    expect(mergeWorkspaceData(local, original).roadmapAssignments).toEqual([]);
    expect(mergeWorkspaceData(local, { ...original, roadmapAssignments: [{ ...original.roadmapAssignments![0], updatedAt: "2026-10-11T00:00:00.000Z" }] }).roadmapAssignments).toHaveLength(1);
    const legacy = { ...data(), roadmaps: undefined, roadmapLanes: undefined, roadmapAssignments: undefined, requirements: [{ ...requirement, source: "运动基础" as const, updatedAt: "2026-10-12T00:00:00.000Z" }] };
    expect(mergeWorkspaceData(original, legacy).roadmapAssignments).toEqual(original.roadmapAssignments);
  });
  it("同名路标泳道跨设备重映射，医疗草稿保留，类型冲突明确拒绝", () => {
    const remote = { ...data(), roadmaps: data().roadmaps!.map((r) => ({ ...r, id: `remote-${r.id}` })), roadmapLanes: data().roadmapLanes!.map((r) => ({ ...r, id: `remote-${r.id}`, roadmapId: `remote-${r.roadmapId}` })), roadmapAssignments: [{ requirementId: "two", roadmapId: "remote-custom", laneId: "remote-lane", updatedAt: time }] };
    const merged = mergeWorkspaceData(data(), remote);
    expect(merged.roadmaps).toHaveLength(2); expect(merged.roadmapLanes).toHaveLength(2); expect(merged.roadmapAssignments).toContainEqual({ requirementId: "two", roadmapId: "custom", laneId: "lane", updatedAt: time });
    expect(() => mergeWorkspaceData(data(), { ...remote, roadmaps: remote.roadmaps.map((r) => r.name === "自定义健康规划" ? { ...r, kind: "medical" as const } : r) })).toThrow("展示类型冲突");
  });
  it("新增与解除匹配的预览显示路标、泳道名称而非内部编号", () => {
    const state = data(); const empty = { ...state,roadmapAssignments:[] };
    const added = diffWorkspaceData(empty,state).find((item)=>item.entity==="路标匹配")!;
    expect(added.fields).toEqual([{field:"路标",currentValue:"未匹配",incomingValue:"自定义健康规划"},{field:"泳道",currentValue:"—",incomingValue:"重点迭代"}]);
    const removed = diffWorkspaceData(state,empty).find((item)=>item.entity==="路标匹配")!;
    expect(removed.title).toBe(requirement.title); expect(removed.fields[0].incomingValue).toBe("未匹配");
  });
  it("医疗时间与通用上线月独立，任意节点命中时间窗，所有空节点仍保留", () => {
    const state = { ...data(), roadmapAssignments: [{ requirementId: "one", roadmapId: "clinical", laneId: "ce", updatedAt: time }], medicalRoadmapPlans: [{ requirementId: "one", roadmapId: "clinical", submissionMonth: "2028-02", registrationStartMonth: "2027-04", certificateMonth: "2026-01", updatedAt: time }] };
    validateRoadmapManagement(state);
    const groups = buildRoadmapGroups({ ...state, templates: [] }, "2027H1", "2027H1", { medicalDemandRows: true, medicalRange: "selected" });
    expect(groups).toHaveLength(1); expect(groups[0].requirements[0].targetMonth).toBe("2028-02"); expect(state.requirements[0].targetMonth).toBe("2027-03");
    expect(medicalRoadmapTimelineMonths(groups, "2027H1", "2027H1")).toMatchObject({ 0: "2027-01" });
    expect(buildRoadmapGroups({ ...state, medicalRoadmapPlans: [], templates: [] }, "2027H1", "2027H1", { medicalDemandRows: true, medicalRange: "selected" })).toHaveLength(1);
    expect(() => validateRoadmapManagement({ ...state, medicalRoadmapPlans: [{ ...state.medicalRoadmapPlans[0], submissionMonth: "2027-13" }] })).toThrow("有效年月");
  });
  it("新包往返字典、匹配、医疗配置及解除记录", async () => {
    const state = normalizeWorkspaceTieringScope(data());
    const bytes = await createWorkspacePackage(state, "0.9.0", time);
    expect(JSON.parse(await (await JSZip.loadAsync(bytes)).file("manifest.json")!.async("string")).formatVersion).toBe(11);
    const result = await inspectWorkspacePackage(bytes, "v10.roadmap");
    expect(result.data.roadmapAssignments).toEqual(state.roadmapAssignments); expect(result.data.roadmaps).toEqual(normalizeRoadmapManagement(state).roadmaps);
    expect(result.data.requirements.every((r) => r.source === undefined)).toBe(true);
  });
  it("真实 v9 包来源迁移，已有归属不被旧包覆盖；新包缺少匹配结构拒绝", async () => {
    const archive = await JSZip.loadAsync(await createWorkspacePackage(data(), "0.9.0", time));
    const packed = JSON.parse(await archive.file("data.json")!.async("string"));
    for (const key of ["roadmaps","roadmapLanes","roadmapAssignments","roadmapAssignmentTombstones","medicalRoadmapPlans"]) delete packed[key];
    packed.requirements[0].source = "健康基础"; packed.requirements[1].source = "CE认证";
    packed.requirements[1].medicalSubmissionMilestones = { registrationStartMonth:"2028-01", certificateMonth:"2026-02" };
    const bytes = Buffer.from(JSON.stringify(packed)); archive.file("data.json",bytes);
    const manifest = JSON.parse(await archive.file("manifest.json")!.async("string")); manifest.formatVersion=9;
    for (const key of ["roadmaps","roadmapLanes","roadmapAssignments","roadmapAssignmentTombstones","medicalRoadmapPlans"]) delete manifest.counts[key];
    manifest.files["data.json"] = { sizeBytes:bytes.length,sha256:createHash("sha256").update(bytes).digest("hex") }; archive.file("manifest.json",JSON.stringify(manifest));
    const old = await inspectWorkspacePackage(await archive.generateAsync({type:"nodebuffer"}),"旧版v9.roadmap");
    expect(old.data.roadmaps).toHaveLength(7); expect(old.data.roadmapLanes).toHaveLength(14);
    expect(old.data.medicalRoadmapPlans?.[0]).toMatchObject({ requirementId:"two",submissionMonth:"2027-03",registrationStartMonth:"2028-01",certificateMonth:"2026-02" });
    expect(mergeWorkspaceData(data(),old.data).roadmapAssignments).toEqual(data().roadmapAssignments);
    manifest.formatVersion=10; archive.file("manifest.json",JSON.stringify(manifest));
    await expect(inspectWorkspacePackage(await archive.generateAsync({type:"nodebuffer"}),"缺结构v10.roadmap")).rejects.toThrow("缺少完整");
  });
  it("协作 v11 可修改全部需求系列；逐项回导和删除保持路标数据一致", async () => {
    const state = normalizeWorkspaceTieringScope(data());
    const original = await createWorkspaceWorkbook(state, "0.9.0", time, undefined, "v11");
    const book = new ExcelJS.Workbook(); await book.xlsx.load(original as never);
    expect(book.getWorksheet("需求清单")!.getCell("AF3").text).toBe("未判断");
    const sourceZip = await JSZip.loadAsync(original);
    const sheet = await sourceZip.file("xl/worksheets/sheet1.xml")!.async("string");
    const edit = (xml: string, address: string, value: string) => xml.replace(new RegExp(`<c r="${address}"([^>]*)>[\\s\\S]*?<\\/c>`), (_all, attributes: string) => `<c r="${address}" ${/s="\d+"/.exec(attributes)?.[0] ?? ""} t="inlineStr"><is><t>${value}</t></is></c>`);
    sourceZip.file("xl/worksheets/sheet1.xml", edit(edit(sheet,"C2","编辑新标题"),"R3","支持"));
    const preview = await inspectWorkspaceWorkbook(await sourceZip.generateAsync({ type: "nodebuffer" }), "v11.xlsx", state);
    expect(preview.preview.formatVersion).toBe("v11"); expect(preview.preview.errors).toEqual([]);
    const result = applyPreparedWorkspaceWorkbook(state, preview, "local-wins", time, { selectedChangeIds: ["one:field:title", "two:series:PS01"], conflictDecisions: {} });
    expect(result.requirements[0].title).toBe("编辑新标题"); expect(result.roadmapAssignments).toEqual(state.roadmapAssignments);
    expect(result.tiering!.supportRecords.find((r)=>r.id === "REQ-two::PS01")!.status).toBe("支持");
  });
  it("动态 PPT 选项、空页及四泳道续页使用字典并保持原生可编辑", async () => {
    const state = { ...data(), roadmapLanes: Array.from({ length: 4 }, (_,index) => ({ id: index === 0 ? "lane" : `lane${index}`, roadmapId: "custom", name: `自定义泳道${index+1}`, sortOrder: index, updatedAt: time })) };
    const snapshot: AppSnapshot = { ...state, templates: [] };
    const request = { start: "2027H1", end: "2027H2", materialName: "v0.9.0 模块拆分验证", sections: ["summary", "custom", "clinical"] };
    const plan = buildPptExportPlan(snapshot, request);
    expect(plan.roadmapPages).toHaveLength(4); expect(plan.requirements).toHaveLength(2);
    const ppt = await createRoadmapPresentation(snapshot, request);
    expect(presentationSlideCount(plan)).toBe(8); await validateGeneratedPptx(ppt, 8);
    const zip = await JSZip.loadAsync(ppt); const xml = (await Promise.all(Object.keys(zip.files).filter((path)=>/^ppt\/slides\/slide\d+\.xml$/.test(path)).map((path)=>zip.file(path)!.async("string")))).join("");
    expect(xml).toContain("自定义健康规划"); expect(xml).toContain("自定义泳道4"); expect(xml).toContain("国际认证规划"); expect(xml).toContain("当前区间暂无");
    const qa = process.env.ROADMAP_V090_QA_DIR;
    if (qa) { await mkdir(qa,{recursive:true}); const workspace = normalizeWorkspaceTieringScope(state); await writeFile(join(qa,"v0.9.0-动态路标样稿.pptx"),ppt); await writeFile(join(qa,"v0.9.0-需求池协作.xlsx"),await createWorkspaceWorkbook(workspace,"0.9.0",time)); await writeFile(join(qa,"v0.9.0-模块迁移验证.roadmap"),await createWorkspacePackage(workspace,"0.9.0",time)); }
  });
});
