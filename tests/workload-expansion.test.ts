import { describe, expect, it } from "vitest";
import { createHash } from "node:crypto";
import JSZip from "jszip";
import { createWorkspacePackage, inspectWorkspacePackage } from "../src/main/workspacePackage";
import { cloneRequirementAsCopy, normalizeRequirement } from "../src/shared/requirements";
import { mergeWorkspaceData } from "../src/shared/workspace";
import { buildHalfYearSummaries, buildRoadmapGroups } from "../src/shared/roadmap";
import { isWorkloadPendingEvaluation, sumWorkloadBreakdown } from "../src/shared/workload";
import { compareMedicalRoadmapRequirements, moveMedicalRoadmapRow, reorderedMedicalRequirements } from "../src/shared/medicalRoadmapOrder";
import { filterMembers, normalizeMember, ownerField } from "../src/shared/members";
import type { AppSnapshot, Requirement, WorkspaceData } from "../src/shared/types";

const time = "2026-09-09T00:00:00.000Z";
const row: Requirement = {
  id: "r1", title: "研究与算法需求", description: "保留原业务描述", domainL0Id: "l0", domainId: "l1", source: "健康基础", scope: "全球", category: "平台演进", targetMonth: "2027-03", allowLatestVersionUpgrade: false, overseasRegions: [], images: [], productIds: [],
  deviceWorkloadPm: 1, appWorkloadPm: 2, cloudWorkloadPm: 3, researchAppWorkloadPm: 4, algorithmWorkloadPm: 5, unallocatedWorkloadPm: 0, workloadPm: 15, createdAt: time, updatedAt: time,
};
const workspace: WorkspaceData = {
  requirements: [row], domains: [{ id: "l0", name: "健康", level: "L0", active: true, sortOrder: 0 }, { id: "l1", name: "研究", level: "L1", parentId: "l0", active: true, sortOrder: 1 }], products: [], groupOverrides: [], members: [], pptAccessPolicy: { entries: [], updatedAt: time },
};

async function rewritePackage(bytes: Buffer, edit: (data: any, manifest: any) => void): Promise<Buffer> {
  const zip = await JSZip.loadAsync(bytes);
  const manifest = JSON.parse(await zip.file("manifest.json")!.async("string"));
  const data = JSON.parse(await zip.file("data.json")!.async("string"));
  edit(data, manifest);
  const body = Buffer.from(JSON.stringify(data));
  zip.file("data.json", body);
  manifest.files["data.json"] = { sizeBytes: body.length, sha256: createHash("sha256").update(body).digest("hex") };
  zip.file("manifest.json", JSON.stringify(manifest));
  return zip.generateAsync({ type: "nodebuffer" });
}

describe("五项工作量与兼容", () => {
  it("历史补零不拆分原 App；新工作量不被误归入待拆分", () => {
    const legacy = { ...row, researchAppWorkloadPm: undefined, algorithmWorkloadPm: undefined, workloadPm: 6 };
    expect(normalizeRequirement(legacy)).toMatchObject({ appWorkloadPm: 2, researchAppWorkloadPm: 0, algorithmWorkloadPm: 0, workloadPm: 6 });
    expect(normalizeRequirement({ ...row, unallocatedWorkloadPm: undefined } as unknown as Requirement)).toMatchObject({ unallocatedWorkloadPm: 0, workloadPm: 15 });
    expect(sumWorkloadBreakdown([row, { ...row, researchAppWorkloadPm: 0.11, algorithmWorkloadPm: 0.22, unallocatedWorkloadPm: 1 }])).toMatchObject({ researchApp: 4.11, algorithm: 5.22, unallocated: 1, total: 22.33 });
    const onlyNew = { ...row, deviceWorkloadPm: 0, appWorkloadPm: 0, cloudWorkloadPm: 0 };
    expect(isWorkloadPendingEvaluation(onlyNew)).toBe(false);
    expect(isWorkloadPendingEvaluation({ ...onlyNew, researchAppWorkloadPm: 0, algorithmWorkloadPm: 0 })).toBe(true);
  });

  it("半年总量和卡片合计包含五项，所有原三项图表保持原口径", () => {
    const summary = buildHalfYearSummaries([row], "2027H1", "2027H1")[0];
    expect(summary).toMatchObject({ totalWorkload: 15, researchAppWorkload: 4, algorithmWorkload: 5, bySide: { device: 1, app: 2, cloud: 3 } });
    expect(summary.byCategory.平台演进).toBe(6);
    expect(summary.byTrack.健康).toBe(6);
    expect(summary.categoryBySide.平台演进).toEqual({ device: 1, app: 2, cloud: 3 });
    expect(buildRoadmapGroups({ ...workspace, templates: [] } as AppSnapshot, "2027H1", "2027H1")[0].totalWorkloadPm).toBe(15);
  });

  it("复制保留工作量但不继承医疗排序", () => {
    const copy = cloneRequirementAsCopy({ ...row, source: "CE认证", targetMonth: "", medicalRoadmapSortOrder: 0 }, [], () => "copy", time);
    expect(copy).toMatchObject({ researchAppWorkloadPm: 4, algorithmWorkloadPm: 5, targetMonth: "" });
    expect(copy.medicalRoadmapSortOrder).toBeUndefined();
  });

  it("v7 往返新增工作量、算法岗位、白名单和医疗顺序，并预览明确清零", async () => {
    const member = normalizeMember({ id: "algo", name: "算法甲", employeeNumber: "A1", microsoftAccount: "algo@example.com", role: "算法", domainL1Ids: ["l1"], createdAt: time, updatedAt: time, sortOrder: 0 });
    const data = { ...workspace, requirements: [{ ...row, source: "CE认证" as const, medicalRoadmapSortOrder: 2 }], members: [member], pptAccessPolicy: { entries: [{ memberId: "algo", permission: "read" as const }], updatedAt: time } };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(data, "0.8.17"), "v7.roadmap");
    expect(inspected.data.requirements[0]).toMatchObject({ researchAppWorkloadPm: 4, algorithmWorkloadPm: 5, medicalRoadmapSortOrder: 2 });
    expect(inspected.data.members[0].role).toBe("算法");
    const cleared = { ...data, requirements: [{ ...data.requirements[0], researchAppWorkloadPm: 0, algorithmWorkloadPm: 0, workloadPm: 6, updatedAt: "2026-09-10T00:00:00.000Z" }] };
    const preview = await inspectWorkspacePackage(await createWorkspacePackage(cleared, "0.8.17"), "清零.roadmap", data);
    expect(preview.preview.differences[0].fields.map((item) => item.field)).toEqual(expect.arrayContaining(["创新研究 App 工作量", "算法工作量"]));
  });

  it.each([1, 2, 3, 4, 5, 6])("v%i 缺少新增字段：合并保留本机，替换规范化补零", async (version) => {
    const bytes = await rewritePackage(await createWorkspacePackage(workspace, "0.8.17"), (data, manifest) => {
      manifest.formatVersion = version;
      delete data.requirements[0].researchAppWorkloadPm;
      delete data.requirements[0].algorithmWorkloadPm;
      Object.assign(data.requirements[0], { workloadPm: 6, title: "旧版修改标题", updatedAt: "2026-09-10T00:00:00.000Z" });
    });
    const imported = await inspectWorkspacePackage(bytes, "旧版.roadmap", workspace);
    expect(mergeWorkspaceData(workspace, imported.data).requirements[0]).toMatchObject({ title: "旧版修改标题", researchAppWorkloadPm: 4, algorithmWorkloadPm: 5, workloadPm: 15 });
    expect(normalizeRequirement(imported.data.requirements[0])).toMatchObject({ researchAppWorkloadPm: 0, algorithmWorkloadPm: 0, workloadPm: 6 });
    expect(imported.preview.differences[0].fields.map((item) => item.field)).not.toContain("算法工作量");
  });

  it.each([-1, null, "2", "NaN"])("校验值有效也拒绝非法新增工作量 %s，不静默归零", async (value) => {
    const bytes = await rewritePackage(await createWorkspacePackage(workspace, "0.8.17"), (data) => { data.requirements[0].researchAppWorkloadPm = value; });
    await expect(inspectWorkspacePackage(bytes, "非法.roadmap")).rejects.toThrow(/工作量/);
  });

  it("算法参与人员岗位和领域筛选，并提供需求责任人映射", () => {
    const member = normalizeMember({ id: "a", name: "算法", employeeNumber: "A1", microsoftAccount: "a@example.com", role: "算法", domainL1Ids: ["l1"], sortOrder: 0, createdAt: time, updatedAt: time });
    expect(member.role).toBe("算法");
    expect(filterMembers([member], workspace.domains, { role: "算法", domainL0Id: "l0", domainL1Id: "" })).toEqual([member]);
    expect(ownerField("算法")).toBe("algorithmMemberId");
  });
});

describe("医疗泳道独立排序", () => {
  const rows = ["a", "b", "c", "d"].map((id, medicalRoadmapSortOrder) => ({ ...row, id, source: "CE认证" as const, medicalRoadmapSortOrder }));
  it("拖动插入前后、同一行不变，过滤后的移动保留隐藏项位置", () => {
    expect(moveMedicalRoadmapRow(["a", "b", "c"], "a", "b", "after")).toEqual(["b", "a", "c"]);
    expect(moveMedicalRoadmapRow(["a", "b", "c"], "c", "a", "before")).toEqual(["c", "a", "b"]);
    expect(moveMedicalRoadmapRow(["a", "b"], "a", "a", "after")).toEqual(["a", "b"]);
    expect(reorderedMedicalRequirements(rows, ["d", "b"]).map((item) => item.id)).toEqual(["a", "d", "c", "b"]);
    expect([...rows, { ...row, id: "new", source: "NMPA认证" as const, targetMonth: "" }].sort(compareMedicalRoadmapRequirements).at(-1)?.id).toBe("new");
  });
  it("未知、重复、非医疗需求或不足两项时拒绝写入", () => {
    for (const ids of [["a"], ["a", "a"], ["a", "missing"], ["a", "r1"]]) expect(() => reorderedMedicalRequirements([...rows, row], ids)).toThrow();
  });
});
