import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { createHash } from "node:crypto";
import { createWorkspacePackage, inspectWorkspacePackage } from "../src/main/workspacePackage";
import type { RequirementImage, WorkspaceData } from "../src/shared/types";
import { BUILT_IN_TIERING_PRODUCT_SERIES, createTieringSupportRecord, EMPTY_TIERING_DATA } from "../src/shared/tiering";

const imageBytes = Buffer.from("offline-image-bytes", "utf8");
const image: RequirementImage = {
  id: "image-1",
  name: "睡眠截图.png",
  mimeType: "image/png",
  sizeBytes: imageBytes.length,
  dataUrl: `data:image/png;base64,${imageBytes.toString("base64")}`,
};

const workspace: WorkspaceData = {
  requirements: [{
    id: "r1",
    title: "社交时差",
    description: "评估周末和周内的睡眠规律性",
    images: [image],
    domainL0Id: "l0",
    domainId: "d1",
    source: "健康基础",
    scope: "全球",
    allowLatestVersionUpgrade: false,
    overseasRegions: [],
    category: "产品专属",
    targetMonth: "2027-03",
    productIds: ["p1"],
    deviceWorkloadPm: 0.25,
    appWorkloadPm: 0.5,
    cloudWorkloadPm: 0.25,
    researchAppWorkloadPm: 0,
    algorithmWorkloadPm: 0,
    unallocatedWorkloadPm: 0,
    workloadPm: 1,
    createdAt: "2026-08-19T00:00:00.000Z",
    updatedAt: "2026-08-20T00:00:00.000Z",
  }],
  domains: [
    { id: "l0", name: "健康", sortOrder: 0, active: true, level: "L0" },
    { id: "d1", name: "睡眠", sortOrder: 0, active: true, level: "L1", parentId: "l0" },
  ],
  products: [{ id: "p1", name: "手表", sortOrder: 0, active: true }],
  groupOverrides: [{ groupKey: "d1::健康基础::2027-03", cardTitle: "睡眠", cardSummary: "摘要", sortOrder: 2, updatedAt: "2026-08-20T00:00:00.000Z" }],
  members: [],
  pptAccessPolicy: { entries: [], updatedAt: "1970-01-01T00:00:00.000Z" },
};

describe(".roadmap 数据包", () => {
  it("v11 完整包保存待完善需求并在合并预览中列出", async () => {
    const draft = { ...workspace.requirements[0], id: "NEW-00000000-0000-4000-8000-000000000002", title: "协作待完善",
      category: "平台演进" as const, productIds: [], images: [], planningMemberId: "", productManagerMemberId: "", seMemberId: "",
      pendingSeries: { PS01: "支持" as const } };
    const bytes = await createWorkspacePackage({ ...workspace, draftRequirements: [draft] }, "0.9.2", "2026-10-10T00:00:00.000Z");
    const inspected = await inspectWorkspacePackage(bytes, "协作待完善.roadmap", { ...workspace, draftRequirements: [] });
    expect(inspected.data.draftRequirements?.map((item) => item.id)).toContain(draft.id);
    expect(inspected.preview.differences.some((item) => item.entity === "待完善需求" && item.key === draft.id)).toBe(true);
  });
  it("v6 往返空送检版本和任意倒序节点；日期清空列入合并差异", async () => {
    const medical: WorkspaceData = { ...workspace, requirements: workspace.requirements.map((item) => ({ ...item, source: "CE认证", targetMonth: "", medicalSubmissionMilestones: { registrationStartMonth: "2029-01", certificateMonth: "2025-01" }, updatedAt: "2026-09-09T00:00:00.000Z" })) };
    const archive = await createWorkspacePackage(medical, "0.8.16");
    const inspected = await inspectWorkspacePackage(archive, "医疗选填.roadmap", workspace);
    expect(inspected.data.requirements[0].targetMonth).toBe("");
    expect(inspected.data.medicalRoadmapPlans?.[0]).toMatchObject(medical.requirements[0].medicalSubmissionMilestones!);
    expect(inspected.preview.differences.some((diff) => diff.fields.some((field) => field.field === "上线月份"))).toBe(true);
    await expect(createWorkspacePackage({ ...medical, requirements: medical.requirements.map((item) => ({ ...item, source: "健康基础" })) }, "0.8.16")).resolves.toBeDefined();
    await expect(createWorkspacePackage({ ...medical, requirements: medical.requirements.map((item) => ({ ...item, targetMonth: "2027-13" })) }, "0.8.16")).rejects.toThrow(/送检时间|送检版本/);
  });

  it("完整往返需求、描述、图片、字典和卡片编辑", async () => {
    const archive = await createWorkspacePackage(workspace, "0.1.0", "2026-08-20T08:00:00.000Z");
    const zip = await JSZip.loadAsync(archive);
    expect(JSON.parse(await zip.file("manifest.json")!.async("string")).formatVersion).toBe(11);
    const inspected = await inspectWorkspacePackage(archive, "交接.roadmap");
    expect(inspected.data.requirements[0]).toMatchObject({ id: "r1", title: "社交时差", tieringFeatureId: "REQ-r1" });
    expect(inspected.data.requirements[0].images).toEqual([image]);
    expect(inspected.data.domains).toEqual(workspace.domains);
    expect(inspected.data.products).toEqual(workspace.products);
    expect(inspected.data.groupOverrides).toEqual(workspace.groupOverrides.map((item)=>({...item,groupKey:"健康::健康基础::d1::2027-03"})));
    expect(inspected.data.tiering).toMatchObject({
      features: [{ id: "REQ-r1", requirementId: "r1", requirementTitle: "社交时差" }],
    });
    expect(inspected.data.tiering?.productSeries).toHaveLength(15);
    expect(inspected.data.tiering?.supportRecords).toHaveLength(15);
    expect(inspected.preview).toMatchObject({
      sourceFormat: "roadmap",
      exportedAt: "2026-08-20T08:00:00.000Z",
      counts: { requirements: 1, images: 1, domains: 2, products: 1, groupOverrides: 1, tieringFeatures: 1, tieringProductSeries: 15, tieringSupportRecords: 15 },
    });
  });

  it("检查时为缺少分层行的纳入范围需求补齐功能和十五系列矩阵", async () => {
    const current: WorkspaceData = { ...workspace, requirements: [], groupOverrides: [], tiering: EMPTY_TIERING_DATA };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(workspace, "0.8.24"), "缺少分层行.roadmap", current);
    expect(inspected.data.requirements[0].tieringFeatureId).toBe("REQ-r1");
    expect(inspected.preview.counts).toMatchObject({ tieringFeatures: 1, tieringProductSeries: 15, tieringSupportRecords: 15 });
    expect(inspected.preview.differences).toContainEqual(expect.objectContaining({ entity: "分层功能", key: "REQ-r1", change: "新增" }));
    expect(inspected.preview.differences.filter((item) => item.entity === "系列支持" && item.change === "新增")).toHaveLength(15);
  });

  it("全部需求纳入分层，旧行业规格与已有支持结论保留", async () => {
    const updatedAt = workspace.requirements[0].updatedAt;
    const industryRequirement = { ...workspace.requirements[0], source: "行业" as const, tieringFeatureId: "REQ-r1" };
    const industryFeature = { id: "REQ-r1", l0: "健康", l1: "睡眠", requirementTitle: industryRequirement.title, requirementId: "r1", sortOrder: 0, createdAt: updatedAt, updatedAt };
    const oldPackage: WorkspaceData = {
      ...workspace,
      requirements: [industryRequirement],
      tiering: {
        ...EMPTY_TIERING_DATA,
        features: [industryFeature],
        productSeries: [{ ...BUILT_IN_TIERING_PRODUCT_SERIES[0], sourceIds: [] }],
        supportRecords: [createTieringSupportRecord(industryFeature.id, BUILT_IN_TIERING_PRODUCT_SERIES[0].id, updatedAt)],
      },
    };
    const current: WorkspaceData = { ...oldPackage, requirements: [{ ...industryRequirement, tieringFeatureId: undefined }], tiering: EMPTY_TIERING_DATA };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(oldPackage, "0.8.23"), "旧行业规格.roadmap", current);
    expect(inspected.data.requirements[0].tieringFeatureId).toBe("REQ-r1");
    expect(inspected.data.tiering?.features).toHaveLength(1);
    expect(inspected.preview.counts).toMatchObject({ tieringFeatures: 1, tieringProductSeries: 15, tieringSupportRecords: 15 });
  });

  it("v6 完整往返医疗节点并在差异预览中列出明确清除", async () => {
    const medical: WorkspaceData = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({
        ...item,
        source: "NMPA认证" as const,
        medicalSubmissionMilestones: { registrationStartMonth: "2027-01", clinicalStartMonth: "2027-04", certificateMonth: "2027-08" },
      })),
    };
    const roundTrip = await inspectWorkspacePackage(await createWorkspacePackage(medical, "0.8.12"), "医疗节点.roadmap");
    expect(roundTrip.data.medicalRoadmapPlans?.[0]).toMatchObject(medical.requirements[0].medicalSubmissionMilestones!);
    const cleared: WorkspaceData = { ...medical, requirements: medical.requirements.map((item) => ({ ...item, medicalSubmissionMilestones: undefined, updatedAt: "2026-09-04T08:00:00.000Z" })) };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(cleared, "0.8.12"), "清除医疗节点.roadmap", medical);
    expect(inspected.preview.differences.filter((item)=>item.entity==="医疗路标配置").flatMap((item)=>item.fields.map((field)=>field.field))).toContain("注册启动时间");
  });

  it("旧 JSON 包缺少医疗节点时保留本机已配置节点", async () => {
    const current: WorkspaceData = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, source: "CE认证" as const, medicalSubmissionMilestones: { registrationStartMonth: "2027-01" } })),
    };
    const legacyData = structuredClone(current);
    delete legacyData.requirements[0].medicalSubmissionMilestones;
    legacyData.requirements[0].updatedAt = "2026-09-04T08:00:00.000Z";
    const bytes = Buffer.from(JSON.stringify({ schemaVersion: 1, exportedAt: "2026-09-04T08:00:00.000Z", data: legacyData }));
    const inspected = await inspectWorkspacePackage(bytes, "旧版.json", current);
    expect(inspected.data.requirements[0].medicalSubmissionMilestones).toEqual({ registrationStartMonth: "2027-01" });
  });

  it("合并检查返回全部逐字段差异和确认摘要", async () => {
    const incoming: WorkspaceData = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, title: "社交时差升级", scope: "仅国内" as const, updatedAt: "2026-08-30T00:00:00.000Z" })),
    };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(incoming, "0.7.28"), "差异.roadmap", workspace);
    const requirementDifference = inspected.preview.differences.find((item) => item.entity === "需求");
    expect(requirementDifference).toMatchObject({ entity: "需求", title: "社交时差升级", change: "修改" });
    expect(requirementDifference?.fields.map((item) => item.field)).toEqual(expect.arrayContaining(["需求标题", "适用范围", "更新时间"]));
    expect(inspected.preview.differences).toContainEqual(expect.objectContaining({ entity: "分层功能", key: "REQ-r1", change: "修改" }));
    expect(inspected.preview.differenceDigest).toMatch(/^diff-/);
  });

  it("删除记录在完整包中往返，并把旧需求标为忽略、较新需求标为恢复", async () => {
    const deleted: WorkspaceData = {
      ...workspace,
      requirements: [],
      requirementTombstones: [{ id: "r1", title: "社交时差", deletedAt: "2026-09-10T00:00:00.000Z" }],
      requirementDeletionTrackingStartedAt: "2026-09-01T00:00:00.000Z",
    };
    const roundTrip = await inspectWorkspacePackage(await createWorkspacePackage(deleted, "0.8.23"), "删除同步.roadmap");
    expect(roundTrip.data.requirementTombstones).toEqual(deleted.requirementTombstones);
    expect(roundTrip.preview.counts.requirementTombstones).toBe(1);

    const older = await inspectWorkspacePackage(await createWorkspacePackage(workspace, "0.8.22"), "旧包.roadmap", deleted);
    expect(older.preview.ambiguousRequirementIds).toEqual([]);
    expect(older.preview.differences).toContainEqual(expect.objectContaining({ entity: "需求", key: "r1", change: "忽略" }));

    const newerWorkspace: WorkspaceData = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, title: "对方后续修改", updatedAt: "2026-09-11T00:00:00.000Z" })),
    };
    const newer = await inspectWorkspacePackage(await createWorkspacePackage(newerWorkspace, "0.8.23"), "新包.roadmap", deleted);
    expect(newer.preview.differences).toContainEqual(expect.objectContaining({ entity: "需求", key: "r1", change: "恢复" }));
  });

  it("删除追踪启用前的未知旧需求必须逐条选择，差异摘要包含待确认项", async () => {
    const local: WorkspaceData = {
      ...workspace,
      requirements: [],
      requirementTombstones: [],
      requirementDeletionTrackingStartedAt: "2026-09-01T00:00:00.000Z",
    };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(workspace, "0.8.22"), "旧设备协作包.roadmap", local);
    expect(inspected.preview.ambiguousRequirementIds).toEqual(["r1"]);
    expect(inspected.preview.differences).toContainEqual(expect.objectContaining({ entity: "需求", key: "r1", change: "待确认", decisionRequired: true }));
    expect(inspected.preview.differenceDigest).toMatch(/^diff-/);
  });

  it("完整往返人员责任矩阵和 PPT 白名单", async () => {
    const withMembers: WorkspaceData = {
      ...workspace,
      members: [{ id: "m1", name: "林规划", employeeNumber: "E001", microsoftAccount: "planner@company.com", role: "规划", domainL1Ids: ["d1"], sortOrder: 0, createdAt: "2026-08-20T00:00:00.000Z", updatedAt: "2026-08-28T00:00:00.000Z" }],
      pptAccessPolicy: { entries: [{ memberId: "m1", permission: "read" }], updatedAt: "2026-08-28T00:00:00.000Z" },
    };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(withMembers, "0.7.23"), "人员.roadmap");
    expect(inspected.data.members).toEqual(withMembers.members);
    expect(inspected.data.pptAccessPolicy).toEqual(withMembers.pptAccessPolicy);
    expect(inspected.preview.counts).toMatchObject({ members: 1, whitelistEntries: 1 });
  });

  it("分层分级初始包带 merge-only 类型及完整数量", async () => {
    const tieringWorkspace: WorkspaceData = {
      ...workspace,
      tiering: {
        ...EMPTY_TIERING_DATA,
        features: [{ id: "F001", l0: "基础体征测量", l1: "心率", requirementTitle: "单次心率测量", requirementId: "r1", sortOrder: 0, createdAt: "2026-08-05T00:00:00.000Z", updatedAt: "2026-08-05T00:00:00.000Z" }],
      },
      requirements: workspace.requirements.map((item) => ({ ...item, tieringFeatureId: "F001" })),
    };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(tieringWorkspace, "0.8.0", "2026-09-01T00:00:00.000Z", "tiering-seed"), "初始包.roadmap");
    expect(inspected.preview.packageKind).toBe("tiering-seed");
    expect(inspected.preview.counts.tieringFeatures).toBe(1);
  });

  it("完整往返历史上线时间证据并在差异预览中列出", async () => {
    const historicalWorkspace: WorkspaceData = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({
        ...item,
        targetMonth: "2018-09",
        updatedAt: "2026-09-01T00:00:00.000Z",
        historicalLaunch: {
          suggestedTargetMonth: "2018-09",
          precision: "half-year" as const,
          verificationStatus: "pending" as const,
          confidence: "B" as const,
          evidenceUrls: ["https://consumer.huawei.com/cn/example"],
          note: "H2 可信，月份待核验",
        },
      })),
    };
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(historicalWorkspace, "0.8.0"), "历史证据.roadmap", workspace);
    expect(inspected.data.requirements[0].historicalLaunch).toMatchObject({ precision: "half-year", verificationStatus: "pending", confidence: "B" });
    expect(inspected.preview.differences[0].fields.map((item) => item.field)).toEqual(expect.arrayContaining(["历史建议时间", "历史时间核验", "历史证据"]));
  });

  it("兼容缺少历史上线时间元数据的旧数据包", async () => {
    const inspected = await inspectWorkspacePackage(await createWorkspacePackage(workspace, "0.7.38"), "旧需求包.roadmap");
    expect(inspected.data.requirements[0].historicalLaunch).toBeUndefined();
  });

  it("允许三端工作量全部为零的数据完整往返", async () => {
    const zeroWorkload = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({
        ...item,
        deviceWorkloadPm: 0,
        appWorkloadPm: 0,
        cloudWorkloadPm: 0,
        unallocatedWorkloadPm: 0,
        workloadPm: 0,
      })),
    };
    const archive = await createWorkspacePackage(zeroWorkload, "0.7.0");
    const inspected = await inspectWorkspacePackage(archive, "待评估需求.roadmap");
    expect(inspected.data.requirements[0]).toMatchObject({
      deviceWorkloadPm: 0,
      appWorkloadPm: 0,
      cloudWorkloadPm: 0,
      workloadPm: 0,
    });
  });

  it("完整往返平台演进分类及其可选产品", async () => {
    const platformWorkspace = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, category: "平台演进" as const, productIds: ["p1"] })),
    };
    const archive = await createWorkspacePackage(platformWorkspace, "0.7.3");
    const inspected = await inspectWorkspacePackage(archive, "平台演进.roadmap");
    expect(inspected.data.requirements[0]).toMatchObject({ category: "平台演进", productIds: ["p1"] });
  });

  it("导出时清除体验优化需求中的旧产品关联", async () => {
    const experienceWorkspace = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, category: "体验优化" as const, productIds: ["p1"] })),
    };
    const archive = await createWorkspacePackage(experienceWorkspace, "0.7.18");
    const inspected = await inspectWorkspacePackage(archive, "体验优化.roadmap");
    expect(inspected.data.requirements[0]).toMatchObject({ category: "体验优化", productIds: [] });
  });

  it("产品专属在完整数据包中仍必须匹配一个产品", async () => {
    const invalid = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, productIds: [] })),
    };
    await expect(createWorkspacePackage(invalid, "0.7.10")).rejects.toThrow("必须且只能匹配一个产品");
  });

  it("文件被修改后拒绝导入", async () => {
    const archive = await createWorkspacePackage(workspace, "0.1.0");
    const zip = await JSZip.loadAsync(archive);
    zip.file("data.json", "{}");
    const damaged = await zip.generateAsync({ type: "nodebuffer" });
    await expect(inspectWorkspacePackage(damaged, "损坏.roadmap")).rejects.toThrow("文件校验失败");
  });

  it("兼容旧版 JSON 备份", async () => {
    const legacyData = {
      ...workspace,
      requirements: workspace.requirements.map(({ domainL0Id: _legacyMissingField, scope: _legacyMissingScope, ...item }) => item),
      domains: [{ ...workspace.domains.find((item) => item.id === "d1")!, parentId: undefined }],
    };
    const legacy = Buffer.from(JSON.stringify({ schemaVersion: 1, exportedAt: "2026-08-19", data: legacyData }));
    const inspected = await inspectWorkspacePackage(legacy, "旧备份.json");
    expect(inspected.preview.sourceFormat).toBe("legacy-json");
    expect(inspected.data.requirements[0].images[0].dataUrl).toBe(image.dataUrl);
    expect(inspected.data.requirements[0].domainL0Id).toBe("");
    expect(inspected.data.requirements[0].scope).toBe("全球");
    expect(inspected.data.domains[0]).toMatchObject({ id: "d1", level: "L1" });
  });

  it("兼容缺少人员与白名单的 v1 roadmap 包", async () => {
    const zip = await JSZip.loadAsync(await createWorkspacePackage(workspace, "0.7.22"));
    const data = JSON.parse(await zip.file("data.json")!.async("string"));
    delete data.members;
    delete data.pptAccessPolicy;
    const dataBytes = Buffer.from(`${JSON.stringify(data, null, 2)}\n`);
    zip.file("data.json", dataBytes);
    const manifest = JSON.parse(await zip.file("manifest.json")!.async("string"));
    manifest.formatVersion = 1;
    delete manifest.counts.members;
    delete manifest.counts.whitelistEntries;
    manifest.files["data.json"] = { sha256: createHash("sha256").update(dataBytes).digest("hex"), sizeBytes: dataBytes.length };
    zip.file("manifest.json", `${JSON.stringify(manifest, null, 2)}\n`);
    const inspected = await inspectWorkspacePackage(await zip.generateAsync({ type: "nodebuffer" }), "v1.roadmap");
    expect(inspected.data.members).toEqual([]);
    expect(inspected.data.pptAccessPolicy.entries).toEqual([]);
  });

  it("旧数据包中的 AI 来源和卡片键迁移为 AI综合", async () => {
    const legacyAiData = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, source: "AI" })),
      groupOverrides: workspace.groupOverrides.map((item) => ({ ...item, groupKey: "d1::AI::2027-03" })),
    };
    const legacy = Buffer.from(JSON.stringify({ schemaVersion: 1, exportedAt: "2026-08-24", data: legacyAiData }));
    const inspected = await inspectWorkspacePackage(legacy, "旧AI备份.json");
    expect(inspected.data.requirements[0].source).toBe("AI综合");
    expect(inspected.data.groupOverrides[0].groupKey).toBe("AI::AI综合::d1::2027-03");
  });

  it("旧数据包中的医疗认证来源和卡片键迁移为 NMPA认证", async () => {
    const legacyMedicalData = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, source: "医疗送检" })),
      groupOverrides: workspace.groupOverrides.map((item) => ({ ...item, groupKey: "d1::医疗送检::2027-03" })),
    };
    const legacy = Buffer.from(JSON.stringify({ schemaVersion: 1, exportedAt: "2026-08-24", data: legacyMedicalData }));
    const inspected = await inspectWorkspacePackage(legacy, "旧医疗认证备份.json");
    expect(inspected.data.requirements[0].source).toBe("NMPA认证");
    expect(inspected.data.groupOverrides[0].groupKey).toBe("医疗认证::NMPA认证::d1::2027-03");
  });

  it("拒绝包含非法适用范围的数据包", async () => {
    const invalid = {
      ...workspace,
      requirements: [{ ...workspace.requirements[0], scope: "未知" as never }],
    };
    await expect(createWorkspacePackage(invalid, "0.7.2")).rejects.toThrow("适用范围无效");
  });

  it("拒绝包含非法卡片顺序的数据包", async () => {
    const invalid = {
      ...workspace,
      groupOverrides: workspace.groupOverrides.map((item) => ({ ...item, sortOrder: -1 })),
    };
    await expect(createWorkspacePackage(invalid, "0.7.6")).rejects.toThrow("排序无效");
  });
});
