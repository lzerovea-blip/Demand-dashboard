import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { createWorkspacePackage, inspectWorkspacePackage } from "../src/main/workspacePackage";
import type { RequirementImage, WorkspaceData } from "../src/shared/types";

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
    overseasRegions: [],
    category: "产品专属",
    targetMonth: "2027-03",
    productIds: ["p1"],
    deviceWorkloadPm: 0.25,
    appWorkloadPm: 0.5,
    cloudWorkloadPm: 0.25,
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
};

describe(".roadmap 数据包", () => {
  it("完整往返需求、描述、图片、字典和卡片编辑", async () => {
    const archive = await createWorkspacePackage(workspace, "0.1.0", "2026-08-20T08:00:00.000Z");
    const inspected = await inspectWorkspacePackage(archive, "交接.roadmap");
    expect(inspected.data).toEqual(workspace);
    expect(inspected.preview).toMatchObject({
      sourceFormat: "roadmap",
      exportedAt: "2026-08-20T08:00:00.000Z",
      counts: { requirements: 1, images: 1, domains: 2, products: 1, groupOverrides: 1 },
    });
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

  it("完整往返平台演进分类且不携带产品", async () => {
    const platformWorkspace = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, category: "平台演进" as const, productIds: [] })),
    };
    const archive = await createWorkspacePackage(platformWorkspace, "0.7.3");
    const inspected = await inspectWorkspacePackage(archive, "平台演进.roadmap");
    expect(inspected.data.requirements[0]).toMatchObject({ category: "平台演进", productIds: [] });
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

  it("旧数据包中的 AI 来源和卡片键迁移为 AI综合", async () => {
    const legacyAiData = {
      ...workspace,
      requirements: workspace.requirements.map((item) => ({ ...item, source: "AI" })),
      groupOverrides: workspace.groupOverrides.map((item) => ({ ...item, groupKey: "d1::AI::2027-03" })),
    };
    const legacy = Buffer.from(JSON.stringify({ schemaVersion: 1, exportedAt: "2026-08-24", data: legacyAiData }));
    const inspected = await inspectWorkspacePackage(legacy, "旧AI备份.json");
    expect(inspected.data.requirements[0].source).toBe("AI综合");
    expect(inspected.data.groupOverrides[0].groupKey).toBe("d1::AI综合::2027-03");
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
