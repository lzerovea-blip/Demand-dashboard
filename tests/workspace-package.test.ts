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
    domainId: "d1",
    source: "健康基础",
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
  domains: [{ id: "d1", name: "睡眠", sortOrder: 0, active: true }],
  products: [{ id: "p1", name: "手表", sortOrder: 0, active: true }],
  groupOverrides: [{ groupKey: "d1::健康基础::2027-03", cardTitle: "睡眠", cardSummary: "摘要", updatedAt: "2026-08-20T00:00:00.000Z" }],
};

describe(".roadmap 数据包", () => {
  it("完整往返需求、描述、图片、字典和卡片编辑", async () => {
    const archive = await createWorkspacePackage(workspace, "0.1.0", "2026-08-20T08:00:00.000Z");
    const inspected = await inspectWorkspacePackage(archive, "交接.roadmap");
    expect(inspected.data).toEqual(workspace);
    expect(inspected.preview).toMatchObject({
      sourceFormat: "roadmap",
      exportedAt: "2026-08-20T08:00:00.000Z",
      counts: { requirements: 1, images: 1, domains: 1, products: 1, groupOverrides: 1 },
    });
  });

  it("文件被修改后拒绝导入", async () => {
    const archive = await createWorkspacePackage(workspace, "0.1.0");
    const zip = await JSZip.loadAsync(archive);
    zip.file("data.json", "{}");
    const damaged = await zip.generateAsync({ type: "nodebuffer" });
    await expect(inspectWorkspacePackage(damaged, "损坏.roadmap")).rejects.toThrow("文件校验失败");
  });

  it("兼容旧版 JSON 备份", async () => {
    const legacy = Buffer.from(JSON.stringify({ schemaVersion: 1, exportedAt: "2026-08-19", data: workspace }));
    const inspected = await inspectWorkspacePackage(legacy, "旧备份.json");
    expect(inspected.preview.sourceFormat).toBe("legacy-json");
    expect(inspected.data.requirements[0].images[0].dataUrl).toBe(image.dataUrl);
  });
});
