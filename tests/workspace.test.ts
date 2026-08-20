import { describe, expect, it } from "vitest";
import { mergeWorkspaceData, workspaceCounts } from "../src/shared/workspace";
import type { Requirement, WorkspaceData } from "../src/shared/types";

describe("工作区合并", () => {
  it("按名称合并字典并重映射需求和卡片键", () => {
    const local = data({
      domains: [{ id: "local-sleep", name: "睡眠", sortOrder: 0, active: true }],
      products: [{ id: "local-watch", name: "手表", sortOrder: 0, active: true }],
    });
    const incoming = data({
      domains: [{ id: "incoming-sleep", name: " 睡眠 ", sortOrder: 0, active: true }],
      products: [{ id: "incoming-watch", name: "手表", sortOrder: 0, active: true }],
      requirements: [requirement("r1", "incoming-sleep", ["incoming-watch"], "2026-08-20")],
      groupOverrides: [{ groupKey: "incoming-sleep::健康基础::2027-03", cardTitle: "睡眠", cardSummary: "摘要", updatedAt: "2026-08-20" }],
    });

    const merged = mergeWorkspaceData(local, incoming);
    expect(merged.domains).toHaveLength(1);
    expect(merged.products).toHaveLength(1);
    expect(merged.requirements[0]).toMatchObject({ domainId: "local-sleep", productIds: ["local-watch"] });
    expect(merged.groupOverrides[0].groupKey).toBe("local-sleep::健康基础::2027-03");
  });

  it("同ID需求保留更新时间较新版本", () => {
    const local = data({ requirements: [requirement("r1", "d1", [], "2026-08-21", "本机新版")] });
    const incoming = data({ requirements: [requirement("r1", "d1", [], "2026-08-20", "导入旧版")] });
    expect(mergeWorkspaceData(local, incoming).requirements[0].title).toBe("本机新版");
  });

  it("统计完整工作区内容", () => {
    const item = requirement("r1", "d1", [], "2026-08-20");
    item.images = [{ id: "i1", name: "a.png", mimeType: "image/png", sizeBytes: 1, dataUrl: "data:image/png;base64,YQ==" }];
    expect(workspaceCounts(data({ requirements: [item] }))).toMatchObject({ requirements: 1, images: 1 });
  });
});

function data(overrides: Partial<WorkspaceData> = {}): WorkspaceData {
  return {
    requirements: [],
    domains: [{ id: "d1", name: "默认领域", sortOrder: 0, active: true }],
    products: [],
    groupOverrides: [],
    ...overrides,
  };
}

function requirement(id: string, domainId: string, productIds: string[], updatedAt: string, title = "需求"): Requirement {
  return {
    id,
    title,
    description: "",
    images: [],
    domainId,
    source: "健康基础",
    category: productIds.length ? "产品专属" : "体验优化",
    targetMonth: "2027-03",
    productIds,
    deviceWorkloadPm: 0,
    appWorkloadPm: 1,
    cloudWorkloadPm: 0,
    unallocatedWorkloadPm: 0,
    workloadPm: 1,
    createdAt: "2026-08-19",
    updatedAt,
  };
}
