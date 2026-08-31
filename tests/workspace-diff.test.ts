import { describe, expect, it } from "vitest";
import type { WorkspaceData } from "../src/shared/types";
import { diffWorkspaceData, workspaceDifferenceDigest } from "../src/shared/workspaceDiff";

function workspace(title = "原需求"): WorkspaceData {
  return {
    requirements: [{
      id: "r1", title, description: "描述", images: [], domainL0Id: "l0", domainId: "l1", source: "健康基础", scope: "全球",
      allowLatestVersionUpgrade: false, overseasRegions: [], category: "体验优化", targetMonth: "2027-03", productIds: [],
      deviceWorkloadPm: 1, appWorkloadPm: 0, cloudWorkloadPm: 0, unallocatedWorkloadPm: 0, workloadPm: 1,
      createdAt: "2026-08-01T00:00:00.000Z", updatedAt: "2026-08-01T00:00:00.000Z",
    }],
    domains: [
      { id: "l0", name: "健康", level: "L0", sortOrder: 0, active: true },
      { id: "l1", name: "睡眠", level: "L1", parentId: "l0", sortOrder: 0, active: true },
    ],
    products: [], groupOverrides: [], members: [],
    pptAccessPolicy: { entries: [], updatedAt: "1970-01-01T00:00:00.000Z" },
  };
}

describe("导入差异清单", () => {
  it("逐字段列出修改、新增和删除，并生成稳定摘要", () => {
    const before = workspace();
    const after = workspace("修改后的需求");
    after.requirements[0].scope = "仅国内";
    after.requirements[0].deviceWorkloadPm = 2;
    after.requirements[0].workloadPm = 2;
    after.products.push({ id: "p1", name: "watch", sortOrder: 0, active: true });
    after.domains = after.domains.filter((item) => item.id !== "l1");
    const differences = diffWorkspaceData(before, after);
    expect(differences.map((item) => `${item.entity}:${item.change}`)).toEqual(expect.arrayContaining(["需求:修改", "领域:删除", "产品:新增"]));
    expect(differences.find((item) => item.entity === "需求")?.fields.map((item) => item.field)).toEqual(expect.arrayContaining(["需求标题", "适用范围", "设备工作量"]));
    expect(workspaceDifferenceDigest(differences)).toBe(workspaceDifferenceDigest(diffWorkspaceData(before, after)));
  });

  it("无变化时返回空清单", () => {
    expect(diffWorkspaceData(workspace(), workspace())).toEqual([]);
    expect(workspaceDifferenceDigest([])).toMatch(/^diff-/);
  });
});
