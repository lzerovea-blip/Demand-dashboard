import { describe, expect, it, vi } from "vitest";
import { McpProposalStore } from "../src/main/mcpProposals";
import { MCP_PROPOSAL_TTL_MS, mcpRequirementView, type McpRequirementOperation } from "../src/shared/mcp";
import type { AppSnapshot, Requirement } from "../src/shared/types";

const requirement: Requirement = {
  id: "r1", title: "原需求", description: "描述", images: [], domainL0Id: "l0", domainId: "l1",
  scope: "全球", allowLatestVersionUpgrade: true, overseasRegions: [], category: "平台演进",
  targetMonth: "2027-03", productIds: [], planningMemberId: "p", productManagerMemberId: "m", seMemberId: "s",
  deviceWorkloadPm: 1, appWorkloadPm: 0, cloudWorkloadPm: 0, researchAppWorkloadPm: 0, algorithmWorkloadPm: 0,
  unallocatedWorkloadPm: 0, workloadPm: 1, createdAt: "2026-01-01T00:00:00.000Z", updatedAt: "2026-01-01T00:00:00.000Z",
};
const snapshot = (requirements: Requirement[]) => ({ requirements, domains: [], products: [], members: [], groupOverrides: [],
  pptAccessPolicy: { entries: [], updatedAt: "2026-01-01T00:00:00.000Z" }, templates: [] }) as AppSnapshot;

describe("MCP proposal safety", () => {
  it("validates date format before creating a pending proposal even through the bridge", () => {
    const store = new McpProposalStore({ getSnapshot: () => snapshot([requirement]), applyMcpRequirementBatch: async () => snapshot([requirement]) });
    for (const targetMonth of ["2027年3月", "2027-3", "2027-13", "2027/03", ""]) {
      expect(() => store.propose([{ type: "update", id: "r1", fields: { targetMonth } }])).toThrow("YYYY-MM");
    }
    expect(store.listPending()).toHaveLength(0);
    expect(store.propose([{ type: "update", id: "r1", fields: { targetMonth: "2027-03" } }]).state).toBe("pending");
  });
  it("returns business fields but not image bytes, email, or rights", () => {
    expect(mcpRequirementView({ ...requirement, images: [{ id: "i", name: "x", mimeType: "image/png", sizeBytes: 1, dataUrl: "data:image/png;base64,AAAA" }] })).toMatchObject({ imageCount: 1 });
    expect(JSON.stringify(mcpRequirementView(requirement))).not.toMatch(/dataUrl|microsoftAccount|pptAccessPolicy/);
  });
  it("previews field changes and applies only after approval", async () => {
    let current = snapshot([requirement]);
    const apply = vi.fn(async (operations: McpRequirementOperation[]) => {
      const patch = operations[0].type === "update" ? operations[0].fields : {};
      current = snapshot([{ ...requirement, ...patch, updatedAt: "later" }]);
      return current;
    });
    const store = new McpProposalStore({ getSnapshot: () => current, applyMcpRequirementBatch: apply });
    const proposal = store.propose([{ type: "update", id: "r1", fields: { title: "新需求" } }]);
    expect(proposal.items[0].differences).toEqual([{ field: "title", before: "原需求", after: "新需求" }]);
    expect(apply).not.toHaveBeenCalled();
    expect((await store.decide(proposal.id, true)).proposal.state).toBe("approved");
    expect(apply).toHaveBeenCalledTimes(1);
  });
  it("rejects stale requirements, duplicate targets and forbidden fields", async () => {
    let current = snapshot([requirement]);
    const apply = vi.fn(async () => current);
    const store = new McpProposalStore({ getSnapshot: () => current, applyMcpRequirementBatch: apply });
    const proposal = store.propose([{ type: "delete", id: "r1" }]);
    current = snapshot([{ ...requirement, updatedAt: "later" }]);
    await expect(store.decide(proposal.id, true)).rejects.toThrow("重新提出");
    expect(store.get(proposal.id).state).toBe("conflict");
    expect(apply).not.toHaveBeenCalled();
    expect(() => store.propose([{ type: "delete", id: "r1" }, { type: "duplicate", id: "r1" }])).toThrow("重复操作");
    expect(() => store.propose([{ type: "update", id: "r1", fields: { roadmapId: "x" } as never }])).toThrow("未开放字段");
  });
  it("cancellation and 30-minute expiry never write", async () => {
    let now = Date.parse("2026-10-22T00:00:00Z");
    const apply = vi.fn(async () => snapshot([requirement]));
    const store = new McpProposalStore({ getSnapshot: () => snapshot([requirement]), applyMcpRequirementBatch: apply }, undefined, () => now);
    const cancelled = store.propose([{ type: "delete", id: "r1" }]);
    expect((await store.decide(cancelled.id, false)).proposal.state).toBe("cancelled");
    const expired = store.propose([{ type: "delete", id: "r1" }]);
    now += MCP_PROPOSAL_TTL_MS;
    expect(store.get(expired.id).state).toBe("expired");
    await expect(store.decide(expired.id, true)).rejects.toThrow("不可处理");
    expect(apply).not.toHaveBeenCalled();
  });
});
