import { describe, it, expect, vi } from "vitest";
import type { AppSnapshot, Requirement } from "../src/shared/types";
import { mcpExportPreview, validateMcpExport, mcpPptRequest } from "../src/shared/mcpExports";
import { McpProposalStore } from "../src/main/mcpProposals";

const req = (id: string, targetMonth: string): Requirement => ({ id, title: id, description: "", targetMonth, images: [], source: "健康基础", productIds: [], overseasRegions: [], category: "平台演进", workloadPm: 0, scope: "全球", domainL0Id: "l0", domainId: "l1", allowLatestVersionUpgrade: true, deviceWorkloadPm: 0, appWorkloadPm: 0, cloudWorkloadPm: 0, unallocatedWorkloadPm: 0, createdAt: "2026-01-01", updatedAt: "2026-01-01" });
const snapshot = (): AppSnapshot => ({ requirements: [req("within", "2027-03"), req("outside", "2025-01")], domains: [], products: [], members: [], groupOverrides: [], pptAccessPolicy: { entries: [], updatedAt: "2026-01-01" }, templates: [], roadmaps: [{ id: "custom", name: "自定义", kind: "standard", sortOrder: 0, updatedAt: "2026-01-01" }], roadmapLanes: [], roadmapAssignments: [], medicalRoadmapPlans: [] });

describe("MCP exports", () => {
  it("supports all formats with explicit scope and full-package warning", () => {
    const snap = snapshot();
    expect(mcpExportPreview(snap, validateMcpExport(snap, { format: "workspace" }))).toMatchObject({ exportCount: 2, exportWarning: expect.stringContaining("人员工号与邮箱") });
    for (const format of ["requirements_excel", "tiering_excel"] as const) {
      expect(mcpExportPreview(snap, validateMcpExport(snap, { format, start: "2027H1", end: "2027H2" }))).toMatchObject({ exportCount: 1 });
    }
  });
  it("rejects model-provided paths, permissions, malformed and reversed dates", () => {
    const snap = snapshot();
    for (const input of [
      { format: "workspace", outputPath: "/secret/overwrite" },
      { format: "workspace", start: "2027H1", end: "2027H2" },
      { format: "ppt", start: "2027H1", end: "2027H2", unprotected: true },
      { format: "requirements_excel", start: "27H1", end: "2027H2" },
      { format: "requirements_excel", start: "2027H2", end: "2027H1" },
      { format: "requirements_excel" },
    ]) expect(() => validateMcpExport(snap, input)).toThrow();
  });
  it("uses dynamic roadmap IDs, handles cover-only and enforces PPT capacity", () => {
    const snap = snapshot();
    const request = validateMcpExport(snap, { format: "ppt", start: "2027H1", end: "2028H2" });
    if (request.format !== "ppt") throw Error();
    expect(mcpPptRequest(snap, request).sections).toEqual(["summary", "custom"]);
    expect(() => validateMcpExport(snap, { format: "ppt", start: "2027H1", end: "2029H1" })).toThrow("最多支持");
    expect(() => validateMcpExport(snap, { format: "ppt", start: "2027H1", end: "2027H2", roadmapIds: ["missing"] })).toThrow("不存在");
    const cover = validateMcpExport(snap, { format: "ppt", start: "2027H1", end: "2027H1", roadmapIds: [], includeSummary: false });
    expect(mcpExportPreview(snap, cover).exportSections).toBe("仅封面");
  });
  it("does not export before approval, uses one operation and returns only after file creation", async () => {
    const snap = snapshot();
    const execute = vi.fn(async () => ({ canceled: false, path: "/chosen/output.roadmap" }));
    const apply = vi.fn(async () => snap);
    const store = new McpProposalStore({ getSnapshot: () => snap, applyMcpRequirementBatch: apply }, undefined, Date.now, execute);
    const proposal = store.proposeExport({ format: "workspace" });
    expect(execute).not.toHaveBeenCalled();
    expect(proposal.items).toHaveLength(1);
    await store.decide(proposal.id, true);
    expect(proposal.state).toBe("approved"); expect(proposal.exportResult?.path).toBe("/chosen/output.roadmap");
    expect(apply).not.toHaveBeenCalled();
  });
  it("reject, expiry and cancelled save never count as successful exports", async () => {
    let now = 0;
    const execute = vi.fn(async () => ({ canceled: true }));
    const snap = snapshot();
    const store = new McpProposalStore({ getSnapshot: () => snap, applyMcpRequirementBatch: async () => snap }, undefined, () => now, execute);
    const rejected = store.proposeExport({ format: "workspace" }); await store.decide(rejected.id, false);
    expect(execute).not.toHaveBeenCalled();
    const expired = store.proposeExport({ format: "workspace" }); now = 31 * 60 * 1000;
    await expect(store.decide(expired.id, true)).rejects.toThrow("不可处理");
    expect(execute).not.toHaveBeenCalled();
    const cancelled = store.proposeExport({ format: "workspace" }); await store.decide(cancelled.id, true);
    expect(cancelled.state).toBe("cancelled"); expect(cancelled.exportResult?.path).toBeUndefined();
  });
  it("data or whitelist changes require a new export proposal", async () => {
    const snap = snapshot(); const execute = vi.fn(async () => ({ canceled: false, path: "/x" }));
    const store = new McpProposalStore({ getSnapshot: () => snap, applyMcpRequirementBatch: async () => snap }, undefined, Date.now, execute);
    const first = store.proposeExport({ format: "workspace" }); snap.requirements.push(req("new", "2027-01"));
    await expect(store.decide(first.id, true)).rejects.toThrow("已变化");
    const second = store.proposeExport({ format: "workspace" }); snap.pptAccessPolicy.updatedAt = "later";
    await expect(store.decide(second.id, true)).rejects.toThrow("已变化");
    expect(execute).not.toHaveBeenCalled();
  });
  it("an in-progress export cannot be approved twice", async () => {
    const snap = snapshot(); let finish!: () => void;
    const execute = vi.fn(() => new Promise<{ canceled: boolean }>(resolve => { finish = () => resolve({ canceled: true }); }));
    const store = new McpProposalStore({ getSnapshot: () => snap, applyMcpRequirementBatch: async () => snap }, undefined, Date.now, execute);
    const proposal = store.proposeExport({ format: "workspace" }); const pending = store.decide(proposal.id, true);
    expect(proposal.state).toBe("exporting");
    await expect(store.decide(proposal.id, true)).rejects.toThrow("不可处理"); finish(); await pending;
    expect(execute).toHaveBeenCalledTimes(1);
  });
  it("IRM/export errors do not report success or fallback file paths", async () => {
    const snap = snapshot();
    const store = new McpProposalStore({ getSnapshot: () => snap, applyMcpRequirementBatch: async () => snap }, undefined, Date.now, async () => { throw new Error("IRM 服务不可用"); });
    const proposal = store.proposeExport({ format: "ppt", start: "2027H1", end: "2027H1" });
    await expect(store.decide(proposal.id, true)).rejects.toThrow("IRM");
    expect(proposal.state).toBe("failed"); expect(proposal.exportResult).toBeUndefined();
  });
});
