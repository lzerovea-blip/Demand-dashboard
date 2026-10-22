import type { AppSnapshot } from "../shared/types.js";
import {
  MCP_EDITABLE_FIELDS, MCP_PROPOSAL_TTL_MS, mcpDifferences, mcpRequirementView,
  type McpProposal, type McpRequirementOperation,
} from "../shared/mcp.js";

export interface McpProposalDatabase {
  getSnapshot(): AppSnapshot;
  applyMcpRequirementBatch(operations: McpRequirementOperation[]): Promise<AppSnapshot>;
}

export class McpProposalStore {
  private readonly proposals = new Map<string, McpProposal>();
  constructor(private readonly database: McpProposalDatabase, private readonly changed: () => void = () => undefined,
    private readonly clock: () => number = Date.now) {}

  propose(operations: McpRequirementOperation[]): McpProposal {
    if (!Array.isArray(operations) || operations.length < 1 || operations.length > 25) throw new Error("每批须包含 1 至 25 项需求操作");
    const snapshot = this.database.getSnapshot();
    const byId = new Map(snapshot.requirements.map((item) => [item.id, item]));
    const touched = new Set<string>();
    const items = operations.map((operation) => {
      if (!operation || !["create", "update", "duplicate", "delete"].includes(operation.type)) throw new Error("无效的 MCP 需求操作");
      if (operation.type === "create") {
        if (!operation.fields || Object.keys(operation.fields).some((key) => !(MCP_EDITABLE_FIELDS as readonly string[]).includes(key))) throw new Error("新增包含未开放字段");
        return { operation, title: String(operation.fields.title ?? ""), differences: mcpDifferences({}, operation.fields as unknown as Record<string, unknown>) };
      }
      if (!operation.id || touched.has(operation.id)) throw new Error("同一批次不能重复操作同一条需求");
      touched.add(operation.id);
      const current = byId.get(operation.id);
      if (!current) throw new Error(`需求不存在：${operation.id}`);
      const before = mcpRequirementView(current);
      if (operation.type === "update") {
        if (!operation.fields || !Object.keys(operation.fields).length || Object.keys(operation.fields).some((key) => !(MCP_EDITABLE_FIELDS as readonly string[]).includes(key))) throw new Error("修改包含空值或未开放字段");
        return { operation, title: current.title, baselineUpdatedAt: current.updatedAt,
          differences: mcpDifferences(before, operation.fields as Record<string, unknown>) };
      }
      return { operation, title: current.title, baselineUpdatedAt: current.updatedAt,
        differences: [{ field: operation.type === "delete" ? "整条需求" : "复制需求", before: current.title, after: operation.type === "delete" ? "删除" : "新建副本" }] };
    });
    const now = this.clock();
    const proposal: McpProposal = { id: crypto.randomUUID(), state: "pending", createdAt: new Date(now).toISOString(),
      expiresAt: new Date(now + MCP_PROPOSAL_TTL_MS).toISOString(), items };
    this.proposals.set(proposal.id, proposal);
    setTimeout(() => this.expire(proposal), MCP_PROPOSAL_TTL_MS + 1).unref();
    this.changed();
    return proposal;
  }

  get(id: string): McpProposal {
    const proposal = this.proposals.get(id);
    if (!proposal) throw new Error("提案不存在或工作台已重启，请重新提出");
    this.expire(proposal);
    return proposal;
  }

  listPending(): McpProposal[] {
    for (const proposal of this.proposals.values()) this.expire(proposal);
    return [...this.proposals.values()].filter((item) => item.state === "pending");
  }

  async decide(id: string, accept: boolean): Promise<{ proposal: McpProposal; snapshot?: AppSnapshot }> {
    const proposal = this.get(id);
    if (proposal.state !== "pending") throw new Error(`提案不可处理：${proposal.state}`);
    if (!accept) { proposal.state = "cancelled"; this.changed(); return { proposal }; }
    const byId = new Map(this.database.getSnapshot().requirements.map((item) => [item.id, item]));
    for (const item of proposal.items) {
      if (item.operation.type === "create") continue;
      if (byId.get(item.operation.id)?.updatedAt !== item.baselineUpdatedAt) {
        proposal.state = "conflict";
        proposal.error = "需求已在提案后被修改或删除，请重新提出";
        this.changed();
        throw new Error(proposal.error);
      }
    }
    try {
      const snapshot = await this.database.applyMcpRequirementBatch(proposal.items.map((item) => item.operation));
      proposal.state = "approved";
      this.changed();
      return { proposal, snapshot };
    } catch (error) {
      proposal.state = "failed";
      proposal.error = error instanceof Error ? error.message : "需求校验失败";
      this.changed();
      throw error;
    }
  }

  private expire(proposal: McpProposal): void {
    if (proposal.state === "pending" && this.clock() >= Date.parse(proposal.expiresAt)) {
      proposal.state = "expired";
      this.changed();
    }
  }
}
