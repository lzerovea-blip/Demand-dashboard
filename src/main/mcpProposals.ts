import type { AppSnapshot } from "../shared/types.js";
import { createHash } from "node:crypto";
import { validateMcpExport, mcpExportPreview, MCP_EXPORT_NAMES, type McpExportRequest, type McpExportResult } from "../shared/mcpExports.js";
import { MCP_DICTIONARY_NAMES, validateMcpDictionaryAdditions, type McpDictionaryAddition } from "../shared/mcpDictionaries.js";
import {
  MCP_EDITABLE_FIELDS, MCP_PROPOSAL_TTL_MS, mcpDifferences, mcpRequirementView, validateMcpTargetMonth,
  type McpProposal, type McpRequirementOperation,
} from "../shared/mcp.js";

export interface McpProposalDatabase {
  getSnapshot(): AppSnapshot;
  applyMcpRequirementBatch(operations: McpRequirementOperation[]): Promise<AppSnapshot>;
  applyMcpDictionaryBatch?(entries: McpDictionaryAddition[]): Promise<AppSnapshot>;
  getWorkspaceData?(): unknown;
}

export class McpProposalStore {
  private readonly proposals = new Map<string, McpProposal>();
  private readonly exportBaselines = new Map<string, string>();
  constructor(private readonly database: McpProposalDatabase, private readonly changed: () => void = () => undefined,
    private readonly clock: () => number = Date.now,
    private readonly exportFile?: (request: McpExportRequest) => Promise<McpExportResult>) {}

  propose(operations: McpRequirementOperation[]): McpProposal {
    if (!Array.isArray(operations) || operations.length < 1 || operations.length > 25) throw new Error("每批须包含 1 至 25 项需求操作");
    const snapshot = this.database.getSnapshot();
    const byId = new Map(snapshot.requirements.map((item) => [item.id, item]));
    const touched = new Set<string>();
    const items = operations.map((operation) => {
      if (!operation || !["create", "update", "duplicate", "delete"].includes(operation.type)) throw new Error("无效的 MCP 需求操作");
      if (operation.type === "create") {
        if (!operation.fields || Object.keys(operation.fields).some((key) => !(MCP_EDITABLE_FIELDS as readonly string[]).includes(key))) throw new Error("新增包含未开放字段");
        if (!operation.fields.targetMonth) throw new Error("上线年月必填，格式 YYYY-MM，例如 2027-03");
        validateMcpTargetMonth(operation.fields);
        return { operation, title: String(operation.fields.title ?? ""), differences: mcpDifferences({}, operation.fields as unknown as Record<string, unknown>) };
      }
      if (!operation.id || touched.has(operation.id)) throw new Error("同一批次不能重复操作同一条需求");
      touched.add(operation.id);
      const current = byId.get(operation.id);
      if (!current) throw new Error(`需求不存在：${operation.id}`);
      const before = mcpRequirementView(current);
      if (operation.type === "update") {
        if (!operation.fields || !Object.keys(operation.fields).length || Object.keys(operation.fields).some((key) => !(MCP_EDITABLE_FIELDS as readonly string[]).includes(key))) throw new Error("修改包含空值或未开放字段");
        validateMcpTargetMonth(operation.fields);
        return { operation, title: current.title, baselineUpdatedAt: current.updatedAt,
          differences: mcpDifferences(before, operation.fields as Record<string, unknown>) };
      }
      return { operation, title: current.title, baselineUpdatedAt: current.updatedAt,
        differences: [{ field: operation.type === "delete" ? "整条需求" : "复制需求", before: current.title, after: operation.type === "delete" ? "删除" : "新建副本" }] };
    });
    return this.createProposal(items);
  }

  proposeDictionaries(input: unknown): McpProposal {
    const entries = validateMcpDictionaryAdditions(this.database.getSnapshot(), input);
    return this.createProposal(entries.map(entry => ({ operation: { type: "create_dictionary" as const, entry },
      title: `新增${MCP_DICTIONARY_NAMES[entry.kind]}：${entry.name}`,
      differences: mcpDifferences({}, Object.fromEntries(Object.entries(entry).filter(([field]) => field !== "kind"))) })));
  }

  proposeExport(input: unknown): McpProposal {
    const snapshot = this.database.getSnapshot();
    const request = validateMcpExport(snapshot, input);
    const proposal = this.createProposal([{ operation: { type: "export", request }, title: `导出${MCP_EXPORT_NAMES[request.format]}`,
      differences: mcpDifferences({}, mcpExportPreview(snapshot, request)) }]);
    this.exportBaselines.set(proposal.id, this.exportFingerprint());
    return proposal;
  }

  private exportFingerprint(): string {
    return createHash("sha256").update(JSON.stringify(this.database.getWorkspaceData?.() ?? this.database.getSnapshot())).digest("hex");
  }

  private createProposal(items: McpProposal["items"]): McpProposal {
    // Detach proposal data from the caller so pending content cannot be replaced.
    items = structuredClone(items);
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
    if (!accept) { proposal.state = "cancelled"; this.exportBaselines.delete(id); this.changed(); return { proposal }; }
    const exportItem = proposal.items.find(item => item.operation.type === "export");
    if (exportItem?.operation.type === "export") {
      if (proposal.items.length !== 1 || !this.exportFile) throw new Error("导出功能尚未就绪");
      if (this.exportBaselines.get(id) !== this.exportFingerprint()) {
        proposal.state = "conflict"; proposal.error = "导出提案后数据或权限已变化，请重新提出并核对范围"; this.changed(); throw new Error(proposal.error);
      }
      proposal.state = "exporting"; this.changed();
      try {
        const result = await this.exportFile(validateMcpExport(this.database.getSnapshot(), exportItem.operation.request));
        if (!result.canceled && !result.path) throw new Error("导出未返回有效文件路径");
        proposal.exportResult = result;
        proposal.state = result.canceled ? "cancelled" : "approved";
        this.changed(); return { proposal };
      } catch (error) {
        proposal.state = "failed"; proposal.error = error instanceof Error ? error.message : "导出失败"; this.changed(); throw error;
      } finally { this.exportBaselines.delete(id); }
    }
    const byId = new Map(this.database.getSnapshot().requirements.map((item) => [item.id, item]));
    for (const item of proposal.items) {
      if (item.operation.type === "create" || item.operation.type === "create_dictionary" || item.operation.type === "export") continue;
      if (byId.get(item.operation.id)?.updatedAt !== item.baselineUpdatedAt) {
        proposal.state = "conflict";
        proposal.error = "需求已在提案后被修改或删除，请重新提出";
        this.changed();
        throw new Error(proposal.error);
      }
    }
    try {
      const dictionaries = proposal.items.filter(item => item.operation.type === "create_dictionary");
      let snapshot: AppSnapshot;
      if (dictionaries.length) {
        if (dictionaries.length !== proposal.items.length || !this.database.applyMcpDictionaryBatch) throw new Error("不支持的字典提案批次");
        const entries = dictionaries.map(item => { if (item.operation.type !== "create_dictionary") throw new Error("无效字典操作"); return item.operation.entry; });
        validateMcpDictionaryAdditions(this.database.getSnapshot(), entries);
        snapshot = await this.database.applyMcpDictionaryBatch(entries);
      } else {
        snapshot = await this.database.applyMcpRequirementBatch(proposal.items.map(item => { if (item.operation.type === "create_dictionary" || item.operation.type === "export") throw new Error("无效需求操作"); return item.operation; }));
      }
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
      this.exportBaselines.delete(proposal.id);
      this.changed();
    }
  }
}
