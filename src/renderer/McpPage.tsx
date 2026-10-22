import { useState } from "react";
import type { McpConnectionStatus, McpProposal } from "../shared/mcp";
import type { AppSnapshot } from "../shared/types";

const FIELD_NAMES: Record<string, string> = {
  title: "需求标题", description: "需求描述", domainL0Id: "领域 L0", domainId: "领域 L1",
  scope: "适用范围", allowLatestVersionUpgrade: "允许最近版本升级", overseasRegions: "海外区域",
  category: "分类", targetMonth: "上线年月", productIds: "匹配产品", planningMemberId: "规划",
  productManagerMemberId: "产品经理", seMemberId: "SE", algorithmMemberId: "算法",
  deviceWorkloadPm: "设备工作量", appWorkloadPm: "运动健康 App 工作量",
  researchAppWorkloadPm: "创新研究 App 工作量", cloudWorkloadPm: "云侧工作量",
  algorithmWorkloadPm: "算法工作量", "整条需求": "删除整条需求", "复制需求": "创建副本",
};

function display(field: string, value: unknown, snapshot: AppSnapshot): string {
  if (value === undefined) return "—";
  if (typeof value === "boolean") return value ? "是" : "否";
  if (["domainL0Id", "domainId"].includes(field) && typeof value === "string") return snapshot.domains.find((item) => item.id === value)?.name ?? value;
  if (["planningMemberId", "productManagerMemberId", "seMemberId", "algorithmMemberId"].includes(field) && typeof value === "string") return snapshot.members.find((item) => item.id === value)?.name ?? value;
  if (field === "productIds" && Array.isArray(value)) return value.map((id) => snapshot.products.find((item) => item.id === id)?.name ?? id).join("、") || "（空）";
  if (typeof value === "string") return value || "（空）";
  return JSON.stringify(value);
}

export function McpPage({ status, proposals, snapshot, onToggle, onDecide }: {
  status: McpConnectionStatus | null;
  proposals: McpProposal[];
  snapshot: AppSnapshot;
  onToggle: (enabled: boolean) => Promise<void>;
  onDecide: (proposal: McpProposal, accept: boolean) => void;
}) {
  const [busy, setBusy] = useState(false);
  return <div className="mcp-page">
    <section className="panel mcp-intro">
      <div><h2>本机模型接入</h2><p>仅向本机 MCP 客户端开放需求池查询与修改提案。模型无法直接写入；每批修改都需你在这里确认。</p>
        <p>不开放路标匹配、分层支持结论、字典、人员管理和 PPT 权限修改。重启后待确认提案作废。</p></div>
      <button type="button" className={`button ${status?.enabled ? "ghost" : "primary"}`} disabled={busy || !status}
        onClick={async () => { setBusy(true); try { await onToggle(!status?.enabled); } finally { setBusy(false); } }}>
        {status?.enabled ? "停用 MCP 接入" : "启用 MCP 接入"}
      </button>
    </section>
    <section className="panel mcp-proposals">
      <div className="panel-title"><div><h3>待确认的 MCP 修改</h3><p>提案 30 分钟后自动失效；确认时重新检查需求版本和字段有效性，整批事务提交。</p></div><span>{proposals.length} 批</span></div>
      {!proposals.length && <p className="mcp-empty">当前没有待确认的修改。</p>}
      {proposals.map((proposal) => {
        const deleteCount = proposal.items.filter((item) => item.operation.type === "delete").length;
        return <article className="mcp-proposal" key={proposal.id}>
          <div className="mcp-proposal-head"><strong>{proposal.items.length} 项修改{deleteCount ? ` · 删除 ${deleteCount} 条` : ""}</strong><small>到期：{new Date(proposal.expiresAt).toLocaleString("zh-CN")}</small></div>
          {proposal.items.map((item, index) => <div className="mcp-proposal-item" key={index}>
            <b>{index + 1}. {item.operation.type === "create" ? "新增" : item.operation.type === "update" ? "修改" : item.operation.type === "duplicate" ? "复制" : "删除"}：{item.title}</b>
            <table><thead><tr><th>字段</th><th>当前</th><th>提议</th></tr></thead><tbody>{item.differences.map((difference) =>
              <tr key={difference.field}><td>{FIELD_NAMES[difference.field] ?? difference.field}</td><td>{display(difference.field, difference.before, snapshot)}</td><td>{display(difference.field, difference.after, snapshot)}</td></tr>)}</tbody></table>
          </div>)}
          <div className="mcp-proposal-actions"><button type="button" className="button ghost" onClick={() => onDecide(proposal, false)}>拒绝整批</button>
            <button type="button" className="button primary" onClick={() => onDecide(proposal, true)}>确认并写入整批</button></div>
        </article>;
      })}
    </section>
  </div>;
}
