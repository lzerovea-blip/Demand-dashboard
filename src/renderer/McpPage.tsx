import { useEffect, useState } from "react";
import { MCP_EVENT_NAMES, MCP_STAGE_NAMES, type McpDiagnosticsSnapshot, type McpSelfTestResult } from "../shared/mcpDiagnostics";
import type { McpConnectionStatus, McpProposal } from "../shared/mcp";
import type { AppSnapshot } from "../shared/types";
import type { McpLaunchConfiguration } from "../shared/mcpLaunch";

const FIELD_NAMES: Record<string, string> = {
  title: "需求标题", description: "需求描述", domainL0Id: "领域 L0", domainId: "领域 L1",
  scope: "适用范围", allowLatestVersionUpgrade: "允许最近版本升级", overseasRegions: "海外区域",
  category: "分类", targetMonth: "上线年月", productIds: "匹配产品", planningMemberId: "规划",
  productManagerMemberId: "产品经理", seMemberId: "SE", algorithmMemberId: "算法",
  deviceWorkloadPm: "设备工作量", appWorkloadPm: "运动健康 App 工作量",
  researchAppWorkloadPm: "创新研究 App 工作量", cloudWorkloadPm: "云侧工作量",
  algorithmWorkloadPm: "算法工作量", "整条需求": "删除整条需求", "复制需求": "创建副本",
  name: "名称", parentId: "所属领域 L0", employeeNumber: "工号", microsoftAccount: "企业邮箱",
  role: "岗位角色", domainL1Ids: "负责领域 L1", roadmapKind: "展示类型", roadmapId: "所属路标",
  exportFormat: "导出类型", exportRange: "导出区间", exportCount: "区间内需求数", materialName: "材料名称", exportSections: "PPT 选择页", exportWarning: "导出范围与安全提示",
  exportDraftCount: "草稿需求数",
};

function display(field: string, value: unknown, snapshot: AppSnapshot): string {
  if (value === undefined) return "—";
  if (typeof value === "boolean") return value ? "是" : "否";
  if (["domainL0Id", "domainId", "parentId"].includes(field) && typeof value === "string") return snapshot.domains.find((item) => item.id === value)?.name ?? value;
  if (field === "domainL1Ids" && Array.isArray(value)) return value.map(id => snapshot.domains.find(item => item.id === id)?.name ?? id).join("、") || "（空）";
  if (field === "roadmapId") return snapshot.roadmaps?.find(item => item.id === value)?.name ?? String(value);
  if (field === "roadmapKind") return value === "medical" ? "医疗认证时间轴" : "普通卡片";
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
  const [diagnostics, setDiagnostics] = useState<McpDiagnosticsSnapshot | null>(null);
  const [diagnosticMessage, setDiagnosticMessage] = useState("");
  const [selfTest, setSelfTest] = useState<McpSelfTestResult | null>(null);
  const [testing, setTesting] = useState(false);
  const [launch, setLaunch] = useState<McpLaunchConfiguration | null>(null);
  useEffect(() => { let disposed = false; void window.roadmapApi.getMcpLaunchConfiguration().then(value => { if (!disposed) setLaunch(value); }).catch(() => { if (!disposed) setDiagnosticMessage("读取独立 MCP 配置失败"); }); return () => { disposed = true; }; }, []);
  useEffect(() => {
    let disposed = false, loading = false;
    const refresh = async () => {
      if (loading || document.visibilityState === "hidden") return;
      loading = true;
      try { const result = await window.roadmapApi.getMcpDiagnostics(); if (!disposed) setDiagnostics(result); }
      catch { if (!disposed) setDiagnosticMessage("读取 MCP 诊断失败，请查看日志目录权限。"); }
      finally { loading = false; }
    };
    void refresh();
    const interval = setInterval(() => { void refresh(); }, 1000);
    document.addEventListener("visibilitychange", refresh);
    return () => { disposed = true; clearInterval(interval); document.removeEventListener("visibilitychange", refresh); };
  }, []);
  const external = diagnostics?.sessions.find(item => item.role === "stdio" && item.source === "external");
  const action = async (run: () => Promise<unknown>, success = "") => {
    try { await run(); setDiagnosticMessage(success); } catch { setDiagnosticMessage("操作失败，请检查 MCP 日志目录权限。"); }
  };
  return <div className="mcp-page">
    <section className="panel mcp-intro">
      <div><h2>本机模型接入</h2><p>仅向本机 MCP 客户端开放需求池查询与修改提案。模型无法直接写入；每批修改都需你在这里确认。</p>
        <p>支持字典新增及 PPT、需求池 Excel、分层分级 Excel、完整需求包导出提案。导出须确认范围并由你选择保存位置；PPT 权限不对模型开放。重启后待确认提案作废。</p></div>
      <button type="button" className={`button ${status?.enabled ? "ghost" : "primary"}`} disabled={busy || !status}
        onClick={async () => { setBusy(true); try { await onToggle(!status?.enabled); } finally { setBusy(false); } }}>
        {status?.enabled ? "停用 MCP 接入" : "启用 MCP 接入"}
      </button>
    </section>
    <section className="panel mcp-cli-configuration">
      <h3>独立 MCP 连接配置</h3>
      <p>传输类型选择 stdio。命令填写下方 Node 可执行文件，参数填写脚本完整路径；不再使用 --mcp-stdio。保持工作台打开并启用接入。</p>
      <p>命令：<code>{launch?.command ?? "读取中…"}</code></p>
      <p>参数（单独一项）：<code>{launch?.args[0] ?? "读取中…"}</code></p>
      {launch?.error && <p role="alert">{launch.error}</p>}
      <button className="button ghost" disabled={!launch?.available} onClick={() => void action(() => window.roadmapApi.copyMcpLaunchConfiguration(), "独立 MCP 配置已复制")}>复制连接配置</button>
    </section>
    <section className="panel mcp-diagnostics">
      <div className="panel-title"><div><h3>MCP 连接诊断</h3><p>自动记录程序启动、握手、工具发现和查询。响应已发送仅表示交给标准输出。</p></div>
        <span>{status?.enabled ? "接入已启用" : "接入未启用"}</span></div>
      {diagnostics?.warning && <p role="alert">{diagnostics.warning}</p>}
      <div className="mcp-diagnostic-status">
        <div><small>外部客户端</small><strong>{external?.clientName ?? "客户端身份未知"}{external?.clientVersion ? ` · ${external.clientVersion}` : ""}</strong></div>
        <div><small>最后观察阶段</small><strong>{external ? `${MCP_STAGE_NAMES[external.stage] ?? external.stage} · ${MCP_EVENT_NAMES[external.event] ?? external.event}` : "尚未观察到程序启动，需结合客户端日志检查"}</strong></div>
        <div><small>最后活动</small><strong>{external ? new Date(external.lastActivity).toLocaleString("zh-CN") : "—"}</strong></div>
        <div><small>最近异常阶段</small><strong>{external?.lastFailure ? `${MCP_STAGE_NAMES[external.lastFailure.stage] ?? external.lastFailure.stage} · ${external.lastFailure.errorCode ?? "等待初始化"}` : "未记录异常"}</strong></div>
      </div>
      <div className="mcp-diagnostic-actions">
        <button className="button ghost" disabled={testing || !status?.enabled} onClick={async () => {
          setTesting(true); try { setSelfTest(await window.roadmapApi.runMcpSelfTest()); }
          catch { setDiagnosticMessage("本机自检启动失败，请查看诊断日志。"); } finally { setTesting(false); }
        }}>{testing ? "正在自检…" : "本机连接自检"}</button>
        <button className="button ghost" onClick={() => void action(() => window.roadmapApi.copyMcpDiagnosticSummary(), "诊断摘要已复制")}>复制诊断摘要</button>
        <button className="button ghost" onClick={() => void action(() => window.roadmapApi.openMcpDiagnosticDirectory())}>打开 MCP 日志目录</button>
        <button className="button ghost" onClick={() => void action(() => window.roadmapApi.exportMcpDiagnostics())}>导出诊断日志</button>
      </div>
      <p className="mcp-diagnostic-hint">本机自检仅验证当前程序和工作台链路，不计作外部客户端连接；不会写入需求。</p>
      {selfTest && <p role="status">{selfTest.success ? "✓" : "!"} {selfTest.message}</p>}
      {diagnosticMessage && <p role="status">{diagnosticMessage}</p>}
      <div className="mcp-diagnostic-events"><table><thead><tr><th>时间</th><th>来源</th><th>阶段／事件</th><th>工具／结果</th></tr></thead><tbody>
        {diagnostics?.events.map((event, index) => <tr key={`${event.sessionId}-${event.time}-${index}`} className={`mcp-event-${event.level}`}>
          <td>{new Date(event.time).toLocaleTimeString("zh-CN")}</td><td>{event.source === "self-test" ? "内部自检" : "外部接入"} · {event.role === "stdio" ? "MCP进程" : "工作台"}</td>
          <td>{MCP_STAGE_NAMES[event.stage] ?? event.stage} · {MCP_EVENT_NAMES[event.event] ?? event.event}</td>
          <td>{event.tool ?? event.method ?? ""}{event.errorCode ? ` · ${event.errorCode}` : ""}{event.durationMs !== undefined ? ` · ${event.durationMs}ms` : ""}{event.count !== undefined ? ` · ${event.count}项` : ""}</td>
        </tr>)}
      </tbody></table></div>
    </section>
    <section className="panel mcp-proposals">
      <div className="panel-title"><div><h3>待确认的 MCP 修改</h3><p>提案 30 分钟后自动失效；确认时重新检查需求版本和字段有效性，整批事务提交。</p></div><span>{proposals.length} 批</span></div>
      {!proposals.length && <p className="mcp-empty">当前没有待确认的修改。</p>}
      {proposals.map((proposal) => {
        const deleteCount = proposal.items.filter((item) => item.operation.type === "delete").length;
        return <article className="mcp-proposal" key={proposal.id}>
          <div className="mcp-proposal-head"><strong>{proposal.items.length} 项修改{deleteCount ? ` · 删除 ${deleteCount} 条` : ""}</strong><small>到期：{new Date(proposal.expiresAt).toLocaleString("zh-CN")}</small></div>
          {proposal.items.map((item, index) => <div className="mcp-proposal-item" key={index}>
            <b>{index + 1}. {item.operation.type === "export" ? "文件导出" : item.operation.type === "create_dictionary" ? "字典新增" : item.operation.type === "create" ? "新增" : item.operation.type === "update" ? "修改" : item.operation.type === "duplicate" ? "复制" : "删除"}：{item.title}</b>
            <table><thead><tr><th>字段</th><th>当前</th><th>提议</th></tr></thead><tbody>{item.differences.map((difference) =>
              <tr key={difference.field}><td>{FIELD_NAMES[difference.field] ?? difference.field}</td><td>{display(difference.field, difference.before, snapshot)}</td><td>{display(difference.field, difference.after, snapshot)}</td></tr>)}</tbody></table>
          </div>)}
          <div className="mcp-proposal-actions"><button type="button" className="button ghost" onClick={() => onDecide(proposal, false)}>拒绝整批</button>
            <button type="button" className="button primary" onClick={() => onDecide(proposal, true)}>{proposal.items.some(item => item.operation.type === "export") ? "确认导出" : "确认并写入整批"}</button></div>
        </article>;
      })}
    </section>
  </div>;
}
