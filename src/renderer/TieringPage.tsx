import { useEffect, useMemo, useState, type FormEvent } from "react";
import {
  TIERING_SUPPORT_STATUSES,
  type AppSnapshot,
  type Requirement,
  type SaveTieringSupportInput,
  type TieringChangeLog,
  type TieringSupportRecord,
  type TieringSupportStatus,
  type UpdateTieringChangeLogInput,
} from "../shared/types";
import { EMPTY_TIERING_DATA, isHistoricalTieringFeature, normalizeTieringData, resolveTieringFeature, tieringSupportId } from "../shared/tiering";
import { formatHistoricalLaunch } from "../shared/tieringLaunches";
import { isMemberFilterValid, UNASSIGNED_MEMBER, visibleMembersForRole } from "../shared/members";
import { listReleaseVersionOptions } from "../shared/requirements";
import { shortHalfYear as shortHalf } from "../shared/roadmap";
import { tieringMatchesOwnersAndVersion } from "../shared/tieringFilters";

type View = "matrix" | "logs";

interface Props {
  snapshot: AppSnapshot;
  onSaveSupport: (input: SaveTieringSupportInput) => Promise<void>;
  onOpenRequirement: (requirement: Requirement) => void;
  onRegenerateRequirement: (featureId: string) => Promise<void>;
  onUpdateChangeLog: (input: UpdateTieringChangeLogInput) => Promise<void>;
  onDeleteChangeLog: (log: TieringChangeLog, trigger: HTMLElement) => void;
}

const SERIES_SESSION_KEY = "tiering-visible-series";

export function TieringPage({ snapshot, onSaveSupport, onOpenRequirement, onRegenerateRequirement, onUpdateChangeLog, onDeleteChangeLog }: Props) {
  const tiering = normalizeTieringData(snapshot.tiering ?? EMPTY_TIERING_DATA);
  const resolvedFeatures = tiering.features.map((feature) => resolveTieringFeature(feature, snapshot.requirements, snapshot.domains));
  tiering.features = resolvedFeatures.map((item) => item.feature);
  const allSeriesIds = tiering.productSeries.map((item) => item.id);
  const [view, setView] = useState<View>("matrix");
  const [query, setQuery] = useState("");
  const [l0, setL0] = useState("");
  const [l1, setL1] = useState("");
  const [status, setStatus] = useState<TieringSupportStatus | "">("");
  const [planningMemberId, setPlanningMemberId] = useState("");
  const [productManagerMemberId, setProductManagerMemberId] = useState("");
  const [halfYear, setHalfYear] = useState("");
  const planningMembers = visibleMembersForRole(snapshot.members, snapshot.requirements, "规划");
  const productManagers = visibleMembersForRole(snapshot.members, snapshot.requirements, "产品经理");
  const requirementByFeature = new Map(resolvedFeatures.map((item) => [item.feature.id, item.requirement]));
  const versions = listReleaseVersionOptions(resolvedFeatures.flatMap((item) => {
    const targetMonth = item.requirement?.targetMonth ?? item.feature.historicalLaunch?.suggestedTargetMonth;
    return targetMonth ? [{ targetMonth }] : [];
  }));
  useEffect(() => {
    if (!isMemberFilterValid(planningMemberId, planningMembers)) setPlanningMemberId("");
    if (!isMemberFilterValid(productManagerMemberId, productManagers)) setProductManagerMemberId("");
    if (halfYear && !versions.some((item) => item.halfYear === halfYear)) setHalfYear("");
  }, [planningMemberId, productManagerMemberId, halfYear, planningMembers.map((item) => item.id).join("|"), productManagers.map((item) => item.id).join("|"), versions.map((item) => item.halfYear).join("|")]);
  const [busyRecordId, setBusyRecordId] = useState<string | null>(null);
  const [editingLog, setEditingLog] = useState<TieringChangeLog | null>(null);
  const [selectedSeriesIds, setSelectedSeriesIds] = useState<string[]>(() => readSeriesSelection(allSeriesIds));
  const selectedSeriesSet = useMemo(() => new Set(selectedSeriesIds), [selectedSeriesIds]);
  const selectedSeries = tiering.productSeries.filter((item) => selectedSeriesSet.has(item.id));
  const l0Options = [...new Set(tiering.features.map((item) => item.l0))];
  const l1Options = [...new Set(tiering.features.filter((item) => !l0 || item.l0 === l0).map((item) => item.l1))];
  const supportByKey = useMemo(() => new Map(tiering.supportRecords.map((item) => [item.id, item])), [tiering.supportRecords]);
  const normalizedQuery = query.trim().toLocaleLowerCase("zh-CN");

  useEffect(() => {
    if (l0 && !l0Options.includes(l0)) setL0("");
    if (l1 && !l1Options.includes(l1)) setL1("");
  }, [l0, l1, l0Options.join("|"), l1Options.join("|")]);

  useEffect(() => {
    const valid = selectedSeriesIds.filter((id) => allSeriesIds.includes(id));
    if (valid.length !== selectedSeriesIds.length) setSelectedSeriesIds(valid);
  }, [allSeriesIds.join("|"), selectedSeriesIds.join("|")]);
  useEffect(() => {
    sessionStorage.setItem(SERIES_SESSION_KEY, JSON.stringify(selectedSeriesIds));
    if (!selectedSeriesIds.length && status) setStatus("");
  }, [selectedSeriesIds.join("|"), status]);

  const visibleFeatures = tiering.features.filter((feature) => {
    const scopedRecords = selectedSeries.map((series) => supportByKey.get(tieringSupportId(feature.id, series.id))).filter(Boolean) as TieringSupportRecord[];
    return (!normalizedQuery || `${feature.id} ${feature.l0} ${feature.l1} ${feature.requirementTitle}`.toLocaleLowerCase("zh-CN").includes(normalizedQuery))
      && (!l0 || feature.l0 === l0)
      && (!l1 || feature.l1 === l1)
      && tieringMatchesOwnersAndVersion(feature, requirementByFeature.get(feature.id), snapshot.members, { planningMemberId, productManagerMemberId, halfYear })
      && (!status || scopedRecords.some((item) => item.status === status));
  });

  async function changeSupport(record: TieringSupportRecord, nextStatus: TieringSupportStatus) {
    if (busyRecordId) return;
    setBusyRecordId(record.id);
    try {
      await onSaveSupport({ recordId: record.id, status: nextStatus });
    } finally {
      setBusyRecordId(null);
    }
  }

  function toggleSeries(seriesId: string) {
    setSelectedSeriesIds((current) => current.includes(seriesId) ? current.filter((id) => id !== seriesId) : [...current, seriesId]);
  }

  return <div className="tiering-page">
    <div className="tiering-tabs" role="tablist" aria-label="分层分级视图">
      {([['matrix', '支持矩阵'], ['logs', '变更记录']] as Array<[View, string]>).map(([id, label]) => <button key={id} className={view === id ? "active" : ""} onClick={() => setView(id)}>{label}<b>{id === "matrix" ? tiering.features.length : tiering.changeLogs.length}</b></button>)}
    </div>
    {view === "matrix" && <>
      {resolvedFeatures.some((item) => item.invalidLink) && <p className="tiering-no-series" role="status">部分规格的需求或领域关联无效，已保留原规格及支持状态。请在需求池修复领域，或在分层分级恢复历史需求关联。</p>}
      <div className="panel tiering-filterbar">
        <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="搜索功能 ID、L0、L1或需求" />
        <select value={l0} onChange={(event) => { setL0(event.target.value); setL1(""); }}><option value="">全部 L0</option>{l0Options.map((item) => <option key={item}>{item}</option>)}</select>
        <select value={l1} onChange={(event) => setL1(event.target.value)}><option value="">全部 L1</option>{l1Options.map((item) => <option key={item}>{item}</option>)}</select>
        <select value={status} disabled={!selectedSeries.length} onChange={(event) => setStatus(event.target.value as TieringSupportStatus | "")}><option value="">全部支持状态</option>{TIERING_SUPPORT_STATUSES.map((item) => <option key={item}>{item}</option>)}</select>
        <details className="tiering-series-picker">
          <summary>显示系列 {selectedSeries.length} / {tiering.productSeries.length}</summary>
          <div>
            <header><b>选择展示系列</b><span><button type="button" onClick={() => setSelectedSeriesIds(allSeriesIds)}>全选</button><button type="button" onClick={() => setSelectedSeriesIds([])}>清空</button></span></header>
            {tiering.productSeries.map((series) => <label key={series.id}><input type="checkbox" checked={selectedSeriesSet.has(series.id)} onChange={() => toggleSeries(series.id)} /><span>{series.name}</span></label>)}
          </div>
        </details>
        <select aria-label="分层筛选规划" value={planningMemberId} onChange={(event) => setPlanningMemberId(event.target.value)}><option value="">全部规划</option><option value={UNASSIGNED_MEMBER}>未配置规划</option>{planningMembers.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select>
        <select aria-label="分层筛选产品经理" value={productManagerMemberId} onChange={(event) => setProductManagerMemberId(event.target.value)}><option value="">全部产品经理</option><option value={UNASSIGNED_MEMBER}>未配置产品经理</option>{productManagers.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select>
        <select aria-label="分层筛选版本" value={halfYear} onChange={(event) => setHalfYear(event.target.value)}><option value="">全部版本</option>{versions.map((item) => <option key={item.halfYear} value={item.halfYear}>{item.version}（{shortHalf(item.halfYear)}）</option>)}</select>
        <span>显示 {visibleFeatures.length} / {tiering.features.length}</span>
      </div>
      <div className="panel tiering-matrix-wrap">
        <table className="tiering-matrix">
          <thead><tr><th className="sticky c1">L0</th><th className="sticky c2">L1</th><th className="sticky c3">需求</th>{selectedSeries.map((series) => <th key={series.id}>{series.name}</th>)}</tr></thead>
          <tbody>{visibleFeatures.map((feature) => {
            const requirement = snapshot.requirements.find((item) => item.id === feature.requirementId || item.tieringFeatureId === feature.id);
            return <tr key={feature.id}>
              <td className="sticky c1">{feature.l0}</td>
              <td className="sticky c2">{feature.l1}</td>
              <td className="sticky c3"><div className="tiering-demand-cell"><strong>{feature.requirementTitle}</strong><small>{feature.id}{feature.historicalLaunch ? ` · ${formatHistoricalLaunch(feature.historicalLaunch)}` : ""}</small>{requirement ? <button onClick={() => onOpenRequirement(requirement)}>查看需求</button> : isHistoricalTieringFeature(feature) ? <button className="warning" onClick={() => void onRegenerateRequirement(feature.id)}>重新生成历史需求</button> : <span className="tiering-unlinked">需求已删除</span>}</div></td>
              {selectedSeries.map((series) => {
                const record = supportByKey.get(tieringSupportId(feature.id, series.id));
                return <td key={series.id}>{record ? <SupportStateControl record={record} busy={busyRecordId === record.id} onChange={(nextStatus) => void changeSupport(record, nextStatus)} /> : <span className="tiering-missing-status">未判断</span>}</td>;
              })}
            </tr>;
          })}</tbody>
        </table>
        {!selectedSeries.length && <div className="tiering-no-series">尚未选择展示系列，可从上方“显示系列”中勾选。</div>}
        {selectedSeries.length > 0 && !visibleFeatures.length && <div className="tiering-no-series">暂无符合条件的需求；13 个内置系列已就绪，保存需求后会自动生成支持规格。</div>}
      </div>
    </>}
    {view === "logs" && <div className="panel tiering-ledger"><div className="tiering-section-title"><div><h2>变更记录</h2><p>支持编辑管理信息或永久删除记录；操作不会改变当前系列支持状态。</p></div></div><table><thead><tr><th>时间</th><th>触发类型</th><th>功能 / 系列</th><th>原因</th><th>审核人</th><th>记录状态</th><th>操作</th></tr></thead><tbody>{tiering.changeLogs.map((log) => <tr key={log.id}><td>{formatDate(log.changedAt)}</td><td>{log.triggerType}</td><td>{log.featureId} / {log.seriesId || "—"}</td><td>{log.reason}</td><td>{log.reviewer || "—"}</td><td>{log.recordStatus}</td><td><div className="row-actions"><button type="button" onClick={() => setEditingLog(log)}>编辑</button><button type="button" className="danger" onClick={(event) => onDeleteChangeLog(log, event.currentTarget)}>删除</button></div></td></tr>)}</tbody></table></div>}
    {editingLog && <ChangeLogEditor log={editingLog} onClose={() => setEditingLog(null)} onSave={async (input) => { await onUpdateChangeLog(input); setEditingLog(null); }} />}
  </div>;
}

function ChangeLogEditor({ log, onClose, onSave }: { log: TieringChangeLog; onClose: () => void; onSave: (input: UpdateTieringChangeLogInput) => Promise<void> }) {
  const [reason, setReason] = useState(log.reason);
  const [reviewer, setReviewer] = useState(log.reviewer);
  const [recordStatus, setRecordStatus] = useState(log.recordStatus || "有效");
  const [note, setNote] = useState(log.note);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  async function submit(event: FormEvent) {
    event.preventDefault();
    if (busy) return;
    setBusy(true);
    setError("");
    try {
      await onSave({ logId: log.id, reason, reviewer, recordStatus, note });
    } catch (value) {
      setError(value instanceof Error ? value.message : String(value));
    } finally {
      setBusy(false);
    }
  }
  return <div className="overlay" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget && !busy) onClose(); }}>
    <form className="drawer tiering-log-editor" role="dialog" aria-modal="true" aria-labelledby="tiering-log-editor-title" onSubmit={(event) => void submit(event)}>
      <div className="drawer-header"><div><h2 id="tiering-log-editor-title">编辑变更记录</h2><p>{log.featureId} / {log.seriesId || "—"}</p></div><button type="button" className="icon-button" aria-label="关闭" disabled={busy} onClick={onClose}>×</button></div>
      <div className="form-body">
        <div className="tiering-log-readonly"><span>变更时间<b>{formatDate(log.changedAt)}</b></span><span>触发类型<b>{log.triggerType}</b></span><span>变更前<b>{log.before || "—"}</b></span><span>变更后<b>{log.after || "—"}</b></span></div>
        <label className="field"><span>变更原因 *</span><textarea value={reason} maxLength={500} required onChange={(event) => setReason(event.target.value)} /></label>
        <div className="field-row"><label className="field"><span>审核人</span><input value={reviewer} maxLength={100} onChange={(event) => setReviewer(event.target.value)} /></label><label className="field"><span>记录状态</span><select value={recordStatus} onChange={(event) => setRecordStatus(event.target.value)}><option>有效</option><option>待审核</option><option>已归档</option></select></label></div>
        <label className="field"><span>备注</span><textarea value={note} maxLength={1000} onChange={(event) => setNote(event.target.value)} /></label>
        {error && <p className="form-error" role="alert">{error}</p>}
      </div>
      <div className="drawer-footer"><span>编辑只更新管理信息，不改变支持状态。</span><div><button type="button" className="button" disabled={busy} onClick={onClose}>取消</button><button type="submit" className="button primary" disabled={busy}>{busy ? "保存中…" : "保存"}</button></div></div>
    </form>
  </div>;
}

function SupportStateControl({ record, busy, onChange }: { record: TieringSupportRecord; busy: boolean; onChange: (status: TieringSupportStatus) => void }) {
  return <div className={`tiering-state-control ${tone(record.status)}${busy ? " busy" : ""}`} aria-label={`当前状态：${record.status}`}>
    <button type="button" disabled={busy} aria-pressed={record.status === "支持"} className={record.status === "支持" ? "active support" : "support"} onClick={() => onChange(record.status === "支持" ? "未判断" : "支持")}>支持</button>
    <button type="button" disabled={busy} aria-pressed={record.status === "不支持"} className={record.status === "不支持" ? "active unsupported" : "unsupported"} onClick={() => onChange(record.status === "不支持" ? "未判断" : "不支持")}>不支持</button>
    <small>{busy ? "保存中…" : record.status}</small>
  </div>;
}

function readSeriesSelection(allSeriesIds: string[]): string[] {
  try {
    const saved = JSON.parse(sessionStorage.getItem(SERIES_SESSION_KEY) ?? "null") as unknown;
    if (Array.isArray(saved) && saved.every((item) => typeof item === "string")) return saved.filter((item) => allSeriesIds.includes(item));
  } catch {
    // Ignore stale session state and fall back to all series.
  }
  return allSeriesIds;
}

function tone(status: TieringSupportStatus): string {
  return status === "支持" ? "supported" : status === "不支持" ? "unsupported" : "pending";
}

function formatDate(value: string): string {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : new Intl.DateTimeFormat("zh-CN", { dateStyle: "medium", timeStyle: "short" }).format(date);
}
