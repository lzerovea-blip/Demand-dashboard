import { useEffect, useMemo, useState } from "react";
import type { AppSnapshot, MedicalRoadmapPlan, Requirement, SaveRoadmapAssignmentsInput, SaveRoadmapInput, SaveRoadmapLaneInput } from "../shared/types";
import { MEDICAL_PLAN_LABELS, MEDICAL_PLAN_MONTH_KEYS, roadmapDefinitionsFor, roadmapLanesFor } from "../shared/roadmapManagement";
import { listReleaseVersionOptions } from "../shared/requirements";
import { halfYearOf } from "../shared/roadmap";

type Confirm = (input: { title: string; message: string; confirmLabel: string; action: () => Promise<void> }) => void;

export function RoadmapMatchingPage({ snapshot, onSave, onOpenRequirement, onMedical, confirm }: {
  snapshot: AppSnapshot; onSave: (input: SaveRoadmapAssignmentsInput) => Promise<void>;
  onOpenRequirement: (requirement: Requirement) => void; onMedical: (requirement: Requirement, roadmapId: string) => void; confirm: Confirm;
}) {
  const [query, setQuery] = useState(""); const [l0, setL0] = useState(""); const [l1, setL1] = useState("");
  const [version, setVersion] = useState(""); const [status, setStatus] = useState("");
  const [roadmapId, setRoadmapId] = useState(""); const [laneId, setLaneId] = useState("");
  const [selected, setSelected] = useState<string[]>([]); const [busy, setBusy] = useState(false);
  const roadmaps = roadmapDefinitionsFor(snapshot); const lanes = roadmapLanesFor(snapshot, roadmapId);
  const names = new Map(snapshot.domains.map((item) => [item.id, item.name]));
  const assignments = new Map(snapshot.roadmapAssignments?.map((item) => [item.requirementId, item]));
  const children = snapshot.domains.filter((item) => item.level === "L1" && (!l0 || item.parentId === l0));
  useEffect(() => { if (l1 && !children.some((item) => item.id === l1)) setL1(""); }, [l0, l1, snapshot.domains]);
  useEffect(() => { if (!lanes.some((item) => item.id === laneId)) setLaneId(""); }, [roadmapId, laneId, snapshot.roadmapLanes]);
  useEffect(() => { setSelected((old) => old.filter((id) => snapshot.requirements.some((item) => item.id === id))); }, [snapshot.requirements]);
  const rows = useMemo(() => snapshot.requirements.filter((req) => {
    let matchVersion = true; if (version) { try { matchVersion = halfYearOf(req.targetMonth) === version; } catch { matchVersion = false; } }
    return (!query || `${req.title} ${names.get(req.domainId) ?? ""} ${names.get(req.domainL0Id) ?? ""}`.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase())) && (!l0 || req.domainL0Id === l0) && (!l1 || req.domainId === l1) && matchVersion && (!status || (status === "unmatched" ? !assignments.has(req.id) : assignments.has(req.id)));
  }), [snapshot, query, l0, l1, version, status]);
  const describe = (id: string) => { const a = assignments.get(id); return a ? `${roadmaps.find((item) => item.id === a.roadmapId)?.name ?? a.roadmapId} · ${snapshot.roadmapLanes?.find((item) => item.id === a.laneId)?.name ?? a.laneId}` : "未匹配"; };
  async function save(remove = false) {
    if (busy || !selected.length) return;
    const action = async () => { setBusy(true); try { await onSave({ requirementIds: selected, ...(remove ? {} : { roadmapId, laneId }) }); setSelected([]); } finally { setBusy(false); } };
    const changed = selected.filter((id) => assignments.has(id) && (remove || assignments.get(id)!.roadmapId !== roadmapId || assignments.get(id)!.laneId !== laneId));
    if (changed.length) confirm({ title: remove ? "确认移出路标" : "确认调整路标归属", message: changed.map((id) => `${snapshot.requirements.find((req) => req.id === id)?.title}：${describe(id)} → ${remove ? "未匹配" : `${roadmaps.find((map) => map.id === roadmapId)?.name} · ${lanes.find((lane) => lane.id === laneId)?.name}`}`).join("\n"), confirmLabel: remove ? "移出路标" : "保存匹配", action });
    else await action();
  }
  return <div className="panel roadmap-matching">
    <div className="panel-title"><div><h2>需求匹配</h2><p>先录入需求，再勾选匹配到一个路标和一个泳道。</p></div></div>
    <div className="matching-filters">
      <input placeholder="搜索需求或领域" value={query} onChange={(e) => setQuery(e.target.value)} />
      <select value={l0} onChange={(e) => setL0(e.target.value)}><option value="">全部领域 L0</option>{snapshot.domains.filter((item) => item.level === "L0").map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select>
      <select value={l1} onChange={(e) => setL1(e.target.value)}><option value="">全部领域 L1</option>{children.map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select>
      <select value={version} onChange={(e) => setVersion(e.target.value)}><option value="">全部版本</option>{listReleaseVersionOptions(snapshot.requirements).map((item) => <option key={item.halfYear} value={item.halfYear}>{item.version}（{item.halfYear.slice(2)}）</option>)}</select>
      <select value={status} onChange={(e) => setStatus(e.target.value)}><option value="">全部匹配状态</option><option value="unmatched">未匹配</option><option value="matched">已匹配</option></select>
    </div>
    <div className="matching-actions">
      <select aria-label="目标路标" value={roadmapId} onChange={(e) => setRoadmapId(e.target.value)}><option value="">选择目标路标</option>{roadmaps.map((map) => <option key={map.id} value={map.id}>{map.name}</option>)}</select>
      <select aria-label="目标泳道" value={laneId} disabled={!roadmapId || !lanes.length} onChange={(e) => setLaneId(e.target.value)}><option value="">{roadmapId && !lanes.length ? "请先在字典添加泳道" : "选择目标泳道"}</option>{lanes.map((lane) => <option key={lane.id} value={lane.id}>{lane.name}</option>)}</select>
      <button className="button primary" disabled={busy || !selected.length || !roadmapId || !laneId} onClick={() => void save()}>匹配已选 {selected.length} 条</button>
      <button className="button ghost" disabled={busy || !selected.some((id) => assignments.has(id))} onClick={() => void save(true)}>移出路标</button>
      <button className="button ghost" disabled={busy || !rows.length} onClick={() => setSelected(rows.map((item) => item.id))}>全选当前结果</button>
      <button className="button ghost" disabled={busy || !selected.length} onClick={() => setSelected([])}>取消全选</button>
    </div>
    <div className="matching-table-wrap"><table className="matching-table"><thead><tr><th>勾选</th><th>需求</th><th>适用范围</th><th>上线月份</th><th>当前路标 · 泳道</th><th>操作</th></tr></thead><tbody>
      {rows.map((req) => { const a = assignments.get(req.id); const map = a && roadmaps.find((item) => item.id === a.roadmapId); return <tr key={req.id}><td><input aria-label={`选择${req.title}`} type="checkbox" disabled={busy} checked={selected.includes(req.id)} onChange={(e) => setSelected((old) => e.target.checked ? [...old, req.id] : old.filter((id) => id !== req.id))} /></td><td><strong>{names.get(req.domainId)}：{req.title}</strong></td><td>{req.scope}</td><td>{req.targetMonth || "待排期"}</td><td>{describe(req.id)}</td><td><button className="text-button" onClick={() => onOpenRequirement(req)}>查看需求</button>{map?.kind === "medical" && <button className="text-button" onClick={() => onMedical(req, map.id)}>医疗节点</button>}</td></tr>; })}
    </tbody></table>{!rows.length && <p className="empty">当前条件暂无需求</p>}</div>
    {(snapshot.medicalRoadmapPlans ?? []).filter((plan) => assignments.get(plan.requirementId)?.roadmapId !== plan.roadmapId).length > 0 && <div className="matching-actions"><strong>保留的医疗草稿</strong>{snapshot.medicalRoadmapPlans!.filter((plan) => assignments.get(plan.requirementId)?.roadmapId !== plan.roadmapId).map((plan) => { const req = snapshot.requirements.find((item)=>item.id===plan.requirementId); return req && <button key={`${plan.roadmapId}:${plan.requirementId}`} className="button ghost" onClick={()=>onMedical(req,plan.roadmapId)}>{req.title} · {roadmaps.find((item)=>item.id===plan.roadmapId)?.name}</button>; })}</div>}
  </div>;
}

export function RoadmapDictionaryPanel({ snapshot, onSaveRoadmap, onDeleteRoadmap, onSaveLane, onDeleteLane, confirm }: {
  snapshot: AppSnapshot; onSaveRoadmap: (input: SaveRoadmapInput) => Promise<void>; onDeleteRoadmap: (id: string) => Promise<void>;
  onSaveLane: (input: SaveRoadmapLaneInput) => Promise<void>; onDeleteLane: (id: string) => Promise<void>; confirm: Confirm;
}) {
  const [mapForm, setMapForm] = useState<SaveRoadmapInput>({ name: "", kind: "standard" });
  const [laneForm, setLaneForm] = useState<SaveRoadmapLaneInput>({ name: "", roadmapId: "" });
  const [busy, setBusy] = useState(false);
  const roadmaps = roadmapDefinitionsFor(snapshot);
  const run = async (action: () => Promise<void>) => { setBusy(true); try { await action(); } finally { setBusy(false); } };
  return <section className="panel roadmap-dictionary"><div className="panel-title"><div><h2>路标与泳道字典</h2><p>名称和顺序由此处维护，路标查看及 PPT 同步使用。</p></div></div>
    <form className="matching-actions" onSubmit={(e) => { e.preventDefault(); void run(async () => { await onSaveRoadmap(mapForm); setMapForm({ name: "", kind: "standard" }); }); }}>
      <input required placeholder="路标名称，例如健康路标" value={mapForm.name} onChange={(e) => setMapForm({ ...mapForm, name: e.target.value })} />
      <select value={mapForm.kind} onChange={(e) => setMapForm({ ...mapForm, kind: e.target.value as SaveRoadmapInput["kind"] })}><option value="standard">普通卡片</option><option value="medical">医疗认证时间轴</option></select>
      <input className="order-input" type="number" min="0" placeholder="顺序" value={mapForm.sortOrder ?? ""} onChange={(e) => setMapForm({ ...mapForm, sortOrder: e.target.value === "" ? undefined : Number(e.target.value) })} />
      <button className="button primary" disabled={busy}>{mapForm.id ? "保存路标" : "添加路标"}</button>{mapForm.id && <button className="button ghost" type="button" onClick={() => setMapForm({ name: "", kind: "standard" })}>取消编辑</button>}
    </form>
    <form className="matching-actions" onSubmit={(e) => { e.preventDefault(); void run(async () => { await onSaveLane(laneForm); setLaneForm({ name: "", roadmapId: laneForm.roadmapId }); }); }}>
      <select required value={laneForm.roadmapId} onChange={(e) => setLaneForm({ name: "", roadmapId: e.target.value })}><option value="">选择所属路标</option>{roadmaps.map((map) => <option key={map.id} value={map.id}>{map.name}</option>)}</select>
      <input required placeholder="泳道名称，例如高阶" value={laneForm.name} onChange={(e) => setLaneForm({ ...laneForm, name: e.target.value })} />
      <input className="order-input" type="number" min="0" placeholder="顺序" value={laneForm.sortOrder ?? ""} onChange={(e) => setLaneForm({ ...laneForm, sortOrder: e.target.value === "" ? undefined : Number(e.target.value) })} />
      <button className="button primary" disabled={busy}>{laneForm.id ? "保存泳道" : "添加泳道"}</button>{laneForm.id && <button type="button" className="button ghost" onClick={() => setLaneForm({ name: "", roadmapId: laneForm.roadmapId })}>取消编辑</button>}
    </form>
    <div className="roadmap-dictionary-list">{roadmaps.map((map) => <div key={map.id} className="roadmap-dictionary-item"><div><strong>{map.name}</strong><span>{map.kind === "medical" ? "医疗认证时间轴" : "普通卡片"} · 顺序 {map.sortOrder}</span><button className="text-button" onClick={() => setMapForm(map)}>编辑</button><button className="text-button danger" disabled={busy} onClick={() => confirm({ title: "删除路标", message: `删除“${map.name}”及其未使用泳道？有匹配或医疗配置时会阻止删除。`, confirmLabel: "删除", action: () => run(() => onDeleteRoadmap(map.id)) })}>删除</button></div>
      <div className="roadmap-lane-chips">{roadmapLanesFor(snapshot, map.id).map((lane) => <span key={lane.id}><b>{lane.name}</b><small>顺序 {lane.sortOrder}</small><button className="text-button" onClick={() => setLaneForm(lane)}>编辑</button><button className="text-button danger" disabled={busy} onClick={() => confirm({ title: "删除泳道", message: `删除“${map.name} · ${lane.name}”？`, confirmLabel: "删除", action: () => run(() => onDeleteLane(lane.id)) })}>删除</button></span>)}{!roadmapLanesFor(snapshot,map.id).length && <small>尚未配置泳道</small>}</div>
    </div>)}</div>
  </section>;
}

export function MedicalPlanEditor({ requirement, roadmapId, snapshot, onClose, onSave, onDelete }: {
  requirement: Requirement; roadmapId: string; snapshot: AppSnapshot; onClose: () => void; onSave: (input: MedicalRoadmapPlan) => Promise<void>;
  onDelete?: () => void;
}) {
  const [form, setForm] = useState<MedicalRoadmapPlan>(() => snapshot.medicalRoadmapPlans?.find((item) => item.requirementId === requirement.id && item.roadmapId === roadmapId) ?? { requirementId: requirement.id, roadmapId, updatedAt: "" });
  const [busy, setBusy] = useState(false);
  return <div className="overlay" onMouseDown={(e) => { if (e.target === e.currentTarget && !busy) onClose(); }}><form className="drawer editor-drawer" onSubmit={async (e) => { e.preventDefault(); setBusy(true); try { await onSave(form); onClose(); } finally { setBusy(false); } }}><div className="drawer-header"><div><h2>医疗认证时间节点</h2><p>{requirement.title}</p></div><button type="button" className="icon-button" disabled={busy} onClick={onClose}>×</button></div><div className="form-body"><p>五个节点全部选填，不限制先后顺序。需求上线月份仍为 {requirement.targetMonth || "待排期"}，两处独立维护。</p><div className="medical-milestone-grid">{MEDICAL_PLAN_MONTH_KEYS.map((key,index) => <label key={key}><span>{MEDICAL_PLAN_LABELS[index]}</span><input type="month" value={form[key] ?? ""} onChange={(e) => setForm({ ...form, [key]: e.target.value || undefined })} /></label>)}</div></div><div className="drawer-footer">{onDelete && <button type="button" className="button ghost" onClick={onDelete} disabled={busy}>清除医疗配置</button>}<button type="button" className="button ghost" onClick={onClose} disabled={busy}>取消</button><button className="button primary" disabled={busy}>{busy ? "保存中…" : "保存医疗节点"}</button></div></form></div>;
}
