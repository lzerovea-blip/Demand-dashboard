import { useEffect, useMemo, useState } from "react";
import type { AppSnapshot, MedicalRoadmapPlan, Requirement, RoadmapDefinition, SaveRoadmapAssignmentsInput, SaveRoadmapInput, SaveRoadmapLaneInput } from "../shared/types";
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
      <select aria-label="目标泳道" value={laneId} disabled={!roadmapId || !lanes.length} onChange={(e) => setLaneId(e.target.value)}><option value="">{roadmapId && !lanes.length ? "请先在路标字典设置添加泳道" : "选择目标泳道"}</option>{lanes.map((lane) => <option key={lane.id} value={lane.id}>{lane.name}</option>)}</select>
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
  snapshot: AppSnapshot; onSaveRoadmap: (input: SaveRoadmapInput) => Promise<AppSnapshot>; onDeleteRoadmap: (id: string) => Promise<void>;
  onSaveLane: (input: SaveRoadmapLaneInput) => Promise<void>; onDeleteLane: (id: string) => Promise<void>; confirm: Confirm;
}) {
  const [newMapForm, setNewMapForm] = useState<SaveRoadmapInput>({ name: "", kind: "standard" });
  const [selectedMapId, setSelectedMapId] = useState<string | null>(null);
  const [editMapForm, setEditMapForm] = useState<SaveRoadmapInput | null>(null);
  const [laneForm, setLaneForm] = useState<SaveRoadmapLaneInput>({ name: "", roadmapId: "" });
  const [busy, setBusy] = useState(false);
  const roadmaps = roadmapDefinitionsFor(snapshot);
  const selectedMap = roadmaps.find((map) => map.id === selectedMapId);
  useEffect(() => {
    if (selectedMapId && !selectedMap) {
      setSelectedMapId(null);
      setEditMapForm(null);
      setLaneForm({ name: "", roadmapId: "" });
    }
  }, [selectedMapId, selectedMap]);
  function selectMap(map: RoadmapDefinition) {
    if (selectedMapId === map.id) {
      setSelectedMapId(null);
      setEditMapForm(null);
      setLaneForm({ name: "", roadmapId: "" });
      return;
    }
    setSelectedMapId(map.id);
    setEditMapForm({ id: map.id, name: map.name, kind: map.kind, sortOrder: map.sortOrder });
    setLaneForm({ name: "", roadmapId: map.id });
  }
  const run = async (action: () => Promise<void>) => {
    if (busy) return;
    setBusy(true);
    try { await action(); } catch { /* The application-level update helper shows the error. */ }
    finally { setBusy(false); }
  };
  return <section className="panel roadmap-dictionary"><div className="panel-title"><div><h2>路标字典设置</h2><p>顶部新增路标；在下方选中路标后维护其名称、顺序和泳道。</p></div></div>
    <form className="matching-actions roadmap-dictionary-create" onSubmit={(e) => { e.preventDefault(); void run(async () => {
      const existingIds = new Set(roadmaps.map((map) => map.id));
      const next = await onSaveRoadmap(newMapForm);
      setNewMapForm({ name: "", kind: "standard" });
      const created = roadmapDefinitionsFor(next).find((map) => !existingIds.has(map.id));
      if (created) {
        setSelectedMapId(created.id);
        setEditMapForm({ id: created.id, name: created.name, kind: created.kind, sortOrder: created.sortOrder });
        setLaneForm({ name: "", roadmapId: created.id });
      }
    }); }}>
      <input required aria-label="新增路标名称" placeholder="路标名称，例如健康路标" value={newMapForm.name} onChange={(e) => setNewMapForm({ ...newMapForm, name: e.target.value })} />
      <select aria-label="新增路标展示类型" value={newMapForm.kind} onChange={(e) => setNewMapForm({ ...newMapForm, kind: e.target.value as SaveRoadmapInput["kind"] })}><option value="standard">普通卡片</option><option value="medical">医疗认证时间轴</option></select>
      <input aria-label="新增路标顺序" className="order-input" type="number" min="0" placeholder="顺序" value={newMapForm.sortOrder ?? ""} onChange={(e) => setNewMapForm({ ...newMapForm, sortOrder: e.target.value === "" ? undefined : Number(e.target.value) })} />
      <button className="button primary" disabled={busy}>添加路标</button>
    </form>
    <div className="roadmap-dictionary-list">{roadmaps.map((map) => {
      const expanded = selectedMapId === map.id;
      const lanes = roadmapLanesFor(snapshot, map.id);
      return <div key={map.id} className={`roadmap-dictionary-item${expanded ? " is-selected" : ""}`}>
        <div className="roadmap-dictionary-heading">
          <button type="button" className="roadmap-dictionary-select" disabled={busy} aria-expanded={expanded} aria-controls={expanded ? `roadmap-dictionary-detail-${map.id}` : undefined} onClick={() => selectMap(map)}><strong>{map.name}</strong><span>{map.kind === "medical" ? "医疗认证时间轴" : "普通卡片"} · 顺序 {map.sortOrder} · {lanes.length} 条泳道</span><i aria-hidden="true">{expanded ? "⌃" : "⌄"}</i></button>
          <button type="button" className="text-button danger" disabled={busy} onClick={() => confirm({ title: "删除路标", message: `删除“${map.name}”及其未使用泳道？有匹配或医疗配置时会阻止删除。`, confirmLabel: "删除", action: () => run(() => onDeleteRoadmap(map.id)) })}>删除</button>
        </div>
        {expanded && editMapForm && <div id={`roadmap-dictionary-detail-${map.id}`} className="roadmap-dictionary-detail">
          <form className="matching-actions" onSubmit={(e) => { e.preventDefault(); void run(async () => { await onSaveRoadmap({ ...editMapForm, id: map.id }); }); }}>
            <input required aria-label={`编辑路标名称（${map.name}）`} value={editMapForm.name} onChange={(e) => setEditMapForm({ ...editMapForm, name: e.target.value })} />
            <select aria-label={`${map.name}展示类型`} value={editMapForm.kind} onChange={(e) => setEditMapForm({ ...editMapForm, kind: e.target.value as SaveRoadmapInput["kind"] })}><option value="standard">普通卡片</option><option value="medical">医疗认证时间轴</option></select>
            <input aria-label={`编辑路标顺序（${map.name}）`} className="order-input" type="number" min="0" value={editMapForm.sortOrder ?? ""} onChange={(e) => setEditMapForm({ ...editMapForm, sortOrder: e.target.value === "" ? undefined : Number(e.target.value) })} />
            <button className="button primary" disabled={busy}>保存路标</button>
          </form>
          <div className="roadmap-dictionary-lane-title"><strong>{map.name}的泳道</strong><span>仅管理当前选中的路标</span></div>
          <form className="matching-actions" onSubmit={(e) => { e.preventDefault(); void run(async () => { await onSaveLane({ ...laneForm, roadmapId: map.id }); setLaneForm({ name: "", roadmapId: map.id }); }); }}>
            <input required aria-label={`${map.name}泳道名称`} placeholder="泳道名称，例如高阶" value={laneForm.name} onChange={(e) => setLaneForm({ ...laneForm, name: e.target.value })} />
            <input aria-label={`${map.name}泳道顺序`} className="order-input" type="number" min="0" placeholder="顺序" value={laneForm.sortOrder ?? ""} onChange={(e) => setLaneForm({ ...laneForm, sortOrder: e.target.value === "" ? undefined : Number(e.target.value) })} />
            <button className="button primary" disabled={busy}>{laneForm.id ? "保存泳道" : "添加泳道"}</button>{laneForm.id && <button type="button" className="button ghost" disabled={busy} onClick={() => setLaneForm({ name: "", roadmapId: map.id })}>取消编辑</button>}
          </form>
          <div className="roadmap-lane-chips">{lanes.map((lane) => <span key={lane.id}><b>{lane.name}</b><small>顺序 {lane.sortOrder}</small><button type="button" className="text-button" disabled={busy} onClick={() => setLaneForm({ id: lane.id, roadmapId: map.id, name: lane.name, sortOrder: lane.sortOrder })}>编辑</button><button type="button" className="text-button danger" disabled={busy} onClick={() => confirm({ title: "删除泳道", message: `删除“${map.name} · ${lane.name}”？`, confirmLabel: "删除", action: () => run(async () => { await onDeleteLane(lane.id); if (laneForm.id === lane.id) setLaneForm({ name: "", roadmapId: map.id }); }) })}>删除</button></span>)}{!lanes.length && <small>尚未配置泳道，请在上方添加。</small>}</div>
        </div>}
      </div>;
    })}{!roadmaps.length && <p className="roadmap-dictionary-empty">暂无路标，请先在上方新增。</p>}</div>
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
