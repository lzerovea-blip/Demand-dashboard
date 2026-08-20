import { useEffect, useMemo, useRef, useState, type FormEvent, type ReactNode } from "react";
import {
  buildHalfYearSummaries,
  buildRoadmapGroups,
  groupMatchesRoadmapLane,
  halfYearSequence,
  monthsForHalfYear,
  roadmapCardLabel,
  roadmapLanesForTrack,
  roadmapMonthColumnTemplate,
  roadmapTrackOrder,
  type HalfYearSummary,
  type RoadmapLane,
  type RoadmapGroup,
} from "../shared/roadmap";
import { latestRequirementSelections } from "../shared/entryDefaults";
import { filterRequirements, listTargetMonths } from "../shared/requirements";
import {
  CATEGORIES,
  MAX_REQUIREMENT_IMAGE_BYTES,
  MAX_REQUIREMENT_IMAGES,
  OVERSEAS_REGIONS,
  REQUIREMENT_IMAGE_MIME_TYPES,
  SOURCES,
  type AppSnapshot,
  type DictionaryItem,
  type Requirement,
  type RequirementCategory,
  type RequirementImage,
  type RequirementSource,
  type SaveRequirementInput,
  type Track,
  type WorkspaceImportMode,
  type WorkspaceImportPreview,
  type WorkspaceWorkbookConflictMode,
  type WorkspaceWorkbookImportPreview,
} from "../shared/types";
import {
  normalizeWorkloadEdit,
  parseWorkloadPart,
  sumWorkloadBreakdown,
  workloadBreakdownOf,
  type WorkloadBreakdown,
  type WorkloadSide,
} from "../shared/workload";

type Page = "requirements" | "summary" | "roadmap" | "dictionaries" | "template";

const EMPTY_SNAPSHOT: AppSnapshot = {
  requirements: [],
  domains: [],
  products: [],
  groupOverrides: [],
  templates: [],
};

const HALF_YEARS = Array.from({ length: 18 }, (_, index) => {
  const serial = 2024 * 2 + index;
  return `${Math.floor(serial / 2)}H${(serial % 2) + 1}`;
});

const SOURCE_COLORS: Record<RequirementSource, string> = {
  运动基础: "#2687ff",
  运动进阶: "#1cb7a5",
  运动高阶: "#7357d9",
  健康基础: "#ef6b54",
  健康进阶: "#f1a33c",
  健康高阶: "#d64d85",
  海外研究: "#14927f",
};

const WORKLOAD_SIDE_META: Record<WorkloadSide, { label: string; color: string }> = {
  device: { label: "设备侧", color: "#12a99a" },
  app: { label: "App 侧", color: "#287ef0" },
  cloud: { label: "云侧", color: "#7357d9" },
};

export default function App() {
  const [snapshot, setSnapshot] = useState<AppSnapshot>(EMPTY_SNAPSHOT);
  const [page, setPage] = useState<Page>("requirements");
  const [loading, setLoading] = useState(true);
  const [notice, setNotice] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [editor, setEditor] = useState<Requirement | "new" | null>(null);
  const [selectedGroupKey, setSelectedGroupKey] = useState<string | null>(null);
  const [startHalf, setStartHalf] = useState("2026H1");
  const [endHalf, setEndHalf] = useState("2027H2");

  useEffect(() => {
    window.roadmapApi
      .getSnapshot()
      .then(setSnapshot)
      .catch((reason) => setError(messageOf(reason)))
      .finally(() => setLoading(false));
  }, []);

  useEffect(() => {
    if (!notice && !error) return;
    const timer = window.setTimeout(() => {
      setNotice("");
      setError("");
    }, 4200);
    return () => window.clearTimeout(timer);
  }, [notice, error]);

  const groups = useMemo(() => {
    try {
      return buildRoadmapGroups(snapshot, startHalf, endHalf);
    } catch {
      return [];
    }
  }, [snapshot, startHalf, endHalf]);

  const selectedGroup = groups.find((item) => item.key === selectedGroupKey) ?? null;
  const summaries = useMemo(() => {
    try {
      return buildHalfYearSummaries(snapshot.requirements, startHalf, endHalf);
    } catch {
      return [];
    }
  }, [snapshot.requirements, startHalf, endHalf]);

  async function update(action: () => Promise<AppSnapshot>, success: string) {
    try {
      setError("");
      setSnapshot(await action());
      setNotice(success);
    } catch (reason) {
      setError(messageOf(reason));
      throw reason;
    }
  }

  if (loading) return <div className="boot">正在打开离线需求库…</div>;

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">R</div>
          <div>
            <strong>需求路标</strong>
            <span>OFFLINE STUDIO</span>
          </div>
        </div>
        <nav>
          <NavButton active={page === "requirements"} icon="⌘" label="需求池" onClick={() => setPage("requirements")} />
          <NavButton active={page === "summary"} icon="↗" label="汇总分析" onClick={() => setPage("summary")} />
          <NavButton active={page === "roadmap"} icon="▦" label="路标预览" onClick={() => setPage("roadmap")} />
          <NavButton active={page === "dictionaries"} icon="◇" label="字典设置" onClick={() => setPage("dictionaries")} />
          <NavButton active={page === "template"} icon="P" label="PPT 输出" onClick={() => setPage("template")} />
        </nav>
        <div className="sidebar-foot">
          <div className="offline-dot"><i /> 全程本地处理</div>
          <small>{snapshot.requirements.length} 条需求 · {groups.length} 张领域卡片</small>
        </div>
      </aside>

      <main className="main-panel">
        <header className="topbar">
          <div>
            <h1>{pageTitle(page)}</h1>
            <p>{pageSubtitle(page)}</p>
          </div>
          <div className="top-actions">
            {(page === "summary" || page === "roadmap" || page === "template") && (
              <PeriodSelector start={startHalf} end={endHalf} onStart={setStartHalf} onEnd={setEndHalf} />
            )}
            {page === "requirements" && (
              <button className="button primary" onClick={() => setEditor("new")}>＋ 新建需求</button>
            )}
          </div>
        </header>

        <section className="content">
          {page === "requirements" && (
            <RequirementsPage
              snapshot={snapshot}
              onEdit={setEditor}
              onDelete={async (id) => {
                if (!window.confirm("确定删除这条需求？此操作可通过备份恢复。")) return;
                await update(() => window.roadmapApi.deleteRequirement(id), "需求已删除");
              }}
            />
          )}
          {page === "summary" && <SummaryPage summaries={summaries} />}
          {page === "roadmap" && (
            <RoadmapPage groups={groups} start={startHalf} end={endHalf} onOpenGroup={(group) => setSelectedGroupKey(group.key)} />
          )}
          {page === "dictionaries" && (
            <DictionariesPage
              domains={snapshot.domains}
              products={snapshot.products}
              onSaveDomain={(input) => update(() => window.roadmapApi.saveDomain(input), "领域已保存")}
              onSaveProduct={(input) => update(() => window.roadmapApi.saveProduct(input), "产品已保存")}
              onDeleteDomain={(id) => update(() => window.roadmapApi.deleteDomain(id), "领域已删除")}
              onDeleteProduct={(id) => update(() => window.roadmapApi.deleteProduct(id), "产品已删除")}
            />
          )}
          {page === "template" && (
            <TemplatePage
              snapshot={snapshot}
              start={startHalf}
              end={endHalf}
              onExportPpt={async () => {
                try {
                  const result = await window.roadmapApi.exportRoadmapPresentation({ start: startHalf, end: endHalf });
                  if (!result.canceled) setNotice(`PPT 已导出（${result.slideCount} 页）：${result.path}`);
                } catch (reason) {
                  setError(messageOf(reason));
                }
              }}
              onExportDraft={async () => {
                try {
                  const result = await window.roadmapApi.exportTemplateDraft();
                  if (!result.canceled) setNotice(`共创版已导出：${result.path}`);
                } catch (reason) {
                  setError(messageOf(reason));
                }
              }}
              onExportWorkspace={async () => {
                try {
                  const result = await window.roadmapApi.exportWorkspacePackage();
                  if (!result.canceled) setNotice(`完整数据包已导出：${result.path}`);
                } catch (reason) {
                  setError(messageOf(reason));
                }
              }}
              onInspectWorkspace={async () => {
                try {
                  const result = await window.roadmapApi.inspectWorkspacePackage();
                  return result.preview;
                } catch (reason) {
                  setError(messageOf(reason));
                  return undefined;
                }
              }}
              onApplyWorkspace={async (token, mode) => {
                try {
                  const result = await window.roadmapApi.applyWorkspacePackage({ token, mode });
                  setSnapshot(result.snapshot);
                  setNotice(mode === "merge" ? "数据包已合并导入" : "数据包已整体替换导入");
                  return true;
                } catch (reason) {
                  setError(messageOf(reason));
                  return false;
                }
              }}
              onExportWorkbook={async () => {
                try {
                  const result = await window.roadmapApi.exportWorkspaceWorkbook();
                  if (!result.canceled) setNotice(`协作 Excel 已导出：${result.path}`);
                } catch (reason) {
                  setError(messageOf(reason));
                }
              }}
              onInspectWorkbook={async () => {
                try {
                  const result = await window.roadmapApi.inspectWorkspaceWorkbook();
                  return result.preview;
                } catch (reason) {
                  setError(messageOf(reason));
                  return undefined;
                }
              }}
              onApplyWorkbook={async (token, conflictMode) => {
                try {
                  const result = await window.roadmapApi.applyWorkspaceWorkbook({ token, conflictMode });
                  setSnapshot(result.snapshot);
                  setNotice(conflictMode === "local-wins" ? "协作 Excel 已合并，冲突保留本机版本" : "协作 Excel 已合并，冲突采用 Excel 版本");
                  return true;
                } catch (reason) {
                  setError(messageOf(reason));
                  return false;
                }
              }}
            />
          )}
        </section>
      </main>

      {editor && (
        <RequirementEditor
          value={editor === "new" ? undefined : editor}
          snapshot={snapshot}
          onClose={() => setEditor(null)}
          onSave={async (input, keepOpen) => {
            await update(() => window.roadmapApi.saveRequirement(input), input.id ? "需求已更新" : "需求已加入路标");
            setEditor(keepOpen ? "new" : null);
          }}
        />
      )}
      {selectedGroup && (
        <GroupDrawer
          group={selectedGroup}
          onClose={() => setSelectedGroupKey(null)}
          onSave={async (cardTitle, cardSummary) => {
            await update(
              () => window.roadmapApi.saveGroupOverride({ groupKey: selectedGroup.key, cardTitle, cardSummary }),
              "领域卡片已更新",
            );
          }}
        />
      )}
      {(notice || error) && <div className={`toast ${error ? "error" : "success"}`}>{error || notice}</div>}
    </div>
  );
}

function NavButton({ active, icon, label, onClick }: { active: boolean; icon: string; label: string; onClick: () => void }) {
  return <button className={active ? "nav-button active" : "nav-button"} onClick={onClick}><b>{icon}</b><span>{label}</span></button>;
}

function PeriodSelector({ start, end, onStart, onEnd }: { start: string; end: string; onStart: (value: string) => void; onEnd: (value: string) => void }) {
  return (
    <div className="period-selector">
      <select value={start} onChange={(event) => onStart(event.target.value)}>{HALF_YEARS.map((item) => <option key={item}>{item}</option>)}</select>
      <span>至</span>
      <select value={end} onChange={(event) => onEnd(event.target.value)}>{HALF_YEARS.map((item) => <option key={item}>{item}</option>)}</select>
    </div>
  );
}

function RequirementsPage({ snapshot, onEdit, onDelete }: { snapshot: AppSnapshot; onEdit: (value: Requirement) => void; onDelete: (id: string) => void }) {
  const [query, setQuery] = useState("");
  const [source, setSource] = useState<RequirementSource | "">("");
  const [category, setCategory] = useState<RequirementCategory | "">("");
  const [targetMonth, setTargetMonth] = useState("");
  const domains = new Map(snapshot.domains.map((item) => [item.id, item.name]));
  const products = new Map(snapshot.products.map((item) => [item.id, item.name]));
  const targetMonths = listTargetMonths(snapshot.requirements);
  useEffect(() => {
    if (targetMonth && !targetMonths.includes(targetMonth)) setTargetMonth("");
  }, [targetMonth, targetMonths.join("|")]);
  const rows = filterRequirements(snapshot.requirements, domains, { query, source, category, targetMonth });

  return (
    <>
      <div className="panel table-panel">
        <div className="toolbar">
          <label className="search"><span>⌕</span><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="搜索标题或领域" /></label>
          <select value={source} onChange={(event) => setSource(event.target.value as RequirementSource | "")}><option value="">全部来源</option>{SOURCES.map((item) => <option key={item}>{item}</option>)}</select>
          <select value={category} onChange={(event) => setCategory(event.target.value as RequirementCategory | "")}><option value="">全部分类</option>{CATEGORIES.map((item) => <option key={item}>{item}</option>)}</select>
          <select value={targetMonth} onChange={(event) => setTargetMonth(event.target.value)}><option value="">全部月份</option>{targetMonths.map((item) => <option key={item} value={item}>{formatMonth(item)}</option>)}</select>
          <span className="toolbar-count">显示 {rows.length} / {snapshot.requirements.length}</span>
        </div>
        {rows.length ? (
          <div className="table-wrap">
            <table>
              <thead><tr><th>需求标题</th><th>领域</th><th>来源</th><th>海外区域</th><th>分类</th><th>上线月份</th><th>匹配产品</th><th className="numeric">工作量</th><th /></tr></thead>
              <tbody>
                {rows.map((item) => (
                  <tr key={item.id} onDoubleClick={() => onEdit(item)}>
                    <td><strong>{item.title}</strong></td>
                    <td>{domains.get(item.domainId) ?? "未命名"}</td>
                    <td><SourceBadge source={item.source} /></td>
                    <td>{item.overseasRegions.join("、") || "—"}</td>
                    <td><span className={`category ${item.category === "产品专属" ? "exclusive" : "experience"}`}>{item.category}</span></td>
                    <td>{formatMonth(item.targetMonth)}</td>
                    <td>{item.productIds.map((id) => products.get(id)).filter(Boolean).join("、") || "—"}</td>
                    <td className="numeric workload-cell"><b>{formatNumber(item.workloadPm)} PM</b><span>{formatWorkloadBreakdown(workloadBreakdownOf(item))}</span></td>
                    <td className="row-actions"><button onClick={() => onEdit(item)}>编辑</button><button className="danger-text" onClick={() => onDelete(item.id)}>删除</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : <Empty title="还没有符合条件的需求" body={snapshot.requirements.length ? "尝试清空搜索或调整来源、分类、上线月份。" : "先维护领域与产品字典，再新建第一条需求。"} />}
      </div>
    </>
  );
}

function SummaryPage({ summaries }: { summaries: HalfYearSummary[] }) {
  const [selectedSide, setSelectedSide] = useState<WorkloadSide>("device");
  const device = summaries.reduce((sum, item) => sum + item.deviceWorkload, 0);
  const app = summaries.reduce((sum, item) => sum + item.appWorkload, 0);
  const cloud = summaries.reduce((sum, item) => sum + item.cloudWorkload, 0);
  const total = device + app + cloud;
  const sports = summaries.reduce((sum, item) => sum + item.sportsWorkload, 0);
  const health = summaries.reduce((sum, item) => sum + item.healthWorkload, 0);
  const overseas = summaries.reduce((sum, item) => sum + item.overseasWorkload, 0);
  const experience = summaries.reduce((sum, item) => sum + item.experienceWorkload, 0);
  const exclusive = summaries.reduce((sum, item) => sum + item.exclusiveWorkload, 0);
  const unallocated = summaries.reduce((sum, item) => sum + item.unallocatedWorkload, 0);
  return (
    <>
      <div className="metric-row summary-kpis">
        <Metric label="总工作量（三端合计）" value={formatNumber(total)} suffix="人月" accent="ink" />
        <Metric label="设备侧总工作量" value={formatNumber(device)} suffix="人月" accent="cyan" />
        <Metric label="App 侧总工作量" value={formatNumber(app)} suffix="人月" accent="blue" />
        <Metric label="云侧总工作量" value={formatNumber(cloud)} suffix="人月" accent="violet" />
      </div>
      {unallocated > 0 && <div className="workload-warning">选定区间仍有 {formatNumber(unallocated)} 人月历史工作量待拆分，当前不会计入三侧明细。</div>}
      <div className="summary-pies">
        <DonutChart title="三端投入占比" subtitle="设备 / App / 云侧" items={[
          { label: "设备侧", value: device, color: WORKLOAD_SIDE_META.device.color },
          { label: "App 侧", value: app, color: WORKLOAD_SIDE_META.app.color },
          { label: "云侧", value: cloud, color: WORKLOAD_SIDE_META.cloud.color },
        ]} />
        <DonutChart title="运动 / 健康 / 海外研究占比" subtitle="按需求来源归属" items={[
          { label: "运动", value: sports, color: "#287ef0" },
          { label: "健康", value: health, color: "#ef6b54" },
          { label: "海外研究", value: overseas, color: "#14927f" },
        ]} />
        <DonutChart title="体验优化 / 产品专属" subtitle="按需求分类归属" items={[
          { label: "体验优化", value: experience, color: "#7357d9" },
          { label: "产品专属", value: exclusive, color: "#17202a" },
        ]} />
      </div>
      <div className="summary-charts">
        <div className="panel summary-chart-panel">
          <PanelTitle title="三端投入工作量" subtitle="固定展示设备侧、App 侧、云侧 · 单位：人月" />
          <ThreeSideBarChart summaries={summaries} />
          <div className="chart-legend">{(["device", "app", "cloud"] as WorkloadSide[]).map((side) => <span key={side}><i style={{ background: WORKLOAD_SIDE_META[side].color }} />{WORKLOAD_SIDE_META[side].label}</span>)}</div>
        </div>
        <div className="panel summary-chart-panel">
          <PanelTitle
            title="所选端投入与体验优化占比"
            subtitle="曲线：工作量（左轴） · 柱：体验优化占比（右轴）"
            action={<div className="side-tabs">{(["device", "app", "cloud"] as WorkloadSide[]).map((side) => <button key={side} className={selectedSide === side ? "active" : ""} onClick={() => setSelectedSide(side)}>{WORKLOAD_SIDE_META[side].label}</button>)}</div>}
          />
          <SideExperienceChart summaries={summaries} side={selectedSide} />
          <div className="chart-legend"><span><i style={{ background: WORKLOAD_SIDE_META[selectedSide].color }} />{WORKLOAD_SIDE_META[selectedSide].label}投入曲线</span><span><i className="experience-legend" />体验优化占比柱</span></div>
        </div>
      </div>
    </>
  );
}

interface DonutItem {
  label: string;
  value: number;
  color: string;
}

function DonutChart({ title, subtitle, items }: { title: string; subtitle: string; items: DonutItem[] }) {
  const total = items.reduce((sum, item) => sum + item.value, 0);
  let offset = 0;
  return (
    <div className="panel donut-panel">
      <PanelTitle title={title} subtitle={subtitle} />
      <div className="donut-content">
        <svg className="donut-chart" viewBox="0 0 42 42" role="img" aria-label={`${title}，合计${formatNumber(total)}人月`}>
          <circle cx="21" cy="21" r="15.9155" fill="none" stroke="#edf1f3" strokeWidth="7" />
          {items.map((item) => {
            const segment = total ? (item.value / total) * 100 : 0;
            const start = offset;
            offset += segment;
            return <circle key={item.label} cx="21" cy="21" r="15.9155" fill="none" stroke={item.color} strokeWidth="7" pathLength="100" strokeDasharray={`${segment} ${100 - segment}`} strokeDashoffset={-start} transform="rotate(-90 21 21)"><title>{item.label}：{formatNumber(item.value)} 人月，占 {total ? Math.round(segment) : 0}%</title></circle>;
          })}
          <text x="21" y="20" textAnchor="middle" className="donut-total">{formatNumber(total)}</text>
          <text x="21" y="25" textAnchor="middle" className="donut-unit">人月</text>
        </svg>
        <div className="donut-legend">{items.map((item) => <div key={item.label}><i style={{ background: item.color }} /><span>{item.label}</span><b>{formatNumber(item.value)} 人月</b><em>{total ? Math.round((item.value / total) * 100) : 0}%</em></div>)}</div>
      </div>
    </div>
  );
}

function ThreeSideBarChart({ summaries }: { summaries: HalfYearSummary[] }) {
  const width = 860;
  const height = 330;
  const left = 55;
  const top = 24;
  const right = 20;
  const bottom = 46;
  const chartWidth = width - left - right;
  const chartHeight = height - top - bottom;
  const max = Math.max(1, ...summaries.flatMap((item) => Object.values(item.bySide)));
  const groupWidth = chartWidth / Math.max(1, summaries.length);
  const gap = 3;
  const barWidth = Math.min(28, Math.max(5, (groupWidth - 18) / 3));
  const groupBarsWidth = barWidth * 3 + gap * 2;
  const y = (value: number) => top + chartHeight - (value / max) * chartHeight;
  const ticks = [0, 0.25, 0.5, 0.75, 1];
  return (
    <svg className="summary-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="设备侧、App侧、云侧投入工作量柱状图">
      {ticks.map((ratio) => { const tickY = top + chartHeight * (1 - ratio); return <g key={ratio}><line x1={left} x2={width - right} y1={tickY} y2={tickY} stroke="#e8edef" /><text x={left - 10} y={tickY + 4} textAnchor="end">{formatNumber(max * ratio)}</text></g>; })}
      {summaries.map((item, index) => {
        const center = left + groupWidth * (index + 0.5);
        const start = center - groupBarsWidth / 2;
        return <g key={item.halfYear}>{(["device", "app", "cloud"] as WorkloadSide[]).map((side, sideIndex) => { const value = item.bySide[side]; return <rect key={side} x={start + sideIndex * (barWidth + gap)} y={y(value)} width={barWidth} height={Math.max(0, top + chartHeight - y(value))} rx="3" fill={WORKLOAD_SIDE_META[side].color}><title>{shortHalf(item.halfYear)} · {WORKLOAD_SIDE_META[side].label}：{formatNumber(value)} 人月</title></rect>; })}<text x={center} y={height - 17} textAnchor="middle">{shortHalf(item.halfYear)}</text></g>;
      })}
    </svg>
  );
}

function SideExperienceChart({ summaries, side }: { summaries: HalfYearSummary[]; side: WorkloadSide }) {
  const width = 760;
  const height = 330;
  const left = 55;
  const top = 24;
  const right = 48;
  const bottom = 46;
  const chartWidth = width - left - right;
  const chartHeight = height - top - bottom;
  const maxWorkload = Math.max(1, ...summaries.map((item) => item.bySide[side]));
  const groupWidth = chartWidth / Math.max(1, summaries.length);
  const x = (index: number) => left + groupWidth * (index + 0.5);
  const workloadY = (value: number) => top + chartHeight - (value / maxWorkload) * chartHeight;
  const percentOf = (item: HalfYearSummary) => item.bySide[side] ? (item.experienceBySide[side] / item.bySide[side]) * 100 : 0;
  const percentY = (value: number) => top + chartHeight - (value / 100) * chartHeight;
  const barWidth = Math.min(45, groupWidth * 0.42);
  const points = summaries.map((item, index) => `${x(index)},${workloadY(item.bySide[side])}`).join(" ");
  return (
    <svg className="summary-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label={`${WORKLOAD_SIDE_META[side].label}投入曲线和体验优化占比柱状图`}>
      {[0, 0.25, 0.5, 0.75, 1].map((ratio) => { const tickY = top + chartHeight * (1 - ratio); return <g key={ratio}><line x1={left} x2={width - right} y1={tickY} y2={tickY} stroke="#e8edef" /><text x={left - 10} y={tickY + 4} textAnchor="end">{formatNumber(maxWorkload * ratio)}</text><text x={width - right + 10} y={tickY + 4}>{Math.round(ratio * 100)}%</text></g>; })}
      {summaries.map((item, index) => { const percentage = percentOf(item); return <g key={item.halfYear}><rect x={x(index) - barWidth / 2} y={percentY(percentage)} width={barWidth} height={Math.max(0, top + chartHeight - percentY(percentage))} rx="4" fill="#dcd5ff"><title>{shortHalf(item.halfYear)} · 体验优化占比：{Math.round(percentage)}%</title></rect><text x={x(index)} y={height - 17} textAnchor="middle">{shortHalf(item.halfYear)}</text></g>; })}
      <polyline points={points} fill="none" stroke={WORKLOAD_SIDE_META[side].color} strokeWidth="3" strokeLinejoin="round" strokeLinecap="round" />
      {summaries.map((item, index) => <circle key={item.halfYear} cx={x(index)} cy={workloadY(item.bySide[side])} r="4" fill="#fff" stroke={WORKLOAD_SIDE_META[side].color} strokeWidth="3"><title>{shortHalf(item.halfYear)} · {WORKLOAD_SIDE_META[side].label}：{formatNumber(item.bySide[side])} 人月</title></circle>)}
    </svg>
  );
}

function RoadmapPage({ groups, start, end, onOpenGroup }: { groups: RoadmapGroup[]; start: string; end: string; onOpenGroup: (group: RoadmapGroup) => void }) {
  const tracks = roadmapTrackOrder(groups);
  return (
    <div className="roadmap-stack">
      {tracks.map((track) => (
        <RoadmapBoard key={track} title={`${track}路标`} track={track} groups={groups} start={start} end={end} onOpenGroup={onOpenGroup} />
      ))}
    </div>
  );
}

function RoadmapBoard({ title, track, groups, start, end, onOpenGroup }: { title: string; track: Track; groups: RoadmapGroup[]; start: string; end: string; onOpenGroup: (group: RoadmapGroup) => void }) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const halves = halfYearSequence(start, end);
  const months = halves.flatMap(monthsForHalfYear);
  const scoped = groups.filter((item) => item.track === track);
  const firstGroup = scoped[0];
  const lanes = roadmapLanesForTrack(track);
  const firstLane = firstGroup ? (firstGroup.track === "海外研究" ? firstGroup.overseasRegions[0] : firstGroup.level) : undefined;
  const monthColumnTemplate = roadmapMonthColumnTemplate(months, groups);

  function locateFirstGroup(behavior: ScrollBehavior = "auto") {
    if (!firstGroup || !firstLane || !scrollRef.current) return;
    const target = scrollRef.current.querySelector<HTMLElement>(
      `[data-roadmap-month="${firstGroup.targetMonth}"][data-roadmap-lane="${firstLane}"]`,
    );
    if (!target) return;
    const centeredOffset = Math.max(116, (scrollRef.current.clientWidth - target.clientWidth) / 2);
    scrollRef.current.scrollTo({ left: Math.max(0, target.offsetLeft - centeredOffset), behavior });
  }

  useEffect(() => {
    const frame = window.requestAnimationFrame(() => locateFirstGroup());
    return () => window.cancelAnimationFrame(frame);
  }, [firstGroup?.key, firstLane, start, end]);

  return (
    <div className={`panel roadmap-board ${track === "运动" ? "sports" : track === "健康" ? "health" : "overseas"}`}>
      <PanelTitle
        title={title}
        subtitle={`${shortHalf(start)} – ${shortHalf(end)} · ${scoped.length} 张卡片 · 点击卡片查看领域全量需求`}
        action={firstGroup
          ? <button type="button" className="roadmap-locate" onClick={() => locateFirstGroup("smooth")}>定位 {formatMonth(firstGroup.targetMonth)}</button>
          : <span className="roadmap-empty-count">暂无卡片</span>}
      />
      <div className="roadmap-scroll" ref={scrollRef}>
        <div className="roadmap-grid" style={{ gridTemplateColumns: `116px ${monthColumnTemplate}` }}>
          <div className="lane-corner">来源泳道</div>
          {halves.map((half) => <div key={half} className="half-header" style={{ gridColumn: "span 6" }}>{shortHalf(half)}</div>)}
          <div className="lane-corner muted">上线月份</div>
          {months.map((month) => <div key={month} className="month-header">{Number(month.slice(5))}月</div>)}
          {lanes.map((lane) => (
            <RoadmapLane key={lane} lane={lane} months={months} groups={scoped.filter((item) => groupMatchesRoadmapLane(item, lane))} onOpenGroup={onOpenGroup} />
          ))}
        </div>
      </div>
    </div>
  );
}

function RoadmapLane({ lane, months, groups, onOpenGroup }: { lane: RoadmapLane; months: string[]; groups: RoadmapGroup[]; onOpenGroup: (group: RoadmapGroup) => void }) {
  return (
    <>
      <div className="lane-label"><strong>{lane}</strong><span>{groups.length} 卡片</span></div>
      {months.map((month) => (
        <div key={`${lane}-${month}`} className="month-cell" data-roadmap-month={month} data-roadmap-lane={lane}>
          {groups.filter((item) => item.targetMonth === month).map((group) => <RoadmapCard key={group.key} group={group} onClick={() => onOpenGroup(group)} />)}
        </div>
      ))}
    </>
  );
}

function RoadmapCard({ group, onClick }: { group: RoadmapGroup; onClick: () => void }) {
  const hasExclusive = group.categories.includes("产品专属");
  const label = roadmapCardLabel(group);
  const requirementText = `${group.requirements.slice(0, 3).map((item) => item.title).join("、")}${group.requirements.length > 3 ? "等" : ""}`;
  const visibleProducts = group.productNames.slice(0, 2);
  const hiddenProductCount = Math.max(group.productNames.length - visibleProducts.length, 0);
  const productLabel = group.productNames.length ? `涉及产品：${group.productNames.join("、")}` : "";
  const regionLabel = group.overseasRegions.length ? `海外区域：${group.overseasRegions.join("、")}` : "";
  const accessibleLabel = [label, regionLabel, productLabel, "点击查看全部需求"].filter(Boolean).join("；");
  return (
    <button className={`roadmap-card ${hasExclusive ? "has-exclusive" : ""}`} onClick={onClick} title={accessibleLabel} aria-label={accessibleLabel}>
      <span className="roadmap-card-copy"><strong>{group.cardTitle}：</strong><span>{requirementText}</span></span>
      {visibleProducts.length > 0 && (
        <span className="roadmap-card-products" aria-label={productLabel}>
          <i>涉及产品</i>
          {visibleProducts.map((name) => <b key={name}>{name}</b>)}
          {hiddenProductCount > 0 && <b title={group.productNames.slice(visibleProducts.length).join("、")}>+{hiddenProductCount}</b>}
        </span>
      )}
    </button>
  );
}

type WorkloadTexts = { device: string; app: string; cloud: string };

function RequirementEditor({ value, snapshot, onClose, onSave }: { value?: Requirement; snapshot: AppSnapshot; onClose: () => void; onSave: (input: SaveRequirementInput, keepOpen: boolean) => Promise<void> }) {
  const [form, setForm] = useState<SaveRequirementInput>(() => {
    if (value) return {
      id: value.id,
      title: value.title,
      description: value.description ?? "",
      images: value.images ?? [],
      domainId: value.domainId,
      source: value.source,
      overseasRegions: value.overseasRegions ?? [],
      category: value.category,
      targetMonth: value.targetMonth,
      productIds: value.productIds,
      deviceWorkloadPm: value.deviceWorkloadPm,
      appWorkloadPm: value.appWorkloadPm,
      cloudWorkloadPm: value.cloudWorkloadPm,
    };
    const previous = latestRequirementSelections(snapshot.requirements);
    return {
      title: "",
      description: "",
      images: [],
      domainId: snapshot.domains.find((item) => item.active)?.id ?? "",
      source: previous?.source ?? "运动基础",
      overseasRegions: previous?.overseasRegions ?? [],
      category: previous?.category ?? "体验优化",
      targetMonth: previous?.targetMonth ?? new Date().toISOString().slice(0, 7),
      productIds: [],
      deviceWorkloadPm: 0,
      appWorkloadPm: 0,
      cloudWorkloadPm: 0,
    };
  });
  const [workloadTexts, setWorkloadTexts] = useState<WorkloadTexts>(() => ({
    device: value?.deviceWorkloadPm ? String(value.deviceWorkloadPm) : "",
    app: value?.appWorkloadPm ? String(value.appWorkloadPm) : "",
    cloud: value?.cloudWorkloadPm ? String(value.cloudWorkloadPm) : "",
  }));
  const [workloadError, setWorkloadError] = useState("");
  const [regionError, setRegionError] = useState("");
  const [imageError, setImageError] = useState("");
  const imageInputRef = useRef<HTMLInputElement>(null);
  const [saving, setSaving] = useState(false);
  const [keepOpen, setKeepOpen] = useState(false);
  const draftParts = {
    device: parseWorkloadPart(workloadTexts.device),
    app: parseWorkloadPart(workloadTexts.app),
    cloud: parseWorkloadPart(workloadTexts.cloud),
  };
  const draftTotal = Object.values(draftParts).every((item) => item !== null)
    ? (draftParts.device ?? 0) + (draftParts.app ?? 0) + (draftParts.cloud ?? 0)
    : null;

  function updateWorkloadText(key: keyof WorkloadTexts, raw: string) {
    const normalized = normalizeWorkloadEdit(raw);
    if (normalized === null) return;
    setWorkloadTexts((current) => ({ ...current, [key]: normalized }));
    setWorkloadError("");
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    if (form.source === "海外研究" && form.overseasRegions.length === 0) {
      setRegionError("请至少选择一个海外研究区域");
      return;
    }
    const parsed = [draftParts.device, draftParts.app, draftParts.cloud];
    if (parsed.some((item) => item === null)) {
      setWorkloadError("各侧工作量请输入 0、正整数或小数，例如 0.5");
      return;
    }
    if ((draftTotal ?? 0) <= 0) {
      setWorkloadError("设备、App、云侧工作量至少填写一项");
      return;
    }
    setSaving(true);
    try {
      await onSave({
        ...form,
        deviceWorkloadPm: draftParts.device ?? 0,
        appWorkloadPm: draftParts.app ?? 0,
        cloudWorkloadPm: draftParts.cloud ?? 0,
      }, keepOpen && !value);
      if (keepOpen && !value) {
        setForm((current) => ({ ...current, title: "", description: "", images: [], productIds: [], deviceWorkloadPm: 0, appWorkloadPm: 0, cloudWorkloadPm: 0 }));
        setWorkloadTexts({ device: "", app: "", cloud: "" });
        setImageError("");
      }
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="overlay" onMouseDown={(event) => event.currentTarget === event.target && onClose()}>
      <form className="drawer editor-drawer" onSubmit={submit}>
        <div className="drawer-header"><div><span className="eyebrow">STRUCTURED INPUT</span><h2>{value ? "编辑需求" : "新建需求"}</h2></div><button type="button" className="icon-button" onClick={onClose}>×</button></div>
        <div className="form-body">
          {snapshot.domains.length === 0 && <div className="callout warning">请先到“字典设置”新增至少一个领域。</div>}
          <Field label="需求标题" required><input autoFocus value={form.title} onChange={(event) => setForm({ ...form, title: event.target.value })} placeholder="例如：支持早间血压测量" /></Field>
          <Field label="需求描述">
            <textarea rows={5} maxLength={5000} value={form.description} onChange={(event) => setForm({ ...form, description: event.target.value })} placeholder="补充使用场景、用户问题、方案范围或验收要求" />
            <small>{form.description.length} / 5000 字</small>
          </Field>
          <Field label={`需求图片（${form.images.length}/${MAX_REQUIREMENT_IMAGES}）`}>
            <input
              ref={imageInputRef}
              className="visually-hidden"
              type="file"
              multiple
              accept={REQUIREMENT_IMAGE_MIME_TYPES.join(",")}
              onChange={async (event) => {
                try {
                  const result = await readRequirementImages(event.target.files, form.images.length);
                  if (result.images.length) {
                    setForm((current) => ({ ...current, images: [...current.images, ...result.images].slice(0, MAX_REQUIREMENT_IMAGES) }));
                  }
                  setImageError(result.error ?? "");
                } catch (reason) {
                  setImageError(messageOf(reason));
                } finally {
                  event.target.value = "";
                }
              }}
            />
            <div className="image-picker">
              <button type="button" className="button ghost image-pick-button" disabled={form.images.length >= MAX_REQUIREMENT_IMAGES} onClick={() => imageInputRef.current?.click()}>＋ 选择本地图片</button>
              <span>PNG / JPG / WebP / GIF，单张不超过 5MB；图片仅保存在本机。</span>
            </div>
            {form.images.length > 0 && <div className="image-preview-grid">{form.images.map((image) => (
              <div className="image-preview" key={image.id}>
                <img src={image.dataUrl} alt={image.name} />
                <button type="button" aria-label={`移除${image.name}`} onClick={() => setForm({ ...form, images: form.images.filter((item) => item.id !== image.id) })}>×</button>
                <span title={image.name}>{image.name}</span>
              </div>
            ))}</div>}
            {imageError && <small className="input-error">{imageError}</small>}
          </Field>
          <div className="form-grid">
            <Field label="领域" required><select value={form.domainId} onChange={(event) => setForm({ ...form, domainId: event.target.value })}><option value="">请选择</option>{snapshot.domains.filter((item) => item.active).map((item) => <option key={item.id} value={item.id}>{item.name}</option>)}</select></Field>
            <Field label="来源" required><select value={form.source} onChange={(event) => {
              const source = event.target.value as RequirementSource;
              setForm({ ...form, source, overseasRegions: source === "海外研究" ? form.overseasRegions : [] });
              setRegionError("");
            }}>{SOURCES.map((item) => <option key={item}>{item}</option>)}</select></Field>
            <Field label="分类" required><div className="segmented">{CATEGORIES.map((item) => <button type="button" key={item} className={form.category === item ? "active" : ""} onClick={() => setForm({ ...form, category: item, productIds: item === "体验优化" ? form.productIds : form.productIds })}>{item}</button>)}</div></Field>
            <Field label="上线年月" required><input type="month" value={form.targetMonth} onChange={(event) => setForm({ ...form, targetMonth: event.target.value })} /></Field>
          </div>
          {form.source === "海外研究" && (
            <Field label="海外研究区域" required>
              <div className="check-grid overseas-region-grid">{OVERSEAS_REGIONS.map((region) => <label key={region}><input type="checkbox" checked={form.overseasRegions.includes(region)} onChange={(event) => {
                setForm({ ...form, overseasRegions: event.target.checked ? [...form.overseasRegions, region] : form.overseasRegions.filter((item) => item !== region) });
                setRegionError("");
              }} /><span>{region}</span></label>)}</div>
              {regionError && <small className="input-error">{regionError}</small>}
            </Field>
          )}
          <Field label={`匹配产品${form.category === "产品专属" ? "（必选）" : "（可选）"}`} required={form.category === "产品专属"}>
            {snapshot.products.length ? <div className="check-grid">{snapshot.products.filter((item) => item.active).map((item) => <label key={item.id}><input type="checkbox" checked={form.productIds.includes(item.id)} onChange={(event) => setForm({ ...form, productIds: event.target.checked ? [...form.productIds, item.id] : form.productIds.filter((id) => id !== item.id) })} /><span>{item.name}</span></label>)}</div> : <div className="inline-empty">暂无产品，可先保存体验优化需求，或前往字典设置新增产品。</div>}
          </Field>
          <div className="workload-section">
            <div className="workload-section-head"><span>工作量（人月） <b>*</b></span><strong>合计 {draftTotal === null ? "—" : formatNumber(draftTotal)} 人月</strong></div>
            {Boolean(value?.unallocatedWorkloadPm) && <div className="legacy-workload-warning">这条历史需求有 {formatNumber(value?.unallocatedWorkloadPm ?? 0)} 人月尚未拆分，请分配到下方三项后保存。</div>}
            <div className="workload-input-grid">
              <label><span>设备工作量</span><input type="text" inputMode="decimal" value={workloadTexts.device} className={workloadError ? "input-invalid" : undefined} aria-invalid={Boolean(workloadError)} placeholder="0" onChange={(event) => updateWorkloadText("device", event.target.value)} /></label>
              <label><span>App 工作量</span><input type="text" inputMode="decimal" value={workloadTexts.app} className={workloadError ? "input-invalid" : undefined} aria-invalid={Boolean(workloadError)} placeholder="0" onChange={(event) => updateWorkloadText("app", event.target.value)} /></label>
              <label><span>云侧工作量</span><input type="text" inputMode="decimal" value={workloadTexts.cloud} className={workloadError ? "input-invalid" : undefined} aria-invalid={Boolean(workloadError)} placeholder="0" onChange={(event) => updateWorkloadText("cloud", event.target.value)} /></label>
            </div>
            {workloadError
              ? <small className="input-error">{workloadError}</small>
              : <small>每项可填整数、小数或 0，系统自动合计；工作量不展示在路标轴卡片上。</small>}
          </div>
          <div className="merge-hint"><b>自动合并规则</b><span>同领域、同来源、同上线年月的需求会合并为一张路标卡片；分类可以不同。</span></div>
        </div>
        <div className="drawer-footer">
          {!value && <label className="keep-open"><input type="checkbox" checked={keepOpen} onChange={(event) => setKeepOpen(event.target.checked)} /> 保存后继续录入</label>}
          <div><button type="button" className="button ghost" onClick={onClose}>取消</button><button className="button primary" disabled={saving || snapshot.domains.length === 0}>{saving ? "保存中…" : "保存需求"}</button></div>
        </div>
      </form>
    </div>
  );
}

function GroupDrawer({ group, onClose, onSave }: { group: RoadmapGroup; onClose: () => void; onSave: (title: string, summary: string) => Promise<void> }) {
  const [title, setTitle] = useState(group.cardTitle);
  const [summary, setSummary] = useState(group.cardSummary);
  return (
    <div className="overlay" onMouseDown={(event) => event.currentTarget === event.target && onClose()}>
      <div className="drawer group-drawer">
        <div className="drawer-header"><div><span className="eyebrow">ROADMAP GROUP</span><h2>{group.domainName}</h2></div><button className="icon-button" onClick={onClose}>×</button></div>
        <div className="group-hero"><SourceBadge source={group.source} /><b>{formatMonth(group.targetMonth)}</b><span>合计 {formatNumber(group.totalWorkloadPm)} 人月</span><span>{formatWorkloadBreakdown(group.workloadBreakdown)}</span><span>{group.requirements.length} 条需求</span></div>
        <div className="form-body">
          <Field label="路标卡片标题" required><input value={title} onChange={(event) => setTitle(event.target.value)} /></Field>
          <Field label="详情 / PPT 摘要"><textarea rows={3} value={summary} onChange={(event) => setSummary(event.target.value)} /><small>该摘要保留给详情页与未来 PPT，不在路标轴卡片上显示。</small></Field>
          {group.overseasRegions.length > 0 && <div className="detail-regions"><span>海外区域</span>{group.overseasRegions.map((region) => <b key={region}>{region}</b>)}</div>}
          {group.productNames.length > 0 && <div className="detail-products"><span>涉及产品</span>{group.productNames.map((name) => <b key={name}>{name}</b>)}</div>}
          <div className="detail-list">
            <div className="detail-list-head"><h3>领域全量需求</h3><span>工作量合计 {formatNumber(group.totalWorkloadPm)} PM</span></div>
            {group.requirements.map((item, index) => (
              <details className="requirement-detail" key={item.id}>
                <summary>
                  <em>{String(index + 1).padStart(2, "0")}</em>
                  <div><strong>{item.title}</strong><span>{item.category}{item.overseasRegions.length ? ` · ${item.overseasRegions.join(" / ")}` : ""}{item.productIds.length ? ` · ${group.requirementProductNames[item.id].join(" / ")}` : ""}</span></div>
                  <b>{formatNumber(item.workloadPm)} PM</b>
                  <i>查看详情</i>
                </summary>
                <div className="requirement-detail-body">
                  <div className="detail-workload-breakdown">{workloadBreakdownItems(workloadBreakdownOf(item)).map(([label, value]) => <span key={label}><b>{label}</b>{formatNumber(value)} 人月</span>)}</div>
                  {item.description ? <p>{item.description}</p> : <p className="muted-detail">暂无需求描述</p>}
                  {item.images.length > 0 && <div className="detail-image-grid">{item.images.map((image) => <figure key={image.id}><img src={image.dataUrl} alt={image.name} /><figcaption>{image.name}</figcaption></figure>)}</div>}
                  {!item.description && item.images.length === 0 && <small>编辑该需求后可补充描述和图片。</small>}
                </div>
              </details>
            ))}
          </div>
        </div>
        <div className="drawer-footer"><span>导出PPT后，路标卡片将链接到对应详情页。</span><div><button className="button ghost" onClick={onClose}>关闭</button><button className="button primary" onClick={() => onSave(title, summary)}>保存卡片</button></div></div>
      </div>
    </div>
  );
}

function DictionariesPage({ domains, products, onSaveDomain, onSaveProduct, onDeleteDomain, onDeleteProduct }: { domains: DictionaryItem[]; products: DictionaryItem[]; onSaveDomain: (input: { name: string }) => Promise<void>; onSaveProduct: (input: { name: string }) => Promise<void>; onDeleteDomain: (id: string) => Promise<void>; onDeleteProduct: (id: string) => Promise<void> }) {
  return (
    <div className="dictionary-grid">
      <DictionaryPanel title="领域字典" subtitle="用于合并相同领域需求，例如血压、睡眠、跑步" items={domains} onSave={onSaveDomain} onDelete={onDeleteDomain} />
      <DictionaryPanel title="产品字典" subtitle="产品专属需求必须选择至少一个产品" items={products} onSave={onSaveProduct} onDelete={onDeleteProduct} />
    </div>
  );
}

function DictionaryPanel({ title, subtitle, items, onSave, onDelete }: { title: string; subtitle: string; items: DictionaryItem[]; onSave: (input: { name: string }) => Promise<void>; onDelete: (id: string) => Promise<void> }) {
  const [name, setName] = useState("");
  async function submit(event: FormEvent) {
    event.preventDefault();
    if (!name.trim()) return;
    await onSave({ name });
    setName("");
  }
  return (
    <div className="panel dictionary-panel">
      <PanelTitle title={title} subtitle={subtitle} />
      <form className="dictionary-add" onSubmit={submit}><input value={name} onChange={(event) => setName(event.target.value)} placeholder={`新增${title.replace("字典", "")}`} /><button className="button primary">添加</button></form>
      <div className="dictionary-list">
        {items.map((item, index) => <div key={item.id}><span><em>{String(index + 1).padStart(2, "0")}</em><b>{item.name}</b></span><button className="danger-text" onClick={() => window.confirm(`确定删除“${item.name}”？`) && onDelete(item.id)}>删除</button></div>)}
        {!items.length && <Empty title="字典为空" body="添加后即可在需求表单中选择。" />}
      </div>
    </div>
  );
}

function TemplatePage({ snapshot, start, end, onExportPpt, onExportDraft, onExportWorkspace, onInspectWorkspace, onApplyWorkspace, onExportWorkbook, onInspectWorkbook, onApplyWorkbook }: {
  snapshot: AppSnapshot;
  start: string;
  end: string;
  onExportPpt: () => Promise<void>;
  onExportDraft: () => Promise<void>;
  onExportWorkspace: () => Promise<void>;
  onInspectWorkspace: () => Promise<WorkspaceImportPreview | undefined>;
  onApplyWorkspace: (token: string, mode: WorkspaceImportMode) => Promise<boolean>;
  onExportWorkbook: () => Promise<void>;
  onInspectWorkbook: () => Promise<WorkspaceWorkbookImportPreview | undefined>;
  onApplyWorkbook: (token: string, conflictMode: WorkspaceWorkbookConflictMode) => Promise<boolean>;
}) {
  const [preview, setPreview] = useState<WorkspaceImportPreview | null>(null);
  const [workbookPreview, setWorkbookPreview] = useState<WorkspaceWorkbookImportPreview | null>(null);
  const [inspecting, setInspecting] = useState(false);
  const [workbookInspecting, setWorkbookInspecting] = useState(false);
  const [applying, setApplying] = useState<WorkspaceImportMode | null>(null);
  const [workbookApplying, setWorkbookApplying] = useState<WorkspaceWorkbookConflictMode | null>(null);
  const [exporting, setExporting] = useState(false);
  const imageCount = snapshot.requirements.reduce((sum, item) => sum + item.images.length, 0);

  async function inspectPackage() {
    setInspecting(true);
    try {
      const result = await onInspectWorkspace();
      if (result) setPreview(result);
    } finally {
      setInspecting(false);
    }
  }

  async function applyPackage(mode: WorkspaceImportMode) {
    if (!preview) return;
    setApplying(mode);
    try {
      if (await onApplyWorkspace(preview.token, mode)) setPreview(null);
    } finally {
      setApplying(null);
    }
  }

  async function inspectWorkbook() {
    setWorkbookInspecting(true);
    try {
      const result = await onInspectWorkbook();
      if (result) setWorkbookPreview(result);
    } finally {
      setWorkbookInspecting(false);
    }
  }

  async function applyWorkbook(conflictMode: WorkspaceWorkbookConflictMode) {
    if (!workbookPreview || workbookPreview.errors.length) return;
    setWorkbookApplying(conflictMode);
    try {
      if (await onApplyWorkbook(workbookPreview.token, conflictMode)) setWorkbookPreview(null);
    } finally {
      setWorkbookApplying(null);
    }
  }

  async function exportPpt() {
    setExporting(true);
    try {
      await onExportPpt();
    } finally {
      setExporting(false);
    }
  }

  return (
    <>
      <div className="template-layout">
        <div className="panel template-main">
          <div className="template-icon">P</div>
          <span className="eyebrow">OFFLINE POWERPOINT</span>
          <h2>需求路标模板 · 生产预览版 v0.3</h2>
          <p>按当前选择的半年区间，将汇总、运动/健康/海外研究三类路标、需求描述和本地图片自动生成可编辑 PowerPoint；共创版继续用于一起确认版式。</p>
          <div className="template-meta"><span>导出区间</span><b>{shortHalf(start)} – {shortHalf(end)}</b><span>状态</span><b className="ready">可导出 / 待确认</b><span>处理方式</span><b>全程本地</b></div>
          <div className="template-actions"><button className="button primary" onClick={exportPpt} disabled={exporting}>{exporting ? "正在生成…" : "生成需求路标 PPT"}</button><button className="button dark" onClick={onExportDraft}>导出 6 页共创版</button></div>
          <div className="template-contract"><h3>固定输出页面</h3><div>{["汇总分析页", "运动三泳道路标页", "健康三泳道路标页", "海外研究区域路标页", "领域全量详情页"].map((item, index) => <span key={item}><b>0{index + 1}</b>{item}</span>)}</div></div>
        </div>
        <div className="panel export-side transfer-side">
          <PanelTitle title="数据协作与交接" subtitle="Excel 协作编辑，.roadmap 无损备份" />
          <dl><div><dt>需求</dt><dd>{snapshot.requirements.length} 条</dd></div><div><dt>原始图片</dt><dd>{imageCount} 张</dd></div><div><dt>领域 / 产品</dt><dd>{snapshot.domains.length} / {snapshot.products.length}</dd></div><div><dt>路标卡片编辑</dt><dd>{snapshot.groupOverrides.length} 项</dd></div></dl>
          <div className="callout workbook-callout"><b>企业协作 Excel</b><span>导出全部结构化数据，上传到企业已有在线表格共同编辑；下载定稿后检查并合并导入。Excel 不携带图片。</span></div>
          <div className="transfer-actions"><button className="button primary" onClick={onExportWorkbook}>导出协作 Excel</button><button className="button ghost" onClick={inspectWorkbook} disabled={workbookInspecting}>{workbookInspecting ? "检查中…" : "导入协作 Excel"}</button></div>
          <div className="transfer-divider" />
          <div className="callout"><b>.roadmap 完整工作区</b><span>导出不受当前筛选和半年区间影响；包含需求描述、图片、字典、产品匹配和卡片编辑。</span></div>
          <div className="transfer-actions"><button className="button primary" onClick={onExportWorkspace}>导出完整数据包</button><button className="button ghost" onClick={inspectPackage} disabled={inspecting}>{inspecting ? "检查中…" : "导入数据包"}</button></div>
          <small className="transfer-note">Excel 负责多人协作；.roadmap 负责跨电脑完整交接原始图片，并兼容旧版 JSON 备份。</small>
        </div>
      </div>
      {workbookPreview && <div className="overlay import-overlay" onMouseDown={(event) => event.currentTarget === event.target && !workbookApplying && setWorkbookPreview(null)}>
        <div className="import-dialog workbook-import-dialog panel" role="dialog" aria-modal="true" aria-labelledby="workbook-import-title">
          <div className="import-dialog-head"><div><span className="eyebrow">EXCEL COLLABORATION IMPORT</span><h2 id="workbook-import-title">检查协作 Excel</h2></div><button className="icon-button" onClick={() => setWorkbookPreview(null)} disabled={Boolean(workbookApplying)}>×</button></div>
          <div className="import-file"><b>{workbookPreview.fileName}</b><span>{workbookPreview.formatVersion} · 导出于 {formatDateTime(workbookPreview.exportedAt)}</span></div>
          <dl className="import-counts workbook-counts"><div><dt>新增需求</dt><dd>{workbookPreview.counts.added} 条</dd></div><div><dt>修改需求</dt><dd>{workbookPreview.counts.updated} 条</dd></div><div><dt>删除需求</dt><dd>{workbookPreview.counts.deleted} 条</dd></div><div><dt>未变化</dt><dd>{workbookPreview.counts.unchanged} 条</dd></div><div><dt>领域 / 产品调整</dt><dd>{workbookPreview.counts.domainsChanged} / {workbookPreview.counts.productsChanged}</dd></div><div><dt>卡片调整</dt><dd>{workbookPreview.counts.groupOverridesChanged} 项</dd></div><div><dt>冲突</dt><dd>{workbookPreview.counts.conflicts} 项</dd></div><div><dt>校验错误</dt><dd>{workbookPreview.errors.length} 项</dd></div></dl>
          {workbookPreview.errors.length > 0 && <ImportIssueList title="需要先修正的错误" tone="error" items={workbookPreview.errors} />}
          {workbookPreview.errors.length === 0 && workbookPreview.conflicts.length > 0 && <ImportIssueList title="本机与 Excel 同时修改" tone="warning" items={workbookPreview.conflicts} />}
          {workbookPreview.errors.length === 0 && <div className="import-guidance"><b>图片处理</b><span>同 ID 需求保留本机图片；Excel 新增需求不带图片。默认按钮会保留发生冲突的本机版本。</span></div>}
          <div className="import-dialog-actions"><button className="button ghost" onClick={() => setWorkbookPreview(null)} disabled={Boolean(workbookApplying)}>取消</button><button className="button dark" onClick={() => applyWorkbook("local-wins")} disabled={Boolean(workbookApplying) || workbookPreview.errors.length > 0}>{workbookApplying === "local-wins" ? "合并中…" : "合并并保留本机冲突"}</button>{workbookPreview.conflicts.length > 0 && <button className="button danger" onClick={() => applyWorkbook("excel-wins")} disabled={Boolean(workbookApplying) || workbookPreview.errors.length > 0}>{workbookApplying === "excel-wins" ? "合并中…" : "Excel 覆盖冲突"}</button>}</div>
        </div>
      </div>}
      {preview && <div className="overlay import-overlay" onMouseDown={(event) => event.currentTarget === event.target && !applying && setPreview(null)}>
        <div className="import-dialog panel" role="dialog" aria-modal="true" aria-labelledby="import-title">
          <div className="import-dialog-head"><div><span className="eyebrow">WORKSPACE IMPORT</span><h2 id="import-title">确认导入数据包</h2></div><button className="icon-button" onClick={() => setPreview(null)} disabled={Boolean(applying)}>×</button></div>
          <div className="import-file"><b>{preview.fileName}</b><span>{preview.sourceFormat === "roadmap" ? `.roadmap · 模板 ${preview.templateVersion}` : "旧版 JSON 备份"}</span></div>
          <dl className="import-counts"><div><dt>需求</dt><dd>{preview.counts.requirements} 条</dd></div><div><dt>图片</dt><dd>{preview.counts.images} 张</dd></div><div><dt>领域</dt><dd>{preview.counts.domains} 个</dd></div><div><dt>产品</dt><dd>{preview.counts.products} 个</dd></div><div><dt>卡片编辑</dt><dd>{preview.counts.groupOverrides} 项</dd></div><div><dt>导出时间</dt><dd>{formatDateTime(preview.exportedAt)}</dd></div></dl>
          <div className="import-guidance"><b>选择导入方式</b><span>合并导入会按名称合并字典，并保留更新时间较新的同 ID 需求；整体替换会用数据包完整覆盖当前业务数据。</span></div>
          <div className="import-dialog-actions"><button className="button ghost" onClick={() => setPreview(null)} disabled={Boolean(applying)}>取消</button><button className="button dark" onClick={() => applyPackage("merge")} disabled={Boolean(applying)}>{applying === "merge" ? "合并中…" : "合并导入"}</button><button className="button danger" onClick={() => applyPackage("replace")} disabled={Boolean(applying)}>{applying === "replace" ? "替换中…" : "整体替换"}</button></div>
        </div>
      </div>}
    </>
  );
}

function ImportIssueList({ title, tone, items }: { title: string; tone: "error" | "warning"; items: WorkspaceWorkbookImportPreview["errors"] }) {
  return <div className={`import-issues ${tone}`}><b>{title}</b><ul>{items.slice(0, 8).map((item, index) => <li key={`${item.sheet}-${item.cell}-${index}`}><span>{item.sheet}!{item.cell}</span>{item.message}</li>)}</ul>{items.length > 8 && <small>另有 {items.length - 8} 项，请修正后重新选择文件检查。</small>}</div>;
}

function Metric({ label, value, suffix, accent = "plain" }: { label: string; value: string | number; suffix: string; accent?: string }) {
  return <div className={`metric ${accent}`}><span>{label}</span><div><strong>{value}</strong><small>{suffix}</small></div></div>;
}

function PanelTitle({ title, subtitle, action }: { title: string; subtitle: string; action?: ReactNode }) {
  return <div className="panel-title"><div><h3>{title}</h3><p>{subtitle}</p></div>{action ?? <span>•••</span>}</div>;
}

function SourceBadge({ source }: { source: RequirementSource }) {
  return <span className="source-badge"><i style={{ background: SOURCE_COLORS[source] }} />{source}</span>;
}

function Field({ label, required, children }: { label: string; required?: boolean; children: ReactNode }) {
  return <label className="field"><span>{label}{required && <b>*</b>}</span>{children}</label>;
}

function Empty({ title, body }: { title: string; body: string }) {
  return <div className="empty"><i>＋</i><strong>{title}</strong><span>{body}</span></div>;
}

function pageTitle(page: Page): string {
  return ({ requirements: "需求池", summary: "汇总分析", roadmap: "路标预览", dictionaries: "字典设置", template: "PPT 输出" })[page];
}

function pageSubtitle(page: Page): string {
  return ({ requirements: "结构化维护上线需求，双击任意行快速编辑", summary: "按时间观察三端投入、业务占比与体验优化趋势", roadmap: "运动、健康、海外研究三路标，按层级或区域展开月份", dictionaries: "维护领域与产品的统一口径", template: "共同确认内置模板，后续按规范匹配需求内容输出" })[page];
}

function formatNumber(value: number): string {
  return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 2 }).format(value);
}

function workloadBreakdownItems(workload: WorkloadBreakdown): Array<[string, number]> {
  const items: Array<[string, number]> = [["设备", workload.device], ["App", workload.app], ["云侧", workload.cloud]];
  if (workload.unallocated > 0) items.push(["待拆分", workload.unallocated]);
  return items;
}

function formatWorkloadBreakdown(workload: WorkloadBreakdown): string {
  return workloadBreakdownItems(workload).map(([label, value]) => `${label} ${formatNumber(value)}`).join(" · ");
}

function formatMonth(value: string): string {
  const [year, month] = value.split("-");
  return `${year}年${Number(month)}月`;
}

function formatDateTime(value: string): string {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : new Intl.DateTimeFormat("zh-CN", { dateStyle: "medium", timeStyle: "short" }).format(date);
}

function shortHalf(value: string): string {
  const match = /^(\d{4})H([12])$/.exec(value);
  return match ? `${match[1].slice(2)}H${match[2]}` : value;
}

function messageOf(reason: unknown): string {
  if (reason instanceof Error) return reason.message.replace(/^Error invoking remote method '[^']+': Error: /, "");
  return String(reason);
}

async function readRequirementImages(files: FileList | null, existingCount: number): Promise<{ images: RequirementImage[]; error?: string }> {
  if (!files?.length) return { images: [] };
  const available = Math.max(0, MAX_REQUIREMENT_IMAGES - existingCount);
  if (available === 0) return { images: [], error: `每条需求最多添加 ${MAX_REQUIREMENT_IMAGES} 张图片` };
  const selected = Array.from(files).slice(0, available);
  const images: RequirementImage[] = [];
  const errors: string[] = [];

  if (files.length > available) errors.push(`仅添加前 ${available} 张图片`);
  for (const file of selected) {
    if (!REQUIREMENT_IMAGE_MIME_TYPES.includes(file.type as RequirementImage["mimeType"])) {
      errors.push(`${file.name} 格式不支持`);
      continue;
    }
    if (file.size > MAX_REQUIREMENT_IMAGE_BYTES) {
      errors.push(`${file.name} 超过 5MB`);
      continue;
    }
    images.push({
      id: globalThis.crypto.randomUUID(),
      name: file.name,
      mimeType: file.type as RequirementImage["mimeType"],
      sizeBytes: file.size,
      dataUrl: await fileToDataUrl(file),
    });
  }
  return { images, error: errors.length ? errors.join("；") : undefined };
}

function fileToDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error(`无法读取图片：${file.name}`));
    reader.readAsDataURL(file);
  });
}
