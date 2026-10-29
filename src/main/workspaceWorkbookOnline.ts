import { createHash, randomUUID } from "node:crypto";
import ExcelJS from "exceljs";
import type { Requirement, RequirementDraft, TieringSupportStatus, WorkspaceData, WorkspaceWorkbookFieldChange, WorkspaceWorkbookFieldSelection, WorkspaceWorkbookImportCounts, WorkspaceWorkbookIssue } from "../shared/types.js";
import { CATEGORIES, REQUIREMENT_SCOPES } from "../shared/types.js";
import { BUILT_IN_TIERING_PRODUCT_SERIES, normalizeWorkspaceTieringScope, normalizeTieringData, tieringSupportId } from "../shared/tiering.js";
import { roundWorkload } from "../shared/workload.js";
import { validateWorkspaceData } from "./workspacePackage.js";

const SHEET_NAME = "需求清单";
const NEW_ROW_COUNT = 300;
const FIELD_HEADERS = {
  title: "需求标题", description: "需求描述", domainL0Id: "领域 L0", domainId: "领域 L1",
  scope: "适用范围", allowLatestVersionUpgrade: "允许最近版本升级", category: "分类", targetMonth: "上线月份",
  deviceWorkloadPm: "设备工作量", appWorkloadPm: "App 工作量", cloudWorkloadPm: "云侧工作量",
  planningMemberId: "规划", productManagerMemberId: "产品经理", seMemberId: "SE", algorithmMemberId: "算法",
} as const;
type FieldKey = keyof typeof FIELD_HEADERS;
type Fields = Record<FieldKey, string>;
const FIELD_KEYS = Object.keys(FIELD_HEADERS) as FieldKey[];
const OWNER_ROLES: Partial<Record<FieldKey, string>> = { planningMemberId: "规划", productManagerMemberId: "产品经理", seMemberId: "SE", algorithmMemberId: "算法" };
const WORKLOAD_KEYS = new Set<FieldKey>(["deviceWorkloadPm", "appWorkloadPm", "cloudWorkloadPm"]);
const WORKLOAD_DECIMAL = /^(?:\d+(?:\.\d{1,2})?|\.\d{1,2})$/;
export const ONLINE_WORKBOOK_HEADERS = ["序号", "需求编号", "操作", ...Object.values(FIELD_HEADERS), ...BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => item.name)] as const;

interface OnlineRow {
  row: number;
  id: string;
  action: "保留" | "删除" | "新增";
  fields: Fields;
  series: Record<string, TieringSupportStatus>;
  isNew: boolean;
  wasDraft: boolean;
  tombstoned: boolean;
}

export interface OnlineWorkbookInspection {
  rows: OnlineRow[];
  changes: WorkspaceWorkbookFieldChange[];
  errors: WorkspaceWorkbookIssue[];
  counts: WorkspaceWorkbookImportCounts;
  digest: string;
  localFingerprint: string;
  exportedAt: string;
}

function hash(value: unknown): string { return createHash("sha256").update(JSON.stringify(value)).digest("hex"); }
function issue(errors: WorkspaceWorkbookIssue[], row: number, column: number, message: string): void {
  errors.push({ sheet: SHEET_NAME, cell: `${columnLetter(column)}${row}`, message });
}
function columnLetter(value: number): string {
  let result = "";
  for (let n = value; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result;
  return result;
}
function text(value: ExcelJS.CellValue): string {
  if (value instanceof Date) return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, "0")}`;
  if (typeof value === "object" && value !== null) {
    if ("text" in value) return String(value.text).trim();
    if ("richText" in value && Array.isArray(value.richText)) return value.richText.map((part) => part.text).join("").trim();
  }
  return String(value ?? "").trim();
}
function month(value: ExcelJS.CellValue): string {
  if (value instanceof Date) return text(value);
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86_400_000);
    return text(date);
  }
  const raw = text(value);
  const match = /^(\d{4})[-/.年](\d{1,2})(?:月)?$/.exec(raw);
  return match ? `${match[1]}-${match[2].padStart(2, "0")}` : raw;
}
function memberLabel(data: WorkspaceData, id: string, role: string): string {
  const member = data.members.find((item) => item.id === id && item.role === role);
  if (!member) return "";
  return data.members.filter((item) => item.name === member.name && item.role === role).length > 1
    ? `${member.name}（${member.employeeNumber}）` : member.name;
}
function fieldsFor(data: WorkspaceData, item: Requirement): Fields {
  return {
    title: item.title, description: item.description,
    domainL0Id: data.domains.find((domain) => domain.id === item.domainL0Id)?.name ?? item.domainL0Id,
    domainId: data.domains.find((domain) => domain.id === item.domainId)?.name ?? item.domainId,
    scope: item.scope, allowLatestVersionUpgrade: item.allowLatestVersionUpgrade ? "是" : "否",
    category: item.category, targetMonth: item.targetMonth,
    deviceWorkloadPm: String(roundWorkload(item.deviceWorkloadPm)), appWorkloadPm: String(roundWorkload(item.appWorkloadPm)), cloudWorkloadPm: String(roundWorkload(item.cloudWorkloadPm)),
    planningMemberId: memberLabel(data, item.planningMemberId ?? "", "规划"), productManagerMemberId: memberLabel(data, item.productManagerMemberId ?? "", "产品经理"),
    seMemberId: memberLabel(data, item.seMemberId ?? "", "SE"), algorithmMemberId: memberLabel(data, item.algorithmMemberId ?? "", "算法"),
  };
}
function statusesFor(data: WorkspaceData, item: Requirement, draft: boolean): Record<string, TieringSupportStatus> {
  const feature = data.tiering?.features.find((candidate) => candidate.requirementId === item.id);
  const pending = draft ? (item as RequirementDraft).pendingSeries ?? {} : {};
  return Object.fromEntries(BUILT_IN_TIERING_PRODUCT_SERIES.map((series) => [series.id,
    pending[series.id] ?? (feature ? data.tiering?.supportRecords.find((record) => record.id === tieringSupportId(feature.id, series.id))?.status : undefined) ?? "未判断",
  ])) as Record<string, TieringSupportStatus>;
}

/** The v12 sheet deliberately carries no custom XML, hidden cells, row formats or workbook properties. */
export async function createOnlineWorkbook(data: WorkspaceData, exportedAt: string): Promise<Buffer> {
  validateWorkspaceData(data);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "解决方案需求管理";
  workbook.created = new Date(exportedAt);
  const sheet = workbook.addWorksheet(SHEET_NAME, { properties: { tabColor: { argb: "FF228A70" } } });
  sheet.views = [{ state: "frozen", xSplit: 3, ySplit: 1, showGridLines: false }];
  sheet.addRow([...ONLINE_WORKBOOK_HEADERS]);
  const items = [...data.requirements.map((item) => ({ item, draft: false })), ...(data.draftRequirements ?? []).map((item) => ({ item, draft: true }))];
  items.forEach(({ item, draft }, index) => {
    const fields = fieldsFor(data, item);
    const series = statusesFor(data, item, draft);
    sheet.addRow([index + 1, item.id, "保留", ...FIELD_KEYS.map((key) => fields[key]), ...BUILT_IN_TIERING_PRODUCT_SERIES.map((entry) => series[entry.id])]);
  });
  for (let index = 0; index < NEW_ROW_COUNT; index += 1) sheet.addRow([null, `NEW-${randomUUID()}`, "新增"]);
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: items.length + NEW_ROW_COUNT + 1, column: ONLINE_WORKBOOK_HEADERS.length } };
  sheet.columns = [8, 43, 10, 30, 46, 18, 18, 16, 18, 16, 16, 16, 16, 16, 19, 19, 19, 19, ...BUILT_IN_TIERING_PRODUCT_SERIES.map(() => 19)].map((width) => ({ width }));
  sheet.getRow(1).height = 30;
  sheet.getRow(1).eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9F2E6" } };
    cell.font = { name: "Microsoft YaHei", size: 10, bold: true, color: { argb: "FF000000" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    sheet.getRow(row).height = 27;
    sheet.getRow(row).eachCell({ includeEmpty: true }, (cell) => {
      cell.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF000000" } };
      cell.alignment = { vertical: "middle", wrapText: cell.address.startsWith("D") || cell.address.startsWith("E") };
    });
  }
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function localFingerprint(data: WorkspaceData): string {
  return hash({ requirements: data.requirements, drafts: data.draftRequirements ?? [], tombstones: data.requirementTombstones ?? [], domains: data.domains, members: data.members, tiering: data.tiering?.supportRecords });
}
function emptyCounts(): WorkspaceWorkbookImportCounts {
  return { added: 0, updated: 0, deleted: 0, unchanged: 0, conflicts: 0, domainsChanged: 0, productsChanged: 0, groupOverridesChanged: 0 };
}
function resolveField(data: WorkspaceData, key: FieldKey, raw: string, row: number, column: number, errors: WorkspaceWorkbookIssue[], allowEmptyMonth = false): string {
  if (key === "domainL0Id" || key === "domainId") {
    const level = key === "domainL0Id" ? "L0" : "L1";
    const candidates = data.domains.filter((item) => (item.level ?? "L1") === level && (item.name === raw || item.id === raw));
    if (candidates.length !== 1) issue(errors, row, column, `${FIELD_HEADERS[key]}“${raw || "空"}”无法在本机字典中唯一匹配`);
    return candidates[0]?.id ?? raw;
  }
  if (OWNER_ROLES[key]) {
    if (!raw) return "";
    const candidates = data.members.filter((item) => item.role === OWNER_ROLES[key] && (item.name === raw || `${item.name}（${item.employeeNumber}）` === raw || item.id === raw));
    if (candidates.length !== 1) issue(errors, row, column, `${FIELD_HEADERS[key]}“${raw}”无法在本机人员字典中唯一匹配`);
    return candidates[0]?.id ?? raw;
  }
  if (key === "targetMonth") {
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(raw) && !(allowEmptyMonth && !raw)) issue(errors, row, column, "上线月份必须填写有效年月");
    return raw;
  }
  if (WORKLOAD_KEYS.has(key)) {
    const value = raw === "" ? 0 : Number(raw);
    if (!Number.isFinite(value) || value < 0 || value > 99999999 || (raw !== "" && !WORKLOAD_DECIMAL.test(raw))) issue(errors, row, column, "工作量须为非负数字，最多两位小数");
    return String(roundWorkload(value || 0));
  }
  if (key === "allowLatestVersionUpgrade") {
    if (!["是", "否"].includes(raw)) issue(errors, row, column, "请选择是或否");
  }
  if (key === "scope" && !REQUIREMENT_SCOPES.includes(raw as Requirement["scope"])) issue(errors, row, column, "适用范围无效");
  if (key === "category" && !CATEGORIES.includes(raw as Requirement["category"])) issue(errors, row, column, "分类无效");
  if (key === "title" && !raw) issue(errors, row, column, "需求标题不能为空");
  if (key === "description" && raw.length > 5000) issue(errors, row, column, "需求描述不能超过 5000 字");
  return raw;
}
function normalizedFields(data: WorkspaceData, fields: Fields, row: number, columns: Map<string, number>, errors: WorkspaceWorkbookIssue[], allowEmptyMonth = false): Fields {
  const result = { ...fields };
  for (const key of FIELD_KEYS) result[key] = resolveField(data, key, fields[key], row, columns.get(FIELD_HEADERS[key])!, errors, allowEmptyMonth);
  if (data.domains.find((item) => item.id === result.domainId)?.parentId !== result.domainL0Id) issue(errors, row, columns.get("领域 L1")!, "领域 L1 不属于所选领域 L0");
  return result;
}
function valuesFor(data: WorkspaceData, item: Requirement): Fields {
  const displayed = fieldsFor(data, item);
  return Object.fromEntries(FIELD_KEYS.map((key) => [key, key === "domainL0Id" || key === "domainId" ? item[key] : OWNER_ROLES[key] ? item[key] ?? "" : displayed[key]])) as Fields;
}

export function inspectOnlineWorkbook(sheet: ExcelJS.Worksheet, current: WorkspaceData, _fileName: string): OnlineWorkbookInspection {
  const errors: WorkspaceWorkbookIssue[] = [];
  const headers = new Map<string, number>();
  sheet.getRow(1).eachCell((cell, index) => {
    const label = text(cell.value);
    if (headers.has(label)) issue(errors, 1, index, `表头“${label}”重复`);
    headers.set(label, index);
  });
  for (const header of ONLINE_WORKBOOK_HEADERS) if (!headers.has(header)) issue(errors, 1, 1, `缺少列“${header}”`);
  const rows: OnlineRow[] = [];
  const changes: WorkspaceWorkbookFieldChange[] = [];
  const counts = emptyCounts();
  const seen = new Set<string>();
  const active = new Map(current.requirements.map((item) => [item.id, item]));
  const drafts = new Map((current.draftRequirements ?? []).map((item) => [item.id, item]));
  const tombstones = new Set((current.requirementTombstones ?? []).map((item) => item.id));
  if (errors.length) return { rows, changes, errors, counts, digest: hash([]), localFingerprint: localFingerprint(current), exportedAt: new Date().toISOString() };
  for (let row = 2; row <= sheet.actualRowCount; row += 1) {
    const value = (header: string) => {
      const cell = sheet.getCell(row, headers.get(header)!);
      if (cell.type === ExcelJS.ValueType.Formula) issue(errors, row, headers.get(header)!, "业务字段不能使用公式");
      return header === "上线月份" ? month(cell.value) : text(cell.value);
    };
    const id = value("需求编号");
    const title = value("需求标题");
    const actionText = value("操作");
    const otherBusiness = FIELD_KEYS.filter((key) => key !== "title").some((key) => Boolean(value(FIELD_HEADERS[key])));
    const anySeries = BUILT_IN_TIERING_PRODUCT_SERIES.some((series) => Boolean(value(series.name)));
    if (!id && !title && !otherBusiness && !anySeries) continue;
    if (!id) { issue(errors, row, headers.get("需求编号")!, "需求编号缺失；请从原导出表复制该编号"); continue; }
    if (seen.has(id)) { issue(errors, row, headers.get("需求编号")!, `需求编号“${id}”重复`); continue; }
    seen.add(id);
    const existing = active.get(id) ?? drafts.get(id);
    const isNew = !existing && !tombstones.has(id);
    if (isNew && !/^NEW-[0-9a-f-]{36}$/i.test(id)) { issue(errors, row, headers.get("需求编号")!, `需求编号“${id}”不属于本机工作区或预留新增行`); continue; }
    if (!existing && !title && !otherBusiness && !anySeries) continue;
    const action = actionText || (isNew ? "新增" : "保留");
    if (!["保留", "删除", "新增"].includes(action)) issue(errors, row, headers.get("操作")!, "操作只能为保留、删除或新增");
    if (isNew && action === "删除") { issue(errors, row, headers.get("操作")!, "未导入的新行无需标记删除"); continue; }
    if (isNew && action !== "新增") issue(errors, row, headers.get("操作")!, "预留行填写后操作应为“新增”");
    const rawFields = Object.fromEntries(FIELD_KEYS.map((key) => [key, value(FIELD_HEADERS[key])])) as Fields;
    if (isNew || drafts.has(id)) {
      if (!rawFields.scope) rawFields.scope = "全球";
      if (!rawFields.allowLatestVersionUpgrade) rawFields.allowLatestVersionUpgrade = "否";
      if (!rawFields.category) rawFields.category = "平台演进";
    }
    const fields = action === "删除" && existing ? valuesFor(current, existing) : normalizedFields(current, rawFields, row, headers, errors, Boolean(existing && !existing.targetMonth));
    const series = Object.fromEntries(BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => [item.id, (value(item.name) || (isNew || drafts.has(id) ? "未判断" : "")) as TieringSupportStatus])) as Record<string, TieringSupportStatus>;
    for (const item of BUILT_IN_TIERING_PRODUCT_SERIES) if (action !== "删除" && !["支持", "不支持", "未判断"].includes(series[item.id])) issue(errors, row, headers.get(item.name)!, `${item.name}只能填写支持、不支持、未判断`);
    const parsed: OnlineRow = { row, id, action: action as OnlineRow["action"], fields, series, isNew, wasDraft: drafts.has(id), tombstoned: tombstones.has(id) };
    rows.push(parsed);
    const displayTitle = existing?.title ?? title;
    if (action === "删除") {
      if (existing) changes.push({ id: `${id}:delete`, requirementId: id, row, title: displayTitle, kind: "delete", field: "删除整条需求", baselineValue: "", currentValue: displayTitle, excelValue: "删除", conflict: true });
      continue;
    }
    if (!existing) {
      changes.push({ id: `${id}:${tombstones.has(id) ? "restore" : "add"}`, requirementId: id, row, title: displayTitle, kind: tombstones.has(id) ? "restore" : "add", field: tombstones.has(id) ? "恢复已删除需求" : "新增需求", baselineValue: "", currentValue: tombstones.has(id) ? "本机已删除" : "本机无此需求", excelValue: displayTitle, conflict: tombstones.has(id) });
      counts.added += 1;
      continue;
    }
    const local = valuesFor(current, existing);
    for (const key of FIELD_KEYS) if (fields[key] !== local[key]) changes.push({ id: `${id}:field:${key}`, requirementId: id, row, title: displayTitle, kind: "field", field: FIELD_HEADERS[key], baselineValue: "", currentValue: fieldsFor(current, existing)[key], excelValue: rawFields[key], conflict: true });
    const localSeries = statusesFor(current, existing, drafts.has(id));
    for (const item of BUILT_IN_TIERING_PRODUCT_SERIES) if (series[item.id] !== localSeries[item.id]) changes.push({ id: `${id}:series:${item.id}`, requirementId: id, row, title: displayTitle, kind: "series", field: item.name, baselineValue: "", currentValue: localSeries[item.id], excelValue: series[item.id], conflict: true });
    if (drafts.has(id) && fields.planningMemberId && fields.productManagerMemberId && fields.seMemberId && fields.category !== "产品专属") {
      changes.push({ id: `${id}:promote`, requirementId: id, row, title: displayTitle, kind: "promote", field: "转为正式需求", baselineValue: "", currentValue: "待完善", excelValue: "正式需求", conflict: false });
    }
    counts.updated += 1;
  }
  counts.deleted = changes.filter((change) => change.kind === "delete").length;
  counts.updated = new Set(changes.filter((change) => change.kind === "field" || change.kind === "series").map((change) => change.requirementId)).size;
  counts.unchanged = rows.length - counts.added - counts.updated - counts.deleted;
  counts.conflicts = changes.filter((change) => change.conflict).length;
  return { rows, changes, errors, counts, digest: hash({ rows, changes }), localFingerprint: localFingerprint(current), exportedAt: new Date().toISOString() };
}

function applyFields(item: Requirement, fields: Fields, keys: readonly FieldKey[], now: string): Requirement {
  const updated = { ...item };
  for (const key of keys) {
    const value = fields[key];
    if (key === "allowLatestVersionUpgrade") updated.allowLatestVersionUpgrade = value === "是";
    else if (WORKLOAD_KEYS.has(key)) (updated as unknown as Record<string, unknown>)[key] = Number(value);
    else (updated as unknown as Record<string, unknown>)[key] = value;
  }
  updated.workloadPm = roundWorkload(updated.deviceWorkloadPm + updated.appWorkloadPm + updated.cloudWorkloadPm + (updated.researchAppWorkloadPm ?? 0) + (updated.algorithmWorkloadPm ?? 0) + updated.unallocatedWorkloadPm);
  updated.updatedAt = now;
  return updated;
}
function buildNewRequirement(row: OnlineRow, now: string): RequirementDraft {
  const fields = row.fields;
  const base: Requirement = {
    id: row.id, title: fields.title, description: fields.description, images: [], domainL0Id: fields.domainL0Id, domainId: fields.domainId,
    scope: fields.scope as Requirement["scope"], allowLatestVersionUpgrade: fields.allowLatestVersionUpgrade === "是", overseasRegions: [],
    category: fields.category as Requirement["category"], targetMonth: fields.targetMonth, productIds: [],
    deviceWorkloadPm: Number(fields.deviceWorkloadPm), appWorkloadPm: Number(fields.appWorkloadPm), cloudWorkloadPm: Number(fields.cloudWorkloadPm),
    researchAppWorkloadPm: 0, algorithmWorkloadPm: 0, unallocatedWorkloadPm: 0,
    workloadPm: roundWorkload(Number(fields.deviceWorkloadPm) + Number(fields.appWorkloadPm) + Number(fields.cloudWorkloadPm)),
    planningMemberId: fields.planningMemberId, productManagerMemberId: fields.productManagerMemberId, seMemberId: fields.seMemberId, algorithmMemberId: fields.algorithmMemberId,
    createdAt: now, updatedAt: now,
  };
  return { ...base, pendingSeries: row.series };
}
function ready(item: Requirement): boolean {
  return Boolean(item.planningMemberId && item.productManagerMemberId && item.seMemberId && (item.category !== "产品专属" || item.productIds.length === 1));
}
export function applyOnlineWorkbook(current: WorkspaceData, inspection: OnlineWorkbookInspection, selection: WorkspaceWorkbookFieldSelection | undefined, now: string): WorkspaceData {
  if (inspection.errors.length) throw new Error("在线协作 Excel 有校验错误，请先修正");
  if (!selection || !Array.isArray(selection.selectedChangeIds) || !selection.conflictDecisions) throw new Error("请选择需要合入的差异");
  if (inspection.localFingerprint !== localFingerprint(current)) throw new Error("本机数据在预览后变化，请重新检查 Excel");
  if (inspection.digest !== hash({ rows: inspection.rows, changes: inspection.changes })) throw new Error("差异校验失败，请重新检查 Excel");
  const byId = new Map(inspection.changes.map((item) => [item.id, item]));
  const selected = new Set(selection.selectedChangeIds);
  if (selected.size !== selection.selectedChangeIds.length || [...selected].some((id) => !byId.has(id))) throw new Error("所选差异不属于当前预览");
  for (const id of selected) if (byId.get(id)?.conflict && !["local", "excel"].includes(selection.conflictDecisions[id])) throw new Error("选中的冲突项尚未决策");
  const adopted = (id: string) => selected.has(id) && (!byId.get(id)?.conflict || selection.conflictDecisions[id] === "excel");
  if (![...selected].some(adopted)) return current;
  const requirements = new Map(current.requirements.map((item) => [item.id, item]));
  const drafts = new Map((current.draftRequirements ?? []).map((item) => [item.id, item]));
  const tombstones = new Map((current.requirementTombstones ?? []).map((item) => [item.id, item]));
  const seriesChanges: Array<{ id: string; seriesId: string; status: TieringSupportStatus }> = [];
  for (const row of inspection.rows) {
    if (adopted(`${row.id}:delete`)) {
      const existing = requirements.get(row.id) ?? drafts.get(row.id);
      if (existing) tombstones.set(row.id, { id: row.id, title: existing.title, deletedAt: now });
      requirements.delete(row.id); drafts.delete(row.id);
      continue;
    }
    if (row.action === "删除") continue;
    if (adopted(`${row.id}:add`) || adopted(`${row.id}:restore`)) {
      const item = buildNewRequirement(row, now);
      if (ready(item)) { const { pendingSeries: _pendingSeries, ...formal } = item; requirements.set(row.id, formal); }
      else drafts.set(row.id, item);
      tombstones.delete(row.id);
      for (const [seriesId, status] of Object.entries(row.series)) seriesChanges.push({ id: row.id, seriesId, status });
      continue;
    }
    const existing = requirements.get(row.id) ?? drafts.get(row.id);
    if (!existing) continue;
    const keys = FIELD_KEYS.filter((key) => adopted(`${row.id}:field:${key}`));
    let next = keys.length ? applyFields(existing, row.fields, keys, now) : existing;
    const changedSeries = BUILT_IN_TIERING_PRODUCT_SERIES.filter((series) => adopted(`${row.id}:series:${series.id}`));
    if (drafts.has(row.id)) {
      const pendingSeries = { ...(existing as RequirementDraft).pendingSeries };
      for (const series of changedSeries) pendingSeries[series.id] = row.series[series.id];
      next = { ...next, pendingSeries } as RequirementDraft;
      if (adopted(`${row.id}:promote`) && !ready(next)) throw new Error(`需求“${next.title}”尚未补齐规划、产品经理和 SE，不能转为正式需求`);
      if (ready(next) && adopted(`${row.id}:promote`)) {
        const { pendingSeries: _pendingSeries, ...formal } = next as RequirementDraft;
        drafts.delete(row.id); requirements.set(row.id, formal);
        for (const [seriesId, status] of Object.entries(pendingSeries)) seriesChanges.push({ id: row.id, seriesId, status });
      } else drafts.set(row.id, next as RequirementDraft);
    } else {
      requirements.set(row.id, next);
      for (const series of changedSeries) seriesChanges.push({ id: row.id, seriesId: series.id, status: row.series[series.id] });
    }
  }
  let result: WorkspaceData = normalizeWorkspaceTieringScope({ ...current, requirements: [...requirements.values()], draftRequirements: [...drafts.values()], requirementTombstones: [...tombstones.values()],
    roadmapAssignments: current.roadmapAssignments?.filter((item) => requirements.has(item.requirementId)),
    medicalRoadmapPlans: current.medicalRoadmapPlans?.filter((item) => requirements.has(item.requirementId)) });
  const tiering = normalizeTieringData(result.tiering);
  const features = new Map(tiering.features.filter((item) => item.requirementId).map((item) => [item.requirementId!, item]));
  const records = new Map(tiering.supportRecords.map((item) => [item.id, item]));
  for (const change of seriesChanges) {
    const feature = features.get(change.id);
    if (!feature) continue;
    const key = tieringSupportId(feature.id, change.seriesId);
    const previous = records.get(key);
    if (!previous || previous.status === change.status) continue;
    const updated = { ...previous, status: change.status, updatedAt: now };
    records.set(key, updated);
    tiering.changeLogs.push({ id: randomUUID(), changedAt: now, triggerType: "在线协作 Excel", featureId: feature.id, seriesId: change.seriesId,
      before: JSON.stringify(previous), after: JSON.stringify(updated), reason: "协作 Excel v12 勾选合入", sourceIds: [...updated.sourceIds], reviewer: "", recordStatus: "有效", note: "", updatedAt: now });
  }
  tiering.supportRecords = [...records.values()];
  result = { ...result, tiering };
  try { validateWorkspaceData(result); }
  catch (error) { throw new Error(`所选差异组合无效，未写入任何数据：${error instanceof Error ? error.message : String(error)}`); }
  return result;
}
