import { createHash } from "node:crypto";
import { randomUUID } from "node:crypto";
import { deflateRawSync, inflateRawSync } from "node:zlib";
import { REQUIREMENT_OWNER_FIELDS } from "../shared/members.js";
import type { Member } from "../shared/types.js";
import { isMedicalSubmissionSource } from "../shared/medicalSubmission.js";
import { normalizeRoadmapManagement } from "../shared/roadmapManagement.js";
import { isTieringEligibleSource } from "../shared/tiering.js";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import type {
  DictionaryItem,
  DomainLevel,
  GroupOverride,
  HalfYearRange,
  Requirement,
  WorkspaceData,
  WorkspaceWorkbookConflictMode,
  WorkspaceWorkbookImportCounts,
  WorkspaceWorkbookImportPreview,
  WorkspaceWorkbookFieldChange,
  WorkspaceWorkbookFieldSelection,
  WorkspaceWorkbookIssue,
} from "../shared/types.js";
import { CATEGORIES, REQUIREMENT_SCOPES, SOURCES } from "../shared/types.js";
import { normalizeLegacyGroupOverride, normalizeLegacySource } from "../shared/requirements.js";
import { halfYearSequence, monthsForHalfYear, requirementsInRange } from "../shared/roadmap.js";
import { roundWorkload } from "../shared/workload.js";
import { diffWorkspaceData, workspaceDifferenceDigest } from "../shared/workspaceDiff.js";
import { validateWorkspaceData } from "./workspacePackage.js";
import { BUILT_IN_TIERING_PRODUCT_SERIES, isHistoricalTieringFeature, isTieringEligibleRequirement, normalizeTieringData, normalizeWorkspaceTieringScope, renameIndependentTieringDomain, resolveTieringFeature, tieringSupportId } from "../shared/tiering.js";

const FORMAT_KEY = "health-roadmap-collaboration";
const FORMAT_VERSION = "v11";
const LEGACY_FORMAT_VERSIONS = ["v2", "v3", "v4", "v5", "v6", "v7", "v8", "v9", "v10"] as const;
const SUPPORTED_FORMAT_VERSIONS: readonly string[] = [FORMAT_VERSION, ...LEGACY_FORMAT_VERSIONS];
const MAX_WORKBOOK_BYTES = 25 * 1024 * 1024;
const LEGACY_HEADER_ROW = 3;
const LEGACY_FIRST_DATA_ROW = LEGACY_HEADER_ROW + 1;
const HEADER_ROW = LEGACY_HEADER_ROW;
const FIRST_DATA_ROW = LEGACY_FIRST_DATA_ROW;
const V6_HEADER_ROW = 1;
const V6_FIRST_DATA_ROW = V6_HEADER_ROW + 1;
const V6_CUSTOM_XML_PATH = "customXml/item1.xml";
const V6_CUSTOM_XML_NAMESPACE = "urn:solution-roadmap:collaboration:v6";
const V6_CUSTOM_XML_REL_ID = "rIdRoadmapCollaborationV6";
const WORKBOOK_CUSTOM_PROPERTIES_PATH = "docProps/custom.xml";
const WORKBOOK_CUSTOM_PROPERTIES_PREFIX = "RoadmapCollaboration_";
const WORKBOOK_CUSTOM_PROPERTIES_CHUNK_SIZE = 240;
const WORKBOOK_CUSTOM_PROPERTIES_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
const WORKBOOK_CUSTOM_PROPERTIES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
const SHEETS = {
  guide: "填写说明",
  requirements: "需求清单",
  domainL0s: "领域L0字典",
  domainL1s: "领域L1字典",
  products: "产品字典",
  overrides: "路标卡片",
} as const;

const REQUIREMENT_HEADERS = [
  "操作",
  "需求标题",
  "需求描述",
  "领域L0",
  "领域L1",
  "来源",
  "适用范围",
  "分类",
  "上线月份",
  "匹配产品（单选）",
  "设备工作量（人月）",
  "App工作量（人月）",
  "云侧工作量（人月）",
  "待拆分工作量（只读）",
  "总工作量（自动合计）",
  "图片数量（只读）",
  "需求ID（请勿修改）",
  "领域L0 ID（请勿修改）",
  "领域L1 ID（请勿修改）",
  "产品ID（请勿修改）",
  "基线领域L0名称（请勿修改）",
  "基线领域L1名称（请勿修改）",
  "基线产品名称（请勿修改）",
  "创建时间（请勿修改）",
  "基线更新时间（请勿修改）",
  "基线校验（请勿修改）",
] as const;

const V6_REQUIREMENT_HEADERS = [
  "操作",
  "需求标题",
  "需求描述",
  "领域 L0",
  "领域 L1",
  "来源",
  "适用范围",
  "允许最近版本升级",
  "分类",
  "上线月份",
  "匹配产品",
  "设备工作量",
  "App 工作量",
  "云侧工作量",
] as const;

const WORKBOOK_OWNER_FIELDS = [REQUIREMENT_OWNER_FIELDS[1], REQUIREMENT_OWNER_FIELDS[0], REQUIREMENT_OWNER_FIELDS[2], REQUIREMENT_OWNER_FIELDS[3]];
const V7_REQUIREMENT_HEADERS = [...V6_REQUIREMENT_HEADERS, "规划", "产品经理", "SE", "算法"] as const;
const V8_BUSINESS_HEADERS = V7_REQUIREMENT_HEADERS.filter((header) => header !== "来源" && header !== "匹配产品");
// Keep the original 13-series layouts immutable: older workbooks must still
// validate and import after the application gains new built-in series.
const LEGACY_TIERING_PRODUCT_SERIES = BUILT_IN_TIERING_PRODUCT_SERIES.filter((series) => series.id !== "PS14" && series.id !== "PS15");
const V8_REQUIREMENT_HEADERS = [...V8_BUSINESS_HEADERS, ...LEGACY_TIERING_PRODUCT_SERIES.map((series) => series.name)];
const V9_REQUIREMENT_HEADERS = ["序号", ...V8_REQUIREMENT_HEADERS];
const V10_REQUIREMENT_HEADERS = ["序号", ...V8_BUSINESS_HEADERS, ...BUILT_IN_TIERING_PRODUCT_SERIES.map((series) => series.name)];
type SelectiveWorkbookVersion = "v8" | "v9" | "v10" | "v11";
const SELECTIVE_WORKBOOK_VERSIONS: readonly string[] = ["v8", "v9", "v10", "v11"];
function isSelectiveWorkbookVersion(version: string): version is SelectiveWorkbookVersion {
  return SELECTIVE_WORKBOOK_VERSIONS.includes(version);
}
function workbookSeriesForVersion(version: SelectiveWorkbookVersion) {
  return version === "v10" || version === "v11" ? BUILT_IN_TIERING_PRODUCT_SERIES : LEGACY_TIERING_PRODUCT_SERIES;
}
function workbookHeadersForVersion(version: SelectiveWorkbookVersion): readonly string[] {
  return version === "v10" || version === "v11" ? V10_REQUIREMENT_HEADERS : version === "v9" ? V9_REQUIREMENT_HEADERS : V8_REQUIREMENT_HEADERS;
}
function workbookRequirementEligible(item: Requirement, version: SelectiveWorkbookVersion): boolean { return version === "v11" || isTieringEligibleSource(item.source); }
function normalizeWorkbookRoadmaps(data: WorkspaceData, current: WorkspaceData): WorkspaceData { return current.roadmapAssignments !== undefined ? normalizeRoadmapManagement(data) : data; }
const V8_FIELD_KEYS = ["title", "description", "domainL0Id", "domainId", "scope", "allowLatestVersionUpgrade", "category", "targetMonth", "deviceWorkloadPm", "appWorkloadPm", "cloudWorkloadPm", "planningMemberId", "productManagerMemberId", "seMemberId", "algorithmMemberId"] as const;
type V8FieldKey = (typeof V8_FIELD_KEYS)[number];
const V8_FIELD_HEADERS: Record<V8FieldKey, string> = {
  title: "需求标题", description: "需求描述", domainL0Id: "领域 L0", domainId: "领域 L1", scope: "适用范围", allowLatestVersionUpgrade: "允许最近版本升级", category: "分类", targetMonth: "上线月份", deviceWorkloadPm: "设备工作量", appWorkloadPm: "App 工作量", cloudWorkloadPm: "云侧工作量", planningMemberId: "规划", productManagerMemberId: "产品经理", seMemberId: "SE", algorithmMemberId: "算法",
};
function workbookMemberLabel(member: Member, members: readonly Member[]): string {
  return members.some((other) => other.id !== member.id && other.role === member.role && other.name === member.name)
    ? `${member.name}（${member.employeeNumber}）` : member.name;
}

interface V6WorkbookMetadata {
  formatKey: typeof FORMAT_KEY;
  formatVersion: "v6" | "v7" | SelectiveWorkbookVersion;
  exportedAt: string;
  appVersion: string;
  range: HalfYearRange | null;
}

interface V6WorkbookDictionarySnapshot {
  members?: Array<{ id: string; role: string; label: string }>;
  domainL0: Array<Pick<DictionaryItem, "id" | "name">>;
  domainL1: Array<Pick<DictionaryItem, "id" | "name" | "parentId">>;
  products: Array<Pick<DictionaryItem, "id" | "name">>;
}

interface V6WorkbookRowMetadata {
  requirementId: string;
  rowToken?: string;
  baselineUpdatedAt: string;
  baselineHash: string;
  baselineVisible: string[];
  baselineFields?: Record<V8FieldKey, string>;
  baselineSeries?: Record<string, string>;
}

interface V6WorkbookPayload {
  metadata: V6WorkbookMetadata;
  dictionaries: V6WorkbookDictionarySnapshot;
  rows: V6WorkbookRowMetadata[];
  integrity?: string;
}

interface V8ParsedRow {
  row: number;
  requirementId: string;
  action: RequirementAction;
  incomingFields: Record<V8FieldKey, string>;
  baselineFields: Record<V8FieldKey, string>;
  incomingSeries: Record<string, string>;
  baselineSeries: Record<string, string>;
  baselineHash: string;
}

const LEGACY_REQUIREMENT_HEADERS = REQUIREMENT_HEADERS.filter((header) => header !== "适用范围");

const DICTIONARY_HEADERS = ["名称", "状态", "排序", "字典ID（请勿修改）", "基线校验（请勿修改）"] as const;
const DOMAIN_L1_HEADERS = [
  "名称",
  "所属领域L0",
  "状态",
  "排序",
  "字典ID（请勿修改）",
  "所属L0 ID（请勿修改）",
  "基线L0名称（请勿修改）",
  "基线校验（请勿修改）",
] as const;
const LEGACY_OVERRIDE_HEADERS = [
  "操作",
  "卡片标题",
  "详情/PPT摘要",
  "领域",
  "来源",
  "上线月份",
  "卡片键（请勿修改）",
  "基线更新时间（请勿修改）",
  "基线校验（请勿修改）",
] as const;
const V4_OVERRIDE_HEADERS = [
  "操作",
  "卡片标题",
  "详情/PPT摘要",
  "领域",
  "来源",
  "上线月份",
  "卡片顺序（请勿修改）",
  "卡片键（请勿修改）",
  "基线更新时间（请勿修改）",
  "基线校验（请勿修改）",
] as const;
const OVERRIDE_HEADERS = [
  "领域 L1 标签（自动）",
  "来源",
  "上线月份",
  "卡片顺序（请勿修改）",
  "卡片键（请勿修改）",
  "基线更新时间（请勿修改）",
  "基线校验（请勿修改）",
] as const;

type RequirementAction = "保留" | "删除";

interface ParsedDictionaryRow {
  sheet: string;
  row: number;
  item: DictionaryItem;
  baselineHash: string;
}

interface ParsedRequirementRow {
  sheet: string;
  row: number;
  action: RequirementAction;
  requirement: Requirement;
  baselineHash: string;
}

interface ParsedOverrideRow {
  sheet: string;
  row: number;
  action: RequirementAction;
  override: GroupOverride;
  baselineHash: string;
}

interface ParsedWorkspaceWorkbook {
  formatVersion: typeof FORMAT_VERSION | (typeof LEGACY_FORMAT_VERSIONS)[number];
  exportedAt: string;
  domains: ParsedDictionaryRow[];
  products: ParsedDictionaryRow[];
  requirements: ParsedRequirementRow[];
  overrides: ParsedOverrideRow[];
  errors: WorkspaceWorkbookIssue[];
  v8Rows?: V8ParsedRow[];
}

export interface PreparedWorkspaceWorkbookImport {
  parsed: ParsedWorkspaceWorkbook;
  preview: Omit<WorkspaceWorkbookImportPreview, "token" | "fileName">;
  localFingerprint?: string;
  sourceFilePath?: string;
  sourceSha256?: string;
}

interface ImportPlan {
  data: WorkspaceData;
  counts: WorkspaceWorkbookImportCounts;
  conflicts: WorkspaceWorkbookIssue[];
}

interface DictionaryMergeResult {
  items: DictionaryItem[];
  idMap: Map<string, string>;
  changed: number;
  conflicts: WorkspaceWorkbookIssue[];
}

export async function createWorkspaceWorkbook(
  data: WorkspaceData,
  appVersion: string,
  exportedAt = new Date().toISOString(),
  range?: HalfYearRange,
  formatVersion: "v11" | "v10" | "v9" | "v8" | "v7" = "v11",
): Promise<Buffer> {
  let exportedData = range ? workspaceDataForWorkbookRange(data, range.start, range.end) : data;
  if (formatVersion !== "v11") exportedData = { ...exportedData, requirements: exportedData.requirements.map((item) => ({ ...item, source: item.source ?? (data.roadmapAssignments?.find((a) => a.requirementId === item.id)?.laneId as Requirement["source"]) ?? "行业" })) };
  validateWorkspaceData(exportedData);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "解决方案需求管理";
  workbook.created = new Date(exportedAt);
  workbook.modified = new Date(exportedAt);
  workbook.calcProperties.fullCalcOnLoad = true;

  const payload = formatVersion === "v7"
    ? createRequirementSheetV6(workbook, exportedData, appVersion, exportedAt, range)
    : createRequirementSheetV8(workbook, exportedData, appVersion, exportedAt, range, formatVersion);

  const bytes = await workbook.xlsx.writeBuffer();
  return injectV6Payload(Buffer.from(bytes), payload);
}

export function workspaceDataForWorkbookRange(data: WorkspaceData, start: string, end: string): WorkspaceData {
  const months = new Set(halfYearSequence(start, end).flatMap(monthsForHalfYear));
  return {
    ...data,
    requirements: requirementsInRange(data.requirements, start, end),
    domains: data.domains,
    products: data.products,
    groupOverrides: data.groupOverrides.filter((item) => months.has(item.groupKey.split("::").at(-1) ?? "")),
    roadmapAssignments: data.roadmapAssignments?.filter((item) => requirementsInRange(data.requirements, start, end).some((req) => req.id === item.requirementId)),
    medicalRoadmapPlans: data.medicalRoadmapPlans?.filter((item) => requirementsInRange(data.requirements, start, end).some((req) => req.id === item.requirementId)),
    members: data.members,
    pptAccessPolicy: data.pptAccessPolicy,
    // Retain link targets for validation; the v6 writer still emits business columns only.
    tiering: data.tiering,
  };
}

export async function inspectWorkspaceWorkbook(
  bytes: Buffer,
  fileName: string,
  current: WorkspaceData,
): Promise<PreparedWorkspaceWorkbookImport> {
  if (!fileName.toLowerCase().endsWith(".xlsx")) throw new Error("请选择 .xlsx 协作工作簿");
  if (!bytes.length || bytes.length > MAX_WORKBOOK_BYTES) throw new Error("Excel 文件为空或超过 25MB 限制");
  const storedPayload = await readV6Payload(bytes);
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(bytes as never);
  } catch {
    throw new Error("Excel 文件损坏或不是有效的 .xlsx 工作簿");
  }

  const errors: WorkspaceWorkbookIssue[] = [];
  const requirementSheet = workbook.getWorksheet(SHEETS.requirements);
  const recoverableVersion = requirementSheet && (["v11", "v10", "v9", "v8"] as const).find((version) =>
    workbookHeadersForVersion(version).every((header, index) => cellText(requirementSheet.getCell(1, index + 1)) === header));
  const v6Payload = storedPayload ?? (requirementSheet && recoverableVersion
    ? recoverOperationOnlyV8Payload(requirementSheet, current, recoverableVersion) : undefined);
  let parsed: ParsedWorkspaceWorkbook;
  if (v6Payload) {
    if (workbook.worksheets.length !== 1) addIssue(errors, SHEETS.requirements, "A1", "协作 Excel 只能包含“需求清单”一个工作表");
    if (!requirementSheet) throw new Error("协作 Excel 缺少“需求清单”工作表");
    if (isSelectiveWorkbookVersion(v6Payload.metadata.formatVersion)) {
      const v8Rows = parseRequirementSheetV8(requirementSheet, current, v6Payload, errors);
      parsed = { formatVersion: v6Payload.metadata.formatVersion, exportedAt: v6Payload.metadata.exportedAt, domains: [], products: [], requirements: [], overrides: [], errors, v8Rows };
      const fieldChanges = createV8FieldChanges(current, v8Rows, v6Payload.metadata.formatVersion);
      const counts = emptyImportCounts();
      counts.updated = new Set(fieldChanges.filter((change) => change.kind !== "delete").map((change) => change.requirementId)).size;
      counts.deleted = fieldChanges.filter((change) => change.kind === "delete").length;
      counts.unchanged = v8Rows.length - counts.updated - counts.deleted;
      counts.conflicts = fieldChanges.filter((change) => change.conflict).length;
      return {
        parsed,
        localFingerprint: v8WorkspaceFingerprint(current),
        preview: {
          formatVersion: parsed.formatVersion, exportedAt: parsed.exportedAt, counts, errors,
          conflicts: fieldChanges.filter((change) => change.conflict).map((change) => conflictIssue(SHEETS.requirements, change.row, `${change.title} · ${change.field} 本机和 Excel 均已修改`)),
          differences: [], fieldChanges, differenceDigest: stableHash(fieldChanges),
        },
      };
    }
    const requirements = parseRequirementSheetV6(requirementSheet, current, v6Payload, errors);
    parsed = {
      formatVersion: v6Payload.metadata.formatVersion,
      exportedAt: v6Payload.metadata.exportedAt,
      domains: [],
      products: [],
      requirements,
      overrides: [],
      errors,
    };
  } else {
    if (requirementSheet && (V6_REQUIREMENT_HEADERS.every((header, index) => cellText(requirementSheet.getCell(1, index + 1)) === header)
      || V8_REQUIREMENT_HEADERS.every((header, index) => cellText(requirementSheet.getCell(1, index + 1)) === header)
      || V9_REQUIREMENT_HEADERS.every((header, index) => cellText(requirementSheet.getCell(1, index + 1)) === header)
      || V10_REQUIREMENT_HEADERS.every((header, index) => cellText(requirementSheet.getCell(1, index + 1)) === header))) {
      throw new Error("协作 Excel 的内部系统信息缺失或损坏，请重新导出后再导入");
    }
    const guide = workbook.getWorksheet(SHEETS.guide);
    if (!guide) throw new Error("不支持的协作 Excel 格式或版本");
    const formatVersion = cellText(guide.getCell("B14"));
    if (cellText(guide.getCell("B13")) !== FORMAT_KEY || !SUPPORTED_FORMAT_VERSIONS.includes(formatVersion) || formatVersion === FORMAT_VERSION) {
      throw new Error("不支持的协作 Excel 格式或版本");
    }
    const exportedAt = cellText(guide.getCell("B15"));
    if (!validTimestamp(exportedAt)) throw new Error("协作 Excel 缺少有效的导出时间");
    const domainL0Sheet = requireSheet(workbook, SHEETS.domainL0s, errors);
    const domainL1Sheet = requireSheet(workbook, SHEETS.domainL1s, errors);
    const productSheet = requireSheet(workbook, SHEETS.products, errors);
    const legacyRequirementSheet = requireSheet(workbook, SHEETS.requirements, errors);
    const overrideSheet = requireSheet(workbook, SHEETS.overrides, errors);
    const domainL0s = domainL0Sheet ? parseDictionarySheet(domainL0Sheet, errors, "L0") : [];
    const domainL1s = domainL1Sheet ? parseDomainL1Sheet(domainL1Sheet, domainL0s, errors) : [];
    const domains = [...domainL0s, ...domainL1s];
    const products = productSheet ? parseDictionarySheet(productSheet, errors) : [];
    validateDictionaryRows(domains, "领域", errors);
    validateDictionaryRows(products, "产品", errors);
    const parsedFormatVersion = formatVersion as ParsedWorkspaceWorkbook["formatVersion"];
    const requirements = legacyRequirementSheet ? parseRequirementSheet(legacyRequirementSheet, domainL0s, domainL1s, products, exportedAt, parsedFormatVersion, errors) : [];
    const overrides = overrideSheet ? parseOverrideSheet(overrideSheet, domainL1s, exportedAt, parsedFormatVersion, errors) : [];
    parsed = { formatVersion: parsedFormatVersion, exportedAt, domains, products, requirements, overrides, errors };
  }
  if (current.roadmapAssignments !== undefined) for (const row of parsed.requirements) {
    const existing = current.requirements.find((item) => item.id === row.requirement.id);
    if (existing && row.action !== "删除") row.requirement = { ...row.requirement, source: existing.source, medicalSubmissionMilestones: existing.medicalSubmissionMilestones, medicalRoadmapSortOrder: existing.medicalRoadmapSortOrder };
  }
  const plan = planWorkspaceWorkbookImport(current, parsed, "local-wins", new Date().toISOString());
  const excelPlan = planWorkspaceWorkbookImport(current, parsed, "excel-wins", new Date().toISOString());
  const differences = diffWorkspaceData(current, excelPlan.data);
  try {
    validateWorkspaceData(plan.data);
  } catch (error) {
    addIssue(errors, SHEETS.requirements, "A1", error instanceof Error ? error.message : String(error));
  }

  return {
    parsed,
    preview: {
      formatVersion: parsed.formatVersion,
      warnings: current.roadmapAssignments !== undefined ? ["旧版 Excel 中的来源修改将忽略；路标匹配和医疗节点只在路标管理中调整。"] : [],
      exportedAt: parsed.exportedAt,
      counts: plan.counts,
      errors,
      conflicts: plan.conflicts,
      differences,
      differenceDigest: workspaceDifferenceDigest(differences),
    },
  };
}

export function applyPreparedWorkspaceWorkbook(
  current: WorkspaceData,
  prepared: PreparedWorkspaceWorkbookImport,
  conflictMode: WorkspaceWorkbookConflictMode,
  now = new Date().toISOString(),
  fieldSelection?: WorkspaceWorkbookFieldSelection,
): WorkspaceData {
  if (prepared.parsed.errors.length) throw new Error("协作 Excel 存在校验错误，请修正后重新导入");
  if (isSelectiveWorkbookVersion(prepared.parsed.formatVersion)) return applyV8FieldSelection(current, prepared, fieldSelection, now);
  const plan = planWorkspaceWorkbookImport(current, prepared.parsed, conflictMode, now);
  validateWorkspaceData(plan.data);
  return plan.data;
}

function applyV8FieldSelection(current: WorkspaceData, prepared: PreparedWorkspaceWorkbookImport, selection: WorkspaceWorkbookFieldSelection | undefined, now: string): WorkspaceData {
  if (!selection || !Array.isArray(selection.selectedChangeIds) || !selection.conflictDecisions || typeof selection.conflictDecisions !== "object") throw new Error("请选择要合入的逐项差异");
  if (v8WorkspaceFingerprint(current) !== prepared.localFingerprint) throw new Error("本机需求或分层状态在检查后已变化，请重新选择 Excel 检查");
  const rows = prepared.parsed.v8Rows ?? [];
  const changes = createV8FieldChanges(current, rows, prepared.parsed.formatVersion as SelectiveWorkbookVersion);
  if (stableHash(changes) !== prepared.preview.differenceDigest) throw new Error("Excel 差异已变化，请重新检查文件");
  const byChangeId = new Map(changes.map((item) => [item.id, item]));
  const submitted = new Set(selection.selectedChangeIds);
  if (submitted.size !== selection.selectedChangeIds.length || [...submitted].some((id) => !byChangeId.has(id))) throw new Error("勾选项不属于当前 Excel 差异，请重新检查");
  const selected = new Set([...submitted].filter((id) => byChangeId.get(id)!.kind === "delete" || !submitted.has(`${byChangeId.get(id)!.requirementId}:delete`)));
  for (const id of selected) {
    const change = byChangeId.get(id)!;
    if (change.conflict && !["local", "excel"].includes(selection.conflictDecisions[id])) throw new Error(`请为“${change.title} · ${change.field}”选择保留本机或采用 Excel`);
  }
  const adopted = new Set([...selected].filter((id) => !byChangeId.get(id)!.conflict || selection.conflictDecisions[id] === "excel"));
  if (!adopted.size) return current;
  const deleted = new Set(rows.filter((row) => adopted.has(`${row.requirementId}:delete`)).map((row) => row.requirementId));
  const nextRequirements = current.requirements.filter((item) => !deleted.has(item.id)).map((item) => {
    const row = rows.find((candidate) => candidate.requirementId === item.id);
    if (!row || row.action === "删除") return item;
    const patchKeys = V8_FIELD_KEYS.filter((key) => adopted.has(`${item.id}:field:${key}`));
    if (!patchKeys.length) return item;
    const next = { ...item };
    for (const key of patchKeys) {
      const value = row.incomingFields[key];
      if (key === "allowLatestVersionUpgrade") next.allowLatestVersionUpgrade = value === "是";
      else if (key === "deviceWorkloadPm" || key === "appWorkloadPm" || key === "cloudWorkloadPm") next[key] = Number(value);
      else if (key === "scope") next.scope = value as Requirement["scope"];
      else if (key === "category") next.category = value as Requirement["category"];
      else next[key] = value;
    }
    if (patchKeys.includes("targetMonth") && next.historicalLaunch && next.targetMonth) next.historicalLaunch = { ...next.historicalLaunch, suggestedTargetMonth: next.targetMonth };
    next.workloadPm = roundWorkload(next.deviceWorkloadPm + next.appWorkloadPm + next.cloudWorkloadPm + (next.researchAppWorkloadPm ?? 0) + (next.algorithmWorkloadPm ?? 0) + next.unallocatedWorkloadPm);
    next.updatedAt = now;
    return next;
  });
  const domainsById = new Map(current.domains.map((item) => [item.id, item]));
  for (const item of nextRequirements) {
    const l1 = domainsById.get(item.domainId);
    if (l1?.parentId !== item.domainL0Id) throw new Error(`所选差异组合无效，未写入任何数据：需求“${item.title}”的领域 L1 不属于所选领域 L0`);
  }
  const scoped = normalizeWorkspaceTieringScope({ ...current, requirements: nextRequirements });
  const tiering = normalizeTieringData(scoped.tiering);
  const featureByRequirement = new Map(tiering.features.filter((item) => item.requirementId).map((item) => [item.requirementId!, item]));
  const records = new Map(tiering.supportRecords.map((item) => [item.id, item]));
  for (const row of rows) {
    if (deleted.has(row.requirementId)) continue;
    const feature = featureByRequirement.get(row.requirementId);
    if (!feature) continue;
    for (const series of workbookSeriesForVersion(prepared.parsed.formatVersion as SelectiveWorkbookVersion)) {
      if (!adopted.has(`${row.requirementId}:series:${series.id}`)) continue;
      const id = tieringSupportId(feature.id, series.id);
      const previous = records.get(id);
      if (!previous) throw new Error(`需求“${row.requirementId}”缺少${series.name}支持记录，请先修复分层分级数据`);
      const record = { ...previous, status: row.incomingSeries[series.id] as typeof previous.status, updatedAt: now };
      records.set(id, record);
      tiering.changeLogs.push({
        id: randomUUID(), changedAt: now, triggerType: "协作 Excel", featureId: feature.id, seriesId: series.id,
        before: JSON.stringify(previous), after: JSON.stringify(record), reason: `协作 Excel ${prepared.parsed.formatVersion} 勾选合入系列支持状态`,
        sourceIds: [...record.sourceIds], reviewer: "", recordStatus: "有效", note: "", updatedAt: now,
      });
    }
  }
  tiering.supportRecords = [...records.values()];
  const result: WorkspaceData = normalizeWorkbookRoadmaps({ ...scoped, tiering }, current);
  try { validateWorkspaceData(result); }
  catch (error) { throw new Error(`所选差异组合无效，未写入任何数据：${error instanceof Error ? error.message : String(error)}`); }
  return result;
}

function v8RequirementFields(item: Requirement): Record<V8FieldKey, string> {
  return {
    title: item.title.trim(), description: item.description.trim(), domainL0Id: item.domainL0Id,
    domainId: item.domainId, scope: item.scope, allowLatestVersionUpgrade: item.allowLatestVersionUpgrade ? "是" : "否",
    category: item.category, targetMonth: item.targetMonth,
    deviceWorkloadPm: String(roundWorkload(item.deviceWorkloadPm)),
    appWorkloadPm: String(roundWorkload(item.appWorkloadPm)),
    cloudWorkloadPm: String(roundWorkload(item.cloudWorkloadPm)),
    planningMemberId: item.planningMemberId ?? "", productManagerMemberId: item.productManagerMemberId ?? "",
    seMemberId: item.seMemberId ?? "", algorithmMemberId: item.algorithmMemberId ?? "",
  };
}

function emptyImportCounts(): WorkspaceWorkbookImportCounts {
  return { added: 0, updated: 0, deleted: 0, unchanged: 0, conflicts: 0, domainsChanged: 0, productsChanged: 0, groupOverridesChanged: 0 };
}

function v8WorkspaceFingerprint(data: WorkspaceData): string {
  return stableHash({
    requirements: data.requirements.map((item) => [item.id, item.updatedAt, v8RequirementFields(item), item.source, item.productIds]),
    domains: data.domains.map((item) => [item.id, item.name, item.parentId]),
    members: (data.members ?? []).map((item) => [item.id, item.name, item.role]),
    tiering: data.tiering?.supportRecords.map((item) => [item.id, item.status, item.updatedAt]),
  });
}

function currentSeriesStatus(data: WorkspaceData, requirementId: string, seriesId: string): string {
  const feature = data.tiering?.features.find((item) => item.requirementId === requirementId);
  return feature ? data.tiering?.supportRecords.find((item) => item.id === tieringSupportId(feature.id, seriesId))?.status ?? "未判断" : "未判断";
}

function v8DisplayValue(data: WorkspaceData, key: V8FieldKey, value: string): string {
  if (key === "domainL0Id" || key === "domainId") return data.domains.find((item) => item.id === value)?.name ?? value;
  if (key === "planningMemberId" || key === "productManagerMemberId" || key === "seMemberId" || key === "algorithmMemberId") return data.members?.find((item) => item.id === value)?.name ?? value;
  return value;
}

function createV8FieldChanges(current: WorkspaceData, rows: readonly V8ParsedRow[], formatVersion: SelectiveWorkbookVersion): WorkspaceWorkbookFieldChange[] {
  const byId = new Map(current.requirements.map((item) => [item.id, item]));
  const changes: WorkspaceWorkbookFieldChange[] = [];
  for (const row of rows) {
    const existing = byId.get(row.requirementId);
    if (!existing) continue;
    if (row.action === "删除") {
      const conflict = !requirementBaselineMatches(existing, row.baselineHash, formatVersion) || (isTieringEligibleRequirement(existing) && (
        workbookSeriesForVersion(formatVersion).some((series) => currentSeriesStatus(current, existing.id, series.id) !== (row.baselineSeries[series.id] || "未判断"))
        || BUILT_IN_TIERING_PRODUCT_SERIES.some((series) => !(series.id in row.baselineSeries) && currentSeriesStatus(current, existing.id, series.id) !== "未判断")
      ));
      changes.push({ id: `${existing.id}:delete`, requirementId: existing.id, row: row.row, title: existing.title, kind: "delete", field: "删除整条需求", baselineValue: row.baselineFields.title, currentValue: existing.title, excelValue: "删除", conflict });
      continue;
    }
    const localFields = v8RequirementFields(existing);
    for (const key of V8_FIELD_KEYS) {
      const incoming = row.incomingFields[key];
      if (incoming === row.baselineFields[key] || incoming === localFields[key]) continue;
      changes.push({ id: `${existing.id}:field:${key}`, requirementId: existing.id, row: row.row, title: existing.title, kind: "field", field: V8_FIELD_HEADERS[key], baselineValue: v8DisplayValue(current, key, row.baselineFields[key]), currentValue: v8DisplayValue(current, key, localFields[key]), excelValue: v8DisplayValue(current, key, incoming), conflict: localFields[key] !== row.baselineFields[key] });
    }
    if (!isTieringEligibleRequirement(existing)) continue;
    for (const series of workbookSeriesForVersion(formatVersion)) {
      const incoming = row.incomingSeries[series.id];
      const local = currentSeriesStatus(current, existing.id, series.id);
      if (incoming === row.baselineSeries[series.id] || incoming === local) continue;
      changes.push({ id: `${existing.id}:series:${series.id}`, requirementId: existing.id, row: row.row, title: existing.title, kind: "series", field: series.name, baselineValue: row.baselineSeries[series.id], currentValue: local, excelValue: incoming, conflict: local !== row.baselineSeries[series.id] });
    }
  }
  return changes;
}

function createRequirementSheetV8(
  workbook: ExcelJS.Workbook, data: WorkspaceData, appVersion: string, exportedAt: string, range?: HalfYearRange,
  formatVersion: SelectiveWorkbookVersion = "v8",
): V6WorkbookPayload {
  const sheet = workbook.addWorksheet(SHEETS.requirements, { properties: { tabColor: { argb: "FF228A70" } } });
  const offset = formatVersion === "v8" ? 0 : 1;
  const headers = workbookHeadersForVersion(formatVersion);
  const seriesForWorkbook = workbookSeriesForVersion(formatVersion);
  sheet.views = [{ state: "frozen", xSplit: 2 + offset, ySplit: 1, showGridLines: false }];
  const domains = new Map(data.domains.map((item) => [item.id, item.name]));
  const tiering = normalizeWorkspaceTieringScope(data).tiering!;
  const featureByRequirement = new Map(tiering.features.filter((feature) => feature.requirementId).map((feature) => [feature.requirementId!, feature]));
  const supportById = new Map(tiering.supportRecords.map((record) => [record.id, record]));
  const rows = data.requirements.map((item, index) => {
    const fields = v8RequirementFields(item);
    const values: ExcelJS.CellValue[] = [
      "保留", fields.title, fields.description || null, domains.get(fields.domainL0Id) ?? fields.domainL0Id,
      domains.get(fields.domainId) ?? fields.domainId, fields.scope, fields.allowLatestVersionUpgrade,
      fields.category, monthDate(fields.targetMonth), item.deviceWorkloadPm, item.appWorkloadPm,
      item.cloudWorkloadPm,
      ...WORKBOOK_OWNER_FIELDS.map(({ key }) => {
        const member = data.members?.find((candidate) => candidate.id === item[key]);
        return member ? workbookMemberLabel(member, data.members ?? []) : null;
      }),
    ];
    const feature = featureByRequirement.get(item.id);
    for (const series of seriesForWorkbook) {
      values.push(workbookRequirementEligible(item, formatVersion) && feature
        ? supportById.get(tieringSupportId(feature.id, series.id))?.status ?? "未判断" : null);
    }
    return offset ? [index + 1, ...values] : values;
  });
  const rowMetadata: V6WorkbookRowMetadata[] = data.requirements.map((item, index) => {
    const feature = featureByRequirement.get(item.id);
    return {
      requirementId: item.id, rowToken: workbookRowToken(item), baselineUpdatedAt: item.updatedAt, baselineHash: requirementHashV7(item),
      baselineVisible: canonicalVisibleRowV8(offset ? rows[index].slice(1) : rows[index]), baselineFields: v8RequirementFields(item),
      baselineSeries: Object.fromEntries(seriesForWorkbook.map((series) => [series.id,
        workbookRequirementEligible(item, formatVersion) && feature
          ? supportById.get(tieringSupportId(feature.id, series.id))?.status ?? "未判断" : ""])),
    };
  });
  sheet.addTable({
    name: "RequirementsTable", ref: "A1", headerRow: true, totalsRow: false,
    style: { theme: "TableStyleMedium2", showRowStripes: false },
    columns: headers.map((name) => ({ name, filterButton: true })), rows,
  });
  sheet.columns = [...(offset ? [8] : []), ...[10, 28, 48, 18, 18, 16, 21, 16, 16, 16, 16, 16, 20, 20, 20, 20], ...seriesForWorkbook.map(() => 19)].map((width) => ({ width }));
  sheet.getColumn(9 + offset).numFmt = "yyyy-mm";
  for (const column of [10, 11, 12]) sheet.getColumn(column + offset).numFmt = "0.00";
  sheet.getRow(1).height = 31;
  forEachCell(sheet, `A1:${columnLetter(headers.length)}1`, (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9F2E6" } };
    cell.font = { name: "Microsoft YaHei", size: 10, bold: true, color: { argb: "FF000000" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = lightBorders();
  });
  const lastRow = rows.length + 1;
  if (rows.length) {
    forEachCell(sheet, `A2:${columnLetter(headers.length)}${lastRow}`, (cell) => {
      cell.fill = { type: "pattern", pattern: "none" };
      cell.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF000000" } };
      cell.alignment = { vertical: "top", wrapText: cell.col === "C" };
      cell.border = lightBorders();
    });
    for (let row = 2; row <= lastRow; row += 1) sheet.getRow(row).height = 28;
    rowMetadata.forEach((metadata, index) => applyWorkbookRowIdentityFormat(sheet, index + 2, metadata.rowToken!, offset));
    const rangeFor = (column: number) => `${columnLetter(column + offset)}2:${columnLetter(column + offset)}${lastRow}`;
    applyListValidation(sheet, rangeFor(1), ["保留", "删除"]);
    applyListValidation(sheet, rangeFor(6), [...REQUIREMENT_SCOPES]);
    applyListValidation(sheet, rangeFor(7), ["是", "否"]);
    applyListValidation(sheet, rangeFor(8), [...CATEGORIES]);
    applyListValidation(sheet, rangeFor(4), data.domains.filter((item) => item.level === "L0").map((item) => item.name));
    applyListValidation(sheet, rangeFor(5), data.domains.filter((item) => (item.level ?? "L1") === "L1").map((item) => item.name));
    for (const column of [10, 11, 12]) applyDecimalValidation(sheet, rangeFor(column), 0, true);
    WORKBOOK_OWNER_FIELDS.forEach(({ role }, index) => {
      const labels = (data.members ?? []).filter((member) => member.role === role).map((member) => workbookMemberLabel(member, data.members ?? []));
      applyListValidation(sheet, rangeFor(13 + index), labels, true);
    });
    data.requirements.forEach((item, index) => {
      if (!workbookRequirementEligible(item, formatVersion)) return;
      for (let column = 17; column < 17 + seriesForWorkbook.length; column += 1) applyListValidation(sheet, `${columnLetter(column + offset)}${index + 2}:${columnLetter(column + offset)}${index + 2}`, ["支持", "不支持", "未判断"]);
    });
  }
  const payload: V6WorkbookPayload = {
    metadata: { formatKey: FORMAT_KEY, formatVersion, exportedAt, appVersion, range: range ?? null },
    dictionaries: {
      members: (data.members ?? []).map((member) => ({ id: member.id, role: member.role ?? "", label: workbookMemberLabel(member, data.members ?? []) })),
      domainL0: data.domains.filter((item) => item.level === "L0").map(({ id, name }) => ({ id, name })),
      domainL1: data.domains.filter((item) => (item.level ?? "L1") === "L1").map(({ id, name, parentId }) => ({ id, name, parentId })),
      products: [],
    },
    rows: rowMetadata,
  };
  payload.integrity = stableHash({ metadata: payload.metadata, dictionaries: payload.dictionaries, rows: payload.rows });
  return payload;
}

function canonicalVisibleRowV8(values: readonly ExcelJS.CellValue[]): string[] {
  return values.map((value, index) => {
    if (index === 8) return parseMonth(value) ?? "";
    if (index >= 9 && index <= 11) return String(roundWorkload(Number(value ?? 0)));
    return String(value ?? "").trim();
  });
}

const WORKBOOK_ROW_TOKEN_PREFIX = "__HRB_";
const WORKBOOK_ROW_TOKEN_COLUMNS = [1, 2, 3, 4, 5, 6, 7, 8, ...Array.from({ length: 17 }, (_, index) => index + 13)] as const;

function workbookRowToken(item: Requirement): string {
  return stableHash({ kind: "requirement-row", id: item.id, createdAt: item.createdAt }).slice(0, 24);
}

function applyWorkbookRowIdentityFormat(sheet: ExcelJS.Worksheet, row: number, token: string, offset = 0): void {
  // Spread the 96-bit token across text cells using only 16 reusable number
  // formats. This avoids Excel's custom-number-format count limit while keeping
  // the identifier outside user-visible values and columns.
  [...token].forEach((nibble, index) => {
    sheet.getCell(row, WORKBOOK_ROW_TOKEN_COLUMNS[index] + offset).numFmt = `0"${WORKBOOK_ROW_TOKEN_PREFIX}${nibble}";0;0;@`;
  });
}

function rowTokenFromWorksheetRow(sheet: ExcelJS.Worksheet, row: number, offset = 0): string | undefined {
  const nibblePattern = new RegExp(`^0"${WORKBOOK_ROW_TOKEN_PREFIX}([a-f0-9])";0;0;@$`, "i");
  const token = WORKBOOK_ROW_TOKEN_COLUMNS.slice(0, 24).map((column) => nibblePattern.exec(sheet.getCell(row, column + offset).numFmt ?? "")?.[1] ?? "").join("");
  return token.length === 24 ? token.toLowerCase() : undefined;
}

function createRequirementSheetV6(
  workbook: ExcelJS.Workbook,
  data: WorkspaceData,
  appVersion: string,
  exportedAt: string,
  range?: HalfYearRange,
): V6WorkbookPayload {
  const sheet = workbook.addWorksheet(SHEETS.requirements, { properties: { tabColor: { argb: "FFED6A5A" } } });
  sheet.views = [{ state: "frozen", xSplit: 2, ySplit: V6_HEADER_ROW, showGridLines: false }];
  const domains = new Map(data.domains.map((item) => [item.id, item.name]));
  const products = new Map(data.products.map((item) => [item.id, item.name]));
  const rows = data.requirements.map((item) => [
    "保留",
    item.title,
    item.description || null,
    domains.get(item.domainL0Id) ?? item.domainL0Id,
    domains.get(item.domainId) ?? item.domainId,
    item.source,
    item.scope,
    item.allowLatestVersionUpgrade ? "是" : "否",
    item.category,
    monthDate(item.targetMonth),
    item.productIds.map((id) => products.get(id) ?? id)[0] ?? null,
    item.deviceWorkloadPm,
    item.appWorkloadPm,
    item.cloudWorkloadPm,
    ...WORKBOOK_OWNER_FIELDS.map(({ key }) => {
      const member = data.members?.find((member) => member.id === item[key]);
      return member ? workbookMemberLabel(member, data.members ?? []) : null;
    }),
  ] as ExcelJS.CellValue[]);

  const rowMetadata: V6WorkbookRowMetadata[] = data.requirements.map((item, index) => ({
    requirementId: item.id,
    baselineUpdatedAt: item.updatedAt,
    baselineHash: requirementHashV7(item),
    baselineVisible: canonicalVisibleRow(rows[index]),
  }));

  sheet.addTable({
    name: "RequirementsTable",
    ref: `A${V6_HEADER_ROW}`,
    headerRow: true,
    totalsRow: false,
    style: { theme: "TableStyleMedium2", showRowStripes: false },
    columns: V7_REQUIREMENT_HEADERS.map((header) => ({ name: header, filterButton: true })),
    rows,
  });

  const widths = [10, 28, 48, 18, 18, 18, 16, 21, 16, 16, 22, 16, 16, 16, 20, 20, 20, 20];
  sheet.columns = widths.map((width) => ({ width }));
  sheet.getColumn(10).numFmt = "yyyy-mm";
  for (const column of [12, 13, 14]) sheet.getColumn(column).numFmt = "0.00";

  sheet.getRow(V6_HEADER_ROW).height = 30;
  forEachCell(sheet, `A${V6_HEADER_ROW}:R${V6_HEADER_ROW}`, (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9F2E6" } };
    cell.font = { name: "Microsoft YaHei", size: 10, bold: true, color: { argb: "FF17352A" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = lightBorders();
  });
  const lastDataRow = V6_FIRST_DATA_ROW + rows.length - 1;
  if (rows.length) {
    forEachCell(sheet, `A${V6_FIRST_DATA_ROW}:R${lastDataRow}`, (cell) => {
      cell.fill = { type: "pattern", pattern: "none" };
      cell.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF000000" } };
      cell.alignment = { vertical: "top", wrapText: cell.col === "C" };
      cell.border = lightBorders();
    });
    for (let row = V6_FIRST_DATA_ROW; row <= lastDataRow; row += 1) sheet.getRow(row).height = 28;
  }

  const metadata: V6WorkbookMetadata = {
    formatKey: FORMAT_KEY,
    formatVersion: "v7",
    exportedAt,
    appVersion,
    range: range ?? null,
  };

  if (rows.length) {
    WORKBOOK_OWNER_FIELDS.forEach(({ role }, index) => {
      const labels = (data.members ?? []).filter((member) => member.role === role).map((member) => workbookMemberLabel(member, data.members ?? []));
      if (labels.join(",").length > 253 || labels.some((label) => /[,"\r\n]/.test(label))) {
        throw new Error(`${role}成员下拉超出单 Sheet 内嵌列表容量，或姓名包含英文逗号、引号、换行。请缩短成员显示名称后重试；未生成不完整表格。`);
      }
      applyListValidation(sheet, `${columnLetter(15 + index)}${V6_FIRST_DATA_ROW}:${columnLetter(15 + index)}${lastDataRow}`, labels, true);
    });
    applyListValidation(sheet, `A${V6_FIRST_DATA_ROW}:A${lastDataRow}`, ["保留", "删除"]);
    applyListValidation(sheet, `F${V6_FIRST_DATA_ROW}:F${lastDataRow}`, [...SOURCES]);
    applyListValidation(sheet, `G${V6_FIRST_DATA_ROW}:G${lastDataRow}`, [...REQUIREMENT_SCOPES]);
    applyListValidation(sheet, `H${V6_FIRST_DATA_ROW}:H${lastDataRow}`, ["是", "否"]);
    applyListValidation(sheet, `I${V6_FIRST_DATA_ROW}:I${lastDataRow}`, [...CATEGORIES]);
    applyListValidation(sheet, `D${V6_FIRST_DATA_ROW}:D${lastDataRow}`, data.domains.filter((item) => item.level === "L0").map((item) => item.name));
    applyListValidation(sheet, `E${V6_FIRST_DATA_ROW}:E${lastDataRow}`, data.domains.filter((item) => (item.level ?? "L1") === "L1").map((item) => item.name));
    applyListValidation(sheet, `K${V6_FIRST_DATA_ROW}:K${lastDataRow}`, data.products.map((item) => item.name));
    for (const column of [12, 13, 14]) applyDecimalValidation(sheet, `${columnLetter(column)}${V6_FIRST_DATA_ROW}:${columnLetter(column)}${lastDataRow}`, 0, true);
  }
  return {
    metadata,
    dictionaries: {
      members: (data.members ?? []).map((member) => ({ id: member.id, role: member.role ?? "", label: workbookMemberLabel(member, data.members ?? []) })),
      domainL0: data.domains.filter((item) => item.level === "L0").map(({ id, name }) => ({ id, name })),
      domainL1: data.domains.filter((item) => (item.level ?? "L1") === "L1").map(({ id, name, parentId }) => ({ id, name, parentId })),
      products: data.products.map(({ id, name }) => ({ id, name })),
    },
    rows: rowMetadata,
  };
}

async function injectV6Payload(bytes: Buffer, payload: V6WorkbookPayload): Promise<Buffer> {
  const zip = await JSZip.loadAsync(bytes);
  const json = JSON.stringify(payload).replaceAll("]]>", "]]]]><![CDATA[>");
  zip.file(V6_CUSTOM_XML_PATH, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><roadmapCollaboration xmlns="${V6_CUSTOM_XML_NAMESPACE}"><payload><![CDATA[${json}]]></payload></roadmapCollaboration>`);
  // WPS and some Excel save paths discard unknown customXml parts. Custom file
  // properties are an Office-supported, non-worksheet location that survives
  // those saves; retain customXml as a reader-compatible copy for older files.
  if (isSelectiveWorkbookVersion(payload.metadata.formatVersion)) {
    const encoded = deflateRawSync(Buffer.from(JSON.stringify(payload), "utf8")).toString("base64");
    const chunks = encoded.match(new RegExp(`.{1,${WORKBOOK_CUSTOM_PROPERTIES_CHUNK_SIZE}}`, "g")) ?? [];
    const properties = [
      `${chunks.length}:${createHash("sha256").update(encoded).digest("hex")}`,
      ...chunks,
    ].map((value, index) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${index + 2}" name="${WORKBOOK_CUSTOM_PROPERTIES_PREFIX}${index === 0 ? "Manifest" : String(index - 1).padStart(3, "0")}"><vt:lpwstr>${value}</vt:lpwstr></property>`);
    zip.file(WORKBOOK_CUSTOM_PROPERTIES_PATH, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${properties.join("")}</Properties>`);
  }
  const relsPath = "_rels/.rels";
  const relsFile = zip.file(relsPath);
  if (!relsFile) throw new Error("Excel 文件缺少关系清单");
  let rels = await relsFile.async("string");
  if (!rels.includes(`Target="${V6_CUSTOM_XML_PATH}"`)) {
    const relationship = `<Relationship Id="${V6_CUSTOM_XML_REL_ID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="${V6_CUSTOM_XML_PATH}"/>`;
    rels = rels.replace("</Relationships>", `${relationship}</Relationships>`);
  }
  if (isSelectiveWorkbookVersion(payload.metadata.formatVersion) && !rels.includes(`Target="${WORKBOOK_CUSTOM_PROPERTIES_PATH}"`)) {
    const relationship = `<Relationship Id="rIdRoadmapCollaborationProperties" Type="${WORKBOOK_CUSTOM_PROPERTIES_REL}" Target="${WORKBOOK_CUSTOM_PROPERTIES_PATH}"/>`;
    rels = rels.replace("</Relationships>", `${relationship}</Relationships>`);
  }
  zip.file(relsPath, rels);
  if (isSelectiveWorkbookVersion(payload.metadata.formatVersion)) {
    const contentTypesFile = zip.file("[Content_Types].xml");
    if (!contentTypesFile) throw new Error("Excel 文件缺少内容类型清单");
    let contentTypes = await contentTypesFile.async("string");
    if (!contentTypes.includes(`PartName="/${WORKBOOK_CUSTOM_PROPERTIES_PATH}"`)) {
      const override = `<Override PartName="/${WORKBOOK_CUSTOM_PROPERTIES_PATH}" ContentType="${WORKBOOK_CUSTOM_PROPERTIES_TYPE}"/>`;
      contentTypes = contentTypes.replace("</Types>", `${override}</Types>`);
      zip.file("[Content_Types].xml", contentTypes);
    }
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

async function readV6Payload(bytes: Buffer): Promise<V6WorkbookPayload | undefined> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return undefined;
  }
  try {
    const customPropertiesFile = zip.file(WORKBOOK_CUSTOM_PROPERTIES_PATH);
    let propertiesPayload: V6WorkbookPayload | undefined;
    if (customPropertiesFile) {
      const xml = await customPropertiesFile.async("string");
      const properties = new Map<string, string>();
      for (const match of xml.matchAll(/<property\b([^>]*)>([\s\S]*?)<\/property>/g)) {
        const name = /\bname="([^"]+)"/.exec(match[1])?.[1];
        const value = /<vt:lpwstr>([\s\S]*?)<\/vt:lpwstr>/.exec(match[2])?.[1];
        if (name?.startsWith(WORKBOOK_CUSTOM_PROPERTIES_PREFIX)) {
          if (value === undefined || properties.has(name)) throw new Error("custom-property");
          properties.set(name, value);
        }
      }
      if (properties.size) {
        const manifest = properties.get(`${WORKBOOK_CUSTOM_PROPERTIES_PREFIX}Manifest`);
        const [countText, digest] = manifest?.split(":") ?? [];
        const count = Number(countText);
        if (!Number.isInteger(count) || count < 1 || count > 1000 || !/^[a-f0-9]{64}$/i.test(digest ?? "") || properties.size !== count + 1) throw new Error("custom-property-manifest");
        const chunks = Array.from({ length: count }, (_, index) => properties.get(`${WORKBOOK_CUSTOM_PROPERTIES_PREFIX}${String(index).padStart(3, "0")}`));
        if (chunks.some((chunk) => !chunk || !/^[A-Za-z0-9+/=]+$/.test(chunk))) throw new Error("custom-property-chunk");
        const encoded = chunks.join("");
        if (createHash("sha256").update(encoded).digest("hex") !== digest.toLowerCase()) throw new Error("custom-property-digest");
        propertiesPayload = JSON.parse(inflateRawSync(Buffer.from(encoded, "base64"), { maxOutputLength: MAX_WORKBOOK_BYTES }).toString("utf8")) as V6WorkbookPayload;
      }
    }
    const file = zip.file(V6_CUSTOM_XML_PATH);
    let xmlPayload: V6WorkbookPayload | undefined;
    if (file) {
      const xml = await file.async("string");
      if (!xml.includes(V6_CUSTOM_XML_NAMESPACE)) throw new Error("namespace");
      const match = /<payload><!\[CDATA\[([\s\S]*?)\]\]><\/payload>/.exec(xml);
      if (!match) throw new Error("payload");
      xmlPayload = JSON.parse(match[1].replaceAll("]]]]><![CDATA[>", "]]>") ) as V6WorkbookPayload;
    }
    if (propertiesPayload && xmlPayload && stableHash(propertiesPayload) !== stableHash(xmlPayload)) throw new Error("metadata-copies-disagree");
    const payload = propertiesPayload ?? xmlPayload;
    if (!payload) return undefined;
    if (payload?.metadata?.formatKey !== FORMAT_KEY || !["v11", "v10", "v9", "v8", "v7", "v6"].includes(payload.metadata.formatVersion) || !validTimestamp(payload.metadata.exportedAt)
      || !Array.isArray(payload.rows) || !payload.dictionaries || !Array.isArray(payload.dictionaries.domainL0)
      || !Array.isArray(payload.dictionaries.domainL1) || !Array.isArray(payload.dictionaries.products)) {
      throw new Error("schema");
    }
    const ids = new Set<string>();
    const rowTokens = new Set<string>();
    if (["v7", "v8", "v9", "v10", "v11"].includes(payload.metadata.formatVersion) && (!Array.isArray(payload.dictionaries.members) || payload.dictionaries.members.some((member) => !member.id || typeof member.label !== "string" || typeof member.role !== "string"))) throw new Error("members");
    if (isSelectiveWorkbookVersion(payload.metadata.formatVersion) && payload.integrity !== stableHash({ metadata: payload.metadata, dictionaries: payload.dictionaries, rows: payload.rows })) throw new Error("integrity");
    for (const row of payload.rows) {
      if (!row.requirementId || ids.has(row.requirementId) || !validTimestamp(row.baselineUpdatedAt)
        || !/^[a-f0-9]{64}$/i.test(row.baselineHash) || !Array.isArray(row.baselineVisible)
        || (row.rowToken !== undefined && (!/^[a-f0-9]{24}$/i.test(row.rowToken) || rowTokens.has(row.rowToken.toLowerCase())))
        || row.baselineVisible.length !== (isSelectiveWorkbookVersion(payload.metadata.formatVersion) ? 16 + workbookSeriesForVersion(payload.metadata.formatVersion).length : payload.metadata.formatVersion === "v7" ? 18 : 14)
        || (isSelectiveWorkbookVersion(payload.metadata.formatVersion) && (!row.baselineFields || !row.baselineSeries || V8_FIELD_KEYS.some((key) => typeof row.baselineFields?.[key] !== "string") || workbookSeriesForVersion(payload.metadata.formatVersion).some((series) => typeof row.baselineSeries?.[series.id] !== "string")))) {
        throw new Error("row");
      }
      ids.add(row.requirementId);
      if (row.rowToken) rowTokens.add(row.rowToken.toLowerCase());
    }
    return payload;
  } catch {
    throw new Error("协作 Excel 的内部系统信息缺失或损坏，请重新导出后再导入");
  }
}

function canonicalVisibleRow(values: readonly ExcelJS.CellValue[]): string[] {
  return values.map((value, index) => {
    if (index === 9) return parseMonth(value) ?? "";
    if (index >= 11 && index <= 13) {
      const numeric = Number(value ?? 0);
      return Number.isFinite(numeric) ? String(roundWorkload(numeric)) : "";
    }
    if (value instanceof Date) return value.toISOString();
    return String(value ?? "").trim();
  });
}

function createGuideSheet(workbook: ExcelJS.Workbook, data: WorkspaceData, appVersion: string, exportedAt: string, range?: HalfYearRange): void {
  const sheet = workbook.addWorksheet(SHEETS.guide, { properties: { tabColor: { argb: "FF6B5CDA" } } });
  sheet.views = [{ state: "frozen", ySplit: 2, showGridLines: false }];
  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "需求路标协作工作簿";
  sheet.getCell("A1").font = { name: "Microsoft YaHei", size: 20, bold: true, color: { argb: "FF17212B" } };
  sheet.getCell("A1").alignment = { vertical: "middle" };
  sheet.getRow(1).height = 38;
  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value = "上传到企业在线表格协作，下载定稿后回到解决方案需求管理导入。不要修改标记为“请勿修改”的系统字段。";
  sheet.getCell("A2").font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF66717B" } };
  sheet.getCell("A2").alignment = { wrapText: true, vertical: "middle" };
  sheet.getRow(2).height = 34;

  const instructions = [
    ["步骤", "操作", "说明"],
    [1, "导出", range ? `工作台仅导出 ${range.start} 至 ${range.end} 区间内的需求和路标卡片；领域与产品字典完整保留。` : "工作台导出全部结构化数据。"],
    [2, "协作", "将本文件上传到企业已有在线协作平台，多人共同编辑同一份文件。"],
    [3, "下载", "协作完成后下载为标准 .xlsx 文件，不要转换成 CSV。"],
    [4, "导入", "工作台先检查新增、修改、删除、冲突和错误，确认后再合并。"],
    ["填写规则", "删除需求", "在“需求清单”的操作列选择“删除”；直接删除整行不会删除应用中的需求。"],
    ["填写规则", "领域层级", "领域 L1 必须关联一个上游 L0；需求中的 L1 必须属于所选 L0。路标卡片标题由领域 L1 标签自动生成，不支持手工标题或摘要。"],
    ["填写规则", "匹配产品", "体验优化不填写产品；平台演进可选一个产品；产品专属必填且只能选择一个产品。"],
    ["填写规则", "来源与范围", `来源必须从下拉列表选择：${SOURCES.join("、")}；适用范围必须选择：${REQUIREMENT_SCOPES.join("、")}。`],
  ];
  writeMatrix(sheet, 4, 1, instructions);
  styleHeader(sheet, "A4:C4");
  forEachCell(sheet, "A5:C12", (cell) => {
    cell.alignment = { vertical: "top", wrapText: true };
    cell.font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF3E4952" } };
    cell.border = lightBorders();
  });

  const domainL0Count = data.domains.filter((item) => item.level === "L0").length;
  const domainL1Count = data.domains.filter((item) => (item.level ?? "L1") === "L1").length;
  sheet.getCell("A13").value = "格式标识（请勿修改）";
  sheet.getCell("B13").value = FORMAT_KEY;
  sheet.getCell("A14").value = "格式版本（请勿修改）";
  sheet.getCell("B14").value = FORMAT_VERSION;
  sheet.getCell("A15").value = "导出时间（请勿修改）";
  sheet.getCell("B15").value = exportedAt;
  sheet.getCell("A16").value = "应用版本（请勿修改）";
  sheet.getCell("B16").value = appVersion;
  sheet.getCell("A17").value = "数据摘要";
  sheet.getCell("B17").value = `${data.requirements.length} 条需求 · L0 ${domainL0Count} 个 / L1 ${domainL1Count} 个 · ${data.products.length} 个产品 · ${data.groupOverrides.length} 项卡片编辑`;
  sheet.getCell("A18").value = "导出区间";
  sheet.getCell("B18").value = range ? `${range.start} – ${range.end}` : "完整工作区";
  forEachCell(sheet, "A13:B18", (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F4F6" } };
    cell.border = lightBorders();
  });
  forEachCell(sheet, "A13:A18", (cell) => { cell.font = { name: "Microsoft YaHei", size: 9, bold: true, color: { argb: "FF66717B" } }; });
  forEachCell(sheet, "B13:B18", (cell) => { cell.font = { name: "Consolas", size: 9, color: { argb: "FF36414A" } }; });
  sheet.columns = [{ width: 20 }, { width: 36 }, { width: 82 }, { width: 4 }];
}

function createDictionarySheet(workbook: ExcelJS.Workbook, name: string, tableName: string, items: DictionaryItem[]): void {
  const sheet = workbook.addWorksheet(name, { properties: { tabColor: { argb: "FF6B5CDA" } } });
  prepareDataSheet(sheet, `${name} · 可新增、改名、排序和启用/停用`, DICTIONARY_HEADERS.length);
  const rows = items.map((item) => [item.name, item.active ? "启用" : "停用", item.sortOrder, item.id, dictionaryHash(item)]);
  addTable(sheet, tableName, DICTIONARY_HEADERS, rows.length ? rows : [blankCells(DICTIONARY_HEADERS.length)]);
  sheet.columns = [{ width: 30 }, { width: 12 }, { width: 10 }, { width: 40 }, { width: 68 }];
  const lastRow = Math.max(34, FIRST_DATA_ROW + rows.length + 29);
  applyListValidation(sheet, `B${FIRST_DATA_ROW}:B${lastRow}`, ["启用", "停用"]);
  applyWholeNumberValidation(sheet, `C${FIRST_DATA_ROW}:C${lastRow}`, 0);
  styleEditable(sheet, `A${FIRST_DATA_ROW}:C${Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + rows.length - 1)}`);
  styleSystem(sheet, `D${FIRST_DATA_ROW}:E${Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + rows.length - 1)}`);
}

function createDomainL1Sheet(workbook: ExcelJS.Workbook, data: WorkspaceData): void {
  const sheet = workbook.addWorksheet(SHEETS.domainL1s, { properties: { tabColor: { argb: "FF6B5CDA" } } });
  prepareDataSheet(sheet, `${SHEETS.domainL1s} · 每个 L1 必须选择一个所属 L0`, DOMAIN_L1_HEADERS.length);
  const l0Names = new Map(data.domains.filter((item) => item.level === "L0").map((item) => [item.id, item.name]));
  const items = data.domains.filter((item) => (item.level ?? "L1") === "L1");
  const rows = items.map((item) => [
    item.name,
    item.parentId ? (l0Names.get(item.parentId) ?? item.parentId) : null,
    item.active ? "启用" : "停用",
    item.sortOrder,
    item.id,
    item.parentId ?? null,
    item.parentId ? (l0Names.get(item.parentId) ?? "") : null,
    dictionaryHash(item),
  ]);
  addTable(sheet, "DomainL1Table", DOMAIN_L1_HEADERS, rows.length ? rows : [blankCells(DOMAIN_L1_HEADERS.length)]);
  sheet.columns = [{ width: 28 }, { width: 28 }, { width: 12 }, { width: 10 }, { width: 40 }, { width: 40 }, { width: 28 }, { width: 68 }];
  const usedLast = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + rows.length - 1);
  const validationLast = Math.max(34, usedLast + 30);
  const l0Count = data.domains.filter((item) => item.level === "L0").length;
  const l0Last = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + l0Count - 1);
  for (let row = FIRST_DATA_ROW; row <= validationLast; row += 1) {
    sheet.getCell(row, 2).dataValidation = {
      type: "list",
      allowBlank: false,
      formulae: [`'${SHEETS.domainL0s}'!$A$${FIRST_DATA_ROW}:$A$${l0Last}`],
      showErrorMessage: true,
      errorTitle: "请选择所属领域 L0",
      error: "领域 L1 必须关联领域 L0 字典中的一个项目。",
    };
  }
  applyListValidation(sheet, `C${FIRST_DATA_ROW}:C${validationLast}`, ["启用", "停用"]);
  applyWholeNumberValidation(sheet, `D${FIRST_DATA_ROW}:D${validationLast}`, 0);
  styleEditable(sheet, `A${FIRST_DATA_ROW}:D${usedLast}`);
  styleSystem(sheet, `E${FIRST_DATA_ROW}:H${usedLast}`);
}

function createRequirementSheet(workbook: ExcelJS.Workbook, data: WorkspaceData): void {
  const sheet = workbook.addWorksheet(SHEETS.requirements, { properties: { tabColor: { argb: "FFED6A5A" } } });
  prepareDataSheet(sheet, "需求清单 · 三端工作量可留空；合计为 0 时回到需求池标记为待评估", REQUIREMENT_HEADERS.length, 2);
  const domains = new Map(data.domains.map((item) => [item.id, item.name]));
  const products = new Map(data.products.map((item) => [item.id, item.name]));
  const rows = data.requirements.map((item, index) => {
    const rowNumber = FIRST_DATA_ROW + index;
    const productNames = item.productIds.map((id) => products.get(id) ?? id);
    return [
      "保留",
      item.title,
      item.description || null,
      domains.get(item.domainL0Id) ?? item.domainL0Id,
      domains.get(item.domainId) ?? item.domainId,
      item.source,
      item.scope,
      item.category,
      monthDate(item.targetMonth),
      productNames[0] ?? null,
      item.deviceWorkloadPm,
      item.appWorkloadPm,
      item.cloudWorkloadPm,
      item.unallocatedWorkloadPm,
      { formula: `SUM(K${rowNumber}:N${rowNumber})`, result: item.workloadPm },
      item.images.length,
      item.id,
      item.domainL0Id,
      item.domainId,
      item.productIds[0] ?? null,
      domains.get(item.domainL0Id) ?? "",
      domains.get(item.domainId) ?? "",
      productNames[0] ?? null,
      item.createdAt,
      item.updatedAt,
      requirementHash(item),
    ];
  });
  addTable(sheet, "RequirementsTable", REQUIREMENT_HEADERS, rows.length ? rows : [blankCells(REQUIREMENT_HEADERS.length)]);

  const widths = [10, 28, 48, 20, 20, 16, 14, 14, 14, 28, 16, 16, 16, 19, 19, 16, 40, 40, 40, 40, 24, 24, 32, 28, 28, 68];
  sheet.columns = widths.map((width) => ({ width }));
  sheet.getColumn(9).numFmt = "yyyy-mm";
  for (const column of [11, 12, 13, 14, 15]) sheet.getColumn(column).numFmt = "0.00";
  const usedLast = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + rows.length - 1);
  const validationLast = Math.max(usedLast + 30, 34);
  applyListValidation(sheet, `A${FIRST_DATA_ROW}:A${validationLast}`, ["保留", "删除"]);
  applyListValidation(sheet, `F${FIRST_DATA_ROW}:F${validationLast}`, [...SOURCES]);
  applyListValidation(sheet, `G${FIRST_DATA_ROW}:G${validationLast}`, [...REQUIREMENT_SCOPES]);
  applyListValidation(sheet, `H${FIRST_DATA_ROW}:H${validationLast}`, [...CATEGORIES]);
  const domainL0Count = data.domains.filter((item) => item.level === "L0").length;
  const domainL1Count = data.domains.filter((item) => (item.level ?? "L1") === "L1").length;
  const domainL0Last = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + domainL0Count - 1);
  const domainL1Last = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + domainL1Count - 1);
  const productLast = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + data.products.length - 1);
  for (let row = FIRST_DATA_ROW; row <= validationLast; row += 1) {
    sheet.getCell(row, 4).dataValidation = {
      type: "list",
      allowBlank: false,
      formulae: [`'${SHEETS.domainL0s}'!$A$${FIRST_DATA_ROW}:$A$${domainL0Last}`],
      showErrorMessage: true,
      errorTitle: "请选择领域 L0",
      error: "领域 L0 必须来自领域 L0 字典。",
    };
    sheet.getCell(row, 5).dataValidation = {
      type: "list",
      allowBlank: false,
      formulae: [`'${SHEETS.domainL1s}'!$A$${FIRST_DATA_ROW}:$A$${domainL1Last}`],
      showErrorMessage: true,
      errorTitle: "请选择领域 L1",
      error: "领域 L1 必须来自领域 L1 字典。",
    };
    sheet.getCell(row, 10).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: [`'${SHEETS.products}'!$A$${FIRST_DATA_ROW}:$A$${productLast}`],
      showErrorMessage: true,
      errorTitle: "请选择一个产品",
      error: "匹配产品必须来自产品字典，且只能选择一个。",
    };
  }
  for (const column of [11, 12, 13]) applyDecimalValidation(sheet, `${columnLetter(column)}${FIRST_DATA_ROW}:${columnLetter(column)}${validationLast}`, 0, true);
  styleEditable(sheet, `A${FIRST_DATA_ROW}:M${usedLast}`);
  styleSystem(sheet, `N${FIRST_DATA_ROW}:Z${usedLast}`);
  sheet.getColumn(3).alignment = { wrapText: true, vertical: "top" };
}

function createOverrideSheet(workbook: ExcelJS.Workbook, data: WorkspaceData): void {
  const sheet = workbook.addWorksheet(SHEETS.overrides, { properties: { tabColor: { argb: "FF2F9D82" } } });
  prepareDataSheet(sheet, "路标卡片 · 标题由领域 L1 标签自动生成；本表仅保留卡片顺序和系统识别字段", OVERRIDE_HEADERS.length);
  const domains = new Map(data.domains.map((item) => [item.id, item.name]));
  const rows = data.groupOverrides.map((item) => {
    const [domainId, source, targetMonth] = item.groupKey.split("::");
    return [domains.get(domainId) ?? domainId, source, monthDate(targetMonth), item.sortOrder ?? null, item.groupKey, item.updatedAt, overrideOrderHash(item)];
  });
  addTable(sheet, "RoadmapCardsTable", OVERRIDE_HEADERS, rows.length ? rows : [blankCells(OVERRIDE_HEADERS.length)]);
  sheet.columns = [{ width: 24 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 58 }, { width: 28 }, { width: 68 }];
  sheet.getColumn(3).numFmt = "yyyy-mm";
  const usedLast = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + rows.length - 1);
  styleSystem(sheet, `A${FIRST_DATA_ROW}:G${usedLast}`);
}

function prepareDataSheet(sheet: ExcelJS.Worksheet, title: string, columnCount: number, freezeColumns = 1): void {
  sheet.views = [{ state: "frozen", xSplit: freezeColumns, ySplit: HEADER_ROW, showGridLines: false }];
  sheet.mergeCells(1, 1, 1, columnCount);
  sheet.getCell(1, 1).value = title;
  sheet.getCell(1, 1).font = { name: "Microsoft YaHei", size: 16, bold: true, color: { argb: "FF17212B" } };
  sheet.getCell(1, 1).alignment = { vertical: "middle" };
  sheet.getRow(1).height = 34;
  sheet.mergeCells(2, 1, 2, columnCount);
  sheet.getCell(2, 1).value = "黄色为协作填写区；灰色为系统识别字段。请保留工作表名称、表头和系统字段。";
  sheet.getCell(2, 1).font = { name: "Microsoft YaHei", size: 9, color: { argb: "FF66717B" } };
  sheet.getCell(2, 1).alignment = { vertical: "middle" };
  sheet.getRow(2).height = 24;
}

function addTable(sheet: ExcelJS.Worksheet, name: string, headers: readonly string[], rows: ExcelJS.CellValue[][]): void {
  sheet.addTable({
    name,
    ref: `A${HEADER_ROW}`,
    headerRow: true,
    totalsRow: false,
    style: { theme: "TableStyleMedium2", showRowStripes: true },
    columns: headers.map((header) => ({ name: header, filterButton: true })),
    rows,
  });
  sheet.getRow(HEADER_ROW).height = 30;
  sheet.getRow(HEADER_ROW).font = { name: "Microsoft YaHei", size: 9, bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(HEADER_ROW).alignment = { vertical: "middle", wrapText: true };
  for (let row = FIRST_DATA_ROW; row <= FIRST_DATA_ROW + rows.length - 1; row += 1) {
    sheet.getRow(row).font = { name: "Microsoft YaHei", size: 10, color: { argb: "FF26313A" } };
    sheet.getRow(row).alignment = { vertical: "top" };
    sheet.getRow(row).height = 28;
  }
}

function blankCells(length: number): ExcelJS.CellValue[] {
  return Array.from({ length }, () => null);
}

function parseDictionarySheet(sheet: ExcelJS.Worksheet, errors: WorkspaceWorkbookIssue[], level?: DomainLevel): ParsedDictionaryRow[] {
  const columns = headerColumns(sheet, DICTIONARY_HEADERS, errors);
  if (!columns) return [];
  const rows: ParsedDictionaryRow[] = [];
  for (let row = FIRST_DATA_ROW; row <= sheet.actualRowCount; row += 1) {
    const name = valueAt(sheet, row, columns, DICTIONARY_HEADERS[0]);
    const status = valueAt(sheet, row, columns, DICTIONARY_HEADERS[1]);
    const sortText = valueAt(sheet, row, columns, DICTIONARY_HEADERS[2]);
    const rawId = valueAt(sheet, row, columns, DICTIONARY_HEADERS[3]);
    const baselineHash = valueAt(sheet, row, columns, DICTIONARY_HEADERS[4]);
    if (![name, status, sortText, rawId, baselineHash].some(Boolean)) continue;
    if (!name) addIssue(errors, sheet.name, cellAddress(row, columns.get(DICTIONARY_HEADERS[0])!), "名称不能为空");
    if (!rawId && baselineHash) addIssue(errors, sheet.name, cellAddress(row, columns.get(DICTIONARY_HEADERS[3])!), "系统ID缺失，请不要修改系统字段");
    if (status !== "启用" && status !== "停用") addIssue(errors, sheet.name, cellAddress(row, columns.get(DICTIONARY_HEADERS[1])!), "状态只能是“启用”或“停用”");
    const sortOrder = parseNonNegativeInteger(sortText);
    if (sortOrder === null) addIssue(errors, sheet.name, cellAddress(row, columns.get(DICTIONARY_HEADERS[2])!), "排序必须是大于或等于0的整数");
    const id = rawId || crypto.randomUUID();
    rows.push({
      sheet: sheet.name,
      row,
      item: { id, name: name.trim(), sortOrder: sortOrder ?? 0, active: status !== "停用", ...(level ? { level } : {}) },
      baselineHash,
    });
  }
  return rows;
}

function parseDomainL1Sheet(
  sheet: ExcelJS.Worksheet,
  domainL0s: ParsedDictionaryRow[],
  errors: WorkspaceWorkbookIssue[],
): ParsedDictionaryRow[] {
  const columns = headerColumns(sheet, DOMAIN_L1_HEADERS, errors);
  if (!columns) return [];
  const l0ByName = dictionaryByName(domainL0s);
  const l0ById = new Map(domainL0s.map((row) => [row.item.id, row.item]));
  const rows: ParsedDictionaryRow[] = [];
  for (let row = FIRST_DATA_ROW; row <= sheet.actualRowCount; row += 1) {
    const values = Object.fromEntries(DOMAIN_L1_HEADERS.map((header) => [header, valueAt(sheet, row, columns, header)])) as Record<(typeof DOMAIN_L1_HEADERS)[number], string>;
    if (!Object.values(values).some(Boolean)) continue;
    const name = values[DOMAIN_L1_HEADERS[0]].trim();
    const parentName = values[DOMAIN_L1_HEADERS[1]].trim();
    const status = values[DOMAIN_L1_HEADERS[2]];
    const sortText = values[DOMAIN_L1_HEADERS[3]];
    const rawId = values[DOMAIN_L1_HEADERS[4]];
    const rawParentId = values[DOMAIN_L1_HEADERS[5]];
    const baselineParentName = values[DOMAIN_L1_HEADERS[6]].trim();
    const baselineHash = values[DOMAIN_L1_HEADERS[7]];
    if (!name) addIssue(errors, sheet.name, cellAddress(row, columns.get(DOMAIN_L1_HEADERS[0])!), "名称不能为空");
    if (!rawId && baselineHash) addIssue(errors, sheet.name, cellAddress(row, columns.get(DOMAIN_L1_HEADERS[4])!), "系统ID缺失，请不要修改系统字段");
    if (status !== "启用" && status !== "停用") addIssue(errors, sheet.name, cellAddress(row, columns.get(DOMAIN_L1_HEADERS[2])!), "状态只能是“启用”或“停用”");
    const sortOrder = parseNonNegativeInteger(sortText);
    if (sortOrder === null) addIssue(errors, sheet.name, cellAddress(row, columns.get(DOMAIN_L1_HEADERS[3])!), "排序必须是大于或等于0的整数");
    const parent = l0ByName.get(normalizeName(parentName))
      ?? (parentName === baselineParentName ? l0ById.get(rawParentId) : undefined);
    if (!parent) addIssue(errors, sheet.name, cellAddress(row, columns.get(DOMAIN_L1_HEADERS[1])!), `所属领域 L0“${parentName || "空"}”不存在于领域 L0 字典`);
    rows.push({
      sheet: sheet.name,
      row,
      item: {
        id: rawId || crypto.randomUUID(),
        name,
        sortOrder: sortOrder ?? 0,
        active: status !== "停用",
        level: "L1",
        parentId: parent?.id ?? rawParentId,
      },
      baselineHash,
    });
  }
  return rows;
}

function parseRequirementSheet(
  sheet: ExcelJS.Worksheet,
  domainL0s: ParsedDictionaryRow[],
  domainL1s: ParsedDictionaryRow[],
  products: ParsedDictionaryRow[],
  exportedAt: string,
  formatVersion: ParsedWorkspaceWorkbook["formatVersion"],
  errors: WorkspaceWorkbookIssue[],
): ParsedRequirementRow[] {
  const headers = formatVersion === "v2" ? LEGACY_REQUIREMENT_HEADERS : REQUIREMENT_HEADERS;
  const columns = headerColumns(sheet, headers, errors) as Map<string, number> | null;
  if (!columns) return [];
  const domainL0ByName = dictionaryByName(domainL0s);
  const domainL0ById = new Map(domainL0s.map((row) => [row.item.id, row.item]));
  const domainL1ByName = dictionaryByName(domainL1s);
  const domainL1ById = new Map(domainL1s.map((row) => [row.item.id, row.item]));
  const productByName = dictionaryByName(products);
  const productById = new Map(products.map((row) => [row.item.id, row.item]));
  const rows: ParsedRequirementRow[] = [];
  const usedIds = new Set<string>();

  for (let row = FIRST_DATA_ROW; row <= sheet.actualRowCount; row += 1) {
    const values = Object.fromEntries(headers.map((header) => [header, valueAt(sheet, row, columns, header)])) as Record<string, string>;
    if (!Object.values(values).some(Boolean)) continue;
    const action = normalizeAction(values["操作"], sheet.name, row, columns.get("操作")!, errors);
    const rawId = values["需求ID（请勿修改）"];
    const baselineHash = values["基线校验（请勿修改）"];
    if (action === "删除") {
      if (!rawId) addIssue(errors, sheet.name, cellAddress(row, columns.get("需求ID（请勿修改）")!), "新增行不能标记删除");
      rows.push({ sheet: sheet.name, row, action, baselineHash, requirement: deletedRequirementPlaceholder(rawId, exportedAt) });
      continue;
    }

    const title = values["需求标题"].trim();
    const description = values["需求描述"].trim();
    if (!title) addIssue(errors, sheet.name, cellAddress(row, columns.get("需求标题")!), "需求标题不能为空");
    if (description.length > 5000) addIssue(errors, sheet.name, cellAddress(row, columns.get("需求描述")!), "需求描述不能超过5000字");
    if (!rawId && baselineHash) addIssue(errors, sheet.name, cellAddress(row, columns.get("需求ID（请勿修改）")!), "系统ID缺失，请不要修改系统字段");
    const id = rawId || crypto.randomUUID();
    if (usedIds.has(id)) addIssue(errors, sheet.name, cellAddress(row, columns.get("需求ID（请勿修改）")!), `需求ID重复：${id}`);
    usedIds.add(id);

    const domainL0Name = values["领域L0"].trim();
    const baselineDomainL0Id = values["领域L0 ID（请勿修改）"];
    const baselineDomainL0Name = values["基线领域L0名称（请勿修改）"].trim();
    const domainL0 = domainL0ByName.get(normalizeName(domainL0Name))
      ?? (domainL0Name === baselineDomainL0Name ? domainL0ById.get(baselineDomainL0Id) : undefined);
    if (!domainL0) addIssue(errors, sheet.name, cellAddress(row, columns.get("领域L0")!), `领域 L0“${domainL0Name || "空"}”不存在于领域 L0 字典`);

    const domainL1Name = values["领域L1"].trim();
    const baselineDomainL1Id = values["领域L1 ID（请勿修改）"];
    const baselineDomainL1Name = values["基线领域L1名称（请勿修改）"].trim();
    const domainL1 = domainL1ByName.get(normalizeName(domainL1Name))
      ?? (domainL1Name === baselineDomainL1Name ? domainL1ById.get(baselineDomainL1Id) : undefined);
    if (!domainL1) addIssue(errors, sheet.name, cellAddress(row, columns.get("领域L1")!), `领域 L1“${domainL1Name || "空"}”不存在于领域 L1 字典`);
    if (domainL0 && domainL1 && domainL1.parentId !== domainL0.id) addIssue(errors, sheet.name, cellAddress(row, columns.get("领域L1")!), `领域 L1“${domainL1Name}”不属于领域 L0“${domainL0Name}”`);

    const rawSource = values["来源"];
    const source = normalizeLegacySource(rawSource);
    const scope = formatVersion === "v2" ? "全球" : values["适用范围"];
    const category = values["分类"];
    if (!SOURCES.includes(source as (typeof SOURCES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get("来源")!), "来源不在允许范围内");
    if (!REQUIREMENT_SCOPES.includes(scope as (typeof REQUIREMENT_SCOPES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get("适用范围")!), "适用范围必须选择仅国内、仅海外或全球");
    if (!CATEGORIES.includes(category as (typeof CATEGORIES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get("分类")!), "分类不在允许范围内");
    const targetMonth = parseMonth(sheet.getCell(row, columns.get("上线月份")!).value);
    if (!targetMonth && !(isMedicalSubmissionSource(source) && cellText(sheet.getCell(row, columns.get("上线月份")!)) === "")) addIssue(errors, sheet.name, cellAddress(row, columns.get("上线月份")!), "上线月份应为有效年月；仅医疗认证来源可留空");

    const productNames = splitList(values["匹配产品（单选）"]);
    const baselineProductIds = splitList(values["产品ID（请勿修改）"]);
    const baselineProductNames = splitList(values["基线产品名称（请勿修改）"]);
    const productIds: string[] = [];
    if (productNames.length > 1) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品（单选）")!), "匹配产品只能选择一个");
    productNames.forEach((name) => {
      const direct = productByName.get(normalizeName(name));
      const baselineIndex = baselineProductNames.findIndex((item) => normalizeName(item) === normalizeName(name));
      const fallback = baselineIndex >= 0 ? productById.get(baselineProductIds[baselineIndex] ?? "") : undefined;
      const product = direct ?? fallback;
      if (!product) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品（单选）")!), `产品“${name}”不存在于产品字典`);
      else productIds.push(product.id);
    });
    if (category === "体验优化" && productIds.length > 0) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品（单选）")!), "体验优化需求不匹配产品");
    if (category === "产品专属" && productIds.length !== 1) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品（单选）")!), "产品专属需求必须且只能匹配一个产品");
    if (category === "平台演进" && productIds.length > 1) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品（单选）")!), "平台演进需求最多匹配一个产品");

    const workloadHeaders = ["设备工作量（人月）", "App工作量（人月）", "云侧工作量（人月）"] as const;
    const device = parseNonNegativeNumber(values[workloadHeaders[0]]);
    const app = parseNonNegativeNumber(values[workloadHeaders[1]]);
    const cloud = parseNonNegativeNumber(values[workloadHeaders[2]]);
    const unallocated = parseNonNegativeNumber(values["待拆分工作量（只读）"]);
    [device, app, cloud].forEach((value, index) => {
      if (value === null) addIssue(errors, sheet.name, cellAddress(row, columns.get(workloadHeaders[index])!), "工作量应留空或输入大于或等于0的数字");
    });
    if (unallocated === null) addIssue(errors, sheet.name, cellAddress(row, columns.get("待拆分工作量（只读）")!), "待拆分工作量必须是大于或等于0的数字");
    const total = roundWorkload((device ?? 0) + (app ?? 0) + (cloud ?? 0) + (unallocated ?? 0));

    const createdAt = values["创建时间（请勿修改）"] || exportedAt;
    const updatedAt = values["基线更新时间（请勿修改）"] || exportedAt;
    if (!validTimestamp(createdAt) || !validTimestamp(updatedAt)) addIssue(errors, sheet.name, cellAddress(row, columns.get("基线更新时间（请勿修改）")!), "系统时间字段无效，请不要修改");
    const requirement: Requirement = {
      id,
      title,
      description,
      images: [],
      domainL0Id: domainL0?.id ?? baselineDomainL0Id,
      domainId: domainL1?.id ?? baselineDomainL1Id,
      source: source as Requirement["source"],
      scope: scope as Requirement["scope"],
      allowLatestVersionUpgrade: false,
      category: category as Requirement["category"],
      targetMonth: targetMonth ?? "",
      productIds: category === "体验优化" ? [] : [...new Set(productIds)],
      overseasRegions: [],
      deviceWorkloadPm: roundWorkload(device ?? 0),
      appWorkloadPm: roundWorkload(app ?? 0),
      cloudWorkloadPm: roundWorkload(cloud ?? 0),
      unallocatedWorkloadPm: roundWorkload(unallocated ?? 0),
      workloadPm: total,
      createdAt,
      updatedAt,
    };
    rows.push({
      sheet: sheet.name,
      row,
      action,
      baselineHash,
      requirement,
    });
  }
  return rows;
}

function parseRequirementSheetV6(
  sheet: ExcelJS.Worksheet,
  current: WorkspaceData,
  payload: V6WorkbookPayload,
  errors: WorkspaceWorkbookIssue[],
): ParsedRequirementRow[] {
  const headers = payload.metadata.formatVersion === "v7" ? V7_REQUIREMENT_HEADERS : V6_REQUIREMENT_HEADERS;
  const columns = headerColumnsAt(sheet, headers, V6_HEADER_ROW, errors);
  if (!columns) return [];
  const currentRequirements = new Map(current.requirements.map((item) => [item.id, item]));
  const currentDomainsById = new Map(current.domains.map((item) => [item.id, item]));
  const currentProductsById = new Map(current.products.map((item) => [item.id, item]));
  const domainL0ByName = new Map(current.domains.filter((item) => item.level === "L0").map((item) => [normalizeName(item.name), item]));
  const domainL1ByName = new Map(current.domains.filter((item) => (item.level ?? "L1") === "L1").map((item) => [normalizeName(item.name), item]));
  const productByName = new Map(current.products.map((item) => [normalizeName(item.name), item]));
  for (const item of payload.dictionaries.domainL0) {
    const currentItem = currentDomainsById.get(item.id);
    if (currentItem?.level === "L0" && !domainL0ByName.has(normalizeName(item.name))) domainL0ByName.set(normalizeName(item.name), currentItem);
  }
  for (const item of payload.dictionaries.domainL1) {
    const currentItem = currentDomainsById.get(item.id);
    if (currentItem && (currentItem.level ?? "L1") === "L1" && !domainL1ByName.has(normalizeName(item.name))) domainL1ByName.set(normalizeName(item.name), currentItem);
  }
  for (const item of payload.dictionaries.products) {
    const currentItem = currentProductsById.get(item.id);
    if (currentItem && !productByName.has(normalizeName(item.name))) productByName.set(normalizeName(item.name), currentItem);
  }
  const rows: ParsedRequirementRow[] = [];
  const remainingMetadata = new Map(payload.rows.map((item) => [item.requirementId, item]));

  for (let row = V6_FIRST_DATA_ROW; row <= sheet.actualRowCount; row += 1) {
    const values = Object.fromEntries(headers.map((header) => [header, valueAt(sheet, row, columns, header)])) as Record<(typeof V7_REQUIREMENT_HEADERS)[number], string>;
    const visibleValues = headers.map((header, index) => {
      if (index === 9) return parseMonth(sheet.getCell(row, columns.get(header)!).value) ?? values[header];
      if (index >= 11 && index <= 13) {
        const numeric = parseNonNegativeNumber(values[header]);
        return numeric === null ? values[header] : String(roundWorkload(numeric));
      }
      return values[header];
    });
    if (!visibleValues.some(Boolean)) continue;
    const matchedMetadata = matchV6RowMetadata(visibleValues, [...remainingMetadata.values()]);
    if (!matchedMetadata) {
      addIssue(errors, sheet.name, `A${row}`, "无法匹配原始需求：修改专用 Excel 不支持新增需求、复制行或破坏内部系统信息");
      continue;
    }
    remainingMetadata.delete(matchedMetadata.requirementId);
    const rawId = matchedMetadata.requirementId;
    const existing = currentRequirements.get(rawId);
    if (!existing) {
      addIssue(errors, sheet.name, `A${row}`, "内部需求 ID 不存在或不属于当前工作区，不能作为新增需求导入");
      continue;
    }
    const baselineUpdatedAt = matchedMetadata.baselineUpdatedAt;
    const baselineHash = matchedMetadata.baselineHash;
    const action = normalizeAction(values["操作"], sheet.name, row, columns.get("操作")!, errors);
    if (action === "删除") {
      rows.push({ sheet: sheet.name, row, action, baselineHash, requirement: deletedRequirementPlaceholder(rawId, baselineUpdatedAt || payload.metadata.exportedAt) });
      continue;
    }

    const title = values["需求标题"].trim();
    const description = values["需求描述"].trim();
    if (!title) addIssue(errors, sheet.name, cellAddress(row, columns.get("需求标题")!), "需求标题不能为空");
    if (description.length > 5000) addIssue(errors, sheet.name, cellAddress(row, columns.get("需求描述")!), "需求描述不能超过5000字");

    const domainL0Name = values["领域 L0"].trim();
    const domainL1Name = values["领域 L1"].trim();
    const domainL0 = domainL0ByName.get(normalizeName(domainL0Name));
    const domainL1 = domainL1ByName.get(normalizeName(domainL1Name));
    if (!domainL0) addIssue(errors, sheet.name, cellAddress(row, columns.get("领域 L0")!), `领域 L0“${domainL0Name || "空"}”不存在于当前字典`);
    if (!domainL1) addIssue(errors, sheet.name, cellAddress(row, columns.get("领域 L1")!), `领域 L1“${domainL1Name || "空"}”不存在于当前字典`);
    if (domainL0 && domainL1 && domainL1.parentId !== domainL0.id) addIssue(errors, sheet.name, cellAddress(row, columns.get("领域 L1")!), `领域 L1“${domainL1Name}”不属于领域 L0“${domainL0Name}”`);

    const source = current.roadmapAssignments !== undefined ? existing.source ?? (existing.targetMonth === "" ? "NMPA认证" : "行业") : normalizeLegacySource(values["来源"]);
    const scope = values["适用范围"];
    const allowUpgradeText = values["允许最近版本升级"];
    const category = values["分类"];
    if (!SOURCES.includes(source as (typeof SOURCES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get("来源")!), "来源不在允许范围内");
    if (!REQUIREMENT_SCOPES.includes(scope as (typeof REQUIREMENT_SCOPES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get("适用范围")!), "适用范围必须选择仅国内、仅海外或全球");
    if (allowUpgradeText !== "是" && allowUpgradeText !== "否") addIssue(errors, sheet.name, cellAddress(row, columns.get("允许最近版本升级")!), "允许最近版本升级必须选择“是”或“否”");
    if (!CATEGORIES.includes(category as (typeof CATEGORIES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get("分类")!), "分类不在允许范围内");
    const targetMonth = parseMonth(sheet.getCell(row, columns.get("上线月份")!).value);
    if (!targetMonth && !(isMedicalSubmissionSource(source) && cellText(sheet.getCell(row, columns.get("上线月份")!)) === "")) addIssue(errors, sheet.name, cellAddress(row, columns.get("上线月份")!), "上线月份应为有效年月；仅医疗认证来源可留空");

    const productNames = splitList(values["匹配产品"]);
    const productIds: string[] = [];
    if (productNames.length > 1) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品")!), "匹配产品只能选择一个");
    productNames.forEach((name) => {
      const product = productByName.get(normalizeName(name));
      if (!product) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品")!), `产品“${name}”不存在于当前字典`);
      else productIds.push(product.id);
    });
    if (category === "体验优化" && productIds.length > 0) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品")!), "体验优化需求不匹配产品");
    if (category === "产品专属" && productIds.length !== 1) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品")!), "产品专属需求必须且只能匹配一个产品");
    if (category === "平台演进" && productIds.length > 1) addIssue(errors, sheet.name, cellAddress(row, columns.get("匹配产品")!), "平台演进需求最多匹配一个产品");

    const workloadHeaders = ["设备工作量", "App 工作量", "云侧工作量"] as const;
    const device = parseNonNegativeNumber(values[workloadHeaders[0]]);
    const app = parseNonNegativeNumber(values[workloadHeaders[1]]);
    const cloud = parseNonNegativeNumber(values[workloadHeaders[2]]);
    [device, app, cloud].forEach((value, index) => {
      if (value === null) addIssue(errors, sheet.name, cellAddress(row, columns.get(workloadHeaders[index])!), "工作量应留空或输入大于或等于0的数字");
    });
    const unallocated = existing?.unallocatedWorkloadPm ?? 0;
    const requirement: Requirement = {
      id: rawId,
      title,
      description,
      images: [],
      domainL0Id: domainL0?.id ?? existing?.domainL0Id ?? "",
      domainId: domainL1?.id ?? existing?.domainId ?? "",
      source: source as Requirement["source"],
      scope: scope as Requirement["scope"],
      allowLatestVersionUpgrade: allowUpgradeText === "是",
      category: category as Requirement["category"],
      targetMonth: targetMonth ?? "",
      productIds: category === "体验优化" ? [] : [...new Set(productIds)],
      overseasRegions: existing?.overseasRegions ?? [],
      deviceWorkloadPm: roundWorkload(device ?? 0),
      appWorkloadPm: roundWorkload(app ?? 0),
      cloudWorkloadPm: roundWorkload(cloud ?? 0),
      unallocatedWorkloadPm: roundWorkload(unallocated),
      workloadPm: roundWorkload((device ?? 0) + (app ?? 0) + (cloud ?? 0) + unallocated),
      createdAt: existing.createdAt,
      updatedAt: baselineUpdatedAt || payload.metadata.exportedAt,
    };
    if (payload.metadata.formatVersion === "v7") {
      for (const { key, role, required } of WORKBOOK_OWNER_FIELDS) {
        const label = values[role as "规划" | "产品经理" | "SE" | "算法"]?.trim() ?? "";
        const candidates = (payload.dictionaries.members ?? []).filter((member) => member.role === role && member.label === label);
        const selected = candidates.length === 1 ? (current.members ?? []).find((member) => member.id === candidates[0].id && member.role === role) : undefined;
        if (label && !selected) addIssue(errors, sheet.name, cellAddress(row, columns.get(role as "规划")!), `${role}成员不存在、岗位不匹配或名称不唯一，请重新导出`);
        if (!label && required && (matchedMetadata.baselineVisible[14 + WORKBOOK_OWNER_FIELDS.findIndex((field) => field.key === key)] || existing[key])) addIssue(errors, sheet.name, cellAddress(row, columns.get(role as "规划")!), `已配置的${role}不可清空`);
        requirement[key] = selected?.id ?? "";
      }
    }
    rows.push({ sheet: sheet.name, row, action, baselineHash, requirement });
  }
  return rows;
}

function visibleV8WorksheetRow(sheet: ExcelJS.Worksheet, row: number, columns: ReadonlyMap<string, number>, formatVersion: SelectiveWorkbookVersion): string[] {
  const headers = formatVersion === "v10" || formatVersion === "v11" ? V10_REQUIREMENT_HEADERS.slice(1) : V8_REQUIREMENT_HEADERS;
  return headers.map((header, index) => {
    const cell = sheet.getCell(row, columns.get(header)!);
    if (index === 8) return parseMonth(cell.value) ?? cellText(cell);
    if (index >= 9 && index <= 11) {
      const number = parseNonNegativeNumber(cellText(cell));
      return number === null ? cellText(cell) : String(roundWorkload(number));
    }
    return cellText(cell);
  });
}

function recoverOperationOnlyV8Payload(sheet: ExcelJS.Worksheet, current: WorkspaceData, formatVersion: SelectiveWorkbookVersion): V6WorkbookPayload {
  // v0.8.28 wrote row identities into cell formatting, but WPS discarded its
  // customXml payload on save. Reconstruct only when every business value still
  // exactly matches this machine's current record; otherwise the lost baseline
  // cannot be inferred safely and no update or deletion may be applied.
  const errorMessage = "此表格的内部系统信息被编辑器移除；仅操作列修改可安全恢复，其他修改请重新导出后再导入";
  const offset = formatVersion === "v8" ? 0 : 1;
  const columns = new Map(workbookHeadersForVersion(formatVersion).slice(offset).map((header, index) => [header, index + 1 + offset]));
  const currentByToken = new Map(current.requirements.map((item) => [workbookRowToken(item), item]));
  if (currentByToken.size !== current.requirements.length) throw new Error(errorMessage);
  const selected: Requirement[] = [];
  const seen = new Set<string>();
  const visibleRows: string[][] = [];
  for (let row = 2; row <= sheet.actualRowCount; row += 1) {
    const visible = visibleV8WorksheetRow(sheet, row, columns, formatVersion);
    if (!visible.some(Boolean)) continue;
    if (!["保留", "删除"].includes(visible[0])) throw new Error(errorMessage);
    const token = rowTokenFromWorksheetRow(sheet, row, offset);
    const requirement = token && currentByToken.get(token);
    if (!requirement || seen.has(token!)) throw new Error(errorMessage);
    selected.push(requirement);
    seen.add(token!);
    visibleRows.push(visible);
  }
  if (!selected.length) throw new Error(errorMessage);
  const reconstructed = createRequirementSheetV8(new ExcelJS.Workbook(), { ...current, requirements: selected }, "recovery", new Date().toISOString(), undefined, formatVersion);
  for (let index = 0; index < selected.length; index += 1) {
    if (visibleRows[index].slice(1).some((value, column) => value !== reconstructed.rows[index].baselineVisible[column + 1])) throw new Error(errorMessage);
  }
  return reconstructed;
}

function parseRequirementSheetV8(sheet: ExcelJS.Worksheet, current: WorkspaceData, payload: V6WorkbookPayload, errors: WorkspaceWorkbookIssue[]): V8ParsedRow[] {
  const formatVersion = payload.metadata.formatVersion as SelectiveWorkbookVersion;
  const offset = formatVersion === "v8" ? 0 : 1;
  const headers = workbookHeadersForVersion(formatVersion);
  const seriesForWorkbook = workbookSeriesForVersion(formatVersion);
  const columns = headerColumnsAt(sheet, headers, 1, errors);
  if (!columns) return [];
  headers.forEach((header, index) => {
    if (cellText(sheet.getCell(1, index + 1)) !== header) addIssue(errors, sheet.name, `${columnLetter(index + 1)}1`, `${payload.metadata.formatVersion} 列顺序不可修改：此列应为“${header}”`);
  });
  if (sheet.columnCount > headers.length) {
    for (let column = headers.length + 1; column <= sheet.columnCount; column += 1) {
      if (sheet.getColumn(column).values.some((value) => value !== null && value !== undefined && value !== "")) addIssue(errors, sheet.name, `${columnLetter(column)}1`, `${payload.metadata.formatVersion} 仅允许 A–${columnLetter(headers.length)} 业务列，不支持新增列`);
    }
  }
  const currentById = new Map(current.requirements.map((item) => [item.id, item]));
  const domainL0ByName = new Map(current.domains.filter((item) => item.level === "L0").map((item) => [normalizeName(item.name), item.id]));
  const domainL1ByName = new Map(current.domains.filter((item) => (item.level ?? "L1") === "L1").map((item) => [normalizeName(item.name), item.id]));
  for (const item of payload.dictionaries.domainL0) {
    if (current.domains.some((candidate) => candidate.id === item.id && candidate.level === "L0")) domainL0ByName.set(normalizeName(item.name), item.id);
  }
  for (const item of payload.dictionaries.domainL1) {
    if (current.domains.some((candidate) => candidate.id === item.id && (candidate.level ?? "L1") === "L1")) domainL1ByName.set(normalizeName(item.name), item.id);
  }
  const unclaimed = new Map(payload.rows.map((item) => [item.requirementId, item]));
  const rawRows: Array<{ row: number; visible: string[]; action: RequirementAction; rowToken?: string }> = [];
  for (let row = 2; row <= sheet.actualRowCount; row += 1) {
    const visible = visibleV8WorksheetRow(sheet, row, columns, formatVersion);
    if (!visible.some(Boolean)) continue;
    rawRows.push({
      row,
      visible,
      action: normalizeAction(visible[0], sheet.name, row, columns.get("操作")!, errors),
      rowToken: rowTokenFromWorksheetRow(sheet, row, offset),
    });
  }
  const matched = new Map<number, V6WorkbookRowMetadata>();
  const metadataByToken = new Map(payload.rows.filter((item) => item.rowToken).map((item) => [item.rowToken!.toLowerCase(), item]));
  const claim = (raw: (typeof rawRows)[number], item: V6WorkbookRowMetadata) => {
    matched.set(raw.row, item);
    unclaimed.delete(item.requirementId);
  };
  // Exact rows first. If duplicate business rows are identical, the style token
  // distinguishes them while remaining outside user-visible worksheet columns.
  for (const raw of rawRows) {
    const candidates = [...unclaimed.values()].filter((item) => item.baselineVisible.every((value, index) => value === raw.visible[index]));
    if (candidates.length === 1) {
      claim(raw, candidates[0]);
      continue;
    }
    const tokenCandidate = raw.rowToken ? metadataByToken.get(raw.rowToken) : undefined;
    if (tokenCandidate && candidates.some((item) => item.requirementId === tokenCandidate.requirementId)) claim(raw, tokenCandidate);
  }
  const weights = [0, 12, 8, 4, 6, 3, 2, 4, 6, 1, 1, 1, 3, 3, 3, 3, ...seriesForWorkbook.map(() => 1)];
  for (const raw of rawRows) {
    if (matched.has(raw.row)) continue;
    const candidates = [...unclaimed.values()].map((item) => {
      return { item, score: item.baselineVisible.reduce((sum, value, index) => sum + (index > 0 && value === raw.visible[index] ? weights[index] : 0), 0) };
    }).sort((a, b) => b.score - a.score);
    const tokenCandidate = raw.rowToken ? metadataByToken.get(raw.rowToken) : undefined;
    const availableTokenCandidate = tokenCandidate && unclaimed.has(tokenCandidate.requirementId) ? tokenCandidate : undefined;
    const tokenScore = availableTokenCandidate
      ? candidates.find((candidate) => candidate.item.requirementId === availableTokenCandidate.requirementId)?.score ?? -1
      : -1;
    // A uniquely better content match wins when a user reordered by copying cell
    // values without formats. Otherwise the row token resolves edited duplicates.
    if (candidates[0] && candidates[0].score >= 15 && candidates[0].score !== candidates[1]?.score && candidates[0].score > tokenScore) {
      claim(raw, candidates[0].item);
      continue;
    }
    if (availableTokenCandidate) {
      claim(raw, availableTokenCandidate);
      continue;
    }
    if (!candidates[0] || candidates[0].score < 15 || candidates[0].score === candidates[1]?.score) {
      addIssue(errors, sheet.name, `A${raw.row}`, "无法唯一匹配原始需求；不支持新增、复制行或同时大幅改动所有识别字段");
      continue;
    }
    claim(raw, candidates[0].item);
  }
  const result: V8ParsedRow[] = [];
  for (const raw of rawRows) {
    const metadata = matched.get(raw.row);
    if (!metadata?.baselineFields || !metadata.baselineSeries) continue;
    const existing = currentById.get(metadata.requirementId);
    if (!existing) {
      addIssue(errors, sheet.name, `A${raw.row}`, "原始需求在本机不存在，协作 Excel 不能创建新需求");
      continue;
    }
    const incomingFields = { ...metadata.baselineFields };
    const incomingSeries = { ...metadata.baselineSeries };
    if (raw.action !== "删除") {
      for (const key of V8_FIELD_KEYS) {
        const header = V8_FIELD_HEADERS[key];
        const column = columns.get(header)!;
        const value = raw.visible[V8_BUSINESS_HEADERS.findIndex((item) => item === header)];
        if (key === "domainL0Id" || key === "domainId") {
          const id = (key === "domainL0Id" ? domainL0ByName : domainL1ByName).get(normalizeName(value));
          if (!id) addIssue(errors, sheet.name, cellAddress(raw.row, column), `${header}“${value || "空"}”不存在于当前字典`);
          else incomingFields[key] = id;
        } else if (key === "planningMemberId" || key === "productManagerMemberId" || key === "seMemberId" || key === "algorithmMemberId") {
          const field = WORKBOOK_OWNER_FIELDS.find((item) => item.key === key)!;
          const candidates = (payload.dictionaries.members ?? []).filter((item) => item.role === field.role && item.label === value);
          const member = candidates.length === 1 ? current.members?.find((item) => item.id === candidates[0].id && item.role === field.role) : undefined;
          if (value && !member) addIssue(errors, sheet.name, cellAddress(raw.row, column), `${field.role}成员不存在、岗位不匹配或名称不唯一`);
          if (!value && field.required && metadata.baselineFields[key]) addIssue(errors, sheet.name, cellAddress(raw.row, column), `已配置的${field.role}不可清空`);
          incomingFields[key] = member?.id ?? "";
        } else if (key === "targetMonth") {
          const parsedMonth = parseMonth(sheet.getCell(raw.row, column).value);
          if (!parsedMonth && !(existing.targetMonth === "" && !value)) addIssue(errors, sheet.name, cellAddress(raw.row, column), "上线月份应为有效年月；历史空日期允许保留");
          incomingFields[key] = parsedMonth ?? "";
        } else if (key === "deviceWorkloadPm" || key === "appWorkloadPm" || key === "cloudWorkloadPm") {
          const number = parseNonNegativeNumber(value);
          if (number === null || number > 99999999) addIssue(errors, sheet.name, cellAddress(raw.row, column), "工作量应为非负数字，最多两位小数");
          incomingFields[key] = number === null ? value : String(roundWorkload(number));
        } else incomingFields[key] = value;
      }
      if (!incomingFields.title) addIssue(errors, sheet.name, cellAddress(raw.row, columns.get("需求标题")!), "需求标题不能为空");
      if (incomingFields.description.length > 5000) addIssue(errors, sheet.name, cellAddress(raw.row, columns.get("需求描述")!), "需求描述不能超过5000字");
      if (!REQUIREMENT_SCOPES.includes(incomingFields.scope as Requirement["scope"])) addIssue(errors, sheet.name, cellAddress(raw.row, columns.get("适用范围")!), "适用范围无效");
      if (!["是", "否"].includes(incomingFields.allowLatestVersionUpgrade)) addIssue(errors, sheet.name, cellAddress(raw.row, columns.get("允许最近版本升级")!), "请选择是或否");
      if (!CATEGORIES.includes(incomingFields.category as Requirement["category"])) addIssue(errors, sheet.name, cellAddress(raw.row, columns.get("分类")!), "分类无效");
      seriesForWorkbook.forEach((series, index) => {
        const value = raw.visible[16 + index];
        if (formatVersion === "v11" || metadata.baselineSeries?.[series.id]) {
          if (!["支持", "不支持", "未判断"].includes(value)) addIssue(errors, sheet.name, cellAddress(raw.row, columns.get(series.name)!), `${series.name}必须选择支持、不支持或未判断，不可清空`);
          incomingSeries[series.id] = value;
        } else if (value) addIssue(errors, sheet.name, cellAddress(raw.row, columns.get(series.name)!), "旧版表格此行未包含分层记录，系列单元格不可回写，请重新导出");
      });
    }
    result.push({ row: raw.row, requirementId: metadata.requirementId, action: raw.action, incomingFields, baselineFields: metadata.baselineFields, incomingSeries, baselineSeries: metadata.baselineSeries, baselineHash: metadata.baselineHash });
  }
  return result;
}

function matchV6RowMetadata(visibleValues: readonly string[], candidates: readonly V6WorkbookRowMetadata[]): V6WorkbookRowMetadata | undefined {
  const exact = candidates.filter((item) => item.baselineVisible.every((value, index) => value === visibleValues[index]));
  if (exact.length === 1) return exact[0];
  const weights = [0, 7, 5, 3, 4, 3, 2, 2, 3, 4, 3, 1, 1, 1, 2, 2, 2, 2];
  const scored = candidates.map((item) => ({
    item,
    score: item.baselineVisible.reduce((sum, value, index) => sum + (index > 0 && value === visibleValues[index] ? weights[index] : 0), 0),
  })).sort((a, b) => b.score - a.score);
  if (!scored[0] || scored[0].score < 7 || scored[0].score === scored[1]?.score) return undefined;
  return scored[0].item;
}

function parseOverrideSheet(
  sheet: ExcelJS.Worksheet,
  domains: ParsedDictionaryRow[],
  exportedAt: string,
  formatVersion: ParsedWorkspaceWorkbook["formatVersion"],
  errors: WorkspaceWorkbookIssue[],
): ParsedOverrideRow[] {
  const currentFormat = formatVersion === "v5";
  const headers = currentFormat ? OVERRIDE_HEADERS : formatVersion === "v4" ? V4_OVERRIDE_HEADERS : LEGACY_OVERRIDE_HEADERS;
  const columns = headerColumns(sheet, headers, errors);
  if (!columns) return [];
  const domainByName = dictionaryByName(domains);
  const domainById = new Map(domains.map((row) => [row.item.id, row.item]));
  const rows: ParsedOverrideRow[] = [];
  const keys = new Set<string>();
  for (let row = FIRST_DATA_ROW; row <= sheet.actualRowCount; row += 1) {
    const values = Object.fromEntries(headers.map((header) => [header, valueAt(sheet, row, columns, header)])) as Record<string, string>;
    if (!Object.values(values).some(Boolean)) continue;
    const action = currentFormat ? "保留" : normalizeAction(values["操作"], sheet.name, row, columns.get("操作")!, errors);
    const sortOrderValue = currentFormat || formatVersion === "v4" ? values["卡片顺序（请勿修改）"] : "";
    const parsedSortOrder = sortOrderValue === "" ? undefined : Number(sortOrderValue);
    if (parsedSortOrder !== undefined && (!Number.isInteger(parsedSortOrder) || parsedSortOrder < 0)) {
      addIssue(errors, sheet.name, cellAddress(row, columns.get("卡片顺序（请勿修改）")!), "卡片顺序必须是大于或等于0的整数，请不要修改");
    }
    const originalKey = normalizeLegacyGroupOverride({
      groupKey: values["卡片键（请勿修改）"],
      cardTitle: currentFormat ? values["领域 L1 标签（自动）"] || "兼容迁移" : values["卡片标题"] || "兼容迁移",
      cardSummary: currentFormat ? "" : values["详情/PPT摘要"],
      ...(parsedSortOrder === undefined ? {} : { sortOrder: parsedSortOrder }),
      updatedAt: values["基线更新时间（请勿修改）"] || exportedAt,
    }).groupKey;
    const baselineHash = values["基线校验（请勿修改）"];
    if (action === "删除") {
      if (!originalKey) addIssue(errors, sheet.name, cellAddress(row, columns.get("卡片键（请勿修改）")!), "新增卡片行不能标记删除");
      rows.push({ sheet: sheet.name, row, action, baselineHash, override: {
        groupKey: originalKey,
        cardTitle: "删除",
        cardSummary: "",
        ...(parsedSortOrder === undefined ? {} : { sortOrder: parsedSortOrder }),
        updatedAt: values["基线更新时间（请勿修改）"] || exportedAt,
      } });
      continue;
    }
    const titleHeader = currentFormat ? "领域 L1 标签（自动）" : "卡片标题";
    const title = values[titleHeader].trim();
    if (!title) addIssue(errors, sheet.name, cellAddress(row, columns.get(titleHeader)!), currentFormat ? "领域 L1 标签不能为空，请不要修改系统字段" : "卡片标题不能为空");
    let groupKey = originalKey;
    if (originalKey) {
      const [domainId, source, month] = originalKey.split("::");
      const originalDomain = domainById.get(domainId);
      const visibleRawSource = values["来源"];
      const visibleSource = normalizeLegacySource(visibleRawSource);
      const visibleMonth = parseMonth(sheet.getCell(row, columns.get("上线月份")!).value);
      if (!originalDomain || visibleSource !== source || visibleMonth !== month) {
        addIssue(errors, sheet.name, cellAddress(row, columns.get(titleHeader)!), "现有卡片的领域 L1 标签、来源和上线月份不可修改");
      }
    } else {
      if (currentFormat) {
        addIssue(errors, sheet.name, cellAddress(row, columns.get("卡片键（请勿修改）")!), "路标卡片系统行缺少卡片键，请重新导出协作 Excel");
      }
      const domainName = currentFormat ? values["领域 L1 标签（自动）"] : values["领域"];
      const domain = domainByName.get(normalizeName(domainName));
      const rawSource = values["来源"];
      const source = normalizeLegacySource(rawSource);
      const month = parseMonth(sheet.getCell(row, columns.get("上线月份")!).value);
      if (!domain) addIssue(errors, sheet.name, cellAddress(row, columns.get(titleHeader)!), "领域不存在于领域 L1 字典");
      if (!SOURCES.includes(source as (typeof SOURCES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get("来源")!), "来源不在允许范围内");
      if (!month) addIssue(errors, sheet.name, cellAddress(row, columns.get("上线月份")!), "上线月份应为有效年月");
      groupKey = `${domain?.id ?? ""}::${source}::${month ?? ""}`;
    }
    if (keys.has(groupKey)) addIssue(errors, sheet.name, cellAddress(row, columns.get("卡片键（请勿修改）")!), "路标卡片重复");
    keys.add(groupKey);
    const updatedAt = values["基线更新时间（请勿修改）"] || exportedAt;
    if (!validTimestamp(updatedAt)) addIssue(errors, sheet.name, cellAddress(row, columns.get("基线更新时间（请勿修改）")!), "系统时间字段无效，请不要修改");
    rows.push({ sheet: sheet.name, row, action, baselineHash, override: {
      groupKey,
      cardTitle: title,
      cardSummary: currentFormat ? "" : values["详情/PPT摘要"].trim(),
      ...(parsedSortOrder === undefined ? {} : { sortOrder: parsedSortOrder }),
      updatedAt,
    } });
  }
  return rows;
}

function planWorkspaceWorkbookImport(
  current: WorkspaceData,
  parsed: ParsedWorkspaceWorkbook,
  conflictMode: WorkspaceWorkbookConflictMode,
  now: string,
): ImportPlan {
  const counts: WorkspaceWorkbookImportCounts = {
    added: 0,
    updated: 0,
    deleted: 0,
    unchanged: 0,
    conflicts: 0,
    domainsChanged: 0,
    productsChanged: 0,
    groupOverridesChanged: 0,
  };
  const currentDomainL0s = current.domains.filter((item) => item.level === "L0");
  const currentDomainL1s = current.domains.filter((item) => (item.level ?? "L1") === "L1");
  const parsedDomainL0s = parsed.domains.filter((row) => row.item.level === "L0");
  const parsedDomainL1s = parsed.domains.filter((row) => (row.item.level ?? "L1") === "L1");
  const domainL0Merge = mergeDictionaryRows(currentDomainL0s, parsedDomainL0s, conflictMode, "领域 L0");
  const remappedDomainL1Rows = parsedDomainL1s.map((row) => ({
    ...row,
    item: {
      ...row.item,
      parentId: row.item.parentId ? (domainL0Merge.idMap.get(row.item.parentId) ?? row.item.parentId) : undefined,
    },
  }));
  const domainL1Merge = mergeDictionaryRows(currentDomainL1s, remappedDomainL1Rows, conflictMode, "领域 L1");
  const productMerge = mergeDictionaryRows(current.products, parsed.products, conflictMode, "产品");
  const baselineRequirementHash = parsed.formatVersion === "v2"
    ? requirementHashV2
    : parsed.formatVersion === "v7" ? requirementHashV7 : parsed.formatVersion === "v6" ? requirementHashV6 : requirementHash;
  counts.domainsChanged = domainL0Merge.changed + domainL1Merge.changed;
  counts.productsChanged = productMerge.changed;
  const conflicts = [...domainL0Merge.conflicts, ...domainL1Merge.conflicts, ...productMerge.conflicts];
  const requirements = new Map(current.requirements.map((item) => [item.id, item]));

  for (const row of parsed.requirements) {
    const existing = requirements.get(row.requirement.id);
    if (row.action === "删除") {
      if (!existing) {
        counts.unchanged += 1;
        continue;
      }
      const localChanged = Boolean(row.baselineHash) && !requirementBaselineMatches(existing, row.baselineHash, parsed.formatVersion);
      if (localChanged) {
        const issue = conflictIssue(row.sheet, row.row, "本机需求在 Excel 导出后已更新，删除操作发生冲突");
        conflicts.push(issue);
        if (conflictMode === "local-wins") {
          counts.unchanged += 1;
          continue;
        }
      }
      requirements.delete(existing.id);
      counts.deleted += 1;
      continue;
    }

    const incoming = {
      ...row.requirement,
      scope: parsed.formatVersion === "v2" && existing ? existing.scope : row.requirement.scope,
      allowLatestVersionUpgrade: ["v7", "v6"].includes(parsed.formatVersion)
        ? row.requirement.allowLatestVersionUpgrade
        : existing?.allowLatestVersionUpgrade ?? false,
      domainL0Id: domainL0Merge.idMap.get(row.requirement.domainL0Id) ?? row.requirement.domainL0Id,
      domainId: domainL1Merge.idMap.get(row.requirement.domainId) ?? row.requirement.domainId,
      productIds: [...new Set(row.requirement.productIds.map((id) => productMerge.idMap.get(id) ?? id))],
    };
    if (!existing) {
      if (["v7", "v6"].includes(parsed.formatVersion)) continue;
      requirements.set(incoming.id, { ...incoming, images: [], updatedAt: row.baselineHash ? incoming.updatedAt : now });
      counts.added += 1;
      continue;
    }
    if (row.baselineHash && requirementBaselineMatches(incoming, row.baselineHash, parsed.formatVersion)) {
      counts.unchanged += 1;
      continue;
    }
    const localChanged = Boolean(row.baselineHash) && !requirementBaselineMatches(existing, row.baselineHash, parsed.formatVersion);
    if (localChanged) {
      conflicts.push(conflictIssue(row.sheet, row.row, `需求“${incoming.title}”在本机和 Excel 中都已修改`));
      if (conflictMode === "local-wins") {
        counts.unchanged += 1;
        continue;
      }
    }
    requirements.set(incoming.id, {
      ...existing, ...incoming,
      tieringFeatureId: existing.tieringFeatureId,
      historicalLaunch: existing.historicalLaunch && incoming.targetMonth ? { ...existing.historicalLaunch, suggestedTargetMonth: incoming.targetMonth } : undefined,
      unallocatedWorkloadPm: existing.unallocatedWorkloadPm,
      researchAppWorkloadPm: existing.researchAppWorkloadPm ?? 0,
      algorithmWorkloadPm: existing.algorithmWorkloadPm ?? 0,
      medicalRoadmapSortOrder: existing.medicalRoadmapSortOrder,
      workloadPm: roundWorkload(incoming.deviceWorkloadPm + incoming.appWorkloadPm + incoming.cloudWorkloadPm + existing.unallocatedWorkloadPm + (existing.researchAppWorkloadPm ?? 0) + (existing.algorithmWorkloadPm ?? 0)),
      images: existing.images, createdAt: existing.createdAt, updatedAt: now,
      planningMemberId: parsed.formatVersion === "v7" ? incoming.planningMemberId : existing.planningMemberId,
      productManagerMemberId: parsed.formatVersion === "v7" ? incoming.productManagerMemberId : existing.productManagerMemberId,
      seMemberId: parsed.formatVersion === "v7" ? incoming.seMemberId : existing.seMemberId,
      algorithmMemberId: parsed.formatVersion === "v7" ? incoming.algorithmMemberId : existing.algorithmMemberId,
      medicalSubmissionMilestones: existing.medicalSubmissionMilestones
        ? { ...existing.medicalSubmissionMilestones }
        : undefined,
    });
    counts.updated += 1;
  }

  const overrides = new Map(current.groupOverrides.map((item) => [item.groupKey, item]));
  const baselineOverrideHash = parsed.formatVersion === "v5"
    ? overrideOrderHash
    : parsed.formatVersion === "v4" ? overrideHash : overrideHashV3;
  for (const row of parsed.overrides) {
    const [domainId, source, month] = row.override.groupKey.split("::");
    const groupKey = `${domainL1Merge.idMap.get(domainId) ?? domainId}::${source}::${month}`;
    const incoming = { ...row.override, groupKey };
    const existing = overrides.get(groupKey);
    if (row.action === "删除") {
      if (!existing) continue;
      const localChanged = Boolean(row.baselineHash) && baselineOverrideHash(existing) !== row.baselineHash;
      if (localChanged) {
        conflicts.push(conflictIssue(row.sheet, row.row, "本机卡片在 Excel 导出后已更新，删除操作发生冲突"));
        if (conflictMode === "local-wins") continue;
      }
      overrides.delete(groupKey);
      counts.groupOverridesChanged += 1;
      continue;
    }
    if (!existing) {
      overrides.set(groupKey, { ...incoming, updatedAt: row.baselineHash ? incoming.updatedAt : now });
      counts.groupOverridesChanged += 1;
      continue;
    }
    const incomingWithCompatibleOrder = parsed.formatVersion === "v5"
      ? { ...incoming, cardTitle: existing.cardTitle, cardSummary: existing.cardSummary }
      : parsed.formatVersion === "v4"
        ? incoming
        : { ...incoming, ...(existing.sortOrder === undefined ? {} : { sortOrder: existing.sortOrder }) };
    if (row.baselineHash && baselineOverrideHash(incomingWithCompatibleOrder) === row.baselineHash) continue;
    const localChanged = Boolean(row.baselineHash) && baselineOverrideHash(existing) !== row.baselineHash;
    if (localChanged) {
      conflicts.push(conflictIssue(row.sheet, row.row, `路标卡片“${incoming.cardTitle}”在本机和 Excel 中都已修改`));
      if (conflictMode === "local-wins") continue;
    }
    overrides.set(groupKey, { ...incomingWithCompatibleOrder, updatedAt: now });
    counts.groupOverridesChanged += 1;
  }

  counts.conflicts = conflicts.length;
  const nextDomains = [...domainL0Merge.items, ...domainL1Merge.items];
  const nextRequirements = [...requirements.values()];
  const tiering = normalizeTieringData(current.tiering);
  // Only explicit Excel deletions may remove normal specifications. Broken existing links stay intact.
  const deletedRequirementIds = new Set(current.requirements.filter((item) => !requirements.has(item.id)).map((item) => item.id));
  const removedFeatureIds = new Set(tiering.features.filter((item) => !isHistoricalTieringFeature(item) && item.requirementId && deletedRequirementIds.has(item.requirementId)).map((item) => item.id));
  tiering.features = tiering.features.filter((item) => !removedFeatureIds.has(item.id)).map((feature) => {
    if (feature.requirementId && deletedRequirementIds.has(feature.requirementId)) return { ...feature, requirementId: undefined, updatedAt: now };
    let updated = feature;
    for (const previous of current.domains) {
      const next = nextDomains.find((item) => item.id === previous.id);
      if (next) updated = renameIndependentTieringDomain(updated, previous, next, current.domains, current.requirements);
    }
    return resolveTieringFeature(updated, nextRequirements, nextDomains).feature;
  });
  tiering.supportRecords = tiering.supportRecords.filter((item) => !removedFeatureIds.has(item.featureId));
  tiering.changeLogs = tiering.changeLogs.filter((item) => !removedFeatureIds.has(item.featureId));
  tiering.verificationIssues = tiering.verificationIssues.filter((item) => !removedFeatureIds.has(item.featureId));
  return {
    data: normalizeWorkbookRoadmaps({
      ...current,
      requirements: nextRequirements,
      domains: nextDomains,
      products: productMerge.items,
      groupOverrides: current.roadmapAssignments !== undefined ? current.groupOverrides : [...overrides.values()],
      members: current.members,
      pptAccessPolicy: current.pptAccessPolicy,
      ...(current.tiering ? { tiering } : {}),
    }, current),
    counts,
    conflicts,
  };
}

function mergeDictionaryRows(
  current: DictionaryItem[],
  rows: ParsedDictionaryRow[],
  conflictMode: WorkspaceWorkbookConflictMode,
  label: string,
): DictionaryMergeResult {
  const items = new Map(current.map((item) => [item.id, { ...item }]));
  const idMap = new Map<string, string>();
  const conflicts: WorkspaceWorkbookIssue[] = [];
  let changed = 0;
  for (const row of rows) {
    const incoming = row.item;
    const sameId = items.get(incoming.id);
    const sameName = [...items.values()].find((item) => (item.level ?? "L1") === (incoming.level ?? "L1") && normalizeName(item.name) === normalizeName(incoming.name));
    if (!sameId && sameName) {
      idMap.set(incoming.id, sameName.id);
      if (!row.baselineHash && (sameName.active !== incoming.active || sameName.sortOrder !== incoming.sortOrder || sameName.parentId !== incoming.parentId)) {
        items.set(sameName.id, { ...sameName, active: incoming.active, sortOrder: incoming.sortOrder, parentId: incoming.parentId });
        changed += 1;
      }
      continue;
    }
    idMap.set(incoming.id, incoming.id);
    if (!sameId) {
      items.set(incoming.id, { ...incoming });
      changed += 1;
      continue;
    }
    if (row.baselineHash && dictionaryHash(incoming) === row.baselineHash) continue;
    const localChanged = Boolean(row.baselineHash) && dictionaryHash(sameId) !== row.baselineHash;
    if (localChanged) {
      conflicts.push(conflictIssue(row.sheet, row.row, `${label}“${incoming.name}”在本机和 Excel 中都已修改`));
      if (conflictMode === "local-wins") continue;
    }
    items.set(incoming.id, { ...incoming });
    changed += 1;
  }
  return { items: [...items.values()], idMap, changed, conflicts };
}

function validateDictionaryRows(rows: ParsedDictionaryRow[], label: string, errors: WorkspaceWorkbookIssue[]): void {
  const ids = new Set<string>();
  const names = new Set<string>();
  for (const row of rows) {
    const normalizedName = `${row.item.level ?? "L1"}:${normalizeName(row.item.name)}`;
    if (ids.has(row.item.id)) addIssue(errors, row.sheet, `${row.sheet === SHEETS.domainL1s ? "E" : "D"}${row.row}`, `${label}ID重复：${row.item.id}`);
    if (normalizedName && names.has(normalizedName)) addIssue(errors, row.sheet, `A${row.row}`, `${label}名称重复：${row.item.name}`);
    ids.add(row.item.id);
    if (normalizedName) names.add(normalizedName);
  }
}

function requireSheet(workbook: ExcelJS.Workbook, name: string, errors: WorkspaceWorkbookIssue[]): ExcelJS.Worksheet | undefined {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) addIssue(errors, name, "A1", `工作簿缺少“${name}”工作表`);
  return sheet;
}

function headerColumns<T extends readonly string[]>(sheet: ExcelJS.Worksheet, headers: T, errors: WorkspaceWorkbookIssue[]): Map<T[number], number> | null {
  return headerColumnsAt(sheet, headers, HEADER_ROW, errors);
}

function headerColumnsAt<T extends readonly string[]>(
  sheet: ExcelJS.Worksheet,
  headers: T,
  headerRow: number,
  errors: WorkspaceWorkbookIssue[],
): Map<T[number], number> | null {
  const found = new Map<T[number], number>();
  sheet.getRow(headerRow).eachCell({ includeEmpty: false }, (cell, column) => {
    const text = cellText(cell);
    if ((headers as readonly string[]).includes(text)) found.set(text as T[number], column);
  });
  for (const header of headers) {
    if (!found.has(header)) addIssue(errors, sheet.name, `${columnLetter(headers.indexOf(header) + 1)}${headerRow}`, `缺少必要列“${header}”`);
  }
  return found.size === headers.length ? found : null;
}

function valueAt<T extends string>(sheet: ExcelJS.Worksheet, row: number, columns: ReadonlyMap<T, number>, header: T): string {
  return cellText(sheet.getCell(row, columns.get(header)!));
}

function cellText(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value).trim();
  if (typeof value === "object") {
    if ("result" in value && value.result !== undefined && value.result !== null) return String(value.result).trim();
    if ("text" in value && typeof value.text === "string") return value.text.trim();
    if ("richText" in value && Array.isArray(value.richText)) return value.richText.map((item) => item.text).join("").trim();
  }
  return cell.text.trim();
}

function parseMonth(value: ExcelJS.CellValue): string | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}`;
  let text = "";
  if (typeof value === "object" && value && "result" in value) text = String(value.result ?? "");
  else text = String(value ?? "");
  const match = /^\s*(\d{4})\s*(?:-|\/|年)\s*(0?[1-9]|1[0-2])\s*(?:月)?\s*$/.exec(text);
  return match ? `${match[1]}-${match[2].padStart(2, "0")}` : null;
}

function monthDate(value: string): Date | null {
  const match = /^(\d{4})-(0[1-9]|1[0-2])$/.exec(value);
  return match ? new Date(Number(match[1]), Number(match[2]) - 1, 1, 12) : null;
}

function normalizeAction(value: string, sheet: string, row: number, column: number, errors: WorkspaceWorkbookIssue[]): RequirementAction {
  const action = value.trim() || "保留";
  if (action !== "保留" && action !== "删除") {
    addIssue(errors, sheet, cellAddress(row, column), "操作只能是“保留”或“删除”");
    return "保留";
  }
  return action;
}

function parseNonNegativeNumber(value: string): number | null {
  if (!value.trim()) return 0;
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : null;
}

function parseNonNegativeInteger(value: string): number | null {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed >= 0 ? parsed : null;
}

function splitList(value: string): string[] {
  return [...new Set(value.split(/[、,，;；\n]+/).map((item) => item.trim()).filter(Boolean))];
}

function dictionaryByName(rows: ParsedDictionaryRow[]): Map<string, DictionaryItem> {
  return new Map(rows.map((row) => [normalizeName(row.item.name), row.item]));
}

function deletedRequirementPlaceholder(id: string, timestamp: string): Requirement {
  return {
    id,
    title: "删除",
    description: "",
    images: [],
    domainL0Id: "",
    domainId: "",
    source: SOURCES[0],
    scope: "全球",
    allowLatestVersionUpgrade: false,
    overseasRegions: [],
    category: CATEGORIES[0],
    targetMonth: "2000-01",
    productIds: [],
    deviceWorkloadPm: 1,
    appWorkloadPm: 0,
    cloudWorkloadPm: 0,
    unallocatedWorkloadPm: 0,
    workloadPm: 1,
    createdAt: timestamp,
    updatedAt: timestamp,
  };
}

function requirementHash(item: Requirement): string {
  return stableHash({
    id: item.id,
    title: item.title.trim(),
    description: item.description.trim(),
    domainL0Id: item.domainL0Id,
    domainId: item.domainId,
    source: legacyWorkbookSource(item.source),
    scope: item.scope,
    category: item.category,
    targetMonth: item.targetMonth,
    productIds: [...new Set(item.productIds)].sort(),
    overseasRegions: [...new Set(item.overseasRegions)].sort(),
    deviceWorkloadPm: roundWorkload(item.deviceWorkloadPm),
    appWorkloadPm: roundWorkload(item.appWorkloadPm),
    cloudWorkloadPm: roundWorkload(item.cloudWorkloadPm),
    unallocatedWorkloadPm: roundWorkload(item.unallocatedWorkloadPm),
  });
}

function requirementHashV7(item: Requirement): string {
  return stableHash({ baseline: requirementHashV6(item, true), owners: WORKBOOK_OWNER_FIELDS.map(({ key }) => item[key] ?? "") });
}

/** Source removal is a schema migration, not a concurrent requirement edit. */
function requirementBaselineMatches(item: Requirement, hash: string, version: string): boolean {
  const hashOf = version === "v2" ? requirementHashV2 : ["v7", "v8", "v9", "v10", "v11"].includes(version) ? requirementHashV7 : version === "v6" ? requirementHashV6 : requirementHash;
  return hashOf(item) === hash || (item.source === undefined && version !== "v11" && SOURCES.some((source) => hashOf({ ...item, source }) === hash));
}

function requirementHashV6(item: Requirement, canonicalSource = false): string {
  return stableHash({
    id: item.id,
    title: item.title.trim(),
    description: item.description.trim(),
    domainL0Id: item.domainL0Id,
    domainId: item.domainId,
    source: canonicalSource ? item.source : item.source === "NMPA认证" ? "NMPA送检" : item.source === "CE认证" ? "CE送检" : item.source,
    scope: item.scope,
    allowLatestVersionUpgrade: item.allowLatestVersionUpgrade,
    category: item.category,
    targetMonth: item.targetMonth,
    productIds: [...new Set(item.productIds)].sort(),
    overseasRegions: [...new Set(item.overseasRegions)].sort(),
    deviceWorkloadPm: roundWorkload(item.deviceWorkloadPm),
    appWorkloadPm: roundWorkload(item.appWorkloadPm),
    cloudWorkloadPm: roundWorkload(item.cloudWorkloadPm),
    unallocatedWorkloadPm: roundWorkload(item.unallocatedWorkloadPm),
  });
}

function requirementHashV2(item: Requirement): string {
  return stableHash({
    id: item.id,
    title: item.title.trim(),
    description: item.description.trim(),
    domainL0Id: item.domainL0Id,
    domainId: item.domainId,
    source: legacyWorkbookSource(item.source),
    category: item.category,
    targetMonth: item.targetMonth,
    productIds: [...new Set(item.productIds)].sort(),
    overseasRegions: [...new Set(item.overseasRegions)].sort(),
    deviceWorkloadPm: roundWorkload(item.deviceWorkloadPm),
    appWorkloadPm: roundWorkload(item.appWorkloadPm),
    cloudWorkloadPm: roundWorkload(item.cloudWorkloadPm),
    unallocatedWorkloadPm: roundWorkload(item.unallocatedWorkloadPm),
  });
}

function legacyWorkbookSource(source: Requirement["source"]): string {
  return source === "NMPA认证" ? "NMPA送检" : source === "CE认证" ? "CE送检" : source ?? "";
}

function dictionaryHash(item: DictionaryItem): string {
  return stableHash({ id: item.id, name: item.name.trim(), sortOrder: item.sortOrder, active: item.active, level: item.level ?? null, parentId: item.parentId ?? null });
}

function overrideHash(item: GroupOverride): string {
  return stableHash({ groupKey: item.groupKey, cardTitle: item.cardTitle.trim(), cardSummary: item.cardSummary.trim(), sortOrder: item.sortOrder ?? null });
}

function overrideHashV3(item: GroupOverride): string {
  return stableHash({ groupKey: item.groupKey, cardTitle: item.cardTitle.trim(), cardSummary: item.cardSummary.trim() });
}

function overrideOrderHash(item: GroupOverride): string {
  return stableHash({ groupKey: item.groupKey, sortOrder: item.sortOrder ?? null });
}

function stableHash(value: unknown): string {
  return createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

function normalizeName(value: string): string {
  return value.trim().toLocaleLowerCase("zh-CN");
}

function validTimestamp(value: string): boolean {
  return Boolean(value) && Number.isFinite(Date.parse(value));
}

function addIssue(errors: WorkspaceWorkbookIssue[], sheet: string, cell: string, message: string): void {
  errors.push({ sheet, cell, message });
}

function conflictIssue(sheet: string, row: number, message: string): WorkspaceWorkbookIssue {
  return { sheet, cell: `A${row}`, message };
}

function cellAddress(row: number, column: number): string {
  return `${columnLetter(column)}${row}`;
}

function columnLetter(column: number): string {
  let value = column;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function styleHeader(sheet: ExcelJS.Worksheet, range: string): void {
  forEachCell(sheet, range, (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF39424A" } };
    cell.font = { name: "Microsoft YaHei", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { vertical: "middle" };
  });
}

function styleEditable(sheet: ExcelJS.Worksheet, range: string): void {
  forEachCell(sheet, range, (cell) => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFAE8" } }; });
}

function styleSystem(sheet: ExcelJS.Worksheet, range: string): void {
  forEachCell(sheet, range, (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF1F3F5" } };
    cell.font = { name: "Consolas", size: 9, color: { argb: "FF69737B" } };
  });
}

function writeMatrix(sheet: ExcelJS.Worksheet, startRow: number, startColumn: number, values: Array<Array<string | number>>): void {
  values.forEach((rowValues, rowOffset) => {
    rowValues.forEach((value, columnOffset) => {
      sheet.getCell(startRow + rowOffset, startColumn + columnOffset).value = value;
    });
  });
}

function forEachCell(sheet: ExcelJS.Worksheet, range: string, action: (cell: ExcelJS.Cell) => void): void {
  const [start, end] = range.split(":");
  const startCell = sheet.getCell(start);
  const endCell = sheet.getCell(end);
  for (let row = startCell.row; row <= endCell.row; row += 1) {
    for (let column = startCell.col; column <= endCell.col; column += 1) action(sheet.getCell(row, column));
  }
}

function applyListValidation(sheet: ExcelJS.Worksheet, range: string, values: string[], allowBlank = false): void {
  // Excel only accepts a nonempty, comma-separated literal list up to 255 characters.
  // Without a hidden dictionary sheet, an oversized list cannot be represented as a
  // dropdown: leave those cells editable and validate their contents on import.
  if (!values.length || values.some((value) => !value || /[,"\r\n]/.test(value))) return;
  const formula = `"${values.join(",")}"`;
  if (formula.length > 255) return;
  const allowedValues = `允许值：${values.join("、")}`;
  const [start, end] = range.split(":");
  const startCell = sheet.getCell(start);
  const endCell = sheet.getCell(end);
  for (let row = startCell.row; row <= endCell.row; row += 1) {
    for (let column = startCell.col; column <= endCell.col; column += 1) {
      sheet.getCell(row, column).dataValidation = {
        type: "list",
        allowBlank,
        formulae: [formula],
        showErrorMessage: true,
        errorTitle: "请选择有效值",
        error: allowedValues.length <= 225 ? allowedValues : "请填写现有选项，导入时将再次校验。",
      };
    }
  }
}

function applyDecimalValidation(sheet: ExcelJS.Worksheet, range: string, minimum: number, allowBlank = true): void {
  const [start, end] = range.split(":");
  const startCell = sheet.getCell(start);
  const endCell = sheet.getCell(end);
  for (let row = startCell.row; row <= endCell.row; row += 1) {
    sheet.getCell(row, startCell.col).dataValidation = {
      type: "decimal",
      operator: "greaterThanOrEqual",
      allowBlank,
      formulae: [minimum],
      showErrorMessage: true,
      errorTitle: "工作量格式错误",
      error: "请输入大于或等于0的整数或小数。",
    };
  }
}

function applyWholeNumberValidation(sheet: ExcelJS.Worksheet, range: string, minimum: number): void {
  const [start, end] = range.split(":");
  const startCell = sheet.getCell(start);
  const endCell = sheet.getCell(end);
  for (let row = startCell.row; row <= endCell.row; row += 1) {
    sheet.getCell(row, startCell.col).dataValidation = {
      type: "whole",
      operator: "greaterThanOrEqual",
      allowBlank: false,
      formulae: [minimum],
      showErrorMessage: true,
      errorTitle: "排序格式错误",
      error: "请输入大于或等于0的整数。",
    };
  }
}

function lightBorders(): Partial<ExcelJS.Borders> {
  const border: Partial<ExcelJS.Border> = { style: "thin", color: { argb: "FFE2E7EA" } };
  return { top: border, left: border, bottom: border, right: border };
}
