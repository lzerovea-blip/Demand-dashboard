import { createHash } from "node:crypto";
import ExcelJS from "exceljs";
import type {
  DictionaryItem,
  DomainLevel,
  GroupOverride,
  Requirement,
  WorkspaceData,
  WorkspaceWorkbookConflictMode,
  WorkspaceWorkbookImportCounts,
  WorkspaceWorkbookImportPreview,
  WorkspaceWorkbookIssue,
} from "../shared/types.js";
import { CATEGORIES, SOURCES } from "../shared/types.js";
import { normalizeLegacyGroupOverride, normalizeLegacySource } from "../shared/requirements.js";
import { roundWorkload } from "../shared/workload.js";
import { validateWorkspaceData } from "./workspacePackage.js";

const FORMAT_KEY = "health-roadmap-collaboration";
const FORMAT_VERSION = "v2";
const MAX_WORKBOOK_BYTES = 25 * 1024 * 1024;
const HEADER_ROW = 3;
const FIRST_DATA_ROW = HEADER_ROW + 1;
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
const OVERRIDE_HEADERS = [
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
  exportedAt: string;
  domains: ParsedDictionaryRow[];
  products: ParsedDictionaryRow[];
  requirements: ParsedRequirementRow[];
  overrides: ParsedOverrideRow[];
  errors: WorkspaceWorkbookIssue[];
}

export interface PreparedWorkspaceWorkbookImport {
  parsed: ParsedWorkspaceWorkbook;
  preview: Omit<WorkspaceWorkbookImportPreview, "token" | "fileName">;
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
): Promise<Buffer> {
  validateWorkspaceData(data);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "解决方案需求管理";
  workbook.created = new Date(exportedAt);
  workbook.modified = new Date(exportedAt);
  workbook.calcProperties.fullCalcOnLoad = true;

  createGuideSheet(workbook, data, appVersion, exportedAt);
  createDictionarySheet(workbook, SHEETS.domainL0s, "DomainL0Table", data.domains.filter((item) => item.level === "L0"));
  createDomainL1Sheet(workbook, data);
  createDictionarySheet(workbook, SHEETS.products, "ProductsTable", data.products);
  createRequirementSheet(workbook, data);
  createOverrideSheet(workbook, data);

  const bytes = await workbook.xlsx.writeBuffer();
  return Buffer.from(bytes);
}

export async function inspectWorkspaceWorkbook(
  bytes: Buffer,
  fileName: string,
  current: WorkspaceData,
): Promise<PreparedWorkspaceWorkbookImport> {
  if (!fileName.toLowerCase().endsWith(".xlsx")) throw new Error("请选择 .xlsx 协作工作簿");
  if (!bytes.length || bytes.length > MAX_WORKBOOK_BYTES) throw new Error("Excel 文件为空或超过 25MB 限制");
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(bytes as never);
  } catch {
    throw new Error("Excel 文件损坏或不是有效的 .xlsx 工作簿");
  }

  const guide = workbook.getWorksheet(SHEETS.guide);
  if (!guide) throw new Error(`工作簿缺少“${SHEETS.guide}”工作表`);
  if (cellText(guide.getCell("B13")) !== FORMAT_KEY || cellText(guide.getCell("B14")) !== FORMAT_VERSION) {
    throw new Error("不支持的协作 Excel 格式或版本");
  }
  const exportedAt = cellText(guide.getCell("B15"));
  if (!validTimestamp(exportedAt)) throw new Error("协作 Excel 缺少有效的导出时间");

  const errors: WorkspaceWorkbookIssue[] = [];
  const domainL0Sheet = requireSheet(workbook, SHEETS.domainL0s, errors);
  const domainL1Sheet = requireSheet(workbook, SHEETS.domainL1s, errors);
  const productSheet = requireSheet(workbook, SHEETS.products, errors);
  const requirementSheet = requireSheet(workbook, SHEETS.requirements, errors);
  const overrideSheet = requireSheet(workbook, SHEETS.overrides, errors);

  const domainL0s = domainL0Sheet ? parseDictionarySheet(domainL0Sheet, errors, "L0") : [];
  const domainL1s = domainL1Sheet ? parseDomainL1Sheet(domainL1Sheet, domainL0s, errors) : [];
  const domains = [...domainL0s, ...domainL1s];
  const products = productSheet ? parseDictionarySheet(productSheet, errors) : [];
  validateDictionaryRows(domains, "领域", errors);
  validateDictionaryRows(products, "产品", errors);
  const requirements = requirementSheet ? parseRequirementSheet(requirementSheet, domainL0s, domainL1s, products, exportedAt, errors) : [];
  const overrides = overrideSheet ? parseOverrideSheet(overrideSheet, domainL1s, exportedAt, errors) : [];
  const parsed: ParsedWorkspaceWorkbook = { exportedAt, domains, products, requirements, overrides, errors };
  const plan = planWorkspaceWorkbookImport(current, parsed, "local-wins", new Date().toISOString());
  try {
    validateWorkspaceData(plan.data);
  } catch (error) {
    addIssue(errors, SHEETS.guide, "A1", error instanceof Error ? error.message : String(error));
  }

  return {
    parsed,
    preview: {
      formatVersion: FORMAT_VERSION,
      exportedAt,
      counts: plan.counts,
      errors,
      conflicts: plan.conflicts,
    },
  };
}

export function applyPreparedWorkspaceWorkbook(
  current: WorkspaceData,
  prepared: PreparedWorkspaceWorkbookImport,
  conflictMode: WorkspaceWorkbookConflictMode,
  now = new Date().toISOString(),
): WorkspaceData {
  if (prepared.parsed.errors.length) throw new Error("协作 Excel 存在校验错误，请修正后重新导入");
  const plan = planWorkspaceWorkbookImport(current, prepared.parsed, conflictMode, now);
  validateWorkspaceData(plan.data);
  return plan.data;
}

function createGuideSheet(workbook: ExcelJS.Workbook, data: WorkspaceData, appVersion: string, exportedAt: string): void {
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
    [1, "导出", "工作台导出全部结构化数据，不受当前筛选和半年区间影响。"],
    [2, "协作", "将本文件上传到企业已有在线协作平台，多人共同编辑同一份文件。"],
    [3, "下载", "协作完成后下载为标准 .xlsx 文件，不要转换成 CSV。"],
    [4, "导入", "工作台先检查新增、修改、删除、冲突和错误，确认后再合并。"],
    ["填写规则", "删除需求", "在“需求清单”的操作列选择“删除”；直接删除整行不会删除应用中的需求。"],
    ["填写规则", "领域层级", "领域 L1 必须关联一个上游 L0；需求中的 L1 必须属于所选 L0。路标卡片只使用领域 L1。"],
    ["填写规则", "匹配产品", "仅产品专属需求填写，且只能选择一个产品。"],
    ["填写规则", "需求来源", `来源必须从下拉列表选择：${SOURCES.join("、")}。`],
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
  forEachCell(sheet, "A13:B17", (cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F4F6" } };
    cell.border = lightBorders();
  });
  forEachCell(sheet, "A13:A17", (cell) => { cell.font = { name: "Microsoft YaHei", size: 9, bold: true, color: { argb: "FF66717B" } }; });
  forEachCell(sheet, "B13:B17", (cell) => { cell.font = { name: "Consolas", size: 9, color: { argb: "FF36414A" } }; });
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
  prepareDataSheet(sheet, "需求清单 · 黄色字段可协作编辑，灰色系统字段请勿修改", REQUIREMENT_HEADERS.length, 2);
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
      item.category,
      monthDate(item.targetMonth),
      productNames[0] ?? null,
      item.deviceWorkloadPm,
      item.appWorkloadPm,
      item.cloudWorkloadPm,
      item.unallocatedWorkloadPm,
      { formula: `SUM(J${rowNumber}:M${rowNumber})`, result: item.workloadPm },
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

  const widths = [10, 28, 48, 20, 20, 16, 14, 14, 28, 16, 16, 16, 19, 19, 16, 40, 40, 40, 40, 24, 24, 32, 28, 28, 68];
  sheet.columns = widths.map((width) => ({ width }));
  sheet.getColumn(8).numFmt = "yyyy-mm";
  for (const column of [10, 11, 12, 13, 14]) sheet.getColumn(column).numFmt = "0.00";
  const usedLast = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + rows.length - 1);
  const validationLast = Math.max(usedLast + 30, 34);
  applyListValidation(sheet, `A${FIRST_DATA_ROW}:A${validationLast}`, ["保留", "删除"]);
  applyListValidation(sheet, `F${FIRST_DATA_ROW}:F${validationLast}`, [...SOURCES]);
  applyListValidation(sheet, `G${FIRST_DATA_ROW}:G${validationLast}`, [...CATEGORIES]);
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
    sheet.getCell(row, 9).dataValidation = {
      type: "list",
      allowBlank: true,
      formulae: [`'${SHEETS.products}'!$A$${FIRST_DATA_ROW}:$A$${productLast}`],
      showErrorMessage: true,
      errorTitle: "请选择一个产品",
      error: "匹配产品必须来自产品字典，且只能选择一个。",
    };
  }
  for (const column of [10, 11, 12]) applyDecimalValidation(sheet, `${columnLetter(column)}${FIRST_DATA_ROW}:${columnLetter(column)}${validationLast}`, 0, false);
  styleEditable(sheet, `A${FIRST_DATA_ROW}:L${usedLast}`);
  styleSystem(sheet, `M${FIRST_DATA_ROW}:Y${usedLast}`);
  sheet.getColumn(3).alignment = { wrapText: true, vertical: "top" };
}

function createOverrideSheet(workbook: ExcelJS.Workbook, data: WorkspaceData): void {
  const sheet = workbook.addWorksheet(SHEETS.overrides, { properties: { tabColor: { argb: "FF2F9D82" } } });
  prepareDataSheet(sheet, "路标卡片 · 标题和摘要可编辑，现有卡片的领域/来源/月不可修改", OVERRIDE_HEADERS.length);
  const domains = new Map(data.domains.map((item) => [item.id, item.name]));
  const rows = data.groupOverrides.map((item) => {
    const [domainId, source, targetMonth] = item.groupKey.split("::");
    return ["保留", item.cardTitle, item.cardSummary || null, domains.get(domainId) ?? domainId, source, monthDate(targetMonth), item.groupKey, item.updatedAt, overrideHash(item)];
  });
  addTable(sheet, "RoadmapCardsTable", OVERRIDE_HEADERS, rows.length ? rows : [blankCells(OVERRIDE_HEADERS.length)]);
  sheet.columns = [{ width: 10 }, { width: 28 }, { width: 48 }, { width: 20 }, { width: 16 }, { width: 14 }, { width: 58 }, { width: 28 }, { width: 68 }];
  sheet.getColumn(6).numFmt = "yyyy-mm";
  const usedLast = Math.max(FIRST_DATA_ROW, FIRST_DATA_ROW + rows.length - 1);
  const validationLast = Math.max(usedLast + 20, 24);
  applyListValidation(sheet, `A${FIRST_DATA_ROW}:A${validationLast}`, ["保留", "删除"]);
  applyListValidation(sheet, `E${FIRST_DATA_ROW}:E${validationLast}`, [...SOURCES]);
  styleEditable(sheet, `A${FIRST_DATA_ROW}:C${usedLast}`);
  styleSystem(sheet, `D${FIRST_DATA_ROW}:I${usedLast}`);
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
  errors: WorkspaceWorkbookIssue[],
): ParsedRequirementRow[] {
  const columns = headerColumns(sheet, REQUIREMENT_HEADERS, errors);
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
    const values = Object.fromEntries(REQUIREMENT_HEADERS.map((header) => [header, valueAt(sheet, row, columns, header)])) as Record<(typeof REQUIREMENT_HEADERS)[number], string>;
    if (!Object.values(values).some(Boolean)) continue;
    const action = normalizeAction(values[REQUIREMENT_HEADERS[0]], sheet.name, row, columns.get(REQUIREMENT_HEADERS[0])!, errors);
    const rawId = values[REQUIREMENT_HEADERS[15]];
    const baselineHash = values[REQUIREMENT_HEADERS[24]];
    if (action === "删除") {
      if (!rawId) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[15])!), "新增行不能标记删除");
      rows.push({ sheet: sheet.name, row, action, baselineHash, requirement: deletedRequirementPlaceholder(rawId, exportedAt) });
      continue;
    }

    const title = values[REQUIREMENT_HEADERS[1]].trim();
    const description = values[REQUIREMENT_HEADERS[2]].trim();
    if (!title) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[1])!), "需求标题不能为空");
    if (description.length > 5000) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[2])!), "需求描述不能超过5000字");
    if (!rawId && baselineHash) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[15])!), "系统ID缺失，请不要修改系统字段");
    const id = rawId || crypto.randomUUID();
    if (usedIds.has(id)) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[15])!), `需求ID重复：${id}`);
    usedIds.add(id);

    const domainL0Name = values[REQUIREMENT_HEADERS[3]].trim();
    const baselineDomainL0Id = values[REQUIREMENT_HEADERS[16]];
    const baselineDomainL0Name = values[REQUIREMENT_HEADERS[19]].trim();
    const domainL0 = domainL0ByName.get(normalizeName(domainL0Name))
      ?? (domainL0Name === baselineDomainL0Name ? domainL0ById.get(baselineDomainL0Id) : undefined);
    if (!domainL0) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[3])!), `领域 L0“${domainL0Name || "空"}”不存在于领域 L0 字典`);

    const domainL1Name = values[REQUIREMENT_HEADERS[4]].trim();
    const baselineDomainL1Id = values[REQUIREMENT_HEADERS[17]];
    const baselineDomainL1Name = values[REQUIREMENT_HEADERS[20]].trim();
    const domainL1 = domainL1ByName.get(normalizeName(domainL1Name))
      ?? (domainL1Name === baselineDomainL1Name ? domainL1ById.get(baselineDomainL1Id) : undefined);
    if (!domainL1) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[4])!), `领域 L1“${domainL1Name || "空"}”不存在于领域 L1 字典`);
    if (domainL0 && domainL1 && domainL1.parentId !== domainL0.id) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[4])!), `领域 L1“${domainL1Name}”不属于领域 L0“${domainL0Name}”`);

    const rawSource = values[REQUIREMENT_HEADERS[5]];
    const source = normalizeLegacySource(rawSource);
    const category = values[REQUIREMENT_HEADERS[6]];
    if (!SOURCES.includes(source as (typeof SOURCES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[5])!), "来源不在允许范围内");
    if (!CATEGORIES.includes(category as (typeof CATEGORIES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[6])!), "分类不在允许范围内");
    const targetMonth = parseMonth(sheet.getCell(row, columns.get(REQUIREMENT_HEADERS[7])!).value);
    if (!targetMonth) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[7])!), "上线月份应为有效年月");

    const productNames = splitList(values[REQUIREMENT_HEADERS[8]]);
    const baselineProductIds = splitList(values[REQUIREMENT_HEADERS[18]]);
    const baselineProductNames = splitList(values[REQUIREMENT_HEADERS[21]]);
    const productIds: string[] = [];
    if (productNames.length > 1) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[8])!), "匹配产品只能选择一个");
    productNames.forEach((name) => {
      const direct = productByName.get(normalizeName(name));
      const baselineIndex = baselineProductNames.findIndex((item) => normalizeName(item) === normalizeName(name));
      const fallback = baselineIndex >= 0 ? productById.get(baselineProductIds[baselineIndex] ?? "") : undefined;
      const product = direct ?? fallback;
      if (!product) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[8])!), `产品“${name}”不存在于产品字典`);
      else productIds.push(product.id);
    });
    if (category === "产品专属" && productIds.length !== 1) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[8])!), "产品专属需求必须且只能匹配一个产品");
    if (category === "体验优化" && productIds.length > 0) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[8])!), "体验优化需求不需要匹配产品");

    const device = parseRequiredNonNegativeNumber(values[REQUIREMENT_HEADERS[9]]);
    const app = parseRequiredNonNegativeNumber(values[REQUIREMENT_HEADERS[10]]);
    const cloud = parseRequiredNonNegativeNumber(values[REQUIREMENT_HEADERS[11]]);
    const unallocated = parseNonNegativeNumber(values[REQUIREMENT_HEADERS[12]]);
    [device, app, cloud].forEach((value, index) => {
      if (value === null) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[9 + index])!), "三端工作量均为必填，请输入大于或等于0的数字");
    });
    if (unallocated === null) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[12])!), "待拆分工作量必须是大于或等于0的数字");
    const total = roundWorkload((device ?? 0) + (app ?? 0) + (cloud ?? 0) + (unallocated ?? 0));
    if (total <= 0) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[9])!), "设备、App、云侧和待拆分工作量合计必须大于0");

    const createdAt = values[REQUIREMENT_HEADERS[22]] || exportedAt;
    const updatedAt = values[REQUIREMENT_HEADERS[23]] || exportedAt;
    if (!validTimestamp(createdAt) || !validTimestamp(updatedAt)) addIssue(errors, sheet.name, cellAddress(row, columns.get(REQUIREMENT_HEADERS[23])!), "系统时间字段无效，请不要修改");
    const requirement: Requirement = {
      id,
      title,
      description,
      images: [],
      domainL0Id: domainL0?.id ?? baselineDomainL0Id,
      domainId: domainL1?.id ?? baselineDomainL1Id,
      source: source as Requirement["source"],
      category: category as Requirement["category"],
      targetMonth: targetMonth ?? "",
      productIds: [...new Set(productIds)],
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

function parseOverrideSheet(
  sheet: ExcelJS.Worksheet,
  domains: ParsedDictionaryRow[],
  exportedAt: string,
  errors: WorkspaceWorkbookIssue[],
): ParsedOverrideRow[] {
  const columns = headerColumns(sheet, OVERRIDE_HEADERS, errors);
  if (!columns) return [];
  const domainByName = dictionaryByName(domains);
  const domainById = new Map(domains.map((row) => [row.item.id, row.item]));
  const rows: ParsedOverrideRow[] = [];
  const keys = new Set<string>();
  for (let row = FIRST_DATA_ROW; row <= sheet.actualRowCount; row += 1) {
    const values = Object.fromEntries(OVERRIDE_HEADERS.map((header) => [header, valueAt(sheet, row, columns, header)])) as Record<(typeof OVERRIDE_HEADERS)[number], string>;
    if (!Object.values(values).some(Boolean)) continue;
    const action = normalizeAction(values[OVERRIDE_HEADERS[0]], sheet.name, row, columns.get(OVERRIDE_HEADERS[0])!, errors);
    const originalKey = normalizeLegacyGroupOverride({
      groupKey: values[OVERRIDE_HEADERS[6]],
      cardTitle: values[OVERRIDE_HEADERS[1]] || "兼容迁移",
      cardSummary: values[OVERRIDE_HEADERS[2]],
      updatedAt: values[OVERRIDE_HEADERS[7]] || exportedAt,
    }).groupKey;
    const baselineHash = values[OVERRIDE_HEADERS[8]];
    if (action === "删除") {
      if (!originalKey) addIssue(errors, sheet.name, cellAddress(row, columns.get(OVERRIDE_HEADERS[6])!), "新增卡片行不能标记删除");
      rows.push({ sheet: sheet.name, row, action, baselineHash, override: { groupKey: originalKey, cardTitle: "删除", cardSummary: "", updatedAt: values[OVERRIDE_HEADERS[7]] || exportedAt } });
      continue;
    }
    const title = values[OVERRIDE_HEADERS[1]].trim();
    if (!title) addIssue(errors, sheet.name, cellAddress(row, columns.get(OVERRIDE_HEADERS[1])!), "卡片标题不能为空");
    let groupKey = originalKey;
    if (originalKey) {
      const [domainId, source, month] = originalKey.split("::");
      const originalDomain = domainById.get(domainId);
      const visibleRawSource = values[OVERRIDE_HEADERS[4]];
      const visibleSource = normalizeLegacySource(visibleRawSource);
      const visibleMonth = parseMonth(sheet.getCell(row, columns.get(OVERRIDE_HEADERS[5])!).value);
      if (!originalDomain || visibleSource !== source || visibleMonth !== month) {
        addIssue(errors, sheet.name, cellAddress(row, columns.get(OVERRIDE_HEADERS[3])!), "现有卡片的领域、来源和上线月份不可修改");
      }
    } else {
      const domain = domainByName.get(normalizeName(values[OVERRIDE_HEADERS[3]]));
      const rawSource = values[OVERRIDE_HEADERS[4]];
      const source = normalizeLegacySource(rawSource);
      const month = parseMonth(sheet.getCell(row, columns.get(OVERRIDE_HEADERS[5])!).value);
      if (!domain) addIssue(errors, sheet.name, cellAddress(row, columns.get(OVERRIDE_HEADERS[3])!), "领域不存在于领域字典");
      if (!SOURCES.includes(source as (typeof SOURCES)[number])) addIssue(errors, sheet.name, cellAddress(row, columns.get(OVERRIDE_HEADERS[4])!), "来源不在允许范围内");
      if (!month) addIssue(errors, sheet.name, cellAddress(row, columns.get(OVERRIDE_HEADERS[5])!), "上线月份应为有效年月");
      groupKey = `${domain?.id ?? ""}::${source}::${month ?? ""}`;
    }
    if (keys.has(groupKey)) addIssue(errors, sheet.name, cellAddress(row, columns.get(OVERRIDE_HEADERS[6])!), "路标卡片重复");
    keys.add(groupKey);
    const updatedAt = values[OVERRIDE_HEADERS[7]] || exportedAt;
    if (!validTimestamp(updatedAt)) addIssue(errors, sheet.name, cellAddress(row, columns.get(OVERRIDE_HEADERS[7])!), "系统时间字段无效，请不要修改");
    rows.push({ sheet: sheet.name, row, action, baselineHash, override: { groupKey, cardTitle: title, cardSummary: values[OVERRIDE_HEADERS[2]].trim(), updatedAt } });
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
      const localChanged = Boolean(row.baselineHash) && requirementHash(existing) !== row.baselineHash;
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
      domainL0Id: domainL0Merge.idMap.get(row.requirement.domainL0Id) ?? row.requirement.domainL0Id,
      domainId: domainL1Merge.idMap.get(row.requirement.domainId) ?? row.requirement.domainId,
      productIds: [...new Set(row.requirement.productIds.map((id) => productMerge.idMap.get(id) ?? id))],
    };
    if (!existing) {
      requirements.set(incoming.id, { ...incoming, images: [], updatedAt: row.baselineHash ? incoming.updatedAt : now });
      counts.added += 1;
      continue;
    }
    if (row.baselineHash && requirementHash(incoming) === row.baselineHash) {
      counts.unchanged += 1;
      continue;
    }
    const localChanged = Boolean(row.baselineHash) && requirementHash(existing) !== row.baselineHash;
    if (localChanged) {
      conflicts.push(conflictIssue(row.sheet, row.row, `需求“${incoming.title}”在本机和 Excel 中都已修改`));
      if (conflictMode === "local-wins") {
        counts.unchanged += 1;
        continue;
      }
    }
    requirements.set(incoming.id, { ...incoming, images: existing.images, createdAt: existing.createdAt, updatedAt: now });
    counts.updated += 1;
  }

  const overrides = new Map(current.groupOverrides.map((item) => [item.groupKey, item]));
  for (const row of parsed.overrides) {
    const [domainId, source, month] = row.override.groupKey.split("::");
    const groupKey = `${domainL1Merge.idMap.get(domainId) ?? domainId}::${source}::${month}`;
    const incoming = { ...row.override, groupKey };
    const existing = overrides.get(groupKey);
    if (row.action === "删除") {
      if (!existing) continue;
      const localChanged = Boolean(row.baselineHash) && overrideHash(existing) !== row.baselineHash;
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
    if (row.baselineHash && overrideHash(incoming) === row.baselineHash) continue;
    const localChanged = Boolean(row.baselineHash) && overrideHash(existing) !== row.baselineHash;
    if (localChanged) {
      conflicts.push(conflictIssue(row.sheet, row.row, `路标卡片“${incoming.cardTitle}”在本机和 Excel 中都已修改`));
      if (conflictMode === "local-wins") continue;
    }
    overrides.set(groupKey, { ...incoming, updatedAt: now });
    counts.groupOverridesChanged += 1;
  }

  counts.conflicts = conflicts.length;
  return {
    data: {
      requirements: [...requirements.values()],
      domains: [...domainL0Merge.items, ...domainL1Merge.items],
      products: productMerge.items,
      groupOverrides: [...overrides.values()],
    },
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
  const found = new Map<T[number], number>();
  sheet.getRow(HEADER_ROW).eachCell({ includeEmpty: false }, (cell, column) => {
    const text = cellText(cell);
    if ((headers as readonly string[]).includes(text)) found.set(text as T[number], column);
  });
  for (const header of headers) {
    if (!found.has(header)) addIssue(errors, sheet.name, `${columnLetter(headers.indexOf(header) + 1)}${HEADER_ROW}`, `缺少必要列“${header}”`);
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

function parseRequiredNonNegativeNumber(value: string): number | null {
  if (!value.trim()) return null;
  return parseNonNegativeNumber(value);
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
    source: item.source,
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

function dictionaryHash(item: DictionaryItem): string {
  return stableHash({ id: item.id, name: item.name.trim(), sortOrder: item.sortOrder, active: item.active, level: item.level ?? null, parentId: item.parentId ?? null });
}

function overrideHash(item: GroupOverride): string {
  return stableHash({ groupKey: item.groupKey, cardTitle: item.cardTitle.trim(), cardSummary: item.cardSummary.trim() });
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

function applyListValidation(sheet: ExcelJS.Worksheet, range: string, values: string[]): void {
  const [start, end] = range.split(":");
  const startCell = sheet.getCell(start);
  const endCell = sheet.getCell(end);
  for (let row = startCell.row; row <= endCell.row; row += 1) {
    for (let column = startCell.col; column <= endCell.col; column += 1) {
      sheet.getCell(row, column).dataValidation = {
        type: "list",
        allowBlank: false,
        formulae: [`"${values.join(",")}"`],
        showErrorMessage: true,
        errorTitle: "请选择有效值",
        error: `允许值：${values.join("、")}`,
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
