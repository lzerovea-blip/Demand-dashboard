import ExcelJS from "exceljs";
import type {
  HalfYearRange,
  Member,
  Requirement,
  TieringEvidenceSource,
  TieringFeature,
  TieringSupportRecord,
  WorkspaceData,
} from "../shared/types.js";
import { isValidRequirementMonth } from "../shared/medicalSubmission.js";
import { halfYearSequence, hmosVersionForMonth, requirementsInRange, shortHalfYear } from "../shared/roadmap.js";
import {
  BUILT_IN_TIERING_PRODUCT_SERIES,
  isTieringEligibleRequirement,
  tieringFeatureIdForRequirement,
} from "../shared/tiering.js";

const MATRIX_SHEET = "支持矩阵";
const DETAIL_SHEET = "规格明细";
const HEADER_ROW = 4;
const FIRST_DATA_ROW = HEADER_ROW + 1;
const FONT_NAME = "Microsoft YaHei";

const MATRIX_BASE_HEADERS = [
  "序号",
  "功能ID",
  "L0",
  "L1",
  "需求",
  "路标匹配",
  "上线月份",
  "HMOS版本",
  "规划",
  "产品经理",
  "SE",
  "算法",
] as const;

const DETAIL_HEADERS = [
  "序号",
  "功能ID",
  "L0",
  "L1",
  "需求",
  "路标匹配",
  "上线月份",
  "HMOS版本",
  "系列",
  "支持状态",
  "支持型号",
  "排除型号",
  "系统/市场",
  "最低运动健康App版本",
  "最低设备版本",
  "前置条件",
  "证据编号",
  "证据链接",
  "核验日期",
  "证据等级",
  "审核状态",
  "备注",
  "更新时间",
] as const;

interface TieringExportDemand {
  roadmapLabel: string;
  requirement: Requirement;
  featureId: string;
  l0: string;
  l1: string;
}

/**
 * 生成只读核对用的分层分级 Excel。文件不携带回导元数据，不能通过应用回导。
 */
export async function createTieringWorkbook(
  data: WorkspaceData,
  appVersion: string,
  generatedAt: string,
  range: HalfYearRange,
): Promise<Buffer> {
  const generatedDate = requireDate(generatedAt, "生成时间");
  // 先校验范围，避免空数据时静默接受无效半年。
  halfYearSequence(range.start, range.end);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "解决方案需求管理";
  workbook.company = "Huawei";
  workbook.created = generatedDate;
  workbook.modified = generatedDate;

  const demands = tieringDemandsForRange(data, range);
  const featureMap = new Map((data.tiering?.features ?? []).map((item) => [item.id, item]));
  const supportMap = new Map((data.tiering?.supportRecords ?? []).map((item) => [item.id, item]));
  const evidenceMap = new Map((data.tiering?.evidenceSources ?? []).map((item) => [item.id, item]));
  const memberMap = new Map((data.members ?? []).map((item) => [item.id, item]));

  createMatrixSheet(workbook, demands, supportMap, memberMap, range, generatedDate, appVersion);
  createDetailSheet(workbook, demands, featureMap, supportMap, evidenceMap, range, generatedDate, appVersion);

  return Buffer.from(await workbook.xlsx.writeBuffer());
}

export function tieringDemandsForRange(data: WorkspaceData, range: HalfYearRange): TieringExportDemand[] {
  const domainsById = new Map(data.domains.map((item) => [item.id, item]));
  const features = data.tiering?.features ?? [];
  const featureById = new Map(features.map((item) => [item.id, item]));
  const featureByRequirementId = new Map(
    features.filter((item) => item.requirementId).map((item) => [item.requirementId as string, item]),
  );

  return requirementsInRange(data.requirements.filter(isTieringEligibleRequirement), range.start, range.end)
    .map((requirement): TieringExportDemand => {
      const linkedFeature = (requirement.tieringFeatureId && featureById.get(requirement.tieringFeatureId))
        || featureByRequirementId.get(requirement.id);
      const domainL1 = domainsById.get(requirement.domainId);
      const domainL0 = domainsById.get(domainL1?.parentId ?? requirement.domainL0Id);
      return {
        roadmapLabel: (() => { const assignment = data.roadmapAssignments?.find((item) => item.requirementId === requirement.id); return assignment ? `${data.roadmaps?.find((r)=>r.id===assignment.roadmapId)?.name ?? assignment.roadmapId} · ${data.roadmapLanes?.find((r)=>r.id===assignment.laneId)?.name ?? assignment.laneId}` : "未匹配"; })(),
        requirement,
        featureId: linkedFeature?.id ?? requirement.tieringFeatureId ?? tieringFeatureIdForRequirement(requirement.id),
        l0: domainL0?.name ?? linkedFeature?.l0 ?? requirement.domainL0Id,
        l1: domainL1?.name ?? linkedFeature?.l1 ?? requirement.domainId,
      };
    })
    .sort((left, right) => {
      const leftFeature = featureById.get(left.featureId);
      const rightFeature = featureById.get(right.featureId);
      return (leftFeature?.sortOrder ?? Number.MAX_SAFE_INTEGER) - (rightFeature?.sortOrder ?? Number.MAX_SAFE_INTEGER)
        || left.requirement.targetMonth.localeCompare(right.requirement.targetMonth)
        || left.requirement.title.localeCompare(right.requirement.title, "zh-CN")
        || left.requirement.id.localeCompare(right.requirement.id);
    });
}

function createMatrixSheet(
  workbook: ExcelJS.Workbook,
  demands: readonly TieringExportDemand[],
  supportMap: ReadonlyMap<string, TieringSupportRecord>,
  memberMap: ReadonlyMap<string, Member>,
  range: HalfYearRange,
  generatedAt: Date,
  appVersion: string,
): void {
  const sheet = workbook.addWorksheet(MATRIX_SHEET, { properties: { tabColor: { argb: "FF245B4D" } } });
  const headers = [...MATRIX_BASE_HEADERS, ...BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => item.name)];
  prepareSheet(sheet, "分层分级支持矩阵", headers, range, generatedAt, appVersion, demands.length, 5);

  const rows = demands.map(({ requirement, featureId, l0, l1, roadmapLabel }, index) => [
    index + 1,
    featureId,
    l0,
    l1,
    requirement.title,
    roadmapLabel,
    monthDate(requirement.targetMonth),
    hmosVersionForMonth(requirement.targetMonth),
    memberName(requirement.planningMemberId, memberMap),
    memberName(requirement.productManagerMemberId, memberMap),
    memberName(requirement.seMemberId, memberMap),
    memberName(requirement.algorithmMemberId, memberMap),
    ...BUILT_IN_TIERING_PRODUCT_SERIES.map((series) => supportMap.get(`${featureId}::${series.id}`)?.status ?? "未判断"),
  ] satisfies ExcelJS.CellValue[]);
  addRowsOrEmptyState(sheet, rows, headers.length, "当前区间暂无需求");

  const widths = [8, 14, 18, 18, 32, 16, 14, 14, 14, 14, 12, 12, ...BUILT_IN_TIERING_PRODUCT_SERIES.map(() => 18)];
  sheet.columns = widths.map((width) => ({ width }));
  styleDataArea(sheet, headers.length, rows.length, 36, widths);
  sheet.getColumn(7).numFmt = 'yyyy"-"mm';

  if (rows.length) {
    const statusStartColumn = MATRIX_BASE_HEADERS.length + 1;
    for (let rowNumber = FIRST_DATA_ROW; rowNumber < FIRST_DATA_ROW + rows.length; rowNumber += 1) {
      for (let column = statusStartColumn; column <= headers.length; column += 1) {
        styleStatusCell(sheet.getCell(rowNumber, column));
      }
    }
  }
  finishFilter(sheet, headers.length, rows.length);
}

function createDetailSheet(
  workbook: ExcelJS.Workbook,
  demands: readonly TieringExportDemand[],
  featureMap: ReadonlyMap<string, TieringFeature>,
  supportMap: ReadonlyMap<string, TieringSupportRecord>,
  evidenceMap: ReadonlyMap<string, TieringEvidenceSource>,
  range: HalfYearRange,
  generatedAt: Date,
  appVersion: string,
): void {
  const sheet = workbook.addWorksheet(DETAIL_SHEET, { properties: { tabColor: { argb: "FF6B8F83" } } });
  prepareSheet(sheet, "分层分级规格明细", [...DETAIL_HEADERS], range, generatedAt, appVersion, demands.length, 5);

  const rows: ExcelJS.CellValue[][] = [];
  for (const { requirement, featureId, l0, l1, roadmapLabel } of demands) {
    const feature = featureMap.get(featureId);
    for (const series of BUILT_IN_TIERING_PRODUCT_SERIES) {
      const record = supportMap.get(`${featureId}::${series.id}`);
      rows.push([
        rows.length + 1,
        feature?.id ?? featureId,
        l0,
        l1,
        requirement.title,
        roadmapLabel,
        monthDate(requirement.targetMonth),
        hmosVersionForMonth(requirement.targetMonth),
        series.name,
        record?.status ?? "未判断",
        record?.supportedModels ?? "",
        record?.excludedModels ?? "",
        record?.phoneOsMarket ?? "",
        record?.minHealthAppVersion ?? "",
        record?.minDeviceVersion ?? "",
        record?.prerequisites ?? "",
        record?.sourceIds.join("、") ?? "",
        record ? evidenceLinks(record, evidenceMap) : "",
        dateValue(record?.verifiedAt),
        record?.evidenceLevel ?? "",
        record?.auditStatus ?? "待内审",
        record?.note ?? "",
        dateValue(record?.updatedAt),
      ]);
    }
  }
  addRowsOrEmptyState(sheet, rows, DETAIL_HEADERS.length, "当前区间暂无需求");

  const widths = [8, 14, 18, 18, 32, 16, 14, 14, 18, 14, 24, 24, 22, 20, 18, 28, 18, 38, 14, 14, 14, 32, 22];
  sheet.columns = widths.map((width) => ({ width }));
  styleDataArea(sheet, DETAIL_HEADERS.length, rows.length, 54, widths);
  sheet.getColumn(7).numFmt = 'yyyy"-"mm';
  sheet.getColumn(19).numFmt = "yyyy-mm-dd";
  sheet.getColumn(23).numFmt = "yyyy-mm-dd hh:mm";
  for (let rowNumber = FIRST_DATA_ROW; rowNumber < FIRST_DATA_ROW + rows.length; rowNumber += 1) {
    styleStatusCell(sheet.getCell(rowNumber, 10));
  }
  finishFilter(sheet, DETAIL_HEADERS.length, rows.length);
}

function prepareSheet(
  sheet: ExcelJS.Worksheet,
  title: string,
  headers: readonly string[],
  range: HalfYearRange,
  generatedAt: Date,
  appVersion: string,
  demandCount: number,
  freezeColumns: number,
): void {
  sheet.views = [{
    state: "frozen",
    xSplit: freezeColumns,
    ySplit: HEADER_ROW,
    topLeftCell: `${excelColumn(freezeColumns + 1)}${FIRST_DATA_ROW}`,
    activeCell: `${excelColumn(freezeColumns + 1)}${FIRST_DATA_ROW}`,
    showGridLines: false,
  }];
  sheet.getCell("A1").value = title;
  sheet.getCell("A1").font = { name: FONT_NAME, size: 15, bold: true, color: { argb: "FF17242C" } };
  sheet.getRow(1).height = 26;

  sheet.getCell("A2").value = `导出范围：${shortHalfYear(range.start)}–${shortHalfYear(range.end)}　生成时间：${formatTimestamp(generatedAt)}`;
  sheet.getCell("A2").font = { name: FONT_NAME, size: 10, italic: true, color: { argb: "FF5F6D75" } };
  sheet.getCell("A3").value = `应用版本：v${appVersion}　需求：${demandCount} 条　口径：全部需求；固定 ${BUILT_IN_TIERING_PRODUCT_SERIES.length} 个产品系列。本文件仅用于查看与核对，不支持回导。`;
  sheet.getCell("A3").font = { name: FONT_NAME, size: 9, color: { argb: "FF6E7C84" } };

  const headerRow = sheet.getRow(HEADER_ROW);
  headerRow.values = [...headers];
  headerRow.height = 34;
  for (let column = 1; column <= headers.length; column += 1) {
    const cell = headerRow.getCell(column);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF245B4D" } };
    cell.font = { name: FONT_NAME, size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = { right: { style: "thin", color: { argb: "FFFFFFFF" } } };
  }
  sheet.getRow(2).height = 20;
  sheet.getRow(3).height = 20;
  sheet.pageSetup = {
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.2, right: 0.2, top: 0.4, bottom: 0.4, header: 0.15, footer: 0.15 },
  };
  sheet.headerFooter.oddFooter = "&L分层分级导出&C&P / &N&R解决方案需求管理";
}

function addRowsOrEmptyState(
  sheet: ExcelJS.Worksheet,
  rows: readonly ExcelJS.CellValue[][],
  columnCount: number,
  emptyMessage: string,
): void {
  if (rows.length) {
    sheet.addRows(rows as ExcelJS.CellValue[][]);
    return;
  }
  sheet.getCell(FIRST_DATA_ROW, 1).value = emptyMessage;
  sheet.mergeCells(FIRST_DATA_ROW, 1, FIRST_DATA_ROW, columnCount);
  const cell = sheet.getCell(FIRST_DATA_ROW, 1);
  cell.font = { name: FONT_NAME, size: 10, italic: true, color: { argb: "FF7C8990" } };
  cell.alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(FIRST_DATA_ROW).height = 34;
}

function styleDataArea(
  sheet: ExcelJS.Worksheet,
  columnCount: number,
  dataRowCount: number,
  minimumRowHeight: number,
  columnWidths: readonly number[],
): void {
  if (!dataRowCount) return;
  for (let rowNumber = FIRST_DATA_ROW; rowNumber < FIRST_DATA_ROW + dataRowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    let estimatedLines = 1;
    for (let column = 1; column <= columnCount; column += 1) {
      const cell = row.getCell(column);
      cell.font = { name: FONT_NAME, size: 9, color: { argb: "FF1D2930" } };
      cell.alignment = { vertical: "top", horizontal: column === 1 ? "center" : "left", wrapText: true };
      cell.border = { bottom: { style: "thin", color: { argb: "FFDCE3E6" } } };
      estimatedLines = Math.max(estimatedLines, wrappedLineCount(cell.text, columnWidths[column - 1] ?? 16));
    }
    // ExcelJS 不会替自动换行的内容计算行高，按中英文显示宽度估算，避免长规格被固定行高截断。
    row.height = Math.max(minimumRowHeight, estimatedLines * 12 + 8);
  }
}

function wrappedLineCount(value: string, columnWidth: number): number {
  if (!value) return 1;
  const availableWidth = Math.max(4, columnWidth - 2);
  return value.split("\n").reduce((total, line) => {
    const displayWidth = [...line].reduce((sum, character) => sum + (/^[\u0000-\u00ff]$/.test(character) ? 1 : 2), 0);
    return total + Math.max(1, Math.ceil(displayWidth / availableWidth));
  }, 0);
}

function finishFilter(sheet: ExcelJS.Worksheet, columnCount: number, dataRowCount: number): void {
  sheet.autoFilter = {
    from: { row: HEADER_ROW, column: 1 },
    to: { row: Math.max(HEADER_ROW, HEADER_ROW + dataRowCount), column: columnCount },
  };
}

function styleStatusCell(cell: ExcelJS.Cell): void {
  const colors: Record<string, { fill: string; font: string }> = {
    支持: { fill: "FFE6F4EC", font: "FF226548" },
    不支持: { fill: "FFFDEAE7", font: "FFA63F33" },
    未判断: { fill: "FFF0F2F3", font: "FF66747C" },
  };
  const color = colors[cell.text] ?? colors.未判断;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color.fill } };
  cell.font = { name: FONT_NAME, size: 9, bold: true, color: { argb: color.font } };
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}

function memberName(id: string | undefined, members: ReadonlyMap<string, Member>): string {
  return (id && members.get(id)?.name) || "未配置";
}

function evidenceLinks(record: TieringSupportRecord, evidenceMap: ReadonlyMap<string, TieringEvidenceSource>): string {
  const values = [
    record.sourceUrl,
    ...record.sourceIds.map((id) => evidenceMap.get(id)?.url ?? ""),
  ].map((item) => item.trim()).filter(Boolean);
  return [...new Set(values)].join("\n");
}

function monthDate(value: string): Date {
  if (!isValidRequirementMonth(value)) throw new Error(`无效上线月份：${value}`);
  const [year, month] = value.split("-").map(Number);
  return new Date(Date.UTC(year, month - 1, 1));
}

function dateValue(value: string | undefined): Date | string {
  if (!value) return "";
  const result = new Date(value);
  return Number.isNaN(result.getTime()) ? value : result;
}

function requireDate(value: string, label: string): Date {
  const result = new Date(value);
  if (Number.isNaN(result.getTime())) throw new Error(`${label}无效`);
  return result;
}

function formatTimestamp(value: Date): string {
  const parts = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(value);
  const pick = (type: Intl.DateTimeFormatPartTypes) => parts.find((item) => item.type === type)?.value ?? "";
  return `${pick("year")}-${pick("month")}-${pick("day")} ${pick("hour")}:${pick("minute")}`;
}

function excelColumn(column: number): string {
  let result = "";
  for (let current = column; current > 0; current = Math.floor((current - 1) / 26)) {
    result = String.fromCharCode(65 + ((current - 1) % 26)) + result;
  }
  return result;
}
