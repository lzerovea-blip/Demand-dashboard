import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  applyPreparedWorkspaceWorkbook,
  createWorkspaceWorkbook as createWorkbookUnderTest,
  inspectWorkspaceWorkbook,
  workspaceDataForWorkbookRange,
} from "../src/main/workspaceWorkbook";
import type { DictionaryItem, GroupOverride, Requirement, RequirementImage, WorkspaceData } from "../src/shared/types";
import { BUILT_IN_TIERING_PRODUCT_SERIES, normalizeWorkspaceTieringScope, tieringSupportId } from "../src/shared/tiering";

const createWorkspaceWorkbook = (data: WorkspaceData, version: string, exportedAt?: string, range?: { start: string; end: string }) =>
  createWorkbookUnderTest(data, version, exportedAt, range, "v7");
const createCurrentWorkspaceWorkbook = (data: WorkspaceData, version: string, exportedAt?: string, range?: { start: string; end: string }) =>
  createWorkbookUnderTest(data, version, exportedAt, range, "v8");
const createV9WorkspaceWorkbook = (data: WorkspaceData, version: string) =>
  createWorkbookUnderTest(data, version, undefined, undefined, "v9");

describe("协作 Excel v10 十五系列", () => {
  it("导出 15 个系列及两项新系列下拉，内部基线保持在非工作表结构", async () => {
    const data = workspace();
    const bytes = await createWorkbookUnderTest(data, "0.8.31", undefined, undefined, "v11");
    if (process.env.ROADMAP_QA_V10_WORKBOOK_PATH) await writeFile(process.env.ROADMAP_QA_V10_WORKBOOK_PATH, bytes);
    const book = new ExcelJS.Workbook(); await book.xlsx.load(bytes as never);
    const sheet = book.getWorksheet("需求清单")!;
    expect(book.worksheets.map((item) => item.name)).toEqual(["需求清单"]);
    expect(sheet.getRow(1).values).toHaveLength(33);
    expect(sheet.getCell("A1").text).toBe("序号");
    expect(sheet.getCell("B1").text).toBe("操作");
    expect(Array.from({ length: 15 }, (_, index) => sheet.getCell(1, 18 + index).text)).toEqual(BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => item.name));
    expect(sheet.getCell("S1").text).toBe("WATCH Pro系列");
    expect(sheet.getCell("AB1").text).toBe("Band Pro系列");
    expect(sheet.getCell("S2").text).toBe("未判断");
    expect(sheet.getCell("AB2").text).toBe("未判断");
    expect(sheet.getCell("S2").dataValidation.type).toBe("list");
    expect(sheet.getCell("AB2").dataValidation.type).toBe("list");
    expect(sheet.getCell("AF1").text).toBe("无屏系列");
    const stored = await payload(bytes) as { metadata: { formatVersion: string }; rows: Array<{ baselineVisible: string[]; baselineSeries: Record<string, string> }> };
    expect(stored.metadata.formatVersion).toBe("v11");
    expect(stored.rows[0].baselineVisible).toHaveLength(31);
    expect(Object.keys(stored.rows[0].baselineSeries)).toHaveLength(15);
    const xml = await (await JSZip.loadAsync(bytes)).file("xl/worksheets/sheet1.xml")!.async("string");
    expect(xml).not.toContain('<c r="AG');
    expect(xml).not.toContain("health-roadmap-collaboration");
  });

  it("两项新系列分别产生差异，未勾选的系列和需求字段不写入", async () => {
    const data = workspace();
    const edited = await editWorkbook(await createWorkbookUnderTest(data, "0.8.31", undefined, undefined, "v11"), (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      sheet.getCell("C2").value = "Excel 新标题";
      sheet.getCell("S2").value = "支持";
      sheet.getCell("AB2").value = "不支持";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "v10.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((item) => item.id)).toEqual(["r1:field:title", "r1:series:PS14", "r1:series:PS15"]);
    const changed = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-24T02:00:00.000Z", {
      selectedChangeIds: ["r1:series:PS14"], conflictDecisions: {},
    });
    const feature = changed.tiering!.features.find((item) => item.requirementId === "r1")!;
    expect(changed.requirements[0].title).toBe(data.requirements[0].title);
    expect(changed.tiering!.supportRecords.find((item) => item.id === tieringSupportId(feature.id, "PS14"))?.status).toBe("支持");
    expect(changed.tiering!.supportRecords.find((item) => item.id === tieringSupportId(feature.id, "PS15"))?.status).toBe("未判断");
  });

  it("行重排后只改操作列为删除，仍按需求 ID 匹配", async () => {
    const base = workspace();
    const data = workspace({ requirements: [base.requirements[0], { ...base.requirements[0], id: "r2", title: "第二条需求" }] });
    const edited = await editWorkbook(await createWorkbookUnderTest(data, "0.8.31", undefined, undefined, "v11"), (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      const first = Array.from({ length: 32 }, (_, index) => ({ value: sheet.getCell(2, index + 1).value, numFmt: sheet.getCell(2, index + 1).numFmt }));
      const second = Array.from({ length: 32 }, (_, index) => ({ value: sheet.getCell(3, index + 1).value, numFmt: sheet.getCell(3, index + 1).numFmt }));
      for (let column = 1; column <= 32; column += 1) {
        sheet.getCell(2, column).value = second[column - 1].value;
        sheet.getCell(2, column).numFmt = second[column - 1].numFmt;
        sheet.getCell(3, column).value = first[column - 1].value;
        sheet.getCell(3, column).numFmt = first[column - 1].numFmt;
      }
      sheet.getCell("B2").value = "删除";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "v10-已排序.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((item) => item.id)).toEqual(["r2:delete"]);
    const changed = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-24T02:00:00.000Z", {
      selectedChangeIds: ["r2:delete"], conflictDecisions: {},
    });
    expect(changed.requirements.map((item) => item.id)).toEqual(["r1"]);
  });

  it.each(["v8", "v9"] as const)("旧 %s 文件保留新系列的本机状态", async (version) => {
    const data = normalizeWorkspaceTieringScope(workspace());
    const feature = data.tiering!.features.find((item) => item.requirementId === "r1")!;
    data.tiering!.supportRecords.find((item) => item.id === tieringSupportId(feature.id, "PS14"))!.status = "支持";
    data.tiering!.supportRecords.find((item) => item.id === tieringSupportId(feature.id, "PS15"))!.status = "不支持";
    const edited = await editWorkbook(await createWorkbookUnderTest(data, "0.8.30", undefined, undefined, version), (book) => {
      book.getWorksheet("需求清单")!.getCell(version === "v8" ? "Q2" : "R2").value = "支持";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, `${version}.xlsx`, data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((item) => item.id)).toEqual(["r1:series:PS01"]);
    const changed = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-24T02:00:00.000Z", {
      selectedChangeIds: ["r1:series:PS01"], conflictDecisions: {},
    });
    expect(changed.tiering!.supportRecords.find((item) => item.id === tieringSupportId(feature.id, "PS14"))?.status).toBe("支持");
    expect(changed.tiering!.supportRecords.find((item) => item.id === tieringSupportId(feature.id, "PS15"))?.status).toBe("不支持");
  });
});

describe("协作 Excel v9 序号与兼容回导", () => {
  it("导出 A 列连续序号、B 列操作与顺移后的系列下拉，序号不参与回导", async () => {
    const base = workspace();
    const data = workspace({ requirements: [base.requirements[0], { ...base.requirements[0], id: "r2", title: "第二条需求" }] });
    const bytes = await createV9WorkspaceWorkbook(data, "0.8.30");
    if (process.env.ROADMAP_QA_V9_WORKBOOK_PATH) await writeFile(process.env.ROADMAP_QA_V9_WORKBOOK_PATH, bytes);
    const book = new ExcelJS.Workbook(); await book.xlsx.load(bytes as never);
    const sheet = book.getWorksheet("需求清单")!;
    expect(sheet.getRow(1).values).toHaveLength(31);
    expect(sheet.getCell("A1").text).toBe("序号");
    expect(sheet.getCell("B1").text).toBe("操作");
    expect([sheet.getCell("A2").value, sheet.getCell("A3").value]).toEqual([1, 2]);
    expect(sheet.getCell("B2").dataValidation.type).toBe("list");
    expect(sheet.getCell("R2").text).toBe("未判断");
    expect(sheet.getCell("R2").dataValidation.type).toBe("list");
    expect(sheet.views[0]).toMatchObject({ state: "frozen", xSplit: 3, ySplit: 1 });
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(xml).not.toContain('<c r="AE');
    const edited = await editWorkbook(bytes, (workbook) => {
      const cell = workbook.getWorksheet("需求清单")!.getCell("A2");
      cell.value = 999;
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "v9.xlsx", data);
    expect(prepared.preview.formatVersion).toBe("v9");
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges).toEqual([]);
  });

  it("排序后序号随需求移动，B 列删除和另一行标题修改仍按原 ID 合入", async () => {
    const base = workspace();
    const data = workspace({ requirements: [base.requirements[0], { ...base.requirements[0], id: "r2", title: "第二条需求" }] });
    const edited = await editWorkbook(await createV9WorkspaceWorkbook(data, "0.8.30"), (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      const first = Array.from({ length: 30 }, (_, index) => ({ value: sheet.getCell(2, index + 1).value, numFmt: sheet.getCell(2, index + 1).numFmt }));
      const second = Array.from({ length: 30 }, (_, index) => ({ value: sheet.getCell(3, index + 1).value, numFmt: sheet.getCell(3, index + 1).numFmt }));
      for (let column = 1; column <= 30; column += 1) {
        sheet.getCell(2, column).value = second[column - 1].value;
        sheet.getCell(2, column).numFmt = second[column - 1].numFmt;
        sheet.getCell(3, column).value = first[column - 1].value;
        sheet.getCell(3, column).numFmt = first[column - 1].numFmt;
      }
      expect([sheet.getCell("A2").value, sheet.getCell("A3").value]).toEqual([2, 1]);
      sheet.getCell("B2").value = "删除";
      sheet.getCell("C3").value = "第一条需求已修改";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "已排序.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((change) => change.id)).toEqual(["r2:delete", "r1:field:title"]);
    const result = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-24T02:00:00.000Z", {
      selectedChangeIds: ["r2:delete", "r1:field:title"], conflictDecisions: {},
    });
    expect(result.requirements.map((item) => [item.id, item.title])).toEqual([["r1", "第一条需求已修改"]]);
  });

  it("WPS 移除 customXml 后可从自定义属性读取 v9；两份元数据均丢失时仅操作列可安全恢复", async () => {
    const data = workspace();
    const original = await createV9WorkspaceWorkbook(data, "0.8.30");
    const edited = await editWorkbook(original, (book) => { book.getWorksheet("需求清单")!.getCell("B2").value = "删除"; });
    if (process.env.ROADMAP_QA_V9_EDITED_PATH) await writeFile(process.env.ROADMAP_QA_V9_EDITED_PATH, edited);
    const originalZip = await JSZip.loadAsync(original);
    const zip = await JSZip.loadAsync(edited);
    zip.remove("customXml/item1.xml");
    const recovered = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "无元数据.xlsx", data);
    expect(recovered.preview.fieldChanges?.map((change) => change.kind)).toEqual(["delete"]);
    zip.file("docProps/custom.xml", await originalZip.file("docProps/custom.xml")!.async("string"));
    const fromProperties = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "WPS保存.xlsx", data);
    expect(fromProperties.preview.errors).toEqual([]);
    expect(fromProperties.preview.fieldChanges?.map((change) => change.kind)).toEqual(["delete"]);
  });

  it.skipIf(!process.env.ROADMAP_QA_V9_ROUNDTRIP_PATH)("外部表格软件另存的 v9 文件仍可回导", async () => {
    const data = workspace();
    const saved = await readFile(process.env.ROADMAP_QA_V9_ROUNDTRIP_PATH!);
    const preview = await inspectWorkspaceWorkbook(saved, "外部另存.xlsx", data);
    expect(preview.preview.errors).toEqual([]);
    expect(preview.preview.fieldChanges?.map((change) => change.kind)).toEqual(["delete"]);
  });
});

describe("协作 Excel v8 逐项合入", () => {
  it("仅显示 A–AC，一张表、13 系列有下拉，非分层行空白", async () => {
    const data = workspace({ requirements: [workspace().requirements[0], { ...workspace().requirements[0], id: "medical", title: "认证事项", source: "NMPA认证", category: "平台演进", productIds: [], targetMonth: "2027-04" }] });
    const bytes = await createCurrentWorkspaceWorkbook(data, "0.8.26");
    if (process.env.ROADMAP_QA_WORKBOOK_PATH) await writeFile(process.env.ROADMAP_QA_WORKBOOK_PATH, bytes);
    const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(bytes as never);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["需求清单"]);
    const sheet = workbook.worksheets[0];
    expect(sheet.getRow(1).values).toHaveLength(30);
    expect(sheet.getRow(1).values).not.toContain("来源");
    expect(sheet.getRow(1).values).not.toContain("匹配产品");
    expect(Array.from({ length: 13 }, (_, index) => sheet.getCell(1, index + 17).text)).toEqual(BUILT_IN_TIERING_PRODUCT_SERIES.filter((item) => !["PS14", "PS15"].includes(item.id)).map((item) => item.name));
    expect(sheet.getCell("Q2").text).toBe("未判断");
    expect(sheet.getCell("Q2").dataValidation.formulae![0]).toContain("支持,不支持,未判断");
    expect(sheet.getCell("Q3").text).toBe("");
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(xml).not.toContain("<c r=\"AD");
    expect(xml).not.toContain("health-roadmap-collaboration");
  });

  it("空成员与超长字典不写损坏的 Excel 数据验证，仍可手填并回导", async () => {
    const base = workspace();
    const data = workspace({
      domains: [...base.domains, ...Array.from({ length: 40 }, (_, index) => ({
        id: `extra-${index}`, name: `扩展领域${index}的较长名称`, level: "L1" as const,
        parentId: "l0", sortOrder: index + 1, active: true,
      }))],
      members: [],
      pptAccessPolicy: { entries: [], updatedAt: base.pptAccessPolicy!.updatedAt },
    });
    const bytes = await createCurrentWorkspaceWorkbook(data, "0.8.27");
    const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(bytes as never);
    const sheet = workbook.getWorksheet("需求清单")!;
    expect(sheet.getCell("A2").dataValidation.type).toBe("list");
    expect(sheet.getCell("D2").dataValidation.type).toBe("list");
    expect(sheet.getCell("E2").dataValidation).toBeUndefined();
    for (const column of ["M", "N", "O", "P"]) expect(sheet.getCell(`${column}2`).dataValidation).toBeUndefined();
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    const formulas = Array.from(xml.matchAll(/<formula1>([^<]*)<\/formula1>/g), ([, formula]) => formula);
    expect(formulas).not.toContain('""');
    expect(formulas.every((formula) => formula.length <= 255)).toBe(true);
    const edited = await editWorkbook(bytes, (book) => { book.getWorksheet("需求清单")!.getCell("E2").value = "扩展领域1的较长名称"; });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((change) => change.field)).toEqual(["领域 L1"]);
  });

  it("过长或含分隔符的成员选项省略下拉，不妨碍文件生成", async () => {
    const base = workspace();
    const members = Array.from({ length: 25 }, (_, index) => ({
      ...base.members![0], id: `member-${index}`, name: `规划成员${index}测试较长名字`,
      employeeNumber: `E${index}`, microsoftAccount: `planner${index}@company.com`,
    }));
    const policy = { entries: [], updatedAt: base.pptAccessPolicy!.updatedAt };
    const bytes = await createCurrentWorkspaceWorkbook(workspace({ members, pptAccessPolicy: policy }), "0.8.27");
    const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(bytes as never);
    expect(workbook.getWorksheet("需求清单")!.getCell("M2").dataValidation).toBeUndefined();
    const withComma = await createCurrentWorkspaceWorkbook(workspace({ members: [{ ...base.members![0], name: "甲,乙" }], pptAccessPolicy: policy }), "0.8.27");
    const commaWorkbook = new ExcelJS.Workbook(); await commaWorkbook.xlsx.load(withComma as never);
    expect(commaWorkbook.getWorksheet("需求清单")!.getCell("M2").dataValidation).toBeUndefined();
  });

  it("标题与系列可分别勾选，默认零项写入且保留来源、产品、图片", async () => {
    const data = workspace();
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.26"), (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      sheet.getCell("B2").value = "新标题";
      sheet.getCell("Q2").value = "支持";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((item) => item.field)).toEqual(["需求标题", "WATCH数字系列"]);
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-22T01:00:00.000Z", { selectedChangeIds: [], conflictDecisions: {} })).toEqual(data);
    const changed = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-22T01:00:00.000Z", { selectedChangeIds: ["r1:series:PS01"], conflictDecisions: {} });
    expect(changed.requirements[0].title).toBe("社交时差");
    expect(changed.requirements[0].source).toBe("健康基础");
    expect(changed.requirements[0].productIds).toEqual(["p1"]);
    expect(changed.requirements[0].images).toEqual([image]);
    const feature = changed.tiering!.features.find((item) => item.requirementId === "r1")!;
    expect(changed.tiering!.supportRecords.find((item) => item.id === tieringSupportId(feature.id, "PS01"))?.status).toBe("支持");
    expect(changed.tiering!.changeLogs).toHaveLength(1);
  });

  it("并发修改同一字段必须逐项决策，本机其他字段始终保留", async () => {
    const data = workspace();
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.26"), (book) => { book.getWorksheet("需求清单")!.getCell("B2").value = "Excel 标题"; });
    const local = workspace({ requirements: [{ ...data.requirements[0], title: "本机标题", description: "本机另改描述", updatedAt: "2026-09-22T00:00:00.000Z" }] });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", local);
    expect(prepared.preview.fieldChanges?.[0]).toMatchObject({ id: "r1:field:title", conflict: true });
    expect(() => applyPreparedWorkspaceWorkbook(local, prepared, "local-wins", undefined, { selectedChangeIds: ["r1:field:title"], conflictDecisions: {} })).toThrow(/选择/);
    const result = applyPreparedWorkspaceWorkbook(local, prepared, "local-wins", "2026-09-22T01:00:00.000Z", { selectedChangeIds: ["r1:field:title"], conflictDecisions: { "r1:field:title": "excel" } });
    expect(result.requirements[0]).toMatchObject({ title: "Excel 标题", description: "本机另改描述" });
    expect(() => applyPreparedWorkspaceWorkbook({ ...local, requirements: [{ ...local.requirements[0], updatedAt: "2026-09-22T01:01:00.000Z" }] }, prepared, "local-wins", undefined, { selectedChangeIds: [], conflictDecisions: {} })).toThrow(/本机需求/);
  });

  it("仅操作列删除有效，整行移除不删本机需求", async () => {
    const data = workspace();
    const deletedRow = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.26"), (book) => { book.getWorksheet("需求清单")!.spliceRows(2, 1); });
    const missing = await inspectWorkspaceWorkbook(deletedRow, "v8.xlsx", data);
    expect(missing.preview.fieldChanges).toEqual([]);
    const marked = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.26"), (book) => { book.getWorksheet("需求清单")!.getCell("A2").value = "删除"; });
    const prepared = await inspectWorkspaceWorkbook(marked, "v8.xlsx", data);
    expect(prepared.preview.fieldChanges?.[0].kind).toBe("delete");
    const result = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-22T01:00:00.000Z", { selectedChangeIds: ["r1:delete"], conflictDecisions: {} });
    expect(result.requirements).toEqual([]);
  });

  it("WPS 移除旧 customXml 后，仅改操作列仍可按随行标识安全恢复", async () => {
    const data = workspace();
    const original = await createCurrentWorkspaceWorkbook(data, "0.8.28");
    const edited = await editWorkbook(original, (book) => { book.getWorksheet("需求清单")!.getCell("A2").value = "删除"; });
    const zip = await JSZip.loadAsync(edited);
    zip.remove("customXml/item1.xml");
    const wpsSaved = await zip.generateAsync({ type: "nodebuffer" });
    const prepared = await inspectWorkspaceWorkbook(wpsSaved, "WPS保存.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((change) => change.kind)).toEqual(["delete"]);
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-24T02:00:00.000Z", { selectedChangeIds: ["r1:delete"], conflictDecisions: {} });
    expect(applied.requirements).toEqual([]);
    const changedBusiness = await editWorkbook(original, (book) => { book.getWorksheet("需求清单")!.getCell("B2").value = "无法核对基线"; });
    const changedZip = await JSZip.loadAsync(changedBusiness);
    changedZip.remove("customXml/item1.xml");
    await expect(inspectWorkspaceWorkbook(await changedZip.generateAsync({ type: "nodebuffer" }), "WPS改字段.xlsx", data)).rejects.toThrow(/仅操作列修改可安全恢复/);
  });

  it("新导出文件在 customXml 被移除后仍可从标准自定义属性读取完整基线", async () => {
    const data = workspace();
    const original = await createCurrentWorkspaceWorkbook(data, "0.8.29");
    const sourceZip = await JSZip.loadAsync(original);
    expect(await sourceZip.file("docProps/custom.xml")?.async("string")).toContain("RoadmapCollaboration_Manifest");
    expect(await sourceZip.file("[Content_Types].xml")?.async("string")).toContain("custom-properties+xml");
    const edited = await editWorkbook(original, (book) => { book.getWorksheet("需求清单")!.getCell("A2").value = "删除"; });
    const zip = await JSZip.loadAsync(edited);
    zip.remove("customXml/item1.xml");
    zip.file("docProps/custom.xml", await sourceZip.file("docProps/custom.xml")!.async("string"));
    const prepared = await inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "保留自定义属性.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((change) => change.kind)).toEqual(["delete"]);
    const customXml = await sourceZip.file("docProps/custom.xml")!.async("string");
    zip.file("docProps/custom.xml", customXml.replace(/name="RoadmapCollaboration_Manifest"><vt:lpwstr>[^<]+/, 'name="RoadmapCollaboration_Manifest"><vt:lpwstr>1:broken'));
    await expect(inspectWorkspaceWorkbook(await zip.generateAsync({ type: "nodebuffer" }), "损坏自定义属性.xlsx", data)).rejects.toThrow(/内部系统信息缺失或损坏/);
  });

  it("无效系列、伪造行及损坏内部元数据均拒绝", async () => {
    const data = workspace();
    const original = await createCurrentWorkspaceWorkbook(data, "0.8.26");
    const badStatus = await editWorkbook(original, (book) => { book.getWorksheet("需求清单")!.getCell("Q2").value = "未知"; });
    expect((await inspectWorkspaceWorkbook(badStatus, "v8.xlsx", data)).preview.errors.length).toBeGreaterThan(0);
    const fake = await editWorkbook(original, (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      for (let column = 1; column <= 29; column += 1) sheet.getCell(3, column).value = sheet.getCell(2, column).value;
    });
    expect((await inspectWorkspaceWorkbook(fake, "v8.xlsx", data)).preview.errors.length).toBeGreaterThan(0);
    const corrupted = await editPayload(original, (value) => { value.rows[0].baselineFields.title = "伪造"; });
    await expect(inspectWorkspaceWorkbook(corrupted, "v8.xlsx", data)).rejects.toThrow(/内部系统信息/);
  });

  it("行重排后仍匹配原需求，非分层来源不得填写系列", async () => {
    const base = workspace();
    const data = workspace({ requirements: [base.requirements[0], secondRequirement(base)] });
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.26"), (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      const first = Array.from({ length: 29 }, (_, index) => sheet.getCell(2, index + 1).value);
      const second = Array.from({ length: 29 }, (_, index) => sheet.getCell(3, index + 1).value);
      first.forEach((value, index) => { sheet.getCell(3, index + 1).value = value; });
      second.forEach((value, index) => { sheet.getCell(2, index + 1).value = value; });
      sheet.getCell("B2").value = "重新排序后的午睡";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.[0]).toMatchObject({ requirementId: "r2", field: "需求标题" });
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-22T01:00:00.000Z", { selectedChangeIds: ["r2:field:title"], conflictDecisions: {} });
    expect(applied.requirements.find((item) => item.id === "r2")?.title).toBe("重新排序后的午睡");
    expect(applied.requirements.find((item) => item.id === "r1")?.title).toBe("社交时差");
    const noneligible = workspace({ requirements: [{ ...base.requirements[0], source: "行业" }] });
    const invalid = await editWorkbook(await createCurrentWorkspaceWorkbook(noneligible, "0.8.26"), (book) => { book.getWorksheet("需求清单")!.getCell("Q2").value = "支持"; });
    expect((await inspectWorkspaceWorkbook(invalid, "v8.xlsx", noneligible)).preview.errors.some((item) => item.message.includes("不可回写"))).toBe(true);
  });

  it("内容相同的需求在调整顺序并修改后仍按内部行标识准确匹配", async () => {
    const base = workspace();
    const duplicate: Requirement = {
      ...base.requirements[0],
      id: "r2",
      images: [],
      createdAt: "2026-08-20T00:00:00.000Z",
      updatedAt: "2026-08-21T00:00:00.000Z",
    };
    const data = workspace({ requirements: [base.requirements[0], duplicate] });
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.28"), (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      const first = Array.from({ length: 29 }, (_, index) => sheet.getCell(2, index + 1).value);
      const second = Array.from({ length: 29 }, (_, index) => sheet.getCell(3, index + 1).value);
      const firstStyles = Array.from({ length: 29 }, (_, index) => structuredClone(sheet.getCell(2, index + 1).style));
      const secondStyles = Array.from({ length: 29 }, (_, index) => structuredClone(sheet.getCell(3, index + 1).style));
      first.forEach((value, index) => { sheet.getCell(3, index + 1).value = value; });
      second.forEach((value, index) => { sheet.getCell(2, index + 1).value = value; });
      firstStyles.forEach((value, index) => { sheet.getCell(3, index + 1).style = value; });
      secondStyles.forEach((value, index) => { sheet.getCell(2, index + 1).style = value; });
      sheet.getCell("B2").value = "排序后的第二条需求";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges).toHaveLength(1);
    expect(prepared.preview.fieldChanges?.[0]).toMatchObject({ requirementId: "r2", field: "需求标题" });
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-23T00:00:00.000Z", {
      selectedChangeIds: ["r2:field:title"], conflictDecisions: {},
    });
    expect(applied.requirements.find((item) => item.id === "r2")?.title).toBe("排序后的第二条需求");
    expect(applied.requirements.find((item) => item.id === "r1")?.title).toBe("社交时差");
  });

  it("内容相同的需求调整顺序后仅修改 A 列删除也能准确匹配", async () => {
    const base = workspace();
    const duplicate: Requirement = {
      ...base.requirements[0],
      id: "r2",
      images: [],
      createdAt: "2026-08-20T00:00:00.000Z",
      updatedAt: "2026-08-21T00:00:00.000Z",
    };
    const data = workspace({ requirements: [base.requirements[0], duplicate] });
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.28"), (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      const first = Array.from({ length: 29 }, (_, index) => sheet.getCell(2, index + 1).value);
      const second = Array.from({ length: 29 }, (_, index) => sheet.getCell(3, index + 1).value);
      const firstStyles = Array.from({ length: 29 }, (_, index) => structuredClone(sheet.getCell(2, index + 1).style));
      const secondStyles = Array.from({ length: 29 }, (_, index) => structuredClone(sheet.getCell(3, index + 1).style));
      first.forEach((value, index) => { sheet.getCell(3, index + 1).value = value; });
      second.forEach((value, index) => { sheet.getCell(2, index + 1).value = value; });
      firstStyles.forEach((value, index) => { sheet.getCell(3, index + 1).style = value; });
      secondStyles.forEach((value, index) => { sheet.getCell(2, index + 1).style = value; });
      sheet.getCell("A2").value = "删除";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges).toHaveLength(1);
    expect(prepared.preview.fieldChanges?.[0]).toMatchObject({ requirementId: "r2", kind: "delete" });
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-23T00:00:00.000Z", {
      selectedChangeIds: ["r2:delete"], conflictDecisions: {},
    });
    expect(applied.requirements.map((item) => item.id)).toEqual(["r1"]);
  });

  it("不调整顺序且仅修改 A 列删除时也不依赖可见内容匹配", async () => {
    const base = workspace();
    const duplicate: Requirement = {
      ...base.requirements[0],
      id: "r2",
      images: [],
      createdAt: "2026-08-20T00:00:00.000Z",
      updatedAt: "2026-08-21T00:00:00.000Z",
    };
    const data = workspace({ requirements: [base.requirements[0], duplicate] });
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.28"), (book) => {
      book.getWorksheet("需求清单")!.getCell("A2").value = "删除";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges).toHaveLength(1);
    expect(prepared.preview.fieldChanges?.[0]).toMatchObject({ requirementId: "r1", kind: "delete" });
  });

  it("选中无效 L0/L1 组合整次失败；未选字段不受影响", async () => {
    const base = workspace();
    const data = workspace({ domains: [...base.domains, { id: "sport", name: "运动", level: "L0", active: true, sortOrder: 2 }] });
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.26"), (book) => { book.getWorksheet("需求清单")!.getCell("D2").value = "运动"; });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(() => applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-09-22T01:00:00.000Z", { selectedChangeIds: ["r1:field:domainL0Id"], conflictDecisions: {} })).toThrow(/组合无效/);
    expect(data.requirements[0].domainL0Id).toBe("l0");
  });

  it("L0 与 L1 同名时按层级解析，不把 L1 错认成 L0", async () => {
    const base = workspace();
    const data = workspace({ domains: [{ ...base.domains[0], name: "睡眠" }, base.domains[1]] });
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.26"), (book) => { book.getWorksheet("需求清单")!.getCell("B2").value = "睡眠协作修改"; });
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.fieldChanges?.map((item) => item.field)).toEqual(["需求标题"]);
  });

  it("系列状态并发按字段冲突决策，采用 Excel 才写变更日志", async () => {
    const base = workspace();
    const data = normalizeWorkspaceTieringScope(base);
    const edited = await editWorkbook(await createCurrentWorkspaceWorkbook(data, "0.8.26"), (book) => { book.getWorksheet("需求清单")!.getCell("Q2").value = "不支持"; });
    const local = normalizeWorkspaceTieringScope(data);
    local.tiering!.supportRecords.find((item) => item.seriesId === "PS01")!.status = "支持";
    const prepared = await inspectWorkspaceWorkbook(edited, "v8.xlsx", local);
    expect(prepared.preview.fieldChanges?.[0]).toMatchObject({ id: "r1:series:PS01", conflict: true });
    const retained = applyPreparedWorkspaceWorkbook(local, prepared, "local-wins", "2026-09-22T01:00:00.000Z", { selectedChangeIds: ["r1:series:PS01"], conflictDecisions: { "r1:series:PS01": "local" } });
    expect(retained).toEqual(local);
    const adopted = applyPreparedWorkspaceWorkbook(local, prepared, "local-wins", "2026-09-22T01:00:00.000Z", { selectedChangeIds: ["r1:series:PS01"], conflictDecisions: { "r1:series:PS01": "excel" } });
    expect(adopted.tiering!.supportRecords.find((item) => item.seriesId === "PS01")?.status).toBe("不支持");
    expect(adopted.tiering!.changeLogs).toHaveLength(1);
  });
});

const V7_HEADERS = [
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
  "规划", "产品经理", "SE", "算法",
] as const;

const LEGACY_REQUIREMENT_HEADERS = [
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

const imageBytes = Buffer.from("excel-preserved-image", "utf8");
const image: RequirementImage = {
  id: "image-1",
  name: "睡眠截图.png",
  mimeType: "image/png",
  sizeBytes: imageBytes.length,
  dataUrl: `data:image/png;base64,${imageBytes.toString("base64")}`,
};

function workspace(overrides: Partial<WorkspaceData> = {}): WorkspaceData {
  return {
    requirements: [{
      id: "r1",
      title: "社交时差",
      description: "评估周末和周内的睡眠规律性",
      images: [image],
      domainL0Id: "l0",
      domainId: "d1",
      source: "健康基础",
      scope: "全球",
      allowLatestVersionUpgrade: false,
      overseasRegions: [],
      category: "产品专属",
      targetMonth: "2027-03",
      productIds: ["p1"],
      deviceWorkloadPm: 0.25,
      appWorkloadPm: 0.5,
      cloudWorkloadPm: 0.25,
      unallocatedWorkloadPm: 2,
      workloadPm: 3,
      createdAt: "2026-08-19T00:00:00.000Z",
      updatedAt: "2026-08-20T00:00:00.000Z",
    }],
    domains: [
      { id: "l0", name: "健康", sortOrder: 0, active: true, level: "L0" },
      { id: "d1", name: "睡眠", sortOrder: 0, active: true, level: "L1", parentId: "l0" },
    ],
    products: [{ id: "p1", name: "手表", sortOrder: 0, active: true }],
    groupOverrides: [{
      groupKey: "d1::健康基础::2027-03",
      cardTitle: "睡眠",
      cardSummary: "摘要",
      sortOrder: 0,
      updatedAt: "2026-08-20T00:00:00.000Z",
    }],
    members: [{
      id: "m1",
      name: "规划甲",
      employeeNumber: "E001",
      microsoftAccount: "planner@company.com",
      role: "规划",
      domainL1Ids: ["d1"],
      sortOrder: 0,
      createdAt: "2026-08-20T00:00:00.000Z",
      updatedAt: "2026-08-28T00:00:00.000Z",
    }],
    pptAccessPolicy: {
      entries: [{ memberId: "m1", permission: "read" }],
      updatedAt: "2026-08-28T00:00:00.000Z",
    },
    ...overrides,
  };
}

function secondRequirement(base: WorkspaceData): Requirement {
  return {
    ...base.requirements[0],
    id: "r2",
    title: "午睡识别",
    description: "识别午间睡眠",
    images: [],
    category: "体验优化",
    targetMonth: "2028-01",
    productIds: [],
    allowLatestVersionUpgrade: true,
    unallocatedWorkloadPm: 0,
    workloadPm: 1,
    updatedAt: "2026-08-21T00:00:00.000Z",
  };
}

async function editWorkbook(bytes: Buffer, edit: (workbook: ExcelJS.Workbook) => void): Promise<Buffer> {
  const originalZip = await JSZip.loadAsync(bytes);
  const customXml = await originalZip.file("customXml/item1.xml")?.async("string");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes as never);
  edit(workbook);
  const output = Buffer.from(await workbook.xlsx.writeBuffer());
  if (!customXml) return output;
  const zip = await JSZip.loadAsync(output);
  zip.file("customXml/item1.xml", customXml);
  const relsFile = zip.file("_rels/.rels")!;
  let rels = await relsFile.async("string");
  if (!rels.includes("customXml/item1.xml")) rels = rels.replace("</Relationships>", '<Relationship Id="rIdRoadmapCollaborationV6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="customXml/item1.xml"/></Relationships>');
  zip.file("_rels/.rels", rels);
  return zip.generateAsync({ type: "nodebuffer" });
}

function headerColumn(sheet: ExcelJS.Worksheet, header: string, row = 1): number {
  let result = 0;
  sheet.getRow(row).eachCell((cell, column) => {
    if (cell.text === header) result = column;
  });
  if (!result) throw new Error(`missing test header: ${header}`);
  return result;
}

async function payload(bytes: Buffer): Promise<Record<string, unknown>> {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("customXml/item1.xml")!.async("string");
  const match = /<payload><!\[CDATA\[([\s\S]*?)\]\]><\/payload>/.exec(xml);
  if (!match) throw new Error("missing v6 metadata");
  return JSON.parse(match[1]) as Record<string, unknown>;
}

async function editPayload(bytes: Buffer, edit: (value: Record<string, any>) => void): Promise<Buffer> {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("customXml/item1.xml")!.async("string");
  const match = /<payload><!\[CDATA\[([\s\S]*?)\]\]><\/payload>/.exec(xml);
  if (!match) throw new Error("missing v6 metadata");
  const value = JSON.parse(match[1]) as Record<string, any>;
  edit(value);
  zip.file("customXml/item1.xml", xml.replace(match[0], `<payload><![CDATA[${JSON.stringify(value)}]]></payload>`));
  return zip.generateAsync({ type: "nodebuffer" });
}

describe("协作 Excel v7 单 Sheet 修改版", () => {
  function ownerWorkspace() {
    const data = workspace();
    data.pptAccessPolicy = { entries: [], updatedAt: "2026-09-15T00:00:00.000Z" };
    data.members = ["规划", "产品经理", "SE", "算法", "规划"].map((role, index) => ({
      ...data.members![0], id: `owner-${index}`, name: index === 4 ? "规划成员" : `${role}成员`,
      employeeNumber: `E${index}`, microsoftAccount: `member${index}@company.com`, role: role as NonNullable<WorkspaceData["members"]>[number]["role"],
    }));
    Object.assign(data.requirements[0], { planningMemberId: "owner-0", productManagerMemberId: "owner-1", seMemberId: "owner-2", algorithmMemberId: "owner-3" });
    return data;
  }

  it("四类成员按角色下拉，同名附加工号，修改与算法清空进入差异确认", async () => {
    const data = ownerWorkspace();
    const bytes = await createWorkspaceWorkbook(data, "0.8.22");
    const edited = await editWorkbook(bytes, (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      expect(sheet.getCell("O2").text).toBe("规划成员（E0）");
      expect(sheet.getCell("O2").dataValidation.formulae![0]).toContain("规划成员（E4）");
      expect(sheet.getCell("O2").dataValidation.formulae![0]).not.toContain("SE成员");
      expect(sheet.getCell("R2").dataValidation.allowBlank).toBe(true);
      sheet.getCell("O2").value = "规划成员（E4）";
      sheet.getCell("R2").value = null;
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "成员.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(JSON.stringify(prepared.preview.differences)).toContain("E4");
    const merged = applyPreparedWorkspaceWorkbook(data, prepared, "excel-wins");
    expect(merged.requirements[0]).toMatchObject({ planningMemberId: "owner-4", algorithmMemberId: "", productManagerMemberId: "owner-1", seMemberId: "owner-2" });
    expect(merged.members).toEqual(data.members);
  });

  it("错误岗位、同名歧义和清空已配置必选责任人均拒绝整次回导", async () => {
    const data = ownerWorkspace();
    const bytes = await createWorkspaceWorkbook(data, "0.8.22");
    for (const value of ["SE成员", "规划成员", null]) {
      const edited = await editWorkbook(bytes, (book) => { book.getWorksheet("需求清单")!.getCell("O2").value = value; });
      const prepared = await inspectWorkspaceWorkbook(edited, "无效.xlsx", data);
      expect(prepared.preview.errors.length).toBeGreaterThan(0);
      expect(() => applyPreparedWorkspaceWorkbook(data, prepared, "excel-wins")).toThrow();
      expect(data.requirements[0].planningMemberId).toBe("owner-0");
    }
  });

  it("人员并发修改产生基线冲突且不混淆同名成员", async () => {
    const data = ownerWorkspace();
    const bytes = await createWorkspaceWorkbook(data, "0.8.22");
    const edited = await editWorkbook(bytes, (book) => { book.getWorksheet("需求清单")!.getCell("B2").value = "表格更新标题"; });
    data.requirements[0].planningMemberId = "owner-4";
    data.requirements[0].updatedAt = "2026-09-15T00:00:00.000Z";
    const prepared = await inspectWorkspaceWorkbook(edited, "冲突.xlsx", data);
    expect(prepared.preview.conflicts).toHaveLength(1);
  });

  it("旧 v6 的十四列更新保留本机四类责任人", async () => {
    const data = ownerWorkspace();
    let bytes = await createWorkspaceWorkbook(data, "0.8.22");
    bytes = await editWorkbook(bytes, (book) => {
      const sheet = book.getWorksheet("需求清单")!;
      sheet.removeTable("RequirementsTable");
      sheet.spliceColumns(15, 4);
      sheet.getCell("B2").value = "旧版表格标题";
      sheet.addTable({ name: "RequirementsTable", ref: "A1", headerRow: true, columns: V7_HEADERS.slice(0, 14).map((name) => ({ name })), rows: [Array.from({ length: 14 }, (_, index) => sheet.getCell(2, index + 1).value)] });
    });
    const item = data.requirements[0];
    bytes = await editPayload(bytes, (value) => {
      value.metadata.formatVersion = "v6";
      delete value.dictionaries.members;
      value.rows[0].baselineVisible = value.rows[0].baselineVisible.slice(0, 14);
      value.rows[0].baselineHash = createHash("sha256").update(JSON.stringify({
        id: item.id, title: item.title.trim(), description: item.description.trim(), domainL0Id: item.domainL0Id, domainId: item.domainId,
        source: item.source, scope: item.scope, allowLatestVersionUpgrade: item.allowLatestVersionUpgrade, category: item.category,
        targetMonth: item.targetMonth, productIds: [...item.productIds].sort(), overseasRegions: [...item.overseasRegions].sort(),
        deviceWorkloadPm: item.deviceWorkloadPm, appWorkloadPm: item.appWorkloadPm, cloudWorkloadPm: item.cloudWorkloadPm, unallocatedWorkloadPm: item.unallocatedWorkloadPm,
      })).digest("hex");
    });
    const prepared = await inspectWorkspaceWorkbook(bytes, "旧版.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.conflicts).toEqual([]);
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "excel-wins").requirements[0]).toMatchObject({
      title: "旧版表格标题", planningMemberId: "owner-0", productManagerMemberId: "owner-1", seMemberId: "owner-2", algorithmMemberId: "owner-3",
    });
  });
  it("新增两项工作量不进入 A-N 或基线，回导保留本机数值并重算总量", async () => {
    const data = workspace();
    const base = data.requirements[0];
    base.researchAppWorkloadPm = 2.5;
    base.algorithmWorkloadPm = 3.5;
    base.workloadPm += 6;
    const bytes = await createWorkspaceWorkbook(data, "0.8.17", "2026-09-09T00:00:00.000Z");
    const edited = await editWorkbook(bytes, (book) => {
      expect(book.worksheets).toHaveLength(1);
      const sheet = book.getWorksheet("需求清单")!;
      expect(sheet.columnCount).toBe(18);
      expect(sheet.getRow(1).values).toEqual([undefined, ...V7_HEADERS]);
      sheet.getCell("M2").value = 7;
    });
    const current = structuredClone(data);
    current.requirements[0].researchAppWorkloadPm = 8;
    current.requirements[0].workloadPm += 5.5;
    current.requirements[0].updatedAt = "2026-09-10T00:00:00.000Z";
    const inspected = await inspectWorkspaceWorkbook(edited, "协作.xlsx", current);
    expect(inspected.preview.conflicts).toHaveLength(0);
    const merged = applyPreparedWorkspaceWorkbook(current, inspected, "excel-wins");
    expect(merged.requirements[0]).toMatchObject({ researchAppWorkloadPm: 8, algorithmWorkloadPm: 3.5, appWorkloadPm: 7, workloadPm: base.deviceWorkloadPm + 7 + base.cloudWorkloadPm + base.unallocatedWorkloadPm + 11.5 });
  });
  it("只导出需求清单，首行为表头且工作表仅包含 A-N 业务字段", async () => {
    const bytes = await createWorkspaceWorkbook(workspace(), "0.7.28", "2026-08-29T08:00:00.000Z", {
      start: "2027H1",
      end: "2027H2",
    });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes as never);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["需求清单"]);
    const sheet = workbook.getWorksheet("需求清单")!;
    expect(sheet.getRow(1).values).toEqual([undefined, ...V7_HEADERS]);
    expect(sheet.views[0]).toMatchObject({ state: "frozen", xSplit: 2, ySplit: 1 });
    for (let column = 1; column <= 18; column += 1) expect(sheet.getColumn(column).hidden).not.toBe(true);
    for (let column = 19; column <= 30; column += 1) {
      expect(sheet.getColumn(column).hidden).not.toBe(true);
      expect(sheet.getColumn(column).values.filter(Boolean)).toEqual([]);
    }
    expect(sheet.getCell("A1").fill).toMatchObject({ pattern: "solid" });
    expect(sheet.getCell("A1").fill).toMatchObject({ fgColor: { argb: "FFD9F2E6" } });
    expect(sheet.getCell("A1").font).toMatchObject({ bold: true, color: { argb: "FF17352A" } });
    expect(sheet.getCell("B2").fill).toMatchObject({ pattern: "none" });
    expect(sheet.getCell("B2").font).toMatchObject({ color: { argb: "FF000000" } });
    expect(sheet.getCell("A2").dataValidation.formulae).toEqual(expect.arrayContaining([expect.stringContaining("保留")]));
    expect(sheet.getCell("H2").dataValidation.formulae).toEqual(expect.arrayContaining([expect.stringContaining("是")]));
    expect(sheet.getCell("D2").dataValidation.formulae?.[0]).toContain("健康");
    expect(sheet.getCell("K2").dataValidation.formulae?.[0]).toContain("手表");
    expect(await payload(bytes)).toMatchObject({
      metadata: {
        formatKey: "health-roadmap-collaboration",
        formatVersion: "v7",
        appVersion: "0.7.28",
        range: { start: "2027H1", end: "2027H2" },
      },
    });
  });

  it("无修改往返保持所有本机数据", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.28", "2026-08-29T08:00:00.000Z");
    const prepared = await inspectWorkspaceWorkbook(bytes, "需求修改.xlsx", data);
    expect(prepared.preview.formatVersion).toBe("v7");
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts).toMatchObject({ added: 0, updated: 0, deleted: 0, unchanged: 1, conflicts: 0 });
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins")).toEqual(data);
  });

  it("按导出区间仅写入区间内需求，回导不影响范围外数据", async () => {
    const base = workspace();
    const data = workspace({ requirements: [...base.requirements, secondRequirement(base)] });
    const bytes = await createWorkspaceWorkbook(data, "0.7.28", "2026-08-29T08:00:00.000Z", {
      start: "2027H1",
      end: "2027H1",
    });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes as never);
    const sheet = workbook.getWorksheet("需求清单")!;
    expect(sheet.getCell("B2").text).toBe("社交时差");
    expect(sheet.getCell("B3").text).toBe("");
    expect(await payload(bytes)).toMatchObject({ metadata: { range: { start: "2027H1", end: "2027H1" } } });
    const prepared = await inspectWorkspaceWorkbook(bytes, "区间修改.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins")).toEqual(data);
  });

  it("回导可修改可见字段并保留图片、待拆分工作量、字典、卡片、人员和白名单", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.28");
    const edited = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(2, headerColumn(sheet, "需求标题")).value = "社交时差优化";
      sheet.getCell(2, headerColumn(sheet, "需求描述")).value = "Excel 修改后的描述";
      sheet.getCell(2, headerColumn(sheet, "适用范围")).value = "仅国内";
      sheet.getCell(2, headerColumn(sheet, "允许最近版本升级")).value = "是";
      sheet.getCell(2, headerColumn(sheet, "上线月份")).value = new Date(2027, 4, 1, 12);
      sheet.getCell(2, headerColumn(sheet, "设备工作量")).value = 1.25;
      sheet.getCell(2, headerColumn(sheet, "App 工作量")).value = 0;
      sheet.getCell(2, headerColumn(sheet, "云侧工作量")).value = 0.5;
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "修改.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts.updated).toBe(1);
    expect(prepared.preview.differences).toHaveLength(1);
    expect(prepared.preview.differences[0].fields.map((item) => item.field)).toEqual(expect.arrayContaining(["需求标题", "需求描述", "适用范围", "允许最近版本升级", "上线月份", "设备工作量"]));
    expect(prepared.preview.differenceDigest).toMatch(/^diff-/);
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "excel-wins", "2026-08-30T00:00:00.000Z");
    expect(applied.requirements[0]).toMatchObject({
      title: "社交时差优化",
      description: "Excel 修改后的描述",
      scope: "仅国内",
      allowLatestVersionUpgrade: true,
      targetMonth: "2027-05",
      deviceWorkloadPm: 1.25,
      appWorkloadPm: 0,
      cloudWorkloadPm: 0.5,
      unallocatedWorkloadPm: 2,
      workloadPm: 3.75,
      images: [image],
    });
    expect(applied.domains).toEqual(data.domains);
    expect(applied.products).toEqual(data.products);
    expect(applied.groupOverrides).toEqual(data.groupOverrides);
    expect(applied.members).toEqual(data.members);
    expect(applied.pptAccessPolicy).toEqual(data.pptAccessPolicy);
  });

  it("重新排序行后依靠内部系统数据准确回写", async () => {
    const base = workspace();
    const data = workspace({ requirements: [...base.requirements, secondRequirement(base)] });
    const bytes = await createWorkspaceWorkbook(data, "0.7.28");
    const edited = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      const first = Array.from({ length: 18 }, (_, index) => sheet.getCell(2, index + 1).value);
      const second = Array.from({ length: 18 }, (_, index) => sheet.getCell(3, index + 1).value);
      first.forEach((value, index) => { sheet.getCell(3, index + 1).value = value; });
      second.forEach((value, index) => { sheet.getCell(2, index + 1).value = value; });
      sheet.getCell(2, headerColumn(sheet, "需求标题")).value = "午睡识别已排序";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "已排序.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "excel-wins");
    expect(applied.requirements.find((item) => item.id === "r2")?.title).toBe("午睡识别已排序");
    expect(applied.requirements.find((item) => item.id === "r1")?.title).toBe("社交时差");
  });

  it("Excel 保留本机医疗节点，允许修改送检月份为倒序或明确清空", async () => {
    const base = workspace();
    const data = workspace({ requirements: [{
      ...base.requirements[0],
      source: "NMPA认证",
      targetMonth: "2027-03",
      medicalSubmissionMilestones: { registrationStartMonth: "2027-02", clinicalStartMonth: "2027-04" },
    }] });
    const bytes = await createWorkspaceWorkbook(data, "0.8.12");
    const edited = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("需求清单")!.getCell(2, headerColumn(workbook.getWorksheet("需求清单")!, "上线月份")).value = new Date(2027, 0, 1, 12);
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "医疗节点倒序.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "excel-wins");
    expect(applied.requirements[0].targetMonth).toBe("2027-01");
    expect(applied.requirements[0].medicalSubmissionMilestones).toEqual(data.requirements[0].medicalSubmissionMilestones);
    const cleared = await editWorkbook(bytes, (workbook) => { workbook.getWorksheet("需求清单")!.getCell("J2").value = null; });
    const clearPreview = await inspectWorkspaceWorkbook(cleared, "送检日期清空.xlsx", data);
    expect(clearPreview.preview.errors).toEqual([]);
    expect(clearPreview.preview.differences.some((diff) => diff.fields.some((field) => field.field === "上线月份"))).toBe(true);
    const clearApplied = applyPreparedWorkspaceWorkbook(data, clearPreview, "excel-wins");
    expect(clearApplied.requirements[0].targetMonth).toBe("");
    expect(clearApplied.requirements[0].medicalSubmissionMilestones).toEqual(data.requirements[0].medicalSubmissionMilestones);
    expect(data.requirements[0].targetMonth).toBe("2027-03");
    expect(data.requirements[0].medicalSubmissionMilestones).toEqual({ registrationStartMonth: "2027-02", clinicalStartMonth: "2027-04" });
  });

  it("Excel 范围导出排除空送检版本；非法月份或普通需求月份清空仍整次失败", async () => {
    const base = workspace();
    const medical = { ...base.requirements[0], id: "medical", source: "CE认证" as const, targetMonth: "" };
    const data = workspace({ requirements: [...base.requirements, medical] });
    const bytes = await createWorkspaceWorkbook(data, "0.8.16", undefined, { start: "2027H1", end: "2027H1" });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes as never);
    expect(wb.worksheets).toHaveLength(1);
    expect(wb.worksheets[0].rowCount).toBe(2);
    expect(wb.worksheets[0].columnCount).toBe(18);
    for (const value of [null, "2027-13"]) {
      const invalid = await editWorkbook(bytes, (book) => { book.getWorksheet("需求清单")!.getCell("J2").value = value; });
      const prepared = await inspectWorkspaceWorkbook(invalid, "日期无效.xlsx", data);
      expect(prepared.preview.errors.length).toBeGreaterThan(0);
      expect(() => applyPreparedWorkspaceWorkbook(data, prepared, "excel-wins")).toThrow(/存在校验错误/);
      expect(data.requirements[0].targetMonth).toBe("2027-03");
    }
    const medBytes = await createWorkspaceWorkbook(workspace({ requirements: [medical] }), "0.8.16");
    const badMedical = await editWorkbook(medBytes, (book) => { book.getWorksheet("需求清单")!.getCell("J2").value = "bad-date"; });
    const bad = await inspectWorkspaceWorkbook(badMedical, "非法送检月份.xlsx", data);
    expect(bad.preview.errors.length).toBeGreaterThan(0);
  });

  it("只有操作列显式选择删除才删除需求", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.28");
    const cleared = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.spliceRows(2, 1);
    });
    const clearPrepared = await inspectWorkspaceWorkbook(cleared, "删除整行.xlsx", data);
    expect(clearPrepared.preview.errors).toEqual([]);
    expect(applyPreparedWorkspaceWorkbook(data, clearPrepared, "local-wins").requirements).toHaveLength(1);

    const explicit = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("需求清单")!.getCell("A2").value = "删除";
    });
    const deletePrepared = await inspectWorkspaceWorkbook(explicit, "显式删除.xlsx", data);
    expect(deletePrepared.preview.counts.deleted).toBe(1);
    expect(applyPreparedWorkspaceWorkbook(data, deletePrepared, "excel-wins").requirements).toEqual([]);
  });

  it("拒绝新增行、未知 ID 和损坏的内部系统数据", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.28");
    const added = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      for (let column = 1; column <= 18; column += 1) sheet.getCell(3, column).value = sheet.getCell(2, column).value;
    });
    expect((await inspectWorkspaceWorkbook(added, "新增行.xlsx", data)).preview.errors.some((item) => item.message.includes("不支持新增需求"))).toBe(true);

    const unknown = await editPayload(bytes, (value) => { value.rows[0].requirementId = "unknown"; });
    expect((await inspectWorkspaceWorkbook(unknown, "未知ID.xlsx", data)).preview.errors.some((item) => item.message.includes("内部需求 ID 不存在"))).toBe(true);

    const badSystem = await editPayload(bytes, (value) => { value.rows[0].baselineHash = "broken"; });
    await expect(inspectWorkspaceWorkbook(badSystem, "损坏系统数据.xlsx", data)).rejects.toThrow("内部系统信息缺失或损坏");

    const zip = await JSZip.loadAsync(bytes);
    zip.remove("customXml/item1.xml");
    const missingSystem = await zip.generateAsync({ type: "nodebuffer" });
    await expect(inspectWorkspaceWorkbook(missingSystem, "缺少系统数据.xlsx", data)).rejects.toThrow("内部系统信息缺失或损坏");
  });

  it("检测导出后本机与 Excel 的并发修改", async () => {
    const baseline = workspace();
    const bytes = await createWorkspaceWorkbook(baseline, "0.7.28");
    const edited = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("需求清单")!.getCell("B2").value = "Excel版本";
    });
    const local = workspace({
      requirements: [{ ...baseline.requirements[0], title: "本机版本", updatedAt: "2026-08-30T00:00:00.000Z" }],
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "冲突.xlsx", local);
    expect(prepared.preview.counts.conflicts).toBe(1);
    expect(applyPreparedWorkspaceWorkbook(local, prepared, "local-wins").requirements[0].title).toBe("本机版本");
    expect(applyPreparedWorkspaceWorkbook(local, prepared, "excel-wins").requirements[0].title).toBe("Excel版本");
  });

  it("校验领域关联、产品单选、范围、升级选项、月份和工作量", async () => {
    const data = workspace({
      domains: [
        ...workspace().domains,
        { id: "l0-2", name: "运动", sortOrder: 1, active: true, level: "L0" },
      ],
      products: [
        { id: "p1", name: "手表", sortOrder: 0, active: true },
        { id: "p2", name: "手环", sortOrder: 1, active: true },
      ],
    });
    const bytes = await createWorkspaceWorkbook(data, "0.7.28");
    const invalid = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell("D2").value = "运动";
      sheet.getCell("G2").value = "其他";
      sheet.getCell("H2").value = "可能";
      sheet.getCell("J2").value = "2027-13";
      sheet.getCell("K2").value = "手表、手环";
      sheet.getCell("L2").value = -1;
    });
    const prepared = await inspectWorkspaceWorkbook(invalid, "非法内容.xlsx", data);
    const messages = prepared.preview.errors.map((item) => item.message).join("\n");
    expect(messages).toContain("不属于领域 L0");
    expect(messages).toContain("适用范围");
    expect(messages).toContain("允许最近版本升级");
    expect(messages).toContain("上线月份");
    expect(messages).toContain("只能选择一个");
    expect(messages).toContain("工作量");
  });

  it("工作量允许留空或填写 0，空工作区也可往返", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.28");
    const edited = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell("L2").value = null;
      sheet.getCell("M2").value = 0;
      sheet.getCell("N2").value = null;
    });
    expect((await inspectWorkspaceWorkbook(edited, "零工作量.xlsx", data)).preview.errors).toEqual([]);

    const empty = workspace({ requirements: [], domains: [], products: [], groupOverrides: [], members: [], pptAccessPolicy: { entries: [], updatedAt: "1970-01-01T00:00:00.000Z" } });
    const emptyBytes = await createWorkspaceWorkbook(empty, "0.7.28");
    const emptyPrepared = await inspectWorkspaceWorkbook(emptyBytes, "空工作区.xlsx", empty);
    expect(emptyPrepared.preview.errors).toEqual([]);
    expect(emptyPrepared.preview.counts).toMatchObject({ added: 0, updated: 0, deleted: 0, unchanged: 0 });
  });

  it("范围辅助函数保留完整字典、卡片范围、人员和白名单", () => {
    const base = workspace();
    const data = workspace({
      requirements: [...base.requirements, secondRequirement(base)],
      groupOverrides: [
        ...base.groupOverrides,
        { groupKey: "d1::健康基础::2028-01", cardTitle: "午睡", cardSummary: "", sortOrder: 0, updatedAt: "2026-08-21T00:00:00.000Z" },
      ],
    });
    const ranged = workspaceDataForWorkbookRange(data, "2027H1", "2027H1");
    expect(ranged.requirements.map((item) => item.id)).toEqual(["r1"]);
    expect(ranged.groupOverrides.map((item) => item.groupKey)).toEqual(["d1::健康基础::2027-03"]);
    expect(ranged.domains).toEqual(data.domains);
    expect(ranged.products).toEqual(data.products);
    expect(ranged.members).toEqual(data.members);
    expect(ranged.pptAccessPolicy).toEqual(data.pptAccessPolicy);
  });
});

describe("旧版协作 Excel 兼容", () => {
  it.each(["v2", "v3", "v4", "v5"] as const)("兼容 %s 多 Sheet 文件并保留新版本机字段", async (version) => {
    const data = workspace({
      requirements: [{ ...workspace().requirements[0], scope: "仅国内", allowLatestVersionUpgrade: true }],
    });
    const legacy = await createLegacyWorkbook(data, version);
    const prepared = await inspectWorkspaceWorkbook(legacy, `${version}.xlsx`, data);
    expect(prepared.preview.formatVersion).toBe(version);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts).toMatchObject({ unchanged: 1, conflicts: 0 });
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins");
    expect(applied.requirements[0].scope).toBe("仅国内");
    expect(applied.requirements[0].allowLatestVersionUpgrade).toBe(true);
    expect(applied.members).toEqual(data.members);
    expect(applied.pptAccessPolicy).toEqual(data.pptAccessPolicy);
  });
});

async function createLegacyWorkbook(data: WorkspaceData, version: "v2" | "v3" | "v4" | "v5"): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const exportedAt = "2026-08-20T08:00:00.000Z";
  const guide = workbook.addWorksheet("填写说明");
  guide.getCell("B13").value = "health-roadmap-collaboration";
  guide.getCell("B14").value = version;
  guide.getCell("B15").value = exportedAt;

  const l0 = data.domains.find((item) => item.level === "L0")!;
  const l1 = data.domains.find((item) => (item.level ?? "L1") === "L1")!;
  writeLegacyDictionary(workbook.addWorksheet("领域L0字典"), l0);
  writeLegacyDomainL1(workbook.addWorksheet("领域L1字典"), l1, l0);
  writeLegacyDictionary(workbook.addWorksheet("产品字典"), data.products[0]);
  writeLegacyRequirements(workbook.addWorksheet("需求清单"), data, version);
  writeLegacyOverrides(workbook.addWorksheet("路标卡片"), data, version);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function writeLegacyDictionary(sheet: ExcelJS.Worksheet, item: DictionaryItem): void {
  sheet.getCell("A1").value = "旧版字典";
  sheet.getCell("A2").value = "旧版说明";
  sheet.getRow(3).values = ["名称", "状态", "排序", "字典ID（请勿修改）", "基线校验（请勿修改）"];
  sheet.getRow(4).values = [item.name, item.active ? "启用" : "停用", item.sortOrder, item.id, dictionaryHash(item)];
}

function writeLegacyDomainL1(sheet: ExcelJS.Worksheet, item: DictionaryItem, parent: DictionaryItem): void {
  sheet.getCell("A1").value = "旧版字典";
  sheet.getCell("A2").value = "旧版说明";
  sheet.getRow(3).values = [
    "名称",
    "所属领域L0",
    "状态",
    "排序",
    "字典ID（请勿修改）",
    "所属L0 ID（请勿修改）",
    "基线L0名称（请勿修改）",
    "基线校验（请勿修改）",
  ];
  sheet.getRow(4).values = [item.name, parent.name, item.active ? "启用" : "停用", item.sortOrder, item.id, parent.id, parent.name, dictionaryHash(item)];
}

function writeLegacyRequirements(sheet: ExcelJS.Worksheet, data: WorkspaceData, version: "v2" | "v3" | "v4" | "v5"): void {
  sheet.getCell("A1").value = "旧版需求清单";
  sheet.getCell("A2").value = "旧版说明";
  const item = data.requirements[0];
  const l0 = data.domains.find((entry) => entry.id === item.domainL0Id)!;
  const l1 = data.domains.find((entry) => entry.id === item.domainId)!;
  const product = data.products.find((entry) => item.productIds.includes(entry.id));
  const headers = version === "v2" ? LEGACY_REQUIREMENT_HEADERS.filter((header) => header !== "适用范围") : [...LEGACY_REQUIREMENT_HEADERS];
  sheet.getRow(3).values = headers;
  const values: Record<string, ExcelJS.CellValue> = {
    操作: "保留",
    需求标题: item.title,
    需求描述: item.description,
    领域L0: l0.name,
    领域L1: l1.name,
    来源: item.source,
    适用范围: item.scope,
    分类: item.category,
    上线月份: new Date(2027, 2, 1, 12),
    "匹配产品（单选）": product?.name ?? "",
    "设备工作量（人月）": item.deviceWorkloadPm,
    "App工作量（人月）": item.appWorkloadPm,
    "云侧工作量（人月）": item.cloudWorkloadPm,
    "待拆分工作量（只读）": item.unallocatedWorkloadPm,
    "总工作量（自动合计）": item.workloadPm,
    "图片数量（只读）": item.images.length,
    "需求ID（请勿修改）": item.id,
    "领域L0 ID（请勿修改）": item.domainL0Id,
    "领域L1 ID（请勿修改）": item.domainId,
    "产品ID（请勿修改）": item.productIds.join("、"),
    "基线领域L0名称（请勿修改）": l0.name,
    "基线领域L1名称（请勿修改）": l1.name,
    "基线产品名称（请勿修改）": product?.name ?? "",
    "创建时间（请勿修改）": item.createdAt,
    "基线更新时间（请勿修改）": item.updatedAt,
    "基线校验（请勿修改）": version === "v2" ? requirementHashV2(item) : requirementHashLegacy(item),
  };
  sheet.getRow(4).values = headers.map((header) => values[header]);
}

function writeLegacyOverrides(sheet: ExcelJS.Worksheet, data: WorkspaceData, version: "v2" | "v3" | "v4" | "v5"): void {
  sheet.getCell("A1").value = "旧版路标卡片";
  sheet.getCell("A2").value = "旧版说明";
  const item = data.groupOverrides[0];
  const [domainId, source, targetMonth] = item.groupKey.split("::");
  const domain = data.domains.find((entry) => entry.id === domainId)!;
  if (version === "v5") {
    sheet.getRow(3).values = [
      "领域 L1 标签（自动）",
      "来源",
      "上线月份",
      "卡片顺序（请勿修改）",
      "卡片键（请勿修改）",
      "基线更新时间（请勿修改）",
      "基线校验（请勿修改）",
    ];
    sheet.getRow(4).values = [domain.name, source, new Date(`${targetMonth}-01T12:00:00.000Z`), item.sortOrder ?? null, item.groupKey, item.updatedAt, overrideOrderHash(item)];
    return;
  }
  const withOrder = version === "v4";
  sheet.getRow(3).values = [
    "操作",
    "卡片标题",
    "详情/PPT摘要",
    "领域",
    "来源",
    "上线月份",
    ...(withOrder ? ["卡片顺序（请勿修改）"] : []),
    "卡片键（请勿修改）",
    "基线更新时间（请勿修改）",
    "基线校验（请勿修改）",
  ];
  sheet.getRow(4).values = [
    "保留",
    item.cardTitle,
    item.cardSummary,
    domain.name,
    source,
    new Date(`${targetMonth}-01T12:00:00.000Z`),
    ...(withOrder ? [item.sortOrder ?? null] : []),
    item.groupKey,
    item.updatedAt,
    withOrder ? overrideHash(item) : overrideHashV3(item),
  ];
}

function stableHash(value: unknown): string {
  return createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

function dictionaryHash(item: DictionaryItem): string {
  return stableHash({
    id: item.id,
    name: item.name.trim(),
    sortOrder: item.sortOrder,
    active: item.active,
    level: item.level ?? null,
    parentId: item.parentId ?? null,
  });
}

function requirementHashLegacy(item: Requirement): string {
  return stableHash({
    id: item.id,
    title: item.title.trim(),
    description: item.description.trim(),
    domainL0Id: item.domainL0Id,
    domainId: item.domainId,
    source: item.source,
    scope: item.scope,
    category: item.category,
    targetMonth: item.targetMonth,
    productIds: [...new Set(item.productIds)].sort(),
    overseasRegions: [...new Set(item.overseasRegions)].sort(),
    deviceWorkloadPm: item.deviceWorkloadPm,
    appWorkloadPm: item.appWorkloadPm,
    cloudWorkloadPm: item.cloudWorkloadPm,
    unallocatedWorkloadPm: item.unallocatedWorkloadPm,
  });
}

function requirementHashV2(item: Requirement): string {
  const { scope: _scope, ...withoutScope } = {
    id: item.id,
    title: item.title.trim(),
    description: item.description.trim(),
    domainL0Id: item.domainL0Id,
    domainId: item.domainId,
    source: item.source,
    scope: item.scope,
    category: item.category,
    targetMonth: item.targetMonth,
    productIds: [...new Set(item.productIds)].sort(),
    overseasRegions: [...new Set(item.overseasRegions)].sort(),
    deviceWorkloadPm: item.deviceWorkloadPm,
    appWorkloadPm: item.appWorkloadPm,
    cloudWorkloadPm: item.cloudWorkloadPm,
    unallocatedWorkloadPm: item.unallocatedWorkloadPm,
  };
  return stableHash(withoutScope);
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
