import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import type { Member, Requirement, WorkspaceData } from "../src/shared/types";
import { BUILT_IN_TIERING_PRODUCT_SERIES, normalizeWorkspaceTieringScope } from "../src/shared/tiering";
import { createWorkspaceWorkbook, inspectWorkspaceWorkbook, applyPreparedWorkspaceWorkbook } from "../src/main/workspaceWorkbook";

const now = "2026-10-10T00:00:00.000Z";
const members: Member[] = (["规划", "产品经理", "SE"] as const).map((role, index) => ({ id: `m${index}`, name: role, employeeNumber: `E${index}`, microsoftAccount: `m${index}@example.com`, role, domainL1Ids: ["l1"], sortOrder: index, createdAt: now, updatedAt: now }));
const baseRequirement: Requirement = {
  id: "existing-1", title: "原需求", description: "描述", images: [], domainL0Id: "l0", domainId: "l1", scope: "全球", allowLatestVersionUpgrade: false,
  overseasRegions: [], category: "平台演进", targetMonth: "2027-03", productIds: [], deviceWorkloadPm: 1, appWorkloadPm: 0, cloudWorkloadPm: 0,
  researchAppWorkloadPm: 0, algorithmWorkloadPm: 0, unallocatedWorkloadPm: 0, workloadPm: 1,
  planningMemberId: "m0", productManagerMemberId: "m1", seMemberId: "m2", createdAt: now, updatedAt: now,
};
function workspace(requirement = baseRequirement): WorkspaceData {
  return normalizeWorkspaceTieringScope({ requirements: [requirement], domains: [
    { id: "l0", name: "健康", level: "L0", sortOrder: 0, active: true },
    { id: "l1", name: "睡眠", level: "L1", parentId: "l0", sortOrder: 0, active: true },
  ], products: [], groupOverrides: [], members, pptAccessPolicy: { entries: [], updatedAt: now } });
}
async function exportBook(data = workspace()) {
  const bytes = await createWorkspaceWorkbook(data, "0.9.2", now, { start: "2027H1", end: "2027H1" });
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(bytes as never);
  return { bytes, book, sheet: book.worksheets[0] };
}
async function inspect(book: ExcelJS.Workbook, data = workspace()) {
  return inspectWorkspaceWorkbook(Buffer.from(await book.xlsx.writeBuffer()), "在线回收.xlsx", data);
}

describe("v12 可见需求编号在线协作", () => {
  it.each([0.01, 0.29, 0.57, 1.13, 2.55, 99999999])("原样导出回导合法工作量 %s 不报错", async (amount) => {
    const data = workspace({ ...baseRequirement, deviceWorkloadPm: amount, appWorkloadPm: 0, cloudWorkloadPm: 0, workloadPm: amount });
    const { book } = await exportBook(data);
    const preview = await inspect(book, data);
    expect(preview.preview.errors).toEqual([]);
    expect(preview.preview.fieldChanges).toEqual([]);
  });

  it("Excel 将三项工作量保存为数字后仍可原样回导", async () => {
    const data = workspace({ ...baseRequirement, deviceWorkloadPm: 0.29, appWorkloadPm: 0.57, cloudWorkloadPm: 1.13, workloadPm: 1.99 });
    const { book, sheet } = await exportBook(data);
    sheet.getCell("L2").value = 0.29;
    sheet.getCell("M2").value = 0.57;
    sheet.getCell("N2").value = 1.13;
    const preview = await inspect(book, data);
    expect(preview.preview.errors).toEqual([]);
    expect(preview.preview.fieldChanges).toEqual([]);
  });

  it.each(["-0.01", "0.001", "不是数字", "100000000"])("无效工作量 %s 仍被拒绝", async (amount) => {
    const data = workspace();
    const { book, sheet } = await exportBook(data);
    sheet.getCell("L2").value = amount;
    const preview = await inspect(book, data);
    expect(preview.preview.errors).toContainEqual({ sheet: "需求清单", cell: "L2", message: "工作量须为非负数字，最多两位小数" });
  });

  it("只依赖业务工作表，预留 300 行且复制到新工作簿仍可识别", async () => {
    const data = workspace();
    const { bytes, sheet } = await exportBook(data);
    expect(sheet.name).toBe("需求清单");
    expect(sheet.getCell("A1").text).toBe("序号");
    expect(sheet.getCell("B1").text).toBe("需求编号");
    expect(sheet.getCell("B2").text).toBe("existing-1");
    expect(sheet.getCell("C3").text).toBe("新增");
    expect(sheet.getCell("B302").text).toMatch(/^NEW-/);
    const zip = await JSZip.loadAsync(bytes);
    expect(zip.file("customXml/item1.xml")).toBeNull();
    expect(zip.file("docProps/custom.xml")).toBeNull();
    const copied = new ExcelJS.Workbook();
    const target = copied.addWorksheet("公司在线下载");
    for (let row = 1; row <= 2; row += 1) target.addRow((sheet.getRow(row).values as ExcelJS.CellValue[]).slice(1));
    target.getCell("D2").value = "在线修改标题";
    const preview = await inspect(copied, data);
    expect(preview.preview.errors).toEqual([]);
    expect(preview.preview.fieldChanges).toMatchObject([{ id: "existing-1:field:title", kind: "field", conflict: true }]);
  });

  it("明确删除和行顺序变化可勾选合入，重导不新增副本", async () => {
    const data = workspace();
    const { book, sheet } = await exportBook(data);
    sheet.getCell("C2").value = "删除";
    const preview = await inspect(book, data);
    expect(preview.preview.errors).toEqual([]);
    expect(preview.preview.fieldChanges?.map((item) => item.id)).toEqual(["existing-1:delete"]);
    const result = applyPreparedWorkspaceWorkbook(data, preview, "local-wins", now, { selectedChangeIds: ["existing-1:delete"], conflictDecisions: { "existing-1:delete": "excel" } });
    expect(result.requirements).toHaveLength(0);
    expect(result.requirementTombstones?.[0].id).toBe("existing-1");
    expect(result.tiering?.features.some((item) => item.requirementId === "existing-1")).toBe(false);
  });

  it("新增缺少责任人先成为待完善需求，不进入分层；补齐后显式转正", async () => {
    const data = workspace();
    const { book, sheet } = await exportBook(data);
    const row = 3;
    sheet.getCell(`D${row}`).value = "新功能";
    sheet.getCell(`F${row}`).value = "健康";
    sheet.getCell(`G${row}`).value = "睡眠";
    sheet.getCell(`K${row}`).value = "2027-04";
    const first = await inspect(book, data);
    expect(first.preview.errors).toEqual([]);
    const id = sheet.getCell(`B${row}`).text;
    const draftData = applyPreparedWorkspaceWorkbook(data, first, "local-wins", now, { selectedChangeIds: [`${id}:add`], conflictDecisions: {} });
    expect(draftData.draftRequirements?.map((item) => item.id)).toContain(id);
    expect(draftData.requirements).toHaveLength(1);
    expect(draftData.tiering?.features.some((item) => item.requirementId === id)).toBe(false);
    sheet.getCell(`O${row}`).value = "规划";
    sheet.getCell(`P${row}`).value = "产品经理";
    sheet.getCell(`Q${row}`).value = "SE";
    const second = await inspect(book, draftData);
    expect(second.preview.errors).toEqual([]);
    expect(second.preview.fieldChanges?.some((item) => item.id === `${id}:promote`)).toBe(true);
    const choices = second.preview.fieldChanges!.filter((item) => item.requirementId === id).map((item) => item.id);
    const decisions = Object.fromEntries(choices.map((change) => [change, "excel" as const]));
    const formalData = applyPreparedWorkspaceWorkbook(draftData, second, "local-wins", now, { selectedChangeIds: choices, conflictDecisions: decisions });
    expect(formalData.draftRequirements).toHaveLength(0);
    expect(formalData.requirements.some((item) => item.id === id)).toBe(true);
    expect(formalData.tiering?.features.some((item) => item.requirementId === id)).toBe(true);
    expect(formalData.tiering?.supportRecords.filter((item) => item.featureId === `REQ-${id}`)).toHaveLength(BUILT_IN_TIERING_PRODUCT_SERIES.length);
  });
});
