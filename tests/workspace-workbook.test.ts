import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  applyPreparedWorkspaceWorkbook,
  createWorkspaceWorkbook,
  inspectWorkspaceWorkbook,
} from "../src/main/workspaceWorkbook";
import type { RequirementImage, WorkspaceData } from "../src/shared/types";

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
      overseasRegions: [],
      category: "产品专属",
      targetMonth: "2027-03",
      productIds: ["p1"],
      deviceWorkloadPm: 0.25,
      appWorkloadPm: 0.5,
      cloudWorkloadPm: 0.25,
      unallocatedWorkloadPm: 0,
      workloadPm: 1,
      createdAt: "2026-08-19T00:00:00.000Z",
      updatedAt: "2026-08-20T00:00:00.000Z",
    }],
    domains: [
      { id: "l0", name: "健康", sortOrder: 0, active: true, level: "L0" },
      { id: "d1", name: "睡眠", sortOrder: 0, active: true, level: "L1", parentId: "l0" },
    ],
    products: [{ id: "p1", name: "手表", sortOrder: 0, active: true }],
    groupOverrides: [{ groupKey: "d1::健康基础::2027-03", cardTitle: "睡眠", cardSummary: "摘要", updatedAt: "2026-08-20T00:00:00.000Z" }],
    ...overrides,
  };
}

async function editWorkbook(bytes: Buffer, edit: (workbook: ExcelJS.Workbook) => void): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes as never);
  edit(workbook);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

function headerColumn(sheet: ExcelJS.Worksheet, header: string): number {
  let result = 0;
  sheet.getRow(3).eachCell((cell, column) => {
    if (cell.text === header) result = column;
  });
  if (!result) throw new Error(`missing test header: ${header}`);
  return result;
}

describe("企业协作 Excel", () => {
  it("生成标准工作簿并无变化往返", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0", "2026-08-20T08:00:00.000Z");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes as never);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["填写说明", "领域L0字典", "领域L1字典", "产品字典", "需求清单", "路标卡片"]);
    expect(workbook.getWorksheet("填写说明")?.getCell("B13").value).toBe("health-roadmap-collaboration");
    expect(workbook.getWorksheet("填写说明")?.getCell("B14").value).toBe("v2");
    expect(workbook.getWorksheet("需求清单")?.getRow(3).values).not.toContain("海外区域（用、分隔）");
    expect(workbook.getWorksheet("需求清单")?.getRow(3).values).toEqual(expect.arrayContaining(["领域L0", "领域L1", "匹配产品（单选）"]));
    expect(workbook.getWorksheet("领域L1字典")?.getCell("B4").value).toBe("健康");
    expect(workbook.getWorksheet("需求清单")?.getCell("N4").value).toMatchObject({ formula: "SUM(J4:M4)", result: 1 });

    const prepared = await inspectWorkspaceWorkbook(bytes, "协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts).toMatchObject({ added: 0, updated: 0, deleted: 0, unchanged: 1, conflicts: 0 });
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins")).toEqual(data);
  });

  it("回导修改和新增需求，并保留同ID需求图片", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0", "2026-08-20T08:00:00.000Z");
    const edited = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "需求标题")).value = "社交时差优化";
      sheet.getCell(4, headerColumn(sheet, "App工作量（人月）")).value = 0.75;
      const row = 5;
      const values: Record<string, string | number | Date> = {
        操作: "保留",
        需求标题: "午睡识别",
        需求描述: "新增协作需求",
        领域L0: "健康",
        领域L1: "睡眠",
        来源: "健康进阶",
        分类: "体验优化",
        上线月份: new Date(2027, 4, 1, 12),
        "匹配产品（单选）": "",
        "设备工作量（人月）": 0.5,
        "App工作量（人月）": 0.5,
        "云侧工作量（人月）": 0,
        "待拆分工作量（只读）": 0,
      };
      Object.entries(values).forEach(([header, value]) => { sheet.getCell(row, headerColumn(sheet, header)).value = value; });
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts).toMatchObject({ added: 1, updated: 1, conflicts: 0 });
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins", "2026-08-21T00:00:00.000Z");
    expect(applied.requirements).toHaveLength(2);
    expect(applied.requirements.find((item) => item.id === "r1")).toMatchObject({ title: "社交时差优化", appWorkloadPm: 0.75, images: [image] });
    expect(applied.requirements.find((item) => item.title === "午睡识别")).toMatchObject({ targetMonth: "2027-05", images: [], workloadPm: 1 });
  });

  it("仅显式标记删除才删除需求", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0");
    const edited = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "操作")).value = "删除";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "协作.xlsx", data);
    expect(prepared.preview.counts.deleted).toBe(1);
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins").requirements).toEqual([]);
  });

  it("检测双方同时修改并支持两种冲突策略", async () => {
    const baseline = workspace();
    const bytes = await createWorkspaceWorkbook(baseline, "0.1.0");
    const edited = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "需求标题")).value = "Excel版本";
    });
    const local = workspace({ requirements: [{ ...baseline.requirements[0], title: "本机版本", updatedAt: "2026-08-21T00:00:00.000Z" }] });
    const prepared = await inspectWorkspaceWorkbook(edited, "协作.xlsx", local);
    expect(prepared.preview.counts.conflicts).toBe(1);
    expect(applyPreparedWorkspaceWorkbook(local, prepared, "local-wins").requirements[0].title).toBe("本机版本");
    const excelWins = applyPreparedWorkspaceWorkbook(local, prepared, "excel-wins", "2026-08-22T00:00:00.000Z");
    expect(excelWins.requirements[0].title).toBe("Excel版本");
    expect(excelWins.requirements[0].images).toEqual([image]);
  });

  it("支持字典改名且保持需求和卡片关联ID", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0");
    const edited = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("领域L1字典")!.getCell("A4").value = "睡眠健康";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins");
    expect(applied.domains.find((item) => item.id === "d1")?.name).toBe("睡眠健康");
    expect(applied.requirements[0].domainId).toBe("d1");
    expect(applied.groupOverrides[0].groupKey).toBe("d1::健康基础::2027-03");
  });

  it("在预览中报告非法月份和缺失表头，并禁止应用", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0");
    const invalidMonth = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "上线月份")).value = "2027-13";
    });
    const prepared = await inspectWorkspaceWorkbook(invalidMonth, "协作.xlsx", data);
    expect(prepared.preview.errors.some((item) => item.message.includes("上线月份"))).toBe(true);
    expect(() => applyPreparedWorkspaceWorkbook(data, prepared, "local-wins")).toThrow("存在校验错误");

    const missingHeader = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("需求清单")!.getCell("B3").value = "标题被修改";
    });
    const headerPrepared = await inspectWorkspaceWorkbook(missingHeader, "协作.xlsx", data);
    expect(headerPrepared.preview.errors.some((item) => item.message.includes("缺少必要列"))).toBe(true);
  });

  it("旧版海外研究来源回导时迁移为行业并清空旧区域", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0");
    const legacyWorkbook = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "来源")).value = "海外研究";
    });
    const prepared = await inspectWorkspaceWorkbook(legacyWorkbook, "旧版海外协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    const applied = applyPreparedWorkspaceWorkbook(data, prepared, "local-wins");
    expect(applied.requirements[0].source).toBe("行业");
    expect(applied.requirements[0].overseasRegions).toEqual([]);
  });

  it("三端工作量均为必填并允许显式填写0", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0");
    const missing = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "云侧工作量（人月）")).value = null;
    });
    const missingPrepared = await inspectWorkspaceWorkbook(missing, "缺少工作量.xlsx", data);
    expect(missingPrepared.preview.errors.some((item) => item.message.includes("三端工作量均为必填"))).toBe(true);

    const explicitZero = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "云侧工作量（人月）")).value = 0;
    });
    expect((await inspectWorkspaceWorkbook(explicitZero, "零工作量.xlsx", data)).preview.errors).toEqual([]);
  });

  it("产品专属只能匹配一个产品，体验优化不匹配产品", async () => {
    const data = workspace({ products: [
      { id: "p1", name: "手表", sortOrder: 0, active: true },
      { id: "p2", name: "手环", sortOrder: 1, active: true },
    ] });
    const bytes = await createWorkspaceWorkbook(data, "0.1.0");
    const edited = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "匹配产品（单选）")).value = "手表、手环";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "多产品.xlsx", data);
    expect(prepared.preview.errors.some((item) => item.message.includes("只能选择一个"))).toBe(true);
  });

  it("领域L1必须关联L0，需求中的两级领域必须匹配", async () => {
    const data = workspace({ domains: [
      { id: "l0", name: "健康", sortOrder: 0, active: true, level: "L0" },
      { id: "l0-2", name: "运动", sortOrder: 1, active: true, level: "L0" },
      { id: "d1", name: "睡眠", sortOrder: 0, active: true, level: "L1", parentId: "l0" },
    ] });
    const bytes = await createWorkspaceWorkbook(data, "0.1.0");
    const mismatched = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("领域L1字典")!.getCell("B4").value = "运动";
    });
    const prepared = await inspectWorkspaceWorkbook(mismatched, "领域不匹配.xlsx", data);
    expect(prepared.preview.errors.some((item) => item.message.includes("不属于领域 L0"))).toBe(true);
  });

  it("空工作区也能导出并重新检查", async () => {
    const empty = workspace({ requirements: [], domains: [], products: [], groupOverrides: [] });
    const bytes = await createWorkspaceWorkbook(empty, "0.1.0");
    const prepared = await inspectWorkspaceWorkbook(bytes, "空工作区.xlsx", empty);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts).toMatchObject({ added: 0, updated: 0, deleted: 0, unchanged: 0 });
  });
});
