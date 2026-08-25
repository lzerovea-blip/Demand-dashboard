import { createHash } from "node:crypto";
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
      scope: "全球",
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
    groupOverrides: [{ groupKey: "d1::健康基础::2027-03", cardTitle: "睡眠", cardSummary: "摘要", sortOrder: 0, updatedAt: "2026-08-20T00:00:00.000Z" }],
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

function writeLegacyRoadmapSheet(workbook: ExcelJS.Workbook, data: WorkspaceData, version: "v2" | "v3" | "v4"): void {
  const sheet = workbook.getWorksheet("路标卡片")!;
  sheet.removeTable("RoadmapCardsTable");
  const withOrder = version === "v4";
  const headers = [
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
  sheet.getRow(3).values = headers;
  const domains = new Map(data.domains.map((item) => [item.id, item.name]));
  data.groupOverrides.forEach((item, index) => {
    const [domainId, source, targetMonth] = item.groupKey.split("::");
    sheet.getRow(4 + index).values = [
      "保留",
      item.cardTitle,
      item.cardSummary,
      domains.get(domainId) ?? domainId,
      source,
      new Date(`${targetMonth}-01T00:00:00.000Z`),
      ...(withOrder ? [item.sortOrder ?? null] : []),
      item.groupKey,
      item.updatedAt,
      withOrder ? v4OverrideHash(item) : legacyOverrideHash(item),
    ];
  });
}

describe("企业协作 Excel", () => {
  it("生成标准工作簿并无变化往返", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0", "2026-08-20T08:00:00.000Z");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes as never);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["填写说明", "领域L0字典", "领域L1字典", "产品字典", "需求清单", "路标卡片"]);
    expect(workbook.getWorksheet("填写说明")?.getCell("B13").value).toBe("health-roadmap-collaboration");
    expect(workbook.getWorksheet("填写说明")?.getCell("B14").value).toBe("v5");
    expect(workbook.getWorksheet("需求清单")?.getRow(3).values).not.toContain("海外区域（用、分隔）");
    expect(workbook.getWorksheet("需求清单")?.getRow(3).values).toEqual(expect.arrayContaining(["领域L0", "领域L1", "适用范围", "匹配产品（单选）"]));
    expect(workbook.getWorksheet("领域L1字典")?.getCell("B4").value).toBe("健康");
    expect(workbook.getWorksheet("路标卡片")?.getRow(3).values).not.toEqual(expect.arrayContaining(["卡片标题", "详情/PPT摘要"]));
    expect(workbook.getWorksheet("路标卡片")?.getCell(4, headerColumn(workbook.getWorksheet("路标卡片")!, "领域 L1 标签（自动）")).value).toBe("睡眠");
    expect(workbook.getWorksheet("路标卡片")?.getCell(4, headerColumn(workbook.getWorksheet("路标卡片")!, "卡片顺序（请勿修改）")).value).toBe(0);
    expect(workbook.getWorksheet("需求清单")?.getCell("O4").value).toMatchObject({ formula: "SUM(K4:N4)", result: 1 });

    const prepared = await inspectWorkspaceWorkbook(bytes, "协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts).toMatchObject({ added: 0, updated: 0, deleted: 0, unchanged: 1, conflicts: 0 });
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins")).toEqual(data);
  });

  it("按半年区间导出需求和卡片并保留完整字典，回导不影响范围外数据", async () => {
    const base = workspace();
    const outsideRequirement = {
      ...base.requirements[0],
      id: "r2",
      title: "范围外需求",
      category: "体验优化" as const,
      targetMonth: "2028-01",
      productIds: [],
    };
    const data = workspace({
      requirements: [...base.requirements, outsideRequirement],
      products: [...base.products, { id: "p2", name: "手环", sortOrder: 1, active: true }],
      groupOverrides: [
        ...base.groupOverrides,
        { groupKey: "d1::健康基础::2028-01", cardTitle: "范围外卡片", cardSummary: "范围外摘要", sortOrder: 0, updatedAt: "2026-08-20T00:00:00.000Z" },
      ],
    });
    const bytes = await createWorkspaceWorkbook(data, "0.7.13", "2026-08-24T08:00:00.000Z", { start: "2027H1", end: "2027H1" });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes as never);

    expect(workbook.getWorksheet("填写说明")?.getCell("B18").value).toBe("2027H1 – 2027H1");
    expect(workbook.getWorksheet("需求清单")?.getCell("B4").value).toBe("社交时差");
    expect(workbook.getWorksheet("需求清单")?.getCell("B5").value).toBeNull();
    expect(workbook.getWorksheet("路标卡片")?.getCell(4, headerColumn(workbook.getWorksheet("路标卡片")!, "卡片键（请勿修改）")).value).toBe("d1::健康基础::2027-03");
    expect(workbook.getWorksheet("路标卡片")?.getCell(5, headerColumn(workbook.getWorksheet("路标卡片")!, "卡片键（请勿修改）")).value).toBeNull();
    expect(workbook.getWorksheet("产品字典")?.getCell("A5").value).toBe("手环");

    const prepared = await inspectWorkspaceWorkbook(bytes, "区间协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts).toMatchObject({ unchanged: 1, added: 0, updated: 0, deleted: 0 });
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
        适用范围: "全球",
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

  it("旧版 AI 来源回导时迁移为 AI综合", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.3");
    const legacyAi = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "来源")).value = "AI";
    });
    const prepared = await inspectWorkspaceWorkbook(legacyAi, "旧AI协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.parsed.requirements[0].requirement.source).toBe("AI综合");
  });

  it("旧版医疗送检来源回导时迁移为 NMPA送检", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.8");
    const legacyMedical = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "来源")).value = "医疗送检";
    });
    const prepared = await inspectWorkspaceWorkbook(legacyMedical, "旧医疗送检协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.parsed.requirements[0].requirement.source).toBe("NMPA送检");
  });

  it("兼容 v2 协作表并默认全球且不产生虚假冲突", async () => {
    const data = workspace({ requirements: [{ ...workspace().requirements[0], scope: "仅国内" }] });
    const bytes = await createWorkspaceWorkbook(data, "0.7.2");
    const legacy = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("填写说明")!.getCell("B14").value = "v2";
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.spliceColumns(headerColumn(sheet, "适用范围"), 1);
      sheet.getCell(4, headerColumn(sheet, "基线校验（请勿修改）")).value = legacyRequirementHash(data.requirements[0]);
      writeLegacyRoadmapSheet(workbook, data, "v2");
    });
    const prepared = await inspectWorkspaceWorkbook(legacy, "旧版协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.formatVersion).toBe("v2");
    expect(prepared.preview.counts).toMatchObject({ unchanged: 1, conflicts: 0 });
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins").requirements[0].scope).toBe("仅国内");
  });

  it("兼容 v3 协作表并保留本机卡片顺序", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.5");
    const legacy = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("填写说明")!.getCell("B14").value = "v3";
      writeLegacyRoadmapSheet(workbook, data, "v3");
    });
    const prepared = await inspectWorkspaceWorkbook(legacy, "v3协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts.conflicts).toBe(0);
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins").groupOverrides[0].sortOrder).toBe(0);
  });

  it("兼容 v4 卡片标题摘要并保留卡片顺序", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.20");
    const legacy = await editWorkbook(bytes, (workbook) => {
      workbook.getWorksheet("填写说明")!.getCell("B14").value = "v4";
      writeLegacyRoadmapSheet(workbook, data, "v4");
    });
    const prepared = await inspectWorkspaceWorkbook(legacy, "v4协作.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.preview.counts.conflicts).toBe(0);
    expect(applyPreparedWorkspaceWorkbook(data, prepared, "local-wins").groupOverrides[0]).toEqual(data.groupOverrides[0]);
  });

  it("三端工作量允许留空或显式填写0", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.1.0");
    const missing = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "云侧工作量（人月）")).value = null;
    });
    const missingPrepared = await inspectWorkspaceWorkbook(missing, "缺少工作量.xlsx", data);
    expect(missingPrepared.preview.errors).toEqual([]);

    const explicitZero = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "云侧工作量（人月）")).value = 0;
    });
    expect((await inspectWorkspaceWorkbook(explicitZero, "零工作量.xlsx", data)).preview.errors).toEqual([]);
  });

  it("拒绝缺失或非法的适用范围", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.2");
    const invalid = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "适用范围")).value = "其他";
    });
    const prepared = await inspectWorkspaceWorkbook(invalid, "非法范围.xlsx", data);
    expect(prepared.preview.errors.some((item) => item.message.includes("适用范围"))).toBe(true);
  });

  it("三类需求都只能匹配一个产品", async () => {
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

    const missing = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "匹配产品（单选）")).value = "";
    });
    const missingPrepared = await inspectWorkspaceWorkbook(missing, "产品专属缺少产品.xlsx", data);
    expect(missingPrepared.preview.errors.some((item) => item.message.includes("必须且只能匹配一个产品"))).toBe(true);
  });

  it("平台演进可通过 Excel 往返并可选择一个产品", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.3");
    const valid = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "分类")).value = "平台演进";
      sheet.getCell(4, headerColumn(sheet, "匹配产品（单选）")).value = "";
    });
    const prepared = await inspectWorkspaceWorkbook(valid, "平台演进.xlsx", data);
    expect(prepared.preview.errors).toEqual([]);
    expect(prepared.parsed.requirements[0].requirement.category).toBe("平台演进");
    expect(prepared.parsed.requirements[0].requirement.productIds).toEqual([]);

    const withProduct = await editWorkbook(valid, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "匹配产品（单选）")).value = "手表";
    });
    const accepted = await inspectWorkspaceWorkbook(withProduct, "平台演进产品.xlsx", data);
    expect(accepted.preview.errors).toEqual([]);
    expect(accepted.parsed.requirements[0].requirement.productIds).toEqual(["p1"]);
  });

  it("体验优化不能通过 Excel 匹配产品", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.18");
    const edited = await editWorkbook(bytes, (workbook) => {
      const sheet = workbook.getWorksheet("需求清单")!;
      sheet.getCell(4, headerColumn(sheet, "分类")).value = "体验优化";
      sheet.getCell(4, headerColumn(sheet, "匹配产品（单选）")).value = "手表";
    });
    const prepared = await inspectWorkspaceWorkbook(edited, "体验优化产品.xlsx", data);
    expect(prepared.preview.errors.some((item) => item.message.includes("体验优化需求不匹配产品"))).toBe(true);
    expect(prepared.parsed.requirements[0].requirement.productIds).toEqual([]);
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

function legacyRequirementHash(item: WorkspaceData["requirements"][number]): string {
  return createHash("sha256").update(JSON.stringify({
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
    deviceWorkloadPm: item.deviceWorkloadPm,
    appWorkloadPm: item.appWorkloadPm,
    cloudWorkloadPm: item.cloudWorkloadPm,
    unallocatedWorkloadPm: item.unallocatedWorkloadPm,
  })).digest("hex");
}

function legacyOverrideHash(item: WorkspaceData["groupOverrides"][number]): string {
  return createHash("sha256").update(JSON.stringify({
    groupKey: item.groupKey,
    cardTitle: item.cardTitle.trim(),
    cardSummary: item.cardSummary.trim(),
  })).digest("hex");
}

function v4OverrideHash(item: WorkspaceData["groupOverrides"][number]): string {
  return createHash("sha256").update(JSON.stringify({
    groupKey: item.groupKey,
    cardTitle: item.cardTitle.trim(),
    cardSummary: item.cardSummary.trim(),
    sortOrder: item.sortOrder ?? null,
  })).digest("hex");
}
