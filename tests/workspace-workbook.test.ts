import { createHash } from "node:crypto";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  applyPreparedWorkspaceWorkbook,
  createWorkspaceWorkbook,
  inspectWorkspaceWorkbook,
  workspaceDataForWorkbookRange,
} from "../src/main/workspaceWorkbook";
import type { DictionaryItem, GroupOverride, Requirement, RequirementImage, WorkspaceData } from "../src/shared/types";

const V6_HEADERS = [
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

describe("协作 Excel v6 单 Sheet 修改版", () => {
  it("只导出需求清单，首行为表头且工作表仅包含 A-N 业务字段", async () => {
    const bytes = await createWorkspaceWorkbook(workspace(), "0.7.28", "2026-08-29T08:00:00.000Z", {
      start: "2027H1",
      end: "2027H2",
    });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes as never);
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["需求清单"]);
    const sheet = workbook.getWorksheet("需求清单")!;
    expect(sheet.getRow(1).values).toEqual([undefined, ...V6_HEADERS]);
    expect(sheet.views[0]).toMatchObject({ state: "frozen", xSplit: 2, ySplit: 1 });
    for (let column = 1; column <= 14; column += 1) expect(sheet.getColumn(column).hidden).not.toBe(true);
    for (let column = 15; column <= 30; column += 1) {
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
        formatVersion: "v6",
        appVersion: "0.7.28",
        range: { start: "2027H1", end: "2027H2" },
      },
    });
  });

  it("无修改往返保持所有本机数据", async () => {
    const data = workspace();
    const bytes = await createWorkspaceWorkbook(data, "0.7.28", "2026-08-29T08:00:00.000Z");
    const prepared = await inspectWorkspaceWorkbook(bytes, "需求修改.xlsx", data);
    expect(prepared.preview.formatVersion).toBe("v6");
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
      const first = Array.from({ length: 14 }, (_, index) => sheet.getCell(2, index + 1).value);
      const second = Array.from({ length: 14 }, (_, index) => sheet.getCell(3, index + 1).value);
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
      for (let column = 1; column <= 14; column += 1) sheet.getCell(3, column).value = sheet.getCell(2, column).value;
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
