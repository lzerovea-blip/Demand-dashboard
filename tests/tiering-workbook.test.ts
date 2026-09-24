import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { createTieringWorkbook, tieringDemandsForRange } from "../src/main/tieringWorkbook";
import { BUILT_IN_TIERING_PRODUCT_SERIES, TIERING_ELIGIBLE_SOURCES } from "../src/shared/tiering";
import type {
  Member,
  Requirement,
  RequirementSource,
  TieringData,
  TieringEvidenceSource,
  TieringSupportRecord,
  WorkspaceData,
} from "../src/shared/types";

const generatedAt = "2026-09-21T06:30:00.000Z";

function member(id: string, name: string, role: Member["role"]): Member {
  return {
    id,
    name,
    employeeNumber: id.toUpperCase(),
    microsoftAccount: `${id}@company.com`,
    role,
    domainL1Ids: ["l1"],
    sortOrder: 0,
    createdAt: generatedAt,
    updatedAt: generatedAt,
  };
}

function requirement(
  id: string,
  source: RequirementSource,
  targetMonth: string,
  overrides: Partial<Requirement> = {},
): Requirement {
  return {
    id,
    title: `需求-${id}`,
    description: `需求 ${id} 的规格说明`,
    images: [],
    domainL0Id: "l0",
    domainId: "l1",
    source,
    scope: "仅国内",
    allowLatestVersionUpgrade: true,
    overseasRegions: [],
    category: "平台演进",
    targetMonth,
    productIds: [],
    deviceWorkloadPm: 0,
    appWorkloadPm: 0,
    cloudWorkloadPm: 0,
    researchAppWorkloadPm: 0,
    algorithmWorkloadPm: 0,
    unallocatedWorkloadPm: 0,
    workloadPm: 0,
    createdAt: generatedAt,
    updatedAt: generatedAt,
    ...overrides,
  };
}

const evidence: TieringEvidenceSource = {
  id: "S001",
  sourceType: "官方网页",
  title: "官方说明",
  url: "https://example.com/evidence",
  scope: "全部",
  market: "中国大陆",
  versionCondition: "",
  collectedAt: "2026-09-20",
  verifiedAt: "2026-09-21",
  evidenceLevel: "A",
  maintenanceStatus: "有效",
  owner: "规划甲",
  note: "",
  updatedAt: generatedAt,
};

const support: TieringSupportRecord = {
  id: "F001::PS01",
  featureId: "F001",
  seriesId: "PS01",
  status: "支持",
  supportedModels: "WATCH 5",
  excludedModels: "WATCH 3",
  phoneOsMarket: "HarmonyOS／中国大陆",
  minHealthAppVersion: "15.1.0",
  minDeviceVersion: "5.0.0",
  prerequisites: "需要打开健康数据权限",
  sourceIds: ["S001"],
  sourceUrl: "https://example.com/direct",
  verifiedAt: "2026-09-21",
  evidenceLevel: "A",
  auditStatus: "已审核",
  note: "仅中国大陆版本",
  updatedAt: generatedAt,
};

function workspace(overrides: Partial<WorkspaceData> = {}): WorkspaceData {
  const owners = [
    member("planner", "规划甲", "规划"),
    member("pm", "产品乙", "产品经理"),
    member("se", "SE丙", "SE"),
    member("algo", "算法丁", "算法"),
  ];
  const r1 = requirement("r1", "健康基础", "2027-03", {
    title: "睡眠规律评估",
    tieringFeatureId: "F001",
    planningMemberId: "planner",
    productManagerMemberId: "pm",
    seMemberId: "se",
    algorithmMemberId: "algo",
  });
  const tiering: TieringData = {
    features: [{
      id: "F001",
      l0: "旧 L0",
      l1: "旧 L1",
      requirementTitle: "旧标题",
      requirementId: "r1",
      sortOrder: 1,
      createdAt: generatedAt,
      updatedAt: generatedAt,
    }],
    // 故意只提供一个本机系列：导出仍必须固定包含 15 个内置系列。
    productSeries: [BUILT_IN_TIERING_PRODUCT_SERIES[0]],
    supportRecords: [support],
    evidenceSources: [evidence],
    verificationIssues: [],
    changeLogs: [],
  };
  return {
    requirements: [
      r1,
      requirement("industry", "行业", "2027-04"),
      requirement("future", "AI运动", "2028-01"),
    ],
    domains: [
      { id: "l0", name: "健康管理", level: "L0", active: true, sortOrder: 0 },
      { id: "l1", name: "睡眠", level: "L1", parentId: "l0", active: true, sortOrder: 0 },
    ],
    products: [],
    groupOverrides: [],
    members: owners,
    pptAccessPolicy: { entries: [], updatedAt: generatedAt },
    tiering,
    ...overrides,
  };
}

async function load(bytes: Buffer): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes as never);
  return workbook;
}

function headerValues(sheet: ExcelJS.Worksheet): string[] {
  return Array.from({ length: sheet.columnCount }, (_, index) => sheet.getCell(4, index + 1).text);
}

describe("分层分级 Excel 导出", () => {
  it("输出两张表、15个固定系列、责任人与完整规格", async () => {
    const input = workspace();
    input.tiering!.supportRecords[0].note = "长规格备注用于验证自动换行且不会被固定行高截断。".repeat(18);
    const bytes = await createTieringWorkbook(
      input,
      "0.8.24",
      generatedAt,
      { start: "2027H1", end: "2027H1" },
    );
    const workbook = await load(bytes);

    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["支持矩阵", "规格明细"]);
    const matrix = workbook.getWorksheet("支持矩阵")!;
    const details = workbook.getWorksheet("规格明细")!;
    const matrixHeaders = headerValues(matrix);
    expect(matrixHeaders.slice(0, 12)).toEqual([
      "序号", "功能ID", "L0", "L1", "需求", "来源", "上线月份", "HMOS版本", "规划", "产品经理", "SE", "算法",
    ]);
    expect(matrixHeaders.slice(12)).toEqual(BUILT_IN_TIERING_PRODUCT_SERIES.map((item) => item.name));
    expect(matrix.columnCount).toBe(27);
    expect(matrix.getCell("A5").value).toBe(1);
    expect(matrix.getCell("B5").text).toBe("F001");
    expect(matrix.getCell("C5").text).toBe("健康管理");
    expect(matrix.getCell("D5").text).toBe("睡眠");
    expect(matrix.getCell("E5").text).toBe("睡眠规律评估");
    expect(matrix.getCell("H5").text).toBe("HMOS 7.1");
    expect([matrix.getCell("I5").text, matrix.getCell("J5").text, matrix.getCell("K5").text, matrix.getCell("L5").text]).toEqual([
      "规划甲", "产品乙", "SE丙", "算法丁",
    ]);
    expect(matrix.getCell("M5").text).toBe("支持");
    expect(matrix.getCell("N5").text).toBe("未判断");
    expect(matrix.getCell("N4").text).toBe("WATCH Pro系列");
    expect(matrix.getCell("W4").text).toBe("Band Pro系列");
    expect(matrix.getCell("W5").text).toBe("未判断");
    expect(matrix.getCell("A2").text).toContain("27H1–27H1");
    expect(matrix.getCell("A3").text).toContain("不支持回导");
    expect(matrix.views[0]).toMatchObject({ state: "frozen", xSplit: 5, ySplit: 4 });
    expect(matrix.autoFilter).toBe("A4:AA5");

    expect(headerValues(details)).toEqual([
      "序号", "功能ID", "L0", "L1", "需求", "来源", "上线月份", "HMOS版本", "系列", "支持状态", "支持型号", "排除型号", "系统/市场",
      "最低运动健康App版本", "最低设备版本", "前置条件", "证据编号", "证据链接", "核验日期", "证据等级", "审核状态", "备注", "更新时间",
    ]);
    expect(details.actualRowCount).toBe(19);
    expect(details.getCell("A5").value).toBe(1);
    expect(details.getCell("A19").value).toBe(15);
    expect(details.getCell("I5").text).toBe("WATCH数字系列");
    expect(details.getCell("J5").text).toBe("支持");
    expect(details.getCell("K5").text).toBe("WATCH 5");
    expect(details.getCell("Q5").text).toBe("S001");
    expect(details.getCell("R5").text).toContain("https://example.com/direct");
    expect(details.getCell("R5").text).toContain("https://example.com/evidence");
    expect(details.getCell("J6").text).toBe("未判断");
    expect(details.getRow(5).height).toBeGreaterThan(54);
    expect(details.views[0]).toMatchObject({ state: "frozen", xSplit: 5, ySplit: 4 });
    expect(details.autoFilter).toBe("A4:W19");
  });

  it("仅选取时间范围内的运动、健康和 AI 九个来源", () => {
    const base = workspace();
    base.requirements = [
      ...TIERING_ELIGIBLE_SOURCES.map((source, index) => requirement(`eligible-${index}`, source, "2027-03")),
      requirement("industry", "行业", "2027-03"),
      requirement("scene", "全场景", "2027-03"),
      requirement("platform", "health平台", "2027-03"),
      requirement("nmpa", "NMPA认证", "2027-03"),
      requirement("ce", "CE认证", "2027-03"),
      requirement("outside", "健康基础", "2027-09"),
    ];
    base.tiering = { features: [], productSeries: [], supportRecords: [], evidenceSources: [], verificationIssues: [], changeLogs: [] };

    const demands = tieringDemandsForRange(base, { start: "2027H1", end: "2027H1" });
    expect(demands.map((item) => item.requirement.source)).toEqual([...TIERING_ELIGIBLE_SOURCES]);
  });

  it("空数据区间仍保留两张表、完整表头和可用筛选", async () => {
    const bytes = await createTieringWorkbook(
      workspace(),
      "0.8.24",
      generatedAt,
      { start: "2030H1", end: "2030H2" },
    );
    const workbook = await load(bytes);
    const matrix = workbook.getWorksheet("支持矩阵")!;
    const details = workbook.getWorksheet("规格明细")!;

    expect(matrix.getCell("A5").text).toContain("暂无");
    expect(matrix.getCell("A5").isMerged).toBe(true);
    expect(matrix.columnCount).toBe(27);
    expect(matrix.autoFilter).toBe("A4:AA4");
    expect(details.getCell("A5").text).toContain("暂无");
    expect(details.getCell("A5").isMerged).toBe(true);
    expect(details.columnCount).toBe(23);
    expect(details.autoFilter).toBe("A4:W4");
  });
});
