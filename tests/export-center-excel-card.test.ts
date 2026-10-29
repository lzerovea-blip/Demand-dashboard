import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { TemplatePage } from "../src/renderer/App";
import type { AppSnapshot } from "../src/shared/types";

const snapshot: AppSnapshot = {
  requirements: [],
  domains: [],
  products: [],
  members: [],
  groupOverrides: [],
  templates: [],
  pptAccessPolicy: { entries: [], updatedAt: "2026-09-22T00:00:00Z" },
};

describe("导出中心 Excel 卡片", () => {
  it("在同一卡片提供两种导出及原协作导入入口，不再单列分层分级模块", () => {
    const html = renderToStaticMarkup(createElement(TemplatePage, {
      snapshot,
      start: "2026H1",
      end: "2027H2",
      onExportPpt: async () => {},
      onExportWorkspace: async () => {},
      onInspectWorkspace: async () => undefined,
      onApplyWorkspace: async () => false,
      onExportWorkbook: async () => {},
      onExportTieringWorkbook: async () => {},
      onInspectWorkbook: async () => undefined,
      onApplyWorkbook: async () => false,
      onSavePptAccessPolicy: async () => {},
      onGetPptDiagnostic: async () => null,
      onOpenPptDiagnostic: async () => ({ path: "/diagnostics" }),
      onCheckIrm: async () => ({ irmReady: false, message: "" } as Awaited<ReturnType<Parameters<typeof TemplatePage>[0]["onCheckIrm"]>>),
    }));
    const excelCard = html.match(/<section class="panel export-type-card excel-export-card">([\s\S]*?)<\/section>/)?.[1];

    expect(excelCard).toBeDefined();
    expect(excelCard).toContain(">导出需求池</button>");
    expect(excelCard).toContain(">导出分层分级</button>");
    expect(excelCard).toContain(">导入协作 Excel</button>");
    expect(excelCard).toContain("分层分级 Excel 不支持通过应用回导");
    expect(html).not.toContain("tiering-workbook-panel");
  });
});
