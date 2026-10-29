import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { DictionariesPage, MembersPanel } from "../src/renderer/App";
import { RoadmapDictionaryPanel } from "../src/renderer/RoadmapManagementPage";
import type { AppSnapshot, DictionaryItem } from "../src/shared/types";

const now = "2026-10-09T00:00:00.000Z";
const domains: DictionaryItem[] = [
  { id: "l0", name: "健康", level: "L0", sortOrder: 0, active: true },
  { id: "l1", name: "睡眠", level: "L1", parentId: "l0", sortOrder: 0, active: true },
];
const snapshot: AppSnapshot = {
  requirements: [], domains, products: [], members: [], groupOverrides: [], templates: [],
  pptAccessPolicy: { entries: [], updatedAt: now },
  roadmaps: [{ id: "health", name: "健康路标", kind: "standard", sortOrder: 0, updatedAt: now }],
  roadmapLanes: [{ id: "basic", roadmapId: "health", name: "基础", sortOrder: 0, updatedAt: now }],
};
const requestConfirmation = () => {};

describe("独立字典设置页面", () => {
  it("需求字典只包含 L0、L1 和产品，人员单独成页", () => {
    const requirementsHtml = renderToStaticMarkup(createElement(DictionariesPage, {
      domains, products: [], onSaveDomain: async () => {}, onSaveProduct: async () => {},
      onDeleteDomain: async () => {}, onDeleteProduct: async () => {}, requestConfirmation,
    }));
    expect(requirementsHtml).toContain("领域 L0 字典");
    expect(requirementsHtml).toContain("领域 L1 字典");
    expect(requirementsHtml).toContain("产品字典");
    expect(requirementsHtml).not.toContain("人员管理");

    const membersHtml = renderToStaticMarkup(createElement(MembersPanel, {
      domains, members: [], onSave: async () => {}, onDelete: async () => {}, requestConfirmation,
    }));
    expect(membersHtml).toContain("人员管理");
    expect(membersHtml).not.toContain("产品字典");
  });

  it("路标页顶部只有新增路标，未选中时不出现全局泳道表单", () => {
    const html = renderToStaticMarkup(createElement(RoadmapDictionaryPanel, {
      snapshot, onSaveRoadmap: async () => snapshot, onDeleteRoadmap: async () => {},
      onSaveLane: async () => {}, onDeleteLane: async () => {}, confirm: () => {},
    }));
    expect(html).toContain("路标字典设置");
    expect(html).toContain('aria-label="新增路标名称"');
    expect(html).toContain("添加路标");
    expect(html).toContain("健康路标");
    expect(html).toContain("1 条泳道");
    expect(html).toContain('aria-expanded="false"');
    expect(html).not.toContain("选择所属路标");
    expect(html).not.toContain("添加泳道");
  });
});
