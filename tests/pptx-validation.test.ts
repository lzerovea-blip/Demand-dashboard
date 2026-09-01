import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { createRoadmapPresentation, presentationSlideCount } from "../src/main/pptxExport";
import { validateGeneratedPptx } from "../src/main/pptxValidation";
import { buildPptExportPlan } from "../src/shared/pptExport";
import type { AppSnapshot, PptExportRequest } from "../src/shared/types";

const snapshot: AppSnapshot = {
  requirements: [],
  domains: [],
  products: [],
  groupOverrides: [],
  members: [],
  pptAccessPolicy: { entries: [], updatedAt: "1970-01-01T00:00:00.000Z" },
  templates: [],
};
const request: PptExportRequest = {
  start: "2026H1",
  end: "2027H2",
  materialName: "校验样稿",
  sections: ["summary", "运动", "健康"],
};

describe("生成 PPTX 结构校验", () => {
  it("校验压缩包、XML、必要文件和实际页数", async () => {
    const bytes = await createRoadmapPresentation(snapshot, request);
    const expectedSlides = presentationSlideCount(buildPptExportPlan(snapshot, request));
    await expect(validateGeneratedPptx(bytes, expectedSlides)).resolves.toMatchObject({ slides: expectedSlides });
  });

  it("在进入 PowerPoint 前拒绝损坏文件、缺失核心文件和页数不一致", async () => {
    await expect(validateGeneratedPptx(Buffer.from("not-a-pptx"), 1)).rejects.toThrow("压缩包损坏");
    const bytes = await createRoadmapPresentation(snapshot, request);
    const archive = await JSZip.loadAsync(bytes);
    archive.remove("ppt/presentation.xml");
    await expect(validateGeneratedPptx(await archive.generateAsync({ type: "nodebuffer" }), 1)).rejects.toThrow("缺少必要文件");
    await expect(validateGeneratedPptx(bytes, 99)).rejects.toThrow("页数不一致");
  });
});
