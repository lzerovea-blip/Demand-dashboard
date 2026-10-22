import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { createRoadmapPresentation } from "../src/main/pptxExport";
import { hasNonFiniteDrawingValue, validateGeneratedPptx } from "../src/main/pptxValidation";
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
    const expectedSlides = buildPptExportPlan(snapshot, request).slideCount;
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

  it("导出清理生成库的悬空母版声明，检查时拒绝任何缺失的声明目标", async () => {
    const bytes = await createRoadmapPresentation(snapshot, request);
    const archive = await JSZip.loadAsync(bytes);
    const types = await archive.file("[Content_Types].xml")!.async("string");
    const masters = [...types.matchAll(/PartName="(\/ppt\/slideMasters\/[^\"]+)"/g)];
    expect(masters).toHaveLength(1);
    for (const match of masters) expect(archive.file(match[1].slice(1))).toBeTruthy();
    const theme = await archive.file("ppt/theme/theme1.xml")!.async("string");
    archive.file("[Content_Types].xml", types.replace("</Types>", '<Override PartName="/ppt/slideMasters/missing.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>'));
    await expect(validateGeneratedPptx(await archive.generateAsync({ type: "nodebuffer" }), 6)).rejects.toThrow("内容类型声明指向不存在");
    expect(await archive.file("ppt/theme/theme1.xml")!.async("string")).toBe(theme);
  });

  it("拒绝 PowerPoint 会判定损坏的 Infinity/NaN 几何值", async () => {
    expect(hasNonFiniteDrawingValue('<a:gd name="adj" fmla="val Infinity"/>')).toBe(true);
    expect(hasNonFiniteDrawingValue('<a:ext cx="NaN" cy="10"/>')).toBe(true);
    expect(hasNonFiniteDrawingValue("<a:t>Infinity 产品</a:t>")).toBe(false);
    const bytes = await createRoadmapPresentation(snapshot, request);
    const archive = await JSZip.loadAsync(bytes);
    const slide = await archive.file("ppt/slides/slide1.xml")!.async("string");
    archive.file(
      "ppt/slides/slide1.xml",
      slide.replace("</p:spTree>", '<p:sp><p:spPr><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val Infinity"/></a:avLst></a:prstGeom></p:spPr></p:sp></p:spTree>'),
    );
    const corrupted = await archive.generateAsync({ type: "nodebuffer" });
    await expect(validateGeneratedPptx(corrupted, 6)).rejects.toThrow("非有限几何值");
  });
});
