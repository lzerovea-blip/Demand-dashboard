import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { inspectTemplate } from "../src/main/pptxInspector";

describe("PPTX模板检查", () => {
  it("读取页数、画布尺寸、形状名称和占位符", async () => {
    const zip = new JSZip();
    zip.file("ppt/presentation.xml", `
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:sldIdLst><p:sldId id="256"/><p:sldId id="257"/></p:sldIdLst>
        <p:sldSz cx="12192000" cy="6858000"/>
      </p:presentation>`);
    zip.file("ppt/slides/slide1.xml", slideXml("12", "Summary Title", "汇总分析", "title"));
    zip.file("ppt/slides/slide2.xml", slideXml("22", "Roadmap Canvas", "运动路标"));
    const result = await inspectTemplate(await zip.generateAsync({ type: "uint8array" }));
    expect(result.slideCount).toBe(2);
    expect(result.slideSize).toEqual({ width: 12192000, height: 6858000 });
    expect(result.shapes[0]).toMatchObject({ slideNumber: 1, shapeId: "12", shapeName: "Summary Title", text: "汇总分析", placeholderType: "title" });
    expect(result.shapes[1]).toMatchObject({ slideNumber: 2, shapeId: "22", shapeName: "Roadmap Canvas", text: "运动路标" });
  });
});

function slideXml(id: string, name: string, text: string, placeholder?: string): string {
  return `
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree><p:sp>
        <p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:nvPr>${placeholder ? `<p:ph type="${placeholder}"/>` : ""}</p:nvPr></p:nvSpPr>
        <p:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody>
      </p:sp></p:spTree></p:cSld>
    </p:sld>`;
}
