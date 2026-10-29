import JSZip from "jszip";
import { XMLParser, XMLValidator } from "fast-xml-parser";

const REQUIRED_PARTS = ["[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"] as const;

export interface PptxValidationResult {
  entries: number;
  slides: number;
}

export async function validateGeneratedPptx(bytes: Buffer, expectedSlides: number): Promise<PptxValidationResult> {
  let archive: JSZip;
  try {
    archive = await JSZip.loadAsync(bytes, { checkCRC32: true });
  } catch (error) {
    throw new Error(`原始 PPTX 压缩包损坏：${error instanceof Error ? error.message : String(error)}`);
  }
  for (const part of REQUIRED_PARTS) {
    if (!archive.file(part)) throw new Error(`原始 PPTX 缺少必要文件：${part}`);
  }
  const xmlEntries = Object.values(archive.files).filter((entry) => !entry.dir && /(?:\.xml|\.rels)$/iu.test(entry.name));
  for (const entry of xmlEntries) {
    const xml = await entry.async("string");
    if (XMLValidator.validate(xml) !== true) throw new Error(`原始 PPTX XML 无效：${entry.name}`);
    if (hasNonFiniteDrawingValue(xml)) throw new Error(`原始 PPTX 包含 PowerPoint 无法识别的非有限几何值：${entry.name}`);
  }
  const contentTypes = new XMLParser({ ignoreAttributes: false }).parse(await archive.file("[Content_Types].xml")!.async("string"));
  const overrides = contentTypes?.Types?.Override ?? [];
  for (const override of Array.isArray(overrides) ? overrides : [overrides]) {
    const part = String(override["@_PartName"] ?? "").replace(/^\//, "");
    if (!part || !archive.file(part)) throw new Error(`原始 PPTX 内容类型声明指向不存在的文件：${part}`);
  }
  const presentationXml = await archive.file("ppt/presentation.xml")!.async("string");
  const presentation = new XMLParser({ ignoreAttributes: false }).parse(presentationXml);
  const slideIds = presentation?.["p:presentation"]?.["p:sldIdLst"]?.["p:sldId"];
  const slides = Array.isArray(slideIds) ? slideIds.length : slideIds ? 1 : 0;
  if (slides !== expectedSlides) throw new Error(`原始 PPTX 页数不一致：预计 ${expectedSlides} 页，实际 ${slides} 页`);
  for (let index = 1; index <= slides; index += 1) {
    if (!archive.file(`ppt/slides/slide${index}.xml`)) throw new Error(`原始 PPTX 缺少第 ${index} 页内容`);
  }
  return { entries: Object.keys(archive.files).length, slides };
}

export function hasNonFiniteDrawingValue(xml: string): boolean {
  return /\bfmla="[^"]*(?:NaN|[+-]?Infinity)[^"]*"/u.test(xml)
    || /\b(?:x|y|cx|cy|w|h|pos|ang)="(?:NaN|[+-]?Infinity)"/u.test(xml);
}
