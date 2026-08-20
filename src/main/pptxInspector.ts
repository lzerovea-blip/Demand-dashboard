import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import type { TemplateInventory, TemplateShapeInventory } from "../shared/types.js";

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });

export async function inspectTemplate(data: Uint8Array): Promise<TemplateInventory> {
  const zip = await JSZip.loadAsync(data);
  const presentationXml = await readZipText(zip, "ppt/presentation.xml");
  const presentation = parser.parse(presentationXml);
  const slideIds = asArray(presentation?.["p:presentation"]?.["p:sldIdLst"]?.["p:sldId"]);
  const size = presentation?.["p:presentation"]?.["p:sldSz"];
  const shapes: TemplateShapeInventory[] = [];

  for (let slideNumber = 1; slideNumber <= slideIds.length; slideNumber += 1) {
    const slidePath = `ppt/slides/slide${slideNumber}.xml`;
    if (!zip.file(slidePath)) continue;
    const slide = parser.parse(await readZipText(zip, slidePath));
    const tree = slide?.["p:sld"]?.["p:cSld"]?.["p:spTree"];
    for (const shape of asArray(tree?.["p:sp"])) {
      const properties = shape?.["p:nvSpPr"]?.["p:cNvPr"] ?? {};
      const placeholder = shape?.["p:nvSpPr"]?.["p:nvPr"]?.["p:ph"];
      shapes.push({
        slideNumber,
        shapeId: String(properties?.["@_id"] ?? ""),
        shapeName: String(properties?.["@_name"] ?? ""),
        text: collectText(shape?.["p:txBody"]).trim(),
        placeholderType: placeholder?.["@_type"] ? String(placeholder["@_type"]) : undefined,
      });
    }
    for (const frame of asArray(tree?.["p:graphicFrame"])) {
      const properties = frame?.["p:nvGraphicFramePr"]?.["p:cNvPr"] ?? {};
      shapes.push({
        slideNumber,
        shapeId: String(properties?.["@_id"] ?? ""),
        shapeName: String(properties?.["@_name"] ?? ""),
        text: collectText(frame).trim(),
      });
    }
  }

  if (slideIds.length === 0) throw new Error("PPTX中未发现幻灯片");
  return {
    slideCount: slideIds.length,
    slideSize: size ? { width: Number(size["@_cx"] ?? 0), height: Number(size["@_cy"] ?? 0) } : undefined,
    shapes,
  };
}

async function readZipText(zip: JSZip, path: string): Promise<string> {
  const entry = zip.file(path);
  if (!entry) throw new Error(`PPTX缺少必要文件：${path}`);
  return entry.async("string");
}

function collectText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number") return String(value);
  if (Array.isArray(value)) return value.map(collectText).filter(Boolean).join(" ");
  if (typeof value === "object") {
    const record = value as Record<string, unknown>;
    if (record["a:t"] !== undefined) return collectText(record["a:t"]);
    return Object.entries(record)
      .filter(([key]) => !key.startsWith("@_"))
      .map(([, child]) => collectText(child))
      .filter(Boolean)
      .join(" ");
  }
  return "";
}

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}
