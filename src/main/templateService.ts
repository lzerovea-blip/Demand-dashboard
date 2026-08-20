import { createHash } from "node:crypto";
import { basename, join } from "node:path";
import { copyFile, mkdir, readFile } from "node:fs/promises";
import { app } from "electron";
import type { TemplateProfile } from "../shared/types.js";
import { inspectTemplate } from "./pptxInspector.js";

export async function importTemplateFile(path: string): Promise<TemplateProfile> {
  const data = await readFile(path);
  const sha256 = createHash("sha256").update(data).digest("hex");
  const inventory = await inspectTemplate(data);
  const templateDir = join(app.getPath("userData"), "templates");
  await mkdir(templateDir, { recursive: true });
  const storedPath = join(templateDir, `${sha256}.pptx`);
  await copyFile(path, storedPath);
  return {
    id: sha256.slice(0, 20),
    name: basename(path),
    storedPath,
    originalPath: path,
    sha256,
    slideCount: inventory.slideCount,
    importedAt: new Date().toISOString(),
    status: "needs_mapping",
    inventory,
  };
}
