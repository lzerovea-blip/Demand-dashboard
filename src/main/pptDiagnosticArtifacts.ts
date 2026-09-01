import { createHash } from "node:crypto";
import { copyFile, mkdir, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { join } from "node:path";

export interface PptDiagnosticArtifactSource {
  sourcePath?: string;
  fileName: string;
  label: string;
}

export interface PreservePptDiagnosticArtifactsInput {
  userDataPath: string;
  createdAt: string;
  appVersion: string;
  stage: string;
  logContent: string;
  files: PptDiagnosticArtifactSource[];
}

export async function preservePptDiagnosticArtifacts(
  input: PreservePptDiagnosticArtifactsInput,
): Promise<string> {
  const root = join(input.userDataPath, "logs", "ppt-export-diagnostics");
  const directory = join(root, diagnosticDirectoryName(input.createdAt));
  await mkdir(directory, { recursive: true });

  const manifest = [
    "解决方案需求管理 - PPT 诊断文件清单",
    `时间：${input.createdAt}`,
    `应用版本：${input.appVersion}`,
    `失败阶段：${input.stage}`,
    "隐私提示：本目录中的 PPT 包含业务数据，仅保存在本机，请勿直接外发。",
  ];
  for (const file of input.files) {
    if (!file.sourcePath) continue;
    try {
      const fileStat = await stat(file.sourcePath);
      if (!fileStat.isFile() || fileStat.size <= 0) continue;
      const target = join(directory, file.fileName);
      await copyFile(file.sourcePath, target);
      const bytes = await readFile(target);
      manifest.push(`${file.label}：${file.fileName}；${bytes.length} 字节；SHA-256 ${createHash("sha256").update(bytes).digest("hex")}`);
    } catch {
      // Missing intermediate files are expected when a failure occurs early.
    }
  }
  await Promise.all([
    writeFile(join(directory, "diagnostic.txt"), input.logContent, "utf8"),
    writeFile(join(directory, "manifest.txt"), `${manifest.join("\n")}\n`, "utf8"),
  ]);
  await prunePptDiagnosticDirectories(root, 3);
  return directory;
}

export async function prunePptDiagnosticDirectories(root: string, keep = 3): Promise<void> {
  let entries;
  try {
    entries = await readdir(root, { withFileTypes: true });
  } catch {
    return;
  }
  const directories = entries.filter((entry) => entry.isDirectory()).map((entry) => entry.name).sort().reverse();
  await Promise.all(directories.slice(Math.max(0, keep)).map((name) => rm(join(root, name), { recursive: true, force: true })));
}

function diagnosticDirectoryName(createdAt: string): string {
  return `failure-${createdAt.replace(/[:.]/g, "-")}`;
}
