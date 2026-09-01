import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { preservePptDiagnosticArtifacts } from "../src/main/pptDiagnosticArtifacts";

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

describe("PPT 失败诊断文件保留", () => {
  it("保存现有阶段文件、清单与简洁日志，并只保留最近三次", async () => {
    const root = await mkdtemp(join(tmpdir(), "ppt-diagnostic-artifacts-"));
    temporaryDirectories.push(root);
    const source = join(root, "source.pptx");
    await writeFile(source, Buffer.from("ppt-stage-content"));

    for (let index = 0; index < 4; index++) {
      await preservePptDiagnosticArtifacts({
        userDataPath: root,
        createdAt: `2026-09-01T08:00:0${index}.000Z`,
        appVersion: "v0.7.40",
        stage: `阶段 ${index}`,
        logContent: `日志 ${index}`,
        files: [
          { sourcePath: source, fileName: "01-generated-original.pptx", label: "原始生成 PPT" },
          { sourcePath: join(root, "missing.pptx"), fileName: "02-missing.pptx", label: "尚未生成" },
        ],
      });
    }

    const diagnosticRoot = join(root, "logs", "ppt-export-diagnostics");
    const directories = (await readdir(diagnosticRoot)).sort();
    expect(directories).toHaveLength(3);
    expect(directories[0]).toContain("08-00-01");
    expect(directories[2]).toContain("08-00-03");
    const latest = join(diagnosticRoot, directories[2]);
    expect(await readFile(join(latest, "diagnostic.txt"), "utf8")).toBe("日志 3");
    expect(await readFile(join(latest, "01-generated-original.pptx"), "utf8")).toBe("ppt-stage-content");
    const manifest = await readFile(join(latest, "manifest.txt"), "utf8");
    expect(manifest).toContain("包含业务数据，仅保存在本机");
    expect(manifest).toMatch(/SHA-256 [a-f0-9]{64}/u);
    expect(manifest).not.toContain("missing.pptx");
  });
});
