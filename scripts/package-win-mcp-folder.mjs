import { createHash } from "node:crypto";
import { spawn } from "node:child_process";
import { createReadStream, createWriteStream } from "node:fs";
import { readFile, readdir, rename, rm, stat, writeFile } from "node:fs/promises";
import { join, relative, resolve } from "node:path";
import { pipeline } from "node:stream/promises";
import JSZip from "jszip";

const releaseDir = resolve(process.env.ROADMAP_RELEASE_DIR ?? "release");
const packageInfo = JSON.parse(await readFile("package.json", "utf8"));
const appDir = join(releaseDir, "win-unpacked");
const appExe = join(appDir, "Solution-Requirements-Manager.exe");
const builtExe = join(appDir, "解决方案需求管理.exe");
const appAsar = join(appDir, "resources", "app.asar");
const archiveName = `Solution-Requirements-Manager-MCP-Folder-${packageInfo.version}-x64.zip`;
const archivePath = join(releaseDir, archiveName);

if ((await stat(builtExe).catch(() => null))?.isFile()) await rename(builtExe, appExe);
for (const path of [appExe, appAsar, ...["node.exe", "server.cjs", "runtime.json"].map(name => join(appDir, "resources", "mcp", name))]) {
  if (!(await stat(path).catch(() => null))?.isFile()) {
    throw new Error(`Windows unpacked app is incomplete: ${path}`);
  }
}

await rm(archivePath, { force: true });
const zip = new JSZip();
for (const filePath of await filesIn(appDir)) {
  const archiveEntry = relative(releaseDir, filePath).split("\\").join("/");
  zip.file(archiveEntry, createReadStream(filePath), { binary: true, createFolders: true });
}
await pipeline(zip.generateNodeStream({ streamFiles: true, compression: "DEFLATE", compressionOptions: { level: 6 } }), createWriteStream(archivePath));
await run("unzip", ["-tqq", archiveName], releaseDir);

const hash = createHash("sha256");
for await (const chunk of createReadStream(archivePath)) hash.update(chunk);
const digest = hash.digest("hex");
await writeFile(join(releaseDir, "SHA256SUMS-mcp-folder.txt"), `${digest}  ${archiveName}\n`, "utf8");
console.log(`MCP folder package: ${archivePath}`);
console.log(`SHA-256: ${digest}`);

async function filesIn(directory) {
  const paths = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const entryPath = join(directory, entry.name);
    if (entry.isDirectory()) paths.push(...await filesIn(entryPath));
    else if (entry.isFile()) paths.push(entryPath);
    else throw new Error(`Unsupported package entry: ${entryPath}`);
  }
  return paths;
}

function run(command, args, cwd) {
  return new Promise((resolvePromise, rejectPromise) => {
    const child = spawn(command, args, { cwd, stdio: "inherit" });
    child.once("error", rejectPromise);
    child.once("exit", (code) => code === 0 ? resolvePromise() : rejectPromise(new Error(`${command} exited with ${code}`)));
  });
}
