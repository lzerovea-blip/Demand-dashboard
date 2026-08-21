import { createRequire } from "node:module";
import { cp, mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { spawn } from "node:child_process";

const require = createRequire(import.meta.url);
const electronExecutable = require("electron");
const electronApp = resolve(dirname(electronExecutable), "../..");
const releaseDir = resolve(process.env.ROADMAP_RELEASE_DIR ?? "release");
const productName = "解决方案需求管理";
const sourcePackage = JSON.parse(await readFile("package.json", "utf8"));
const targetApp = join(releaseDir, `${productName}.app`);
const zipPath = join(releaseDir, `${productName}-${sourcePackage.version}-mac-arm64.zip`);
const dmgPath = join(releaseDir, `${productName}-${sourcePackage.version}-mac-arm64.dmg`);
const resourcesDir = join(targetApp, "Contents", "Resources");
const appDir = join(resourcesDir, "app");
const templatesDir = join(resourcesDir, "templates");

await mkdir(releaseDir, { recursive: true });
for (const path of [targetApp, zipPath, dmgPath]) await rm(path, { recursive: true, force: true });
await run("/usr/bin/ditto", [electronApp, targetApp]);
await rm(join(resourcesDir, "default_app.asar"), { force: true });
await mkdir(appDir, { recursive: true });

await cp("dist-electron", join(appDir, "dist-electron"), { recursive: true });
await cp("dist-renderer", join(appDir, "dist-renderer"), { recursive: true });
await cp("node_modules/sql.js/dist/sql-wasm.wasm", join(resourcesDir, "sql-wasm.wasm"));
await mkdir(templatesDir, { recursive: true });
await cp("需求路标模板-共创版-v0.1.pptx", join(templatesDir, "需求路标模板-共创版-v0.1.pptx"));

const packagedManifest = {
  name: sourcePackage.name,
  productName,
  version: sourcePackage.version,
  private: true,
  type: "module",
  main: sourcePackage.main,
};
await writeFile(join(appDir, "package.json"), `${JSON.stringify(packagedManifest, null, 2)}\n`, "utf8");

const plist = join(targetApp, "Contents", "Info.plist");
for (const [key, value] of [
  ["CFBundleDisplayName", productName],
  ["CFBundleName", productName],
  ["CFBundleIdentifier", "com.local.healthroadmap"],
  ["CFBundleShortVersionString", sourcePackage.version],
  ["CFBundleVersion", sourcePackage.version],
]) {
  await run("/usr/libexec/PlistBuddy", ["-c", `Set :${key} ${value}`, plist]);
}

await run("/usr/bin/codesign", ["--force", "--deep", "--sign", "-", targetApp]);
await run("/usr/bin/ditto", ["-c", "-k", "--sequesterRsrc", "--keepParent", targetApp, zipPath]);
await run("/usr/bin/hdiutil", ["create", "-volname", productName, "-srcfolder", targetApp, "-ov", "-format", "UDZO", dmgPath]);

console.log(`Packaged ${targetApp}`);
console.log(`Created ${zipPath}`);
console.log(`Created ${dmgPath}`);

function run(command, args) {
  return new Promise((resolvePromise, rejectPromise) => {
    const child = spawn(command, args, { stdio: "inherit" });
    child.on("error", rejectPromise);
    child.on("exit", (code) => code === 0 ? resolvePromise() : rejectPromise(new Error(`${command} exited with ${code}`)));
  });
}
