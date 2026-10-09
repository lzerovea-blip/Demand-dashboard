import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { resolve } from "node:path";
import assert from "node:assert/strict";

const version = JSON.parse(readFileSync("package.json", "utf8")).version;
const output = process.argv[2] ?? `release-v${version}-win`;
const digest = (bytes) => createHash("sha256").update(bytes).digest("hex");
function inspectPe(path) {
  const data = readFileSync(path); const pe = data.readUInt32LE(60);
  assert.equal(data.subarray(pe, pe + 4).toString("hex"), "50450000");
  const optional = pe + 24; const directories = optional + (data.readUInt16LE(optional) === 0x20b ? 112 : 96);
  const sections = optional + data.readUInt16LE(pe + 20);
  function offset(rva) {
    for (let i = 0; i < data.readUInt16LE(pe + 6); i++) { const at = sections + i * 40; const start = data.readUInt32LE(at + 12); const size = Math.max(data.readUInt32LE(at + 8), data.readUInt32LE(at + 16)); if (rva >= start && rva < start + size) return data.readUInt32LE(at + 20) + rva - start; }
    throw Error("Unmapped PE resource address");
  }
  const resources = offset(data.readUInt32LE(directories + 16)); const versions = [];
  function walk(relative, path = []) {
    const at = resources + relative; const count = data.readUInt16LE(at + 12) + data.readUInt16LE(at + 14);
    for (let i = 0; i < count; i++) { const entry = at + 16 + i * 8; const id = data.readUInt32LE(entry); const value = data.readUInt32LE(entry + 4);
      if (value & 0x80000000) walk(value & 0x7fffffff, [...path, id]);
      else if (path[0] === 16) { const leaf = resources + value; const start = offset(data.readUInt32LE(leaf)); const bytes = data.subarray(start, start + data.readUInt32LE(leaf + 4)); const signature = Buffer.from([0xbd,0x04,0xef,0xfe]); const fixed = bytes.indexOf(signature); assert(fixed >= 0); const ms = bytes.readUInt32LE(fixed + 8); const ls = bytes.readUInt32LE(fixed + 12); versions.push(`${ms >>> 16}.${ms & 65535}.${ls >>> 16}.${ls & 65535}`); }
    }
  }
  walk(0); assert(versions.includes(`${version}.0`), `${path}: built-in version ${versions}`);
  return { path: resolve(path), bytes: data.length, machine: data.readUInt16LE(pe + 4).toString(16), versions, sha256: digest(data) };
}
const portable = inspectPe(`${output}/Solution-Requirements-Manager-Portable-${version}-x64.exe`);
const executable = inspectPe(`${output}/win-unpacked/解决方案需求管理.exe`);
assert.equal(executable.machine, "8664", "Electron payload must be Windows x64");
const asar = readFileSync(`${output}/win-unpacked/resources/app.asar`);
const headerSize = asar.readUInt32LE(4); const jsonLength = asar.readUInt32LE(12); const header = JSON.parse(asar.subarray(16, 16 + jsonLength).toString());
function entry(path) { let node = header; for (const part of path.split("/")) node = node.files[part]; const start = 8 + headerSize + Number(node.offset); return asar.subarray(start, start + node.size); }
const manifest = JSON.parse(entry("package.json").toString()); assert.equal(manifest.version, version);
const mainPath = "dist-electron/main/index.cjs";
assert.equal(digest(entry(mainPath)), digest(readFileSync(mainPath)), "Packaged main process differs from verified production build");
assert.equal(digest(entry("dist-renderer/index.html")), digest(readFileSync("dist-renderer/index.html")), "Packaged renderer differs from production build");
console.log(JSON.stringify({ version, portable, executable, asarManifestVersion: manifest.version, asarSha256: digest(asar) }, null, 2));
