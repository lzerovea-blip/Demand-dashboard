import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
const digest = bytes => createHash("sha256").update(bytes).digest("hex");
const ico = readFileSync("build-resources/app-icon.ico");
const expected = new Set();
for (let index = 0; index < ico.readUInt16LE(4); index++) {
  const base = 6 + index * 16;
  const start = ico.readUInt32LE(base + 12);
  expected.add(digest(ico.subarray(start, start + ico.readUInt32LE(base + 8))));
}
for (const file of process.argv.slice(2)) {
  const data = readFileSync(file);
  const pe = data.readUInt32LE(60);
  if (data.subarray(pe, pe + 4).toString("hex") !== "50450000") throw Error("Invalid PE");
  const optional = pe + 24;
  const is64 = data.readUInt16LE(optional) === 0x20b;
  const directories = optional + (is64 ? 112 : 96);
  const sections = optional + data.readUInt16LE(pe + 20);
  function rvaOffset(rva) {
    for (let i = 0; i < data.readUInt16LE(pe + 6); i++) {
      const at = sections + i * 40;
      const virtual = data.readUInt32LE(at + 12);
      const size = Math.max(data.readUInt32LE(at + 8), data.readUInt32LE(at + 16));
      if (rva >= virtual && rva < virtual + size) return data.readUInt32LE(at + 20) + rva - virtual;
    }
    throw Error(`Unmapped RVA ${rva}`);
  }
  const resources = rvaOffset(data.readUInt32LE(directories + 16));
  const found = new Set();
  let groups = 0;
  function walk(relative, path = []) {
    const at = resources + relative;
    const count = data.readUInt16LE(at + 12) + data.readUInt16LE(at + 14);
    for (let i = 0; i < count; i++) {
      const entry = at + 16 + i * 8;
      const id = data.readUInt32LE(entry);
      const value = data.readUInt32LE(entry + 4);
      const next = [...path, id];
      if (value & 0x80000000) walk(value & 0x7fffffff, next);
      else {
        const leaf = resources + value;
        const start = rvaOffset(data.readUInt32LE(leaf));
        const bytes = data.subarray(start, start + data.readUInt32LE(leaf + 4));
        if (next[0] === 3) found.add(digest(bytes));
        if (next[0] === 14) groups++;
      }
    }
  }
  walk(0);
  const matching = [...expected].filter(hash => found.has(hash)).length;
  if (matching !== expected.size || !groups) throw Error(`${file}: icon resources mismatch (${matching}/${expected.size})`);
  console.log(JSON.stringify({ file, machine: data.readUInt16LE(pe + 4).toString(16), iconFrames: matching, groups, sha256: digest(data) }));
}
