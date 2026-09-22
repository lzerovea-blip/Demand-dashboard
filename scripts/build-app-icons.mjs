// macOS maintainer utility: lossless format conversion / multi-resolution packaging.
// Artwork is supplied separately; this script does not redraw the selected icon.
import { mkdtemp, mkdir, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { execFileSync } from "node:child_process";

const directory = resolve("build-resources");
const source = join(directory, "app-icon.png");
const temporary = await mkdtemp(join(tmpdir(), "solution-app-icons-"));
const iconset = join(temporary, "app.iconset");
await mkdir(iconset);
const sizes = [16, 24, 32, 48, 64, 128, 256];
const frames = [];
for (const size of sizes) {
  const path = join(temporary, `${size}.png`);
  execFileSync("sips", ["-z", String(size), String(size), source, "--out", path]);
  frames.push(await readFile(path));
}
const header = Buffer.alloc(6 + sizes.length * 16);
header.writeUInt16LE(1, 2);
header.writeUInt16LE(sizes.length, 4);
let offset = header.length;
sizes.forEach((size, index) => {
  const base = 6 + index * 16;
  header[base] = header[base + 1] = size === 256 ? 0 : size;
  header.writeUInt16LE(1, base + 4);
  header.writeUInt16LE(32, base + 6);
  header.writeUInt32LE(frames[index].length, base + 8);
  header.writeUInt32LE(offset, base + 12);
  offset += frames[index].length;
});
await writeFile(join(directory, "app-icon.ico"), Buffer.concat([header, ...frames]));
for (const size of [16, 32, 128, 256, 512]) {
  for (const scale of [1, 2]) {
    const pixels = String(size * scale);
    execFileSync("sips", ["-z", pixels, pixels, source, "--out", join(iconset, `icon_${size}x${size}${scale === 2 ? "@2x" : ""}.png`)]);
  }
}
execFileSync("iconutil", ["-c", "icns", iconset, "-o", join(directory, "app-icon.icns")]);
console.log(`Created ICO (${sizes.join(", ")} px) and ICNS in ${directory}`);
