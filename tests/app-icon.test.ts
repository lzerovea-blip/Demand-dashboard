import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

describe("应用图标资源", () => {
  it("主图使用带透明通道的正方形 PNG", () => {
    const png = readFileSync("build-resources/app-icon.png");
    expect(png.subarray(0, 8).toString("hex")).toBe("89504e470d0a1a0a");
    expect(png.readUInt32BE(16)).toBe(png.readUInt32BE(20));
    expect(png.readUInt32BE(16)).toBeGreaterThanOrEqual(256);
    expect(png[25]).toBe(6);
  });
  it("ICO 提供完整 16–256px 图像，所有偏移与 PNG 大小一致", () => {
    const icon = readFileSync("build-resources/app-icon.ico");
    expect(icon.readUInt16LE(0)).toBe(0);
    expect(icon.readUInt16LE(2)).toBe(1);
    const count = icon.readUInt16LE(4);
    expect(count).toBe(7);
    const dimensions: number[] = [];
    let end = 6 + count * 16;
    for (let index = 0; index < count; index++) {
      const base = 6 + index * 16;
      const size = icon[base] || 256;
      const length = icon.readUInt32LE(base + 8);
      const offset = icon.readUInt32LE(base + 12);
      expect(offset).toBe(end);
      const png = icon.subarray(offset, offset + length);
      expect(png.subarray(0, 8).toString("hex")).toBe("89504e470d0a1a0a");
      expect(png.readUInt32BE(16)).toBe(size);
      expect(png.readUInt32BE(20)).toBe(size);
      expect(icon.readUInt16LE(base + 6)).toBe(32);
      dimensions.push(size);
      end += length;
    }
    expect(dimensions).toEqual([16, 24, 32, 48, 64, 128, 256]);
    expect(end).toBe(icon.length);
  });
  it("打包配置、运行时及本地 Mac 使用同一组资源", () => {
    const config = readFileSync("electron-builder.yml", "utf8");
    const main = readFileSync("src/main/index.ts", "utf8");
    const mac = readFileSync("scripts/package-mac.mjs", "utf8");
    expect(config).toContain("icon: build-resources/app-icon.ico");
    expect(config).toContain("from: build-resources/app-icon.png");
    expect(main).toContain('join(process.resourcesPath, "app-icon.png")');
    expect(main).toContain('app.setAppUserModelId("com.local.healthroadmap")');
    expect(mac).toContain('["CFBundleIconFile", "app-icon.icns"]');
    const icns = readFileSync("build-resources/app-icon.icns");
    expect(icns.subarray(0, 4).toString()).toBe("icns");
    expect(icns.readUInt32BE(4)).toBe(icns.length);
  });
});
