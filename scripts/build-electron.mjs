import { build } from "esbuild";
import { rm } from "node:fs/promises";

await rm("dist-electron", { recursive: true, force: true });

await build({
  entryPoints: ["src/main/index.ts"],
  outfile: "dist-electron/main/index.cjs",
  bundle: true,
  platform: "node",
  target: "node22",
  format: "cjs",
  external: ["electron"],
  legalComments: "none",
  sourcemap: false,
});

await build({
  entryPoints: ["src/preload/index.cts"],
  outfile: "dist-electron/preload/index.cjs",
  bundle: true,
  platform: "node",
  target: "node22",
  format: "cjs",
  external: ["electron"],
  legalComments: "none",
  sourcemap: false,
});
