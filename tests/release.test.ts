import { describe, expect, it } from "vitest";
import { windowsSetupDownloadUrl } from "../src/shared/release";

describe("windowsSetupDownloadUrl", () => {
  it("builds the direct Windows setup download URL for the app version", () => {
    expect(windowsSetupDownloadUrl("0.7.1")).toBe(
      "https://github.com/lzerovea-blip/Demand-dashboard/releases/download/v0.7.1/Solution-Requirements-Manager-Setup-0.7.1-x64.exe",
    );
  });

  it("accepts a version value that already includes the v prefix", () => {
    expect(windowsSetupDownloadUrl("v0.7.2")).toContain("/releases/download/v0.7.2/");
  });
});
