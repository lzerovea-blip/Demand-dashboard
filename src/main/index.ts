import { app, BrowserWindow, dialog, ipcMain, nativeTheme, shell } from "electron";
import { basename, join } from "node:path";
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { createHash } from "node:crypto";
import { LocalDatabase } from "./database.js";
import { createWorkspacePackage, inspectWorkspacePackage, type PreparedWorkspaceImport } from "./workspacePackage.js";
import type {
  SaveDictionaryInput,
  SaveGroupOrderInput,
  SaveGroupOverrideInput,
  SaveRequirementInput,
  HalfYearRange,
  PptExportRequest,
  SaveMemberInput,
  SavePptAccessPolicyInput,
  SaveTieringSupportInput,
  SaveTieringEvidenceSourceInput,
  UpdateTieringIssueInput,
  UpdateTieringChangeLogInput,
  RendererDiagnosticInput,
  LegacyRequirementDecision,
  WorkspaceImportMode,
  WorkspaceWorkbookFieldSelection,
} from "../shared/types.js";
import { buildPptExportPlan } from "../shared/pptExport.js";
import { halfYearSequence } from "../shared/roadmap.js";
import { createRoadmapPresentation, presentationSlideCount } from "./pptxExport.js";
import { validateGeneratedPptx } from "./pptxValidation.js";
import { checkPowerPointIrmEnvironment, protectPowerPointWithIrm, resolveIrmEnvironmentStatus } from "./pptIrm.js";
import { pptPermissionExpiration } from "../shared/pptAccess.js";
import {
  createWorkspaceWorkbook,
  inspectWorkspaceWorkbook,
  type PreparedWorkspaceWorkbookImport,
} from "./workspaceWorkbook.js";
import { createTieringWorkbook } from "./tieringWorkbook.js";
import type { WorkspaceWorkbookConflictMode } from "../shared/types.js";
import { beginDiagnosticStage, diagnosticInfo, initializeStartupDiagnostics, sanitizeDiagnosticDetail, writeStartupDiagnostic } from "./diagnostics.js";

const PRODUCT_NAME = "解决方案需求管理";
const APP_RELEASED_AT = "2026-10-21";
const APP_AUTHOR = "林忠维";
const LEGACY_USER_DATA_DIRECTORY = "需求路标工作台";

let database: LocalDatabase;
let mainWindow: BrowserWindow | null = null;
let startupPhase = "process-start";
const BUILTIN_TEMPLATE_FILE = "需求路标模板-共创版-v0.1.pptx";
const pendingImports = new Map<string, PreparedWorkspaceImport>();
const pendingWorkbookImports = new Map<string, PreparedWorkspaceWorkbookImport>();

function safeErrorMessage(error: unknown): string {
  return sanitizeDiagnosticDetail(error);
}

async function getResolvedIrmEnvironment() {
  const status = await checkPowerPointIrmEnvironment();
  return resolveIrmEnvironmentStatus(status, database.getSnapshot().members);
}

async function createWindow(): Promise<void> {
  const appWindow = new BrowserWindow({
    width: 1480,
    height: 960,
    minWidth: 1120,
    minHeight: 760,
    backgroundColor: "#f4f6f8",
    title: PRODUCT_NAME,
    icon: app.isPackaged
      ? join(process.resourcesPath, "app-icon.png")
      : join(app.getAppPath(), "build-resources", "app-icon.png"),
    show: true,
    webPreferences: {
      preload: join(app.getAppPath(), "dist-electron", "preload", "index.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  mainWindow = appWindow;
  writeStartupDiagnostic("window-created", `bounds=${appWindow.getBounds().width}x${appWindow.getBounds().height}`);

  const notifyFullscreenChange = (enabled: boolean) => {
    if (!appWindow.webContents.isDestroyed()) {
      appWindow.webContents.send("window:fullscreen:changed", enabled);
    }
  };
  appWindow.on("enter-full-screen", () => notifyFullscreenChange(true));
  appWindow.on("leave-full-screen", () => notifyFullscreenChange(false));
  appWindow.on("unresponsive", () => {
    writeStartupDiagnostic("window-unresponsive", "main-window");
  });
  appWindow.on("responsive", () => writeStartupDiagnostic("window-responsive", "main-window"));
  appWindow.once("ready-to-show", () => writeStartupDiagnostic("window-ready-to-show", "main-window"));
  appWindow.webContents.on("did-fail-load", (_event, errorCode, errorDescription) => {
    writeStartupDiagnostic("renderer-load-failed", `${errorCode} ${errorDescription}`);
  });
  appWindow.webContents.on("did-finish-load", () => writeStartupDiagnostic("preload-and-document-loaded", "main-window"));
  appWindow.webContents.on("render-process-gone", (_event, details) => {
    writeStartupDiagnostic("renderer-process-gone", `${details.reason} exitCode=${details.exitCode}`);
  });

  if (process.env.VITE_DEV_SERVER_URL) {
    await appWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    await appWindow.loadFile(join(app.getAppPath(), "dist-renderer", "index.html"));
  }
  appWindow.center();
  appWindow.show();
  appWindow.focus();
}

function registerIpc(): void {
  ipcMain.handle("app:info", () => ({
    name: PRODUCT_NAME,
    version: app.getVersion(),
    releasedAt: APP_RELEASED_AT,
    author: APP_AUTHOR,
  }));
  ipcMain.handle("snapshot:get", () => database.getSnapshot());
  ipcMain.handle("roadmaps:save", (_event, input) => database.saveRoadmap(input));
  ipcMain.handle("roadmaps:delete", (_event, id) => database.deleteRoadmap(id));
  ipcMain.handle("roadmap-lanes:save", (_event, input) => database.saveRoadmapLane(input));
  ipcMain.handle("roadmap-lanes:delete", (_event, id) => database.deleteRoadmapLane(id));
  ipcMain.handle("roadmap-assignments:save", (_event, input) => database.saveRoadmapAssignments(input));
  ipcMain.handle("medical-roadmap-plans:save", (_event, input) => database.saveMedicalRoadmapPlan(input));
  ipcMain.handle("medical-roadmap-plans:delete", (_event, input) => database.deleteMedicalRoadmapPlan(input));
  ipcMain.handle("medical-roadmap-plans:reorder", (_event, input) => database.saveMedicalRoadmapPlanOrder(input));
  ipcMain.handle("window:fullscreen:get", (event) => BrowserWindow.fromWebContents(event.sender)?.isFullScreen() ?? false);
  ipcMain.handle("window:fullscreen:set", (event, enabled: boolean) => {
    if (typeof enabled !== "boolean") throw new Error("请提供有效的全屏状态");
    const targetWindow = BrowserWindow.fromWebContents(event.sender);
    if (!targetWindow) throw new Error("无法获取当前应用窗口");
    targetWindow.setFullScreen(enabled);
  });
  ipcMain.handle("requirements:save", (_event, input: SaveRequirementInput) => database.saveRequirement(input));
  ipcMain.handle("requirements:duplicate", (_event, id: string) => database.duplicateRequirement(id));
  ipcMain.handle("requirements:delete", (_event, id: string) => database.deleteRequirement(id));
  ipcMain.handle("domains:save", (_event, input: SaveDictionaryInput) => database.saveDomain(input));
  ipcMain.handle("domains:delete", (_event, id: string) => database.deleteDomain(id));
  ipcMain.handle("products:save", (_event, input: SaveDictionaryInput) => database.saveProduct(input));
  ipcMain.handle("products:delete", (_event, id: string) => database.deleteProduct(id));
  ipcMain.handle("members:save", (_event, input: SaveMemberInput) => database.saveMember(input));
  ipcMain.handle("members:delete", (_event, id: string) => database.deleteMember(id));
  ipcMain.handle("ppt-access:save", (_event, input: SavePptAccessPolicyInput) => database.savePptAccessPolicy(input));
  ipcMain.handle("ppt-access:check-environment", () => getResolvedIrmEnvironment());
  ipcMain.handle("groups:save", (_event, input: SaveGroupOverrideInput) => database.saveGroupOverride(input));
  ipcMain.handle("groups:reorder", (_event, input: SaveGroupOrderInput) => database.saveGroupOrder(input));
  ipcMain.handle("medical-roadmap:reorder", (_event, input: { requirementIds: string[] }) => database.saveMedicalRoadmapOrder(input));
  ipcMain.handle("tiering:support:save", (_event, input: SaveTieringSupportInput) => database.saveTieringSupport(input));
  ipcMain.handle("tiering:source:save", (_event, input: SaveTieringEvidenceSourceInput) => database.saveTieringEvidenceSource(input));
  ipcMain.handle("tiering:issue:update", (_event, input: UpdateTieringIssueInput) => database.updateTieringIssue(input));
  ipcMain.handle("tiering:change-log:update", (_event, input: UpdateTieringChangeLogInput) => database.updateTieringChangeLog(input));
  ipcMain.handle("tiering:change-log:delete", (_event, logId: string) => database.deleteTieringChangeLog(logId));
  ipcMain.handle("tiering:requirement:regenerate", (_event, featureId: string) => database.regenerateTieringRequirement(featureId));
  ipcMain.handle("diagnostics:renderer-report", async (_event, input: RendererDiagnosticInput) => {
    const allowed = new Set(["renderer-ready", "renderer-error", "renderer-unhandled-rejection", "react-error"]);
    if (!input || !allowed.has(input.event)) throw new Error("无效的前端诊断事件");
    writeStartupDiagnostic(input.event, input.detail ?? "");
    if (input.event === "renderer-ready") writeStartupDiagnostic("window-interactive", "renderer-reported-ready");
  });
  ipcMain.handle("diagnostics:info", () => diagnosticInfo());
  ipcMain.handle("diagnostics:open-directory", async () => {
    const { directory } = diagnosticInfo();
    const result = await shell.openPath(directory);
    if (result) throw new Error(`无法打开诊断日志目录：${result}`);
    return { directory, currentLogPath: join(directory, "startup.log") };
  });

  ipcMain.handle("workspace:export", async () => {
    const result = await dialog.showSaveDialog(mainWindow!, {
      title: "导出完整工作区数据包",
      defaultPath: `需求路标工作区-${new Date().toISOString().slice(0, 10)}.roadmap`,
      filters: [{ name: "需求路标工作区", extensions: ["roadmap"] }],
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    const outputPath = result.filePath.toLowerCase().endsWith(".roadmap") ? result.filePath : `${result.filePath}.roadmap`;
    await writeFile(outputPath, await createWorkspacePackage(database.getWorkspaceData(), app.getVersion()));
    return { canceled: false, path: outputPath };
  });

  ipcMain.handle("workspace:inspect", async () => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: "选择需求路标数据包",
      properties: ["openFile"],
      filters: [
        { name: "需求路标数据包", extensions: ["roadmap", "json"] },
        { name: "需求路标工作区", extensions: ["roadmap"] },
        { name: "旧版JSON备份", extensions: ["json"] },
      ],
    });
    if (result.canceled || !result.filePaths[0]) return { canceled: true };
    const filePath = result.filePaths[0];
    const prepared = await inspectWorkspacePackage(await readFile(filePath), basename(filePath), database.getWorkspaceData());
    const token = crypto.randomUUID();
    pendingImports.clear();
    pendingImports.set(token, prepared);
    return { canceled: false, preview: { token, fileName: basename(filePath), ...prepared.preview } };
  });

  ipcMain.handle("workspace:apply", async (_event, input: { token: string; mode: WorkspaceImportMode; differenceDigest?: string; legacyRequirementDecisions?: Record<string, LegacyRequirementDecision> }) => {
    if (!input || !["merge", "replace"].includes(input.mode)) throw new Error("请选择有效的导入方式");
    const prepared = pendingImports.get(input.token);
    if (!prepared) throw new Error("导入数据已失效，请重新选择数据包");
    if (prepared.preview.packageKind === "tiering-seed" && input.mode !== "merge") throw new Error("分层分级初始包只能选择合并导入，不能整体替换");
    if (input.mode === "merge" && input.differenceDigest !== prepared.preview.differenceDigest) throw new Error("请先核对并确认全部合并差异");
    if (input.mode === "merge") {
      const decisions = input.legacyRequirementDecisions ?? {};
      const unresolved = prepared.preview.ambiguousRequirementIds.filter((id) => !["merge", "keep-deleted"].includes(decisions[id]));
      if (unresolved.length) throw new Error(`仍有 ${unresolved.length} 条旧版新增需求未选择“合入”或“保持删除”`);
    }
    pendingImports.delete(input.token);
    return { snapshot: await database.importWorkspace(prepared.data, input.mode, input.legacyRequirementDecisions) };
  });

  ipcMain.handle("workbook:export", async (_event, input: HalfYearRange) => {
    if (!input) throw new Error("请选择有效的 Excel 导出区间");
    halfYearSequence(input.start, input.end);
    const result = await dialog.showSaveDialog(mainWindow!, {
      title: "导出企业协作 Excel",
      defaultPath: `需求路标协作-${input.start.slice(2)}-${input.end.slice(2)}-${new Date().toISOString().slice(0, 10)}.xlsx`,
      filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }],
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    const outputPath = result.filePath.toLowerCase().endsWith(".xlsx") ? result.filePath : `${result.filePath}.xlsx`;
    await writeFile(outputPath, await createWorkspaceWorkbook(database.getWorkspaceData(), app.getVersion(), new Date().toISOString(), input));
    return { canceled: false, path: outputPath };
  });

  ipcMain.handle("tiering-workbook:export", async (_event, input: HalfYearRange) => {
    if (!input) throw new Error("请选择有效的分层分级 Excel 导出区间");
    halfYearSequence(input.start, input.end);
    const generatedAt = new Date().toISOString();
    const result = await dialog.showSaveDialog(mainWindow!, {
      title: "导出分层分级 Excel",
      defaultPath: `分层分级-${input.start.slice(2)}-${input.end.slice(2)}-${generatedAt.slice(0, 10)}.xlsx`,
      filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }],
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    const outputPath = result.filePath.toLowerCase().endsWith(".xlsx") ? result.filePath : `${result.filePath}.xlsx`;
    await writeFile(outputPath, await createTieringWorkbook(database.getWorkspaceData(), app.getVersion(), generatedAt, input));
    return { canceled: false, path: outputPath };
  });

  ipcMain.handle("workbook:inspect", async () => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: "选择企业协作 Excel",
      properties: ["openFile"],
      filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }],
    });
    if (result.canceled || !result.filePaths[0]) return { canceled: true };
    const filePath = result.filePaths[0];
    writeStartupDiagnostic("workbook-inspect-start", "file-selected");
    try {
      const bytes = await readFile(filePath);
      const prepared = await inspectWorkspaceWorkbook(bytes, basename(filePath), database.getWorkspaceData());
      if (["v8", "v9", "v10", "v11", "v12"].includes(prepared.parsed.formatVersion)) {
        prepared.sourceFilePath = filePath;
        prepared.sourceSha256 = createHash("sha256").update(bytes).digest("hex");
      }
      const token = crypto.randomUUID();
      pendingWorkbookImports.clear();
      pendingWorkbookImports.set(token, prepared);
      writeStartupDiagnostic("workbook-inspect-complete", `format=${prepared.parsed.formatVersion} sizeBytes=${bytes.length} errors=${prepared.preview.errors.length} changes=${prepared.preview.fieldChanges?.length ?? prepared.preview.differences.length}`);
      return { canceled: false, preview: { token, fileName: basename(filePath), ...prepared.preview } };
    } catch (error) {
      writeStartupDiagnostic("workbook-inspect-failed", error);
      throw error;
    }
  });

  ipcMain.handle("workbook:apply", async (_event, input: { token: string; conflictMode: WorkspaceWorkbookConflictMode; differenceDigest: string; fieldSelection?: WorkspaceWorkbookFieldSelection }) => {
    writeStartupDiagnostic("workbook-apply-start", `hasInput=${Boolean(input)}`);
    try {
      if (!input || !["local-wins", "excel-wins"].includes(input.conflictMode)) throw new Error("请选择有效的冲突处理方式");
      const prepared = pendingWorkbookImports.get(input.token);
      if (!prepared) throw new Error("Excel 导入数据已失效，请重新选择文件");
      writeStartupDiagnostic("workbook-apply-selection", `format=${prepared.parsed.formatVersion} selected=${input.fieldSelection?.selectedChangeIds.length ?? 0} previewErrors=${prepared.parsed.errors.length}`);
      if (input.differenceDigest !== prepared.preview.differenceDigest) throw new Error("请先核对并确认全部 Excel 差异");
      if (prepared.parsed.errors.length) throw new Error("Excel 存在校验错误，请修正后重新导入");
      if (["v8", "v9", "v10", "v11", "v12"].includes(prepared.parsed.formatVersion)) {
        if (!prepared.sourceFilePath || !prepared.sourceSha256) throw new Error("Excel 文件校验信息缺失，请重新选择文件");
        const latestBytes = await readFile(prepared.sourceFilePath).catch(() => { throw new Error("Excel 文件已移动或删除，请重新选择文件"); });
        if (createHash("sha256").update(latestBytes).digest("hex") !== prepared.sourceSha256) throw new Error("Excel 文件在预览后发生变化，请重新选择文件检查");
      }
      const snapshot = await database.importWorkspaceWorkbook(prepared, input.conflictMode, input.fieldSelection);
      pendingWorkbookImports.delete(input.token);
      writeStartupDiagnostic("workbook-apply-complete", `format=${prepared.parsed.formatVersion} requirements=${snapshot.requirements.length}`);
      return { snapshot };
    } catch (error) {
      writeStartupDiagnostic("workbook-apply-failed", error);
      throw error;
    }
  });

  ipcMain.handle("template:draft-export", async () => {
    const result = await dialog.showSaveDialog(mainWindow!, {
      title: "导出PPT模板共创版",
      defaultPath: BUILTIN_TEMPLATE_FILE,
      filters: [{ name: "PowerPoint演示文稿", extensions: ["pptx"] }],
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    const sourcePath = app.isPackaged
      ? join(process.resourcesPath, "templates", BUILTIN_TEMPLATE_FILE)
      : join(app.getAppPath(), BUILTIN_TEMPLATE_FILE);
    await copyFile(sourcePath, result.filePath);
    return { canceled: false, path: result.filePath };
  });

  ipcMain.handle("presentation:export", async (_event, input: PptExportRequest) => {
    const snapshot = database.getSnapshot();
    if (!snapshot.pptAccessPolicy.entries.length) throw new Error("请先配置并保存 PPT 成员白名单");
    const environment = await getResolvedIrmEnvironment();
    if (!environment.supported || !environment.powerPointInstalled || !environment.irmReady) throw new Error(environment.message);
    const plan = buildPptExportPlan(snapshot, input);
    const slideCount = presentationSlideCount(plan);
    const exportedAt = new Date();
    const expiration = pptPermissionExpiration(exportedAt);
    const result = await dialog.showSaveDialog(mainWindow!, {
      title: "导出需求路标PPT",
      defaultPath: `${plan.materialName.replace(/[\\/:*?\"<>|]/g, "-")}-${input.start}-${input.end}-${new Date().toISOString().slice(0, 10)}.pptx`,
      filters: [{ name: "PowerPoint演示文稿", extensions: ["pptx"] }],
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    const outputPath = result.filePath.toLowerCase().endsWith(".pptx") ? result.filePath : `${result.filePath}.pptx`;
    const tempDirectory = await mkdtemp(join(tmpdir(), "solution-ppt-export-"));
    const editablePath = join(tempDirectory, "roadmap-editable.pptx");
    const normalizedPath = join(tempDirectory, "roadmap-normalized.pptx");
    const preferredProtectedPath = join(tempDirectory, "roadmap-protected-preferred-owner.pptx");
    const fallbackProtectedPath = join(tempDirectory, "roadmap-protected-office-owner.pptx");
    try {
      const presentationBytes = await createRoadmapPresentation(snapshot, input, { exportedAt, expiration });
      await validateGeneratedPptx(presentationBytes, slideCount);
      await writeFile(editablePath, presentationBytes);
      const protection = await protectPowerPointWithIrm({
        inputPath: editablePath,
        normalizedPath,
        preferredOutputPath: preferredProtectedPath,
        fallbackOutputPath: fallbackProtectedPath,
        members: snapshot.members,
        policy: snapshot.pptAccessPolicy,
        expiration,
        officeIdentity: environment.officeIdentity,
      });
      await copyFile(protection.outputPath, outputPath);
      return {
        canceled: false,
        path: outputPath,
        slideCount,
        effectiveDocumentOwner: protection.effectiveDocumentOwner,
        fallbackUsed: protection.fallbackUsed,
      };
    } finally {
      await rm(tempDirectory, { recursive: true, force: true });
    }
  });
}

app.whenReady().then(async () => {
  startupPhase = "electron-ready";
  nativeTheme.themeSource = "light";
  if (process.platform === "win32") app.setAppUserModelId("com.local.healthroadmap");
  app.setPath("userData", process.env.SOLUTION_REQUIREMENTS_USER_DATA_DIR || join(app.getPath("appData"), LEGACY_USER_DATA_DIRECTORY));
  initializeStartupDiagnostics(app.getPath("userData"), app.getVersion());
  writeStartupDiagnostic("process-ready", `pid=${process.pid}`);
  startupPhase = "user-data-initialization";
  const finishUserData = beginDiagnosticStage("user-data-initialization");
  await mkdir(app.getPath("userData"), { recursive: true });
  finishUserData();
  database = new LocalDatabase();
  startupPhase = "database-read-migration-and-built-in-series";
  const finishDatabase = beginDiagnosticStage("database-read-migration-and-built-in-series");
  const initialization = await database.initialize();
  finishDatabase();
  if (initialization.skippedTieringRequirements > 0) {
    writeStartupDiagnostic(
      "legacy-domain-links-skipped",
      `检测到 ${initialization.skippedTieringRequirements} 项需求或分层规格的关联失效；原数据和支持状态已保留，请在需求池或分层分级修复关联`,
    );
  }
  startupPhase = "ipc-registration";
  const finishIpc = beginDiagnosticStage("ipc-registration");
  registerIpc();
  finishIpc();
  startupPhase = "window-creation-and-load";
  const finishWindow = beginDiagnosticStage("window-creation-and-load");
  await createWindow();
  finishWindow();
  startupPhase = "running";
  app.on("activate", async () => {
    if (BrowserWindow.getAllWindows().length === 0) await createWindow();
  });
}).catch(async (error: unknown) => {
  const message = safeErrorMessage(error);
  const diagnosticPath = writeStartupDiagnostic("startup-failed", `stage=${startupPhase} ${message}`);
  dialog.showErrorBox(
    `${PRODUCT_NAME}无法启动`,
    [
      "应用启动初始化失败，已停止后台进程。",
      `失败阶段：${startupPhase}`,
      message,
      diagnosticPath ? `诊断日志：${diagnosticPath}` : "未能写入诊断日志。",
    ].join("\n\n"),
  );
  app.exit(1);
});

process.on("uncaughtException", (error) => {
  const message = safeErrorMessage(error);
  const diagnosticPath = writeStartupDiagnostic("main-uncaught-exception", message);
  if (app.isReady()) {
    dialog.showErrorBox(
      `${PRODUCT_NAME}发生致命错误`,
      `主进程已停止以避免后台残留。\n\n${message}\n\n${diagnosticPath ? `诊断日志：${diagnosticPath}` : "未能写入诊断日志。"}`,
    );
  }
  app.exit(1);
});
process.on("unhandledRejection", (reason) => {
  writeStartupDiagnostic("main-unhandled-rejection", reason);
});
app.on("child-process-gone", (_event, details) => {
  writeStartupDiagnostic("child-process-gone", `${details.type} ${details.reason} exitCode=${details.exitCode}`);
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
