import { app, BrowserWindow, dialog, ipcMain, nativeTheme, shell } from "electron";
import { basename, dirname, join } from "node:path";
import { copyFile, mkdir, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { LocalDatabase } from "./database.js";
import { createWorkspacePackage, inspectWorkspacePackage, type PreparedWorkspaceImport } from "./workspacePackage.js";
import type {
  SaveDictionaryInput,
  SaveGroupOrderInput,
  SaveGroupOverrideInput,
  SaveRequirementInput,
  HalfYearRange,
  PptExportRequest,
  PptDiagnosticLog,
  SaveMemberInput,
  SavePptAccessPolicyInput,
  WorkspaceImportMode,
} from "../shared/types.js";
import { buildPptExportPlan } from "../shared/pptExport.js";
import { halfYearSequence } from "../shared/roadmap.js";
import { createRoadmapPresentation, presentationSlideCount } from "./pptxExport.js";
import { checkPowerPointIrmEnvironment, protectPowerPointWithIrm, type PowerPointIrmErrorDetails } from "./pptIrm.js";
import { pptPermissionExpiration } from "../shared/pptAccess.js";
import { resolvePptDocumentOwnerAccount } from "../shared/members.js";
import { compactPptDiagnosticValue, formatPptDiagnosticText } from "../shared/pptDiagnostic.js";
import { preservePptDiagnosticArtifacts } from "./pptDiagnosticArtifacts.js";
import { validateGeneratedPptx } from "./pptxValidation.js";
import {
  createWorkspaceWorkbook,
  inspectWorkspaceWorkbook,
  type PreparedWorkspaceWorkbookImport,
} from "./workspaceWorkbook.js";
import type { WorkspaceWorkbookConflictMode } from "../shared/types.js";

const PRODUCT_NAME = "解决方案需求管理";
const APP_RELEASED_AT = "2026-09-01";
const APP_AUTHOR = "林忠维";
const LEGACY_USER_DATA_DIRECTORY = "需求路标工作台";
const PPT_DIAGNOSTIC_FILE = "ppt-export-last-error.txt";
const PPT_DIAGNOSTIC_POINTER_FILE = "ppt-export-last-error.json";

let database: LocalDatabase;
let mainWindow: BrowserWindow | null = null;
const BUILTIN_TEMPLATE_FILE = "需求路标模板-共创版-v0.1.pptx";
const pendingImports = new Map<string, PreparedWorkspaceImport>();
const pendingWorkbookImports = new Map<string, PreparedWorkspaceWorkbookImport>();

async function optionalFileSize(path: string | undefined): Promise<string> {
  if (!path) return "未生成";
  try {
    return `${(await stat(path)).size} 字节`;
  } catch {
    return "未生成";
  }
}

async function writePptDiagnosticLog(input: {
  error: unknown;
  stage: string;
  environment?: Awaited<ReturnType<typeof checkPowerPointIrmEnvironment>>;
  preferredDocumentOwner?: string;
  effectiveDocumentOwner?: string;
  fallbackUsed?: boolean;
  editablePath?: string;
  normalizedPath?: string;
  preferredProtectedPath?: string;
  fallbackProtectedPath?: string;
}): Promise<PptDiagnosticLog> {
  const createdAt = new Date().toISOString();
  const details = input.error as Partial<PowerPointIrmErrorDetails>;
  const userDataPath = app.getPath("userData");
  const logsPath = join(userDataPath, "logs");
  const path = join(logsPath, PPT_DIAGNOSTIC_FILE);
  const pointerPath = join(logsPath, PPT_DIAGNOSTIC_POINTER_FILE);
  const message = input.error instanceof Error ? input.error.message : String(input.error);
  const fallbackUsed = input.fallbackUsed ?? details.fallbackUsed ?? false;
  const protectedPath = fallbackUsed ? input.fallbackProtectedPath : input.preferredProtectedPath;
  const standardizationStatus = details.normalized
    ? "已完成修复/标准化并由新 PowerPoint 实例常规重开"
    : input.normalizedPath
      ? "未完成或未通过常规重开"
      : "未开始";
  const content = formatPptDiagnosticText({
    createdAt,
    appVersion: `v${app.getVersion()}`,
    stage: details.stage || input.stage,
    powerPointAndIrmStatus: `PowerPoint ${input.environment?.powerPointInstalled ? "已检测" : "未确认"}；企业 IRM ${input.environment?.irmReady ? "可用" : "未确认"}`,
    officeIdentity: details.officeIdentity || input.environment?.officeIdentity || "未识别",
    preferredDocumentOwner: details.preferredDocumentOwner || input.preferredDocumentOwner || input.environment?.preferredDocumentOwner || "未识别",
    effectiveDocumentOwner: details.effectiveDocumentOwner || input.effectiveDocumentOwner || "未生成",
    fallbackUsed,
    standardizationStatus,
    editableFileSize: await optionalFileSize(input.editablePath),
    normalizedFileSize: await optionalFileSize(input.normalizedPath),
    protectedFileSize: await optionalFileSize(protectedPath),
    completedSteps: details.diagnostics ?? [],
    errorMessage: message,
  });
  const directoryPath = await preservePptDiagnosticArtifacts({
    userDataPath,
    createdAt,
    appVersion: `v${app.getVersion()}`,
    stage: details.stage || input.stage,
    logContent: content,
    files: [
      { sourcePath: input.editablePath, fileName: "01-generated-original.pptx", label: "原始生成 PPT" },
      { sourcePath: input.normalizedPath, fileName: "02-powerpoint-normalized.pptx", label: "PowerPoint 标准化 PPT" },
      { sourcePath: input.preferredProtectedPath, fileName: "03-preferred-owner-protected.pptx", label: "首选账号保护结果" },
      { sourcePath: input.fallbackProtectedPath, fileName: "04-office-owner-fallback-protected.pptx", label: "Office 账号回退保护结果" },
    ],
  });
  await mkdir(logsPath, { recursive: true });
  await Promise.all([
    writeFile(path, content, "utf8"),
    writeFile(pointerPath, JSON.stringify({ createdAt, directoryPath }), "utf8"),
  ]);
  return { createdAt, path, content, directoryPath };
}

async function readPptDiagnosticLog(): Promise<PptDiagnosticLog | null> {
  const logsPath = join(app.getPath("userData"), "logs");
  const path = join(logsPath, PPT_DIAGNOSTIC_FILE);
  try {
    const content = await readFile(path, "utf8");
    const createdAt = /^时间：(.+)$/m.exec(content)?.[1] ?? "";
    let directoryPath: string | undefined;
    try {
      const pointer = JSON.parse(await readFile(join(logsPath, PPT_DIAGNOSTIC_POINTER_FILE), "utf8")) as { directoryPath?: string };
      directoryPath = pointer.directoryPath;
    } catch {
      directoryPath = undefined;
    }
    return { createdAt, path, content, directoryPath };
  } catch {
    return null;
  }
}

async function clearPptDiagnosticLog(): Promise<void> {
  const logsPath = join(app.getPath("userData"), "logs");
  await Promise.all([
    rm(join(logsPath, PPT_DIAGNOSTIC_FILE), { force: true }),
    rm(join(logsPath, PPT_DIAGNOSTIC_POINTER_FILE), { force: true }),
  ]);
}

async function createWindow(): Promise<void> {
  const appWindow = new BrowserWindow({
    width: 1480,
    height: 960,
    minWidth: 1120,
    minHeight: 760,
    backgroundColor: "#f4f6f8",
    title: PRODUCT_NAME,
    webPreferences: {
      preload: join(app.getAppPath(), "dist-electron", "preload", "index.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  mainWindow = appWindow;

  const notifyFullscreenChange = (enabled: boolean) => {
    if (!appWindow.webContents.isDestroyed()) {
      appWindow.webContents.send("window:fullscreen:changed", enabled);
    }
  };
  appWindow.on("enter-full-screen", () => notifyFullscreenChange(true));
  appWindow.on("leave-full-screen", () => notifyFullscreenChange(false));

  if (process.env.VITE_DEV_SERVER_URL) {
    await appWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    await appWindow.loadFile(join(app.getAppPath(), "dist-renderer", "index.html"));
  }
}

function registerIpc(): void {
  ipcMain.handle("app:info", () => ({
    name: PRODUCT_NAME,
    version: app.getVersion(),
    releasedAt: APP_RELEASED_AT,
    author: APP_AUTHOR,
  }));
  ipcMain.handle("snapshot:get", () => database.getSnapshot());
  ipcMain.handle("storage:open-local-directory", async () => {
    const directory = app.getPath("userData");
    await mkdir(directory, { recursive: true });
    const error = await shell.openPath(directory);
    if (error) throw new Error(`无法打开本地数据目录：${error}`);
    return { path: directory };
  });
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
  ipcMain.handle("ppt-access:check-environment", async () => {
    const status = await checkPowerPointIrmEnvironment();
    const preferredDocumentOwner = resolvePptDocumentOwnerAccount(status.officeIdentity, database.getSnapshot().members);
    return {
      ...status,
      preferredDocumentOwner,
      message: status.irmReady && preferredDocumentOwner
        ? `PowerPoint 企业 IRM 可用；Office 登录身份：${status.officeIdentity || "未识别"}；首选文档所有者：${preferredDocumentOwner}；必要时会自动回退为 Office 登录身份。`
        : status.message,
    };
  });
  ipcMain.handle("ppt-access:diagnostic-log", () => readPptDiagnosticLog());
  ipcMain.handle("ppt-access:open-diagnostic-directory", async () => {
    const diagnostic = await readPptDiagnosticLog();
    if (!diagnostic) throw new Error("当前没有可打开的 PPT 诊断记录");
    const directory = diagnostic.directoryPath ?? dirname(diagnostic.path);
    await mkdir(directory, { recursive: true });
    const error = await shell.openPath(directory);
    if (error) throw new Error(`无法打开 PPT 诊断目录：${error}`);
    return { path: directory };
  });
  ipcMain.handle("groups:save", (_event, input: SaveGroupOverrideInput) => database.saveGroupOverride(input));
  ipcMain.handle("groups:reorder", (_event, input: SaveGroupOrderInput) => database.saveGroupOrder(input));

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

  ipcMain.handle("workspace:apply", async (_event, input: { token: string; mode: WorkspaceImportMode; differenceDigest?: string }) => {
    if (!input || !["merge", "replace"].includes(input.mode)) throw new Error("请选择有效的导入方式");
    const prepared = pendingImports.get(input.token);
    if (!prepared) throw new Error("导入数据已失效，请重新选择数据包");
    if (input.mode === "merge" && input.differenceDigest !== prepared.preview.differenceDigest) throw new Error("请先核对并确认全部合并差异");
    pendingImports.delete(input.token);
    return { snapshot: await database.importWorkspace(prepared.data, input.mode) };
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

  ipcMain.handle("workbook:inspect", async () => {
    const result = await dialog.showOpenDialog(mainWindow!, {
      title: "选择企业协作 Excel",
      properties: ["openFile"],
      filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }],
    });
    if (result.canceled || !result.filePaths[0]) return { canceled: true };
    const filePath = result.filePaths[0];
    const prepared = await inspectWorkspaceWorkbook(await readFile(filePath), basename(filePath), database.getWorkspaceData());
    const token = crypto.randomUUID();
    pendingWorkbookImports.clear();
    pendingWorkbookImports.set(token, prepared);
    return { canceled: false, preview: { token, fileName: basename(filePath), ...prepared.preview } };
  });

  ipcMain.handle("workbook:apply", async (_event, input: { token: string; conflictMode: WorkspaceWorkbookConflictMode; differenceDigest: string }) => {
    if (!input || !["local-wins", "excel-wins"].includes(input.conflictMode)) throw new Error("请选择有效的冲突处理方式");
    const prepared = pendingWorkbookImports.get(input.token);
    if (!prepared) throw new Error("Excel 导入数据已失效，请重新选择文件");
    if (input.differenceDigest !== prepared.preview.differenceDigest) throw new Error("请先核对并确认全部 Excel 差异");
    if (prepared.parsed.errors.length) throw new Error("Excel 存在校验错误，请修正后重新导入");
    pendingWorkbookImports.delete(input.token);
    return { snapshot: await database.importWorkspaceWorkbook(prepared, input.conflictMode) };
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
    let stage = "准备导出";
    let environment: Awaited<ReturnType<typeof checkPowerPointIrmEnvironment>> | undefined;
    let preferredDocumentOwner: string | undefined;
    let effectiveDocumentOwner: string | undefined;
    let fallbackUsed = false;
    let tempDirectory: string | undefined;
    let editablePath: string | undefined;
    let normalizedPath: string | undefined;
    let preferredProtectedPath: string | undefined;
    let fallbackProtectedPath: string | undefined;
    try {
      const snapshot = database.getSnapshot();
      if (!snapshot.pptAccessPolicy.entries.length) throw new Error("请先配置并保存 PPT 成员白名单");
      stage = "检测 PowerPoint 与 IRM";
      environment = await checkPowerPointIrmEnvironment();
      if (!environment.supported || !environment.powerPointInstalled || !environment.irmReady) throw new Error(environment.message);
      preferredDocumentOwner = resolvePptDocumentOwnerAccount(environment.officeIdentity, snapshot.members);
      stage = "计算 PPT 页面";
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
      tempDirectory = await mkdtemp(join(tmpdir(), "solution-ppt-export-"));
      editablePath = join(tempDirectory, "roadmap-editable.pptx");
      normalizedPath = join(tempDirectory, "roadmap-normalized.pptx");
      preferredProtectedPath = join(tempDirectory, "roadmap-protected-preferred-owner.pptx");
      fallbackProtectedPath = join(tempDirectory, "roadmap-protected-office-owner.pptx");
      stage = "生成原始 PPT";
      const editableBytes = await createRoadmapPresentation(snapshot, input, { exportedAt, expiration });
      stage = "校验原始 PPT";
      await validateGeneratedPptx(editableBytes, slideCount);
      await writeFile(editablePath, editableBytes);
      stage = "标准化并写入 IRM 权限";
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
      effectiveDocumentOwner = protection.effectiveDocumentOwner;
      fallbackUsed = protection.fallbackUsed;
      stage = "复制最终 PPT";
      await copyFile(protection.outputPath, outputPath);
      await clearPptDiagnosticLog();
      return {
        canceled: false,
        path: outputPath,
        slideCount,
        effectiveDocumentOwner,
        fallbackUsed,
      };
    } catch (error) {
      const diagnostic = await writePptDiagnosticLog({
        error,
        stage,
        environment,
        preferredDocumentOwner,
        effectiveDocumentOwner,
        editablePath,
        normalizedPath,
        preferredProtectedPath,
        fallbackProtectedPath,
      });
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(`${compactPptDiagnosticValue(message)}\n诊断日志：${diagnostic.path}\n诊断目录：${diagnostic.directoryPath}`);
    } finally {
      if (tempDirectory) await rm(tempDirectory, { recursive: true, force: true });
    }
  });
}

app.whenReady().then(async () => {
  nativeTheme.themeSource = "light";
  app.setPath("userData", process.env.SOLUTION_REQUIREMENTS_USER_DATA_DIR || join(app.getPath("appData"), LEGACY_USER_DATA_DIRECTORY));
  database = new LocalDatabase();
  await database.initialize();
  registerIpc();
  await createWindow();
  app.on("activate", async () => {
    if (BrowserWindow.getAllWindows().length === 0) await createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
