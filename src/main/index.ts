import { app, BrowserWindow, dialog, ipcMain, nativeTheme } from "electron";
import { basename, join } from "node:path";
import { copyFile, readFile, writeFile } from "node:fs/promises";
import { LocalDatabase } from "./database.js";
import { createWorkspacePackage, inspectWorkspacePackage, type PreparedWorkspaceImport } from "./workspacePackage.js";
import type {
  SaveDictionaryInput,
  SaveGroupOverrideInput,
  SaveRequirementInput,
  HalfYearRange,
  WorkspaceImportMode,
} from "../shared/types.js";
import { buildPptExportPlan } from "../shared/pptExport.js";
import { createRoadmapPresentation } from "./pptxExport.js";
import {
  createWorkspaceWorkbook,
  inspectWorkspaceWorkbook,
  type PreparedWorkspaceWorkbookImport,
} from "./workspaceWorkbook.js";
import type { WorkspaceWorkbookConflictMode } from "../shared/types.js";

const database = new LocalDatabase();
let mainWindow: BrowserWindow | null = null;
const BUILTIN_TEMPLATE_FILE = "需求路标模板-共创版-v0.1.pptx";
const pendingImports = new Map<string, PreparedWorkspaceImport>();
const pendingWorkbookImports = new Map<string, PreparedWorkspaceWorkbookImport>();

async function createWindow(): Promise<void> {
  mainWindow = new BrowserWindow({
    width: 1480,
    height: 960,
    minWidth: 1120,
    minHeight: 760,
    backgroundColor: "#f4f6f8",
    title: "需求路标工作台",
    webPreferences: {
      preload: join(app.getAppPath(), "dist-electron", "preload", "index.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  if (process.env.VITE_DEV_SERVER_URL) {
    await mainWindow.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    await mainWindow.loadFile(join(app.getAppPath(), "dist-renderer", "index.html"));
  }
}

function registerIpc(): void {
  ipcMain.handle("snapshot:get", () => database.getSnapshot());
  ipcMain.handle("requirements:save", (_event, input: SaveRequirementInput) => database.saveRequirement(input));
  ipcMain.handle("requirements:delete", (_event, id: string) => database.deleteRequirement(id));
  ipcMain.handle("domains:save", (_event, input: SaveDictionaryInput) => database.saveDomain(input));
  ipcMain.handle("domains:delete", (_event, id: string) => database.deleteDomain(id));
  ipcMain.handle("products:save", (_event, input: SaveDictionaryInput) => database.saveProduct(input));
  ipcMain.handle("products:delete", (_event, id: string) => database.deleteProduct(id));
  ipcMain.handle("groups:save", (_event, input: SaveGroupOverrideInput) => database.saveGroupOverride(input));

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
    const prepared = await inspectWorkspacePackage(await readFile(filePath), basename(filePath));
    const token = crypto.randomUUID();
    pendingImports.clear();
    pendingImports.set(token, prepared);
    return { canceled: false, preview: { token, fileName: basename(filePath), ...prepared.preview } };
  });

  ipcMain.handle("workspace:apply", async (_event, input: { token: string; mode: WorkspaceImportMode }) => {
    if (!input || !["merge", "replace"].includes(input.mode)) throw new Error("请选择有效的导入方式");
    const prepared = pendingImports.get(input.token);
    if (!prepared) throw new Error("导入数据已失效，请重新选择数据包");
    pendingImports.delete(input.token);
    return { snapshot: await database.importWorkspace(prepared.data, input.mode) };
  });

  ipcMain.handle("workbook:export", async () => {
    const result = await dialog.showSaveDialog(mainWindow!, {
      title: "导出企业协作 Excel",
      defaultPath: `需求路标协作-${new Date().toISOString().slice(0, 10)}.xlsx`,
      filters: [{ name: "Excel 工作簿", extensions: ["xlsx"] }],
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    const outputPath = result.filePath.toLowerCase().endsWith(".xlsx") ? result.filePath : `${result.filePath}.xlsx`;
    await writeFile(outputPath, await createWorkspaceWorkbook(database.getWorkspaceData(), app.getVersion()));
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

  ipcMain.handle("workbook:apply", async (_event, input: { token: string; conflictMode: WorkspaceWorkbookConflictMode }) => {
    if (!input || !["local-wins", "excel-wins"].includes(input.conflictMode)) throw new Error("请选择有效的冲突处理方式");
    const prepared = pendingWorkbookImports.get(input.token);
    if (!prepared) throw new Error("Excel 导入数据已失效，请重新选择文件");
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

  ipcMain.handle("presentation:export", async (_event, input: HalfYearRange) => {
    const snapshot = database.getSnapshot();
    const plan = buildPptExportPlan(snapshot, input.start, input.end);
    const result = await dialog.showSaveDialog(mainWindow!, {
      title: "导出需求路标PPT",
      defaultPath: `需求路标-${input.start}-${input.end}-${new Date().toISOString().slice(0, 10)}.pptx`,
      filters: [{ name: "PowerPoint演示文稿", extensions: ["pptx"] }],
    });
    if (result.canceled || !result.filePath) return { canceled: true };
    const outputPath = result.filePath.toLowerCase().endsWith(".pptx") ? result.filePath : `${result.filePath}.pptx`;
    await writeFile(outputPath, await createRoadmapPresentation(snapshot, input.start, input.end));
    return { canceled: false, path: outputPath, slideCount: plan.slideCount };
  });
}

app.whenReady().then(async () => {
  nativeTheme.themeSource = "light";
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
