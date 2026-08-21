import { contextBridge, ipcRenderer, type IpcRendererEvent } from "electron";
import type {
  ElectronApi,
  SaveDictionaryInput,
  SaveGroupOverrideInput,
  SaveRequirementInput,
  WorkspaceImportMode,
  WorkspaceWorkbookConflictMode,
} from "../shared/types.js";

const api: ElectronApi = {
  getAppInfo: () => ipcRenderer.invoke("app:info"),
  openUpdatePage: () => ipcRenderer.invoke("app:update:open"),
  getSnapshot: () => ipcRenderer.invoke("snapshot:get"),
  getWindowFullscreen: () => ipcRenderer.invoke("window:fullscreen:get"),
  setWindowFullscreen: (enabled: boolean) => ipcRenderer.invoke("window:fullscreen:set", enabled),
  onWindowFullscreenChange: (listener: (enabled: boolean) => void) => {
    const subscription = (_event: IpcRendererEvent, enabled: boolean) => listener(enabled);
    ipcRenderer.on("window:fullscreen:changed", subscription);
    return () => ipcRenderer.removeListener("window:fullscreen:changed", subscription);
  },
  saveRequirement: (input: SaveRequirementInput) => ipcRenderer.invoke("requirements:save", input),
  deleteRequirement: (id: string) => ipcRenderer.invoke("requirements:delete", id),
  saveDomain: (input: SaveDictionaryInput) => ipcRenderer.invoke("domains:save", input),
  deleteDomain: (id: string) => ipcRenderer.invoke("domains:delete", id),
  saveProduct: (input: SaveDictionaryInput) => ipcRenderer.invoke("products:save", input),
  deleteProduct: (id: string) => ipcRenderer.invoke("products:delete", id),
  saveGroupOverride: (input: SaveGroupOverrideInput) => ipcRenderer.invoke("groups:save", input),
  exportWorkspacePackage: () => ipcRenderer.invoke("workspace:export"),
  inspectWorkspacePackage: () => ipcRenderer.invoke("workspace:inspect"),
  applyWorkspacePackage: (input: { token: string; mode: WorkspaceImportMode }) => ipcRenderer.invoke("workspace:apply", input),
  exportWorkspaceWorkbook: () => ipcRenderer.invoke("workbook:export"),
  inspectWorkspaceWorkbook: () => ipcRenderer.invoke("workbook:inspect"),
  applyWorkspaceWorkbook: (input: { token: string; conflictMode: WorkspaceWorkbookConflictMode }) => ipcRenderer.invoke("workbook:apply", input),
  exportTemplateDraft: () => ipcRenderer.invoke("template:draft-export"),
  exportRoadmapPresentation: (input) => ipcRenderer.invoke("presentation:export", input),
};

contextBridge.exposeInMainWorld("roadmapApi", api);
