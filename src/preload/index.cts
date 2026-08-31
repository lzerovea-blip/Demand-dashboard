import { contextBridge, ipcRenderer, type IpcRendererEvent } from "electron";
import type {
  ElectronApi,
  HalfYearRange,
  SaveDictionaryInput,
  SaveGroupOrderInput,
  SaveGroupOverrideInput,
  SaveMemberInput,
  SavePptAccessPolicyInput,
  SaveRequirementInput,
  WorkspaceImportMode,
  WorkspaceWorkbookConflictMode,
} from "../shared/types.js";

const api: ElectronApi = {
  getAppInfo: () => ipcRenderer.invoke("app:info"),
  getSnapshot: () => ipcRenderer.invoke("snapshot:get"),
  getWindowFullscreen: () => ipcRenderer.invoke("window:fullscreen:get"),
  setWindowFullscreen: (enabled: boolean) => ipcRenderer.invoke("window:fullscreen:set", enabled),
  onWindowFullscreenChange: (listener: (enabled: boolean) => void) => {
    const subscription = (_event: IpcRendererEvent, enabled: boolean) => listener(enabled);
    ipcRenderer.on("window:fullscreen:changed", subscription);
    return () => ipcRenderer.removeListener("window:fullscreen:changed", subscription);
  },
  saveRequirement: (input: SaveRequirementInput) => ipcRenderer.invoke("requirements:save", input),
  duplicateRequirement: (id: string) => ipcRenderer.invoke("requirements:duplicate", id),
  deleteRequirement: (id: string) => ipcRenderer.invoke("requirements:delete", id),
  saveDomain: (input: SaveDictionaryInput) => ipcRenderer.invoke("domains:save", input),
  deleteDomain: (id: string) => ipcRenderer.invoke("domains:delete", id),
  saveProduct: (input: SaveDictionaryInput) => ipcRenderer.invoke("products:save", input),
  deleteProduct: (id: string) => ipcRenderer.invoke("products:delete", id),
  saveMember: (input: SaveMemberInput) => ipcRenderer.invoke("members:save", input),
  deleteMember: (id: string) => ipcRenderer.invoke("members:delete", id),
  savePptAccessPolicy: (input: SavePptAccessPolicyInput) => ipcRenderer.invoke("ppt-access:save", input),
  checkIrmEnvironment: () => ipcRenderer.invoke("ppt-access:check-environment"),
  getPptDiagnosticLog: () => ipcRenderer.invoke("ppt-access:diagnostic-log"),
  saveGroupOverride: (input: SaveGroupOverrideInput) => ipcRenderer.invoke("groups:save", input),
  saveGroupOrder: (input: SaveGroupOrderInput) => ipcRenderer.invoke("groups:reorder", input),
  exportWorkspacePackage: () => ipcRenderer.invoke("workspace:export"),
  inspectWorkspacePackage: () => ipcRenderer.invoke("workspace:inspect"),
  applyWorkspacePackage: (input: { token: string; mode: WorkspaceImportMode; differenceDigest?: string }) => ipcRenderer.invoke("workspace:apply", input),
  exportWorkspaceWorkbook: (input: HalfYearRange) => ipcRenderer.invoke("workbook:export", input),
  inspectWorkspaceWorkbook: () => ipcRenderer.invoke("workbook:inspect"),
  applyWorkspaceWorkbook: (input: { token: string; conflictMode: WorkspaceWorkbookConflictMode; differenceDigest: string }) => ipcRenderer.invoke("workbook:apply", input),
  exportTemplateDraft: () => ipcRenderer.invoke("template:draft-export"),
  exportRoadmapPresentation: (input) => ipcRenderer.invoke("presentation:export", input),
};

contextBridge.exposeInMainWorld("roadmapApi", api);
