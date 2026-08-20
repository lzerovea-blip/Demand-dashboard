import { contextBridge, ipcRenderer } from "electron";
import type {
  ElectronApi,
  SaveDictionaryInput,
  SaveGroupOverrideInput,
  SaveRequirementInput,
  WorkspaceImportMode,
  WorkspaceWorkbookConflictMode,
} from "../shared/types.js";

const api: ElectronApi = {
  getSnapshot: () => ipcRenderer.invoke("snapshot:get"),
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
