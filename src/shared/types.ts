export const SOURCES = [
  "运动基础",
  "运动进阶",
  "运动高阶",
  "健康基础",
  "健康进阶",
  "健康高阶",
  "行业",
  "全场景",
  "health平台",
  "AI",
  "医疗送检",
] as const;

export const CATEGORIES = ["体验优化", "产品专属"] as const;
/** 仅用于兼容旧版“海外研究”数据包，当前界面不再提供该字段。 */
export const OVERSEAS_REGIONS = ["欧州", "亚非拉", "欧亚"] as const;
export const ROADMAP_TRACKS = ["运动", "健康", "行业", "全场景", "health平台", "AI", "医疗送检"] as const;

export type RequirementSource = (typeof SOURCES)[number];
export type RequirementCategory = (typeof CATEGORIES)[number];
export type OverseasRegion = (typeof OVERSEAS_REGIONS)[number];
export type Track = (typeof ROADMAP_TRACKS)[number];
export type Level = "基础" | "进阶" | "高阶";
export type DomainLevel = "L0" | "L1";

export const MAX_REQUIREMENT_IMAGES = 5;
export const MAX_REQUIREMENT_IMAGE_BYTES = 5 * 1024 * 1024;
export const REQUIREMENT_IMAGE_MIME_TYPES = ["image/png", "image/jpeg", "image/webp", "image/gif"] as const;

export interface RequirementImage {
  id: string;
  name: string;
  mimeType: (typeof REQUIREMENT_IMAGE_MIME_TYPES)[number];
  sizeBytes: number;
  dataUrl: string;
}

export interface Requirement {
  id: string;
  title: string;
  description: string;
  images: RequirementImage[];
  domainL0Id: string;
  domainId: string;
  source: RequirementSource;
  overseasRegions: OverseasRegion[];
  category: RequirementCategory;
  targetMonth: string;
  productIds: string[];
  deviceWorkloadPm: number;
  appWorkloadPm: number;
  cloudWorkloadPm: number;
  unallocatedWorkloadPm: number;
  workloadPm: number;
  createdAt: string;
  updatedAt: string;
}

export interface DictionaryItem {
  id: string;
  name: string;
  sortOrder: number;
  active: boolean;
  /** 仅领域字典使用；历史领域默认按 L1 处理。 */
  level?: DomainLevel;
  /** 仅领域 L1 使用，指向其上游领域 L0；历史未归属的 L1 可暂时为空。 */
  parentId?: string;
}

export interface GroupOverride {
  groupKey: string;
  cardTitle: string;
  cardSummary: string;
  updatedAt: string;
}

export interface TemplateProfile {
  id: string;
  name: string;
  storedPath: string;
  originalPath: string;
  sha256: string;
  slideCount: number;
  importedAt: string;
  status: "needs_mapping" | "ready" | "invalid";
  inventory: TemplateInventory;
  mapping?: TemplateMapping;
}

export interface TemplateShapeInventory {
  slideNumber: number;
  shapeId: string;
  shapeName: string;
  text: string;
  placeholderType?: string;
}

export interface TemplateInventory {
  slideCount: number;
  slideSize?: { width: number; height: number };
  shapes: TemplateShapeInventory[];
}

export interface TemplateMapping {
  summarySlide: number;
  sportsRoadmapSlide: number;
  healthRoadmapSlide: number;
  detailSlide: number;
  roles: Record<string, { slideNumber: number; shapeId: string }>;
}

export interface AppSnapshot {
  requirements: Requirement[];
  domains: DictionaryItem[];
  products: DictionaryItem[];
  groupOverrides: GroupOverride[];
  templates: TemplateProfile[];
}

export interface SaveRequirementInput {
  id?: string;
  title: string;
  description: string;
  images: RequirementImage[];
  domainL0Id: string;
  domainId: string;
  source: RequirementSource;
  overseasRegions: OverseasRegion[];
  category: RequirementCategory;
  targetMonth: string;
  productIds: string[];
  deviceWorkloadPm: number;
  appWorkloadPm: number;
  cloudWorkloadPm: number;
}

export interface SaveDictionaryInput {
  id?: string;
  name: string;
  sortOrder?: number;
  active?: boolean;
  level?: DomainLevel;
  parentId?: string;
}

export interface SaveGroupOverrideInput {
  groupKey: string;
  cardTitle: string;
  cardSummary: string;
}

export interface HalfYearRange {
  start: string;
  end: string;
}

export interface ExportOptions extends HalfYearRange {
  templateId: string;
  outputPath?: string;
}

export interface BackupEnvelope {
  schemaVersion: 1;
  exportedAt: string;
  data: WorkspaceData;
}

export type WorkspaceData = Omit<AppSnapshot, "templates">;
export type WorkspaceImportMode = "merge" | "replace";

export interface WorkspacePackageCounts {
  requirements: number;
  images: number;
  domains: number;
  products: number;
  groupOverrides: number;
}

export interface WorkspaceImportPreview {
  token: string;
  fileName: string;
  sourceFormat: "roadmap" | "legacy-json";
  exportedAt: string;
  templateVersion: string;
  counts: WorkspacePackageCounts;
}

export type WorkspaceWorkbookConflictMode = "local-wins" | "excel-wins";

export interface WorkspaceWorkbookIssue {
  sheet: string;
  cell: string;
  message: string;
}

export interface WorkspaceWorkbookImportCounts {
  added: number;
  updated: number;
  deleted: number;
  unchanged: number;
  conflicts: number;
  domainsChanged: number;
  productsChanged: number;
  groupOverridesChanged: number;
}

export interface WorkspaceWorkbookImportPreview {
  token: string;
  fileName: string;
  formatVersion: string;
  exportedAt: string;
  counts: WorkspaceWorkbookImportCounts;
  errors: WorkspaceWorkbookIssue[];
  conflicts: WorkspaceWorkbookIssue[];
}

export interface AppInfo {
  name: string;
  version: string;
  releasedAt: string;
  author: string;
  updateUrl: string;
}

export interface ElectronApi {
  getAppInfo(): Promise<AppInfo>;
  openUpdatePage(): Promise<void>;
  getSnapshot(): Promise<AppSnapshot>;
  getWindowFullscreen(): Promise<boolean>;
  setWindowFullscreen(enabled: boolean): Promise<void>;
  onWindowFullscreenChange(listener: (enabled: boolean) => void): () => void;
  saveRequirement(input: SaveRequirementInput): Promise<AppSnapshot>;
  deleteRequirement(id: string): Promise<AppSnapshot>;
  saveDomain(input: SaveDictionaryInput): Promise<AppSnapshot>;
  deleteDomain(id: string): Promise<AppSnapshot>;
  saveProduct(input: SaveDictionaryInput): Promise<AppSnapshot>;
  deleteProduct(id: string): Promise<AppSnapshot>;
  saveGroupOverride(input: SaveGroupOverrideInput): Promise<AppSnapshot>;
  exportWorkspacePackage(): Promise<{ canceled: boolean; path?: string }>;
  inspectWorkspacePackage(): Promise<{ canceled: boolean; preview?: WorkspaceImportPreview }>;
  applyWorkspacePackage(input: { token: string; mode: WorkspaceImportMode }): Promise<{ snapshot: AppSnapshot }>;
  exportWorkspaceWorkbook(): Promise<{ canceled: boolean; path?: string }>;
  inspectWorkspaceWorkbook(): Promise<{ canceled: boolean; preview?: WorkspaceWorkbookImportPreview }>;
  applyWorkspaceWorkbook(input: { token: string; conflictMode: WorkspaceWorkbookConflictMode }): Promise<{ snapshot: AppSnapshot }>;
  exportTemplateDraft(): Promise<{ canceled: boolean; path?: string }>;
  exportRoadmapPresentation(input: HalfYearRange): Promise<{ canceled: boolean; path?: string; slideCount?: number }>;
}

declare global {
  interface Window {
    roadmapApi: ElectronApi;
  }
}
