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
  "AI运动",
  "AI健康",
  "AI综合",
  "NMPA认证",
  "CE认证",
] as const;

export const CATEGORIES = ["体验优化", "平台演进", "产品专属"] as const;
export const REQUIREMENT_SCOPES = ["仅国内", "仅海外", "全球"] as const;
/** 仅用于兼容旧版“海外研究”数据包，当前界面不再提供该字段。 */
export const OVERSEAS_REGIONS = ["欧州", "亚非拉", "欧亚"] as const;
export const ROADMAP_TRACKS = ["运动", "健康", "行业", "全场景", "health平台", "AI", "医疗认证"] as const;
export const MEMBER_ROLES = ["规划", "产品经理", "SE", "PM", "开发", "算法", "UX"] as const;
export const TIERING_SUPPORT_STATUSES = ["支持", "不支持", "未判断"] as const;
export const TIERING_AUDIT_STATUSES = ["待内审", "已公开核验", "已审核", "需复核"] as const;

export type RequirementSource = (typeof SOURCES)[number];
export type RequirementCategory = (typeof CATEGORIES)[number];
export const CATEGORY_META: Record<RequirementCategory, {
  color: string;
  background: string;
  className: string;
  requiresProduct: boolean;
}> = {
  体验优化: { color: "#7357d9", background: "#efecff", className: "experience", requiresProduct: false },
  平台演进: { color: "#c96a08", background: "#fff0d8", className: "platform", requiresProduct: false },
  产品专属: { color: "#17202a", background: "#e9edef", className: "exclusive", requiresProduct: true },
};
export type RequirementScope = (typeof REQUIREMENT_SCOPES)[number];
export type OverseasRegion = (typeof OVERSEAS_REGIONS)[number];
export type Track = (typeof ROADMAP_TRACKS)[number];
export type MemberRole = (typeof MEMBER_ROLES)[number];
export type TieringSupportStatus = (typeof TIERING_SUPPORT_STATUSES)[number];
export type TieringAuditStatus = (typeof TIERING_AUDIT_STATUSES)[number];
export type PptPermissionLevel = "read" | "change";
export const HISTORICAL_LAUNCH_PRECISIONS = ["month", "half-year", "year"] as const;
export const HISTORICAL_LAUNCH_VERIFICATION_STATUSES = ["confirmed", "pending"] as const;
export const HISTORICAL_LAUNCH_CONFIDENCE_LEVELS = ["A", "B", "C"] as const;
export type HistoricalLaunchPrecision = (typeof HISTORICAL_LAUNCH_PRECISIONS)[number];
export type HistoricalLaunchVerificationStatus = (typeof HISTORICAL_LAUNCH_VERIFICATION_STATUSES)[number];
export type HistoricalLaunchConfidence = (typeof HISTORICAL_LAUNCH_CONFIDENCE_LEVELS)[number];

export interface HistoricalLaunchMetadata {
  suggestedTargetMonth: string;
  precision: HistoricalLaunchPrecision;
  verificationStatus: HistoricalLaunchVerificationStatus;
  confidence: HistoricalLaunchConfidence;
  evidenceUrls: string[];
  note: string;
}
export type PptExportSectionId = "summary" | Track;
export interface PptExportRequest extends HalfYearRange {
  materialName: string;
  sections: PptExportSectionId[];
}
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

export interface MedicalSubmissionMilestones {
  registrationStartMonth?: string;
  clinicalStartMonth?: string;
  clinicalCompleteMonth?: string;
  certificateMonth?: string;
}

export interface Requirement {
  algorithmMemberId?: string;
  planningMemberId?: string;
  productManagerMemberId?: string;
  seMemberId?: string;
  id: string;
  title: string;
  description: string;
  images: RequirementImage[];
  domainL0Id: string;
  domainId: string;
  source: RequirementSource;
  scope: RequirementScope;
  allowLatestVersionUpgrade: boolean;
  overseasRegions: OverseasRegion[];
  category: RequirementCategory;
  /** YYYY-MM; empty only for NMPA/CE requirements with no submission version yet. */
  targetMonth: string;
  medicalSubmissionMilestones?: MedicalSubmissionMilestones;
  /** Application-only medical lane order, independent of dates and L1. */
  medicalRoadmapSortOrder?: number;
  productIds: string[];
  deviceWorkloadPm: number;
  appWorkloadPm: number;
  cloudWorkloadPm: number;
  /** Absent only in legacy input; normalized workspace records use 0. */
  researchAppWorkloadPm?: number;
  algorithmWorkloadPm?: number;
  unallocatedWorkloadPm: number;
  workloadPm: number;
  createdAt: string;
  updatedAt: string;
  /** 指向分层分级中的稳定需求行；普通需求与历史需求均使用。 */
  tieringFeatureId?: string;
  /** 历史需求的中国大陆首次公开可用时间证据；普通需求不设置。 */
  historicalLaunch?: HistoricalLaunchMetadata;
}

export interface RequirementTombstone {
  id: string;
  title: string;
  deletedAt: string;
}

export interface TieringFeature {
  id: string;
  l0: string;
  l1: string;
  requirementTitle: string;
  requirementId?: string;
  historicalLaunch?: HistoricalLaunchMetadata;
  sortOrder: number;
  createdAt: string;
  updatedAt: string;
}

export interface TieringProductSeries {
  id: string;
  name: string;
  boundary: string;
  representativeProducts: string;
  managementEntry: string;
  sourceIds: string[];
  lifecycle: string;
  note: string;
  sortOrder: number;
  updatedAt: string;
}

export interface TieringSupportRecord {
  id: string;
  featureId: string;
  seriesId: string;
  status: TieringSupportStatus;
  supportedModels: string;
  excludedModels: string;
  phoneOsMarket: string;
  minHealthAppVersion: string;
  minDeviceVersion: string;
  prerequisites: string;
  sourceIds: string[];
  sourceUrl: string;
  verifiedAt: string;
  evidenceLevel: string;
  auditStatus: TieringAuditStatus;
  note: string;
  updatedAt: string;
}

export interface TieringEvidenceSource {
  id: string;
  sourceType: string;
  title: string;
  url: string;
  scope: string;
  market: string;
  versionCondition: string;
  collectedAt: string;
  verifiedAt: string;
  evidenceLevel: string;
  maintenanceStatus: string;
  owner: string;
  note: string;
  updatedAt: string;
}

export interface TieringVerificationIssue {
  id: string;
  featureId: string;
  seriesIds: string[];
  issueType: string;
  description: string;
  priority: string;
  status: string;
  sourceIds: string[];
  owner: string;
  createdAt: string;
  targetDate: string;
  updatedAt: string;
  resolution: string;
}

export interface TieringChangeLog {
  id: string;
  changedAt: string;
  triggerType: string;
  featureId: string;
  seriesId: string;
  before: string;
  after: string;
  reason: string;
  sourceIds: string[];
  reviewer: string;
  recordStatus: string;
  note: string;
  updatedAt: string;
}

export interface TieringData {
  features: TieringFeature[];
  productSeries: TieringProductSeries[];
  supportRecords: TieringSupportRecord[];
  evidenceSources: TieringEvidenceSource[];
  verificationIssues: TieringVerificationIssue[];
  changeLogs: TieringChangeLog[];
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
  sortOrder?: number;
  updatedAt: string;
}

export interface Member {
  id: string;
  name: string;
  employeeNumber: string;
  microsoftAccount: string;
  role?: MemberRole;
  domainL1Ids: string[];
  sortOrder: number;
  createdAt: string;
  updatedAt: string;
}

export interface PptWhitelistEntry {
  memberId: string;
  permission: PptPermissionLevel;
}

export interface PptAccessPolicy {
  entries: PptWhitelistEntry[];
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
  members: Member[];
  pptAccessPolicy: PptAccessPolicy;
  tiering?: TieringData;
  templates: TemplateProfile[];
}

export interface SaveRequirementInput {
  algorithmMemberId?: string;
  planningMemberId: string;
  productManagerMemberId: string;
  seMemberId: string;
  id?: string;
  title: string;
  description: string;
  images: RequirementImage[];
  domainL0Id: string;
  domainId: string;
  source: RequirementSource;
  scope: RequirementScope;
  allowLatestVersionUpgrade: boolean;
  overseasRegions: OverseasRegion[];
  category: RequirementCategory;
  /** YYYY-MM; empty only for NMPA/CE requirements with no submission version yet. */
  targetMonth: string;
  medicalSubmissionMilestones?: MedicalSubmissionMilestones;
  productIds: string[];
  deviceWorkloadPm: number;
  appWorkloadPm: number;
  cloudWorkloadPm: number;
  researchAppWorkloadPm: number;
  algorithmWorkloadPm: number;
  tieringFeatureId?: string;
}

export interface SaveTieringSupportInput {
  recordId: string;
  status: TieringSupportStatus;
  reason?: string;
}

export interface SaveTieringEvidenceSourceInput {
  source: TieringEvidenceSource;
}

export interface UpdateTieringIssueInput {
  issueId: string;
  status: string;
  owner: string;
  targetDate: string;
  resolution: string;
}

export interface UpdateTieringChangeLogInput {
  logId: string;
  reason: string;
  reviewer: string;
  recordStatus: string;
  note: string;
}

export interface RendererDiagnosticInput {
  event: "renderer-ready" | "renderer-error" | "renderer-unhandled-rejection" | "react-error";
  detail?: string;
}

export interface DiagnosticInfo {
  directory: string;
  currentLogPath: string;
}

export interface SaveDictionaryInput {
  id?: string;
  name: string;
  sortOrder?: number;
  active?: boolean;
  level?: DomainLevel;
  parentId?: string;
}

export interface SaveMemberInput {
  id?: string;
  name: string;
  employeeNumber: string;
  microsoftAccount: string;
  role?: MemberRole;
  domainL1Ids: string[];
  sortOrder?: number;
}

export interface SavePptAccessPolicyInput {
  entries: PptWhitelistEntry[];
}

export interface IrmEnvironmentStatus {
  supported: boolean;
  platform: string;
  powerPointInstalled: boolean;
  irmReady: boolean;
  officeIdentity?: string;
  preferredDocumentOwner?: string;
  message: string;
}

export interface PptExportResult {
  canceled: boolean;
  path?: string;
  slideCount?: number;
  effectiveDocumentOwner?: string;
  fallbackUsed?: boolean;
}

export interface SaveGroupOverrideInput {
  groupKey: string;
  cardTitle: string;
  cardSummary: string;
}

export interface SaveGroupOrderInput {
  source: RequirementSource;
  targetMonth: string;
  groupKeys: string[];
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

export type WorkspaceData = Omit<AppSnapshot, "templates" | "tiering"> & {
  tiering?: TieringData;
  requirementTombstones?: RequirementTombstone[];
  requirementDeletionTrackingStartedAt?: string;
};
export type WorkspaceImportMode = "merge" | "replace";
export type WorkspacePackageKind = "workspace" | "tiering-seed";
export type LegacyRequirementDecision = "merge" | "keep-deleted";

export interface WorkspacePackageCounts {
  requirements: number;
  requirementTombstones: number;
  images: number;
  domains: number;
  products: number;
  groupOverrides: number;
  members: number;
  whitelistEntries: number;
  tieringFeatures: number;
  tieringProductSeries: number;
  tieringSupportRecords: number;
  tieringEvidenceSources: number;
  tieringVerificationIssues: number;
  tieringChangeLogs: number;
}

export type WorkspaceDifferenceChange = "新增" | "修改" | "删除" | "恢复" | "忽略" | "待确认";

export interface WorkspaceFieldDifference {
  field: string;
  currentValue: string;
  incomingValue: string;
}

export interface WorkspaceImportDifference {
  entity: "需求" | "领域" | "产品" | "路标卡片" | "成员" | "PPT白名单" | "分层功能" | "产品系列" | "系列支持" | "证据来源" | "待核验项" | "变更记录";
  key: string;
  title: string;
  change: WorkspaceDifferenceChange;
  fields: WorkspaceFieldDifference[];
  decisionRequired?: boolean;
}

export interface WorkspaceImportPreview {
  token: string;
  fileName: string;
  sourceFormat: "roadmap" | "legacy-json";
  packageKind: WorkspacePackageKind;
  exportedAt: string;
  templateVersion: string;
  counts: WorkspacePackageCounts;
  differences: WorkspaceImportDifference[];
  ambiguousRequirementIds: string[];
  differenceDigest: string;
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
  differences: WorkspaceImportDifference[];
  differenceDigest: string;
  fieldChanges?: WorkspaceWorkbookFieldChange[];
}

export interface WorkspaceWorkbookFieldChange {
  id: string;
  requirementId: string;
  row: number;
  title: string;
  kind: "field" | "series" | "delete";
  field: string;
  baselineValue: string;
  currentValue: string;
  excelValue: string;
  conflict: boolean;
}

export interface WorkspaceWorkbookFieldSelection {
  selectedChangeIds: string[];
  conflictDecisions: Record<string, "local" | "excel">;
}

export interface AppInfo {
  name: string;
  version: string;
  releasedAt: string;
  author: string;
}

export interface ElectronApi {
  getAppInfo(): Promise<AppInfo>;
  getSnapshot(): Promise<AppSnapshot>;
  getWindowFullscreen(): Promise<boolean>;
  setWindowFullscreen(enabled: boolean): Promise<void>;
  onWindowFullscreenChange(listener: (enabled: boolean) => void): () => void;
  saveRequirement(input: SaveRequirementInput): Promise<AppSnapshot>;
  duplicateRequirement(id: string): Promise<AppSnapshot>;
  deleteRequirement(id: string): Promise<AppSnapshot>;
  saveDomain(input: SaveDictionaryInput): Promise<AppSnapshot>;
  deleteDomain(id: string): Promise<AppSnapshot>;
  saveProduct(input: SaveDictionaryInput): Promise<AppSnapshot>;
  deleteProduct(id: string): Promise<AppSnapshot>;
  saveMember(input: SaveMemberInput): Promise<AppSnapshot>;
  deleteMember(id: string): Promise<AppSnapshot>;
  savePptAccessPolicy(input: SavePptAccessPolicyInput): Promise<AppSnapshot>;
  checkIrmEnvironment(): Promise<IrmEnvironmentStatus>;
  saveGroupOverride(input: SaveGroupOverrideInput): Promise<AppSnapshot>;
  saveGroupOrder(input: SaveGroupOrderInput): Promise<AppSnapshot>;
  saveMedicalRoadmapOrder(input: { requirementIds: string[] }): Promise<AppSnapshot>;
  saveTieringSupport(input: SaveTieringSupportInput): Promise<AppSnapshot>;
  saveTieringEvidenceSource(input: SaveTieringEvidenceSourceInput): Promise<AppSnapshot>;
  updateTieringIssue(input: UpdateTieringIssueInput): Promise<AppSnapshot>;
  updateTieringChangeLog(input: UpdateTieringChangeLogInput): Promise<AppSnapshot>;
  deleteTieringChangeLog(logId: string): Promise<AppSnapshot>;
  regenerateTieringRequirement(featureId: string): Promise<AppSnapshot>;
  reportRendererDiagnostic(input: RendererDiagnosticInput): Promise<void>;
  getDiagnosticInfo(): Promise<DiagnosticInfo>;
  openDiagnosticDirectory(): Promise<DiagnosticInfo>;
  exportWorkspacePackage(): Promise<{ canceled: boolean; path?: string }>;
  inspectWorkspacePackage(): Promise<{ canceled: boolean; preview?: WorkspaceImportPreview }>;
  applyWorkspacePackage(input: { token: string; mode: WorkspaceImportMode; differenceDigest?: string; legacyRequirementDecisions?: Record<string, LegacyRequirementDecision> }): Promise<{ snapshot: AppSnapshot }>;
  exportWorkspaceWorkbook(input: HalfYearRange): Promise<{ canceled: boolean; path?: string }>;
  exportTieringWorkbook(input: HalfYearRange): Promise<{ canceled: boolean; path?: string }>;
  inspectWorkspaceWorkbook(): Promise<{ canceled: boolean; preview?: WorkspaceWorkbookImportPreview }>;
  applyWorkspaceWorkbook(input: { token: string; conflictMode: WorkspaceWorkbookConflictMode; differenceDigest: string; fieldSelection?: WorkspaceWorkbookFieldSelection }): Promise<{ snapshot: AppSnapshot }>;
  exportTemplateDraft(): Promise<{ canceled: boolean; path?: string }>;
  exportRoadmapPresentation(input: PptExportRequest): Promise<PptExportResult>;
}

declare global {
  interface Window {
    roadmapApi: ElectronApi;
  }
}
