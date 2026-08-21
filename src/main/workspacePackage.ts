import { createHash } from "node:crypto";
import JSZip from "jszip";
import type {
  BackupEnvelope,
  Requirement,
  RequirementImage,
  WorkspaceData,
  WorkspaceImportPreview,
  WorkspacePackageCounts,
} from "../shared/types.js";
import { CATEGORIES, MAX_REQUIREMENT_IMAGES, MAX_REQUIREMENT_IMAGE_BYTES, REQUIREMENT_IMAGE_MIME_TYPES, SOURCES } from "../shared/types.js";
import { normalizeLegacyGroupOverride, normalizeLegacySource, normalizeRequirement } from "../shared/requirements.js";
import { workspaceCounts } from "../shared/workspace.js";

const PACKAGE_FORMAT = "health-roadmap-workspace";
const PACKAGE_VERSION = 1;
const TEMPLATE_VERSION = "v0.1";
const DATA_PATH = "data.json";

interface PackageFileRecord {
  sha256: string;
  sizeBytes: number;
}

interface WorkspacePackageManifest {
  format: typeof PACKAGE_FORMAT;
  formatVersion: typeof PACKAGE_VERSION;
  appVersion: string;
  templateVersion: string;
  exportedAt: string;
  dataPath: typeof DATA_PATH;
  counts: WorkspacePackageCounts;
  files: Record<string, PackageFileRecord>;
}

interface PackedRequirementImage extends Omit<RequirementImage, "dataUrl"> {
  path: string;
}

interface PackedRequirement extends Omit<Requirement, "images"> {
  images: PackedRequirementImage[];
}

interface PackedWorkspaceData extends Omit<WorkspaceData, "requirements"> {
  requirements: PackedRequirement[];
}

export interface PreparedWorkspaceImport {
  data: WorkspaceData;
  preview: Omit<WorkspaceImportPreview, "token" | "fileName">;
}

export async function createWorkspacePackage(data: WorkspaceData, appVersion: string, exportedAt = new Date().toISOString()): Promise<Buffer> {
  validateWorkspaceData(data);
  const zip = new JSZip();
  const files: Record<string, PackageFileRecord> = {};
  const usedPaths = new Set<string>();
  const packedRequirements: PackedRequirement[] = [];

  for (const requirement of data.requirements) {
    const images: PackedRequirementImage[] = [];
    for (const [index, image] of requirement.images.entries()) {
      const bytes = decodeDataUrl(image);
      const imagePath = uniqueImagePath(requirement.id, image.id, image.mimeType, index, usedPaths);
      usedPaths.add(imagePath);
      zip.file(imagePath, bytes);
      files[imagePath] = fileRecord(bytes);
      images.push({
        id: image.id,
        name: image.name,
        mimeType: image.mimeType,
        sizeBytes: bytes.length,
        path: imagePath,
      });
    }
    packedRequirements.push({ ...requirement, images });
  }

  const packedData: PackedWorkspaceData = {
    requirements: packedRequirements,
    domains: normalizeDomains(data.domains),
    products: data.products,
    groupOverrides: data.groupOverrides,
  };
  const dataBytes = Buffer.from(`${JSON.stringify(packedData, null, 2)}\n`, "utf8");
  zip.file(DATA_PATH, dataBytes);
  files[DATA_PATH] = fileRecord(dataBytes);

  const manifest: WorkspacePackageManifest = {
    format: PACKAGE_FORMAT,
    formatVersion: PACKAGE_VERSION,
    appVersion,
    templateVersion: TEMPLATE_VERSION,
    exportedAt,
    dataPath: DATA_PATH,
    counts: workspaceCounts(data),
    files,
  };
  zip.file("manifest.json", `${JSON.stringify(manifest, null, 2)}\n`);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

export async function inspectWorkspacePackage(bytes: Buffer, fileName: string): Promise<PreparedWorkspaceImport> {
  if (fileName.toLowerCase().endsWith(".json") || firstNonWhitespace(bytes) === "{") {
    return inspectLegacyBackup(bytes);
  }

  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const manifestFile = zip.file("manifest.json");
  if (!manifestFile) throw new Error("数据包缺少 manifest.json");
  const manifest = JSON.parse(await manifestFile.async("string")) as WorkspacePackageManifest;
  validateManifest(manifest);

  const verifiedFiles = new Map<string, Buffer>();
  for (const [filePath, expected] of Object.entries(manifest.files)) {
    validateArchivePath(filePath);
    const file = zip.file(filePath);
    if (!file) throw new Error(`数据包缺少文件：${filePath}`);
    const content = Buffer.from(await file.async("uint8array"));
    if (content.length !== expected.sizeBytes || sha256(content) !== expected.sha256) {
      throw new Error(`文件校验失败：${filePath}`);
    }
    verifiedFiles.set(filePath, content);
  }

  const dataBytes = verifiedFiles.get(manifest.dataPath);
  if (!dataBytes) throw new Error("数据包缺少结构化数据");
  const packedData = JSON.parse(dataBytes.toString("utf8")) as PackedWorkspaceData;
  const data: WorkspaceData = {
    domains: normalizeDomains(packedData.domains),
    products: packedData.products,
    groupOverrides: packedData.groupOverrides.map(normalizeLegacyGroupOverride),
    requirements: packedData.requirements.map((requirement) => ({
      ...normalizeRequirement({ ...requirement, images: [] }),
      images: requirement.images.map((image) => {
        const content = verifiedFiles.get(image.path);
        if (!content) throw new Error(`需求图片缺失：${image.name}`);
        if (content.length !== image.sizeBytes) throw new Error(`需求图片大小不一致：${image.name}`);
        return {
          id: image.id,
          name: image.name,
          mimeType: image.mimeType,
          sizeBytes: content.length,
          dataUrl: `data:${image.mimeType};base64,${content.toString("base64")}`,
        };
      }),
    })),
  };
  validateWorkspaceData(data);
  const actualCounts = workspaceCounts(data);
  if (!sameCounts(actualCounts, manifest.counts)) throw new Error("数据包数量清单与实际内容不一致");

  return {
    data,
    preview: {
      sourceFormat: "roadmap",
      exportedAt: manifest.exportedAt,
      templateVersion: manifest.templateVersion,
      counts: actualCounts,
    },
  };
}

function inspectLegacyBackup(bytes: Buffer): PreparedWorkspaceImport {
  const backup = JSON.parse(bytes.toString("utf8")) as BackupEnvelope;
  if (backup.schemaVersion !== 1 || !backup.data) throw new Error("不支持的旧版备份格式");
  const data: WorkspaceData = {
    requirements: backup.data.requirements.map((item) => normalizeRequirement(item)),
    domains: normalizeDomains(backup.data.domains),
    products: backup.data.products,
    groupOverrides: backup.data.groupOverrides.map(normalizeLegacyGroupOverride),
  };
  validateWorkspaceData(data);
  return {
    data,
    preview: {
      sourceFormat: "legacy-json",
      exportedAt: backup.exportedAt,
      templateVersion: "旧版备份",
      counts: workspaceCounts(data),
    },
  };
}

function validateManifest(value: WorkspacePackageManifest): void {
  if (value?.format !== PACKAGE_FORMAT || value.formatVersion !== PACKAGE_VERSION) throw new Error("不支持的数据包版本");
  if (value.dataPath !== DATA_PATH || !value.files || !value.files[DATA_PATH]) throw new Error("数据包清单不完整");
  if (!value.exportedAt || !value.templateVersion || !value.counts) throw new Error("数据包元信息不完整");
  for (const record of Object.values(value.files)) {
    if (!record || !/^[a-f0-9]{64}$/.test(record.sha256) || !Number.isInteger(record.sizeBytes) || record.sizeBytes < 0) {
      throw new Error("数据包文件清单无效");
    }
  }
}

export function validateWorkspaceData(data: WorkspaceData): void {
  if (!data || !Array.isArray(data.requirements) || !Array.isArray(data.domains) || !Array.isArray(data.products) || !Array.isArray(data.groupOverrides)) {
    throw new Error("数据包结构无效");
  }
  const domainIds = validateDictionaries(data.domains, "领域");
  const domainsById = new Map(data.domains.map((item) => [item.id, { ...item, level: item.level ?? "L1" }]));
  const productIds = validateDictionaries(data.products, "产品");
  const requirementIds = new Set<string>();

  for (const domain of domainsById.values()) {
    if (domain.level === "L0" && domain.parentId) throw new Error(`领域 L0“${domain.name}”不能设置上游领域`);
    if (domain.level === "L1" && domain.parentId && domainsById.get(domain.parentId)?.level !== "L0") {
      throw new Error(`领域 L1“${domain.name}”引用了不存在的上游领域 L0`);
    }
  }

  for (const raw of data.requirements) {
    validateRawWorkload(raw);
    const item = normalizeRequirement(raw);
    if (!item.id || requirementIds.has(item.id) || !item.title.trim()) throw new Error("数据包包含无效或重复需求");
    requirementIds.add(item.id);
    if (item.domainL0Id && domainsById.get(item.domainL0Id)?.level !== "L0") throw new Error(`需求“${item.title}”引用了不存在的领域 L0`);
    if (!domainIds.has(item.domainId) || domainsById.get(item.domainId)?.level !== "L1") throw new Error(`需求“${item.title}”引用了不存在的领域 L1`);
    if (item.domainL0Id && domainsById.get(item.domainId)?.parentId !== item.domainL0Id) throw new Error(`需求“${item.title}”的领域 L1 不属于所选领域 L0`);
    if (!SOURCES.includes(item.source) || !CATEGORIES.includes(item.category)) throw new Error(`需求“${item.title}”的来源或分类无效`);
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(item.targetMonth)) throw new Error(`需求“${item.title}”的上线年月无效`);
    if (!Number.isFinite(item.workloadPm) || item.workloadPm <= 0) throw new Error(`需求“${item.title}”的工作量无效`);
    if (item.description.length > 5000 || !validTimestamp(item.createdAt) || !validTimestamp(item.updatedAt)) throw new Error(`需求“${item.title}”的文本或时间信息无效`);
    if (item.productIds.some((id) => !productIds.has(id))) throw new Error(`需求“${item.title}”引用了不存在的产品`);
    if (item.category === "产品专属" && item.productIds.length !== 1) throw new Error(`产品专属需求“${item.title}”必须且只能匹配一个产品`);
    if (item.category === "体验优化" && item.productIds.length > 0) throw new Error(`体验优化需求“${item.title}”不应匹配产品`);
    if (item.images.length > MAX_REQUIREMENT_IMAGES) throw new Error(`需求“${item.title}”的图片数量超限`);
    for (const image of item.images) decodeDataUrl(image);
  }

  const overrideKeys = new Set<string>();
  for (const item of data.groupOverrides) {
    if (!item.groupKey || overrideKeys.has(item.groupKey) || !item.cardTitle?.trim() || !validTimestamp(item.updatedAt)) throw new Error("数据包包含无效或重复路标卡片编辑");
    overrideKeys.add(item.groupKey);
    const [domainId, source, targetMonth] = item.groupKey.split("::");
    if (!domainIds.has(domainId) || !SOURCES.includes(normalizeLegacySource(source) as (typeof SOURCES)[number]) || !/^\d{4}-(0[1-9]|1[0-2])$/.test(targetMonth ?? "")) {
      throw new Error(`路标卡片“${item.cardTitle}”的关联信息无效`);
    }
  }
}

function validateRawWorkload(raw: Requirement): void {
  const keys = ["deviceWorkloadPm", "appWorkloadPm", "cloudWorkloadPm", "unallocatedWorkloadPm"] as const;
  const hasBreakdown = keys.some((key) => Object.prototype.hasOwnProperty.call(raw, key));
  if (!hasBreakdown) {
    if (!Number.isFinite(raw.workloadPm) || raw.workloadPm <= 0) throw new Error(`需求“${raw.title}”的工作量无效`);
    return;
  }
  for (const key of keys) {
    const value = raw[key];
    if (!Number.isFinite(value) || value < 0) throw new Error(`需求“${raw.title}”的${key === "deviceWorkloadPm" ? "设备" : key === "appWorkloadPm" ? "App" : key === "cloudWorkloadPm" ? "云侧" : "待拆分"}工作量无效`);
  }
  const total = raw.deviceWorkloadPm + raw.appWorkloadPm + raw.cloudWorkloadPm + raw.unallocatedWorkloadPm;
  if (total <= 0 || !Number.isFinite(raw.workloadPm) || Math.abs(total - raw.workloadPm) > 0.001) {
    throw new Error(`需求“${raw.title}”的工作量合计不一致`);
  }
}

function validateDictionaries(items: WorkspaceData["domains"], label: string): Set<string> {
  const ids = new Set<string>();
  const names = new Set<string>();
  for (const item of items) {
    const name = item?.name?.trim().toLocaleLowerCase("zh-CN");
    const dictionaryName = `${item.level ?? "L1"}:${name}`;
    if (!item?.id || !name || ids.has(item.id) || names.has(dictionaryName)) throw new Error(`数据包包含无效或重复${label}`);
    ids.add(item.id);
    names.add(dictionaryName);
  }
  return ids;
}

function normalizeDomains(items: WorkspaceData["domains"]): WorkspaceData["domains"] {
  return items.map((item) => {
    const level = item.level ?? "L1";
    if (level === "L0") {
      const { parentId: _ignoredParent, ...rest } = item;
      return { ...rest, level };
    }
    return { ...item, level };
  });
}

function decodeDataUrl(image: RequirementImage): Buffer {
  if (!REQUIREMENT_IMAGE_MIME_TYPES.includes(image.mimeType)) throw new Error(`不支持图片格式：${image.name}`);
  const prefix = `data:${image.mimeType};base64,`;
  if (!image.dataUrl.startsWith(prefix)) throw new Error(`图片数据无效：${image.name}`);
  const content = Buffer.from(image.dataUrl.slice(prefix.length), "base64");
  if (!content.length || content.length !== image.sizeBytes || content.length > MAX_REQUIREMENT_IMAGE_BYTES) throw new Error(`图片大小无效：${image.name}`);
  return content;
}

function uniqueImagePath(requirementId: string, imageId: string, mimeType: RequirementImage["mimeType"], index: number, used: ReadonlySet<string>): string {
  const extension = ({ "image/png": "png", "image/jpeg": "jpg", "image/webp": "webp", "image/gif": "gif" } as const)[mimeType];
  const base = `images/${safeSegment(requirementId)}/${safeSegment(imageId || String(index + 1))}`;
  let candidate = `${base}.${extension}`;
  let counter = 1;
  while (used.has(candidate)) candidate = `${base}-${counter++}.${extension}`;
  return candidate;
}

function safeSegment(value: string): string {
  return value.replace(/[^a-zA-Z0-9._-]/g, "-").replace(/^\.+/, "") || "item";
}

function validateArchivePath(value: string): void {
  if (!value || value.startsWith("/") || value.split("/").includes("..")) throw new Error("数据包包含不安全的文件路径");
}

function fileRecord(bytes: Buffer): PackageFileRecord {
  return { sha256: sha256(bytes), sizeBytes: bytes.length };
}

function sha256(bytes: Buffer): string {
  return createHash("sha256").update(bytes).digest("hex");
}

function firstNonWhitespace(bytes: Buffer): string {
  return bytes.toString("utf8", 0, Math.min(bytes.length, 128)).trimStart()[0] ?? "";
}

function validTimestamp(value: string): boolean {
  return typeof value === "string" && Number.isFinite(Date.parse(value));
}

function sameCounts(left: WorkspacePackageCounts, right: WorkspacePackageCounts): boolean {
  return left.requirements === right.requirements
    && left.images === right.images
    && left.domains === right.domains
    && left.products === right.products
    && left.groupOverrides === right.groupOverrides;
}
