import initSqlJs, { type Database as SqlDatabase, type SqlJsStatic } from "sql.js";
import { app } from "electron";
import { existsSync } from "node:fs";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { join } from "node:path";
import type {
  AppSnapshot,
  BackupEnvelope,
  DictionaryItem,
  GroupOverride,
  Requirement,
  SaveDictionaryInput,
  SaveGroupOverrideInput,
  SaveRequirementInput,
  TemplateProfile,
  WorkspaceData,
  WorkspaceImportMode,
  WorkspaceWorkbookConflictMode,
} from "../shared/types.js";
import {
  MAX_REQUIREMENT_IMAGES,
  MAX_REQUIREMENT_IMAGE_BYTES,
  OVERSEAS_REGIONS,
  REQUIREMENT_IMAGE_MIME_TYPES,
} from "../shared/types.js";
import { normalizeRequirement } from "../shared/requirements.js";
import { mergeWorkspaceData } from "../shared/workspace.js";
import { roundWorkload } from "../shared/workload.js";
import { applyPreparedWorkspaceWorkbook, type PreparedWorkspaceWorkbookImport } from "./workspaceWorkbook.js";

const SCHEMA_VERSION = 1;

export class LocalDatabase {
  private sql!: SqlJsStatic;
  private db!: SqlDatabase;
  private readonly databasePath: string;

  constructor() {
    this.databasePath = join(app.getPath("userData"), "roadmap.db");
  }

  async initialize(): Promise<void> {
    const packagedWasmPath = join(process.resourcesPath, "sql-wasm.wasm");
    const wasmPath = existsSync(packagedWasmPath)
      ? packagedWasmPath
      : join(app.getAppPath(), "node_modules", "sql.js", "dist", "sql-wasm.wasm");
    this.sql = await initSqlJs({ locateFile: () => wasmPath });
    const bytes = existsSync(this.databasePath) ? await readFile(this.databasePath) : undefined;
    this.db = bytes ? new this.sql.Database(bytes) : new this.sql.Database();
    this.migrate();
    await this.persist();
  }

  getSnapshot(): AppSnapshot {
    return {
      requirements: this.all<{ payload: string }>("SELECT payload FROM requirements ORDER BY updated_at DESC").map((row) => normalizeRequirement(JSON.parse(row.payload))),
      domains: this.all<{ payload: string }>("SELECT payload FROM domains ORDER BY sort_order, name").map((row) => JSON.parse(row.payload) as DictionaryItem),
      products: this.all<{ payload: string }>("SELECT payload FROM products ORDER BY sort_order, name").map((row) => JSON.parse(row.payload) as DictionaryItem),
      groupOverrides: this.all<{ payload: string }>("SELECT payload FROM group_overrides ORDER BY group_key").map((row) => JSON.parse(row.payload) as GroupOverride),
      templates: this.all<{ payload: string }>("SELECT payload FROM templates ORDER BY imported_at DESC").map((row) => JSON.parse(row.payload) as TemplateProfile),
    };
  }

  async saveRequirement(input: SaveRequirementInput): Promise<AppSnapshot> {
    validateRequirement(input);
    const now = new Date().toISOString();
    const existingPayload = input.id ? this.getPayload<Requirement>("requirements", "id", input.id) : undefined;
    const existing = existingPayload ? normalizeRequirement(existingPayload) : undefined;
    const deviceWorkloadPm = roundWorkload(input.deviceWorkloadPm);
    const appWorkloadPm = roundWorkload(input.appWorkloadPm);
    const cloudWorkloadPm = roundWorkload(input.cloudWorkloadPm);
    const workloadPm = roundWorkload(deviceWorkloadPm + appWorkloadPm + cloudWorkloadPm);
    const requirement: Requirement = {
      ...input,
      id: input.id ?? crypto.randomUUID(),
      title: input.title.trim(),
      description: input.description.trim(),
      images: input.images.map((image) => ({ ...image, name: image.name.trim() || "需求图片" })),
      overseasRegions: input.source === "海外研究" ? [...new Set(input.overseasRegions)] : [],
      productIds: [...new Set(input.productIds)],
      deviceWorkloadPm,
      appWorkloadPm,
      cloudWorkloadPm,
      unallocatedWorkloadPm: 0,
      workloadPm,
      createdAt: existing?.createdAt ?? now,
      updatedAt: now,
    };
    this.run(
      "INSERT OR REPLACE INTO requirements (id, updated_at, payload) VALUES (?, ?, ?)",
      [requirement.id, requirement.updatedAt, JSON.stringify(requirement)],
    );
    await this.persist();
    return this.getSnapshot();
  }

  async deleteRequirement(id: string): Promise<AppSnapshot> {
    this.run("DELETE FROM requirements WHERE id = ?", [id]);
    await this.persist();
    return this.getSnapshot();
  }

  async saveDomain(input: SaveDictionaryInput): Promise<AppSnapshot> {
    return this.saveDictionary("domains", input);
  }

  async saveProduct(input: SaveDictionaryInput): Promise<AppSnapshot> {
    return this.saveDictionary("products", input);
  }

  async deleteDomain(id: string): Promise<AppSnapshot> {
    const used = this.getSnapshot().requirements.some((item) => item.domainId === id);
    if (used) throw new Error("该领域已被需求使用，请先迁移相关需求");
    this.run("DELETE FROM domains WHERE id = ?", [id]);
    await this.persist();
    return this.getSnapshot();
  }

  async deleteProduct(id: string): Promise<AppSnapshot> {
    const used = this.getSnapshot().requirements.some((item) => item.productIds.includes(id));
    if (used) throw new Error("该产品已被需求使用，请先移除相关产品匹配");
    this.run("DELETE FROM products WHERE id = ?", [id]);
    await this.persist();
    return this.getSnapshot();
  }

  async saveGroupOverride(input: SaveGroupOverrideInput): Promise<AppSnapshot> {
    if (!input.groupKey.trim() || !input.cardTitle.trim()) throw new Error("卡片标题不能为空");
    const value: GroupOverride = {
      groupKey: input.groupKey,
      cardTitle: input.cardTitle.trim(),
      cardSummary: input.cardSummary.trim(),
      updatedAt: new Date().toISOString(),
    };
    this.run(
      "INSERT OR REPLACE INTO group_overrides (group_key, updated_at, payload) VALUES (?, ?, ?)",
      [value.groupKey, value.updatedAt, JSON.stringify(value)],
    );
    await this.persist();
    return this.getSnapshot();
  }

  async saveTemplate(profile: TemplateProfile): Promise<AppSnapshot> {
    this.run(
      "INSERT OR REPLACE INTO templates (id, imported_at, payload) VALUES (?, ?, ?)",
      [profile.id, profile.importedAt, JSON.stringify(profile)],
    );
    await this.persist();
    return this.getSnapshot();
  }

  createBackup(): BackupEnvelope {
    return { schemaVersion: SCHEMA_VERSION, exportedAt: new Date().toISOString(), data: this.getWorkspaceData() };
  }

  async restoreBackup(backup: BackupEnvelope): Promise<AppSnapshot> {
    if (backup.schemaVersion !== SCHEMA_VERSION) throw new Error(`不支持的备份版本：${backup.schemaVersion}`);
    return this.importWorkspace(backup.data, "replace");
  }

  getWorkspaceData(): WorkspaceData {
    const { templates: _templates, ...data } = this.getSnapshot();
    return data;
  }

  async importWorkspace(incoming: WorkspaceData, mode: WorkspaceImportMode): Promise<AppSnapshot> {
    const data = mode === "merge" ? mergeWorkspaceData(this.getWorkspaceData(), incoming) : incoming;
    this.db.run("BEGIN");
    try {
      for (const table of ["requirements", "domains", "products", "group_overrides"]) this.db.run(`DELETE FROM ${table}`);
      for (const item of data.requirements) {
        const normalized = normalizeRequirement(item);
        this.run("INSERT INTO requirements (id, updated_at, payload) VALUES (?, ?, ?)", [normalized.id, normalized.updatedAt, JSON.stringify(normalized)]);
      }
      for (const [table, items] of [["domains", data.domains], ["products", data.products]] as const) {
        for (const item of items) {
          this.run(`INSERT INTO ${table} (id, sort_order, name, payload) VALUES (?, ?, ?, ?)`, [item.id, item.sortOrder, item.name, JSON.stringify(item)]);
        }
      }
      for (const item of data.groupOverrides) {
        this.run("INSERT INTO group_overrides (group_key, updated_at, payload) VALUES (?, ?, ?)", [item.groupKey, item.updatedAt, JSON.stringify(item)]);
      }
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
    await this.persist();
    return this.getSnapshot();
  }

  async importWorkspaceWorkbook(prepared: PreparedWorkspaceWorkbookImport, conflictMode: WorkspaceWorkbookConflictMode): Promise<AppSnapshot> {
    const data = applyPreparedWorkspaceWorkbook(this.getWorkspaceData(), prepared, conflictMode);
    return this.importWorkspace(data, "replace");
  }

  private async saveDictionary(table: "domains" | "products", input: SaveDictionaryInput): Promise<AppSnapshot> {
    const name = input.name.trim();
    if (!name) throw new Error("名称不能为空");
    const duplicate = this.all<{ id: string }>(`SELECT id FROM ${table} WHERE lower(name) = lower(?) AND id != ?`, [name, input.id ?? ""])[0];
    if (duplicate) throw new Error(`名称“${name}”已存在`);
    const existing = input.id ? this.getPayload<DictionaryItem>(table, "id", input.id) : undefined;
    const count = this.all<{ count: number }>(`SELECT COUNT(*) AS count FROM ${table}`)[0]?.count ?? 0;
    const item: DictionaryItem = {
      id: input.id ?? crypto.randomUUID(),
      name,
      sortOrder: input.sortOrder ?? existing?.sortOrder ?? count,
      active: input.active ?? existing?.active ?? true,
    };
    this.run(
      `INSERT OR REPLACE INTO ${table} (id, sort_order, name, payload) VALUES (?, ?, ?, ?)`,
      [item.id, item.sortOrder, item.name, JSON.stringify(item)],
    );
    await this.persist();
    return this.getSnapshot();
  }

  private migrate(): void {
    this.db.run(`
      CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS requirements (id TEXT PRIMARY KEY, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS domains (id TEXT PRIMARY KEY, sort_order INTEGER NOT NULL, name TEXT NOT NULL COLLATE NOCASE UNIQUE, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS products (id TEXT PRIMARY KEY, sort_order INTEGER NOT NULL, name TEXT NOT NULL COLLATE NOCASE UNIQUE, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS group_overrides (group_key TEXT PRIMARY KEY, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS templates (id TEXT PRIMARY KEY, imported_at TEXT NOT NULL, payload TEXT NOT NULL);
      INSERT OR REPLACE INTO meta (key, value) VALUES ('schema_version', '${SCHEMA_VERSION}');
    `);
  }

  private run(sql: string, params: Array<string | number | null> = []): void {
    this.db.run(sql, params);
  }

  private all<T>(sql: string, params: Array<string | number | null> = []): T[] {
    const statement = this.db.prepare(sql);
    try {
      statement.bind(params);
      const rows: T[] = [];
      while (statement.step()) rows.push(statement.getAsObject() as T);
      return rows;
    } finally {
      statement.free();
    }
  }

  private getPayload<T>(table: string, key: string, value: string): T | undefined {
    const row = this.all<{ payload: string }>(`SELECT payload FROM ${table} WHERE ${key} = ?`, [value])[0];
    return row ? JSON.parse(row.payload) : undefined;
  }

  private async persist(): Promise<void> {
    await mkdir(app.getPath("userData"), { recursive: true });
    const tempPath = `${this.databasePath}.tmp`;
    await writeFile(tempPath, Buffer.from(this.db.export()));
    await rename(tempPath, this.databasePath);
  }
}

function validateRequirement(input: SaveRequirementInput): void {
  if (!input.title.trim()) throw new Error("需求标题不能为空");
  if (!input.domainId) throw new Error("请选择领域");
  if (input.overseasRegions.some((item) => !OVERSEAS_REGIONS.includes(item))) throw new Error("海外研究区域包含无效选项");
  if (input.source === "海外研究" && input.overseasRegions.length === 0) throw new Error("海外研究需求必须选择至少一个区域");
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(input.targetMonth)) throw new Error("上线年月格式应为 YYYY-MM");
  const workloadParts = [input.deviceWorkloadPm, input.appWorkloadPm, input.cloudWorkloadPm];
  if (workloadParts.some((value) => !Number.isFinite(value) || value < 0)) throw new Error("各侧工作量必须是大于或等于0的人月数");
  if (workloadParts.reduce((sum, value) => sum + value, 0) <= 0) throw new Error("设备、App、云侧工作量至少填写一项");
  if (input.description.length > 5000) throw new Error("需求描述不能超过5000字");
  if (input.images.length > MAX_REQUIREMENT_IMAGES) throw new Error(`每条需求最多添加${MAX_REQUIREMENT_IMAGES}张图片`);
  for (const image of input.images) {
    if (!REQUIREMENT_IMAGE_MIME_TYPES.includes(image.mimeType)) throw new Error(`不支持图片格式：${image.name}`);
    if (image.sizeBytes <= 0 || image.sizeBytes > MAX_REQUIREMENT_IMAGE_BYTES) throw new Error(`图片“${image.name}”不能超过5MB`);
    if (!image.dataUrl.startsWith(`data:${image.mimeType};base64,`)) throw new Error(`图片“${image.name}”数据无效`);
  }
  if (input.category === "产品专属" && input.productIds.length === 0) throw new Error("产品专属需求必须选择至少一个产品");
}
