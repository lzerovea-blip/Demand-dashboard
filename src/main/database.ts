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
  Member,
  PptAccessPolicy,
  Requirement,
  SaveDictionaryInput,
  SaveGroupOrderInput,
  SaveGroupOverrideInput,
  SaveMemberInput,
  SavePptAccessPolicyInput,
  SaveRequirementInput,
  SaveTieringSupportInput,
  SaveTieringEvidenceSourceInput,
  UpdateTieringIssueInput,
  TieringChangeLog,
  TieringEvidenceSource,
  TieringFeature,
  TieringProductSeries,
  TieringSupportRecord,
  TieringVerificationIssue,
  TemplateProfile,
  WorkspaceData,
  WorkspaceImportMode,
  WorkspaceWorkbookConflictMode,
} from "../shared/types.js";
import {
  MAX_REQUIREMENT_IMAGES,
  MAX_REQUIREMENT_IMAGE_BYTES,
  REQUIREMENT_SCOPES,
  REQUIREMENT_IMAGE_MIME_TYPES,
  SOURCES,
  MEMBER_ROLES,
} from "../shared/types.js";
import { cloneRequirementAsCopy, LEGACY_AI_SOURCE, LEGACY_MEDICAL_SUBMISSION_SOURCE, LEGACY_OVERSEAS_SOURCE, LEGACY_UPPERCASE_AI_SOURCE, normalizeLegacyGroupOverride, normalizeRequirement, normalizeRequirementProductIds } from "../shared/requirements.js";
import { EMPTY_PPT_ACCESS_POLICY, isValidMicrosoftAccount, normalizeMember, normalizeMicrosoftAccount, normalizePptAccessPolicy, normalizePptWhitelistEntries } from "../shared/members.js";
import { groupKeyFor } from "../shared/roadmap.js";
import { mergeWorkspaceData } from "../shared/workspace.js";
import { roundWorkload } from "../shared/workload.js";
import { applyPreparedWorkspaceWorkbook, type PreparedWorkspaceWorkbookImport } from "./workspaceWorkbook.js";
import {
  buildHistoricalRequirement,
  buildTieringFeatureForRequirement,
  createTieringSupportRecord,
  domainReferenceMessages,
  EMPTY_TIERING_DATA,
  isHistoricalTieringFeature,
  migrateArchivedHistoricalRequirement,
  missingBuiltInTieringProductSeries,
  normalizeTieringData,
  renameIndependentTieringDomain,
  resolveTieringFeature,
  resolveTieringRequirementDomains,
  tieringFeatureIdForRequirement,
  validateTieringSupportRecord,
} from "../shared/tiering.js";
import { historicalLaunchForFeatureId, normalizeHistoricalLaunchMetadata } from "../shared/tieringLaunches.js";

const BACKUP_SCHEMA_VERSION = 1;
const DATABASE_SCHEMA_VERSION = 7;

export class LocalDatabase {
  private sql!: SqlJsStatic;
  private db!: SqlDatabase;
  private readonly databasePath: string;

  constructor() {
    this.databasePath = join(app.getPath("userData"), "roadmap.db");
  }

  async initialize(): Promise<{ skippedTieringRequirements: number }> {
    const packagedWasmPath = join(process.resourcesPath, "sql-wasm.wasm");
    const wasmPath = existsSync(packagedWasmPath)
      ? packagedWasmPath
      : join(app.getAppPath(), "node_modules", "sql.js", "dist", "sql-wasm.wasm");
    this.sql = await initSqlJs({ locateFile: () => wasmPath });
    const bytes = existsSync(this.databasePath) ? await readFile(this.databasePath) : undefined;
    this.db = bytes ? new this.sql.Database(bytes) : new this.sql.Database();
    const skippedTieringRequirements = this.migrate();
    await this.persist();
    return { skippedTieringRequirements };
  }

  getSnapshot(): AppSnapshot {
    return {
      requirements: this.all<{ payload: string }>("SELECT payload FROM requirements ORDER BY updated_at DESC").map((row) => normalizeRequirement(JSON.parse(row.payload))),
      domains: this.all<{ payload: string }>("SELECT payload FROM domains ORDER BY sort_order, name").map((row) => {
        const item = JSON.parse(row.payload) as DictionaryItem;
        return { ...item, level: item.level ?? "L1" };
      }),
      products: this.all<{ payload: string }>("SELECT payload FROM products ORDER BY sort_order, name").map((row) => JSON.parse(row.payload) as DictionaryItem),
      groupOverrides: this.all<{ payload: string }>("SELECT payload FROM group_overrides ORDER BY group_key").map((row) => JSON.parse(row.payload) as GroupOverride),
      members: this.all<{ payload: string }>("SELECT payload FROM members ORDER BY sort_order, name").map((row) => normalizeMember(JSON.parse(row.payload) as Member)),
      pptAccessPolicy: normalizePptAccessPolicy(this.getPayload<PptAccessPolicy>("settings", "key", "ppt_access_policy") ?? EMPTY_PPT_ACCESS_POLICY),
      tiering: normalizeTieringData({
        features: this.payloads<TieringFeature>("tiering_features", "sort_order, id"),
        productSeries: this.payloads<TieringProductSeries>("tiering_product_series", "sort_order, id"),
        supportRecords: this.payloads<TieringSupportRecord>("tiering_support_records", "id"),
        evidenceSources: this.payloads<TieringEvidenceSource>("tiering_evidence_sources", "id"),
        verificationIssues: this.payloads<TieringVerificationIssue>("tiering_issues", "id"),
        changeLogs: this.payloads<TieringChangeLog>("tiering_change_logs", "changed_at DESC, id"),
      }),
      templates: this.all<{ payload: string }>("SELECT payload FROM templates ORDER BY imported_at DESC").map((row) => JSON.parse(row.payload) as TemplateProfile),
    };
  }

  async saveRequirement(input: SaveRequirementInput): Promise<AppSnapshot> {
    validateRequirement(input);
    const domains = new Map(this.getSnapshot().domains.map((item) => [item.id, item]));
    if (domains.get(input.domainL0Id)?.level !== "L0") throw new Error("领域 L0 选择无效");
    const domainL1 = domains.get(input.domainId);
    if ((domainL1?.level ?? "L1") !== "L1" || domainL1?.parentId !== input.domainL0Id) throw new Error("领域 L1 不属于所选领域 L0");
    const now = new Date().toISOString();
    const existingPayload = input.id ? this.getPayload<Requirement>("requirements", "id", input.id) : undefined;
    const existing = existingPayload ? normalizeRequirement(existingPayload) : undefined;
    const deviceWorkloadPm = roundWorkload(input.deviceWorkloadPm);
    const appWorkloadPm = roundWorkload(input.appWorkloadPm);
    const cloudWorkloadPm = roundWorkload(input.cloudWorkloadPm);
    const workloadPm = roundWorkload(deviceWorkloadPm + appWorkloadPm + cloudWorkloadPm);
    const requirementId = input.id ?? crypto.randomUUID();
    const requirement: Requirement = {
      ...input,
      id: requirementId,
      title: input.title.trim(),
      description: input.description.trim(),
      images: input.images.map((image) => ({ ...image, name: image.name.trim() || "需求图片" })),
      overseasRegions: [],
      productIds: normalizeRequirementProductIds(input.category, input.productIds),
      deviceWorkloadPm,
      appWorkloadPm,
      cloudWorkloadPm,
      unallocatedWorkloadPm: 0,
      workloadPm,
      createdAt: existing?.createdAt ?? now,
      updatedAt: now,
      tieringFeatureId: input.tieringFeatureId ?? existing?.tieringFeatureId ?? tieringFeatureIdForRequirement(requirementId),
      historicalLaunch: existing?.historicalLaunch
        ? { ...existing.historicalLaunch, suggestedTargetMonth: input.targetMonth }
        : undefined,
    };
    this.db.run("BEGIN");
    try {
      this.run(
        "INSERT OR REPLACE INTO requirements (id, updated_at, payload) VALUES (?, ?, ?)",
        [requirement.id, requirement.updatedAt, JSON.stringify(requirement)],
      );
      this.upsertTieringForRequirement(requirement, domains, now);
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
    await this.persist();
    return this.getSnapshot();
  }

  async deleteRequirement(id: string): Promise<AppSnapshot> {
    const requirement = this.getPayload<Requirement>("requirements", "id", id);
    this.db.run("BEGIN");
    try {
      this.run("DELETE FROM requirements WHERE id = ?", [id]);
      const tiering = normalizeTieringData({ features: this.payloads<TieringFeature>("tiering_features", "sort_order, id") });
      for (const feature of tiering.features.filter((item) => item.requirementId === id || item.id === requirement?.tieringFeatureId)) {
        if (isHistoricalTieringFeature(feature)) {
          const updated = { ...feature, requirementId: undefined, updatedAt: new Date().toISOString() };
          this.run("UPDATE tiering_features SET updated_at = ?, payload = ? WHERE id = ?", [updated.updatedAt, JSON.stringify(updated), updated.id]);
        } else {
          this.run("DELETE FROM tiering_support_records WHERE id LIKE ?", [`${feature.id}::%`]);
          for (const issue of this.payloads<TieringVerificationIssue>("tiering_issues", "id").filter((item) => item.featureId === feature.id)) this.run("DELETE FROM tiering_issues WHERE id = ?", [issue.id]);
          for (const log of this.payloads<TieringChangeLog>("tiering_change_logs", "changed_at DESC, id").filter((item) => item.featureId === feature.id)) this.run("DELETE FROM tiering_change_logs WHERE id = ?", [log.id]);
          this.run("DELETE FROM tiering_features WHERE id = ?", [feature.id]);
        }
      }
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
    await this.persist();
    return this.getSnapshot();
  }

  async duplicateRequirement(id: string): Promise<AppSnapshot> {
    const snapshot = this.getSnapshot();
    const source = snapshot.requirements.find((item) => item.id === id);
    if (!source) throw new Error("要复制的需求不存在或已被删除");
    const copy = cloneRequirementAsCopy(source, snapshot.requirements.map((item) => item.title));
    const normalizedCopy = { ...copy, tieringFeatureId: tieringFeatureIdForRequirement(copy.id) };
    const domains = new Map(snapshot.domains.map((item) => [item.id, item]));
    this.db.run("BEGIN");
    try {
      this.run(
        "INSERT INTO requirements (id, updated_at, payload) VALUES (?, ?, ?)",
        [normalizedCopy.id, normalizedCopy.updatedAt, JSON.stringify(normalizedCopy)],
      );
      this.upsertTieringForRequirement(normalizedCopy, domains, normalizedCopy.updatedAt);
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
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
    const snapshot = this.getSnapshot();
    const domain = snapshot.domains.find((item) => item.id === id);
    if (!domain) return snapshot;
    const references = domainReferenceMessages(domain, { ...snapshot, features: snapshot.tiering?.features ?? [] });
    if (references.length) throw new Error(`无法删除“${domain.name}”：\n${references.join("\n")}`);
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

  async saveMember(input: SaveMemberInput): Promise<AppSnapshot> {
    const name = input.name.trim();
    const employeeNumber = input.employeeNumber.trim();
    const microsoftAccount = normalizeMicrosoftAccount(input.microsoftAccount);
    if (!name) throw new Error("成员姓名不能为空");
    if (!employeeNumber) throw new Error("成员工号不能为空");
    if (!isValidMicrosoftAccount(microsoftAccount)) throw new Error("请输入完整有效的 Microsoft 365 企业邮箱账号");
    if (input.role && !MEMBER_ROLES.includes(input.role)) throw new Error("成员岗位角色无效");
    const snapshot = this.getSnapshot();
    if (snapshot.members.some((item) => item.id !== input.id && item.employeeNumber.trim().toLocaleLowerCase("zh-CN") === employeeNumber.toLocaleLowerCase("zh-CN"))) {
      throw new Error(`工号“${employeeNumber}”已存在`);
    }
    if (snapshot.members.some((item) => item.id !== input.id && normalizeMicrosoftAccount(item.microsoftAccount) === microsoftAccount)) {
      throw new Error(`Microsoft 365 账号“${microsoftAccount}”已存在`);
    }
    const domainL1Ids = [...new Set(input.domainL1Ids.filter(Boolean))];
    const domainById = new Map(snapshot.domains.map((item) => [item.id, item]));
    if (domainL1Ids.some((id) => !domainById.has(id) || (domainById.get(id)?.level ?? "L1") !== "L1")) throw new Error("负责领域包含无效的领域 L1");
    const now = new Date().toISOString();
    const existing = input.id ? snapshot.members.find((item) => item.id === input.id) : undefined;
    if (input.id && !existing) throw new Error("要编辑的成员不存在");
    const member: Member = {
      id: existing?.id ?? crypto.randomUUID(),
      name,
      employeeNumber,
      microsoftAccount,
      role: input.role,
      domainL1Ids,
      sortOrder: input.sortOrder ?? existing?.sortOrder ?? snapshot.members.length,
      createdAt: existing?.createdAt ?? now,
      updatedAt: now,
    };
    this.run("INSERT OR REPLACE INTO members (id, sort_order, name, payload) VALUES (?, ?, ?, ?)", [member.id, member.sortOrder, member.name, JSON.stringify(member)]);
    await this.persist();
    return this.getSnapshot();
  }

  async deleteMember(id: string): Promise<AppSnapshot> {
    const snapshot = this.getSnapshot();
    if (snapshot.pptAccessPolicy.entries.some((entry) => entry.memberId === id)) throw new Error("该成员仍在 PPT 白名单中，请先从白名单移除");
    this.run("DELETE FROM members WHERE id = ?", [id]);
    await this.persist();
    return this.getSnapshot();
  }

  async savePptAccessPolicy(input: SavePptAccessPolicyInput): Promise<AppSnapshot> {
    const entries = normalizePptWhitelistEntries(input.entries);
    if (entries.length !== input.entries.length) throw new Error("白名单包含重复成员或无效权限");
    const memberIds = new Set(this.getSnapshot().members.map((item) => item.id));
    if (entries.some((entry) => !memberIds.has(entry.memberId))) throw new Error("白名单引用了不存在的成员");
    const policy: PptAccessPolicy = { entries, updatedAt: new Date().toISOString() };
    this.run("INSERT OR REPLACE INTO settings (key, payload) VALUES (?, ?)", ["ppt_access_policy", JSON.stringify(policy)]);
    await this.persist();
    return this.getSnapshot();
  }

  async saveGroupOverride(input: SaveGroupOverrideInput): Promise<AppSnapshot> {
    if (!input.groupKey.trim() || !input.cardTitle.trim()) throw new Error("卡片标题不能为空");
    const existing = this.getPayload<GroupOverride>("group_overrides", "group_key", input.groupKey);
    const value: GroupOverride = {
      groupKey: input.groupKey,
      cardTitle: input.cardTitle.trim(),
      cardSummary: input.cardSummary.trim(),
      ...(existing?.sortOrder === undefined ? {} : { sortOrder: existing.sortOrder }),
      updatedAt: new Date().toISOString(),
    };
    this.run(
      "INSERT OR REPLACE INTO group_overrides (group_key, updated_at, payload) VALUES (?, ?, ?)",
      [value.groupKey, value.updatedAt, JSON.stringify(value)],
    );
    await this.persist();
    return this.getSnapshot();
  }

  async saveGroupOrder(input: SaveGroupOrderInput): Promise<AppSnapshot> {
    if (!SOURCES.includes(input.source)) throw new Error("路标泳道无效");
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(input.targetMonth)) throw new Error("上线月份无效");
    const orderedKeys = input.groupKeys.map((key) => key.trim());
    if (!orderedKeys.length || orderedKeys.some((key) => !key) || new Set(orderedKeys).size !== orderedKeys.length) {
      throw new Error("卡片顺序包含空值或重复项");
    }
    for (const key of orderedKeys) {
      const [, source, targetMonth] = key.split("::");
      if (source !== input.source || targetMonth !== input.targetMonth) throw new Error("只能调整同一泳道、同一月份内的卡片");
    }

    const snapshot = this.getSnapshot();
    const actualKeys = [...new Set(snapshot.requirements
      .filter((item) => item.source === input.source && item.targetMonth === input.targetMonth)
      .map(groupKeyFor))];
    if (actualKeys.length !== orderedKeys.length || actualKeys.some((key) => !orderedKeys.includes(key))) {
      throw new Error("卡片集合已变化，请刷新路标后重试");
    }

    const domains = new Map(snapshot.domains.map((item) => [item.id, item]));
    const overrides = new Map(snapshot.groupOverrides.map((item) => [item.groupKey, item]));
    const now = new Date().toISOString();
    this.db.run("BEGIN");
    try {
      orderedKeys.forEach((groupKey, sortOrder) => {
        const [domainId] = groupKey.split("::");
        const existing = overrides.get(groupKey);
        const value: GroupOverride = {
          groupKey,
          cardTitle: existing?.cardTitle.trim() || domains.get(domainId)?.name || "未命名领域",
          cardSummary: existing?.cardSummary ?? "",
          sortOrder,
          updatedAt: now,
        };
        this.run(
          "INSERT OR REPLACE INTO group_overrides (group_key, updated_at, payload) VALUES (?, ?, ?)",
          [groupKey, now, JSON.stringify(value)],
        );
      });
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
    await this.persist();
    return this.getSnapshot();
  }

  async saveTieringSupport(input: SaveTieringSupportInput): Promise<AppSnapshot> {
    const snapshot = this.getSnapshot();
    const tiering = normalizeTieringData(snapshot.tiering);
    const previous = tiering.supportRecords.find((item) => item.id === input.recordId);
    if (!previous) throw new Error("支持规格不存在，请刷新后重试");
    const record = { ...previous, status: input.status, updatedAt: new Date().toISOString() };
    validateTieringSupportRecord(record);
    if (!tiering.features.some((item) => item.id === record.featureId)) throw new Error("分层功能不存在");
    if (!tiering.productSeries.some((item) => item.id === record.seriesId)) throw new Error("产品系列不存在");
    const reason = input.reason?.trim() || "分层分级页面调整系列支持状态";
    const change: TieringChangeLog = {
      id: crypto.randomUUID(),
      changedAt: record.updatedAt,
      triggerType: "人工维护",
      featureId: record.featureId,
      seriesId: record.seriesId,
      before: JSON.stringify(previous),
      after: JSON.stringify(record),
      reason,
      sourceIds: [...record.sourceIds],
      reviewer: "",
      recordStatus: "有效",
      note: "",
    };
    this.db.run("BEGIN");
    try {
      this.run("UPDATE tiering_support_records SET updated_at = ?, payload = ? WHERE id = ?", [record.updatedAt, JSON.stringify(record), record.id]);
      this.run("INSERT INTO tiering_change_logs (id, changed_at, payload) VALUES (?, ?, ?)", [change.id, change.changedAt, JSON.stringify(change)]);
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
    await this.persist();
    return this.getSnapshot();
  }

  async saveTieringEvidenceSource(input: SaveTieringEvidenceSourceInput): Promise<AppSnapshot> {
    const source = { ...input.source, title: input.source.title.trim(), updatedAt: new Date().toISOString() };
    if (!source.id.trim() || !source.title) throw new Error("来源 ID 和来源标题不能为空");
    this.run("INSERT OR REPLACE INTO tiering_evidence_sources (id, updated_at, payload) VALUES (?, ?, ?)", [source.id, source.updatedAt, JSON.stringify(source)]);
    await this.persist();
    return this.getSnapshot();
  }

  async updateTieringIssue(input: UpdateTieringIssueInput): Promise<AppSnapshot> {
    const issue = this.getPayload<TieringVerificationIssue>("tiering_issues", "id", input.issueId);
    if (!issue) throw new Error("待核验项不存在");
    const updated: TieringVerificationIssue = {
      ...issue,
      status: input.status.trim() || issue.status,
      owner: input.owner.trim(),
      targetDate: input.targetDate.trim(),
      resolution: input.resolution.trim(),
      updatedAt: new Date().toISOString(),
    };
    this.run("UPDATE tiering_issues SET updated_at = ?, payload = ? WHERE id = ?", [updated.updatedAt, JSON.stringify(updated), updated.id]);
    await this.persist();
    return this.getSnapshot();
  }

  async regenerateTieringRequirement(featureId: string): Promise<AppSnapshot> {
    const snapshot = this.getSnapshot();
    const tiering = normalizeTieringData(snapshot.tiering);
    const feature = tiering.features.find((item) => item.id === featureId);
    if (!feature) throw new Error("分层功能不存在");
    const existing = snapshot.requirements.find((item) => item.tieringFeatureId === featureId || item.id === feature.requirementId);
    if (existing) throw new Error("该功能已有关联需求，无需重新生成");
    const domainL0 = snapshot.domains.find((item) => item.level === "L0" && item.name === feature.l0);
    const domainL1 = snapshot.domains.find((item) => (item.level ?? "L1") === "L1" && item.name === feature.l1 && (!domainL0 || item.parentId === domainL0.id));
    if (!domainL0 || !domainL1) throw new Error("历史需求所需的领域字典不存在，请重新导入分层分级初始包");
    const requirement = buildHistoricalRequirement(feature, tiering.supportRecords, tiering.productSeries, domainL0, domainL1);
    const linkedFeature: TieringFeature = { ...feature, requirementId: requirement.id, updatedAt: new Date().toISOString() };
    this.db.run("BEGIN");
    try {
      this.run("INSERT INTO requirements (id, updated_at, payload) VALUES (?, ?, ?)", [requirement.id, requirement.updatedAt, JSON.stringify(requirement)]);
      this.run("UPDATE tiering_features SET updated_at = ?, payload = ? WHERE id = ?", [linkedFeature.updatedAt, JSON.stringify(linkedFeature), linkedFeature.id]);
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
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
    return { schemaVersion: BACKUP_SCHEMA_VERSION, exportedAt: new Date().toISOString(), data: this.getWorkspaceData() };
  }

  async restoreBackup(backup: BackupEnvelope): Promise<AppSnapshot> {
    if (backup.schemaVersion !== BACKUP_SCHEMA_VERSION) throw new Error(`不支持的备份版本：${backup.schemaVersion}`);
    return this.importWorkspace(backup.data, "replace");
  }

  getWorkspaceData(): WorkspaceData {
    const { templates: _templates, ...data } = this.getSnapshot();
    return data;
  }

  async importWorkspace(incoming: WorkspaceData, mode: WorkspaceImportMode): Promise<AppSnapshot> {
    const data = mode === "merge" ? mergeWorkspaceData(this.getWorkspaceData(), incoming) : incoming;
    const tiering = normalizeTieringData(data.tiering ?? EMPTY_TIERING_DATA);
    this.db.run("BEGIN");
    try {
      for (const table of ["requirements", "domains", "products", "group_overrides", "members", "tiering_features", "tiering_product_series", "tiering_support_records", "tiering_evidence_sources", "tiering_issues", "tiering_change_logs"]) this.db.run(`DELETE FROM ${table}`);
      for (const item of data.requirements) {
        const normalized = normalizeRequirement(item);
        this.run("INSERT INTO requirements (id, updated_at, payload) VALUES (?, ?, ?)", [normalized.id, normalized.updatedAt, JSON.stringify(normalized)]);
      }
      for (const item of data.domains) {
        const normalized = { ...item, level: item.level ?? "L1" };
        this.run("INSERT INTO domains (id, level, sort_order, name, payload) VALUES (?, ?, ?, ?, ?)", [normalized.id, normalized.level, normalized.sortOrder, normalized.name, JSON.stringify(normalized)]);
      }
      for (const item of data.products) {
        this.run("INSERT INTO products (id, sort_order, name, payload) VALUES (?, ?, ?, ?)", [item.id, item.sortOrder, item.name, JSON.stringify(item)]);
      }
      for (const item of data.groupOverrides) {
        const normalized = normalizeLegacyGroupOverride(item);
        this.run("INSERT INTO group_overrides (group_key, updated_at, payload) VALUES (?, ?, ?)", [normalized.groupKey, normalized.updatedAt, JSON.stringify(normalized)]);
      }
      for (const raw of data.members ?? []) {
        const item = normalizeMember(raw);
        this.run("INSERT INTO members (id, sort_order, name, payload) VALUES (?, ?, ?, ?)", [item.id, item.sortOrder, item.name, JSON.stringify(item)]);
      }
      const policy = normalizePptAccessPolicy(data.pptAccessPolicy);
      this.run("INSERT OR REPLACE INTO settings (key, payload) VALUES (?, ?)", ["ppt_access_policy", JSON.stringify(policy)]);
      for (const item of tiering.features) this.run("INSERT INTO tiering_features (id, sort_order, updated_at, payload) VALUES (?, ?, ?, ?)", [item.id, item.sortOrder, item.updatedAt, JSON.stringify(item)]);
      for (const item of tiering.productSeries) this.run("INSERT INTO tiering_product_series (id, sort_order, updated_at, payload) VALUES (?, ?, ?, ?)", [item.id, item.sortOrder, item.updatedAt, JSON.stringify(item)]);
      for (const item of tiering.supportRecords) this.run("INSERT INTO tiering_support_records (id, updated_at, payload) VALUES (?, ?, ?)", [item.id, item.updatedAt, JSON.stringify(item)]);
      for (const item of tiering.evidenceSources) this.run("INSERT INTO tiering_evidence_sources (id, updated_at, payload) VALUES (?, ?, ?)", [item.id, item.updatedAt, JSON.stringify(item)]);
      for (const item of tiering.verificationIssues) this.run("INSERT INTO tiering_issues (id, updated_at, payload) VALUES (?, ?, ?)", [item.id, item.updatedAt, JSON.stringify(item)]);
      for (const item of tiering.changeLogs) this.run("INSERT INTO tiering_change_logs (id, changed_at, payload) VALUES (?, ?, ?)", [item.id, item.changedAt, JSON.stringify(item)]);
      this.ensureBuiltInTieringProductSeries();
      this.reconcileTieringRequirements(true);
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
    const before = this.getSnapshot();
    const name = input.name.trim();
    if (!name) throw new Error("名称不能为空");
    const level = input.level ?? (input.id ? this.getPayload<DictionaryItem>(table, "id", input.id)?.level : undefined) ?? "L1";
    const duplicate = table === "domains"
      ? this.all<{ id: string }>("SELECT id FROM domains WHERE lower(name) = lower(?) AND level = ? AND id != ?", [name, level, input.id ?? ""])[0]
      : this.all<{ id: string }>("SELECT id FROM products WHERE lower(name) = lower(?) AND id != ?", [name, input.id ?? ""])[0];
    if (duplicate) throw new Error(`名称“${name}”已存在`);
    const existing = input.id ? this.getPayload<DictionaryItem>(table, "id", input.id) : undefined;
    if (input.id && !existing) throw new Error("要编辑的字典项不存在");
    if (table === "domains" && existing && (existing.level ?? "L1") !== level) throw new Error("不允许修改领域层级，请通过调整 L1 所属 L0 迁移");
    const count = this.all<{ count: number }>(`SELECT COUNT(*) AS count FROM ${table}`)[0]?.count ?? 0;
    let parentId: string | undefined;
    if (table === "domains" && level === "L1") {
      parentId = input.parentId ?? existing?.parentId;
      if (!parentId) throw new Error("领域 L1 必须选择所属领域 L0");
      const parent = this.getSnapshot().domains.find((item) => item.id === parentId);
      if (parent?.level !== "L0") throw new Error("所属领域 L0 选择无效");
      if (parent.id === input.id) throw new Error("领域不能归属于自身");
    }
    const item: DictionaryItem = {
      id: input.id ?? crypto.randomUUID(),
      name,
      sortOrder: input.sortOrder ?? existing?.sortOrder ?? count,
      active: input.active ?? existing?.active ?? true,
      ...(table === "domains" ? { level, ...(level === "L1" ? { parentId } : {}) } : {}),
    };
    const previousBytes = this.db.export();
    let committed = false;
    this.db.run("BEGIN");
    try {
      if (table === "domains") {
        this.run("INSERT OR REPLACE INTO domains (id, level, sort_order, name, payload) VALUES (?, ?, ?, ?, ?)", [item.id, item.level ?? "L1", item.sortOrder, item.name, JSON.stringify(item)]);
        if (existing) {
          for (const feature of before.tiering?.features ?? []) {
            const updated = renameIndependentTieringDomain(feature, existing, item, before.domains, before.requirements);
            if (updated === feature) continue;
            updated.updatedAt = new Date().toISOString();
            this.run("UPDATE tiering_features SET updated_at = ?, payload = ? WHERE id = ?", [updated.updatedAt, JSON.stringify(updated), feature.id]);
          }
        }
        this.reconcileTieringRequirements(true, true);
      } else {
        this.run("INSERT OR REPLACE INTO products (id, sort_order, name, payload) VALUES (?, ?, ?, ?)", [item.id, item.sortOrder, item.name, JSON.stringify(item)]);
      }
      this.db.run("COMMIT");
      committed = true;
      await this.persist();
    } catch (error) {
      if (!committed) this.db.run("ROLLBACK");
      else {
        this.db.close();
        this.db = new this.sql.Database(previousBytes);
      }
      throw error;
    }
    return this.getSnapshot();
  }

  private upsertTieringForRequirement(
    requirement: Requirement,
    domains: ReadonlyMap<string, DictionaryItem>,
    updatedAt: string,
  ): void {
    const resolvedDomains = resolveTieringRequirementDomains(requirement, [...domains.values()]);
    if (!resolvedDomains) {
      throw new Error(`需求“${requirement.title}”的领域层级无效，无法同步分层分级`);
    }
    const { domainL0, domainL1 } = resolvedDomains;
    const tiering = normalizeTieringData({
      features: this.payloads<TieringFeature>("tiering_features", "sort_order, id"),
      productSeries: this.payloads<TieringProductSeries>("tiering_product_series", "sort_order, id"),
      supportRecords: this.payloads<TieringSupportRecord>("tiering_support_records", "id"),
    });
    const existing = tiering.features.find((item) => item.id === requirement.tieringFeatureId)
      ?? tiering.features.find((item) => item.requirementId === requirement.id);
    const sortOrder = tiering.features.reduce((max, item) => Math.max(max, item.sortOrder), -1) + 1;
    const feature = buildTieringFeatureForRequirement(requirement, domainL0.name, domainL1.name, sortOrder, updatedAt, existing);
    this.run(
      "INSERT OR REPLACE INTO tiering_features (id, sort_order, updated_at, payload) VALUES (?, ?, ?, ?)",
      [feature.id, feature.sortOrder, feature.updatedAt, JSON.stringify(feature)],
    );
    const supportIds = new Set(tiering.supportRecords.map((item) => item.id));
    for (const series of tiering.productSeries) {
      const record = createTieringSupportRecord(feature.id, series.id, updatedAt);
      if (supportIds.has(record.id)) continue;
      this.run("INSERT INTO tiering_support_records (id, updated_at, payload) VALUES (?, ?, ?)", [record.id, record.updatedAt, JSON.stringify(record)]);
    }
  }

  private migrateTieringSchema(): void {
    const rawFeatures = this.payloads<TieringFeature>("tiering_features", "sort_order, id");
    const rawSupportRecords = this.payloads<TieringSupportRecord>("tiering_support_records", "id");
    const normalized = normalizeTieringData({ features: rawFeatures, supportRecords: rawSupportRecords });
    this.db.run("BEGIN");
    try {
      for (const feature of normalized.features) {
        this.run("UPDATE tiering_features SET sort_order = ?, updated_at = ?, payload = ? WHERE id = ?", [feature.sortOrder, feature.updatedAt, JSON.stringify(feature), feature.id]);
      }
      for (const record of normalized.supportRecords) {
        this.run("UPDATE tiering_support_records SET updated_at = ?, payload = ? WHERE id = ?", [record.updatedAt, JSON.stringify(record), record.id]);
      }
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
  }

  private ensureBuiltInTieringProductSeries(): void {
    const existing = this.payloads<TieringProductSeries>("tiering_product_series", "sort_order, id");
    for (const item of missingBuiltInTieringProductSeries(existing)) {
      this.run(
        "INSERT INTO tiering_product_series (id, sort_order, updated_at, payload) VALUES (?, ?, ?, ?)",
        [item.id, item.sortOrder, item.updatedAt, JSON.stringify(item)],
      );
    }
  }

  private reconcileTieringRequirements(transactionOpen = false, tolerateInvalidDomainLinks = false): number {
    const now = new Date().toISOString();
    const domains = new Map(this.getSnapshot().domains.map((item) => [item.id, item]));
    const requirements = this.all<{ id: string; payload: string }>("SELECT id, payload FROM requirements")
      .map((row) => normalizeRequirement(JSON.parse(row.payload)));
    const existingFeatures = normalizeTieringData({ features: this.payloads<TieringFeature>("tiering_features", "sort_order, id") }).features;
    let skippedTieringRequirements = 0;
    if (!transactionOpen) this.db.run("BEGIN");
    try {
      for (const requirement of requirements) {
        const resolvedDomains = resolveTieringRequirementDomains(requirement, [...domains.values()]);
        if (!resolvedDomains) {
          if (!tolerateInvalidDomainLinks) {
            throw new Error(`需求“${requirement.title}”的领域层级无效，无法同步分层分级`);
          }
          skippedTieringRequirements += 1;
          continue;
        }
        const requirementWithDomain = requirement.domainL0Id === resolvedDomains.domainL0.id
          ? requirement
          : { ...requirement, domainL0Id: resolvedDomains.domainL0.id };
        const linked = existingFeatures.find((item) => item.id === requirement.tieringFeatureId)
          ?? existingFeatures.find((item) => item.requirementId === requirement.id);
        if (linked && resolveTieringFeature(linked, requirements, [...domains.values()]).invalidLink) {
          if (!tolerateInvalidDomainLinks) throw new Error(`分层规格“${linked.requirementTitle}”的需求关联存在冲突，请先修复关联`);
          skippedTieringRequirements += 1;
          continue;
        }
        const tieringFeatureId = linked?.id ?? requirementWithDomain.tieringFeatureId ?? tieringFeatureIdForRequirement(requirementWithDomain.id);
        const withLink = requirementWithDomain.tieringFeatureId === tieringFeatureId
          ? requirementWithDomain
          : { ...requirementWithDomain, tieringFeatureId };
        // Repair hierarchy/link caches only; user descriptions, dates and estimates are not migrations.
        const normalized = withLink;
        if (normalized !== requirement) {
          this.run("UPDATE requirements SET payload = ? WHERE id = ?", [JSON.stringify(normalized), normalized.id]);
        }
        const changed = linked && (linked.l0 !== resolvedDomains.domainL0.name || linked.l1 !== resolvedDomains.domainL1.name || linked.requirementTitle !== normalized.title || linked.requirementId !== normalized.id);
        this.upsertTieringForRequirement(normalized, domains, !linked || changed ? now : linked.updatedAt);
      }
      // A broken association is not authorization to erase a specification or its support history.
      skippedTieringRequirements += existingFeatures.filter((feature) => feature.requirementId && !requirements.some((item) => item.id === feature.requirementId || item.tieringFeatureId === feature.id)).length;
      if (!transactionOpen) this.db.run("COMMIT");
      return skippedTieringRequirements;
    } catch (error) {
      if (!transactionOpen) this.db.run("ROLLBACK");
      throw error;
    }
  }

  private migrate(): number {
    this.db.run(`
      CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS requirements (id TEXT PRIMARY KEY, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS domains (id TEXT PRIMARY KEY, level TEXT NOT NULL, sort_order INTEGER NOT NULL, name TEXT NOT NULL COLLATE NOCASE, payload TEXT NOT NULL, UNIQUE(level, name));
      CREATE TABLE IF NOT EXISTS products (id TEXT PRIMARY KEY, sort_order INTEGER NOT NULL, name TEXT NOT NULL COLLATE NOCASE UNIQUE, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS group_overrides (group_key TEXT PRIMARY KEY, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS members (id TEXT PRIMARY KEY, sort_order INTEGER NOT NULL, name TEXT NOT NULL COLLATE NOCASE, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS templates (id TEXT PRIMARY KEY, imported_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS tiering_features (id TEXT PRIMARY KEY, sort_order INTEGER NOT NULL, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS tiering_product_series (id TEXT PRIMARY KEY, sort_order INTEGER NOT NULL, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS tiering_support_records (id TEXT PRIMARY KEY, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS tiering_evidence_sources (id TEXT PRIMARY KEY, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS tiering_issues (id TEXT PRIMARY KEY, updated_at TEXT NOT NULL, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS tiering_change_logs (id TEXT PRIMARY KEY, changed_at TEXT NOT NULL, payload TEXT NOT NULL);
    `);
    if (!this.getPayload<PptAccessPolicy>("settings", "key", "ppt_access_policy")) {
      this.run("INSERT INTO settings (key, payload) VALUES (?, ?)", ["ppt_access_policy", JSON.stringify(EMPTY_PPT_ACCESS_POLICY)]);
    }
    this.migrateDomainLevels();
    this.migrateTieringSchema();
    this.ensureBuiltInTieringProductSeries();
    this.migrateLegacySources();
    this.migrateRequirementScopes();
    this.migrateHistoricalLaunches();
    const skippedTieringRequirements = this.reconcileTieringRequirements(false, true);
    this.run("INSERT OR REPLACE INTO meta (key, value) VALUES ('schema_version', ?)", [DATABASE_SCHEMA_VERSION]);
    return skippedTieringRequirements;
  }

  private migrateDomainLevels(): void {
    const definition = this.all<{ sql: string }>("SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'domains'")[0]?.sql ?? "";
    if (/\blevel\s+TEXT\b/i.test(definition)) return;
    const rows = this.all<{ id: string; sort_order: number; name: string; payload: string }>("SELECT id, sort_order, name, payload FROM domains");
    this.db.run("BEGIN");
    try {
      this.db.run("CREATE TABLE domains_v2 (id TEXT PRIMARY KEY, level TEXT NOT NULL, sort_order INTEGER NOT NULL, name TEXT NOT NULL COLLATE NOCASE, payload TEXT NOT NULL, UNIQUE(level, name))");
      for (const row of rows) {
        const raw = JSON.parse(row.payload) as DictionaryItem;
        const item = { ...raw, level: raw.level ?? "L1" };
        this.run("INSERT INTO domains_v2 (id, level, sort_order, name, payload) VALUES (?, ?, ?, ?, ?)", [row.id, item.level ?? "L1", row.sort_order, row.name, JSON.stringify(item)]);
      }
      this.db.run("DROP TABLE domains");
      this.db.run("ALTER TABLE domains_v2 RENAME TO domains");
      this.db.run("COMMIT");
    } catch (error) {
      this.db.run("ROLLBACK");
      throw error;
    }
  }

  private migrateLegacySources(): void {
    for (const row of this.all<{ id: string; payload: string }>("SELECT id, payload FROM requirements")) {
      const raw = JSON.parse(row.payload) as Omit<Requirement, "source"> & {
        source: Requirement["source"] | typeof LEGACY_OVERSEAS_SOURCE | typeof LEGACY_AI_SOURCE | typeof LEGACY_UPPERCASE_AI_SOURCE | typeof LEGACY_MEDICAL_SUBMISSION_SOURCE;
      };
      if (raw.source !== LEGACY_OVERSEAS_SOURCE
        && raw.source !== LEGACY_AI_SOURCE
        && raw.source !== LEGACY_UPPERCASE_AI_SOURCE
        && raw.source !== LEGACY_MEDICAL_SUBMISSION_SOURCE) continue;
      const normalized = normalizeRequirement(raw);
      this.run("UPDATE requirements SET payload = ? WHERE id = ?", [JSON.stringify(normalized), row.id]);
    }

    for (const row of this.all<{ group_key: string; updated_at: string; payload: string }>("SELECT group_key, updated_at, payload FROM group_overrides")) {
      const legacy = JSON.parse(row.payload) as GroupOverride;
      const normalized = normalizeLegacyGroupOverride(legacy);
      if (normalized.groupKey === row.group_key) continue;
      const existing = this.getPayload<GroupOverride>("group_overrides", "group_key", normalized.groupKey);
      this.run("DELETE FROM group_overrides WHERE group_key = ?", [row.group_key]);
      if (!existing || Date.parse(normalized.updatedAt) > Date.parse(existing.updatedAt)) {
        this.run(
          "INSERT OR REPLACE INTO group_overrides (group_key, updated_at, payload) VALUES (?, ?, ?)",
          [normalized.groupKey, normalized.updatedAt, JSON.stringify(normalized)],
        );
      }
    }
  }

  private migrateRequirementScopes(): void {
    for (const row of this.all<{ id: string; payload: string }>("SELECT id, payload FROM requirements")) {
      const raw = JSON.parse(row.payload) as Requirement & { scope?: Requirement["scope"] };
      const normalized = normalizeRequirement(raw);
      if (raw.scope === normalized.scope) continue;
      this.run("UPDATE requirements SET payload = ? WHERE id = ?", [JSON.stringify(normalized), row.id]);
    }
  }

  private migrateHistoricalLaunches(): void {
    const now = new Date().toISOString();
    const features = this.payloads<TieringFeature>("tiering_features", "sort_order, id");
    const featuresById = new Map(features.map((feature) => [feature.id, feature]));
    for (const feature of features) {
      if (normalizeHistoricalLaunchMetadata(feature.historicalLaunch)) continue;
      let historicalLaunch;
      try {
        historicalLaunch = historicalLaunchForFeatureId(feature.id);
      } catch {
        continue;
      }
      const updated = { ...feature, historicalLaunch, updatedAt: now };
      this.run("UPDATE tiering_features SET updated_at = ?, payload = ? WHERE id = ?", [now, JSON.stringify(updated), feature.id]);
      featuresById.set(feature.id, updated);
    }
    for (const row of this.all<{ id: string; payload: string }>("SELECT id, payload FROM requirements")) {
      const raw = JSON.parse(row.payload) as Requirement;
      if (!raw.tieringFeatureId || raw.historicalLaunch || raw.targetMonth !== "2026-09") continue;
      const requirement = normalizeRequirement(raw);
      const feature = featuresById.get(raw.tieringFeatureId);
      if (!feature || !isHistoricalTieringFeature(feature)) continue;
      const updated = {
        ...migrateArchivedHistoricalRequirement({ ...requirement, targetMonth: "2026-09", historicalLaunch: undefined }, feature),
        updatedAt: now,
      };
      this.run("UPDATE requirements SET updated_at = ?, payload = ? WHERE id = ?", [now, JSON.stringify(updated), row.id]);
    }
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

  private payloads<T>(table: string, orderBy: string): T[] {
    return this.all<{ payload: string }>(`SELECT payload FROM ${table} ORDER BY ${orderBy}`).map((row) => JSON.parse(row.payload) as T);
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
  if (!input.domainL0Id) throw new Error("请选择领域 L0");
  if (!input.domainId) throw new Error("请选择领域 L1");
  if (!REQUIREMENT_SCOPES.includes(input.scope)) throw new Error("请选择有效的适用范围");
  if (typeof input.allowLatestVersionUpgrade !== "boolean") throw new Error("请选择是否允许最近版本升级");
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(input.targetMonth)) throw new Error("上线年月格式应为 YYYY-MM");
  const workloadParts = [input.deviceWorkloadPm, input.appWorkloadPm, input.cloudWorkloadPm];
  if (workloadParts.some((value) => !Number.isFinite(value) || value < 0)) throw new Error("各侧工作量必须是大于或等于0的人月数");
  if (input.description.length > 5000) throw new Error("需求描述不能超过5000字");
  if (input.images.length > MAX_REQUIREMENT_IMAGES) throw new Error(`每条需求最多添加${MAX_REQUIREMENT_IMAGES}张图片`);
  for (const image of input.images) {
    if (!REQUIREMENT_IMAGE_MIME_TYPES.includes(image.mimeType)) throw new Error(`不支持图片格式：${image.name}`);
    if (image.sizeBytes <= 0 || image.sizeBytes > MAX_REQUIREMENT_IMAGE_BYTES) throw new Error(`图片“${image.name}”不能超过5MB`);
    if (!image.dataUrl.startsWith(`data:${image.mimeType};base64,`)) throw new Error(`图片“${image.name}”数据无效`);
  }
  if (input.category === "体验优化" && input.productIds.length > 0) throw new Error("体验优化需求不匹配产品");
  if (input.category === "产品专属" && input.productIds.length !== 1) throw new Error("产品专属需求必须且只能选择一个产品");
  if (input.category === "平台演进" && input.productIds.length > 1) throw new Error("平台演进需求最多匹配一个产品");
}
