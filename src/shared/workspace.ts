import type { DictionaryItem, GroupOverride, LegacyRequirementDecision, Member, PptAccessPolicy, Requirement, RequirementTombstone, WorkspaceData } from "./types.js";
import { normalizeLegacyGroupOverride, normalizeRequirement } from "./requirements.js";
import { EMPTY_PPT_ACCESS_POLICY, normalizeMember, normalizePptAccessPolicy, REQUIREMENT_OWNER_FIELDS, backfillRequirementOwners, validateRequirementOwners } from "./members.js";
import { EMPTY_TIERING_DATA, mergeTieringData, normalizeTieringData } from "./tiering.js";
import { mergeRoadmapManagement } from "./roadmapManagement.js";

export interface RequirementMergeResolution {
  id: string;
  title: string;
  outcome: "restored" | "ignored" | "ambiguous";
  requirementUpdatedAt: string;
  deletedAt?: string;
}

export interface WorkspaceMergeOptions {
  legacyRequirementDecisions?: Record<string, LegacyRequirementDecision>;
  deferAmbiguousRequirements?: boolean;
  decisionAt?: string;
}

export function mergeWorkspaceData(local: WorkspaceData, incoming: WorkspaceData, options: WorkspaceMergeOptions = {}): WorkspaceData {
  return mergeWorkspaceDataWithReport(local, incoming, options).data;
}

export function mergeWorkspaceDataWithReport(local: WorkspaceData, incoming: WorkspaceData, options: WorkspaceMergeOptions = {}): { data: WorkspaceData; requirementResolutions: RequirementMergeResolution[] } {
  const domainMerge = mergeDomainDictionaries(local.domains, incoming.domains);
  const productMerge = mergeDictionaries(local.products, incoming.products);
  const memberMerge = mergeMembers(local.members ?? [], incoming.members ?? [], domainMerge.idMap);

  const remappedRequirements = incoming.requirements.map((item) => {
    const normalized = normalizeRequirement(item);
    const existing = local.requirements.find((localItem) => localItem.id === item.id);
    // Missing legacy fields are not an instruction to clear locally entered work.
    normalized.researchAppWorkloadPm = item.researchAppWorkloadPm ?? existing?.researchAppWorkloadPm ?? 0;
    normalized.algorithmWorkloadPm = item.algorithmWorkloadPm ?? existing?.algorithmWorkloadPm ?? 0;
    normalized.medicalRoadmapSortOrder = item.medicalRoadmapSortOrder ?? existing?.medicalRoadmapSortOrder;
    for (const { key } of REQUIREMENT_OWNER_FIELDS) {
      normalized[key] = item[key] === undefined
        ? existing?.[key]
        : (memberMerge.idMap.get(item[key]!) ?? item[key]);
    }
    const domainId = domainMerge.idMap.get(normalized.domainId) ?? normalized.domainId;
    const mergedDomainL1 = domainMerge.items.find((domain) => domain.id === domainId);
    return {
      ...normalizeRequirement(normalized),
      domainL0Id: mergedDomainL1?.parentId ?? domainMerge.idMap.get(normalized.domainL0Id) ?? normalized.domainL0Id,
      domainId,
      productIds: [...new Set(normalized.productIds.map((id) => productMerge.idMap.get(id) ?? id))],
    };
  });

  const requirementMerge = mergeRequirementsWithTombstones(local, incoming, remappedRequirements, options);
  const requirements = requirementMerge.requirements
    .map((item) => backfillRequirementOwners(item, memberMerge.items));
  for (const item of requirements) validateRequirementOwners(item, memberMerge.items);
  const remappedOverrides = incoming.groupOverrides.map((item) => {
    const normalized = normalizeLegacyGroupOverride(item);
    return {
      ...normalized,
      groupKey: remapGroupKey(normalized.groupKey, domainMerge.idMap),
    };
  });
  const localPolicy = normalizePptAccessPolicy(local.pptAccessPolicy);
  const incomingPolicy = normalizePptAccessPolicy(incoming.pptAccessPolicy);
  const pptAccessPolicy = timestamp(incomingPolicy.updatedAt) > timestamp(localPolicy.updatedAt)
    ? remapPolicy(incomingPolicy, memberMerge.idMap)
    : localPolicy;

  const data: WorkspaceData = {
    requirements,
    requirementTombstones: requirementMerge.tombstones,
    requirementDeletionTrackingStartedAt: laterTimestamp(local.requirementDeletionTrackingStartedAt, incoming.requirementDeletionTrackingStartedAt),
    domains: domainMerge.items,
    products: productMerge.items,
    groupOverrides: mergeByUpdatedAt(local.groupOverrides, remappedOverrides, (item) => item.groupKey),
    members: memberMerge.items,
    pptAccessPolicy,
    tiering: mergeTieringData(normalizeTieringData(local.tiering), normalizeTieringData(incoming.tiering)),
  };
  return { data: mergeRoadmapManagement(local, incoming, data, domainMerge.idMap), requirementResolutions: requirementMerge.resolutions };
}

function mergeRequirementsWithTombstones(
  local: WorkspaceData,
  incoming: WorkspaceData,
  remappedIncoming: Requirement[],
  options: WorkspaceMergeOptions,
): { requirements: Requirement[]; tombstones: RequirementTombstone[]; resolutions: RequirementMergeResolution[] } {
  const localRequirements = local.requirements.map(normalizeRequirement);
  const localById = new Map(localRequirements.map((item) => [item.id, item]));
  const localTombstones = new Map((local.requirementTombstones ?? []).map((item) => [item.id, item]));
  const incomingTombstones = new Map((incoming.requirementTombstones ?? []).map((item) => [item.id, item]));
  const resolutions: RequirementMergeResolution[] = [];
  const incomingAllowed: Requirement[] = [];
  const trackingStartedAt = timestamp(local.requirementDeletionTrackingStartedAt ?? "");

  for (const requirement of remappedIncoming) {
    const isAmbiguous = !localById.has(requirement.id)
      && !localTombstones.has(requirement.id)
      && trackingStartedAt > 0
      && timestamp(requirement.updatedAt) < trackingStartedAt;
    if (!isAmbiguous) {
      incomingAllowed.push(requirement);
      continue;
    }
    const decision = options.legacyRequirementDecisions?.[requirement.id];
    if (decision === "merge" || (!options.deferAmbiguousRequirements && !decision)) {
      incomingAllowed.push(requirement);
      continue;
    }
    resolutions.push({ id: requirement.id, title: requirement.title, outcome: "ambiguous", requirementUpdatedAt: requirement.updatedAt });
    if (decision === "keep-deleted") {
      const deletedAt = options.decisionAt ?? new Date().toISOString();
      incomingTombstones.set(requirement.id, { id: requirement.id, title: requirement.title, deletedAt });
    }
  }

  const active = new Map(localRequirements.map((item) => [item.id, item]));
  for (const requirement of incomingAllowed) {
    const existing = active.get(requirement.id);
    if (!existing || timestamp(requirement.updatedAt) > timestamp(existing.updatedAt)) active.set(requirement.id, requirement);
  }
  const tombstones = new Map(localTombstones);
  for (const tombstone of incomingTombstones.values()) {
    const existing = tombstones.get(tombstone.id);
    if (!existing || timestamp(tombstone.deletedAt) > timestamp(existing.deletedAt)) tombstones.set(tombstone.id, tombstone);
  }

  const keptRequirements: Requirement[] = [];
  const keptTombstones: RequirementTombstone[] = [];
  for (const id of new Set([...active.keys(), ...tombstones.keys()])) {
    const requirement = active.get(id);
    const tombstone = tombstones.get(id);
    if (tombstone && (!requirement || timestamp(tombstone.deletedAt) >= timestamp(requirement.updatedAt))) {
      keptTombstones.push(tombstone);
      const incomingRequirement = incomingAllowed.find((item) => item.id === id);
      if (incomingRequirement && localTombstones.has(id)) {
        resolutions.push({ id, title: tombstone.title || incomingRequirement.title, outcome: "ignored", requirementUpdatedAt: incomingRequirement.updatedAt, deletedAt: tombstone.deletedAt });
      }
      continue;
    }
    if (requirement) {
      keptRequirements.push(requirement);
      if (localTombstones.has(id) && incomingAllowed.some((item) => item.id === id)) {
        resolutions.push({ id, title: requirement.title, outcome: "restored", requirementUpdatedAt: requirement.updatedAt, deletedAt: localTombstones.get(id)?.deletedAt });
      }
    }
  }
  return { requirements: keptRequirements, tombstones: keptTombstones, resolutions };
}

function laterTimestamp(left?: string, right?: string): string | undefined {
  if (!left) return right;
  if (!right) return left;
  return timestamp(left) >= timestamp(right) ? left : right;
}

function mergeMembers(
  local: Member[],
  incoming: Member[],
  domainIds: ReadonlyMap<string, string>,
): { items: Member[]; idMap: Map<string, string> } {
  const items = local.map(normalizeMember);
  const byEmployeeNumber = new Map(items.map((item) => [normalizeName(item.employeeNumber), item]));
  const usedIds = new Set(items.map((item) => item.id));
  const idMap = new Map<string, string>();

  for (const raw of incoming) {
    const normalized = normalizeMember({
      ...raw,
      domainL1Ids: raw.domainL1Ids.map((id) => domainIds.get(id) ?? id),
    });
    const key = normalizeName(normalized.employeeNumber);
    const existing = byEmployeeNumber.get(key);
    if (existing) {
      idMap.set(normalized.id, existing.id);
      if (timestamp(normalized.updatedAt) > timestamp(existing.updatedAt)) {
        Object.assign(existing, normalized, { id: existing.id });
      }
      continue;
    }
    const id = uniqueId(normalized.id, usedIds);
    const item = { ...normalized, id, sortOrder: items.length };
    items.push(item);
    usedIds.add(id);
    byEmployeeNumber.set(key, item);
    idMap.set(normalized.id, id);
  }

  const accounts = new Set<string>();
  for (const member of items) {
    const account = normalizeName(member.microsoftAccount);
    if (accounts.has(account)) throw new Error(`成员账号“${member.microsoftAccount}”重复，请先处理冲突`);
    accounts.add(account);
  }
  return { items, idMap };
}

function remapPolicy(policy: PptAccessPolicy, memberIds: ReadonlyMap<string, string>): PptAccessPolicy {
  return {
    ...policy,
    entries: policy.entries.map((entry) => ({
      ...entry,
      memberId: memberIds.get(entry.memberId) ?? entry.memberId,
    })),
  };
}

function mergeDomainDictionaries(local: DictionaryItem[], incoming: DictionaryItem[]): { items: DictionaryItem[]; idMap: Map<string, string> } {
  const localL0s = local.filter((item) => item.level === "L0").map((item) => ({ ...item, level: "L0" as const, parentId: undefined }));
  const incomingL0s = incoming.filter((item) => item.level === "L0").map((item) => ({ ...item, level: "L0" as const, parentId: undefined }));
  const l0Merge = mergeDictionaries(localL0s, incomingL0s);
  const localL1s = local.filter((item) => (item.level ?? "L1") === "L1").map((item) => ({ ...item, level: "L1" as const }));
  const incomingL1s = incoming
    .filter((item) => (item.level ?? "L1") === "L1")
    .map((item) => ({ ...item, level: "L1" as const, parentId: item.parentId ? (l0Merge.idMap.get(item.parentId) ?? item.parentId) : undefined }));
  const l1Merge = mergeDictionaries(localL1s, incomingL1s);
  return {
    items: [...l0Merge.items, ...l1Merge.items],
    idMap: new Map([...l0Merge.idMap, ...l1Merge.idMap]),
  };
}

function mergeDictionaries(local: DictionaryItem[], incoming: DictionaryItem[]): { items: DictionaryItem[]; idMap: Map<string, string> } {
  const items = local.map((item) => ({ ...item }));
  const byName = new Map(items.map((item) => [dictionaryKey(item), item]));
  const usedIds = new Set(items.map((item) => item.id));
  const idMap = new Map<string, string>();

  for (const incomingItem of incoming) {
    const sameName = byName.get(dictionaryKey(incomingItem));
    if (sameName) {
      idMap.set(incomingItem.id, sameName.id);
      if (!sameName.parentId && incomingItem.parentId) sameName.parentId = incomingItem.parentId;
      continue;
    }
    const id = uniqueId(incomingItem.id, usedIds);
    const item = { ...incomingItem, id, sortOrder: items.length };
    items.push(item);
    usedIds.add(id);
    byName.set(dictionaryKey(item), item);
    idMap.set(incomingItem.id, id);
  }

  return { items, idMap };
}

function mergeByUpdatedAt<T extends { id: string; updatedAt: string }>(local: T[], incoming: T[]): T[];
function mergeByUpdatedAt<T extends { updatedAt: string }>(local: T[], incoming: T[], keyOf: (item: T) => string): T[];
function mergeByUpdatedAt<T extends { updatedAt: string }>(local: T[], incoming: T[], keyOf: (item: T) => string = (item) => (item as T & { id: string }).id): T[] {
  const merged = new Map(local.map((item) => [keyOf(item), item]));
  for (const item of incoming) {
    const key = keyOf(item);
    const existing = merged.get(key);
    if (!existing || timestamp(item.updatedAt) > timestamp(existing.updatedAt)) merged.set(key, item);
  }
  return [...merged.values()];
}

function remapGroupKey(groupKey: string, domainIds: ReadonlyMap<string, string>): string {
  const [domainId, source, targetMonth] = groupKey.split("::");
  if (!domainId || !source || !targetMonth) return groupKey;
  return `${domainIds.get(domainId) ?? domainId}::${source}::${targetMonth}`;
}

function uniqueId(preferred: string, used: ReadonlySet<string>): string {
  if (!used.has(preferred)) return preferred;
  let counter = 1;
  while (used.has(`${preferred}-imported-${counter}`)) counter += 1;
  return `${preferred}-imported-${counter}`;
}

function normalizeName(value: string): string {
  return value.trim().toLocaleLowerCase("zh-CN");
}

function dictionaryKey(item: DictionaryItem): string {
  return `${item.level ?? "L1"}:${normalizeName(item.name)}`;
}

function timestamp(value: string): number {
  const parsed = Date.parse(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function workspaceCounts(data: WorkspaceData) {
  const tiering = normalizeTieringData(data.tiering ?? EMPTY_TIERING_DATA);
  return {
    roadmaps: data.roadmaps?.length ?? 0, roadmapLanes: data.roadmapLanes?.length ?? 0,
    roadmapAssignments: data.roadmapAssignments?.length ?? 0, medicalRoadmapPlans: data.medicalRoadmapPlans?.length ?? 0,
    roadmapAssignmentTombstones: data.roadmapAssignmentTombstones?.length ?? 0,
    requirements: data.requirements.length,
    requirementTombstones: (data.requirementTombstones ?? []).length,
    images: data.requirements.reduce((sum, item) => sum + item.images.length, 0),
    domains: data.domains.length,
    products: data.products.length,
    groupOverrides: data.groupOverrides.length,
    members: (data.members ?? []).length,
    whitelistEntries: normalizePptAccessPolicy(data.pptAccessPolicy ?? EMPTY_PPT_ACCESS_POLICY).entries.length,
    tieringFeatures: tiering.features.length,
    tieringProductSeries: tiering.productSeries.length,
    tieringSupportRecords: tiering.supportRecords.length,
    tieringEvidenceSources: tiering.evidenceSources.length,
    tieringVerificationIssues: tiering.verificationIssues.length,
    tieringChangeLogs: tiering.changeLogs.length,
  };
}
