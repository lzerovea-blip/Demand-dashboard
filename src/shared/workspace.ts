import type { DictionaryItem, GroupOverride, Member, PptAccessPolicy, Requirement, WorkspaceData } from "./types.js";
import { normalizeLegacyGroupOverride, normalizeRequirement } from "./requirements.js";
import { EMPTY_PPT_ACCESS_POLICY, normalizeMember, normalizePptAccessPolicy, REQUIREMENT_OWNER_FIELDS, backfillRequirementOwners, validateRequirementOwners } from "./members.js";
import { EMPTY_TIERING_DATA, mergeTieringData, normalizeTieringData } from "./tiering.js";

export function mergeWorkspaceData(local: WorkspaceData, incoming: WorkspaceData): WorkspaceData {
  const domainMerge = mergeDomainDictionaries(local.domains, incoming.domains);
  const productMerge = mergeDictionaries(local.products, incoming.products);
  const memberMerge = mergeMembers(local.members ?? [], incoming.members ?? [], domainMerge.idMap);

  const remappedRequirements = incoming.requirements.map((item) => {
    const normalized = normalizeRequirement(item);
    const existing = local.requirements.find((localItem) => localItem.id === item.id);
    for (const { key } of REQUIREMENT_OWNER_FIELDS) {
      normalized[key] = item[key] === undefined
        ? existing?.[key]
        : (memberMerge.idMap.get(item[key]!) ?? item[key]);
    }
    const domainId = domainMerge.idMap.get(normalized.domainId) ?? normalized.domainId;
    const mergedDomainL1 = domainMerge.items.find((domain) => domain.id === domainId);
    return {
      ...normalized,
      domainL0Id: mergedDomainL1?.parentId ?? domainMerge.idMap.get(normalized.domainL0Id) ?? normalized.domainL0Id,
      domainId,
      productIds: [...new Set(normalized.productIds.map((id) => productMerge.idMap.get(id) ?? id))],
    };
  });

  const requirements = mergeByUpdatedAt(local.requirements.map((item) => normalizeRequirement(item)), remappedRequirements)
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

  return {
    requirements,
    domains: domainMerge.items,
    products: productMerge.items,
    groupOverrides: mergeByUpdatedAt(local.groupOverrides, remappedOverrides, (item) => item.groupKey),
    members: memberMerge.items,
    pptAccessPolicy,
    tiering: mergeTieringData(normalizeTieringData(local.tiering), normalizeTieringData(incoming.tiering)),
  };
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
    requirements: data.requirements.length,
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
