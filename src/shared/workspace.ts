import type { DictionaryItem, GroupOverride, Requirement, WorkspaceData } from "./types.js";
import { normalizeLegacyGroupOverride, normalizeRequirement } from "./requirements.js";

export function mergeWorkspaceData(local: WorkspaceData, incoming: WorkspaceData): WorkspaceData {
  const domainMerge = mergeDomainDictionaries(local.domains, incoming.domains);
  const productMerge = mergeDictionaries(local.products, incoming.products);

  const remappedRequirements = incoming.requirements.map((item) => {
    const normalized = normalizeRequirement(item);
    const domainId = domainMerge.idMap.get(normalized.domainId) ?? normalized.domainId;
    const mergedDomainL1 = domainMerge.items.find((domain) => domain.id === domainId);
    return {
      ...normalized,
      domainL0Id: mergedDomainL1?.parentId ?? domainMerge.idMap.get(normalized.domainL0Id) ?? normalized.domainL0Id,
      domainId,
      productIds: [...new Set(normalized.productIds.map((id) => productMerge.idMap.get(id) ?? id))],
    };
  });

  const requirements = mergeByUpdatedAt(local.requirements.map((item) => normalizeRequirement(item)), remappedRequirements);
  const remappedOverrides = incoming.groupOverrides.map((item) => {
    const normalized = normalizeLegacyGroupOverride(item);
    return {
      ...normalized,
      groupKey: remapGroupKey(normalized.groupKey, domainMerge.idMap),
    };
  });

  return {
    requirements,
    domains: domainMerge.items,
    products: productMerge.items,
    groupOverrides: mergeByUpdatedAt(local.groupOverrides, remappedOverrides, (item) => item.groupKey),
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
  return {
    requirements: data.requirements.length,
    images: data.requirements.reduce((sum, item) => sum + item.images.length, 0),
    domains: data.domains.length,
    products: data.products.length,
    groupOverrides: data.groupOverrides.length,
  };
}
