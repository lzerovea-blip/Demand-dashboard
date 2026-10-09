import { SOURCES, ROADMAP_TRACKS } from "./types.js";
import type { AppSnapshot, WorkspaceData, RoadmapDefinition, RoadmapLaneDefinition, RequirementRoadmapAssignment, MedicalRoadmapPlan, Requirement } from "./types.js";
import { isValidRequirementMonth, isMedicalSubmissionSource } from "./medicalSubmission.js";

const SEED_TIME = "1970-01-01T00:00:00.000Z";
export const MEDICAL_PLAN_MONTH_KEYS = ["registrationStartMonth", "submissionMonth", "clinicalStartMonth", "clinicalCompleteMonth", "certificateMonth"] as const;
export const MEDICAL_PLAN_LABELS = ["注册启动时间", "送检时间", "临床启动时间", "临床完成时间", "获证时间"] as const;

export function legacyRoadmapId(source: string): string {
  if (source.startsWith("运动")) return "运动";
  if (source.startsWith("健康")) return "健康";
  if (source.startsWith("AI")) return "AI";
  if (isMedicalSubmissionSource(source)) return "医疗认证";
  return source;
}

export function defaultRoadmapDefinitions(): RoadmapDefinition[] {
  return ROADMAP_TRACKS.map((id, sortOrder) => ({ id, name: `${id}路标`, kind: id === "医疗认证" ? "medical" : "standard", sortOrder, updatedAt: SEED_TIME }));
}
export function defaultRoadmapLanes(): RoadmapLaneDefinition[] {
  return SOURCES.map((id) => {
    const roadmapId = legacyRoadmapId(id);
    const name = roadmapId === "运动" || roadmapId === "健康" ? id.slice(2) : roadmapId === "AI" || roadmapId === "医疗认证" ? id : "需求";
    const sortOrder = name === "高阶" ? 0 : name === "进阶" ? 1 : name === "基础" ? 2 : SOURCES.filter((source) => legacyRoadmapId(source) === roadmapId).indexOf(id);
    return { id, roadmapId, name, sortOrder, updatedAt: SEED_TIME };
  });
}

export function normalizeRoadmapManagement<T extends Pick<AppSnapshot, "requirements" | "groupOverrides"> & Partial<AppSnapshot>>(input: T): Omit<T, "roadmaps" | "roadmapLanes" | "roadmapAssignments" | "roadmapAssignmentTombstones" | "medicalRoadmapPlans"> & Required<Pick<AppSnapshot, "roadmaps" | "roadmapLanes" | "roadmapAssignments" | "roadmapAssignmentTombstones" | "medicalRoadmapPlans">> {
  const legacy = input.roadmapAssignments === undefined;
  const roadmaps = (input.roadmaps ?? defaultRoadmapDefinitions()).map((item) => ({ ...item })).sort(order);
  const roadmapLanes = (input.roadmapLanes ?? defaultRoadmapLanes()).map((item) => ({ ...item })).sort(order);
  const roadmapAssignments = legacy ? input.requirements.filter((item) => item.source && SOURCES.includes(item.source)).map((item) => ({
    requirementId: item.id, roadmapId: legacyRoadmapId(item.source!), laneId: item.source!, updatedAt: item.updatedAt,
  })) : input.roadmapAssignments!.map((item) => ({ ...item }));
  const medicalRoadmapPlans = input.medicalRoadmapPlans?.map((item) => ({ ...item })) ?? (legacy ? input.requirements.filter((item) => isMedicalSubmissionSource(item.source ?? "")).map((item) => ({
    requirementId: item.id, roadmapId: "医疗认证", ...item.medicalSubmissionMilestones, submissionMonth: item.targetMonth || undefined,
    sortOrder: item.medicalRoadmapSortOrder, updatedAt: item.updatedAt,
  })) : []);
  for (const record of [...roadmapAssignments, ...medicalRoadmapPlans]) { const map = roadmaps.find((item) => item.id === record.roadmapId); if (map && !map.usedAt) map.usedAt = record.updatedAt; }
  const groupOverrides = legacy ? input.groupOverrides.map((item) => {
    const parts = item.groupKey.split("::");
    return parts.length === 3 && SOURCES.includes(parts[1] as never)
      ? { ...item, groupKey: assignmentGroupKey(legacyRoadmapId(parts[1]), parts[1], parts[0], parts[2]) } : { ...item };
  }) : input.groupOverrides;
  const liveIds = new Set(input.requirements.map((item) => item.id));
  return { ...input, groupOverrides, roadmaps, roadmapLanes, roadmapAssignments: roadmapAssignments.filter((item) => liveIds.has(item.requirementId)), medicalRoadmapPlans: medicalRoadmapPlans.filter((item) => liveIds.has(item.requirementId)), roadmapAssignmentTombstones: input.roadmapAssignmentTombstones?.map((item) => ({ ...item })) ?? [] };
}

export function assignmentGroupKey(roadmapId: string, laneId: string, domainId: string, month: string): string {
  return `${roadmapId}::${laneId}::${domainId}::${month}`;
}
export function roadmapDefinitionsFor(snapshot: Partial<AppSnapshot>): RoadmapDefinition[] {
  return (snapshot.roadmaps ?? defaultRoadmapDefinitions()).slice().sort(order);
}
export function roadmapLanesFor(snapshot: Partial<AppSnapshot>, roadmapId: string): RoadmapLaneDefinition[] {
  return (snapshot.roadmapLanes ?? defaultRoadmapLanes()).filter((item) => item.roadmapId === roadmapId).sort(order);
}
export function projectedMedicalRequirement(requirement: Requirement, plan?: MedicalRoadmapPlan, laneId?: string): Requirement {
  return { ...requirement, source: laneId === "CE认证" ? "CE认证" : "NMPA认证", targetMonth: plan?.submissionMonth ?? "", medicalSubmissionMilestones: plan,
    medicalRoadmapSortOrder: plan?.sortOrder };
}

export function validateRoadmapManagement(input: Partial<AppSnapshot> & Pick<AppSnapshot, "requirements">): void {
  if (input.roadmapAssignments === undefined) return;
  const roadmaps = input.roadmaps ?? [];
  const lanes = input.roadmapLanes ?? [];
  const requirements = new Set(input.requirements.map((item) => item.id));
  unique(roadmaps, (item) => item.id, "路标 ID"); unique(roadmaps, (item) => item.name.trim().toLocaleLowerCase(), "路标名称");
  unique(lanes, (item) => item.id, "泳道 ID"); unique(lanes, (item) => `${item.roadmapId}:${item.name.trim().toLocaleLowerCase()}`, "泳道名称");
  unique(input.roadmapAssignments, (item) => item.requirementId, "需求路标匹配");
  for (const item of roadmaps) if (!item.id || !item.name.trim() || !["standard", "medical"].includes(item.kind) || !validTime(item.updatedAt) || !validOrder(item.sortOrder)) throw new Error("路标字典无效");
  for (const item of roadmaps) if (item.usedAt && !validTime(item.usedAt)) throw new Error("路标使用记录无效");
  for (const item of lanes) if (!item.id || !item.name.trim() || !roadmaps.some((map) => map.id === item.roadmapId) || !validTime(item.updatedAt) || !validOrder(item.sortOrder)) throw new Error("泳道字典无效");
  for (const item of input.roadmapAssignments) {
    if (!requirements.has(item.requirementId) || !lanes.some((lane) => lane.id === item.laneId && lane.roadmapId === item.roadmapId) || !validTime(item.updatedAt)) throw new Error("需求的路标或泳道匹配无效");
  }
  unique(input.medicalRoadmapPlans ?? [], (item) => `${item.roadmapId}:${item.requirementId}`, "医疗路标配置");
  for (const plan of input.medicalRoadmapPlans ?? []) {
    if (!requirements.has(plan.requirementId) || !roadmaps.some((item) => item.id === plan.roadmapId && item.kind === "medical") || !validTime(plan.updatedAt)) throw new Error("医疗路标配置引用无效");
    if (plan.sortOrder !== undefined && !validOrder(plan.sortOrder)) throw new Error("医疗泳道顺序无效");
    for (const key of MEDICAL_PLAN_MONTH_KEYS) if (plan[key] && !isValidRequirementMonth(plan[key])) throw new Error(`${MEDICAL_PLAN_LABELS[MEDICAL_PLAN_MONTH_KEYS.indexOf(key)]}必须是有效年月`);
  }
  unique(input.roadmapAssignmentTombstones ?? [], (item) => item.requirementId, "匹配删除记录");
  for (const item of input.roadmapAssignmentTombstones ?? []) if (!item.requirementId || !validTime(item.removedAt)) throw new Error("匹配删除记录无效");
}

export function mergeRoadmapManagement(local: WorkspaceData, incoming: WorkspaceData, merged: WorkspaceData, domainMap: ReadonlyMap<string, string>): WorkspaceData {
  const left = normalizeRoadmapManagement(local);
  const right = normalizeRoadmapManagement(incoming);
  const roadmaps = left.roadmaps.map((item) => ({ ...item }));
  const roadmapIds = new Map<string, string>();
  for (const item of right.roadmaps) {
    const match = roadmaps.find((candidate) => candidate.id === item.id || candidate.name.trim().toLocaleLowerCase() === item.name.trim().toLocaleLowerCase());
    if (match) {
      if (match.kind !== item.kind) throw new Error(`同名路标“${item.name}”展示类型冲突，请先统一字典`);
      roadmapIds.set(item.id, match.id);
      if (Date.parse(item.updatedAt) > Date.parse(match.updatedAt)) Object.assign(match, item, { id: match.id, usedAt: match.usedAt ?? item.usedAt });
    } else { roadmaps.push({ ...item }); roadmapIds.set(item.id, item.id); }
  }
  const roadmapLanes = left.roadmapLanes.map((item) => ({ ...item }));
  const laneIds = new Map<string, string>();
  for (const raw of right.roadmapLanes) {
    const item = { ...raw, roadmapId: roadmapIds.get(raw.roadmapId) ?? raw.roadmapId };
    const match = roadmapLanes.find((candidate) => candidate.id === item.id || (candidate.roadmapId === item.roadmapId && candidate.name.trim().toLocaleLowerCase() === item.name.trim().toLocaleLowerCase()));
    if (match) {
      laneIds.set(item.id, match.id);
      if (Date.parse(item.updatedAt) > Date.parse(match.updatedAt)) Object.assign(match, item, { id: match.id });
    } else { roadmapLanes.push(item); laneIds.set(item.id, item.id); }
  }
  const activeIds = new Set(merged.requirements.map((item) => item.id));
  const assignments = new Map(left.roadmapAssignments.map((item) => [item.requirementId, { ...item }]));
  const removals = new Map(left.roadmapAssignmentTombstones.map((item) => [item.requirementId, { ...item }]));
  const legacyInput = incoming.legacyRoadmapInput || incoming.roadmapAssignments === undefined;
  for (const item of right.roadmapAssignments) {
    if (legacyInput && (local.requirements.some((req) => req.id === item.requirementId) || removals.has(item.requirementId))) continue;
    const old = assignments.get(item.requirementId);
    if (!old || Date.parse(item.updatedAt) > Date.parse(old.updatedAt)) assignments.set(item.requirementId, { ...item, roadmapId: roadmapIds.get(item.roadmapId) ?? item.roadmapId, laneId: laneIds.get(item.laneId) ?? item.laneId });
  }
  for (const item of right.roadmapAssignmentTombstones) {
    const old = removals.get(item.requirementId);
    if (!old || Date.parse(item.removedAt) > Date.parse(old.removedAt)) removals.set(item.requirementId, { ...item });
  }
  for (const [id, assignment] of assignments) if (!activeIds.has(id) || (removals.has(id) && Date.parse(removals.get(id)!.removedAt) >= Date.parse(assignment.updatedAt))) assignments.delete(id);
  for (const [id, assignment] of assignments) if (removals.has(id) && Date.parse(assignment.updatedAt) > Date.parse(removals.get(id)!.removedAt)) removals.delete(id);
  const plans = new Map(left.medicalRoadmapPlans.map((item) => [`${item.roadmapId}:${item.requirementId}`, item]));
  for (const item of right.medicalRoadmapPlans) {
    if (legacyInput && local.requirements.some((req) => req.id === item.requirementId)) continue;
    const next = { ...item, roadmapId: roadmapIds.get(item.roadmapId) ?? item.roadmapId };
    const key = `${next.roadmapId}:${next.requirementId}`;
    if (!plans.has(key) || Date.parse(next.updatedAt) > Date.parse(plans.get(key)!.updatedAt)) plans.set(key, next);
  }
  const remapOverride = (item: WorkspaceData["groupOverrides"][number]) => {
    const parts = item.groupKey.split("::");
    if (parts.length !== 4) return item;
    return { ...item, groupKey: assignmentGroupKey(roadmapIds.get(parts[0]) ?? parts[0], laneIds.get(parts[1]) ?? parts[1], domainMap.get(parts[2]) ?? parts[2], parts[3]) };
  };
  const overrides = new Map(left.groupOverrides.map((item) => [item.groupKey, item]));
  for (const raw of right.groupOverrides) {
    const item = remapOverride(raw);
    if (!overrides.has(item.groupKey) || Date.parse(item.updatedAt) > Date.parse(overrides.get(item.groupKey)!.updatedAt)) overrides.set(item.groupKey, item);
  }
  return { ...merged, roadmaps, roadmapLanes, roadmapAssignments: [...assignments.values()], roadmapAssignmentTombstones: [...removals.values()], medicalRoadmapPlans: [...plans.values()].filter((item) => activeIds.has(item.requirementId)), groupOverrides: [...overrides.values()] };
}
function order(a: { sortOrder: number; id: string }, b: { sortOrder: number; id: string }): number { return a.sortOrder - b.sortOrder || a.id.localeCompare(b.id); }
function unique<T>(items: readonly T[], key: (item: T) => string, label: string): void { const ids = new Set<string>(); for (const item of items) { const id = key(item); if (ids.has(id)) throw new Error(`${label}重复`); ids.add(id); } }
function validTime(value: string): boolean { return Boolean(value) && Number.isFinite(Date.parse(value)); }
function validOrder(value: number): boolean { return Number.isSafeInteger(value) && value >= 0; }
