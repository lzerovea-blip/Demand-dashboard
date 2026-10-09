import type { Requirement } from "./types.js";
import { isMedicalSubmissionSource } from "./medicalSubmission.js";

export function compareMedicalRoadmapRequirements(a: Requirement, b: Requirement): number {
  const aOrder = a.medicalRoadmapSortOrder ?? Number.MAX_SAFE_INTEGER;
  const bOrder = b.medicalRoadmapSortOrder ?? Number.MAX_SAFE_INTEGER;
  return aOrder - bOrder
    || Number(!a.targetMonth) - Number(!b.targetMonth)
    || a.targetMonth.localeCompare(b.targetMonth)
    || (a.source ?? "").localeCompare(b.source ?? "", "zh-CN")
    || a.title.localeCompare(b.title, "zh-CN")
    || a.id.localeCompare(b.id);
}

/** Reorder visible rows in their global slots, leaving filtered-out rows in place. */
export function reorderedMedicalRequirements(requirements: Requirement[], orderedIds: string[]): Requirement[] {
  if (!Array.isArray(orderedIds) || orderedIds.length < 2 || orderedIds.some((id) => typeof id !== "string" || !id.trim()) || new Set(orderedIds).size !== orderedIds.length) {
    throw new Error("医疗泳道顺序须包含至少两条不重复的需求");
  }
  const rows = requirements.filter((item) => isMedicalSubmissionSource(item.source)).sort(compareMedicalRoadmapRequirements);
  const byId = new Map(rows.map((item) => [item.id, item]));
  if (orderedIds.some((id) => !byId.has(id))) throw new Error("医疗需求集合已变化，请刷新路标后重试");
  const visible = new Set(orderedIds);
  let cursor = 0;
  return rows.map((item) => visible.has(item.id) ? byId.get(orderedIds[cursor++])! : item);
}

export function moveMedicalRoadmapRow(ids: string[], draggedId: string, targetId: string, placement: "before" | "after"): string[] {
  if (draggedId === targetId || !ids.includes(draggedId) || !ids.includes(targetId)) return ids;
  const next = ids.filter((id) => id !== draggedId);
  next.splice(next.indexOf(targetId) + Number(placement === "after"), 0, draggedId);
  return next;
}
