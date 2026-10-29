import type { WorkspaceWorkbookFieldChange } from "./types.js";

/** Bulk selection mirrors the single-row rule: deleting a demand supersedes its edits. */
export function allWorkbookChangeIds(changes: readonly WorkspaceWorkbookFieldChange[]): string[] {
  const deletedRequirements = new Set(changes.filter((change) => change.kind === "delete").map((change) => change.requirementId));
  return changes.filter((change) => change.kind === "delete" || !deletedRequirements.has(change.requirementId)).map((change) => change.id);
}

export function adoptedWorkbookDeleteCount(
  changes: readonly WorkspaceWorkbookFieldChange[],
  selectedIds: readonly string[],
  conflictDecisions: Readonly<Record<string, "local" | "excel">>,
): number {
  const selected = new Set(selectedIds);
  return changes.filter((change) => change.kind === "delete" && selected.has(change.id) && (!change.conflict || conflictDecisions[change.id] === "excel")).length;
}
