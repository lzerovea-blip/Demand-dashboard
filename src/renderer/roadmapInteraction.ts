export interface RoadmapMonthColumnMetrics {
  month: string;
  offsetLeft: number;
  width: number;
}

export interface RoadmapMonthAnchor {
  month: string;
  viewportOffset: number;
}

export function roadmapCardIsL1Related(domainId: string, highlightedDomainId: string | null, enabled: boolean): boolean {
  return enabled && Boolean(highlightedDomainId) && domainId === highlightedDomainId;
}

export function captureRoadmapMonthAnchor(
  scrollLeft: number,
  stickyWidth: number,
  columns: RoadmapMonthColumnMetrics[],
): RoadmapMonthAnchor | undefined {
  if (!columns.length) return undefined;
  const visibleStart = scrollLeft + stickyWidth;
  const column = columns.find((item) => item.offsetLeft + item.width > visibleStart) ?? columns[columns.length - 1];
  return { month: column.month, viewportOffset: column.offsetLeft - scrollLeft };
}

export function roadmapScrollLeftForAnchor(
  anchor: RoadmapMonthAnchor,
  columns: RoadmapMonthColumnMetrics[],
  maxScrollLeft: number,
): number {
  const column = columns.find((item) => item.month === anchor.month);
  if (!column) return Math.max(0, Math.min(maxScrollLeft, 0));
  return Math.max(0, Math.min(maxScrollLeft, column.offsetLeft - anchor.viewportOffset));
}
