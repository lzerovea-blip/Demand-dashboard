export interface RectangleLike {
  left: number;
  top: number;
  right: number;
  bottom: number;
  width: number;
  height: number;
}

export interface TimelinePopoverPosition {
  left: number;
  top: number;
  placement: "above" | "below";
}

export function medicalTimelinePopoverPosition(
  anchor: RectangleLike,
  panel: Pick<RectangleLike, "width" | "height">,
  viewportWidth: number,
  viewportHeight: number,
  gap = 8,
  margin = 12,
): TimelinePopoverPosition {
  const maxLeft = Math.max(margin, viewportWidth - panel.width - margin);
  const left = Math.min(Math.max(anchor.left, margin), maxLeft);
  const fitsBelow = anchor.bottom + gap + panel.height <= viewportHeight - margin;
  const fitsAbove = anchor.top - gap - panel.height >= margin;
  const placement = !fitsBelow && fitsAbove ? "above" : "below";
  const rawTop = placement === "above" ? anchor.top - gap - panel.height : anchor.bottom + gap;
  const top = Math.min(Math.max(rawTop, margin), Math.max(margin, viewportHeight - panel.height - margin));
  return { left, top, placement };
}
