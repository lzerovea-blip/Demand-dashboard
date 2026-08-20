const EDITABLE_WORKLOAD = /^\d*(?:\.\d*)?$/;
const COMPLETE_WORKLOAD = /^(?:\d+|\d*\.\d+)$/;

export interface WorkloadBreakdown {
  device: number;
  app: number;
  cloud: number;
  unallocated: number;
  total: number;
}

export type WorkloadSide = "device" | "app" | "cloud";

export function normalizeWorkloadEdit(raw: string): string | null {
  const normalized = raw.replace(/[。．，,]/g, ".");
  return EDITABLE_WORKLOAD.test(normalized) ? normalized : null;
}

export function parseWorkloadPm(raw: string): number | null {
  const normalized = raw.trim().replace(/[。．，,]/g, ".");
  if (!COMPLETE_WORKLOAD.test(normalized)) return null;
  const value = Number(normalized);
  return Number.isFinite(value) && value > 0 ? value : null;
}

export function parseWorkloadPart(raw: string): number | null {
  const normalized = raw.trim().replace(/[。．，,]/g, ".");
  if (!normalized) return 0;
  if (!COMPLETE_WORKLOAD.test(normalized)) return null;
  const value = Number(normalized);
  return Number.isFinite(value) && value >= 0 ? value : null;
}

export function roundWorkload(value: number): number {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

export function workloadBreakdownOf(value: {
  deviceWorkloadPm: number;
  appWorkloadPm: number;
  cloudWorkloadPm: number;
  unallocatedWorkloadPm: number;
}): WorkloadBreakdown {
  const device = roundWorkload(value.deviceWorkloadPm);
  const app = roundWorkload(value.appWorkloadPm);
  const cloud = roundWorkload(value.cloudWorkloadPm);
  const unallocated = roundWorkload(value.unallocatedWorkloadPm);
  return { device, app, cloud, unallocated, total: roundWorkload(device + app + cloud + unallocated) };
}

export function sumWorkloadBreakdown(values: Array<{
  deviceWorkloadPm: number;
  appWorkloadPm: number;
  cloudWorkloadPm: number;
  unallocatedWorkloadPm: number;
}>): WorkloadBreakdown {
  return workloadBreakdownOf({
    deviceWorkloadPm: values.reduce((sum, item) => sum + item.deviceWorkloadPm, 0),
    appWorkloadPm: values.reduce((sum, item) => sum + item.appWorkloadPm, 0),
    cloudWorkloadPm: values.reduce((sum, item) => sum + item.cloudWorkloadPm, 0),
    unallocatedWorkloadPm: values.reduce((sum, item) => sum + item.unallocatedWorkloadPm, 0),
  });
}
