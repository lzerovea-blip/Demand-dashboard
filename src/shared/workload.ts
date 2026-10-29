const EDITABLE_WORKLOAD = /^\d*(?:\.\d*)?$/;
const COMPLETE_WORKLOAD = /^(?:\d+|\d*\.\d+)$/;

export interface WorkloadBreakdown {
  device: number;
  app: number;
  cloud: number;
  researchApp: number;
  algorithm: number;
  unallocated: number;
  total: number;
}

export type WorkloadSide = "device" | "app" | "cloud";

export const WORKLOAD_INPUTS = [
  { key: "device", field: "deviceWorkloadPm", label: "设备工作量" },
  { key: "app", field: "appWorkloadPm", label: "运动健康 App 工作量" },
  { key: "researchApp", field: "researchAppWorkloadPm", label: "创新研究 App 工作量" },
  { key: "cloud", field: "cloudWorkloadPm", label: "云侧工作量" },
  { key: "algorithm", field: "algorithmWorkloadPm", label: "算法工作量" },
] as const;

export interface WorkloadParts {
  deviceWorkloadPm: number;
  appWorkloadPm: number;
  cloudWorkloadPm: number;
  researchAppWorkloadPm?: number;
  algorithmWorkloadPm?: number;
}

export function isWorkloadPendingEvaluation(value: WorkloadParts): boolean {
  return roundWorkload(value.deviceWorkloadPm)
    + roundWorkload(value.appWorkloadPm)
    + roundWorkload(value.cloudWorkloadPm)
    + roundWorkload(value.researchAppWorkloadPm ?? 0)
    + roundWorkload(value.algorithmWorkloadPm ?? 0) === 0;
}

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

export function workloadBreakdownOf(value: WorkloadParts & {
  unallocatedWorkloadPm: number;
}): WorkloadBreakdown {
  const device = roundWorkload(value.deviceWorkloadPm);
  const app = roundWorkload(value.appWorkloadPm);
  const cloud = roundWorkload(value.cloudWorkloadPm);
  const researchApp = roundWorkload(value.researchAppWorkloadPm ?? 0);
  const algorithm = roundWorkload(value.algorithmWorkloadPm ?? 0);
  const unallocated = roundWorkload(value.unallocatedWorkloadPm);
  return { device, app, cloud, researchApp, algorithm, unallocated, total: roundWorkload(device + app + cloud + researchApp + algorithm + unallocated) };
}

export function sumWorkloadBreakdown(values: Array<WorkloadParts & {
  unallocatedWorkloadPm: number;
}>): WorkloadBreakdown {
  return workloadBreakdownOf({
    deviceWorkloadPm: values.reduce((sum, item) => sum + item.deviceWorkloadPm, 0),
    appWorkloadPm: values.reduce((sum, item) => sum + item.appWorkloadPm, 0),
    cloudWorkloadPm: values.reduce((sum, item) => sum + item.cloudWorkloadPm, 0),
    researchAppWorkloadPm: values.reduce((sum, item) => sum + (item.researchAppWorkloadPm ?? 0), 0),
    algorithmWorkloadPm: values.reduce((sum, item) => sum + (item.algorithmWorkloadPm ?? 0), 0),
    unallocatedWorkloadPm: values.reduce((sum, item) => sum + item.unallocatedWorkloadPm, 0),
  });
}
