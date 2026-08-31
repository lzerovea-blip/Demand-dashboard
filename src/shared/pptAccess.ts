export function addCalendarMonthsClamped(value: Date, months: number): Date {
  if (!Number.isInteger(months)) throw new Error("月份偏移必须是整数");
  const result = new Date(value.getTime());
  const originalDay = result.getDate();
  result.setDate(1);
  result.setMonth(result.getMonth() + months);
  const lastDay = new Date(result.getFullYear(), result.getMonth() + 1, 0).getDate();
  result.setDate(Math.min(originalDay, lastDay));
  return result;
}

export function pptPermissionExpiration(exportedAt: string | Date): Date {
  const date = exportedAt instanceof Date ? new Date(exportedAt.getTime()) : new Date(exportedAt);
  if (Number.isNaN(date.getTime())) throw new Error("导出时间无效");
  return addCalendarMonthsClamped(date, 6);
}

export function formatPptExpirationDate(value: Date): string {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}
