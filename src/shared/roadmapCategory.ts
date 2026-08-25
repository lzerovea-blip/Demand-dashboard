import { CATEGORIES, CATEGORY_META, type RequirementCategory } from "./types.js";

export interface RoadmapCategoryMark {
  category: RequirementCategory;
  label: string;
  className: string;
  color: string;
}

/**
 * 路标卡片可合并不同分类的需求。分类标识始终按统一顺序输出，
 * 避免只突出其中一个分类而隐藏其他信息。
 */
export function roadmapCategoryMarks(categories: readonly RequirementCategory[]): RoadmapCategoryMark[] {
  const selected = new Set(categories);
  return CATEGORIES
    .filter((category) => selected.has(category))
    .map((category) => ({
      category,
      label: CATEGORY_META[category].roadmapLabel,
      className: CATEGORY_META[category].className,
      color: CATEGORY_META[category].color,
    }));
}
