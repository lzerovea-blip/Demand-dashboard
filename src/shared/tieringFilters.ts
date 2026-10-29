import type { Requirement, TieringFeature, Member } from "./types.js";
import { requirementMatchesResponsibleMember } from "./members.js";
import { halfYearOf } from "./roadmap.js";

export function tieringFeatureHalfYear(feature: TieringFeature, requirement?: Requirement): string {
  const month = requirement?.targetMonth ?? feature.historicalLaunch?.suggestedTargetMonth;
  return month ? halfYearOf(month) : "";
}

export function tieringMatchesOwnersAndVersion(feature: TieringFeature, requirement: Requirement | undefined, members: readonly Member[], filters: { planningMemberId: string; productManagerMemberId: string; halfYear: string }): boolean {
  return requirementMatchesResponsibleMember(requirement ?? {}, members, "规划", filters.planningMemberId)
    && requirementMatchesResponsibleMember(requirement ?? {}, members, "产品经理", filters.productManagerMemberId)
    && (!filters.halfYear || tieringFeatureHalfYear(feature, requirement) === filters.halfYear);
}
