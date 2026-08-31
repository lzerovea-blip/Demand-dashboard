import { MEMBER_ROLES } from "./types.js";
import type { Member, MemberRole, PptAccessPolicy, PptWhitelistEntry, Requirement } from "./types.js";

export const EMPTY_PPT_ACCESS_POLICY: PptAccessPolicy = {
  entries: [],
  updatedAt: "1970-01-01T00:00:00.000Z",
};

export function normalizeMicrosoftAccount(value: string): string {
  return value.trim().toLocaleLowerCase("en-US");
}

export function isValidMicrosoftAccount(value: string): boolean {
  const account = normalizeMicrosoftAccount(value);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/u.test(account);
}

export function normalizeMember(input: Member): Member {
  return {
    ...input,
    name: input.name.trim(),
    employeeNumber: input.employeeNumber.trim(),
    microsoftAccount: normalizeMicrosoftAccount(input.microsoftAccount),
    role: MEMBER_ROLES.includes(input.role as MemberRole) ? input.role : undefined,
    domainL1Ids: [...new Set(Array.isArray(input.domainL1Ids) ? input.domainL1Ids.filter(Boolean) : [])],
  };
}

export function normalizePptAccessPolicy(value: PptAccessPolicy | undefined): PptAccessPolicy {
  const entries = Array.isArray(value?.entries) ? normalizePptWhitelistEntries(value.entries) : [];
  return {
    entries,
    updatedAt: typeof value?.updatedAt === "string" && !Number.isNaN(Date.parse(value.updatedAt))
      ? value.updatedAt
      : EMPTY_PPT_ACCESS_POLICY.updatedAt,
  };
}

export function normalizePptWhitelistEntries(entries: PptWhitelistEntry[]): PptWhitelistEntry[] {
  const byMember = new Map<string, PptWhitelistEntry>();
  for (const entry of entries) {
    if (!entry?.memberId || !["read", "change"].includes(entry.permission)) continue;
    byMember.set(entry.memberId, { memberId: entry.memberId, permission: entry.permission });
  }
  return [...byMember.values()];
}

export function responsibleMembers(
  members: readonly Member[],
  domainL1Id: string,
  role: MemberRole,
): Member[] {
  return members
    .filter((member) => member.role === role && member.domainL1Ids.includes(domainL1Id))
    .sort((a, b) => a.sortOrder - b.sortOrder || a.name.localeCompare(b.name, "zh-CN"));
}

export function requirementMatchesResponsibleMember(
  requirement: Pick<Requirement, "domainId">,
  members: readonly Member[],
  role: MemberRole,
  memberId: string,
): boolean {
  if (!memberId) return true;
  return responsibleMembers(members, requirement.domainId, role).some((member) => member.id === memberId);
}

export function visibleMembersForRole(members: readonly Member[], domains: ReadonlySet<string>, role: MemberRole): Member[] {
  return members.filter((member) => member.role === role && member.domainL1Ids.some((id) => domains.has(id)));
}
