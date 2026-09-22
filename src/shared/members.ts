import { MEMBER_ROLES } from "./types.js";
import type { DictionaryItem, Member, MemberRole, PptAccessPolicy, PptWhitelistEntry, Requirement } from "./types.js";

export const UNASSIGNED_MEMBER = "__unassigned__";

export const REQUIREMENT_OWNER_FIELDS = [
  { key: "productManagerMemberId", role: "产品经理", required: true },
  { key: "planningMemberId", role: "规划", required: true },
  { key: "seMemberId", role: "SE", required: true },
  { key: "algorithmMemberId", role: "算法", required: false },
] as const;
export type RequirementOwnerKey = (typeof REQUIREMENT_OWNER_FIELDS)[number]["key"];
export type RequirementOwners = Pick<Requirement, RequirementOwnerKey>;

export function ownerField(role: MemberRole): RequirementOwnerKey | undefined {
  return REQUIREMENT_OWNER_FIELDS.find((item) => item.role === role)?.key;
}

export function savedResponsibleMembers(requirement: RequirementOwners, members: readonly Member[], role: MemberRole): Member[] {
  const key = ownerField(role);
  return key ? members.filter((member) => member.id === requirement[key]) : [];
}

/** Only absent legacy fields are inferred; an explicitly empty assignment remains empty. */
export function backfillRequirementOwners<T extends RequirementOwners & { domainId: string }>(requirement: T, members: readonly Member[]): T {
  const result = { ...requirement };
  for (const { key, role } of REQUIREMENT_OWNER_FIELDS) {
    if (result[key] !== undefined) continue;
    const matches = responsibleMembers(members, result.domainId, role);
    result[key] = matches.length === 1 ? matches[0].id : "";
  }
  return result;
}

export function validateRequirementOwners(requirement: RequirementOwners, members: readonly Member[], required = false): void {
  for (const { key, role, required: fieldRequired } of REQUIREMENT_OWNER_FIELDS) {
    const id = requirement[key];
    if (id !== undefined && typeof id !== "string") throw new Error(`${role}责任人格式无效`);
    if (!id) {
      if (required && fieldRequired) throw new Error(`请选择${role}责任人`);
      continue;
    }
    if (!members.some((member) => member.id === id && member.role === role)) throw new Error(`${role}责任人不存在或岗位角色不匹配`);
  }
}

export function memberRequirementReferences(requirements: readonly Requirement[], memberId: string): Requirement[] {
  return requirements.filter((item) => REQUIREMENT_OWNER_FIELDS.some(({ key }) => item[key] === memberId));
}

export function isMemberFilterValid(memberId: string, options: readonly Member[]): boolean {
  return !memberId || memberId === UNASSIGNED_MEMBER || options.some((item) => item.id === memberId);
}

export function filterMembers(members: readonly Member[], domains: readonly DictionaryItem[], filters: { role: MemberRole | ""; domainL0Id: string; domainL1Id: string }): Member[] {
  const children = new Set(domains.filter((item) => (item.level ?? "L1") === "L1" && (!filters.domainL0Id || item.parentId === filters.domainL0Id)).map((item) => item.id));
  return members.filter((member) => (!filters.role || member.role === filters.role)
    && (!filters.domainL0Id || member.domainL1Ids.some((id) => children.has(id)))
    && (!filters.domainL1Id || (children.has(filters.domainL1Id) && member.domainL1Ids.includes(filters.domainL1Id))));
}

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

export function resolvePptDocumentOwnerAccount(
  officeIdentity: string | undefined,
  members: readonly Member[],
): string | undefined {
  if (!officeIdentity?.trim()) return undefined;
  const normalizedAccount = normalizeMicrosoftAccount(officeIdentity);
  const localIdentity = normalizedAccount
    .split("@")[0]
    ?.split("\\")
    .at(-1)
    ?.trim()
    .toLocaleLowerCase("en-US");
  const employeeMember = members.find(
    (member) => member.employeeNumber.trim().toLocaleLowerCase("en-US") === localIdentity,
  );
  if (employeeMember) return normalizeMicrosoftAccount(employeeMember.microsoftAccount);

  const exactMember = members.find(
    (member) => normalizeMicrosoftAccount(member.microsoftAccount) === normalizedAccount,
  );
  if (exactMember) return normalizeMicrosoftAccount(exactMember.microsoftAccount);

  return isValidMicrosoftAccount(normalizedAccount) ? normalizedAccount : undefined;
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
  requirement: RequirementOwners,
  members: readonly Member[],
  role: MemberRole,
  memberId: string,
): boolean {
  if (!memberId) return true;
  const key = ownerField(role);
  const assigned = key ? requirement[key] : undefined;
  return memberId === UNASSIGNED_MEMBER ? !assigned : assigned === memberId;
}

export function visibleMembersForRole(members: readonly Member[], requirements: readonly RequirementOwners[], role: MemberRole): Member[] {
  const key = ownerField(role);
  const ids = new Set(key ? requirements.map((item) => item[key]) : []);
  return members.filter((member) => member.role === role && ids.has(member.id));
}
