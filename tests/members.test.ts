import { describe, expect, it } from "vitest";
import { isValidMicrosoftAccount, normalizePptAccessPolicy, requirementMatchesResponsibleMember, responsibleMembers, visibleMembersForRole } from "../src/shared/members";
import type { Member, Requirement } from "../src/shared/types";

const members: Member[] = [
  member("m1", "规划甲", "规划", ["d1"], 1),
  member("m2", "规划乙", "规划", ["d1", "d2"], 0),
  member("m3", "产品甲", "产品经理", ["d2"], 2),
  member("m4", "开发甲", "开发", ["d1"], 3),
];

describe("人员责任矩阵", () => {
  it("同一 L1 可匹配多名同角色成员并按排序展示", () => {
    expect(responsibleMembers(members, "d1", "规划").map((item) => item.id)).toEqual(["m2", "m1"]);
  });

  it("责任人筛选与可见 L1 选项正确联动", () => {
    const requirement = { domainId: "d2" } as Requirement;
    expect(requirementMatchesResponsibleMember(requirement, members, "规划", "m2")).toBe(true);
    expect(requirementMatchesResponsibleMember(requirement, members, "规划", "m1")).toBe(false);
    expect(visibleMembersForRole(members, new Set(["d2"]), "产品经理").map((item) => item.id)).toEqual(["m3"]);
  });

  it("只接受完整邮箱，并去重无效白名单项", () => {
    expect(isValidMicrosoftAccount("user@company.com")).toBe(true);
    expect(isValidMicrosoftAccount("10086")).toBe(false);
    expect(normalizePptAccessPolicy({ entries: [{ memberId: "m1", permission: "read" }, { memberId: "m1", permission: "change" }], updatedAt: "2026-08-28" }).entries).toEqual([{ memberId: "m1", permission: "change" }]);
  });
});

function member(id: string, name: string, role: Member["role"], domainL1Ids: string[], sortOrder: number): Member {
  return { id, name, employeeNumber: `E-${id}`, microsoftAccount: `${id}@company.com`, role, domainL1Ids, sortOrder, createdAt: "2026-08-20T00:00:00.000Z", updatedAt: "2026-08-20T00:00:00.000Z" };
}
