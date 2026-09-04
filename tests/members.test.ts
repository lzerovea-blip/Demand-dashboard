import { describe, expect, it } from "vitest";
import { filterMembers, isMemberFilterValid, isValidMicrosoftAccount, normalizePptAccessPolicy, requirementMatchesResponsibleMember, resolvePptDocumentOwnerAccount, responsibleMembers, UNASSIGNED_MEMBER, visibleMembersForRole, ownerField, backfillRequirementOwners, validateRequirementOwners, savedResponsibleMembers } from "../src/shared/members";
import type { Member, Requirement } from "../src/shared/types";

const members: Member[] = [
  member("m1", "规划甲", "规划", ["d1"], 1),
  member("m2", "规划乙", "规划", ["d1", "d2"], 0),
  member("m3", "产品甲", "产品经理", ["d2"], 2),
  member("m4", "开发甲", "开发", ["d1"], 3),
];

describe("人员责任矩阵", () => {
  it("人员角色、L0 和 L1 采用 AND；L0 匹配任意负责子级", () => {
    const domains = ["d1", "d2"].map((id) => ({ id, name: id, level: "L1" as const, parentId: id === "d1" ? "health" : "sports", active: true, sortOrder: 0 }));
    expect(filterMembers(members, domains, { role: "规划", domainL0Id: "sports", domainL1Id: "" }).map((item) => item.id)).toEqual(["m2"]);
    expect(filterMembers(members, domains, { role: "规划", domainL0Id: "health", domainL1Id: "d2" })).toEqual([]);
    expect(filterMembers(members, domains, { role: "产品经理", domainL0Id: "health", domainL1Id: "" })).toEqual([]);
  });
  it.each(["规划", "产品经理", "SE"] as const)("%s 支持未配置筛选，且不自动清空该特殊选项", (role) => {
    expect(requirementMatchesResponsibleMember({}, members, role, UNASSIGNED_MEMBER)).toBe(true);
    expect(requirementMatchesResponsibleMember({ [ownerField(role)!]: "a" }, [member("a", "同名", role, ["d1"], 0)], role, UNASSIGNED_MEMBER)).toBe(false);
    expect(isMemberFilterValid(UNASSIGNED_MEMBER, [])).toBe(true);
    expect(isMemberFilterValid("deleted", [])).toBe(false);
  });
  it("角色空值与指定成员 AND 组合按角色和成员 ID 判定", () => {
    const sameName = [member("a", "同名", "规划", ["d1"], 0), member("b", "同名", "规划", ["d2"], 1)];
    const rows = ["a", "b", ""].map((planningMemberId) => ({ planningMemberId }));
    expect(rows.filter((row) => requirementMatchesResponsibleMember(row, sameName, "规划", "b") && requirementMatchesResponsibleMember(row, sameName, "SE", UNASSIGNED_MEMBER))).toEqual([{ planningMemberId: "b" }]);
  });
  it("同一 L1 可匹配多名同角色成员并按排序展示", () => {
    expect(responsibleMembers(members, "d1", "规划").map((item) => item.id)).toEqual(["m2", "m1"]);
  });

  it("责任人筛选与可见 L1 选项正确联动", () => {
    const requirement = { domainId: "d2", planningMemberId: "m2", productManagerMemberId: "m3" } as Requirement;
    expect(requirementMatchesResponsibleMember(requirement, members, "规划", "m2")).toBe(true);
    expect(requirementMatchesResponsibleMember(requirement, members, "规划", "m1")).toBe(false);
    expect(visibleMembersForRole(members, [requirement], "产品经理").map((item) => item.id)).toEqual(["m3"]);
  });

  it("旧数据只预填唯一 L1 匹配，多人/无人为空且不覆盖显式分配", () => {
    expect(backfillRequirementOwners({ domainId: "d1" }, members)).toMatchObject({ planningMemberId: "", productManagerMemberId: "", seMemberId: "" });
    expect(backfillRequirementOwners({ domainId: "d2" }, members)).toMatchObject({ planningMemberId: "m2", productManagerMemberId: "m3" });
    expect(backfillRequirementOwners({ domainId: "d2", planningMemberId: "m1", productManagerMemberId: "" }, members)).toMatchObject({ planningMemberId: "m1", productManagerMemberId: "" });
  });

  it("保存必须有三名有效匹配角色，可选择非当前领域成员；显示跟随改名", () => {
    const all = [...members, member("se", "架构师", "SE", [], 4)];
    const assigned = { planningMemberId: "m1", productManagerMemberId: "m3", seMemberId: "se" };
    expect(() => validateRequirementOwners(assigned, all, true)).not.toThrow();
    for (const key of ["planningMemberId", "productManagerMemberId", "seMemberId"]) {
      expect(() => validateRequirementOwners({ ...assigned, [key]: "" }, all, true)).toThrow("请选择");
      expect(() => validateRequirementOwners({ ...assigned, [key]: "m4" }, all, true)).toThrow("角色不匹配");
    }
    expect(() => validateRequirementOwners({}, all)).not.toThrow();
    expect(savedResponsibleMembers(assigned, all.map((item) => ({ ...item, name: "新姓名" })), "规划")[0].name).toBe("新姓名");
  });

  it("只接受完整邮箱，并去重无效白名单项", () => {
    expect(isValidMicrosoftAccount("user@company.com")).toBe(true);
    expect(isValidMicrosoftAccount("10086")).toBe(false);
    expect(normalizePptAccessPolicy({ entries: [{ memberId: "m1", permission: "read" }, { memberId: "m1", permission: "change" }], updatedAt: "2026-08-28" }).entries).toEqual([{ memberId: "m1", permission: "change" }]);
  });

  it("将 Office 内部工号身份映射到人员字典企业账号", () => {
    const owner = member("owner", "林忠维", "规划", [], 0);
    owner.employeeNumber = "l00850411";
    owner.microsoftAccount = "zw.lin@huawei.com";
    expect(resolvePptDocumentOwnerAccount("l00850411@china.huawei.com", [owner])).toBe("zw.lin@huawei.com");
    expect(resolvePptDocumentOwnerAccount("OTHER@HUAWEI.COM", [owner])).toBe("other@huawei.com");
    expect(resolvePptDocumentOwnerAccount(undefined, [owner])).toBeUndefined();
  });
});

function member(id: string, name: string, role: Member["role"], domainL1Ids: string[], sortOrder: number): Member {
  return { id, name, employeeNumber: `E-${id}`, microsoftAccount: `${id}@company.com`, role, domainL1Ids, sortOrder, createdAt: "2026-08-20T00:00:00.000Z", updatedAt: "2026-08-20T00:00:00.000Z" };
}
