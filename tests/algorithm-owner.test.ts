import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { RequirementEditor } from "../src/renderer/App";
import { backfillRequirementOwners, validateRequirementOwners, savedResponsibleMembers, memberRequirementReferences } from "../src/shared/members";
import { cloneRequirementAsCopy } from "../src/shared/requirements";
import { mergeWorkspaceData } from "../src/shared/workspace";
import { createWorkspacePackage, inspectWorkspacePackage } from "../src/main/workspacePackage";
import type { AppSnapshot, Member, Requirement, WorkspaceData } from "../src/shared/types";
const time = "2026-09-10T00:00:00.000Z";
const member: Member = { id: "algo", name: "算法成员", employeeNumber: "ALG1", microsoftAccount: "alg@example.com", role: "算法", domainL1Ids: ["l1"], sortOrder: 0, createdAt: time, updatedAt: time };
const row: Requirement = { id: "r1", title: "算法责任需求", description: "", domainL0Id: "l0", domainId: "l1", source: "健康基础", category: "平台演进", scope: "全球", targetMonth: "2027-03", overseasRegions: [], allowLatestVersionUpgrade: true, productIds: [], images: [], deviceWorkloadPm: 0, appWorkloadPm: 0, cloudWorkloadPm: 0, workloadPm: 0, unallocatedWorkloadPm: 0, algorithmMemberId: "algo", createdAt: time, updatedAt: time };
const data: WorkspaceData = { requirements: [row], members: [member], domains: [{ id: "l0", name: "健康", level: "L0", active: true, sortOrder: 0 }, { id: "l1", name: "睡眠", level: "L1", parentId: "l0", active: true, sortOrder: 0 }], products: [], groupOverrides: [], pptAccessPolicy: { entries: [], updatedAt: time } };
describe("算法责任人", () => {
  it("选填兼容，但已选人员必须存在且为算法岗位；不限制负责领域", () => {
    expect(() => validateRequirementOwners({}, [])).not.toThrow();
    expect(() => validateRequirementOwners(row, [{ ...member, domainL1Ids: [] }])).not.toThrow();
    expect(() => validateRequirementOwners(row, [])).toThrow("算法责任人不存在");
    expect(() => validateRequirementOwners(row, [{ ...member, role: "开发" }])).toThrow("角色不匹配");
    expect(() => validateRequirementOwners({ algorithmMemberId: 123 } as unknown as Requirement, [member])).toThrow("格式无效");
  });
  it("唯一匹配预填，多人不猜测；明确空值或手选不覆盖", () => {
    expect(backfillRequirementOwners({ domainId: "l1" }, [member])).toMatchObject({ algorithmMemberId: "algo" });
    expect(backfillRequirementOwners({ domainId: "l1" }, [member, { ...member, id: "other" }])).toMatchObject({ algorithmMemberId: "" });
    expect(backfillRequirementOwners({ domainId: "l1", algorithmMemberId: "" }, [member])).toMatchObject({ algorithmMemberId: "" });
    expect(backfillRequirementOwners({ ...row, domainId: "other" }, [member]).algorithmMemberId).toBe("algo");
  });
  it("复制保留分配，改名按 ID 显示，引用可供删除保护", () => {
    expect(cloneRequirementAsCopy(row, [], () => "copy", time).algorithmMemberId).toBe("algo");
    expect(savedResponsibleMembers(row, [{ ...member, name: "改名" }], "算法")[0].name).toBe("改名");
    expect(memberRequirementReferences([row], "algo")).toEqual([row]);
  });
  it("v8 往返和人员 ID 重映射，旧包缺字段合并保留本机", async () => {
    const bytes = await createWorkspacePackage(data, "0.8.21");
    expect(JSON.parse(await (await JSZip.loadAsync(bytes)).file("manifest.json")!.async("string")).formatVersion).toBe(10);
    expect((await inspectWorkspacePackage(bytes, "算法.roadmap")).data.requirements[0].algorithmMemberId).toBe("algo");
    const incoming = { ...data, members: [{ ...member, id: "remote" }], requirements: [{ ...row, algorithmMemberId: "remote", updatedAt: "2099-01-01T00:00:00.000Z" }] };
    expect(mergeWorkspaceData(data, incoming).requirements[0].algorithmMemberId).toBe("algo");
    expect(mergeWorkspaceData(data, { ...incoming, requirements: [{ ...incoming.requirements[0], algorithmMemberId: undefined }] }).requirements[0].algorithmMemberId).toBe("algo");
    const cleared = { ...data, requirements: [{ ...row, algorithmMemberId: "", updatedAt: "2099-01-01T00:00:00.000Z" }] };
    expect(mergeWorkspaceData(data, cleared).requirements[0].algorithmMemberId).toBe("");
    const diff = await inspectWorkspacePackage(await createWorkspacePackage(cleared, "0.8.21"), "清空.roadmap", data);
    expect(diff.preview.differences.flatMap(item => item.fields.map(field => field.field))).toContain("算法");
  });
  it("新建和编辑均渲染四个人员单选在同一容器，算法选项只含该岗位", () => {
    const snapshot = { ...data, templates: [], members: [member, { ...member, id: "dev", name: "不应出现在算法选项", role: "开发" }] } as AppSnapshot;
    for (const value of [undefined, row]) {
      const html = renderToStaticMarkup(createElement(RequirementEditor, { value, snapshot, onClose: () => {}, onSave: async () => {}, onDelete: async () => {}, requestConfirmation: async () => false }));
      const owners = html.split('class="requirement-owner-row"')[1].split('class="form-row-pair category-product-row"')[0];
      expect([...owners.matchAll(/aria-label="需求/g)]).toHaveLength(4);
      expect(owners).toContain('aria-label="需求算法"');
      expect(owners).toContain("算法成员");
      expect(owners).not.toContain("不应出现在算法选项");
    }
  });
});
