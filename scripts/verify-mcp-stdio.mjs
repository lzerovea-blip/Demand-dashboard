import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

const appBinary = resolve(process.argv[2] ?? "build-resources/mcp-runtime/darwin-arm64/node");
const serverScript = resolve(process.argv[3] ?? "dist-mcp/server.cjs");
const suppliedUserData = process.argv[4];
const userData = suppliedUserData ?? await mkdtemp(join(tmpdir(), "solution-mcp-stdio-"));
const transport = new StdioClientTransport({ command: appBinary, args: [serverScript],
  env: { ...process.env, SOLUTION_REQUIREMENTS_USER_DATA_DIR: userData }, stderr: "pipe" });
const client = new Client({ name: "solution-mcp-acceptance", version: "1.0.0" });
try {
  await client.connect(transport);
  const names = (await client.listTools()).tools.map((tool) => tool.name);
  const expected = ["list_requirements", "get_requirement", "list_requirement_references", "propose_requirement_changes", "get_proposal_status", "propose_dictionary_additions", "propose_export"];
  if (JSON.stringify(names) !== JSON.stringify(expected)) throw new Error(`Unexpected tools: ${names.join(", ")}`);
  const result = await client.callTool({ name: "list_requirements", arguments: { limit: 10 } });
  if (!suppliedUserData) {
    if (!result.isError || !JSON.stringify(result.content).includes("未运行或未启用")) throw new Error("Missing-workbench error was not returned");
    console.log(JSON.stringify({ handshake: "ok", toolCount: names.length, disabledWorkbench: "explicit-error", stdout: "protocol-only" }));
  } else {
    if (result.isError) throw new Error(`Workbench query failed: ${JSON.stringify(result.content)}`);
    const proposal = await client.callTool({ name: "propose_requirement_changes", arguments: { operations: [{ type: "create", fields: {
      title: "MCP 隔离环境验收提案", description: "仅用于空测试库的提案预览，不批准写入", domainL0Id: "test-l0", domainId: "test-l1",
      scope: "全球", allowLatestVersionUpgrade: true, overseasRegions: [], category: "平台演进", targetMonth: "2027-03",
      productIds: [], planningMemberId: "test-planning", productManagerMemberId: "test-pm", seMemberId: "test-se",
      deviceWorkloadPm: 0, appWorkloadPm: 0, cloudWorkloadPm: 0, researchAppWorkloadPm: 0, algorithmWorkloadPm: 0,
    } }] } });
    if (proposal.isError) throw new Error(`Proposal failed: ${JSON.stringify(proposal.content)}`);
    const id = JSON.parse(proposal.content[0].text).id;
    const status = await client.callTool({ name: "get_proposal_status", arguments: { id } });
    if (status.isError || JSON.parse(status.content[0].text).state !== "pending") throw new Error("Proposal did not remain pending approval");
    console.log(JSON.stringify({ handshake: "ok", toolCount: names.length, query: "ok", proposal: id, state: "pending", stdout: "protocol-only" }));
  }
} finally {
  await client.close();
  if (!suppliedUserData) await rm(userData, { recursive: true, force: true });
}
