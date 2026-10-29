import { describe, expect, it } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { buildMcpServer } from "../src/main/mcpStdio";

describe("standard MCP handshake and tools", () => {
  it("negotiates protocol, publishes seven approved tools and reports a missing workbench", async () => {
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    const server = buildMcpServer("/nonexistent-mcp-workbench");
    const client = new Client({ name: "acceptance-client", version: "1.0.0" });
    try {
      await server.connect(serverTransport);
      await client.connect(clientTransport);
      expect((await client.listTools()).tools.map((tool) => tool.name)).toEqual([
        "list_requirements", "get_requirement", "list_requirement_references",
        "propose_requirement_changes", "get_proposal_status", "propose_dictionary_additions", "propose_export",
      ]);
      const result = await client.callTool({ name: "list_requirements", arguments: { limit: 10 } });
      expect(result.isError).toBe(true);
      expect(JSON.stringify(result.content)).toContain("未运行或未启用");
    } finally { await client.close(); await server.close(); }
  });
  it("rejects malformed months with a correction example before reaching the workbench", async () => {
    const [left, right] = InMemoryTransport.createLinkedPair();
    const server = buildMcpServer("/nonexistent-mcp-workbench");
    const client = new Client({ name: "date-test", version: "1" });
    try {
      await server.connect(right); await client.connect(left);
      for (const targetMonth of ["2027年3月", "2027/03", "2027-3", "2027-13", "2027-00", "2027-03-01", ""]) {
        const result = await client.callTool({ name: "propose_requirement_changes", arguments: { operations: [{ type: "update", id: "r1", fields: { targetMonth } }] } });
        expect(result.isError).toBe(true);
        expect(JSON.stringify(result.content)).toContain("2027-03");
        expect(JSON.stringify(result.content)).not.toContain("未运行或未启用");
      }
      const result = await client.callTool({ name: "propose_requirement_changes", arguments: { operations: [{ type: "update", id: "r1", fields: { targetMonth: "2027-03" } }] } });
      expect(JSON.stringify(result.content)).toContain("未运行或未启用");
    } finally { await client.close(); await server.close(); }
  });
});
