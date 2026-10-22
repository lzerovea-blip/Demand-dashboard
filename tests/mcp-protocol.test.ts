import { describe, expect, it } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { buildMcpServer } from "../src/main/mcpStdio";

describe("standard MCP handshake and tools", () => {
  it("negotiates protocol, publishes only the five requirement tools and reports a missing workbench", async () => {
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    const server = buildMcpServer("/nonexistent-mcp-workbench");
    const client = new Client({ name: "acceptance-client", version: "1.0.0" });
    try {
      await server.connect(serverTransport);
      await client.connect(clientTransport);
      expect((await client.listTools()).tools.map((tool) => tool.name)).toEqual([
        "list_requirements", "get_requirement", "list_requirement_references",
        "propose_requirement_changes", "get_proposal_status",
      ]);
      const result = await client.callTool({ name: "list_requirements", arguments: { limit: 10 } });
      expect(result.isError).toBe(true);
      expect(JSON.stringify(result.content)).toContain("未运行或未启用");
    } finally { await client.close(); await server.close(); }
  });
});
