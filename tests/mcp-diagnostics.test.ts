import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { randomUUID } from "node:crypto";
import { PassThrough } from "node:stream";
import { createServer, createConnection } from "node:net";
import JSZip from "jszip";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { McpDiagnosticLogger, MCP_LOG_LIMIT, createMcpDiagnosticArchive, diagnosticErrorCode, readMcpDiagnostics } from "../src/main/mcpDiagnostics";
import { McpDiagnosticTransport } from "../src/main/mcpDiagnosticTransport";
import { buildMcpServer } from "../src/main/mcpStdio";
import { LocalMcpBridge, callLocalMcpBridge } from "../src/main/mcpBridge";
import { McpProposalStore } from "../src/main/mcpProposals";
import type { AppSnapshot } from "../src/shared/types";

let directory: string;
beforeEach(async () => { directory = await mkdtemp(join(tmpdir(), "mcpdiag-")); });
afterEach(async () => { vi.useRealTimers(); vi.restoreAllMocks(); await rm(directory, { recursive: true, force: true }); });
const snapshot = { requirements: [], domains: [], products: [], members: [] } as unknown as AppSnapshot;

describe("MCP diagnostic metadata", () => {
  it("keeps one trace across real SDK requests and the local bridge without logging query or invalid arguments", async () => {
    const workbench = new McpDiagnosticLogger(directory, "workbench", "0.9.10");
    const stdio = new McpDiagnosticLogger(directory, "stdio", "0.9.10");
    const bridge = new LocalMcpBridge(directory, () => snapshot, new McpProposalStore({ getSnapshot: () => snapshot, applyMcpRequirementBatch: async () => snapshot }), () => undefined, workbench);
    await bridge.start();
    const [clientTransport, raw] = InMemoryTransport.createLinkedPair();
    const transport = new McpDiagnosticTransport(raw, stdio);
    const server = buildMcpServer(directory, stdio);
    server.server.oninitialized = () => transport.initialized(server.server.getClientVersion());
    const client = new Client({ name: "acceptance-client", version: "1.0" });
    try {
      await server.connect(transport); await client.connect(clientTransport);
      expect((await client.listTools()).tools).toHaveLength(7);
      const result = await client.callTool({ name: "list_requirements", arguments: { query: "PRIVATE_REQUIREMENT_SEARCH", limit: 1 } });
      expect(result.isError).not.toBe(true);
      await client.callTool({ name: "list_requirements", arguments: { limit: "PRIVATE_INVALID_ARGUMENT" } }).catch(() => undefined);
      await clientTransport.send({ jsonrpc: "2.0", id: 9999, method: "private_unknown_method" });
      await new Promise(resolve => setTimeout(resolve, 20));
      const info = await readMcpDiagnostics(directory);
      const start = info.events.find(item => item.event === "request-received" && item.tool === "list_requirements");
      const processed = info.events.find(item => item.event === "workbench-complete");
      expect(processed?.connectionId).toBe(stdio.sessionId);
      expect(processed?.traceId).toBeTruthy();
      expect(info.events.some(item => item.event === "response-sent" && item.traceId === processed?.traceId)).toBe(true);
      expect(start).toBeDefined();
      expect(info.events.some(item => item.errorCode === "TOOL_VALIDATION_FAILED")).toBe(true);
      expect(info.events.some(item => item.errorCode === "JSONRPC_-32601")).toBe(true);
      expect(info.sessions.find(item => item.role === "stdio")?.clientName).toBe("acceptance-client");
      const all = JSON.stringify(info);
      expect(all).not.toMatch(/PRIVATE_REQUIREMENT_SEARCH|PRIVATE_INVALID_ARGUMENT|private_unknown_method/);
    } finally { await client.close(); await server.close(); await bridge.stop(); }
  });
  it("reports malformed stdio safely without writing diagnostic text to stdout", async () => {
    const logger = new McpDiagnosticLogger(directory, "stdio", "0.9.10");
    const input = new PassThrough(), output = new PassThrough();
    let stdout = ""; output.on("data", chunk => { stdout += chunk.toString(); });
    const transport = new McpDiagnosticTransport(new StdioServerTransport(input, output), logger);
    await transport.start(); input.write('{"PRIVATE_BAD_JSON":}\n');
    const info = await readMcpDiagnostics(directory);
    expect(info.events.some(item => item.errorCode === "INVALID_PROTOCOL_MESSAGE")).toBe(true);
    expect(JSON.stringify(info)).not.toContain("PRIVATE_BAD_JSON");
    expect(stdout).toBe("");
    await transport.close();
  });
  it("isolates concurrent clients and records a proposal without storing its title", async () => {
    const data = { ...snapshot, requirements: [{ id: "fixture", title: "PRIVATE_PROPOSAL_TITLE", images: [], updatedAt: "2026-10-23T00:00:00Z" }] } as unknown as AppSnapshot;
    const workbench = new McpDiagnosticLogger(directory,"workbench","0.9.10");
    const bridge = new LocalMcpBridge(directory,()=>data,new McpProposalStore({getSnapshot:()=>data,applyMcpRequirementBatch:async()=>data}),()=>undefined,workbench);
    await bridge.start();
    const participants=[1,2].map(index=>{
      const logger=new McpDiagnosticLogger(directory,"stdio","0.9.10");
      const [clientTransport,raw]=InMemoryTransport.createLinkedPair();
      const transport=new McpDiagnosticTransport(raw,logger);
      const server=buildMcpServer(directory,logger);
      server.server.oninitialized=()=>transport.initialized(server.server.getClientVersion());
      return {logger,clientTransport,transport,server,client:new Client({name:`parallel-${index}`,version:"1"})};
    });
    try {
      await Promise.all(participants.map(async p=>{await p.server.connect(p.transport);await p.client.connect(p.clientTransport);}));
      await Promise.all(participants.map(p=>p.client.callTool({name:"list_requirements",arguments:{query:"NO_MATCH"}})));
      const proposal=await participants[1].client.callTool({name:"propose_requirement_changes",arguments:{operations:[{type:"duplicate",id:"fixture"}]}});
      expect(proposal.isError).not.toBe(true);
      const info=await readMcpDiagnostics(directory);
      for(const p of participants) expect(info.events.some(e=>e.event==="workbench-complete"&&e.connectionId===p.logger.sessionId)).toBe(true);
      expect(info.events.some(e=>e.event==="workbench-complete"&&e.tool==="propose_requirement_changes"&&e.count===1)).toBe(true);
      expect(JSON.stringify(info)).not.toMatch(/PRIVATE_PROPOSAL_TITLE|NO_MATCH/);
    } finally {for(const p of participants){await p.client.close();await p.server.close();}await bridge.stop();}
    expect((await readMcpDiagnostics(directory)).sessions.filter(s=>s.role==="stdio").every(s=>!s.active)).toBe(true);
  });
  it("marks unfinished requests when a client disconnects", async () => {
    const logger=new McpDiagnosticLogger(directory,"stdio","0.9.10");
    const [clientTransport,raw]=InMemoryTransport.createLinkedPair();
    const transport=new McpDiagnosticTransport(raw,logger);
    transport.onmessage=()=>undefined;
    await transport.start();await clientTransport.start();
    await clientTransport.send({jsonrpc:"2.0",id:1,method:"tools/list"});
    await transport.close();
    expect((await readMcpDiagnostics(directory)).events.some(e=>e.errorCode==="CLIENT_DISCONNECTED")).toBe(true);
  });
  it("records a 30s handshake warning and accepts a later successful initialization", async () => {
    vi.useFakeTimers();
    const logger = new McpDiagnosticLogger(directory, "stdio", "0.9.10");
    const transport = new McpDiagnosticTransport(new StdioServerTransport(new PassThrough(), new PassThrough()), logger);
    await transport.start(); await vi.advanceTimersByTimeAsync(30001);
    transport.initialized({ name: "late-client", version: "2" });
    const info = await readMcpDiagnostics(directory);
    expect(info.events.some(item => item.event === "handshake-waiting")).toBe(true);
    expect(info.events.some(item => item.event === "handshake-complete")).toBe(true);
    await transport.close();
  });
  it("reports missing workbench, dead pipe and invalid authentication", async () => {
    const logger = new McpDiagnosticLogger(directory, "stdio", "0.9.10");
    await expect(callLocalMcpBridge(directory, "list_requirements", {}, undefined, logger)).rejects.toThrow("未运行或未启用");
    await writeFile(join(directory, "mcp-session.json"), JSON.stringify({ pipe: join(directory,"missing.sock"), token: "PRIVATE_TOKEN" }));
    await expect(callLocalMcpBridge(directory, "list_requirements", {}, undefined, logger)).rejects.toThrow("断开");
    const bridgeLogger = new McpDiagnosticLogger(directory, "workbench", "0.9.10");
    const bridge = new LocalMcpBridge(directory, () => snapshot, new McpProposalStore({ getSnapshot: () => snapshot, applyMcpRequirementBatch: async () => snapshot }), () => undefined, bridgeLogger);
    await bridge.start();
    try {
      const descriptor = JSON.parse(await readFile(join(directory, "mcp-session.json"), "utf8"));
      await new Promise<void>((resolve, reject) => { const socket = createConnection(descriptor.pipe); socket.on("error",reject); socket.on("connect",()=>socket.write(JSON.stringify({token:"PRIVATE_TOKEN",action:"list_requirements"})+'\n')); socket.on("data",()=>undefined); socket.on("end",resolve); });
      const info = await readMcpDiagnostics(directory);
      expect(info.events.some(item => item.errorCode === "BRIDGE_AUTH_FAILED")).toBe(true);
      expect(JSON.stringify(info)).not.toContain("PRIVATE_TOKEN");
    } finally { await bridge.stop(); }
  });
  it("logs a real stalled bridge timeout", async () => {
    const path = join(directory, "stall.sock");
    const server = createServer(socket => socket.on("error",()=>undefined));
    await new Promise<void>(resolve => server.listen(path,resolve));
    await writeFile(join(directory,"mcp-session.json"),JSON.stringify({pipe:path,token:"not-logged"}));
    const logger = new McpDiagnosticLogger(directory,"stdio","0.9.10");
    try { await expect(callLocalMcpBridge(directory,"list_requirements",{},undefined,logger)).rejects.toThrow("超时");
      expect((await readMcpDiagnostics(directory)).events.some(item=>item.errorCode==="TIMEOUT")).toBe(true);
    } finally { server.close(); }
  }, 8000);
  it("rotates bounded files, exports only diagnostic files and separates self-tests", async () => {
    const logger = new McpDiagnosticLogger(directory,"stdio","0.9.10","self-test");
    await writeFile(logger.path,"x".repeat(MCP_LOG_LIMIT-32));
    logger.record("handshake-complete","handshake",{clientName:"user@company.com", errorCode:diagnosticErrorCode(new Error('PRIVATE_BODY'))});
    expect((await stat(logger.path)).size).toBeLessThan(MCP_LOG_LIMIT);
    expect((await readdir(logger.directory)).some(item=>item.endsWith('.previous.log'))).toBe(true);
    await writeFile(join(directory,"mcp-session.json"),'PRIVATE_TOKEN');
    await writeFile(join(directory,"database.sqlite"),'PRIVATE_DB');
    const archive=await JSZip.loadAsync(await createMcpDiagnosticArchive(directory,"0.9.10"));
    expect(Object.keys(archive.files).every(name=>name==="summary.json"||name.endsWith(".log"))).toBe(true);
    const info=await readMcpDiagnostics(directory);
    expect(info.sessions[0].source).toBe("self-test");
    expect(JSON.stringify(info)).not.toMatch(/user@company.com|PRIVATE_BODY|PRIVATE_TOKEN/);
  });
  it("keeps twenty inactive sessions and never prunes a live process", async () => {
    const logDir=join(directory,"logs","mcp"); await mkdir(logDir,{recursive:true});
    for(let i=0;i<23;i++) await writeFile(join(logDir,`mcp-stdio-99999999-${randomUUID()}.log`),"");
    const active=join(logDir,`mcp-stdio-${process.pid}-${randomUUID()}.log`); await writeFile(active,"");
    new McpDiagnosticLogger(directory,"workbench","0.9.10");
    const names=await readdir(logDir);
    expect(names.filter(item=>item.includes("99999999"))).toHaveLength(20);
    expect(names).toContain(active.split("/").at(-1));
  });
  it("never throws when diagnostic writing is unavailable", async () => {
    await writeFile(join(directory,"logs"),"not a directory");
    const stderr=vi.spyOn(process.stderr,"write").mockImplementation(()=>true);
    const logger=new McpDiagnosticLogger(directory,"stdio","0.9.10");
    expect(()=>logger.record("request-received","tool")).not.toThrow();
    expect(logger.available).toBe(false);
    expect(stderr.mock.calls.flat().join("")).toContain("LOG_WRITE_FAILED");
  });
});
