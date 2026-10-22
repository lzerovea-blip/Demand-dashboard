import { randomBytes } from "node:crypto";
import { createServer, createConnection, type Server } from "node:net";
import { chmod, readFile, rm, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import type { AppSnapshot } from "../shared/types.js";
import { mcpRequirementView, type McpRequirementOperation } from "../shared/mcp.js";
import { McpProposalStore } from "./mcpProposals.js";

interface SessionDescriptor { pipe: string; token: string }
const SESSION_FILE = "mcp-session.json";
const MAX_REQUEST_BYTES = 1024 * 1024;
const MAX_RESPONSE_BYTES = 8 * 1024 * 1024;

export class LocalMcpBridge {
  private server: Server | null = null;
  private descriptor: SessionDescriptor | null = null;
  constructor(private readonly userData: string, private readonly getSnapshot: () => AppSnapshot,
    private readonly proposals: McpProposalStore, private readonly diagnostic: (event: string, detail: string) => void) {}

  get isRunning(): boolean { return this.server !== null; }

  async start(): Promise<void> {
    if (this.server) return;
    const nonce = randomBytes(16).toString("hex");
    const pipe = process.platform === "win32" ? `\\\\.\\pipe\\solution-requirements-${nonce}` : join(tmpdir(), `mcp-${nonce}.sock`);
    const descriptor = { pipe, token: randomBytes(32).toString("hex") };
    const server = createServer((socket) => {
      let buffer = "";
      socket.setEncoding("utf8");
      socket.on("data", (chunk: string) => {
        buffer += chunk;
        if (Buffer.byteLength(buffer) > MAX_REQUEST_BYTES) { socket.destroy(); return; }
        const newline = buffer.indexOf("\n");
        if (newline < 0) return;
        const line = buffer.slice(0, newline);
        buffer = "";
        void (async () => {
          try {
            const request = JSON.parse(line) as { token: string; action: string; args?: unknown };
            if (request.token !== descriptor.token) throw new Error("MCP 会话认证失败");
            socket.end(`${JSON.stringify({ ok: true, result: this.handle(request.action, request.args) })}\n`);
          } catch (error) {
            socket.end(`${JSON.stringify({ ok: false, error: error instanceof Error ? error.message : "MCP 请求失败" })}\n`);
          }
        })();
      });
    });
    try {
      await new Promise<void>((resolve, reject) => { server.once("error", reject); server.listen(pipe, resolve); });
      if (process.platform !== "win32") await chmod(pipe, 0o600);
      await writeFile(join(this.userData, SESSION_FILE), JSON.stringify(descriptor), { mode: 0o600 });
      if (process.platform !== "win32") await chmod(join(this.userData, SESSION_FILE), 0o600);
      this.server = server;
      this.descriptor = descriptor;
      this.diagnostic("mcp-started", "stdio bridge enabled");
    } catch (error) { server.close(); throw error; }
  }

  async stop(): Promise<void> {
    const server = this.server;
    const descriptor = this.descriptor;
    this.server = null;
    this.descriptor = null;
    await rm(join(this.userData, SESSION_FILE), { force: true });
    if (server) await new Promise<void>((resolve) => server.close(() => resolve()));
    if (descriptor && process.platform !== "win32") await rm(descriptor.pipe, { force: true });
    if (server) this.diagnostic("mcp-stopped", "stdio bridge disabled");
  }

  private handle(action: string, args: unknown): unknown {
    const snapshot = this.getSnapshot();
    const input = (args ?? {}) as Record<string, unknown>;
    if (action === "list_requirements") {
      const offset = Math.max(0, Math.floor(Number(input.offset ?? 0)));
      const limit = Math.max(1, Math.min(100, Math.floor(Number(input.limit ?? 30))));
      const query = String(input.query ?? "").trim().toLocaleLowerCase();
      const filtered = snapshot.requirements.filter((item) => !query || `${item.title} ${item.description}`.toLocaleLowerCase().includes(query));
      return { total: filtered.length, offset, limit, items: filtered.slice(offset, offset + limit).map(mcpRequirementView) };
    }
    if (action === "get_requirement") {
      const requirement = snapshot.requirements.find((item) => item.id === input.id);
      if (!requirement) throw new Error("需求不存在");
      return mcpRequirementView(requirement);
    }
    if (action === "list_requirement_references") return {
      domains: snapshot.domains.map(({ id, name, level, parentId }) => ({ id, name, level, parentId })),
      products: snapshot.products.map(({ id, name }) => ({ id, name })),
      members: snapshot.members.map(({ id, name, role, domainL1Ids }) => ({ id, name, role, domainL1Ids })),
    };
    if (action === "propose_requirement_changes") {
      const operations = input.operations as McpRequirementOperation[];
      const proposal = this.proposals.propose(operations);
      this.diagnostic("mcp-proposal-created", `operations=${operations.length} create=${operations.filter((item) => item.type === "create").length} update=${operations.filter((item) => item.type === "update").length} duplicate=${operations.filter((item) => item.type === "duplicate").length} delete=${operations.filter((item) => item.type === "delete").length}`);
      return { id: proposal.id, state: proposal.state, expiresAt: proposal.expiresAt, itemCount: proposal.items.length };
    }
    if (action === "get_proposal_status") {
      const proposal = this.proposals.get(String(input.id ?? ""));
      return { id: proposal.id, state: proposal.state, expiresAt: proposal.expiresAt, itemCount: proposal.items.length, error: proposal.error };
    }
    throw new Error("不支持的 MCP 操作");
  }
}

export async function callLocalMcpBridge(userData: string, action: string, args: unknown): Promise<unknown> {
  let descriptor: SessionDescriptor;
  try { descriptor = JSON.parse(await readFile(join(userData, SESSION_FILE), "utf8")) as SessionDescriptor; }
  catch { throw new Error("工作台未运行或未启用 MCP；请先打开工作台并在“模型接入”中启用"); }
  return new Promise((resolve, reject) => {
    const socket = createConnection(descriptor.pipe);
    let response = "";
    socket.setEncoding("utf8");
    socket.setTimeout(5000, () => socket.destroy(new Error("本机工作台连接超时")));
    socket.on("connect", () => socket.write(`${JSON.stringify({ token: descriptor.token, action, args })}\n`));
    socket.on("data", (chunk: string) => { response += chunk; if (Buffer.byteLength(response) > MAX_RESPONSE_BYTES) socket.destroy(new Error("MCP 响应过大，请缩小分页大小")); });
    socket.on("error", (error) => reject(new Error(`工作台连接已断开：${error.message}`)));
    socket.on("end", () => {
      try {
        const parsed = JSON.parse(response) as { ok: boolean; result?: unknown; error?: string };
        if (!parsed.ok) reject(new Error(parsed.error ?? "MCP 请求失败")); else resolve(parsed.result);
      } catch { reject(new Error("工作台连接已断开或响应无效")); }
    });
  });
}
