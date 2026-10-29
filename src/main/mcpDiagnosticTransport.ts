import { createHash } from "node:crypto";
import type { Transport, TransportSendOptions } from "@modelcontextprotocol/sdk/shared/transport.js";
import type { JSONRPCMessage, MessageExtraInfo } from "@modelcontextprotocol/sdk/types.js";
import { McpDiagnosticLogger, mcpTraceContext, type McpTrace } from "./mcpDiagnostics.js";

const METHODS = new Set(["initialize", "notifications/initialized", "tools/list", "tools/call", "ping", "resources/list", "resources/templates/list"]);
const TOOLS = new Set(["list_requirements", "get_requirement", "list_requirement_references", "propose_requirement_changes", "get_proposal_status", "propose_dictionary_additions", "propose_export"]);
export class McpDiagnosticTransport implements Transport {
  onmessage?: Transport["onmessage"];
  onerror?: Transport["onerror"];
  onclose?: Transport["onclose"];
  private pending = new Map<string | number, { trace: McpTrace; stage: string; method: string; tool?: string; start: number }>();
  private waiting?: ReturnType<typeof setTimeout>;
  constructor(private inner: Transport, private logger: McpDiagnosticLogger) {}
  async start() {
    this.inner.onmessage = (message: JSONRPCMessage, extra?: MessageExtraInfo) => {
      if ("method" in message && "id" in message) {
        const method = METHODS.has(message.method) ? message.method : "unknown";
        const stage = method === "initialize" ? "initialize" : method === "tools/list" ? "tools" : method === "tools/call" ? "tool" : "stdio";
        const name = message.params?.name;
        const tool = typeof name === "string" && TOOLS.has(name) ? name : undefined;
        const trace = { connectionId: this.logger.sessionId, traceId: createHash("sha256").update(String(message.id)).digest("hex").slice(0,16), source: this.logger.source };
        this.pending.set(message.id, { trace, stage, method, tool, start: Date.now() });
        this.logger.record("request-received", stage, { ...trace, method, tool });
        mcpTraceContext.run(trace, () => this.onmessage?.(message, extra));
      } else this.onmessage?.(message, extra);
    };
    this.inner.onerror = () => {
      this.logger.record("protocol-error", "stdio", { level: "error", errorCode: "INVALID_PROTOCOL_MESSAGE" });
      this.onerror?.(new Error("MCP protocol message invalid"));
    };
    this.inner.onclose = () => { this.finish(); this.onclose?.(); };
    await this.inner.start();
    this.logger.record("stdio-ready", "stdio");
    this.waiting = setTimeout(() => this.logger.record("handshake-waiting", "initialize", { level: "warning", errorCode: "INITIALIZE_NOT_COMPLETED" }), 30000);
    this.waiting.unref();
  }
  initialized(client?: { name: string; version: string }) {
    clearTimeout(this.waiting);
    this.logger.record("handshake-complete", "handshake", { clientName: client?.name, clientVersion: client?.version });
  }
  async send(message: JSONRPCMessage, options?: TransportSendOptions) {
    const request = "id" in message && message.id !== undefined ? this.pending.get(message.id) : undefined;
    try {
      await this.inner.send(message, options);
      if (request) {
        const result = "result" in message ? message.result : undefined;
        const failed = "error" in message || Boolean(result?.isError);
        const validationFailed = Array.isArray(result?.content) && result.content.some((item: unknown) => item && typeof item === "object" && "text" in item && typeof item.text === "string" && /Input validation error:/.test(item.text));
        this.logger.record(failed ? "request-failed" : "response-sent", request.stage,
          { ...request.trace, method: request.method, tool: request.tool, durationMs: Date.now()-request.start,
            level: failed ? "error" : "info", errorCode: "error" in message ? `JSONRPC_${message.error.code}` : validationFailed ? "TOOL_VALIDATION_FAILED" : failed ? "TOOL_FAILED" : undefined,
            protocolVersion: request.method === "initialize" && typeof result?.protocolVersion === "string" ? result.protocolVersion : undefined,
            count: Array.isArray(result?.tools) ? result.tools.length : undefined });
        if ("id" in message && message.id !== undefined) this.pending.delete(message.id);
      }
    } catch (error) {
      this.logger.record("request-failed", "response", { ...request?.trace, level: "error", errorCode: "STDOUT_WRITE_FAILED" });
      throw error;
    }
  }
  async close() { this.finish(); await this.inner.close(); }
  private finish() {
    clearTimeout(this.waiting);
    for (const request of this.pending.values()) this.logger.record("request-failed", request.stage, { ...request.trace, level: "error", errorCode: "CLIENT_DISCONNECTED" });
    this.pending.clear();
    this.logger.end("connection-closed");
  }
}
