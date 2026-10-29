export type McpDiagnosticRole = "workbench" | "stdio";
export type McpDiagnosticSource = "external" | "self-test";
export interface McpDiagnosticEvent {
  time: string; sequence: number; sessionId: string; pid: number; role: McpDiagnosticRole; source: McpDiagnosticSource;
  event: string; stage: string; level: "info" | "warning" | "error";
  connectionId?: string; traceId?: string; method?: string; tool?: string;
  clientName?: string; clientVersion?: string; protocolVersion?: string;
  version?: string; system?: string; ppid?: number; durationMs?: number; count?: number; errorCode?: string;
  runtime?: string; runtimeVersion?: string;
}
export interface McpDiagnosticSession {
  sessionId: string; role: McpDiagnosticRole; source: McpDiagnosticSource; pid: number;
  lastActivity: string; stage: string; event: string; active: boolean;
  clientName?: string; clientVersion?: string; lastFailure?: McpDiagnosticEvent;
}
export interface McpDiagnosticsSnapshot {
  directory: string; available: boolean; warning?: string;
  sessions: McpDiagnosticSession[]; events: McpDiagnosticEvent[];
}
export interface McpSelfTestResult { success: boolean; stage: string; message: string; checkedAt: string }
export const MCP_STAGE_NAMES: Record<string, string> = {
  export: "文件导出",
  process: "程序启动", stdio: "stdio 就绪", initialize: "初始化请求", handshake: "协议握手",
  tools: "工具列表", tool: "工具调用", bridge: "本机桥接", workbench: "工作台处理",
  response: "响应发送", closed: "连接结束", diagnostics: "诊断日志", proposal: "修改提案",
};
export const MCP_EVENT_NAMES: Record<string, string> = {
  "export-start": "已确认，开始导出", "export-complete": "文件已生成", "export-cancelled": "用户取消保存", "export-failed": "导出失败",
  "process-start": "程序已启动", "process-exit": "进程退出", "process-error": "进程异常",
  "stdin-ended": "标准输入已关闭",
  "stdio-ready": "stdio 已就绪", "request-received": "收到客户端请求", "handshake-complete": "握手完成",
  "handshake-waiting": "等待客户端初始化超过 30 秒", "protocol-error": "协议消息无法解析",
  "response-sent": "响应已发送", "request-failed": "请求失败", "connection-closed": "连接关闭",
  "bridge-enabled": "本机接入已启用", "bridge-disabled": "本机接入已停用",
  "bridge-connecting": "正在连接工作台", "bridge-connected": "已连接工作台", "bridge-request": "工作台收到请求",
  "workbench-start": "工作台开始处理", "workbench-complete": "工作台处理完成", "bridge-complete": "收到工作台响应",
  "bridge-failed": "本机桥接失败", "self-test-start": "本机自检开始", "self-test-complete": "本机自检完成",
  "self-test-failed": "本机自检失败", "proposal-decision": "提案处理结束", "logging-failed": "日志不可用",
};
