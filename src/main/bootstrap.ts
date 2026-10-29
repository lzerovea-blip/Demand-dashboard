import { app } from "electron";
import { join } from "node:path";
import { initializeMcpDiagnostics, diagnosticErrorCode } from "./mcpDiagnostics.js";

const userData = process.env.SOLUTION_REQUIREMENTS_USER_DATA_DIR || join(app.getPath("appData"), "需求路标工作台");
const stdio = process.argv.includes("--mcp-stdio");
const logger = initializeMcpDiagnostics(userData, stdio ? "stdio" : "workbench", app.getVersion(), process.env.SOLUTION_MCP_SELF_TEST === "1" ? "self-test" : "external");
process.on("exit", () => logger.end());
process.on("uncaughtExceptionMonitor", error => logger.record("process-error", "process", { level: "error", errorCode: diagnosticErrorCode(error) }));
process.on("unhandledRejection", error => logger.record("process-error", "process", { level: "error", errorCode: diagnosticErrorCode(error) }));
void import("./index.js").catch(error => {
  logger.record("process-error", "process", { level: "error", errorCode: diagnosticErrorCode(error) });
  process.stderr.write("Application startup failed; inspect MCP diagnostics.\n");
  app.exit(1);
});
