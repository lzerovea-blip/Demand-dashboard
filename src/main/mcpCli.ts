import { homedir } from "node:os";
import { join } from "node:path";
import { initializeMcpDiagnostics, diagnosticErrorCode } from "./mcpDiagnostics.js";
import { serveMcpStdio } from "./mcpStdio.js";
import { MCP_APP_VERSION } from "../shared/mcpVersion.js";
import { resolveMcpUserData } from "../shared/mcpLaunch.js";

const directory = resolveMcpUserData(process.platform, homedir(), process.env);
const logger = initializeMcpDiagnostics(directory, "stdio", MCP_APP_VERSION, process.env.SOLUTION_MCP_SELF_TEST === "1" ? "self-test" : "external");
process.on("exit", () => logger.end());
process.on("uncaughtException", error => {
  logger.record("process-error", "process", { level: "error", errorCode: diagnosticErrorCode(error) });
  process.stderr.write("MCP CLI failed; inspect diagnostics.\n"); process.exit(1);
});
process.on("unhandledRejection", error => {
  logger.record("process-error", "process", { level: "error", errorCode: diagnosticErrorCode(error) });
  process.stderr.write("MCP CLI failed; inspect diagnostics.\n"); process.exit(1);
});
// No Electron imports, GUI lifecycle or database access. EOF closes the transport,
// after which Node exits naturally with all pending output writes completed.
void serveMcpStdio(directory).catch(error => {
  logger.record("process-error", "process", { level: "error", errorCode: diagnosticErrorCode(error) });
  process.stderr.write("MCP CLI startup failed; inspect diagnostics.\n"); process.exitCode = 1;
});
