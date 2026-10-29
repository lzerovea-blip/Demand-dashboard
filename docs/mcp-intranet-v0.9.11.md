# v0.9.11 独立 MCP 入口

## 日志结论

用户提供的 v0.9.10 Windows 日志中，外部客户端及内部自检均在 stdio-ready 后约 1–2ms 关闭连接，没有 initialize 请求记录。因此失败发生在初始化进入程序之前，不能据此确认是 stdout 响应缓冲。v0.9.11 改用独立 Node CLI，绕开 Electron 子进程 stdio 生命周期。

## Windows 连接步骤

1. 解压完整文件夹 ZIP，勿只复制 EXE。打开 win-unpacked/Solution-Requirements-Manager.exe。
2. 在“模型接入”启用接入，点击“本机连接自检”。
3. 从同页“独立 MCP 连接配置”复制命令和参数。移动文件夹后需要重新复制配置。
4. 公司客户端传输类型选 stdio；命令是包内 resources/mcp/node.exe 的绝对路径，参数只有 resources/mcp/server.cjs 的绝对路径。不再填写 --mcp-stdio，也不把 GUI EXE 当 MCP 命令。

示例（须替换成实际解压路径）：

```json
{
  "mcpServers": {
    "solution-requirements": {
      "command": "C:\\Tools\\win-unpacked\\resources\\mcp\\node.exe",
      "args": ["C:\\Tools\\win-unpacked\\resources\\mcp\\server.cjs"]
    }
  }
}
```

界面分开填写时不在路径外手动加引号；参数作为一个完整路径条目。无需另装 Node，无需管理员或开放网络端口。保持图形工作台打开；使用同一 Windows 用户运行客户端和工作台。

## 安全与诊断

- CLI 仅处理协议，通过当前用户本机管道访问工作台，不直接读写数据库；五个工具和人工批量确认规则不变。
- Node 22.21.1 随包提供，下载时校验官方 SHA-256，许可证和运行文件校验值位于 resources/mcp。
- 本机自检使用与外部客户端相同的 Node 及脚本，依次完成握手、五工具发现、无匹配只读查询。
- 日志增加 runtime=node 和 STDIN_CLOSED_BEFORE_INITIALIZE；连接关闭不再只报 OPERATION_FAILED。日志仍不记录查询内容或需求正文。
- 旧 Electron --mcp-stdio 入口保留兼容，但不再作为推荐配置或自检入口。
- 若本机自检成功、公司客户端仍失败，请导出 MCP 诊断并保留客户端错误。若自检也失败，首先检查独立运行文件是否完整以及对应阶段日志。

## 验收边界

自动化测试及 Mac 实际客户端验收不能代替 Windows OfficeAce 实机验收。Windows 无执行环境时只报告构建、文件校验结果，不宣称已连接公司模型。
