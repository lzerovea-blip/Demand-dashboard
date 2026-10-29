# v0.9.13 MCP 文件导出

继续使用包内独立 Node 入口：命令 resources/mcp/node.exe、唯一参数 resources/mcp/server.cjs，填写绝对路径，工作台保持打开并启用 MCP。更新后请断开重连，刷新为七个工具。

## 使用流程

模型调用 `propose_export`，返回 pending 提案 ID。你在工作台“模型接入”页检查类型、范围和安全提示，点击“确认导出”，再选择本机保存位置。模型不能指定路径、上传网址或权限账号。保存完成后，通过 `get_proposal_status` 查询 `exportResult.path`；返回路径不等于已发送邮件或上传，工具不会返回文件二进制。

取消保存返回 cancelled；拒绝、过期、重启不生成文件；生成失败返回 failed。导出期间为 exporting，不可重复确认。若提案后需求、字典、权限等数据变化，提案冲突，须重新提出。点击确认时冻结导出数据，选择保存位置期间的后续编辑不会混入该文件。

## 四种文件

```json
{"request":{"format":"requirements_excel","start":"2026H1","end":"2027H2"}}
```

requirements_excel：现有需求池协作 Excel，按所选时间范围，保持现有在线协作/回导格式。

```json
{"request":{"format":"tiering_excel","start":"2026H1","end":"2027H2"}}
```

tiering_excel：现有只读分层分级 Excel，支持矩阵和规格明细，全系列，不受分层页面临时筛选影响。

```json
{"request":{"format":"workspace"}}
```

workspace：完整 `.roadmap` 包，**全部数据、不按年份过滤**。包括需求与草稿、图片、成员工号和邮箱、PPT 白名单、字典、路标配置、分层数据及删除记录。只交给可信接收人；工作台预览会单独提醒。不能把它当成去敏后的需求清单。

```json
{"request":{"format":"ppt","start":"2026H1","end":"2027H2","materialName":"解决方案需求路标","includeSummary":true,"roadmapIds":["已有路标ID"]}}
```

ppt：复用应用现有原生可编辑 PPT 样式、国内海外分页、密集续页和医疗独立时间轴。不改变已有 IRM 流程：必须先保存白名单，且 Windows 桌面 PowerPoint 企业 IRM 可用；生成并验证保护后才写入目标位置，失败不降级生成无保护版本。有效期仍为实际导出日起六个日历月。

PPT 最多四个半年；includeSummary 默认 true；roadmapIds 省略表示全部路标，空数组表示不选路标。includeSummary=false 且 roadmapIds=[] 仅封面。具体路标 ID 从 list_requirement_references 查询。时间格式为 YYYYH1/YYYYH2（如 2027H2），不是 27H2。超范围、非法区间或未知路标会直接报错。

## 边界

本轮只增加 MCP 调用现有导出的入口，不改变数据库、Excel 或完整包格式，也不支持自动导入、网络上传、邮件发送、模型修改白名单。只读 Excel 不代表加密保护；完整包的敏感范围须用户自行核对。

日志记录工具名、文件类型、阶段及结果，不记录导出正文、完整路径、账号或白名单。MCP 结果只返回完成状态、本机路径及 PPT 页数，不返回文档所有者邮箱。Windows IRM 实机结果需单独验证，不用 Mac 测试替代。
