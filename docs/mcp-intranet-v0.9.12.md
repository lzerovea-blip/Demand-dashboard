# v0.9.12 MCP 字典新增与日期格式校验

继续使用 v0.9.11 独立 Node 入口：命令为包内 resources/mcp/node.exe，唯一参数为 resources/mcp/server.cjs，两者都填写绝对路径。打开工作台并启用 MCP。更新后请在客户端断开重连，刷新六个工具的定义。

## 新增字典条目

新增工具 `propose_dictionary_additions` 接受 `{ "entries": [...] }`，每批 1–25 项：

| kind | 类型 | 必须提供 | 可选项 |
| --- | --- | --- | --- |
| domain_l0 | 领域 L0 | name | 无 |
| domain_l1 | 领域 L1 | name、parentId（已有 L0 ID） | 无 |
| product | 产品 | name | 无 |
| member | 成员 | name、employeeNumber、microsoftAccount（完整企业邮箱） | role、domainL1Ids（已有 L1 ID 数组） |
| roadmap | 路标 | name、roadmapKind | 无 |
| lane | 泳道 | name、roadmapId（已有路标 ID） | 无 |

roadmapKind 为 standard（普通卡片）或 medical（医疗认证时间轴）。role 为规划、产品经理、SE、PM、开发、算法、UX之一；可省略。

先调用 `list_requirement_references` 读取现有 ID、岗位选项、路标和泳道。引用须使用真实已有 ID，不按名称猜测；若父级也是新的，先提出并人工批准父级，再重新查询 ID 后提出子级。成员工号、邮箱必须唯一；新增成员不自动加入 PPT 白名单。

示例：

```json
{"entries":[{"kind":"product","name":"示例产品"},{"kind":"roadmap","name":"示例路标","roadmapKind":"standard"}]}
```

返回提案 ID 后，只是 pending；必须在工作台“待确认的 MCP 修改”检查每个字段并确认，才会整批保存。拒绝、30 分钟过期或重启均不会写入。通过 `get_proposal_status` 查询处理结果；新增后的真实 ID 重新从字典查询获取。

本次仅开放新增条目，不是增加任意数据库字段；不能修改、删除、覆盖已有字典。提案时和批准时都校验，重复名称或父级已失效时整批失败。多个字典操作复用隔离数据库及单次磁盘提交，后续操作失败时撤销整批。

## 日期格式

`propose_requirement_changes` 的 targetMonth 必须为 `YYYY-MM`，如 `2027-03`。工具定义公布格式、正则和示例；在 MCP 参数入口及提案入口都进行校验。

拒绝 `2027年3月`、`2027/03`、`2027-3`、`2027-13`、完整日期、时间戳及空值。错误会提示改用 YYYY-MM；不静默猜测或自动改写。无效日期不会生成待确认提案，原需求不变。只修改其他字段时可以不传 targetMonth。

## 安全及交付边界

字典查询仍不返回成员工号、邮箱或 PPT 权限；人员新增所需工号及邮箱由使用者提供，工作台确认页展示待新增值。诊断仅记录工具名、数量、阶段和结果，不记录新增内容。

不修改数据库版本、数据包、Excel 格式、路标匹配、分层支持结论或现有查询分页。v0.9.11 的 Windows OfficeAce 连接已获用户确认；v0.9.12 仍需更新后在公司客户端复测新工具和日期提示。本轮不上传 GitHub。
