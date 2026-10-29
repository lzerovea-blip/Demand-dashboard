# 应用图标

v0.8.20 使用用户选定的浅灰圆角底、深色圆形原子图案，不使用此前讨论的深蓝/青绿路标方案。

- `app-icon.png`：打包主图及窗口图标。
- `app-icon.ico`：Windows EXE，含 16、24、32、48、64、128、256px。
- `app-icon.icns`：本地 Mac 工作台图标。
- `reference.png`：用户提供的选定图截图，仅用于溯源，不打入应用。

主图通过内置图像工具基于用户截图及 Electron 原有高分辨率原子图案进行背景分离和清晰化，非新标志创作。图案来自 Electron 随附应用图标；保留项目既有 Electron 许可文件。本图标不代表 Electron 官方产品或认证。

最终整理提示：保留浅灰圆角底、深色圆形和浅青原子图案；移除截图绿色背景、棋盘格和圆角外白色杂边；外部使用真正透明通道；不增加文字、路标节点或其他图形。

图像工具的透明边缘输出存在白色杂边；最终使用 `scripts/normalize-icon.swift` 在格式转换时施加圆角 alpha 边界，保留内部原图像素，不重新绘制原子图案。外圈细白边随裁切收敛，保留浅灰圆角底和图案主体。

维护时运行 `node scripts/build-app-icons.mjs` 生成多尺寸格式（macOS 的 sips/iconutil）；脚本只做尺寸和格式转换，不重画图案。运行 `node scripts/verify-windows-icon.mjs <便携EXE> <解包应用EXE>` 校验两层 EXE 实際包含全部预期图标。
