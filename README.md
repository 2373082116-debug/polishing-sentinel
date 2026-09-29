# 抛光哨站

[下载安装包 v0.9.0](https://github.com/2373082116-debug/polishing-sentinel/releases/download/v0.9.0/PolishingSentinel-Setup-0.9.0.exe) · [SHA256 校验](https://github.com/2373082116-debug/polishing-sentinel/releases/download/v0.9.0/SHA256SUMS.txt)

Windows 实验数据整理与分析工具。支持 U 盘扫描、Sa 汇总、Excel 实验块回写、统计分析及用户主动授权的 AI 分析。

## 安装与运行

- 支持 Windows 10/11 x64；需要桌面版 Microsoft Excel（不包含在安装包中）。
- 优先使用已安装的 PowerShell 7，否则使用 Windows PowerShell 5.1。
- 安装包不包含 MATLAB、Excel、API Key 或任何实验数据。内置 Sa 模式无需 MATLAB。
- 首次启动在“软件配置”设置主表路径、U盘卷标和目录；默认新块模式为预览，确认配置后再启用自动建块。
- 需要自动启动时，在软件“自启动”页手动安装登录任务；安装程序不会自动开始扫描。
- 升级或卸载前等待写表结束，停止监控并关闭控制板。安装器不会强制终止 Excel。

## 数据安全

程序安装到当前用户目录。配置、日志、备份、AI历史保存在 `%LOCALAPPDATA%\PolishingSentinel`，通过目录联接与现有运行引擎兼容。卸载移除程序、快捷方式和本安装实例的计划任务，默认保留所有用户数据。实验主表和U盘文件不会由卸载器删除。

API Key 由 Windows 当前用户加密。更换电脑或用户后需重新输入。AI历史包含实验内容，分享前请检查。AI建议需实验验证，不代表可靠预测或因果结论。

安装包未数字签名，Windows 可能提示未知发布者。请核对发布页 SHA256。当前版本为测试版；尚未在独立新电脑验证全部驱动和 Excel 版本组合。

## 从源码构建

在 Windows 上安装 Inno Setup 6，使用 PowerShell 7 执行 `packaging/Build.ps1 -Compiler <ISCC.exe路径>`。构建生成发布白名单，只包含根目录程序脚本、启动器和本文档。禁止把运行目录的 config、state、logs、temp、backups 提交到仓库。
