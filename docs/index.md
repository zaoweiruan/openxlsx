# openxlsx 文档索引

| 文档 | 描述 |
|------|------|
| [architecture-analysis.md](architecture-analysis.md) | 项目架构分析 |
| [batch-rename-usage.md](batch-rename-usage.md) | 批量重命名功能使用说明 |

## 设计文档

| 文档 | 描述 |
|------|------|
| [2026-06-23-Spec-BatchRename-v1.0.md](spec/2026-06-23-Spec-BatchRename-v1.0.md) | 批量重命名功能设计规格 |
| [2026-06-23-Feature-OperationLog-v1.0.md](design/2026-06-23-Feature-OperationLog-v1.0.md) | 操作日志（AppendLogLine）功能设计规格 |
| [2026-06-24-Refactor-ClipboardFlow-v1.0.md](design/2026-06-24-Refactor-ClipboardFlow-v1.0.md) | 剪贴板处理流重构：特殊命令优先于搜索过滤 |

## Bugfix 文档

| 文档 | 描述 |
|------|------|
| [2026-06-24-Bugfix-BatchRenameQuote-v1.0.md](bugfix/2026-06-24-Bugfix-BatchRenameQuote-v1.0.md) | 修复路径双引号导致批量重命名失效 |
| [2026-06-24-Bugfix-ConfigEncodingUtf8-v1.0.md](bugfix/2026-06-24-Bugfix-ConfigEncodingUtf8-v1.0.md) | 修复 config.ini UTF-8 编码读取乱码 |
| [2026-06-24-Bugfix-ConfigCrlfParsing-v1.0.md](bugfix/2026-06-24-Bugfix-ConfigCrlfParsing-v1.0.md) | 修复 CRLF 行尾导致配置值带 `\r` 引起文件打开失败 |
