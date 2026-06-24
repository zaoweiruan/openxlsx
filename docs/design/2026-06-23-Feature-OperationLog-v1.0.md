# 操作日志记录 — 批量重命名 & 刷新目录文件日志

## 摘要

为批量重命名（`BatchRename`）和刷新 Kindle Content 目录（`RunPythonAndShowResult`）两个功能增加文件日志记录能力，复用现有的 `-l`/`--logTofile` 日志框架，使所有操作均有据可查。

## 背景

- 当前只有 `search_keyword`（Excel 搜索书名）在 `-l` 开启时会写入日志文件
- `BatchRename` 和 `RunPythonAndShowResult` 仅输出到控制台（`std::wcout`）
- 用户需要统一的文件日志，便于事后追溯批量重命名历史、Python 脚本执行状态

## 需求

### 功能需求

1. **批量重命名日志**：当 `-l` 启用时，将重命名操作记录逐条写入日志文件
2. **刷新目录日志**：当 `-l` 启用时，将 Python 脚本执行开始、错误、完成状态写入日志文件
3. **日志格式一致**：与现有搜索日志共用同一输出文件（`cfg.output_path`，默认 `search_log.txt`）

### 非功能需求

- **零侵入**：不改变现有 console 输出行为
- **零性能影响**：未启用 `-l` 时不产生文件 I/O

## 设计

### 架构变更

无架构变更。在 `listern_clipboard_1.cpp` 中新增一个辅助函数并修改两个现有函数：

| 变更 | 说明 |
|---|---|
| 新增 `AppendLogLine(const std::wstring&)` | 当 `cfg.logTofile` 为 true 时追加一行到日志文件 |
| 修改 `RunPythonAndShowResult()` | 操作开始/完成时调用 `AppendLogLine` |
| 修改 `BatchRename(...)` | 每条重命名结果同步写入日志 |

### 数据流

```
AppendLogLine(line)
  → if (!cfg.logTofile) return
  → WideCharConvertMultiByte(line, CP_UTF8)
  → ofstream(cfg.output_path, ios::app) << utf8 << endl
  → close
```

### 日志格式

```
=== 刷新目录: 开始 ===
=== 刷新目录: 完成 ===
=== 批量重命名: <目录路径> ===
  ✓ 原名 → 新名
  ✗ 失败: 原名 → 新名 (错误: N)
=== 完成: 成功 N 个, 跳过 M 个 ===
```

## 实现计划

1. 在 `log_file` / `search_keyword` 函数附近添加 `AppendLogLine` 辅助函数
2. 在 `RunPythonAndShowResult` 中插桩：开始 / 错误 / 完成
3. 在 `BatchRename` 中将 `std::wcout` 调用同步追加 `AppendLogLine`
4. 更新 `docs/index.md`

## 验证方案

1. `cmake --build` 零警告零错误
2. 启动时带 `-l search_log.txt`，触发批量重命名，检查日志文件内容
3. 启动时带 `-l search_log.txt -a`，检查刷新目录日志

## 风险评估

- 日志文件累积增长（用户需定期清理）
- 无侵入性风险——纯追加写入，不改变既有行为

## Changelog

| 日期 | 版本 | 变更 |
|---|---|---|
| 2026-06-23 | v1.0 | 初始版本 |
