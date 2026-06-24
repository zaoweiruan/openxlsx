# config.ini CRLF 行尾导致配置值带 `\r` 引起文件打开失败

## 摘要
修复 `bin/config.ini` 使用 CRLF 行尾时，`std::getline` 保留 `\r` 导致 `LoadConfig()` 解析的路径值尾部带回车字符，进而使 `_wfopen_s` 打开 `illegal_chars.txt` 等配置文件失败的问题。

## 背景
- 项目使用 `bin/config.ini` 存储运行时配置（如 `illegalcharsfile`、`renamerprefix`、`renamerpreset` 等）
- `bin/config.ini` 使用 Windows 标准的 CRLF（`0D 0A`）换行
- `LoadConfig()` 使用 `std::getline` 逐行解析，`getline` 以 `\n` 为行结束符，**会移除 `\n` 但保留 `\r`**
- 此前代码中 `codecvt_utf8_utf16` 存在 segfault 问题（已修复），修复后 `LoadIllegalChars()` 的 `_wfopen_s` 报错
- 但最终根因是 `std::getline` 的 CRLF 行尾处理——config 值尾部残留 `\r` 字符，导致文件名参数不合法

## 需求
### 功能需求
- `LoadConfig()` 解析 `config.ini` 后，所有 key 和 value 不应包含行尾 `\r` 字符
- `illegalcharsfile` 值解析后应为裸路径（如 `..\src\illegal_chars.txt`），不含尾随 `\r`
- `renamerprefix`、`renamerpreset` 等其他配置值同理

### 非功能需求
- 向后兼容：`config.ini` 为 LF 结尾时不受影响（`std::getline` 已正确提取，`\r` 不存在时无事可做）
- 纯配置解析层修复，不影响其他模块

## 设计
### 架构变更
无。仅在 `LoadConfig()` 函数内添加一行 key/value 尾随 `\r` 剥离逻辑。

### 接口变更
无。

### 数据流
```
config.ini 行: "illegalcharsfile=..\src\illegal_chars.txt\r\n"
  → std::getline → "illegalcharsfile=..\src\illegal_chars.txt\r"  (含\r)
  → 注释剥离 → key="illegalcharsfile", value="..\src\illegal_chars.txt\r"
  → key.pop_back() / value.pop_back() 剥离 \r → key="illegalcharsfile", value="..\src\illegal_chars.txt"
  → LoadIllegalChars(value)  → _wfopen_s 成功
```

## 实现计划
1. 在 `LoadConfig()` 中 `std::getline` 后的注释剥离逻辑之后，添加 key 和 value 的尾随 `\r` 检查及移除
   ```cpp
   if (!key.empty() && key.back() == L'\r') key.pop_back();
   if (!value.empty() && value.back() == L'\r') value.pop_back();
   ```
2. 执行完整构建验证无编译错误
3. 验证 `bin/config.ini` 加载后 `illegal_chars.txt` 可正常打开

## 验证方案
- 启动程序后检查控制台是否有 `非法字符文件打开失败` 错误
- 分析日志确认 `LoadIllegalChars` 正常加载
- 验证批量重命名功能可用（依赖正确的 `renamerprefix`/`renamerpreset` 值）
- 特别验证：使用 CRLF 行尾的 `config.ini` 测试

## 风险评估
- 修改极小（仅 2 行 key/value 的 `\r` 检查）
- 仅在 value 或 key 尾部确实存在 `\r` 时才执行 `pop_back`，无副作用
- LF 行尾文件不受影响（`\r` 不存在，`pop_back` 不会执行）
- 回滚：直接删除对应 2 行代码即可

## Changelog

| 日期 | 变更 | 关联代码变更 |
|------|------|-------------|
| 2026-06-24 | 初始版本 v1.0 | `src/listern_clipboard_1.cpp`: LoadConfig() 添加 key/value 尾随 `\r` 剥离 |
