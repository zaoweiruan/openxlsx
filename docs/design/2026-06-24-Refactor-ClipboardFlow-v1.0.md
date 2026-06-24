# 剪贴板处理流重构设计文档

## 摘要
将 `GetClipboardText()` 拆分为「原始读取」和「搜索过滤」两层，在 `WndProc` 中先将特殊命令（刷新 Kindle Content、批量重命名）分支到对应的处理函数，再对搜索路径应用字符过滤，避免 `illegal_chars.txt` 中的 `\` 等字符阻塞批量重命名指令。

## 背景
- `illegal_chars.txt` 第 1 行为 `\`，Windows 路径必然包含反斜杠
- 用户复制 `批量重命名 "E:\文件存储卷_1\My Kindle Content"` 到剪贴板
- `GetClipboardText()` 调用 `ContainsIllegalChar(text, cfg.illegalChars)` 发现 `\` → 返回空字符串
- `WndProc` 在 line 662 检查 `!clipboardText.empty()` 为 false → 程序无任何响应

## 需求
### 功能需求
1. 特殊命令（刷新、批量重命名）必须能正确识别和触发，不受非法字符过滤影响
2. 普通搜索依然需要保留现有过滤链（非法字符、IP/域名排除）

### 非功能需求
- 不修改现有命令行参数和行为
- 不增加新的外部依赖
- 后续新增特殊命令时可方便地在 WndProc dispatch 中添加

## 设计
### 接口变更
| 旧函数 | 新函数 | 说明 |
|---|---|---|
| `GetClipboardText()` | → 拆分为 `GetRawClipboardText()` + `FilterForSearch()` | 读取与过滤分离 |
| (WndProc 直接使用 GetClipboardText) | → WndProc 先检查特殊命令，再过滤 | 调度顺序变更 |

### 数据流

```
                    GetRawClipboardText()
                              │
                    ┌─────────▼──────────┐
                    │  rawText empty?     │──→ return (nothing in clipboard)
                    └─────────┬──────────┘
                              │
                    ┌─────────▼──────────┐
                    │  duplicate? (last)  │──→ return (same as previous)
                    └─────────┬──────────┘
                              │
                    ┌─────────▼──────────┐
                    │  == 刷新My Kindle   │──→ RunPythonAndShowResult()
                    └─────────┬──────────┘
                              │
                    ┌─────────▼──────────┐
                    │  match 批量重命名?   │──→ BatchRename(dirPath)
                    └─────────┬──────────┘
                              │
                    ┌─────────▼──────────┐
                    │  FilterForSearch()  │
                    │  ┌────────────────┐ │
                    │  │ 非法字符过滤    │ │  ← 只有搜索路径才应用
                    │  │ 全数字忽略      │ │
                    │  │ 中文 trim       │ │
                    │  │ IPv4/域名忽略   │ │
                    │  └────────────────┘ │
                    └─────────┬──────────┘
                              │
                    ┌─────────▼──────────┐
                    │  search_keyword()   │
                    └────────────────────┘
```

### 函数拆分细节

```cpp
// 1. 纯读取，无过滤
std::wstring GetRawClipboardText() {
    // OpenClipboard → CF_UNICODETEXT / CF_TEXT → return raw text
    // 不含任何 ContainsIllegalChar / has_chinese / isValidIPv4 调用
}

// 2. 静态过滤逻辑（从原 GetClipboardText 提取）
std::wstring FilterForSearch(const std::wstring& rawText) {
    if (rawText.empty()) return L"";
    // 含非法字符 → 返回空
    // 全数字/空白/./- → 返回空
    // 含中文 → trim
    // IPv4/域名 → 返回空
    return filtered;
}
```

## 实现计划
1. 从 `GetClipboardText()` 提取的 clipboard 读取逻辑（OpenClipboard/GetClipboardData/CloseClipboard）保留在同一函数中，改名为 `GetRawClipboardText()`，**移除**所有过滤调用
2. 将过滤逻辑（ContainsIllegalChar、has_chinese、isValidIPv4、domainExists）提取为新函数 `FilterForSearch(const std::wstring&)`
3. 修改 `WndProc` 中 `WM_CLIPBOARDUPDATE` handler：
   - `GetRawClipboardText()` → 判空 → 判重
   - 先检查 `刷新My Kindle Content` 和批量重命名前缀
   - 都不匹配时，调用 `FilterForSearch()` 再 `search_keyword()`
4. 运行诊断日志（`[DEBUG] 收到剪贴板`）应输出**原始**文本而非过滤后文本，便于调试

## 验证方案
- 编译通过（cmake --build）
- 手动测试三个场景：
  1. 复制 `批量重命名 "E:\文件存储卷_1\My Kindle Content"` → 触发 BatchRename
  2. 复制 `刷新My Kindle Content` → 触发 Python 脚本
  3. 复制正常书名（含中文） → 触发 search_keyword
- 运行时日志应显示 `[DEBUG] 收到剪贴板: 批量重命名 "E:\文件..."`（原始文本）

## 风险评估
- 仅限于 `GetClipboardText()` 和 `WndProc` 的调度逻辑改动，不影响 Excel 搜索、日志写入等其他模块
- 回滚方案：`git checkout -- src/listern_clipboard_1.cpp`

## Changelog

| 日期 | 版本 | 变更 |
|---|---|---|
| 2026-06-24 | v1.0 | 初始版本 |
