# openxlsx 项目架构分析

## 一句话概括

Windows 剪贴板监听后台服务，通过 OpenXLSX 读取 `.xlsm`/`.xlsx` 中的书籍列表，自动比对剪贴板内容判断书籍是否已下载，附带 IP/域名过滤、Python VBA 宏触发等辅助功能。

## 核心架构与数据流

```
[用户操作: Ctrl+C] → 剪贴板更新 → WM_CLIPBOARDUPDATE
                                      ↓
                              GetClipboardText()
                              ┌────────────────┐
                              │ 剪贴板内容过滤  │  ← 全空白/全数字/非法字符/IPv4/域名 → 丢弃
                              │ 有中文→trim()  │  ← IPv4或域名→丢弃
                              └───────┬────────┘
                                      ↓ (有效文本)
                              ┌────────────────┐
                              │ search_keyword │  ← 调用 OpenXLSX 打开 Excel
                              │ 遍历所有sheet   │     逐行比对指定列
                              │ 匹配→记录结果   │     处理Datetime类型
                              └───────┬────────┘
                                      ↓
                              控制台/文件日志输出
```

## 关键模块

### 1. 主入口 (`src/listern_clipboard_1.cpp:725`)

```
wWinMain(...) → setup_utf8_console() → parse_command_line() → LoadConfig()
→ WSAStartup() → StartClipboardListener() → 消息循环
```

使用 **Win32 隐藏消息窗口 + `AddClipboardFormatListener`** 监听剪贴板，而非轮询。

### 2. 剪贴板过滤逻辑 (`GetClipboardText()`, 行 561-632)

优先读取 `CF_UNICODETEXT`，降级到 `CF_TEXT`（ANSI→Unicode 转换）。

过滤链（顺序敏感）：
1. **全新空白/数字/点/减号过滤**（行 617-619）：`iswspace || iswdigit || '.' || '-'` 全匹配 → 丢弃
2. **非法字符**：从 `illegal_chars.txt` 逐字符加载的集合
3. **中文** → trim 保留（可能是书名）
4. **IPv4 (`InetPtonW`) 或域名 (`GetAddrInfoW`)** → 丢弃

### 3. Excel 搜索 (`search_excel_all_sheets`, 行 469-507)

- OpenXLSX `XLDocument::open()` 打开 `.xlsm`/`.xlsx`
- `column_letter_to_index()` 转换列字母 → 数字下标
- `detect_max_row()` 逐行检测直到遇到首个空行（最多 200 万行）
- `find_string_in_sheet()` 支持 **DateTime 字段**处理：检测 `XLValueType::Float` + `isValidExcelDate()` → 转换为 `%Y-%m-%d %H:%M:%S` 字符串再比对
- 遍历 workbook 所有 worksheet

### 4. 配置系统 (`LoadConfig` + 命令行参数)

双层配置：
- **命令行**：`-f <excel路径>`、`-d`(详细日志)、`-l <日志路径>`、`-a` (自动刷新)、`-c <列字母>`
- **配置文件 `bin/config.ini`**（`key=value` 格式）：
  - `autorefreshfile` → Python 脚本路径（win32com 驱动 Excel VBA）
  - `illegalcharsfile` → 非法字符列表文件路径

### 5. Python 辅助脚本 (`src/rungetfilenamefrompath.py`)

```python
win32com.client.Dispatch("Excel.Application")
→ 打开 xlsm → 调用 VBA 宏 `ThisWorkbook.GetFilesFromFolder`
→ 刷新 Kindle Content 目录
```

通过 `_wpopen()` 调用，输出通过管道读回控制台。

### 6. IP/域名过滤

- `isValidIPv4()`: `InetPtonW(AF_INET, ...)` — 系统级验证
- `domainExists()`: `GetAddrInfoW()` — 实际 DNS 查询验证域名是否存在

## 编译目标一览

| CMake 目标 | 源文件 | 输出 | 描述 |
|---|---|---|---|
| `search_excel` | `listern_clipboard_1.cpp` | `bin/search_book_service.exe` | 剪贴板监听主服务 |
| `isValidIPv4orDomainName` | `isValidIPv4orDomainName.cpp` | `bin/isValidIPv4orDomainName.exe` | IP/域名验证测试工具 |

两个目标共享 `CMAKE_CXX_FLAGS`：`-lstdc++ -std=c++17 -g -Wall -O2 -mconsole -municode`

链接库：
- **search_excel**：`OpenXLSX::OpenXLSX` + `boost_algorithm_lib` + `Comctl32` + `Shlwapi` + `user32` + `ws2_32`
- **isValidIPv4orDomainName**：仅 `ws2_32`（通过 `-mconsole -municode` 链接）

## 历史遗留源码（不参与构建）

`src/` 下有大量编号变体文件，均为开发迭代遗迹：

| 模式 | 文件数 | 说明 |
|---|---|---|
| `search_excel_all_sheets*.cpp` | 5 (1-4, 无编号) | CLI 搜索，逐步演进 |
| `search_excel_all_sheets_gui*.cpp` | 3 (1-3) | Win32 GUI 版（Comctl32 ListView） |
| `excel_search_service*.cpp` | 4 (1-3, 无编号) | Windows Service 版 |

这些文件不参与构建，不要误认为多模块架构。

## 注意事项

1. **路径硬编码**：OpenXLSX (`E:/OpenXLSX/`) 和 Boost (`D:/boost_1_88_0/`) 是绝对路径，新机器需修改 `CMakeLists.txt`
2. **Boost 实际只用在一个文件**：`search_excel_all_sheets.cpp` 引用了 `boost/algorithm/string.hpp`，但主目标 `listern_clipboard_1.cpp` 并未使用它（链接了但未引用）
3. **`illegal_chars.txt`** 使用 `fin >> ch` 逐字符读取，会跳过空白字符（空格、换行等），所以不需要在文件中列出空格
4. **无测试**、无 CI、无 lint
5. **必须从 `bin/` 目录运行**，否则找不到 `config.ini`
