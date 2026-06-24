# UTF-8 配置文件编码修复

## 摘要
剪贴板批量重命名指令无法触发，根因是 `LoadConfig()` 读取 `config.ini` 时 MinGW 的 `codecvt_utf8_utf16` 实现缺陷（GCC 14+ 模板要求），导致中文前缀 `批量重命名` 读取失败，进而前缀匹配永不成功。

## 背景
- `config.ini` 采用 UTF-8 without BOM 存储（4 行，含中文 `批量重命名`、`常用.rnp`）
- MinGW GCC 14+ 要求 `std::filesystem::path` 具有 `make_preferred()` 成员，`std::wstring` 不满足
- `std::codecvt_utf8_utf16` 模板参数不被 MinGW 支持
- Windows API `MultiByteToWideChar(CP_UTF8)` 是可靠方案

## 背景
- `config.ini` 采用 UTF-8 without BOM 存储
- MinGW `std::wifstream` 默认使用 C locale 或系统 locale（GBK）
- Windows 剪贴板提供 UTF-16LE 文本（CF_UNICODETEXT）
- 配置文件中文字符 → 单字节乱码码点 → 前缀匹配永失败

## 需求

### 功能需求
- 修复 `LoadConfig()` 读取 UTF-8 配置文件

### 非功能需求
- 无性能影响
- 保持向后兼容（ANSI 编码的 config.ini 仍能读）

## 设计

### 架构变更
- `LoadConfig()` 文件流打开后立即 `imbue` UTF-8 locale

### 接口变更
- 新增 `std::codecvt_utf8<wchar_t>` imbue

### 数据流
```
config.ini(UTF-8) → std::wifstream → imbue(codecvt_utf8) → std::wstring 正确解码
```

## 实现计划
1. 在 `LoadConfig()` 第 78 行后添加 imbue
2. 编译验证

## 验证方案
拷贝 `批量重命名 "E:\文件存储卷_1\My Kindle Content"` → 日志显示 `[DBG] 前缀匹配成功`

## 风险评估
- `codecvt_utf8` 在 GCC 需 `<codecvt>` 头，若缺失则替换为手写 UTF-8 解码函数

## Changelog
- 2026-06-24: 创建 v1.0，修复 LoadConfig() UTF-8 编码读取