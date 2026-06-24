# 批量重命名功能 - 集成 ReNamer .rnp 预设文件解析

## 摘要

在 openxlsx 的剪贴板监听服务中集成 ReNamer (.rnp) 预设文件解析能力，通过剪贴板特殊命令触发批量文件重命名，支持 Replace/Remove/RegEx 三种规则类型。

## 背景

- 当前 openxlsx 项目仅有 Excel 搜索书名的单一功能
- 日常需要批量清理下载目录中电子书文件名的杂乱标记（公众号广告、网站水印、多余符号等）
- ReNamer（Windows 批量重命名工具）已有成熟的 .rnp 预设文件积累，但需要手动打开软件加载执行
- 通过集成到剪贴板服务，复制即重命名，减少操作步骤

## 需求

### 功能需求

1. **解析 .rnp 预设文件**：支持 ReNamer 导出的 .rnp 文件中 `Remove`、`Replace`、`RegEx` 三种规则
2. **URL 解码**：.rnp 中 Config 值为 URL 编码格式，需解码还原
3. **剪贴板触发**：复制 `批量重命名 <目录路径>` 触发批量重命名
4. **配置文件集成**：.rnp 文件路径通过 `config.ini` 的 `renamerpreset` 字段指定
5. **批量执行**：遍历目录下所有文件，对每个文件名依次应用全部启用规则
6. **更改日志**：控制台输出原文件名 → 新文件名对照表
7. **安全保护**：不覆盖已存在的文件

### 非功能需求

- **Unicode 安全**：所有文件路径和文件名使用宽字符 Windows API
- **编码**：`config.ini` 路径支持相对路径（相对于 bin/ 运行目录）
- **性能**：目录文件数 < 10000 时无明显延迟
- **兼容性**：.rnp 格式与 ReNamer 7.5 版本兼容

## 设计

### 架构变更

无架构级别变更。在现有 `listern_clipboard_1.cpp` 中新增以下函数：

| 函数 | 职责 |
|---|---|
| `UrlDecode(const std::wstring&) -> std::wstring` | 将 URL 编码的 UTF-8 字节解码为宽字符串 |
| `ParseRnpRules(const std::wstring&) -> std::vector<RnpRule>` | 解析 .rnp 文件为规则列表 |
| `ApplyRules(const std::wstring&, const std::vector<RnpRule>&) -> std::wstring` | 对单个文件名应用全部规则 |
| `BatchRename(const std::wstring&, const std::wstring&)` | 遍历目录 + 执行重命名 + 输出日志 |

### 数据结构

```cpp
enum class RnpRuleType { Remove, Replace, RegEx };

struct RnpRule {
    RnpRuleType type;        // 规则类型
    bool enabled;            // Marked=1
    std::wstring matchText;  // TEXT / TEXTWHAT / EXPRESSION（解码后）
    std::wstring replaceText;// TEXTWITH / REPLACE（解码后）
    bool skipExtension;      // SKIPEXTENSION=1
    bool caseSensitive;      // CASESENSITIVE=1
    bool useWildcards;       // USEWILDCARDS=1（仅 Remove/Replace）
    bool wholeWordsOnly;     // WHOLEWORDSONLY=1
    int which;               // WHICH: 1=first, 2=last, 3=all
};
```

### 数据流

```
config.ini
  → 读取 renamerpreset 字段
  → 拼接为 .rnp 文件绝对路径

剪贴板 "批量重命名 E:\下载"
  → 过滤链检测 "批量重命名 " 前缀
  → 提取目录路径 "E:\下载"
  → 调用 BatchRename(dir, rnpPath)

BatchRename() {
  1. ParseRnpRules() → 解析 .rnp → RnpRule[]
  2. FindFirstFile / FindNextFile 遍历目录
  3. 对每个文件：
     a. 分离 文件名.扩展名
     b. ApplyRules(文件名, 规则) → 新文件名
     c. 如新文件名 ≠ 原文件名且目标不存在 → MoveFileExW()
     d. 输出 "原 → 新" 日志
}

ParseRnpRules() {
  1. 按行读取 .rnp 文件
  2. 检测 [RuleN] 起始新规则
  3. 解析 ID= → 规则类型
  4. 解析 Config= → URL 解码 → key:value 拆分
  5. 解析 Marked= → 是否启用
}
```

### 接口变更

#### config.ini 新增字段

```ini
autorefreshfile=..\src\rungetfilenamefrompath.py
illegalcharsfile=..\src\illegal_chars.txt
renamerpreset=..\Presets\常用.rnp    ; 新增：ReNamer 预设文件路径
```

#### 剪贴板触发协议

| 复制内容 | 当前行为 | 新增行为 |
|---|---|---|
| `刷新My Kindle Content` | 刷新 Kindle 目录 | 不变 |
| `批量重命名 E:\下载` | — | 触发批量重命名 |
| 其他文本 | 搜索 Excel | 不变 |

### .rnp 格式解析细则

Config 字段解析流程：

```
原始 Config:
  TEXT:%E2%80%9C;WHICH:3;SKIPEXTENSION:1

步骤1：以 ; 分割 → ["TEXT:%E2%80%9C", "WHICH:3", "SKIPEXTENSION:1"]
步骤2：每个段以第一个 : 分割 → {"TEXT": "%E2%80%9C", "WHICH": "3", ...}
步骤3：对值部分 URL 解码 → {"TEXT": "\u201C", "WHICH": "3", ...}
```

#### 规则类型与参数映射

| .rnp ID | 关键 Config 键 | C++ 实现方式 |
|---|---|---|
| `Remove` | `TEXT`, `WHICH`, `SKIPEXTENSION`, `CASESENSITIVE`, `USEWILDCARDS` | `std::wstring::find` / `std::wregex` + 替换为空 |
| `Replace` | `TEXTWHAT`, `TEXTWITH`, `WHICH`, `SKIPEXTENSION`, `CASESENSITIVE` | `std::wstring::find` / `std::wregex` + 替换 |
| `RegEx` | `EXPRESSION`, `REPLACE`, `CASESENSITIVE`, `SKIPEXTENSION` | `std::wregex::replace` |

WHICH 参数行为：
- WHICH=1（第一个）：替换/删除第一个匹配
- WHICH=2（最后一个）：替换/删除最后一个匹配
- WHICH=3（全部）：替换/删除所有匹配（等效 `/g`）

## 实现计划

1. **在 `listern_clipboard_1.cpp` 中新增函数实现**
   - `UrlDecode()` — URL 解码工具函数
   - `ParseRnpRules()` — .rnp 解析
   - `ApplyRules()` — 单文件规则执行
   - `BatchRename()` — 目录遍历 + 重命名 + 日志
2. **修改剪贴板过滤链**：在检测 `刷新My Kindle Content` 同一位置之前，检测 `批量重命名 ` 前缀
3. **修改 config.ini 读取逻辑**：增加 `renamerpreset` 字段解析
4. **更新 `bin/config.ini`**：添加默认配置
5. **编译验证**：完整构建确认无错误

## 验证方案

1. **单元验证**：使用 `bin/20250304.xlsm` 旁测试目录，放几个测试文件
2. **功能验证**：
   - 复制 `批量重命名 E:\test_dir`
   - 观察控制台输出改名对照表
   - 确认文件已被正确重命名
3. **边界情况**：
   - .rnp 文件不存在 → 报错但不崩溃
   - 目录不存在 → 报错
   - 无匹配规则 → 不修改文件
   - 目标文件名已存在 → 跳过
4. **回滚**：保留 `illegal_chars.txt` 模式，不改动原文件

## 风险评估

| 风险 | 缓解措施 |
|---|---|
| URL 解码不完整导致规则失效 | 对照 ReNamer 7.5 解析结果逐条验证 |
| 正则语法差异（C++ vs ReNamer） | 优先使用 TEXT 匹配，RegEx 仅支持基础正则 |
| 文件名编码问题 | 统一使用宽字符 API |
| 重命名不可逆 | 先输出对照表日志，确认后再执行 `MoveFileW()` |

## 修订记录

| 日期 | 版本 | 变更说明 |
|---|---|---|---|
| 2026-06-23 | v1.0 | 初始版本：需求、设计、实现计划
