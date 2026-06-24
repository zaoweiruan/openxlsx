# 批量重命名功能使用说明

## 概述

openxlsx 的批量重命名功能通过解析 ReNamer (.rnp) 预设文件，快速对指定目录下的文件执行批量重命名操作。复制一条命令即可触发，无需打开额外软件。

## 触发方式

复制以下格式的文本到剪贴板：

```
批量重命名 <目标目录路径>
```

例如：

```
批量重命名 E:\下载
批量重命名 D:\电子书\待整理
批量重命名 C:\Users\xxx\Desktop\新建文件夹
```

当剪贴板检测到 `批量重命名 `（注意末尾空格）前缀时，会自动调用重命名引擎处理指定目录。

## 规则文件配置

批量重命名使用 ReNamer (.rnp) 格式的预设文件定义替换规则。路径在 `config.ini` 中配置：

```ini
renamerpreset=..\Presets\常用.rnp
```

路径相对于 `bin/` 运行目录。默认指向 `Presets\常用.rnp`。

## .rnp 规则文件格式

.rnp 文件是 ReNamer 导出的预设文件（INI 格式），每个规则由 `[RuleN]` 区间定义：

```ini
[Rule1]
ID=Remove               ; 规则类型：Remove / Replace / RegEx
Marked=1                ; 1=启用, 0=禁用
Config=TEXT:%E6%B5%8B;WHICH:3;SKIPEXTENSION:1;WHOLEWORDSONLY:0;CASESENSITIVE:0;USEWILDCARDS:0
```

### 规则类型

| 类型 | ID | 说明 | Config 关键字段 |
|------|-----|------|----------------|
| 移除 | `Remove` | 删除匹配文本 | `TEXT`=要删除的文本, `WHICH`=第几个(1/2/3) |
| 替换 | `Replace` | 替换匹配文本 | `TEXT`=查找文本, `TEXTWITH`=替换为, `WHICH`=第几个(1/2/3) |
| 正则 | `RegEx` | 正则表达式替换 | `EXPRESSION`=正则表达式, `REPLACE`=替换为, `WHICH`=第一个(1)/全部(3) |

### Config 字段说明

Config 字段使用 URL 编码（UTF-8），解析后键值对以 `;` 分隔：

| 键 | 含义 | 取值 |
|----|------|------|
| `TEXT` | 要删除/查找的文本（Remove/Replace） | URL 编码文本 |
| `TEXTWHAT` | 查找文本（Replace 的别名） | URL 编码文本 |
| `EXPRESSION` | 正则表达式（RegEx） | URL 编码文本 |
| `REPLACE` | 替换为（RegEx） | URL 编码文本 |
| `TEXTWITH` | 替换为（Replace） | URL 编码文本 |
| `WHICH` | 匹配位置 | 1=第一个, 2=最后一个, 3=全部 |
| `SKIPEXTENSION` | 跳过扩展名 | 1=是, 0=否 |
| `CASESENSITIVE` | 区分大小写 | 1=是, 0=否 |
| `USEWILDCARDS` | 使用通配符 | 1=是, 0=否（仅 Remove/Replace） |
| `WHOLEWORDSONLY` | 仅完整单词 | 1=是, 0=否 |

## 示例：创建 Preset

创建 `Presets\常用.rnp`：

```ini
[Rule1]
ID=Remove
Marked=1
Config=TEXT:%E5%85%AC%E4%BC%97%E5%8F%B7;WHICH:3;SKIPEXTENSION:1
; 移除"公众号"（UTF-8 URL 编码）

[Rule2]
ID=Replace
Marked=1
Config=TEXT:www.;TEXTWITH:;WHICH:3;SKIPEXTENSION:1
; 删除所有"www."前缀

[Rule3]
ID=RegEx
Marked=1
Config=EXPRESSION:%5C%5B.%2A%3F%5C%5D;REPLACE:;WHICH:3
; 正则删除所有 [方括号内容]
```

### URL 编码工具

Config 中的中文/特殊字符需做 UTF-8 URL 编码。可使用在线工具或以下 Python 脚本生成：

```python
import urllib.parse
text = "公众号"
encoded = urllib.parse.quote(text)
print(encoded)  # %E5%85%AC%E4%BC%97%E5%8F%B7
```

## 运行输出示例

```
正在解析规则文件: E:\...\Presets\常用.rnp
共加载 3 条规则，启用 3 条
----- 重命名对照表 -----
  ✓ [公众号]测试文件名.pdf
    → 测试文件名.pdf
  ✓ www.example.txt
    → example.txt
  ✗ 失败: locked.docx → clean.docx (错误: 5)
----- 完成: 成功 2 个, 跳过 0 个 -----
```

错误码 5 表示文件被其他程序占用（`ERROR_ACCESS_DENIED`）。

## 安全说明

- 不覆盖已存在的文件（目标文件名冲突时跳过并提示）
- 不处理子目录（仅当前目录下文件）
- 不改动文件扩展名（基于 `SKIPEXTENSION` 约定）
- 重命名前需确保文件未被其他程序占用
- 建议先在测试目录验证规则效果

## 实现参考

详见设计文档：[2026-06-23-Spec-BatchRename-v1.0.md](../2026-06-23-Spec-BatchRename-v1.0.md)
