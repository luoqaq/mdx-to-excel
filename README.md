
# MDX to Excel 转换工具

一个用于将MDX和MD文件内容转换为Excel表格的命令行工具。该工具可以批量处理MDX和MD文件，自动提取标题和内容，并生成结构化的Excel文件。

## 功能特性

- 基于Bun运行时开发，提供极致的性能体验
- 支持批量处理MDX文件
- 自动提取一级标题和二级标题
- 智能分割超长内容，确保Excel单元格不会超出限制
- 生成详细的转换日志
- 支持忽略指定目录
- 自定义源目录和输出目录

## 环境要求

- [Bun](https://bun.sh) 1.0.0 或更高版本

## 安装

```bash
bun install
```

## 使用方法

```bash
bun run index.ts [options]
```

### 命令行参数

- `-s, --source <dir>`: 指定MDX文件目录（默认：./content）
- `-o, --output <dir>`: 指定输出Excel文件目录（默认：./excel）
- `-i, --ignore <dirs...>`: 指定要忽略的目录（可指定多个）

### 示例

```bash
# 使用默认配置
bun run index.ts

# 指定源目录和输出目录
bun run index.ts -s ./docs -o ./output

# 忽略特定目录
bun run index.ts -i drafts temp
```

## 工作原理

1. 扫描指定目录下的所有.mdx和.md文件
2. 解析每个文件的frontmatter和内容
3. 提取一级标题（# 开头）和二级标题（## 开头）
4. 将内容按标题组织，超长内容自动分割
5. 生成Excel文件，每行包含标题和对应内容
6. 记录详细的转换日志

## 日志功能

- 日志文件保存在`./logs`目录下
- 日志文件名格式：`conversion-{timestamp}.log`
- 记录每个文件的处理状态和错误信息
- 使用不同图标标识信息类型（✨ 信息，❌ 错误）

## 输出文件

- Excel文件保存在指定的输出目录（默认：./excel）
- 文件名格式：`output-{timestamp}.xlsx`
- 包含"MDX Content"工作表，列出所有处理后的内容

## 注意事项

- Excel单元格有字符限制（32767字符），超长内容会自动分割
- 分割时会优先在段落之间进行，保持内容的完整性
- 建议在运行前备份重要的MDX文件