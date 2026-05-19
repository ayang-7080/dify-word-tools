# dify-word-tools

面向 Dify 的 Word 文档处理插件，提供 `.docx` 生成、解析、转换、模板渲染和常见格式处理能力。

## 功能概览

- 根据 JSON 数据生成 `.docx`，支持段落、表格、图片、批注、页码、行号、分页和换行。
- 将 `.docx` 提取为纯文本、Markdown 或结构化 JSON。
- 从 `.docx` 中提取图片。
- 基于 Word 模板和 `docxtpl` 渲染新文档。
- 对已有 Word 文档执行文本标蓝、正则标蓝、添加批注、设置上标等处理。
- 提供文本内容转 `.txt` 文件能力。

## 项目信息

| 项目 | 内容 |
| --- | --- |
| 插件名 | `docx-tools` |
| Dify 展示名 | `ay-office-word` |
| 作者 | `ayang` |
| 当前版本 | `0.2.4` |
| 插件类型 | Tool Plugin |
| 运行时 | Python 3.12 |

## 目录结构

```text
.
├── main.py                         # Dify 插件入口
├── manifest.yaml                   # 插件清单
├── provider/
│   ├── docx-tools.yaml             # 工具 Provider 配置
│   └── docx-tools.py               # Provider 凭据校验入口
├── tools/
│   ├── tools.py                    # Word 生成和解析公共能力
│   ├── docx-tools.py               # JSON 生成 docx
│   ├── docx-to-json.py             # docx 转结构化 JSON
│   ├── docx-to-json-text.py        # docx 转 JSON，图片以索引占位
│   ├── docx-to-json-image.py       # 提取 docx 图片
│   ├── docx-to-text.py             # 提取纯文本
│   ├── docx-markdown.py            # docx 转 Markdown
│   ├── content-to-file.py          # 内容转文件
│   ├── expand/                     # 标蓝、批注、上标等扩展工具
│   └── docxtpl/                    # 模板渲染工具
└── _assets/icon.svg                # 插件图标
```

`test/` 目录用于本地样例和手工验证，默认不提交到仓库。

## 已注册工具

| 工具名 | 说明 |
| --- | --- |
| `docx-tools` | 根据 JSON 数组生成 Word 文档 |
| `content-to-file` | 将文本内容输出为 `.txt` 文件 |
| `docx-markdown` | 使用 `mammoth` 将 Word 转为 Markdown |
| `docx-to-json` | 将 Word 内容解析为文本、表格、图片和样式 JSON |
| `docx-to-json-text` | 解析 Word 内容，图片内容替换为图片索引 |
| `docx-to-json-image` | 提取 Word 中的图片，可按下标提取 |
| `docx-to-text` | 提取 Word 段落纯文本 |
| `regex-highlight` | 按内置类型或自定义正则标蓝文本 |
| `text-highlight` | 按指定文本列表标蓝、加粗或斜体 |
| `add-comments` | 给指定文本添加 Word 批注 |
| `set-superscript` | 将引用编号等文本设置为上标 |
| `template-create-word` | 基于 `.docx` 模板和 JSON 数据生成 Word |

`tools/file-pandoc.yaml` 当前在 Provider 配置中被注释，默认不会注册。

## 安装依赖

建议使用 Python 3.12，与 `manifest.yaml` 中的运行时保持一致。

```bash
pip install -r requirements.txt
```

主要依赖：

- `dify_plugin`
- `python-docx`
- `mammoth`
- `pypandoc`
- `docxtpl`
- `Pillow`

## 本地调试

复制环境变量模板：

```bash
cp .env.example .env
```

根据 Dify 插件调试页面填写：

```env
INSTALL_METHOD=remote
REMOTE_INSTALL_URL=debug.dify.ai:5003
REMOTE_INSTALL_KEY=your-debug-key
```

启动插件：

```bash
python -m main
```

启动后，在 Dify 插件管理页面刷新即可看到调试中的插件。

## 生成 Word 示例

`docx-tools` 的核心参数是 `output_data`，它必须是一个合法 JSON 字符串，解析后为对象数组。数组中的每一项按顺序写入 Word。

基础示例：

```json
[
  {
    "data": "文档标题",
    "font_size": 16,
    "bold": true,
    "alignment": "center"
  },
  {
    "data": "这是第一段正文，支持字体、字号、颜色、行距、首行缩进等样式。",
    "font_name": "宋体",
    "font_size": 12,
    "line_spacing": 1.5,
    "first_line_indent": 24
  },
  {
    "type": "table",
    "style": "Table Grid",
    "alignment": "center",
    "data": "| 姓名 | 部门 |\n| --- | --- |\n| 张三 | 研发 |\n| 李四 | 产品 |"
  }
]
```

常用字段：

| 字段 | 说明 |
| --- | --- |
| `type` | 内容类型，支持 `paragraph`、`table`、`table2`、`image`、`float_image` |
| `data` | 内容数据，段落为文本，表格为 Markdown 或二维数组，图片为 URL 或 Data URL |
| `font_name` | 字体，默认 `宋体` |
| `font_size` | 字号，默认 `12` |
| `bold` | 是否加粗 |
| `alignment` | 对齐方式，支持 `left`、`center`、`right`、`justify` |
| `color` | 字体颜色，如 `#000000` |
| `bg_color` | 背景高亮颜色 |
| `line_spacing` | 行距 |
| `first_line_indent` | 首行缩进 |
| `space_after` | 段后距 |
| `break_type` | `page` 插入分页符，`line` 插入换行 |
| `comment` | Word 批注，支持单个对象或对象数组 |

更完整的参数说明见 [docx-tools-使用说明.md](docx-tools-使用说明.md)。

## 打包插件

安装 Dify 插件 CLI 后，在项目根目录执行：

```bash
dify-plugin plugin package .
```

也可以指定输出文件：

```bash
dify-plugin plugin package . -o docx-tools.difypkg
```

## 发布

项目内置 GitHub Actions 工作流：

```text
.github/workflows/plugin-publish.yml
```

发布 GitHub Release 后，工作流会打包插件，并向目标 `dify-plugins` 仓库创建更新 PR。使用前需要在 GitHub 仓库中配置 `PLUGIN_ACTION` secret。

## 隐私

隐私说明见 [PRIVACY.md](PRIVACY.md)。
