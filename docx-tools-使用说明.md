# 生成 docx（docx-tools）使用说明


## 工具做什么？

根据 **`output_data` 中的一段 JSON 数组**，按顺序生成 Word 文档（`.docx`），支持：

- 普通段落（字体、颜色、对齐、行距、首行缩进、段后距、背景高亮等）
- Markdown 或二维数组表格（`table` / `table2`，含三线表样式）
- 嵌入图片（URL 或 Base64）及 Word 批注
- 绝对定位浮动图片
- 分页符、换行
- 可选：行号、页脚页码、统一页边距

执行成功后，工具会返回 **二进制文件消息**（Blob），MIME 为 `application/vnd.openxmlformats-officedocument.wordprocessingml.document`，下载文件名为 `{filename}.docx`。

## 在 Dify 里怎么用？

1. 将本仓库作为 **插件** 安装到 Dify（或按团队流程发布后再安装）。
2. 在工作流或 Agent 中选用 **「生成 docx」** 工具节点。
3. 重点配置 **`output_data`**：必须是 **合法 JSON 的字符串**，其解析结果为 **对象数组**（`[ {...}, {...} ]`）。若 JSON 无效，工具会抛出解析错误。

其余参数见下表。

## 工具参数一览

| 参数名 | 类型 | 必填 | 说明 |
|--------|------|------|------|
| `output_data` | string | 是 | JSON 数组字符串，每一项描述一段内容（段落/表/图等），格式见下文。 |
| `page_padding` | string | 是（YAML 声明） | 页边距，单位 **厘米**，会用于左/右/上/下四边（代码中为 `int(...)`，建议填整数如 `2`）。 |
| `filename` | string | 否 | 生成文件名（**不含扩展名**）；为空则使用时间戳+随机数。 |
| `enable_line` | boolean | 否 | 是否开启行号（默认 `false`）。在 YAML 中为表单字段。 |
| `enable_page` | boolean | 否 | 是否在页脚显示页码（默认 `false`）。 |

## `output_data` 数组：每一项的通用字段

- **`type`**（可选）：内容类型。未写时按段落处理。
  - `paragraph`：普通段落（可省略 `type`，走默认分支）
  - `table`：表格（Markdown 字符串或二维数组）
  - `table2`：增强表格（支持按列对齐、内容加粗、文字颜色等）
  - `image`：行内图片段落
  - `float_image`：页面绝对定位图片
- **`data`**：正文数据。段落为字符串（支持 `\n` 拆成多段）；表格为 Markdown 或 `string[][]`；图片为 **可访问的 `http(s)` URL** 或含 **`base64,`** 的 Data URL。
- **`font_name`**（默认 `宋体`）、**`font_size`**（默认 `12`）、**`bold`**、**`alignment`**（`left` / `center` / `right` / `justify`）
- **`color`**：前景色，如 `#000000`、`#F56C6C`（非法或非 `#` 开头时可能回退为黑色）
- **`bg_color`**：段落背景高亮（十六进制 `#RRGGBB`）；会映射到 Word 内置高亮色，未命中映射时默认偏黄色高亮
- **`line_spacing`**：行距倍数（默认 `1`）
- **`first_line_indent`**：首行缩进，单位按实现为 **磅（pt）** 量级传入（与 `DocxGenerator.add_paragraph` 一致）
- **`space_after`**：段后间距（磅）
- **`break_type`**：在本条内容处理完后插入 **`page`** 分页符或 **`line`** 换行（仅 `line` 时会尽量挂在当前段落上）

### 批注 `comment`

段落与 `image` 支持批注：

- 单条：`{ "text": "批注内容", "author": "可选", "initials": "可选" }`
- 多条：上述对象的数组

### `type: "table"` 额外字段

- **`style`**：Word 表格样式名，如 `Table Grid`；也可使用三线表：`three_line`、`three line`、`Three Line Table`
- **`cell_width`**：大于 0 时，列宽为 **英寸**（`Inches`）

### `type: "table2"` 额外字段

在 `table` 基础上增加：

- **`column_alignments`**：字符串数组，按列指定 `left` / `center` / `right` / `justify`
- **`bold_content`**：表体是否加粗
- 表内文字颜色使用当前条的 **`color`** 字段

### `type: "image"` 额外字段

- **`width`** / **`height`**：厘米（`Cm`），可只填其一
- **`comment`**：同上

### `type: "float_image"` 额外字段

- **`width`** / **`height`**：厘米
- **`pos_x`** / **`pos_y`**：相对页面的偏移（实现中与 EMU 换算有关，默认 `0`）
- **`wrap_type`**：`topBottom`（上下型环绕，默认）或非该值时使用另一种环绕（浮于文字上方）

## 示例

### 示例 1：标题 + 正文 + 彩色修订说明

在 Dify 里 **`output_data` 填一个字符串**（注意转义引号，或在工作流里用变量拼接）。下面为 **未转义** 的 JSON 内容示例：

```json
[
  {
    "font_size": 16,
    "bold": true,
    "data": "文档标题",
    "alignment": "center",
    "font_name": "宋体"
  },
  {
    "data": "这是第一段正文，支持普通样式。",
    "line_spacing": 1.5
  },
  {
    "data": "这是需要突出的修订建议。",
    "color": "#F56C6C",
    "line_spacing": 1.5
  }
]
```

### 示例 2：Markdown 表格 + 分页

```json
[
  {
    "type": "table",
    "style": "Table Grid",
    "alignment": "center",
    "data": "| 姓名 | 部门 |\n| --- | --- |\n| 张三 | 研发 |\n| 李四 | 产品 |"
  },
  {
    "break_type": "page",
    "data": ""
  },
  {
    "data": "第二页从这里开始。"
  }
]
```

### 示例 3：`table2` 按列对齐 + 网络图片

```json
[
  {
    "type": "table2",
    "style": "Table Grid",
    "alignment": "center",
    "column_alignments": ["left", "right"],
    "bold_content": false,
    "color": "#333333",
    "data": "| 项目 | 金额 |\n| --- | --- |\n| 服务器 | 12000 |"
  },
  {
    "type": "image",
    "alignment": "center",
    "width": 8,
    "data": "https://www.example.com/logo.png",
    "comment": {
      "text": "示意图，以实际下载为准。",
      "author": "审核人"
    }
  }
]
```

### 示例 4：Base64 图片（Data URL 片段示意）

图片的 `data` 需为包含 `base64,` 的字符串，例如：

`data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==`

（实际使用请替换为完整 Base64。）

## 常见问题与提示

1. **`output_data` 必须是字符串**：在部分节点里若直接传「对象」而非序列化后的 JSON 字符串，可能导致类型或解析不符合预期；推荐在代码节点或模板中 `JSON.stringify(...)` 后再传入。
2. **`data` 为 JSON 字符串里的字符串**：表格的 Markdown 换行在 JSON 中写成 `\n`。
3. **图片**：仅支持 `http`/`https` 下载或带 `base64,` 的 Data URL；其他形式会报错并在文档中用红字提示失败原因（实现行为以当前版本为准）。
4. **空 `filename`**：会自动生成带时间戳的文件名，无需强行填写。

更多复杂样例可参考仓库内测试数据，例如 `test/生成word/1.json` 中的段落与 `break_type` 组合方式。
