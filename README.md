# wc_word_report_tool [![PyPI Downloads](https://static.pepy.tech/personalized-badge/wc-word-report-tool?period=total&units=INTERNATIONAL_SYSTEM&left_color=BLACK&right_color=GREEN&left_text=downloads)](https://pepy.tech/projects/wc-word-report-tool)

基于 `python-docx` 的中文 Word 报告格式化工具，内置：

1. **数字转中文大写**：`number_to_chinese_upper()` 把数字转成人民币金额大写。
2. **Word 格式化**：`WordFormatter` 提供从封面、正文、表格、页眉页脚到目录、页码的完整覆盖，默认值贴近中国公文标准。

> 版本：v0.4.20
> 作者：willcha
> 许可证：MIT
> Python：>=3.9

## 安装

```bash
pip install wc_word_report_tool
```

## 数字转大写

```python
from wc_word_report_tool import number_to_chinese_upper

print(number_to_chinese_upper(123456.78))
# 壹拾贰万叁仟肆佰伍拾陆元柒角捌分
```

命令行也可以直接用：

```bash
wc-rmb-upper 100200.03
```

## 当前接口约定

1. **覆盖完整**：补齐表格、页眉页脚、文档默认样式等 v0.1.x 缺失能力。
2. **命名一致**：公开方法统一使用 snake_case，例如 `heading`、`body`、`set_header`。
3. **默认合理**：正文默认宋体、四号（14pt）、首行缩进 2 字符、两端对齐和 1.5 倍行距。
4. **可选严格校验**：对齐方式默认兼容回退到左对齐；传 `strict=True` 时，未知值抛出 `ValueError`。

格式化器在创建时绑定文档。实例方法通过 `formatter` 调用，不再重复传入
`doc`；只处理独立对象的方法（例如 `set_header(section, ...)`）保留为静态方法：

```python
formatter = WordFormatter(doc)
formatter.cover_text("项目名称")
formatter.heading("一、项目概况")
formatter.body("正文内容")
formatter.add_toc()
```

## 快速上手

### 完整报告骨架

```python
from docx import Document
from wc_word_report_tool import WordFormatter

doc = Document()
formatter = WordFormatter(doc)

# 1. 文档级设置
formatter.setup_defaults()
formatter.set_document_language()
# setup_defaults 只设置页面参数；正文格式通过 body() 单独设置

# 2. 封面
formatter.insert_img("logo.png", width=5)
formatter.blank_lines(2)
formatter.cover_text("测试项目", font_name="宋体", font_size=22, bold=True)
formatter.blank_lines(8)
formatter.right_text("委托单位：XXX公司")
formatter.right_text("编制单位：YYY公司")

# 3. 正文（新节）
formatter.insert_section()
formatter.heading("1. 概述")
formatter.body("本项目位于……，建设内容包括……")
formatter.body("项目背景说明……")

# 4. 表格
formatter.add_table(
    headers=["序号", "项目", "金额（万元）"],
    rows=[
        ["1", "建筑工程", "1200.50"],
        ["2", "安装工程", "850.00"],
        ["3", "合计", "2050.50"],
    ],
    col_widths=[2, 6, 4],
    font_size=12,
)

# 5. 页眉页脚 + 页码从正文开始
WordFormatter.set_header(doc.sections[1], "测试项目 竣工决算报告", alignment="居中")
formatter.set_page_number_from_section(start_section_idx=1, start=1,
                                       prefix="第 ", suffix=" 页")

# 6. 目录
formatter.add_toc(title="目  录", levels=(1, 3))

doc.save("demo.docx")
```

## API 总览

### 文档级设置

| 方法 | 用途 |
|------|------|
| `set_default_font(*, font_size, cn_font, en_font)` | 设置 Normal 样式默认字体 |
| `setup_defaults(*, top, bottom, left, right, gutter)` | 一键设置默认页边距；不设置正文格式 |
| `set_document_language(lang)` | 设置 DOCX 文档语言 |
| `set_page_margins(*, top, bottom, left, right, gutter, horizontal_alignment)` | 设置页边距和页面水平对齐 |

### 段落与标题

| 方法 | 用途 |
|------|------|
| `body(text, *, font_name, font_size, indent, bold, highlight, alignment, line_spacing, ...)` | 正文段落；默认字符缩进优先、pt 缩进兜底 |
| `blank_lines(count)` | 批量添加空行 |
| `heading(text, *, level, font_name, font_size, bold, indent, ...)` | 通用标题（默认一级标题格式） |
| `cover_text(text, *, font_name, font_size, bold, alignment, indent, ...)` | 封面文本（默认居中） |
| `right_text(text, *, font_name, font_size, indent)` | 右对齐段落（公司名/日期） |
| `created_time(value, *, font_name, font_size, bold)` | 数字日期文本，默认右对齐 |
| `chinese_date(value, *, font_name, font_size, bold, alignment)` | 中文完整日期，默认居中 |
| `chinese_year_month(value, *, font_name, font_size, bold, alignment)` | 中文年月，默认居中 |
| `insert_img(img_path, width, *, alignment)` | 插入图片（width 单位 cm） |

`body(indent=True)` 会同时写入字符单位和绝对长度单位：

```xml
<w:ind w:firstLine="560" w:firstLineChars="200"/>
```

其中 `w:firstLineChars="200"` 表示首行缩进 2 字符，在 Word 中优先生效；
`w:firstLine` 按当前字号计算，用于不识别字符缩进的兼容程序。

### 结算封面与签发页

```python
formatter = WordFormatter(doc)

formatter.fengmian_jiesuan(
    logo="logo.png",
    project_name="某工程",
    entrusting_unit="某委托单位",
    date="2026-07-21",  # 可省略，默认当天
)

formatter.qianfaye(
    logo="logo.png",
    project_name="某工程",
    entrusting_unit="某委托单位",
    personnel={
        "公司签发": {"姓名": "张三", "职务": "总经理", "职称": "正高级工程师"},
        "部门核准": {"姓名": "李四", "部门": "造价部", "职务": "经理", "职称": "高级工程师"},
        "部门审核": {"姓名": "王五", "部门": "造价部", "职务": "副经理", "职称": "高级工程师"},
        "小组初审": {"姓名": "赵六", "部门": "造价部", "职务": "组长", "职称": "工程师"},
        "项目负责人": {
            "姓名": "钱七", "部门": "造价部", "职务": "项目经理",
            "职称": "高级工程师", "联系电话": "13800000000",
        },
    },
    participants=[{"姓名": "孙八", "职称": "工程师"}],
)
```

`compiling_unit`、`report_title`、`logo_width` 可按项目覆盖；封面还可传
`info_blank_lines`，签发页还可传 `footer_text` 和 `logo_y_offset_pt`。
`logo_y_offset_pt` 单位为 pt，正数上移、负数下移。

### 表格

| 方法 | 用途 |
|------|------|
| `add_table(headers, rows, *, col_widths, font_size, header_bold, alignment, border_color, border_size)` | 一站式创建带表头表格 |
| `set_cell(cell, text, *, font_name, font_size, bold, alignment, line_spacing)` | 设置单元格格式（支持加粗/对齐） |
| `set_table_borders(table, *, color, size)` | 为表格添加边框 |

```python
formatter.add_table(
    headers=["序号", "项目", "金额"],
    rows=[["1", "建安费", "1200"], ["2", "设备费", "850"]],
    col_widths=[2, 5, 3],          # cm
    header_bold=True,
    alignment="居中",
)
```

### 页眉页脚

| 方法 | 用途 |
|------|------|
| `set_header(section, text, *, alignment, font_name, font_size)` | 设置节页眉（默认楷体小五号） |
| `set_footer(section, text, *, alignment, font_name, font_size)` | 设置节页脚（默认楷体小五号） |
| `clear_header(section)` / `clear_footer(section)` | 清空节页眉/页脚 |
| `set_header_image(section, image_path, *, width, height, alignment, floating, bottom_border, y_offset_pt)` | 设置页眉图片；可浮于文字上方并调整垂直位置 |
| `set_footer_image(section, image_path, *, width, height, alignment, floating, bottom_border, y_offset_pt)` | 设置页脚图片；可浮于文字上方并调整垂直位置 |

```python
WordFormatter.set_header(doc.sections[1], "项目名称 报告名称", alignment="居中")
WordFormatter.set_footer(doc.sections[1], "编制单位名称", alignment="居右")
WordFormatter.set_header_image(doc.sections[1], "logo.png", width=3)
```

### 节与页码

| 方法 | 用途 |
|------|------|
| `insert_section(start_type, *, add_page_number, restart_page_number, inherit_header, inherit_footer)` | 插入新节并控制页眉页脚继承 |
| `set_page_number_start(section, start)` | 设置节页码起始值 |
| `add_page_number(paragraph, *, prefix, suffix, alignment, font_name, font_size)` | 向段落添加 PAGE 域 |
| `add_footer_page_number(section, *, prefix, suffix, ...)` | 节页脚添加页码 |
| `restart_page_numbering(section, *, start, add_footer_number, ...)` | 重启单节页码 |
| `insert_section_with_page_numbering(*, start_page_number, ...)` | 插入节并重启页码（组合方法） |
| `set_page_number_from_section(start_section_idx, *, start, ...)` | **跨节页码**：从指定节起编号，之前节无页码 |

```python
# 文档有 3 个节：0=前置（封面/扉页/目录）, 1=正文起, 2=正文续
formatter.set_page_number_from_section(
    start_section_idx=1, start=1,
    prefix="第 ", suffix=" 页", alignment="居中",
)
# 节0 无页码，节1 从"第 1 页"开始，节2 链接前节延续编号
```

### 目录

| 方法 | 用途 |
|------|------|
| `add_toc(*, title, levels, title_style, toc_level_styles, ...)` | 插入 Word 目录域（TOC field）；标题默认使用 Normal 样式 |
| `set_toc_level_style(level, *, font_name, ...)` | 设置某一级 TOC 样式 |
| `set_paragraph_style(style_name, *, base_style_name, ...)` | 创建/更新自定义段落样式 |
| `add_custom_heading(text_content, *, style_name, level, ...)` | 使用自定义样式添加标题 |

```python
formatter.add_toc(
    title="目  录",
    levels=(1, 3),
    toc_level_styles={
        1: {"font_name": "黑体", "font_size": 14, "bold": True, "space_after": 6},
        2: {"font_name": "仿宋_GB2312", "font_size": 12, "left_indent": 24},
    },
)
formatter.heading("1. 一级标题")
```

### 底层工具方法

| 方法 | 用途 |
|------|------|
| `set_run_font(run, *, cn_font, en_font, size, bold, color, highlight)` | 设置 run 字体 |
| `set_paragraph_format(paragraph, *, alignment, line_spacing, first_line_indent_chars, first_line_indent_pt, ...)` | 设置段落格式；字符缩进优先，pt 缩进兜底 |
| `resolve_alignment(alignment, *, strict)` | 对齐方式解析 |

## 字号参数

`font_size` 参数支持两种形式：

- **数字（pt）**：`font_size=14` 表示 14pt
- **中文字号字符串**：`font_size="三号"` = 16pt，`font_size="小五"` = 9pt

支持的中文字号：`初号/小初/一号/小一/二号/小二/三号/小三/四号/小四/五号/小五/六号/小六/七号/八号`。

## 对齐方式参数

`alignment` 参数支持多种形式：

| 类型 | 示例 |
|------|------|
| 中文 | `"居中"` / `"居左"` / `"居右"` / `"左对齐"` / `"右对齐"` / `"两端对齐"` |
| 英文 | `"center"` / `"left"` / `"right"` / `"justify"` |
| 单字母 | `"C"` / `"L"` / `"R"` |
| 数字 | `1` / `2` / `3` |
| 枚举 | `WD_PARAGRAPH_ALIGNMENT.CENTER` |

默认 `strict=False` 时未知值回退 LEFT；`strict=True` 时抛 `ValueError`。

## 单位约定

| 参数 | 单位 |
|------|------|
| `top/bottom/left/right`（页边距） | cm |
| `gutter`（装订线） | cm |
| `width`（图片宽度） | cm |
| `col_widths`（列宽） | cm |
| `font_size`（字号） | pt 或中文字号 |
| `space_before/space_after` | pt |
| `first_line_indent_pt` | pt |
| `first_line_indent_chars` | 字符数（优先写入 `w:firstLineChars`，pt 值作为兼容兜底） |
| `border_size`（边框粗细） | 1/8 pt（4 = 0.5pt 细线，8 = 1pt） |

## 版本说明

### v0.4.20

- 正文及带 `indent=True` 的相关接口优先写入 `w:firstLineChars="200"`。
- 按字号计算的 `w:firstLine` 继续保留为兼容兜底。
- `set_paragraph_format()`、TOC 样式和自定义段落样式新增
  `first_line_indent_chars` 参数。
- `right_text()` 的 pt 兜底值改为跟随实际字号计算。

### v0.4.19

- 浮动页眉图片改用 Word 兼容的 VML 结构。
- 页眉图片与标题文字可同时保留，并支持 `logo_y_offset_pt` 垂直调整。

### v0.4.3

- API 收敛为 snake_case；实例方法统一通过 `WordFormatter(doc)` 绑定的格式化器调用。
- 不再提供旧方法名兼容层。

## 注意事项

1. **目录刷新**：`python-docx` 可写入 TOC 域，但目录内容需在 Word/OnlyOffice 打开后刷新（F9 或右键 → 更新域）才会显示页码。
2. **页码分节**：从某页开始重新编号本质上是"从某个分节开始重新编号"。建议在正文起、附录起等关键节点显式插入分节。
3. **字体可用性**：默认字体 `仿宋_GB2312` / `楷体` / `黑体` 在 Windows 上预装；macOS / Linux 可能需要额外安装或回退到 `仿宋` / `STKaiti` / `SimHei`。



## License

MIT
