# wc_word_report_tool

一个可独立上传到 PyPI 的小包，包含两部分：

1. `number_to_chinese_upper()`：把数字转成中文金额大写。
2. `WordFormatter`：基于 `python-docx` 的常用 Word 格式化方法。

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

## Word 方法

```python
from docx import Document
from wc_word_report_tool import WordFormatter

doc = Document()
WordFormatter.set_all_layout(doc)
WordFormatter.fengmian_doc1(doc, "测试项目")
WordFormatter.Heading_1(doc, "1. 概述")
WordFormatter.Normal_doc(doc, "这是正文。")
doc.save("demo.docx")
```

## 页码从指定位置重新按 1 开始

Word 里“页码重新从 1 开始”本质上要靠“分节”实现。

```python
from docx import Document
from wc_word_report_tool import WordFormatter

doc = Document()
WordFormatter.Normal_doc(doc, "前置内容，不参与新的页码计数")

WordFormatter.insert_section_with_page_numbering(
    doc,
    start_page_number=1,
    prefix="第",
    suffix="页",
    alignment="居中",
)

WordFormatter.Heading_1(doc, "1. 正文开始")
WordFormatter.Normal_doc(doc, "这里所在的节，页码会从 1 开始。")
doc.save("page_number_demo.docx")
```

如果你已经自己建好了分节，也可以直接对某个 `section` 调：

```python
section = doc.sections[1]
WordFormatter.restart_page_numbering(section, start=1)
```

## 生成目录并定义目录级别样式

```python
from docx import Document
from wc_word_report_tool import WordFormatter

doc = Document()
WordFormatter.add_toc(
    doc,
    title="目录",
    levels=(1, 3),
    toc_level_styles={
        1: {"font_name": "黑体", "font_size": 14, "bold": True, "space_after": 6},
        2: {"font_name": "仿宋_GB2312", "font_size": 12, "left_indent": 24},
        3: {"font_name": "宋体", "font_size": 11, "left_indent": 48},
    },
)

WordFormatter.Heading_1(doc, "1. 一级标题")
WordFormatter.Heading_2(doc, "1.1 二级标题")
WordFormatter.Heading_3(doc, "1.1.1 三级标题")
doc.save("toc_demo.docx")
```

也可以单独设置某一级目录样式：

```python
WordFormatter.set_toc_level_style(
    doc,
    1,
    font_name="黑体",
    font_size=14,
    bold=True,
)
```

如果你不是用内置 `Heading 1/2/3`，也可以按自定义段落样式进目录：

```python
WordFormatter.set_paragraph_style(
    doc,
    "MyHeading1",
    font_name="黑体",
    font_size=14,
    bold=True,
)

WordFormatter.add_toc(
    doc,
    custom_style_levels={"MyHeading1": 1},
    use_outline_levels=True,
)

WordFormatter.add_custom_heading(
    doc,
    "这是自定义样式标题",
    style_name="MyHeading1",
    level=1,
)
```

## 注意

1. `python-docx` 可以把目录域和页码域写进 `.docx`，但目录内容通常需要在 Word 里打开文档后更新一次。
2. “从第几页开始重新编号”在技术上是“从某个分节开始重新编号”。如果你说的是按最终排版后的物理页自动识别第 N 页，这件事 `python-docx` 本身不擅长，建议在目标位置手动插入分节，或者在生成逻辑里明确分节点。
3. 当前我已经补了最小测试文件在 `tests/test_word_formatter.py`，如果你本机装了 `pytest`，可以直接跑：`cd wc_word_report_tool && PYTHONPATH=src python3 -m pytest tests/test_word_formatter.py -q`

## 打包上传

目录里已经带了 `publish_to_pypi.sh`，典型流程：

```bash
cd wc_word_report_tool
python3 -m pip install --upgrade build twine
./publish_to_pypi.sh
```
