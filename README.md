# Docxnote

**docxnote** 是一个轻量级 **DOCX 批注引擎**，仅依赖 lxml，用于自动化添加 Word 批注。

该库直接操作 **WordprocessingML**，将 DOCX 视为 **ZIP + XML** 文档，并提供一个 **基于文本视图的 API**。

与传统 DOCX 库不同，docxnote **完全隐藏 Word 的 Run 结构**，所有操作都基于 **段落字符串**。

---

## 安装

```
pip install git+https://github.com/touken928/docxnote.git
```

使用 [uv](https://github.com/astral-sh/uv)：

```
uv add git+https://github.com/touken928/docxnote.git
```

---

## 快速开始

```python
from docxnote import DocxDocument, Paragraph, Table

# 读取文档
with open("document.docx", "rb") as f:
    doc = DocxDocument.parse(f.read())

# 遍历文档块
for block in doc.blocks():
    if isinstance(block, Paragraph):
        # 为段落添加批注
        if block.text:
            block.comment("请检查表述", end=5, author="reviewer")
    
    elif isinstance(block, Table):
        # 处理表格
        rows, cols = block.shape()
        for r in range(rows):
            for c in range(cols):
                cell = block[r, c]
                # 为单元格内容添加批注
                for inner in cell.blocks():
                    if isinstance(inner, Paragraph) and inner.text:
                        inner.comment("需复核", end=3, author="reviewer")

# 生成新文档
output = doc.render()
with open("output.docx", "wb") as f:
    f.write(output)
```

---

## API

### DocxDocument

DOCX 文档对象。

#### parse

```python
DocxDocument.parse(docx_bytes)
```

解析 DOCX 并构建文档对象。

---

#### blocks

```python
doc.blocks()
```

返回文档中的块级元素：

```python
(Paragraph | Table, ...)
```

顺序与 Word 文档一致。

---

#### render

```python
doc.render()
```

生成新的 DOCX 并返回 `bytes`。

所有批注在此阶段写入文档。

---

### Paragraph

表示 Word 段落。

#### text

```python
text = paragraph.text
```

返回段落完整文本，保留换行符（`\n`）和制表符（`\t`）。

---

#### comment

```python
paragraph.comment(
    text,           # 批注内容
    start=0,        # 起始字符位置
    end=None,       # 结束字符位置（None 表示到末尾）
    *,
    author="docxnote"  # 批注作者
)
```

为段落文本范围添加批注。

**示例：**

```python
paragraph.comment("需要修改", start=3, end=8, author="张三")
```

docxnote 会自动处理：

- Run 分割
- 批注锚点
- comments.xml 写入
- 文档关系更新

---

### Table

表示 Word 表格。

#### shape

```python
rows, cols = table.shape()
```

返回表格尺寸 `(行数, 列数)`。

---

#### 单元格访问

```python
cell = table[row, col]
```

返回 `Cell` 对象。支持访问所有坐标，包括合并单元格覆盖的区域。

---

### Cell

表示表格单元格。

#### blocks

```python
cell.blocks()
```

返回单元格中的块级元素：

```python
(Paragraph | Table, ...)
```

顺序与 Word 文档一致。

---

#### bounds

```python
top, left, bottom, right = cell.bounds()
```

返回单元格边界 `(top, left, bottom, right)`，使用左闭右开区间 `[top, bottom)` 和 `[left, right)`。

对于未合并的单元格，返回 `(r, c, r+1, c+1)`。

---

## 高级用法

### 处理嵌套表格

```python
for block in doc.blocks():
    if isinstance(block, Table):
        rows, cols = block.shape()
        for r in range(rows):
            for c in range(cols):
                cell = block[r, c]
                # 遍历单元格内的块（可能包含嵌套表格）
                for inner_block in cell.blocks():
                    if isinstance(inner_block, Table):
                        # 处理嵌套表格
                        inner_rows, inner_cols = inner_block.shape()
                        # ...
```

### 多个批注

```python
# 为同一段落的不同位置添加多个批注
paragraph.comment("批注1", start=0, end=5, author="张三")
paragraph.comment("批注2", start=10, end=15, author="李四")
paragraph.comment("批注3", start=20, end=25, author="王五")
```

### 处理合并单元格

```python
table = [b for b in doc.blocks() if isinstance(b, Table)][0]

# 访问合并单元格
cell = table[0, 0]
top, left, bottom, right = cell.bounds()

# 如果单元格跨越多行或多列
if bottom - top > 1 or right - left > 1:
    print(f"合并单元格：跨越 {bottom-top} 行，{right-left} 列")
```

---

## 测试

所有测试文档使用 python-docx 动态生成，不依赖外部文件，详见 [tests/README.md](tests/README.md)。