# 测试

## 安装依赖

```bash
uv sync --extra test
```

## 运行测试

```bash
# 运行所有测试
uv run pytest

# 详细输出
uv run pytest -v

# 运行特定文件
uv run pytest tests/test_xml_validity.py
```

## 测试覆盖

- **test_xml_validity.py** - XML 语法合法性
- **test_structure_comparison.py** - 文档结构与 python-docx 对比
- **test_paragraph_text.py** - 段落文本理解
- **test_table_shape.py** - 表格形状（含合并单元格）
- **test_nested_tables.py** - 嵌套表格
- **test_cell_content.py** - 单元格内容
- **test_comments.py** - 批注功能
- **test_comment_conflicts.py** - 批注冲突和 Run 分割

所有测试文档使用 python-docx 动态生成，不依赖外部文件。
