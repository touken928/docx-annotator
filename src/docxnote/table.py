"""表格和单元格处理"""

from typing import Iterator
from lxml import etree
from .namespaces import NS


class Table:
    """表示 Word 表格"""
    
    def __init__(self, element, document):
        self._element = element
        self._document = document
        self._grid = None
        self._build_grid()
    
    def _build_grid(self):
        """构建表格网格，处理合并单元格"""
        # 只查找直接子行，不包括嵌套表格的行
        rows = self._element.findall("./w:tr", NS)
        if not rows:
            self._grid = []
            return
        
        # 构建实际的表格网格，考虑 gridSpan 和 vMerge
        self._grid = []
        
        for r_idx, row in enumerate(rows):
            # 只查找直接子单元格
            cells = row.findall("./w:tc", NS)
            
            grid_row = []
            col_idx = 0  # 实际列位置
            
            for cell in cells:
                # 检查 gridSpan（横向合并）
                colspan = 1
                tc_pr = cell.find("./w:tcPr", NS)
                if tc_pr is not None:
                    gridspan = tc_pr.find("./w:gridSpan", NS)
                    if gridspan is not None:
                        val = gridspan.get(f"{{{NS['w']}}}val")
                        if val:
                            colspan = int(val)
                
                # 创建 Cell 对象，使用实际的列位置
                grid_row.append(Cell(cell, self._document, r_idx, col_idx, colspan))
                
                # 更新列位置
                col_idx += colspan
            
            self._grid.append(grid_row)
    
    def shape(self) -> tuple[int, int]:
        """返回表格尺寸 (rows, cols)"""
        if not self._grid:
            return (0, 0)
        rows = len(self._grid)
        # 计算最大列数：每行最后一个单元格的结束位置
        cols = 0
        for row in self._grid:
            if row:
                last_cell = row[-1]
                row_cols = last_cell._col + last_cell._colspan
                cols = max(cols, row_cols)
        return (rows, cols)
    
    def __getitem__(self, key: tuple[int, int]) -> "Cell":
        """返回 Cell 对象"""
        row, col = key
        if 0 <= row < len(self._grid):
            # 在该行中查找包含指定列的单元格
            for cell in self._grid[row]:
                if cell._col <= col < cell._col + cell._colspan:
                    return cell
        # 返回空单元格
        return Cell(None, self._document, row, col, 1)



class Cell:
    """表示表格单元格"""
    
    def __init__(self, element, document, row: int, col: int, colspan: int = 1):
        self._element = element
        self._document = document
        self._row = row
        self._col = col
        self._colspan = colspan
    
    def blocks(self) -> Iterator:
        """返回单元格中的块级元素"""
        if self._element is None:
            return
        
        from .paragraph import Paragraph
        
        for child in self._element:
            tag = etree.QName(child.tag).localname
            if tag == "p":
                yield Paragraph(child, self._document)
            elif tag == "tbl":
                yield Table(child, self._document)
    
    def bounds(self) -> tuple[int, int, int, int]:
        """返回单元格边界 (top, left, bottom, right)"""
        if self._element is None:
            return (self._row, self._col, self._row + 1, self._col + 1)
        
        # 检查垂直合并
        rowspan = 1
        tc_pr = self._element.find("./w:tcPr", NS)
        if tc_pr is not None:
            vmerge = tc_pr.find("./w:vMerge", NS)
            if vmerge is not None:
                # vMerge 存在表示参与垂直合并
                # 但我们无法从单个单元格确定 rowspan
                # 简化处理：只返回当前行
                pass
        
        # 返回边界：使用实际的列位置和 colspan
        return (self._row, self._col, self._row + 1, self._col + self._colspan)
