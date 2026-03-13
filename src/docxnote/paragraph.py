"""段落处理"""

from lxml import etree
from .namespaces import NS


class Paragraph:
    """表示 Word 段落"""
    
    def __init__(self, element, document):
        self._element = element
        self._document = document
        self._text_cache = None
    
    @property
    def text(self) -> str:
        """返回段落完整文本"""
        if self._text_cache is not None:
            return self._text_cache
        
        text_parts = []
        for run in self._element.findall(".//w:r", NS):
            # 遍历 run 的所有子元素，保持顺序
            for child in run:
                tag = etree.QName(child.tag).localname
                if tag == "t":
                    # 文本节点
                    if child.text:
                        text_parts.append(child.text)
                elif tag == "br":
                    # 换行符
                    text_parts.append("\n")
                elif tag == "tab":
                    # 制表符
                    text_parts.append("\t")
        
        self._text_cache = "".join(text_parts)
        return self._text_cache
    
    def comment(self, text: str, start: int = 0, end: int | None = None, *, author: str = "docxnote"):
        """为段落文本范围添加批注"""
        if end is None:
            end = len(self.text)
        
        # 获取批注 ID
        comment_id = self._document.add_comment(text, author)
        
        # 在段落中插入批注标记
        self._insert_comment_markers(comment_id, start, end)
    
    def _insert_comment_markers(self, comment_id: int, start: int, end: int):
        """在指定位置插入批注起止标记"""
        runs = list(self._element.findall(".//w:r", NS))
        if not runs:
            return
        
        # 计算字符位置到 run 的映射
        run_positions = []
        current_pos = 0
        
        for run in runs:
            run_start = current_pos
            run_text = ""
            for t in run.findall(".//w:t", NS):
                if t.text:
                    run_text += t.text
            run_end = current_pos + len(run_text)
            run_positions.append((run, run_start, run_end, run_text))
            current_pos = run_end
        
        # 找到需要分割的 run
        start_run_idx = None
        end_run_idx = None
        
        for idx, (run, run_start, run_end, run_text) in enumerate(run_positions):
            if start_run_idx is None and run_start <= start < run_end:
                start_run_idx = idx
            if end_run_idx is None and run_start < end <= run_end:
                end_run_idx = idx
        
        if start_run_idx is None or end_run_idx is None:
            return
        
        # 分割 run 并插入标记
        self._split_and_mark(run_positions, start_run_idx, end_run_idx, start, end, comment_id)
    
    def _split_and_mark(self, run_positions, start_idx, end_idx, start, end, comment_id):
        """分割 run 并插入批注标记"""
        # 简化实现：在第一个 run 前插入开始标记，在最后一个 run 后插入结束标记
        start_run, start_pos, _, _ = run_positions[start_idx]
        end_run, _, end_pos, _ = run_positions[end_idx]
        
        # 创建批注范围开始标记
        comment_start = etree.Element(
            f"{{{NS['w']}}}commentRangeStart",
            attrib={f"{{{NS['w']}}}id": str(comment_id)}
        )
        
        # 创建批注范围结束标记
        comment_end = etree.Element(
            f"{{{NS['w']}}}commentRangeEnd",
            attrib={f"{{{NS['w']}}}id": str(comment_id)}
        )
        
        # 创建批注引用
        comment_ref_run = etree.Element(f"{{{NS['w']}}}r")
        comment_ref = etree.SubElement(
            comment_ref_run,
            f"{{{NS['w']}}}commentReference",
            attrib={f"{{{NS['w']}}}id": str(comment_id)}
        )
        
        # 插入标记
        parent = self._element
        
        # 查找 run 在父元素中的位置
        try:
            children = list(parent)
            start_run_pos = children.index(start_run)
            end_run_pos = children.index(end_run)
        except ValueError:
            # run 不是直接子元素，跳过
            return
        
        # 在开始 run 之前插入开始标记
        parent.insert(start_run_pos, comment_start)
        
        # 在结束 run 之后插入结束标记和引用（注意索引偏移）
        parent.insert(end_run_pos + 2, comment_end)
        parent.insert(end_run_pos + 3, comment_ref_run)
