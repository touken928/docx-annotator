"""DOCX 文档解析和渲染"""

import io
import zipfile
from typing import Iterator
from lxml import etree

from .paragraph import Paragraph
from .table import Table
from .namespaces import NS


class DocxDocument:
    """DOCX 文档对象"""
    
    def __init__(self, zip_data: bytes):
        self._zip_data = zip_data
        self._zip = zipfile.ZipFile(io.BytesIO(zip_data))
        self._document_xml = None
        self._body = None
        self._comments = []
        self._comment_id_counter = 0
        
    @classmethod
    def parse(cls, docx_bytes: bytes) -> "DocxDocument":
        """解析 DOCX 并构建文档对象"""
        doc = cls(docx_bytes)
        doc._load_document()
        return doc
    
    def _load_document(self):
        """加载 document.xml 和已有批注"""
        doc_xml = self._zip.read("word/document.xml")
        self._document_xml = etree.fromstring(doc_xml)
        self._body = self._document_xml.find(".//w:body", NS)
        
        # 加载已有的批注
        self._load_existing_comments()
    
    def _load_existing_comments(self):
        """加载已有的批注"""
        try:
            comments_xml = self._zip.read("word/comments.xml")
            comments_tree = etree.fromstring(comments_xml)
            
            max_id = -1
            for comment in comments_tree:
                comment_id_str = comment.get(f"{{{NS['w']}}}id")
                if comment_id_str:
                    comment_id = int(comment_id_str)
                    max_id = max(max_id, comment_id)
                    
                    # 提取批注内容
                    author = comment.get(f"{{{NS['w']}}}author", "")
                    text_elem = comment.find(f".//{{{NS['w']}}}t", NS)
                    text = text_elem.text if text_elem is not None and text_elem.text else ""
                    
                    self._comments.append((comment_id, text, author))
            
            # 设置下一个批注 ID
            self._comment_id_counter = max_id + 1
        except KeyError:
            # 没有 comments.xml 文件
            pass
    
    def blocks(self) -> Iterator[Paragraph | Table]:
        """返回文档中的块级元素"""
        if self._body is None:
            return
        
        for child in self._body:
            tag = etree.QName(child.tag).localname
            if tag == "p":
                yield Paragraph(child, self)
            elif tag == "tbl":
                yield Table(child, self)
    
    def add_comment(self, text: str, author: str = "docxnote") -> int:
        """添加批注并返回 ID"""
        comment_id = self._comment_id_counter
        self._comment_id_counter += 1
        self._comments.append((comment_id, text, author))
        return comment_id
    
    def render(self) -> bytes:
        """生成新的 DOCX 并返回 bytes"""
        output = io.BytesIO()
        
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as out_zip:
            # 准备 rels 和 content types（如果有批注）
            rels_data = None
            content_types_data = None
            if self._comments:
                rels_data = self._prepare_rels()
                content_types_data = self._prepare_content_types()
            
            # 复制所有原始文件
            for item in self._zip.namelist():
                if item == "word/document.xml":
                    continue
                if item == "word/comments.xml":
                    continue
                if item == "word/_rels/document.xml.rels" and rels_data is not None:
                    continue
                if item == "[Content_Types].xml" and content_types_data is not None:
                    continue
                out_zip.writestr(item, self._zip.read(item))
            
            # 写入修改后的 document.xml
            doc_bytes = etree.tostring(
                self._document_xml,
                xml_declaration=True,
                encoding="UTF-8",
                standalone=True
            )
            out_zip.writestr("word/document.xml", doc_bytes)
            
            # 写入 comments.xml、rels 和 content types
            if self._comments:
                comments_xml = self._build_comments_xml()
                out_zip.writestr("word/comments.xml", comments_xml)
                out_zip.writestr("word/_rels/document.xml.rels", rels_data)
                out_zip.writestr("[Content_Types].xml", content_types_data)
        
        return output.getvalue()
    
    def _build_comments_xml(self) -> bytes:
        """构建 comments.xml"""
        root = etree.Element(
            f"{{{NS['w']}}}comments",
            nsmap=NS
        )
        
        for comment_id, text, author in self._comments:
            comment = etree.SubElement(
                root,
                f"{{{NS['w']}}}comment",
                attrib={
                    f"{{{NS['w']}}}id": str(comment_id),
                    f"{{{NS['w']}}}author": author,
                    f"{{{NS['w']}}}date": "2024-01-01T00:00:00Z",
                    f"{{{NS['w']}}}initials": author[0].upper() if author else "D"
                }
            )
            
            p = etree.SubElement(comment, f"{{{NS['w']}}}p")
            r = etree.SubElement(p, f"{{{NS['w']}}}r")
            t = etree.SubElement(r, f"{{{NS['w']}}}t")
            t.text = text
        
        return etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True
        )
    
    def _prepare_rels(self) -> bytes:
        """准备 document.xml.rels 数据以包含 comments.xml 关系"""
        rels_path = "word/_rels/document.xml.rels"
        
        try:
            rels_data = self._zip.read(rels_path)
            rels_xml = etree.fromstring(rels_data)
        except KeyError:
            # 创建新的 rels
            rels_xml = etree.Element(
                "Relationships",
                nsmap={"": "http://schemas.openxmlformats.org/package/2006/relationships"}
            )
        
        # 检查是否已有 comments 关系
        has_comments = False
        for rel in rels_xml:
            if rel.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":
                has_comments = True
                break
        
        if not has_comments:
            # 添加 comments 关系
            max_id = 0
            for rel in rels_xml:
                rel_id = rel.get("Id", "")
                if rel_id.startswith("rId"):
                    try:
                        num = int(rel_id[3:])
                        max_id = max(max_id, num)
                    except ValueError:
                        pass
            
            etree.SubElement(
                rels_xml,
                "Relationship",
                attrib={
                    "Id": f"rId{max_id + 1}",
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                    "Target": "comments.xml"
                }
            )
        
        return etree.tostring(
            rels_xml,
            xml_declaration=True,
            encoding="UTF-8"
        )
    
    def _prepare_content_types(self) -> bytes:
        """准备 [Content_Types].xml 数据以包含 comments.xml"""
        ct_data = self._zip.read("[Content_Types].xml")
        ct_xml = etree.fromstring(ct_data)
        
        # 获取命名空间
        ns = ct_xml.nsmap.get(None, "http://schemas.openxmlformats.org/package/2006/content-types")
        
        # 检查是否已有 comments.xml 的 Override
        has_comments_override = False
        for override in ct_xml:
            if override.get("PartName") == "/word/comments.xml":
                has_comments_override = True
                break
        
        if not has_comments_override:
            # 添加 comments.xml 的 Override
            override_elem = etree.Element(
                f"{{{ns}}}Override",
                attrib={
                    "PartName": "/word/comments.xml",
                    "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
                }
            )
            ct_xml.append(override_elem)
        
        return etree.tostring(
            ct_xml,
            xml_declaration=True,
            encoding="UTF-8"
        )
