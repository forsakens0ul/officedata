"""Word 文档加载器模块 - 提取 DOCX 文件中的文本内容"""

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

try:
    from docx import Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

from .base_loader import BaseDocumentLoader
from .config import get_config


class DOCXLoader(BaseDocumentLoader):
    """Word 文档加载器

    将 Word 文档内容转换为结构化的 DataFrame，
    每行代表一个段落或表格。
    """

    SUPPORTED_EXTENSIONS = ['.docx', '.doc']

    def __init__(self):
        super().__init__()
        self._total_paragraphs: int = 0
        self._total_tables: int = 0

    def load(self, file_path: str, **kwargs) -> Dict[str, Any]:
        """加载 Word 文档

        Args:
            file_path: Word 文件路径
            **kwargs: 其他参数（保留用于扩展）

        Returns:
            文档结构信息

        Raises:
            ImportError: 如果未安装 python-docx
            FileNotFoundError: 如果文件不存在
            ValueError: 如果文件格式不支持
        """
        if not DOCX_AVAILABLE:
            raise ImportError(
                "未安装 python-docx 库，请运行: uv add python-docx"
            )

        # 验证文件
        path = self._validate_file(file_path)

        # 加载 Word 文档
        try:
            doc = Document(str(path))
        except Exception as e:
            raise ValueError(f"无法加载 Word 文档: {e}")

        self._file_path = file_path

        # 提取所有内容
        records = []
        element_number = 0

        # 遍历文档的所有元素（段落和表格）
        for element in doc.element.body:
            if isinstance(element, CT_P):
                # 段落元素
                paragraph = Paragraph(element, doc)
                if paragraph.text.strip():
                    element_number += 1
                    record = self._extract_paragraph(paragraph, element_number)
                    records.append(record)
                    self._total_paragraphs += 1

            elif isinstance(element, CT_Tbl):
                # 表格元素
                table = Table(element, doc)
                element_number += 1
                record = self._extract_table_element(table, element_number)
                records.append(record)
                self._total_tables += 1

        # 转换为 DataFrame
        if records:
            self._df = pd.DataFrame(records)
        else:
            # 空文档
            self._df = pd.DataFrame(columns=[
                'paragraph_number', 'content_type', 'content_text', 'style_name',
                'hierarchy_level', 'table_data', 'is_bold', 'is_italic', 'hyperlink'
            ])

        # 存储元数据
        self._metadata = {
            'total_paragraphs': self._total_paragraphs,
            'total_tables': self._total_tables,
            'file_type': 'docx',
        }

        return self.get_structure()

    def _extract_paragraph(self, paragraph, element_number: int) -> Dict[str, Any]:
        """提取段落信息

        Args:
            paragraph: python-docx Paragraph 对象
            element_number: 元素编号

        Returns:
            段落信息字典
        """
        # 提取文本
        text = paragraph.text.strip()

        # 获取样式
        style_name = paragraph.style.name if paragraph.style else "Normal"

        # 获取标题层级
        hierarchy_level = self._get_heading_level(paragraph)

        # 推断内容类型
        content_type = self._infer_content_type(paragraph, hierarchy_level)

        # 检查格式
        is_bold = any(run.bold for run in paragraph.runs if run.bold)
        is_italic = any(run.italic for run in paragraph.runs if run.italic)

        # 提取超链接
        hyperlink = self._extract_hyperlink(paragraph)

        return {
            'paragraph_number': element_number,
            'content_type': content_type,
            'content_text': text,
            'style_name': style_name,
            'hierarchy_level': hierarchy_level,
            'table_data': None,
            'is_bold': is_bold,
            'is_italic': is_italic,
            'hyperlink': hyperlink,
        }

    def _extract_table_element(self, table, element_number: int) -> Dict[str, Any]:
        """提取表格信息

        Args:
            table: python-docx Table 对象
            element_number: 元素编号

        Returns:
            表格信息字典
        """
        # 提取表格数据
        table_data = self._extract_table_data(table)

        # 生成表格文本描述
        table_text = f"[表格: {len(table.rows)}行 × {len(table.columns)}列]"

        # 如果表格有内容，也提取纯文本
        all_text = []
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text:
                    all_text.append(cell_text)

        if all_text:
            table_text += " " + " | ".join(all_text[:10])  # 只取前10个单元格文本

        return {
            'paragraph_number': element_number,
            'content_type': 'table',
            'content_text': table_text,
            'style_name': 'Table',
            'hierarchy_level': 0,
            'table_data': json.dumps(table_data, ensure_ascii=False),
            'is_bold': False,
            'is_italic': False,
            'hyperlink': None,
        }

    def _extract_table_data(self, table) -> Dict[str, Any]:
        """提取表格数据为 JSON 结构

        Args:
            table: python-docx Table 对象

        Returns:
            包含表格数据的字典
        """
        rows_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                row_data.append(cell.text.strip())
            rows_data.append(row_data)

        return {
            'rows': rows_data,
            'num_rows': len(table.rows),
            'num_cols': len(table.columns),
        }

    def _get_heading_level(self, paragraph) -> int:
        """获取标题层级

        Args:
            paragraph: python-docx Paragraph 对象

        Returns:
            标题层级（1-9），正文为 0
        """
        if paragraph.style and paragraph.style.name:
            style_name = paragraph.style.name
            # 检查是否为标题样式
            if 'Heading' in style_name or '标题' in style_name:
                # 尝试提取数字
                import re
                match = re.search(r'(\d+)', style_name)
                if match:
                    level = int(match.group(1))
                    if 1 <= level <= 9:
                        return level
        return 0

    def _infer_content_type(self, paragraph, hierarchy_level: int) -> str:
        """推断内容类型

        Args:
            paragraph: python-docx Paragraph 对象
            hierarchy_level: 标题层级

        Returns:
            内容类型字符串
        """
        if hierarchy_level > 0:
            return f"heading_{hierarchy_level}"

        # 检查是否为列表项
        if paragraph.style and 'List' in paragraph.style.name:
            return 'list_item'

        # 默认为段落
        return 'paragraph'

    def _extract_hyperlink(self, paragraph) -> Optional[str]:
        """提取段落中的超链接

        Args:
            paragraph: python-docx Paragraph 对象

        Returns:
            超链接 URL，如果没有返回 None
        """
        try:
            # 查找段落中的超链接
            for run in paragraph.runs:
                # 检查 run 的 XML 元素
                if run.element.rPr is not None:
                    # 遍历子元素查找超链接
                    for child in run.element.rPr:
                        if 'hyperlink' in child.tag.lower():
                            # 尝试获取链接地址
                            # 注意：这是简化的实现，可能不适用于所有情况
                            pass

            # 更简单的方法：查找段落 XML 中的超链接
            for element in paragraph._element.iter():
                if 'hyperlink' in element.tag.lower():
                    # 尝试获取 r:id 属性
                    rId = element.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if rId:
                        # 通过 rId 获取实际 URL
                        rels = paragraph.part.rels
                        if rId in rels:
                            return rels[rId].target_ref

        except Exception:
            # 如果提取失败，忽略
            pass

        return None

    def get_structure(self) -> Dict[str, Any]:
        """获取 Word 文档结构信息"""
        if self._df is None:
            raise ValueError("未加载 Word 文档")

        # 列信息
        columns_info = []
        for col in self._df.columns:
            col_data = self._df[col]
            dtype = str(col_data.dtype)
            non_null = col_data.count()
            null_count = col_data.isna().sum()

            columns_info.append({
                "name": str(col),
                "dtype": dtype,
                "non_null_count": int(non_null),
                "null_count": int(null_count),
            })

        return {
            "file_path": self._file_path,
            "file_type": "docx",
            "total_paragraphs": self._total_paragraphs,
            "total_tables": self._total_tables,
            "total_rows": len(self._df),
            "total_columns": len(self._df.columns),
            "columns": columns_info,
        }

    def get_preview(self, n_rows: Optional[int] = None) -> Dict[str, Any]:
        """获取数据预览"""
        if self._df is None:
            raise ValueError("未加载 Word 文档")

        config = get_config()
        if n_rows is None:
            n_rows = config.document.max_preview_rows

        preview_df = self._df.head(n_rows)

        return {
            "columns": list(self._df.columns),
            "data": preview_df.to_dict(orient="records"),
            "preview_rows": len(preview_df),
            "total_rows": len(self._df),
        }

    def get_summary(self) -> str:
        """获取 Word 文档摘要信息（用于 Agent 上下文）"""
        if self._df is None:
            return "未加载 Word 文档"

        structure = self.get_structure()
        preview = self.get_preview()

        lines = [
            f"📄 **已加载 Word 文档**: {structure['file_path']}",
            f"📏 **段落总数**: {structure['total_paragraphs']} 个",
            f"📊 **表格总数**: {structure['total_tables']} 个",
            f"📐 **元素总数**: {structure['total_rows']} 个",
            "",
            "**数据结构说明**:",
            "  - `paragraph_number`: 元素编号（从 1 开始）",
            "  - `content_type`: 内容类型（heading_1, heading_2, paragraph, table, list_item）",
            "  - `content_text`: 文本内容",
            "  - `hierarchy_level`: 标题层级（1-9，0 为正文）",
            "  - `style_name`: Word 样式名称",
            "  - `is_bold`: 是否包含粗体文本",
            "  - `is_italic`: 是否包含斜体文本",
            "",
            f"**前 {preview['preview_rows']} 个元素预览**:",
        ]

        # 简化的预览表格
        if preview['data']:
            # 只显示关键列
            display_columns = ['paragraph_number', 'content_type', 'hierarchy_level', 'content_text']
            lines.append("| " + " | ".join(display_columns) + " |")
            lines.append("| " + " | ".join("---" for _ in display_columns) + " |")

            for row in preview['data']:
                values = []
                for col in display_columns:
                    value = str(row.get(col, ""))
                    # 截断长文本
                    if col == 'content_text' and len(value) > 40:
                        value = value[:37] + "..."
                    values.append(value)
                lines.append("| " + " | ".join(values) + " |")

        return "\n".join(lines)
