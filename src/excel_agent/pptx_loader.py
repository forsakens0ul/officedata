"""PowerPoint 加载器模块 - 提取 PPTX 文件中的文本内容"""

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

from .base_loader import BaseDocumentLoader
from .config import get_config


class PPTXLoader(BaseDocumentLoader):
    """PowerPoint 文件加载器

    将 PowerPoint 幻灯片内容转换为结构化的 DataFrame，
    每行代表一个内容元素（标题、正文、列表项、表格或备注）。
    """

    SUPPORTED_EXTENSIONS = ['.pptx', '.ppt']

    def __init__(self):
        super().__init__()
        self._total_slides: int = 0

    def load(self, file_path: str, **kwargs) -> Dict[str, Any]:
        """加载 PowerPoint 文件

        Args:
            file_path: PowerPoint 文件路径
            **kwargs: 其他参数（保留用于扩展）

        Returns:
            文档结构信息

        Raises:
            ImportError: 如果未安装 python-pptx
            FileNotFoundError: 如果文件不存在
            ValueError: 如果文件格式不支持
        """
        if not PPTX_AVAILABLE:
            raise ImportError(
                "未安装 python-pptx 库，请运行: uv add python-pptx"
            )

        # 验证文件
        path = self._validate_file(file_path)

        # 加载 PowerPoint
        try:
            prs = Presentation(str(path))
        except Exception as e:
            raise ValueError(f"无法加载 PowerPoint 文件: {e}")

        self._file_path = file_path
        self._total_slides = len(prs.slides)

        # 提取所有文本内容
        records = []

        for slide_idx, slide in enumerate(prs.slides, start=1):
            # 提取幻灯片标题
            slide_title = self._extract_slide_title(slide)

            # 提取幻灯片内容
            content_order = 0
            for shape in slide.shapes:
                # 提取文本框内容
                if shape.has_text_frame:
                    # 跳过标题（已经单独提取）
                    if shape == slide.shapes.title:
                        continue

                    for paragraph in shape.text_frame.paragraphs:
                        text = paragraph.text.strip()
                        if text:
                            records.append({
                                'slide_number': slide_idx,
                                'slide_title': slide_title,
                                'content_type': self._infer_content_type(shape, slide),
                                'content_order': content_order,
                                'content_text': text,
                                'content_hierarchy': f"level_{paragraph.level}",
                                'table_data': None,
                                'speaker_notes': "",
                            })
                            content_order += 1

                # 提取表格
                if shape.has_table:
                    table_data = self._extract_table(shape.table)
                    table_text = self._table_to_text(shape.table)

                    records.append({
                        'slide_number': slide_idx,
                        'slide_title': slide_title,
                        'content_type': 'table',
                        'content_order': content_order,
                        'content_text': table_text,
                        'content_hierarchy': 'level_0',
                        'table_data': json.dumps(table_data, ensure_ascii=False),
                        'speaker_notes': "",
                    })
                    content_order += 1

            # 提取演讲者备注
            notes_text = self._extract_speaker_notes(slide)
            if notes_text and records:
                # 将备注附加到最后一个内容元素
                # 如果当前幻灯片有内容，附加到最后一行
                for record in reversed(records):
                    if record['slide_number'] == slide_idx:
                        record['speaker_notes'] = notes_text
                        break
            elif notes_text and not records:
                # 如果幻灯片只有备注没有其他内容，创建一个备注行
                records.append({
                    'slide_number': slide_idx,
                    'slide_title': slide_title,
                    'content_type': 'notes',
                    'content_order': 0,
                    'content_text': notes_text,
                    'content_hierarchy': 'level_0',
                    'table_data': None,
                    'speaker_notes': notes_text,
                })

        # 转换为 DataFrame
        if records:
            self._df = pd.DataFrame(records)
        else:
            # 空演示文稿
            self._df = pd.DataFrame(columns=[
                'slide_number', 'slide_title', 'content_type', 'content_order',
                'content_text', 'content_hierarchy', 'table_data', 'speaker_notes'
            ])

        # 存储元数据
        self._metadata = {
            'total_slides': self._total_slides,
            'file_type': 'pptx',
        }

        return self.get_structure()

    def _extract_slide_title(self, slide) -> str:
        """提取幻灯片标题

        Args:
            slide: python-pptx Slide 对象

        Returns:
            幻灯片标题文本，如果没有标题返回空字符串
        """
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            return slide.shapes.title.text.strip()
        return ""

    def _extract_speaker_notes(self, slide) -> str:
        """提取演讲者备注

        Args:
            slide: python-pptx Slide 对象

        Returns:
            备注文本，如果没有备注返回空字符串
        """
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes_text = slide.notes_slide.notes_text_frame.text.strip()
                return notes_text
        except Exception:
            # 某些 PPT 文件可能没有备注页
            pass
        return ""

    def _infer_content_type(self, shape, slide) -> str:
        """推断内容类型

        Args:
            shape: python-pptx Shape 对象
            slide: python-pptx Slide 对象

        Returns:
            内容类型：'title', 'body', 'notes'
        """
        # 如果是标题形状
        if shape == slide.shapes.title:
            return 'title'

        # 默认为正文
        return 'body'

    def _extract_table(self, table) -> Dict[str, Any]:
        """提取表格数据为 JSON 结构

        Args:
            table: python-pptx Table 对象

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

    def _table_to_text(self, table) -> str:
        """将表格转换为文本描述

        Args:
            table: python-pptx Table 对象

        Returns:
            表格的文本描述
        """
        return f"[表格: {len(table.rows)}行 × {len(table.columns)}列]"

    def get_structure(self) -> Dict[str, Any]:
        """获取 PowerPoint 结构信息"""
        if self._df is None:
            raise ValueError("未加载 PowerPoint 文件")

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
            "file_type": "pptx",
            "total_slides": self._total_slides,
            "total_rows": len(self._df),
            "total_columns": len(self._df.columns),
            "columns": columns_info,
        }

    def get_preview(self, n_rows: Optional[int] = None) -> Dict[str, Any]:
        """获取数据预览"""
        if self._df is None:
            raise ValueError("未加载 PowerPoint 文件")

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
        """获取 PowerPoint 摘要信息（用于 Agent 上下文）"""
        if self._df is None:
            return "未加载 PowerPoint 文件"

        structure = self.get_structure()
        preview = self.get_preview()

        lines = [
            f"📊 **已加载 PowerPoint 文件**: {structure['file_path']}",
            f"📑 **幻灯片总数**: {structure['total_slides']} 张",
            f"📏 **内容元素总数**: {structure['total_rows']} 个",
            "",
            "**数据结构说明**:",
            "  - `slide_number`: 幻灯片编号（从 1 开始）",
            "  - `slide_title`: 幻灯片标题",
            "  - `content_type`: 内容类型（title/body/table/notes）",
            "  - `content_text`: 文本内容",
            "  - `content_hierarchy`: 列表层级（level_0, level_1, ...）",
            "  - `speaker_notes`: 演讲者备注",
            "",
            f"**前 {preview['preview_rows']} 个内容元素预览**:",
        ]

        # 简化的预览表格
        if preview['data']:
            # 只显示关键列
            display_columns = ['slide_number', 'slide_title', 'content_type', 'content_text']
            lines.append("| " + " | ".join(display_columns) + " |")
            lines.append("| " + " | ".join("---" for _ in display_columns) + " |")

            for row in preview['data']:
                values = []
                for col in display_columns:
                    value = str(row.get(col, ""))
                    # 截断长文本
                    if len(value) > 30:
                        value = value[:27] + "..."
                    values.append(value)
                lines.append("| " + " | ".join(values) + " |")

        return "\n".join(lines)
