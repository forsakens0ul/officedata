"""多文档管理器模块 - 统一管理 Excel、PowerPoint、Word 等文档"""

import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Type

import pandas as pd

from .base_loader import BaseDocumentLoader
from .excel_loader import ExcelLoader


@dataclass
class DocumentInfo:
    """文档的元信息（扩展自 TableInfo）"""
    id: str
    filename: str
    file_path: str
    doc_type: str  # "excel", "pptx", "docx"
    sheet_name: Optional[str] = None  # Excel 专用
    total_rows: int = 0
    total_columns: int = 0
    loaded_at: datetime = field(default_factory=datetime.now)
    is_joined: bool = False  # 是否为连接表
    source_tables: List[str] = field(default_factory=list)  # 源表名称列表


class LoaderFactory:
    """文档加载器工厂类

    根据文件扩展名自动创建对应的加载器实例。
    """

    # 扩展名到加载器类的映射
    _loader_registry: Dict[str, Type[BaseDocumentLoader]] = {}

    @classmethod
    def register_loader(cls, extensions: List[str], loader_class: Type[BaseDocumentLoader]):
        """注册加载器类

        Args:
            extensions: 支持的扩展名列表
            loader_class: 加载器类
        """
        for ext in extensions:
            cls._loader_registry[ext.lower()] = loader_class

    @classmethod
    def create_loader(cls, file_path: str) -> BaseDocumentLoader:
        """创建加载器实例

        Args:
            file_path: 文件路径

        Returns:
            对应的加载器实例

        Raises:
            ValueError: 如果文件格式不支持
        """
        ext = Path(file_path).suffix.lower()

        loader_class = cls._loader_registry.get(ext)
        if loader_class is None:
            supported_exts = ', '.join(sorted(cls._loader_registry.keys()))
            raise ValueError(
                f"不支持的文件格式: {ext}\n"
                f"支持的格式: {supported_exts}"
            )

        return loader_class()

    @classmethod
    def get_supported_extensions(cls) -> List[str]:
        """获取所有支持的文件扩展名

        Returns:
            支持的扩展名列表
        """
        return sorted(cls._loader_registry.keys())

    @classmethod
    def get_doc_type(cls, file_path: str) -> str:
        """根据文件路径推断文档类型

        Args:
            file_path: 文件路径

        Returns:
            文档类型：'excel', 'pptx', 'docx'
        """
        ext = Path(file_path).suffix.lower()

        if ext in ['.xlsx', '.xls', '.xlsm']:
            return 'excel'
        elif ext in ['.pptx', '.ppt']:
            return 'pptx'
        elif ext in ['.docx', '.doc']:
            return 'docx'
        else:
            return 'unknown'


# 注册默认加载器
LoaderFactory.register_loader(['.xlsx', '.xls', '.xlsm'], ExcelLoader)

# 延迟导入以避免循环依赖
try:
    from .pptx_loader import PPTXLoader
    LoaderFactory.register_loader(['.pptx', '.ppt'], PPTXLoader)
except ImportError:
    pass  # python-pptx 未安装

try:
    from .docx_loader import DOCXLoader
    LoaderFactory.register_loader(['.docx', '.doc'], DOCXLoader)
except ImportError:
    pass  # python-docx 未安装


class MultiDocumentLoader:
    """多文档管理器

    管理多个异构文档（Excel、PowerPoint、Word）的加载和切换。
    兼容原有的 MultiExcelLoader 接口。
    """

    def __init__(self):
        self._documents: Dict[str, BaseDocumentLoader] = {}  # doc_id -> Loader
        self._doc_infos: Dict[str, DocumentInfo] = {}  # doc_id -> DocumentInfo
        self._active_doc_id: Optional[str] = None

    @property
    def is_loaded(self) -> bool:
        """是否有任何文档已加载"""
        return len(self._documents) > 0

    @property
    def active_table_id(self) -> Optional[str]:
        """获取当前活跃文档ID（兼容接口）"""
        return self._active_doc_id

    @property
    def active_doc_id(self) -> Optional[str]:
        """获取当前活跃文档ID"""
        return self._active_doc_id

    def add_document(
        self,
        file_path: str,
        sheet_name: Optional[str] = None,
        **kwargs
    ) -> tuple[str, Dict[str, Any]]:
        """添加一个新文档

        Args:
            file_path: 文件路径
            sheet_name: 工作表名称（仅 Excel 使用）
            **kwargs: 其他加载参数

        Returns:
            (文档ID, 结构信息)

        Raises:
            ValueError: 如果文件格式不支持
        """
        # 创建对应的加载器
        loader = LoaderFactory.create_loader(file_path)

        # 加载文档
        if isinstance(loader, ExcelLoader) and sheet_name:
            structure = loader.load(file_path, sheet_name=sheet_name)
        else:
            structure = loader.load(file_path, **kwargs)

        # 生成唯一ID
        doc_id = str(uuid.uuid4())[:8]

        # 获取文件信息
        filename = Path(file_path).name
        doc_type = LoaderFactory.get_doc_type(file_path)

        # 存储文档信息
        self._documents[doc_id] = loader
        self._doc_infos[doc_id] = DocumentInfo(
            id=doc_id,
            filename=filename,
            file_path=file_path,
            doc_type=doc_type,
            sheet_name=structure.get("sheet_name"),
            total_rows=structure.get("total_rows", 0),
            total_columns=structure.get("total_columns", 0),
        )

        # 自动设为活跃文档
        self._active_doc_id = doc_id

        return doc_id, structure

    def add_table(
        self,
        file_path: str,
        sheet_name: Optional[str] = None
    ) -> tuple[str, Dict[str, Any]]:
        """添加一张新表（兼容接口）

        这是为了向后兼容 MultiExcelLoader.add_table() 方法。

        Args:
            file_path: 文件路径
            sheet_name: 工作表名称

        Returns:
            (文档ID, 结构信息)
        """
        return self.add_document(file_path, sheet_name=sheet_name)

    def remove_document(self, doc_id: str) -> bool:
        """删除指定文档

        Args:
            doc_id: 文档ID

        Returns:
            是否删除成功
        """
        if doc_id not in self._documents:
            return False

        del self._documents[doc_id]
        del self._doc_infos[doc_id]

        # 如果删除的是活跃文档，切换到另一个文档或设为None
        if self._active_doc_id == doc_id:
            if self._documents:
                self._active_doc_id = next(iter(self._documents.keys()))
            else:
                self._active_doc_id = None

        return True

    def remove_table(self, table_id: str) -> bool:
        """删除指定表（兼容接口）"""
        return self.remove_document(table_id)

    def get_document(self, doc_id: str) -> Optional[BaseDocumentLoader]:
        """获取指定文档的加载器"""
        return self._documents.get(doc_id)

    def get_table(self, table_id: str) -> Optional[BaseDocumentLoader]:
        """获取指定表的加载器（兼容接口）"""
        return self.get_document(table_id)

    def get_document_info(self, doc_id: str) -> Optional[DocumentInfo]:
        """获取指定文档的元信息"""
        return self._doc_infos.get(doc_id)

    def get_table_info(self, table_id: str) -> Optional[DocumentInfo]:
        """获取指定表的元信息（兼容接口）"""
        return self.get_document_info(table_id)

    def get_active_loader(self) -> Optional[BaseDocumentLoader]:
        """获取当前活跃文档的加载器"""
        if self._active_doc_id:
            return self._documents.get(self._active_doc_id)
        return None

    def get_active_document_info(self) -> Optional[DocumentInfo]:
        """获取当前活跃文档的元信息"""
        if self._active_doc_id:
            return self._doc_infos.get(self._active_doc_id)
        return None

    def get_active_table_info(self) -> Optional[DocumentInfo]:
        """获取当前活跃表的元信息（兼容接口）"""
        return self.get_active_document_info()

    def set_active_document(self, doc_id: str) -> bool:
        """设置当前活跃文档

        Args:
            doc_id: 文档ID

        Returns:
            是否设置成功
        """
        if doc_id not in self._documents:
            return False
        self._active_doc_id = doc_id
        return True

    def set_active_table(self, table_id: str) -> bool:
        """设置当前活跃表（兼容接口）"""
        return self.set_active_document(table_id)

    def list_documents(self) -> List[Dict[str, Any]]:
        """获取所有文档的信息列表"""
        result = []
        for doc_id, info in self._doc_infos.items():
            result.append({
                "id": info.id,
                "filename": info.filename,
                "doc_type": info.doc_type,
                "sheet_name": info.sheet_name,
                "total_rows": info.total_rows,
                "total_columns": info.total_columns,
                "loaded_at": info.loaded_at.isoformat(),
                "is_active": doc_id == self._active_doc_id,
                "is_joined": info.is_joined,
                "source_tables": info.source_tables,
            })
        return result

    def list_tables(self) -> List[Dict[str, Any]]:
        """获取所有表的信息列表（兼容接口）"""
        return self.list_documents()

    def get_table_columns(self, table_id: str) -> List[str]:
        """获取指定表的列名列表（兼容接口）"""
        loader = self.get_document(table_id)
        if loader and loader.is_loaded:
            return list(loader.dataframe.columns)
        return []

    def join_tables(
        self,
        table1_id: str,
        table2_id: str,
        keys1: List[str],
        keys2: List[str],
        join_type: str = "inner",
        new_name: str = "连接表"
    ) -> tuple[str, Dict[str, Any]]:
        """连接两个文档（支持多字段连接）

        注意：仅支持具有 DataFrame 的文档连接。

        Args:
            table1_id: 文档1 ID
            table2_id: 文档2 ID
            keys1: 文档1 连接字段列表
            keys2: 文档2 连接字段列表
            join_type: 连接类型 (inner/left/right/outer)
            new_name: 新表名称

        Returns:
            (新表ID, 结构信息)
        """
        # 验证文档存在
        loader1 = self.get_document(table1_id)
        loader2 = self.get_document(table2_id)
        if not loader1 or not loader2:
            raise ValueError("指定的文档不存在")

        info1 = self.get_document_info(table1_id)
        info2 = self.get_document_info(table2_id)

        df1 = loader1.dataframe
        df2 = loader2.dataframe

        # 验证字段数量一致
        if len(keys1) != len(keys2):
            raise ValueError("两个文档的连接字段数量必须一致")

        if len(keys1) == 0:
            raise ValueError("至少需要指定一个连接字段")

        # 验证字段存在
        for key in keys1:
            if key not in df1.columns:
                raise ValueError(f"文档1中不存在字段: {key}")
        for key in keys2:
            if key not in df2.columns:
                raise ValueError(f"文档2中不存在字段: {key}")

        # 验证连接类型
        valid_join_types = ["inner", "left", "right", "outer"]
        if join_type not in valid_join_types:
            raise ValueError(f"不支持的连接类型: {join_type}，可选: {valid_join_types}")

        # 执行连接
        merged_df = pd.merge(
            df1, df2,
            left_on=keys1,
            right_on=keys2,
            how=join_type,
            suffixes=('_文档1', '_文档2')
        )

        # 创建新的 Excel 加载器来存储连接结果
        new_loader = ExcelLoader()
        new_loader._df = merged_df
        new_loader._file_path = f"[连接表] {new_name}"
        new_loader._sheet_name = "merged"
        new_loader._all_sheets = ["merged"]

        # 生成唯一ID
        doc_id = str(uuid.uuid4())[:8]

        # 存储文档信息
        self._documents[doc_id] = new_loader
        self._doc_infos[doc_id] = DocumentInfo(
            id=doc_id,
            filename=f"🔗 {new_name}",
            file_path=f"[连接表] {new_name}",
            doc_type="excel",  # 连接结果作为 Excel 表
            sheet_name="merged",
            total_rows=len(merged_df),
            total_columns=len(merged_df.columns),
            is_joined=True,
            source_tables=[info1.filename, info2.filename],
        )

        # 自动设为活跃文档
        self._active_doc_id = doc_id

        return doc_id, new_loader.get_structure()

    def get_active_summary(self) -> str:
        """获取当前活跃文档的摘要"""
        loader = self.get_active_loader()
        if loader:
            return loader.get_summary()
        return "未加载任何文档"

    def get_summary(self) -> str:
        """获取当前活跃文档的摘要（兼容接口）"""
        return self.get_active_summary()

    @property
    def dataframe(self) -> pd.DataFrame:
        """获取当前活跃文档的 DataFrame（兼容接口）"""
        loader = self.get_active_loader()
        if loader:
            return loader.dataframe
        raise ValueError("未加载任何文档")


# 向后兼容别名
MultiExcelLoader = MultiDocumentLoader


# 全局实例
_loader: Optional[MultiDocumentLoader] = None


def get_loader() -> MultiDocumentLoader:
    """获取全局 MultiDocumentLoader 实例"""
    global _loader
    if _loader is None:
        _loader = MultiDocumentLoader()
    return _loader


def reset_loader() -> None:
    """重置全局 MultiDocumentLoader 实例"""
    global _loader
    _loader = MultiDocumentLoader()
