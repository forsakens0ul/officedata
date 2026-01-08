"""文档加载器抽象基类 - 定义统一接口"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd


class BaseDocumentLoader(ABC):
    """文档加载器抽象基类

    所有文档加载器（Excel、PowerPoint、Word等）都应继承此类
    并实现其抽象方法，以提供统一的接口。
    """

    # 子类需要重写此类变量，指定支持的文件扩展名
    SUPPORTED_EXTENSIONS: List[str] = []

    def __init__(self):
        """初始化加载器"""
        self._df: Optional[pd.DataFrame] = None
        self._file_path: Optional[str] = None
        self._metadata: Dict[str, Any] = {}

    @property
    def is_loaded(self) -> bool:
        """是否已加载文件"""
        return self._df is not None

    @property
    def dataframe(self) -> pd.DataFrame:
        """获取 DataFrame

        Returns:
            pandas DataFrame，包含文档的结构化数据

        Raises:
            ValueError: 如果未加载文件
        """
        if self._df is None:
            raise ValueError("未加载文件")
        return self._df

    @property
    def file_path(self) -> Optional[str]:
        """获取文件路径"""
        return self._file_path

    @property
    def metadata(self) -> Dict[str, Any]:
        """获取元数据"""
        return self._metadata

    @abstractmethod
    def load(self, file_path: str, **kwargs) -> Dict[str, Any]:
        """加载文档

        Args:
            file_path: 文件路径
            **kwargs: 其他加载参数（由子类定义）

        Returns:
            文档结构信息字典，至少包含：
            - total_rows: 总行数
            - total_columns: 总列数
            - 其他文档特定信息

        Raises:
            FileNotFoundError: 如果文件不存在
            ValueError: 如果文件格式不支持
        """
        pass

    @abstractmethod
    def get_structure(self) -> Dict[str, Any]:
        """获取文档结构信息

        Returns:
            结构信息字典，至少包含：
            - file_path: 文件路径
            - total_rows: 总行数
            - total_columns: 总列数
            - columns: 列信息列表

        Raises:
            ValueError: 如果未加载文件
        """
        pass

    @abstractmethod
    def get_preview(self, n_rows: Optional[int] = None) -> Dict[str, Any]:
        """获取数据预览

        Args:
            n_rows: 预览行数，默认使用配置值

        Returns:
            预览数据字典，至少包含：
            - columns: 列名列表
            - data: 数据行列表（字典列表）
            - preview_rows: 预览行数
            - total_rows: 总行数

        Raises:
            ValueError: 如果未加载文件
        """
        pass

    @abstractmethod
    def get_summary(self) -> str:
        """获取文档摘要信息（用于 Agent 上下文）

        Returns:
            Markdown 格式的摘要字符串，用于 LLM 上下文
            应包含文档基本信息、结构说明和数据预览
        """
        pass

    @classmethod
    def supports_extension(cls, extension: str) -> bool:
        """检查是否支持指定的文件扩展名

        Args:
            extension: 文件扩展名（如 '.xlsx', '.pptx'）

        Returns:
            是否支持
        """
        return extension.lower() in [ext.lower() for ext in cls.SUPPORTED_EXTENSIONS]

    @classmethod
    def get_supported_extensions(cls) -> List[str]:
        """获取支持的文件扩展名列表

        Returns:
            支持的扩展名列表
        """
        return cls.SUPPORTED_EXTENSIONS.copy()

    def _validate_file(self, file_path: str) -> Path:
        """验证文件路径和扩展名

        Args:
            file_path: 文件路径

        Returns:
            Path 对象

        Raises:
            FileNotFoundError: 如果文件不存在
            ValueError: 如果文件格式不支持
        """
        path = Path(file_path)

        if not path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")

        if not self.supports_extension(path.suffix):
            raise ValueError(
                f"不支持的文件格式: {path.suffix}，"
                f"支持的格式: {', '.join(self.SUPPORTED_EXTENSIONS)}"
            )

        return path
