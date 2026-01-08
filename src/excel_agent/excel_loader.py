"""Excel 加载与管理模块 - 支持多表管理"""

import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

from .base_loader import BaseDocumentLoader
from .config import get_config


@dataclass
class TableInfo:
    """表的元信息"""
    id: str
    filename: str
    file_path: str
    sheet_name: str
    total_rows: int
    total_columns: int
    loaded_at: datetime = field(default_factory=datetime.now)
    is_joined: bool = False  # 是否为连接表
    source_tables: List[str] = field(default_factory=list)  # 源表名称列表


class ExcelLoader(BaseDocumentLoader):
    """Excel 文件加载器"""

    SUPPORTED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm']

    def __init__(self):
        super().__init__()
        self._sheet_name: Optional[str] = None
        self._all_sheets: List[str] = []

    def load(self, file_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """加载 Excel 文件
        
        Args:
            file_path: Excel 文件路径
            sheet_name: 工作表名称，默认加载第一个
            
        Returns:
            文件结构信息
        """
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件不存在: {file_path}")
        
        if path.suffix.lower() not in ['.xlsx', '.xls', '.xlsm']:
            raise ValueError(f"不支持的文件格式: {path.suffix}")
        
        # 获取所有工作表名称
        xlsx = pd.ExcelFile(file_path)
        self._all_sheets = xlsx.sheet_names
        
        # 确定要加载的工作表
        if sheet_name is None:
            sheet_name = self._all_sheets[0]
        elif sheet_name not in self._all_sheets:
            raise ValueError(f"工作表 '{sheet_name}' 不存在，可用工作表: {self._all_sheets}")
        
        # 加载数据
        self._df = pd.read_excel(file_path, sheet_name=sheet_name)
        self._file_path = file_path
        self._sheet_name = sheet_name
        
        return self.get_structure()
    
    def get_structure(self) -> Dict[str, Any]:
        """获取 Excel 结构信息"""
        if self._df is None:
            raise ValueError("未加载 Excel 文件")
        
        config = get_config()
        
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
            "sheet_name": self._sheet_name,
            "all_sheets": self._all_sheets,
            "total_rows": len(self._df),
            "total_columns": len(self._df.columns),
            "columns": columns_info,
        }
    
    def get_preview(self, n_rows: Optional[int] = None) -> Dict[str, Any]:
        """获取数据预览
        
        Args:
            n_rows: 预览行数，默认使用配置值
            
        Returns:
            预览数据
        """
        if self._df is None:
            raise ValueError("未加载 Excel 文件")
        
        config = get_config()
        if n_rows is None:
            n_rows = config.excel.max_preview_rows
        
        preview_df = self._df.head(n_rows)
        
        return {
            "columns": list(self._df.columns),
            "data": preview_df.to_dict(orient="records"),
            "preview_rows": len(preview_df),
            "total_rows": len(self._df),
        }
    
    def get_summary(self) -> str:
        """获取 Excel 摘要信息（用于 Agent 上下文）"""
        if self._df is None:
            return "未加载 Excel 文件"
        
        structure = self.get_structure()
        preview = self.get_preview()
        
        lines = [
            f"📊 **已加载 Excel 文件**: {structure['file_path']}",
            f"📋 **当前工作表**: {structure['sheet_name']}",
            f"📑 **所有工作表**: {', '.join(structure['all_sheets'])}",
            f"📏 **数据规模**: {structure['total_rows']} 行 × {structure['total_columns']} 列",
            "",
            "**列信息**:",
        ]
        
        for col in structure['columns']:
            lines.append(f"  - `{col['name']}` ({col['dtype']}): {col['non_null_count']} 非空值")
        
        lines.append("")
        lines.append(f"**前 {preview['preview_rows']} 行数据预览**:")
        
        # 简单表格格式
        if preview['data']:
            headers = preview['columns']
            lines.append("| " + " | ".join(str(h) for h in headers) + " |")
            lines.append("| " + " | ".join("---" for _ in headers) + " |")
            for row in preview['data']:
                values = [str(row.get(h, ""))[:20] for h in headers]  # 截断长值
                lines.append("| " + " | ".join(values) + " |")
        
        return "\n".join(lines)


class MultiExcelLoader:
    """多表管理器 - 管理多个 ExcelLoader 实例"""
    
    def __init__(self):
        self._tables: Dict[str, ExcelLoader] = {}  # table_id -> ExcelLoader
        self._table_infos: Dict[str, TableInfo] = {}  # table_id -> TableInfo
        self._active_table_id: Optional[str] = None
    
    @property
    def is_loaded(self) -> bool:
        """是否有任何表已加载"""
        return len(self._tables) > 0
    
    @property
    def active_table_id(self) -> Optional[str]:
        """获取当前活跃表ID"""
        return self._active_table_id
    
    def add_table(self, file_path: str, sheet_name: Optional[str] = None) -> tuple[str, Dict[str, Any]]:
        """添加一张新表
        
        Args:
            file_path: Excel 文件路径
            sheet_name: 工作表名称
            
        Returns:
            (表ID, 结构信息)
        """
        # 创建新的加载器并加载数据
        loader = ExcelLoader()
        structure = loader.load(file_path, sheet_name)
        
        # 生成唯一ID
        table_id = str(uuid.uuid4())[:8]
        
        # 获取文件名
        filename = Path(file_path).name
        
        # 存储表信息
        self._tables[table_id] = loader
        self._table_infos[table_id] = TableInfo(
            id=table_id,
            filename=filename,
            file_path=file_path,
            sheet_name=structure["sheet_name"],
            total_rows=structure["total_rows"],
            total_columns=structure["total_columns"],
        )
        
        # 自动设为活跃表
        self._active_table_id = table_id
        
        return table_id, structure
    
    def remove_table(self, table_id: str) -> bool:
        """删除指定表
        
        Args:
            table_id: 表ID
            
        Returns:
            是否删除成功
        """
        if table_id not in self._tables:
            return False
        
        del self._tables[table_id]
        del self._table_infos[table_id]
        
        # 如果删除的是活跃表，切换到另一张表或设为None
        if self._active_table_id == table_id:
            if self._tables:
                self._active_table_id = next(iter(self._tables.keys()))
            else:
                self._active_table_id = None
        
        return True
    
    def get_table(self, table_id: str) -> Optional[ExcelLoader]:
        """获取指定表的加载器"""
        return self._tables.get(table_id)
    
    def get_table_info(self, table_id: str) -> Optional[TableInfo]:
        """获取指定表的元信息"""
        return self._table_infos.get(table_id)
    
    def get_active_loader(self) -> Optional[ExcelLoader]:
        """获取当前活跃表的加载器"""
        if self._active_table_id:
            return self._tables.get(self._active_table_id)
        return None
    
    def get_active_table_info(self) -> Optional[TableInfo]:
        """获取当前活跃表的元信息"""
        if self._active_table_id:
            return self._table_infos.get(self._active_table_id)
        return None
    
    def set_active_table(self, table_id: str) -> bool:
        """设置当前活跃表
        
        Args:
            table_id: 表ID
            
        Returns:
            是否设置成功
        """
        if table_id not in self._tables:
            return False
        self._active_table_id = table_id
        return True
    
    def list_tables(self) -> List[Dict[str, Any]]:
        """获取所有表的信息列表"""
        result = []
        for table_id, info in self._table_infos.items():
            result.append({
                "id": info.id,
                "filename": info.filename,
                "sheet_name": info.sheet_name,
                "total_rows": info.total_rows,
                "total_columns": info.total_columns,
                "loaded_at": info.loaded_at.isoformat(),
                "is_active": table_id == self._active_table_id,
                "is_joined": info.is_joined,
                "source_tables": info.source_tables,
            })
        return result
    
    def get_table_columns(self, table_id: str) -> List[str]:
        """获取指定表的列名列表"""
        loader = self.get_table(table_id)
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
        """连接两张表（支持多字段连接）
        
        Args:
            table1_id: 表1 ID
            table2_id: 表2 ID
            keys1: 表1 连接字段列表
            keys2: 表2 连接字段列表
            join_type: 连接类型 (inner/left/right/outer)
            new_name: 新表名称
            
        Returns:
            (新表ID, 结构信息)
        """
        # 验证表存在
        loader1 = self.get_table(table1_id)
        loader2 = self.get_table(table2_id)
        if not loader1 or not loader2:
            raise ValueError("指定的表不存在")
        
        info1 = self.get_table_info(table1_id)
        info2 = self.get_table_info(table2_id)
        
        df1 = loader1.dataframe
        df2 = loader2.dataframe
        
        # 验证字段数量一致
        if len(keys1) != len(keys2):
            raise ValueError("两表的连接字段数量必须一致")
        
        if len(keys1) == 0:
            raise ValueError("至少需要指定一个连接字段")
        
        # 验证字段存在
        for key in keys1:
            if key not in df1.columns:
                raise ValueError(f"表1中不存在字段: {key}")
        for key in keys2:
            if key not in df2.columns:
                raise ValueError(f"表2中不存在字段: {key}")
        
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
            suffixes=('_表1', '_表2')
        )
        
        # 创建新的加载器
        new_loader = ExcelLoader()
        new_loader._df = merged_df
        new_loader._file_path = f"[连接表] {new_name}"
        new_loader._sheet_name = "merged"
        new_loader._all_sheets = ["merged"]
        
        # 生成唯一ID
        table_id = str(uuid.uuid4())[:8]
        
        # 存储表信息
        self._tables[table_id] = new_loader
        self._table_infos[table_id] = TableInfo(
            id=table_id,
            filename=f"🔗 {new_name}",
            file_path=f"[连接表] {new_name}",
            sheet_name="merged",
            total_rows=len(merged_df),
            total_columns=len(merged_df.columns),
            is_joined=True,
            source_tables=[info1.filename, info2.filename],
        )
        
        # 自动设为活跃表
        self._active_table_id = table_id
        
        return table_id, new_loader.get_structure()
    
    def get_active_summary(self) -> str:
        """获取当前活跃表的摘要"""
        loader = self.get_active_loader()
        if loader:
            return loader.get_summary()
        return "未加载 Excel 文件"
    
    def get_summary(self) -> str:
        """获取当前活跃表的摘要（兼容旧接口）"""
        return self.get_active_summary()
    
    @property
    def dataframe(self) -> pd.DataFrame:
        """获取当前活跃表的 DataFrame（兼容旧接口）"""
        loader = self.get_active_loader()
        if loader:
            return loader.dataframe
        raise ValueError("未加载 Excel 文件")


# 全局实例 - 使用多表管理器
_loader: Optional[MultiExcelLoader] = None


def get_loader() -> MultiExcelLoader:
    """获取全局 MultiExcelLoader 实例"""
    global _loader
    if _loader is None:
        _loader = MultiExcelLoader()
    return _loader


def reset_loader() -> None:
    """重置全局 MultiExcelLoader 实例"""
    global _loader
    _loader = MultiExcelLoader()
