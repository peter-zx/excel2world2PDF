"""
Excel处理服务
负责读取Excel数据、验证、格式化
"""
import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from datetime import datetime
import tempfile


class ExcelService:
    """Excel处理服务"""
    
    def read_excel(self, file_bytes: bytes, filename: str) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
        """
        读取Excel文件
        
        Returns:
            (DataFrame, 错误信息)
        """
        try:
            df = pd.read_excel(file_bytes)
            return df, None
        except Exception as e:
            return None, str(e)
    
    def get_columns(self, df: pd.DataFrame) -> List[str]:
        """获取列名列表"""
        return df.columns.tolist()
    
    def validate_columns(
        self, 
        df: pd.DataFrame, 
        required_columns: List[str]
    ) -> Tuple[bool, List[str]]:
        """
        验证Excel是否包含所需列
        
        Returns:
            (是否通过, 缺失的列名列表)
        """
        existing_columns = set(df.columns.tolist())
        required_set = set(required_columns)
        
        missing = required_set - existing_columns
        return len(missing) == 0, list(missing)
    
    def format_row_data(self, row: pd.Series) -> Dict[str, str]:
        """
        格式化单行数据，处理特殊类型
        
        Returns:
            字典格式的行数据
        """
        result = {}
        for key, value in row.items():
            if pd.isna(value):
                result[key] = ""
            elif isinstance(value, (pd.Timestamp, datetime)):
                result[key] = value.strftime("%Y-%m-%d")
            else:
                result[key] = str(value)
        return result
    
    def dataframe_to_dict_list(self, df: pd.DataFrame) -> List[Dict[str, str]]:
        """将DataFrame转换为字典列表"""
        return [self.format_row_data(row) for _, row in df.iterrows()]
    
    def preview_data(self, df: pd.DataFrame, rows: int = 5) -> pd.DataFrame:
        """预览前N行数据"""
        return df.head(rows)


# 单例实例
excel_service = ExcelService()
