"""
数据模型定义
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from datetime import datetime
import json


@dataclass
class TemplateConfig:
    """
    模板配置
    
    location_mapping: {变量名: {element_id, start, end, length, original_text}}
    精确记录每个变量在文档中的位置，避免全文替换错误
    """
    template_id: str
    template_name: str
    original_filename: str
    template_filename: str
    
    # 位置映射模式
    location_mapping: Dict[str, Dict] = field(default_factory=dict)
    
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())
    description: str = ""
    
    def to_dict(self) -> dict:
        return {
            "template_id": self.template_id,
            "template_name": self.template_name,
            "original_filename": self.original_filename,
            "template_filename": self.template_filename,
            "location_mapping": self.location_mapping,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "description": self.description
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> "TemplateConfig":
        return cls(
            template_id=data["template_id"],
            template_name=data["template_name"],
            original_filename=data["original_filename"],
            template_filename=data["template_filename"],
            location_mapping=data.get("location_mapping", {}),
            created_at=data.get("created_at", ""),
            updated_at=data.get("updated_at", ""),
            description=data.get("description", "")
        )
    
    def get_excel_columns(self) -> List[str]:
        """获取需要的Excel列名"""
        return list(self.location_mapping.keys())


@dataclass
class GenerationTask:
    """生成任务"""
    template_id: str
    data_rows: List[Dict[str, str]]
    output_filenames: List[str] = field(default_factory=list)
    status: str = "pending"
    error_message: str = ""
