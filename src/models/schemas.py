"""
数据模型定义
支持两种映射格式：
1. location_mapping: 精确位置映射
2. text_mapping: 简单文本映射
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from datetime import datetime
import json


@dataclass
class TemplateConfig:
    """
    模板配置
    """
    template_id: str
    template_name: str
    original_filename: str
    template_filename: str
    
    # 新格式：位置映射
    location_mapping: Dict[str, Dict] = field(default_factory=dict)
    
    # 旧格式：文本映射（向后兼容）
    text_mapping: Dict[str, str] = field(default_factory=dict)
    
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
            "text_mapping": self.text_mapping,
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
            text_mapping=data.get("text_mapping", {}),
            created_at=data.get("created_at", ""),
            updated_at=data.get("updated_at", ""),
            description=data.get("description", "")
        )
    
    def get_mapping(self) -> Dict:
        """获取有效的映射（优先使用location_mapping）"""
        if self.location_mapping:
            return {"type": "location", "data": self.location_mapping}
        elif self.text_mapping:
            return {"type": "text", "data": self.text_mapping}
        return {"type": "none", "data": {}}
    
    def get_excel_columns(self) -> List[str]:
        """获取需要的Excel列名"""
        if self.location_mapping:
            return list(self.location_mapping.keys())
        elif self.text_mapping:
            return list(self.text_mapping.keys())
        return []


@dataclass
class GenerationTask:
    """生成任务"""
    template_id: str
    data_rows: List[Dict[str, str]]
    output_filenames: List[str] = field(default_factory=list)
    status: str = "pending"
    error_message: str = ""
