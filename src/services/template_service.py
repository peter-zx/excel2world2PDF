"""
模板管理服务
支持位置映射模式：精确记录变量位置，避免全文替换错误
"""
import json
import uuid
from pathlib import Path
from typing import Dict, List, Optional

from ..config import TEMPLATES_DIR, CONFIGS_DIR
from ..models.schemas import TemplateConfig


class TemplateService:
    """模板管理服务"""
    
    def __init__(self):
        self.templates_dir = TEMPLATES_DIR
        self.configs_dir = CONFIGS_DIR
    
    def create_location_template(
        self,
        template_name: str,
        original_filename: str,
        docx_bytes: bytes,
        location_mapping: Dict[str, Dict],  # {变量名: {element_id, start, end, length, original_text}}
        description: str = ""
    ) -> TemplateConfig:
        """
        创建位置映射模式模板
        
        Args:
            location_mapping: 位置映射
                {
                    "姓名": {
                        "element_id": "para_3",
                        "element_type": "paragraph",
                        "start": 10,
                        "end": 12,
                        "length": 2,
                        "original_text": "陈长"
                    }
                }
        """
        template_id = str(uuid.uuid4())[:8]
        template_filename = f"{template_id}_{original_filename}"
        template_path = self.templates_dir / template_filename
        
        with open(template_path, "wb") as f:
            f.write(docx_bytes)
        
        config = TemplateConfig(
            template_id=template_id,
            template_name=template_name,
            original_filename=original_filename,
            template_filename=template_filename,
            location_mapping=location_mapping,
            description=description
        )
        
        self.save_config(config)
        return config
    
    def save_config(self, config: TemplateConfig) -> None:
        """保存模板配置"""
        config_path = self.configs_dir / f"{config.template_id}.json"
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config.to_dict(), f, ensure_ascii=False, indent=2)
    
    def load_config(self, template_id: str) -> Optional[TemplateConfig]:
        """加载模板配置"""
        config_path = self.configs_dir / f"{template_id}.json"
        if not config_path.exists():
            return None
        
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        return TemplateConfig.from_dict(data)
    
    def list_templates(self) -> List[TemplateConfig]:
        """列出所有模板"""
        templates = []
        for config_file in self.configs_dir.glob("*.json"):
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            templates.append(TemplateConfig.from_dict(data))
        
        templates.sort(key=lambda x: x.updated_at, reverse=True)
        return templates
    
    def delete_template(self, template_id: str) -> bool:
        """删除模板"""
        config = self.load_config(template_id)
        if not config:
            return False
        
        template_path = self.templates_dir / config.template_filename
        if template_path.exists():
            template_path.unlink()
        
        config_path = self.configs_dir / f"{template_id}.json"
        if config_path.exists():
            config_path.unlink()
        
        return True
    
    def get_template_path(self, template_id: str) -> Optional[Path]:
        """获取模板文件路径"""
        config = self.load_config(template_id)
        if not config:
            return None
        
        template_path = self.templates_dir / config.template_filename
        if template_path.exists():
            return template_path
        return None
    
    def get_template_bytes(self, template_id: str) -> Optional[bytes]:
        """获取模板文件二进制"""
        path = self.get_template_path(template_id)
        if path:
            with open(path, "rb") as f:
                return f.read()
        return None


# 单例
template_service = TemplateService()
