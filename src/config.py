"""
项目配置文件
"""
import os
from pathlib import Path

# 项目根目录
BASE_DIR = Path(__file__).resolve().parent.parent

# 存储目录
STORAGE_DIR = BASE_DIR / "src" / "storage"
TEMPLATES_DIR = STORAGE_DIR / "templates"
CONFIGS_DIR = STORAGE_DIR / "configs"
OUTPUTS_DIR = STORAGE_DIR / "outputs"

# 确保目录存在
for dir_path in [TEMPLATES_DIR, CONFIGS_DIR, OUTPUTS_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# 占位符模式
PLACEHOLDER_PATTERN = "【(.*?)】"  # 匹配【变量名】格式
PLACEHOLDER_TEMPLATE = "【{}】"    # 生成【变量名】格式

# 模板变量格式（docxtpl使用）
VAR_TEMPLATE = "{{{{ {} }}}}"     # 生成 {{变量名}} 格式
