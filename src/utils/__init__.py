"""
工具函数模块
"""
import re
from typing import Dict, List
import pandas as pd
from io import BytesIO
from datetime import datetime


def extract_candidates(text: str) -> List[Dict]:
    """从文本中提取候选替换内容"""
    candidates = []
    seen = set()
    
    patterns = [
        (r'\d{15,18}[Xx]?', "身份证号"),
        (r'1[3-9]\d{9}', "手机号"),
        (r'20\d{2}', "年份"),
        (r'\d{4,}(?:\.\d{1,2})?', "金额/数字"),
    ]
    
    for pattern, ptype in patterns:
        for m in re.finditer(pattern, text):
            val = m.group()
            if val not in seen:
                candidates.append({
                    "text": val,
                    "type": ptype,
                    "start": m.start(),
                    "end": m.end()
                })
                seen.add(val)
    
    return candidates


def generate_excel_template(mapping: Dict[str, str]) -> bytes:
    """生成Excel模板"""
    df = pd.DataFrame(columns=list(mapping.keys()))
    example = {k: v if v.isdigit() else f"示例{k}" for k, v in mapping.items()}
    df = pd.concat([df, pd.DataFrame([example])], ignore_index=True)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='数据')
    output.seek(0)
    return output.getvalue()


def safe_filename(name: str, default: str = "文件") -> str:
    """生成安全的文件名"""
    safe = "".join(
        c for c in str(name) 
        if c.isalnum() or c in (' ', '-', '_') or '\u4e00' <= c <= '\u9fff'
    )
    return safe if safe else default


def format_datetime(dt: datetime = None, fmt: str = "%Y%m%d_%H%M%S") -> str:
    """格式化日期时间"""
    if dt is None:
        dt = datetime.now()
    return dt.strftime(fmt)
