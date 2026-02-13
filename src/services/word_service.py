"""
Word文档处理服务
支持位置映射模式：根据精确位置替换文本，完美保留格式
"""
from io import BytesIO
from typing import Dict, List, Tuple
from docx import Document
from docx.shared import Pt, RGBColor
from copy import deepcopy


class WordService:
    """Word文档处理服务"""
    
    def replace_by_location(
        self,
        file_bytes: bytes,
        location_mapping: Dict[str, Dict],  # {变量名: {element_id, start, end, original_text}}
        data: Dict[str, str]  # {变量名: 新值}
    ) -> bytes:
        """
        根据位置映射精确替换文本，保留格式
        
        Args:
            file_bytes: Word文档二进制
            location_mapping: 位置映射
            data: 数据映射
        
        Returns:
            替换后的Word文档
        """
        doc = Document(BytesIO(file_bytes))
        
        # 构建元素索引映射
        element_map = {}
        
        # 段落
        for para_idx, para in enumerate(doc.paragraphs):
            element_map[f"para_{para_idx}"] = para
        
        # 表格单元格
        for table_idx, table in enumerate(doc.tables):
            for row_idx, row in enumerate(table.rows):
                for cell_idx, cell in enumerate(row.cells):
                    element_map[f"cell_{table_idx}_{row_idx}_{cell_idx}"] = cell
        
        # 按位置执行替换
        for var_name, new_value in data.items():
            if var_name not in location_mapping:
                continue
            
            loc = location_mapping[var_name]
            element_id = loc["element_id"]
            
            if element_id not in element_map:
                continue
            
            element = element_map[element_id]
            
            # 判断元素类型
            if element_id.startswith("para_"):
                # 段落处理
                self._replace_in_paragraph(element, loc, new_value)
            elif element_id.startswith("cell_"):
                # 表格单元格处理
                self._replace_in_cell(element, loc, new_value)
        
        # 保存
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        return output.getvalue()
    
    def _replace_in_paragraph(self, paragraph, location: Dict, new_text: str):
        """
        在段落中根据位置替换文本，保留格式
        
        关键：段落由多个run组成，每个run有自己的格式
        需要找到目标位置所在的run，保留其格式进行替换
        """
        start = location["start"]
        end = location["end"]
        
        # 获取段落的完整文本
        full_text = paragraph.text
        
        # 验证位置
        if start < 0 or end > len(full_text):
            return
        
        # 构建新文本
        new_full_text = full_text[:start] + new_text + full_text[end:]
        
        # 找到目标run并替换
        # 策略：找到包含start位置的run，在该run中执行替换
        current_pos = 0
        
        for run in paragraph.runs:
            run_start = current_pos
            run_end = current_pos + len(run.text)
            
            # 检查目标位置是否在这个run中
            if run_start <= start < run_end:
                # 计算在run内的位置
                in_run_start = start - run_start
                in_run_end = min(end - run_start, len(run.text))
                
                # 执行替换
                original_run_text = run.text
                new_run_text = (
                    original_run_text[:in_run_start] + 
                    new_text + 
                    original_run_text[in_run_end:]
                )
                run.text = new_run_text
                
                # 格式已自动保留
                break
            
            current_pos = run_end
        
        # 处理跨run的情况（简化：合并到第一个run）
        # 如果替换文本跨越多个run，需要更复杂的处理
    
    def _replace_in_cell(self, cell, location: Dict, new_text: str):
        """在表格单元格中替换文本"""
        # 简化处理：直接替换第一个段落的文本
        if cell.paragraphs:
            self._replace_in_paragraph(cell.paragraphs[0], location, new_text)
    
    def batch_generate_by_location(
        self,
        template_bytes: bytes,
        data_list: List[Dict[str, str]],
        location_mapping: Dict[str, Dict]
    ) -> List[Tuple[str, bytes]]:
        """
        批量生成文档（位置映射模式）
        
        Returns:
            [(文件名, 文档二进制), ...]
        """
        results = []
        
        for idx, data in enumerate(data_list):
            try:
                doc_bytes = self.replace_by_location(
                    template_bytes,
                    location_mapping,
                    data
                )
                
                # 生成文件名
                name = data.get("姓名", data.get("name", f"合同_{idx+1}"))
                safe_name = "".join(
                    c for c in str(name) 
                    if c.isalnum() or c in (' ', '-', '_') or '\u4e00' <= c <= '\u9fff'
                )
                filename = f"{safe_name}_合同.docx"
                
                results.append((filename, doc_bytes))
                
            except Exception as e:
                # 记录错误但继续处理其他数据
                print(f"Error generating contract {idx+1}: {e}")
                continue
        
        return results


# 单例
word_service = WordService()
