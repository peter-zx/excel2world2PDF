"""
数据导入页面
"""
import streamlit as st
import pandas as pd
from datetime import datetime

from src.services.template_service import template_service
from src.services.excel_service import excel_service
from src.utils import generate_excel_template
from src.components import show_success, show_error, show_warning, show_info


def render_template_selector():
    """渲染模板选择器"""
    templates = template_service.list_templates()
    if not templates:
        show_warning("⚠️ 请先创建模板")
        return None
    
    st.subheader("📚 选择模板")
    template_options = {f"{t.template_name}": t for t in templates}
    selected_key = st.selectbox("选择模板", options=list(template_options.keys()))
    
    return template_options[selected_key]


def render_column_mapping(var_names: list, excel_columns: list):
    """渲染列映射配置"""
    st.divider()
    st.subheader("🔗 变量列映射配置")
    st.caption("将模板变量映射到Excel列名")
    
    column_mapping = {}
    cols_per_row = 3
    
    for i in range(0, len(var_names), cols_per_row):
        cols = st.columns(cols_per_row)
        for j, var_name in enumerate(var_names[i:i+cols_per_row]):
            with cols[j]:
                # 尝试自动匹配
                default_idx = 0
                for idx, col in enumerate(excel_columns):
                    if col == var_name or var_name in col or col in var_name:
                        default_idx = idx + 1
                        break
                
                selected_col = st.selectbox(
                    f"**{var_name}**",
                    options=["-- 不映射 --"] + excel_columns,
                    index=default_idx,
                    key=f"col_map_{var_name}"
                )
                
                if selected_col != "-- 不映射 --":
                    column_mapping[var_name] = selected_col
    
    return column_mapping


def render_data_page():
    """渲染数据导入页面"""
    st.header("📊 步骤2: 数据导入")
    
    # 选择模板
    selected = render_template_selector()
    if not selected:
        return
    
    st.session_state.selected_template = selected
    
    # 获取映射信息
    mapping_info = selected.get_mapping()
    var_names = list(mapping_info['data'].keys())
    
    show_info(f"**模板变量:** {', '.join(var_names)}")
    
    # 下载Excel模板
    if st.button("📥 下载Excel模板"):
        if mapping_info['type'] == 'text':
            excel_bytes = generate_excel_template(mapping_info['data'])
        else:
            simple_map = {k: v.get("original_text", "") for k, v in mapping_info['data'].items()}
            excel_bytes = generate_excel_template(simple_map)
        st.download_button(
            label="📥 下载",
            data=excel_bytes,
            file_name=f"{selected.template_name}_模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    st.divider()
    
    # 上传Excel
    st.subheader("📤 上传Excel")
    excel_file = st.file_uploader("选择Excel文件", type=["xlsx", "xls"])
    
    if excel_file:
        df, error = excel_service.read_excel(excel_file.getvalue(), excel_file.name)
        if error:
            show_error(f"读取失败: {error}")
            return
        
        st.session_state.uploaded_df = df
        st.dataframe(df, use_container_width=True)
        show_info(f"共 {len(df)} 条记录")
        
        # 列映射配置
        excel_columns = df.columns.tolist()
        column_mapping = render_column_mapping(var_names, excel_columns)
        
        st.session_state.column_mapping = column_mapping
        
        if column_mapping:
            show_success(f"已配置 {len(column_mapping)} 个映射")
        else:
            show_warning("请配置至少一个映射")
