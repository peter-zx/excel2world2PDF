"""
批量生成页面
"""
import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from datetime import datetime

from src.services.template_service import template_service
from src.services.word_service import word_service
from src.components import show_success, show_error, show_warning


def render_column_mapping_display(column_mapping: dict):
    """显示列映射配置"""
    if column_mapping:
        with st.expander("📋 列映射配置", expanded=False):
            for var_name, col_name in column_mapping.items():
                st.write(f"**{var_name}** ← `{col_name}`")


def transform_data(df: pd.DataFrame, column_mapping: dict) -> list:
    """根据列映射转换数据"""
    transformed_data = []
    
    for _, row in df.iterrows():
        new_row = {}
        for var_name, col_name in column_mapping.items():
            if col_name in df.columns:
                value = row[col_name]
                if pd.isna(value):
                    new_row[var_name] = ""
                elif isinstance(value, (pd.Timestamp, datetime)):
                    new_row[var_name] = value.strftime("%Y-%m-%d")
                else:
                    new_row[var_name] = str(value)
        transformed_data.append(new_row)
    
    return transformed_data


def render_generate_page():
    """渲染批量生成页面"""
    st.header("🚀 步骤3: 批量生成")
    
    # 检查前置条件
    if not st.session_state.selected_template:
        show_warning("⚠️ 请先选择模板")
        return
    
    if st.session_state.uploaded_df is None:
        show_warning("⚠️ 请先上传数据")
        return
    
    template = st.session_state.selected_template
    df = st.session_state.uploaded_df
    column_mapping = st.session_state.get("column_mapping", {})
    
    # 显示状态
    c1, c2 = st.columns(2)
    c1.info(f"**模板:** {template.template_name}")
    c2.info(f"**数据:** {len(df)} 条")
    
    # 显示列映射
    render_column_mapping_display(column_mapping)
    
    # 生成按钮
    if st.button("开始生成", type="primary", use_container_width=True):
        if not column_mapping:
            show_error("请先在「数据导入」页面配置列映射")
            return
        
        with st.spinner("生成中..."):
            try:
                template_bytes = template_service.get_template_bytes(template.template_id)
                if not template_bytes:
                    show_error("模板不存在")
                    return
                
                # 转换数据
                transformed_data = transform_data(df, column_mapping)
                
                # 获取映射信息
                mapping_info = template.get_mapping()
                
                # 生成文档
                if mapping_info['type'] == 'location':
                    files = word_service.batch_generate_by_location(
                        template_bytes, transformed_data, mapping_info['data']
                    )
                elif mapping_info['type'] == 'text':
                    files = word_service.batch_generate_by_text(
                        template_bytes, transformed_data, mapping_info['data']
                    )
                else:
                    show_error("模板没有配置映射")
                    return
                
                st.session_state.generated_files = files
                show_success(f"成功生成 {len(files)} 份合同！")
                
            except Exception as e:
                show_error(f"失败: {e}")
    
    # 下载
    if st.session_state.generated_files:
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as z:
            for fn, fb in st.session_state.generated_files:
                z.writestr(fn, fb)
        zip_buf.seek(0)
        
        st.download_button(
            label="下载全部合同",
            data=zip_buf,
            file_name=f"合同_{datetime.now():%Y%m%d_%H%M%S}.zip",
            mime="application/zip",
            use_container_width=True
        )
