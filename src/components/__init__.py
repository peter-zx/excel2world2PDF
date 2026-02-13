"""
UI组件模块
"""
import streamlit as st
from typing import List
from src.services.template_service import template_service


def render_sidebar(current_step: str) -> str:
    """渲染侧边栏导航"""
    st.sidebar.title("📋 导航")
    
    steps = {
        "template": "📄 模板管理",
        "data": "📊 数据导入",
        "generate": "🚀 批量生成"
    }
    
    for step_key, step_name in steps.items():
        btn_type = "primary" if current_step == step_key else "secondary"
        if st.sidebar.button(step_name, key=f"nav_{step_key}", use_container_width=True, type=btn_type):
            return step_key
    
    st.sidebar.divider()
    templates = template_service.list_templates()
    st.sidebar.info(f"📚 已保存: {len(templates)} 个模板")
    
    return current_step


def show_mapping_config(mapping_type: str, mapping_data: dict, excel_columns: List[str] = None):
    """显示映射配置"""
    with st.expander("查看映射配置", expanded=False):
        st.write(f"**映射类型:** {mapping_type}")
        st.write("**映射内容:**")
        st.json(mapping_data)
        
        if excel_columns:
            st.write("**Excel列名:**")
            st.write(excel_columns)


def show_success(message: str):
    """显示成功消息"""
    st.success(message)


def show_error(message: str):
    """显示错误消息"""
    st.error(message)


def show_warning(message: str):
    """显示警告消息"""
    st.warning(message)


def show_info(message: str):
    """显示信息消息"""
    st.info(message)
