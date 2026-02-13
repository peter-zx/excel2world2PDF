"""
合同自动填写工具 - 主入口

模块化架构：
- pages/: 页面模块（模板管理、数据导入、批量生成）
- components/: UI组件（侧边栏、消息提示等）
- services/: 业务服务（Word处理、Excel处理、模板管理）
- models/: 数据模型
- utils/: 工具函数
"""
import streamlit as st

from src.pages import render_template_page, render_data_page, render_generate_page
from src.components import render_sidebar


# ==================== 页面配置 ====================
st.set_page_config(
    page_title="合同自动填写工具",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ==================== Session State 初始化 ====================
def init_session_state():
    """初始化session状态"""
    defaults = {
        "current_step": "template",
        "uploaded_template_bytes": None,
        "template_name": "",
        "description": "",
        "doc_elements": [],
        "location_mapping": {},
        "selected_element_id": None,
        "uploaded_df": None,
        "selected_template": None,
        "generated_files": [],
        "column_mapping": {},
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# ==================== 路由控制 ====================
def route_to_page(step: str):
    """根据步骤路由到对应页面"""
    page_renderers = {
        "template": render_template_page,
        "data": render_data_page,
        "generate": render_generate_page,
    }
    
    renderer = page_renderers.get(step)
    if renderer:
        renderer()


# ==================== 主函数 ====================
def main():
    """主函数"""
    # 初始化
    init_session_state()
    
    # 页面标题
    st.title("📝 合同自动填写工具")
    st.markdown("上传合同 → 选择段落 → 配置映射 → 批量生成")
    
    # 渲染侧边栏并处理导航
    new_step = render_sidebar(st.session_state.current_step)
    if new_step != st.session_state.current_step:
        st.session_state.current_step = new_step
        st.rerun()
    
    st.divider()
    
    # 渲染当前页面
    route_to_page(st.session_state.current_step)


if __name__ == "__main__":
    main()
