import streamlit as st
import pandas as pd
from utils.query_folder_engine import extract_company_name, query_from_file

st.set_page_config(page_title="📊 财报数据查询BOT", layout="wide")
st.title("📊 财报数据处理与查询 BOT")

st.sidebar.header("📌 功能菜单")
menu_option = st.sidebar.radio("请选择操作", [
    "🔍 查询公司指标",
    "📁 处理原始数据（开发中）",
    "🧠 自定义指标（占位）"
])

# 🔍 功能 1：查询指标
if menu_option == "🔍 查询公司指标":
    st.subheader("🔍 查询公司在某日的某个财务指标值（原始表结构）")

    uploaded_files = st.file_uploader(
        "📤 上传多个原始 Excel 财报文件（支持多个公司）",
        type="xlsx",
        accept_multiple_files=True
    )

    if uploaded_files:
        # 提取所有公司简称
        company_map = {}
        for f in uploaded_files:
            name = extract_company_name(f)
            if name:
                company_map[name] = f

        if company_map:
            st.success(f"✅ 成功识别 {len(company_map)} 家公司")
            selected_company = st.selectbox("请选择公司简称", sorted(company_map.keys()))
            indicator = st.text_input("请输入指标名称（如：资产总计）", value="资产总计")
            date_str = st.text_input("请输入日期（格式 YYYY-MM-DD）", value="2025-03-31")

            if st.button("🔍 开始查询"):
                result = query_from_file(
                    company_map[selected_company],
                    selected_company,
                    indicator,
                    date_str
                )
                st.write("🔎 查询结果：")
                st.success(f"{selected_company} - {indicator} @ {date_str} = {result}")
        else:
            st.warning("⚠️ 未能识别出任何公司简称，请确保上传的是原始表结构（第3行第2列为公司简称）")

# 📁 功能 2：读取并处理原始数据（暂未开发）
elif menu_option == "📁 处理原始数据（开发中）":
    st.subheader("📁 原始数据清洗（敬请期待）")
    st.info("未来将支持上传原始表后，自动生成结构化数据并导出。")

# 🧠 功能 3：自定义指标（占位）
elif menu_option == "🧠 自定义指标（占位）":
    st.subheader("🧠 自定义指标（敬请期待）")
    st.info("如 ROE = 净利润 / 所有者权益，将在后续版本中开放。")
