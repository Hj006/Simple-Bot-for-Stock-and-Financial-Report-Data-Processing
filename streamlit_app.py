# Simple-Bot-for-Stock-and-Financial-Report-Data-Processing\streamlit_app.py
import os
import streamlit as st
import pandas as pd
import win32com.client
import pythoncom
from utils.query_folder_engine import extract_company_name, query_from_file
from processor.process_excel import process_excel_file

def contains_non_ascii(s):
    return not all(ord(c) < 128 for c in s)

def process_raw_data():

    st.subheader("📁 原始数据清洗与结构化处理")
    raw_path = st.text_input("📂 请输入包含原始 Excel 的文件夹路径（例如：D:/原始财报）", value="E:\code\Simple-Bot-for-Stock-and-Financial-Report-Data-Processing\test")
    st.caption("💡 路径建议使用英文且无特殊字符，例如：D:/data/2025")
    folder_path = os.path.normpath(raw_path.strip()) if raw_path else ""

    if folder_path:
        if contains_non_ascii(folder_path):
            st.warning("⚠️ 当前路径包含中文或特殊字符，可能导致读取失败，请使用英文路径。")
            return

        if not os.path.exists(folder_path):
            st.error("❌ 输入的文件夹路径不存在，请确认路径是否正确。")
            return

        files = [f for f in os.listdir(folder_path) if f.lower().endswith('.xlsx')]

        if not files:
            st.warning("⚠️ 该文件夹中未找到任何 .xlsx 文件。")
            return

        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        st.success(f"✅ 找到 {len(files)} 个 Excel 文件，准备开始处理...")

        for f in files:
            full_path = os.path.join(folder_path, f)
            output_path = os.path.join(folder_path, f.replace('.xlsx', '_v1.xlsx'))
            st.write(f"🔄 正在处理：{f}")

            try:
                process_excel_file(full_path, output_path)
                st.success(f"✅ 已处理完成：{f}")
            except Exception as e:
                st.error(f"❌ 处理失败：{f}，原因：{str(e)}")

        st.balloons()
        st.info("🎉 全部处理完成！请在原文件夹中查看生成的 *_v1.xlsx 文件。")

def run_query(company_map, indicator, date_str):
    if not company_map:
        st.warning("⚠️ 未能识别出任何公司简称，请确保文件格式正确")
        return

    st.success(f"✅ 成功识别 {len(company_map)} 家公司")

    if st.button("🔍 开始查询所有公司"):
        results = []
        for company_name, file in company_map.items():
            try:
                result = query_from_file(file, company_name, indicator, date_str)
                results.append((company_name, result))
            except Exception as e:
                results.append((company_name, f"❌ 错误：{str(e)}"))

        st.write("📊 查询结果：")
        for company, res in results:
            st.success(f"{company} - {indicator} @ {date_str} = {res}")

    with st.expander("🔎 仅查询某一个公司（可选）"):
        selected_company = st.selectbox("🎯 选择公司进行单独查询", sorted(company_map.keys()))
        if st.button("📌 单独查询该公司"):
            try:
                result = query_from_file(
                    company_map[selected_company],
                    selected_company,
                    indicator,
                    date_str
                )
                st.success(f"{selected_company} - {indicator} @ {date_str} = {result}")
            except Exception as e:
                st.error(f"❌ 查询失败：{str(e)}")

def query_from_folder(indicator, date_str):
    st.subheader("📁 从指定文件夹读取多个原始 Excel 财报文件")
    raw_path = st.text_input("📂 请输入本地文件夹路径（例如：D:/财报数据）")
    st.caption("💡 输入本地绝对路径，如 D:/data 或 /Users/name/reports（必须为本机路径）")

    folder_path = os.path.normpath(raw_path.strip()) if raw_path else ""

    if folder_path and os.path.exists(folder_path):
        all_files = [
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path)
            if f.lower().endswith(".xlsx")
        ]

        company_map = {}
        for file_path in all_files:
            with open(file_path, "rb") as f:
                name = extract_company_name(f)
                if name:
                    company_map[name] = file_path

        run_query(company_map, indicator, date_str)

    elif raw_path:
        st.error("❌ 文件夹路径不存在，请确认路径无误")



st.set_page_config(page_title="📊 财报数据查询BOT", layout="wide")
st.title("📊 财报数据处理与查询 BOT")

st.sidebar.header("📌 功能菜单")
menu_option = st.sidebar.radio("请选择操作", [
    "🔍 查询公司指标",
    "📁 处理原始数据",
    "🧠 自定义指标（占位）"
])

if menu_option == "🔍 查询公司指标":
    st.subheader("🔍 查询公司财务指标")

    # ✅ 公共输入项（上方统一输入）
    indicator = st.text_input("📌 请输入要查询的指标（如：资产总计）", value="资产总计")
    date_str = st.text_input("📅 请输入查询日期（格式 YYYY-MM-DD）", value="2025-03-31")

    mode = st.radio("📂 请选择数据来源方式：", ["从本地文件夹读取", "上传 Excel 文件"])

    if mode == "从本地文件夹读取":
        query_from_folder(indicator, date_str)

    elif mode == "上传 Excel 文件":
        st.subheader("📤 上传多个原始 Excel 财报文件（支持多个公司）")

        uploaded_files = st.file_uploader(
            "请选择一个或多个 Excel 文件（.xlsx）",
            type="xlsx",
            accept_multiple_files=True
        )

        if uploaded_files:
            company_map = {}
            for f in uploaded_files:
                name = extract_company_name(f)
                if name:
                    company_map[name] = f

            run_query(company_map, indicator, date_str)


# 📁 功能 2：读取并处理原始数据
elif menu_option == "📁 处理原始数据":
    from processor.process_excel import process_excel_file

    process_raw_data()

# 🧠 功能 3：自定义指标（占位）
elif menu_option == "🧠 自定义指标（占位）":
    st.subheader("🧠 自定义指标（敬请期待）")
    st.info("如 ROE = 净利润 / 所有者权益，将在后续版本中开放。")
