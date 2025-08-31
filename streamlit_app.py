# Simple-Bot-for-Stock-and-Financial-Report-Data-Processing\streamlit_app.py
import os
import re  # ← 新增：用于提取列名中的年份
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
    st.info("请依次输入【原始数据路径】和【处理结果保存路径】，然后点击下方按钮开始处理。")

    # ✅ 输入原始路径
    raw_path = st.text_input(
        "📂 原始 Excel 文件夹路径（例如：D:/原始财报）",
        value=r"E:\code\Simple-Bot-for-Stock-and-Financial-Report-Data-Processing\test"
    )

    # ✅ 输入输出路径
    output_base_path = st.text_input(
        "💾 处理结果保存到哪个文件夹（例如：D:/输出结果）",
        value=r"E:\code\Simple-Bot-for-Stock-and-Financial-Report-Data-Processing\output"
    )

    st.caption("⚠️ 路径建议使用英文且无特殊字符；若文件夹不存在将自动创建。")

    # ✅ 按钮触发处理
    if st.button("🚀 开始批量处理并导出"):

        input_folder = os.path.normpath(raw_path.strip()) if raw_path else ""
        output_folder = os.path.normpath(output_base_path.strip()) if output_base_path else ""

        if not input_folder or not output_folder:
            st.error("❌ 请同时填写输入路径和输出路径")
            return

        if contains_non_ascii(input_folder) or contains_non_ascii(output_folder):
            st.warning("⚠️ 路径中包含中文或特殊字符，建议使用英文路径")
            return

        if not os.path.exists(input_folder):
            st.error("❌ 输入路径不存在")
            return

        os.makedirs(output_folder, exist_ok=True)  # 自动创建输出目录（如不存在）

        files = [f for f in os.listdir(input_folder) if f.lower().endswith('.xlsx')]

        if not files:
            st.warning("⚠️ 未找到任何 .xlsx 文件")
            return

        st.success(f"✅ 找到 {len(files)} 个 Excel 文件，开始处理...")

        pythoncom.CoInitialize()
        win32com.client.Dispatch("Excel.Application")  # 初始化 COM，不用实际打开

        for f in files:
            full_input = os.path.join(input_folder, f)
            full_output = os.path.join(output_folder, f.replace(".xlsx", "_v1.xlsx"))

            st.write(f"📄 正在处理：{f}")
            try:
                process_excel_file(full_input, full_output)
                st.success(f"✅ 已导出至：{full_output}")
            except Exception as e:
                st.error(f"❌ 处理失败：{f}，原因：{str(e)}")

        st.balloons()
        st.success("🎉 所有文件处理完成！请前往指定输出文件夹查看结果。")


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
    "🔍 查询数值",
    "🔍 查询满足条件的公司",
    "📁 处理原始数据",
    "🧠 自定义指标（占位）"
])

if menu_option == "🔍 查询数值":
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

elif menu_option == "🔍 查询满足条件的公司":
    st.subheader("🔍 查询满足条件的公司以及具体时间")

    indicator = st.text_input("📌 请输入要查询的指标（如：总资产周转率）", value="总资产周转率")
    threshold = st.text_input("📈 请输入阈值（百分比请直接输入数字，如98表示98%）", value="73")
    sheet_index = st.number_input("📄 请输入Sheet索引（从1开始）", min_value=1, value=1, step=1)

    # ✅ 新增：比较类型 + 年份区间
    compare_label = st.selectbox("请选择比较类型", ["大于", "大于等于", "小于", "小于等于", "等于"], index=0)
    compare_ops = {
        "大于":       (lambda x, y: x > y),
        "大于等于":   (lambda x, y: x >= y),
        "小于":       (lambda x, y: x < y),
        "小于等于":   (lambda x, y: x <= y),
        "等于":       (lambda x, y: abs(x - y) < 1e-6),  # 带容差，避免浮点误差
    }
    c1, c2 = st.columns(2)
    start_year = c1.number_input("开始年份（含）", min_value=1990, max_value=2100, value=2020, step=1)
    end_year   = c2.number_input("结束年份（含）", min_value=1990, max_value=2100, value=2025, step=1)
    if start_year > end_year:
        st.warning("开始年份大于结束年份，已自动交换。")
        start_year, end_year = end_year, start_year

    mode = st.radio("📂 请选择数据来源方式：", ["从本地文件夹读取", "上传 Excel 文件"])

    # 存储文件对象（file path 或 uploaded file）
    file_objs = []

    if mode == "从本地文件夹读取":
        folder_path = st.text_input("📂 请输入本地文件夹路径（如 D:/data）")
        if folder_path and os.path.exists(folder_path):
            file_objs = [
                os.path.join(folder_path, f)
                for f in os.listdir(folder_path)
                if f.lower().endswith('.xlsx')
            ]
        elif folder_path:
            st.error("❌ 文件夹路径不存在，请确认输入。")

    elif mode == "上传 Excel 文件":
        uploaded_files = st.file_uploader("📤 上传多个 Excel 文件", type="xlsx", accept_multiple_files=True)
        if uploaded_files:
            file_objs = uploaded_files

    if file_objs and st.button("🔍 开始筛选"):
        # 阈值转换
        threshold_value = None
        try:
            threshold_value = float(threshold)
        except ValueError:
            st.error("❌ 阈值必须为数字，请重新输入。")

        # 小工具：提取列名中的 4 位年份
        def _extract_year(val):
            m = re.search(r"(19|20)\d{2}", str(val))
            return int(m.group(0)) if m else None

        if threshold_value is not None:
            for f in file_objs:
                file_label = f.name if hasattr(f, "name") else os.path.basename(f)

                try:
                    # 读取 Excel
                    st.markdown(f"### 📄 正在处理文件：`{file_label}`")
                    
                    xl = pd.ExcelFile(f)
                    sheet_names = xl.sheet_names

                    if sheet_index > len(sheet_names):
                        st.warning(f"⚠️ 文件 {file_label} 不存在第 {sheet_index} 个 sheet，总共有 {len(sheet_names)} 个 sheet。")
                        continue

                    df = xl.parse(sheet_name=sheet_names[sheet_index - 1], header=None)
                    # 时间行（第6行，从第二列开始）
                    time_row = df.iloc[5, 1:]

                    # —— 年份过滤掩码（只保留 [start_year, end_year] 的列）
                    years = pd.Series([_extract_year(x) for x in time_row], index=time_row.index)
                    mask_year = years.apply(lambda y: y is not None and (int(start_year) <= y <= int(end_year)))

                    # 查找指标行（模糊匹配：包含关系）
                    matched_rows = df.iloc[:, 0].astype(str).str.strip().str.contains(str(indicator), na=False)
                    if matched_rows.any():

                        idx = matched_rows[matched_rows].index[0]
                        row_data_raw = df.iloc[idx, 1:]  # 原始字符串（可能包含 %）

                        # 将值解析为数字 & 文本显示（百分比→百分数字面；0<x<1 视作百分比）
                        row_data_numeric = []
                        row_data_display = []

                        for val in row_data_raw:
                            try:
                                val_str = str(val).strip()

                                # 默认值
                                display_val = val_str
                                num = None

                                if '%' in val_str:
                                    num = float(val_str.replace('%', '').strip())
                                    display_val = f"{num:.2f}%"
                                else:
                                    num = float(val_str)
                                    if 0 < num < 1:
                                        # Excel 百分比小数 → 百分数
                                        num *= 100
                                        display_val = f"{num:.2f}%"
                                    else:
                                        display_val = f"{num:.2f}"

                                row_data_numeric.append(num)
                                row_data_display.append(display_val)

                            except:
                                row_data_numeric.append(None)
                                row_data_display.append("N/A")

                        # —— 先按年份掩码过滤
                        row_series_all = pd.Series(row_data_numeric, index=row_data_raw.index)
                        row_series = row_series_all[mask_year]
                        time_row_year = time_row[mask_year]
                        row_display_year = [d for d, keep in zip(row_data_display, mask_year.tolist()) if keep]

                        # —— 动态比较（> ≥ < ≤ =）
                        cmp = compare_ops[compare_label]
                        passed = row_series.apply(lambda v: cmp(v, threshold_value) if v is not None and pd.notna(v) else False)

                        if passed.any():
                            st.success(f"✅ 文件 {file_label}（年份 {int(start_year)}–{int(end_year)}）满足条件的时间如下：")
                            # 将命中的列逐个输出
                            di = 0
                            for col, is_ok in passed.items():
                                if is_ok:
                                    st.write(f"📌 时间：{time_row_year.loc[col]}，值：{row_display_year[di]}")
                                di += 1
                        else:
                            st.info(f"📄 {file_label} 在 {int(start_year)}–{int(end_year)} 年内未发现满足“{compare_label} {threshold_value}”的值")

                    else:
                        st.warning(f"⚠️ 文件 {file_label} 中未找到指标 '{indicator}'")

                except Exception as e:
                    st.error(f"❌ 处理文件 {file_label} 时出错：{str(e)}")



# 📁 功能 3：读取并处理原始数据
elif menu_option == "📁 处理原始数据":
    from processor.process_excel import process_excel_file

    process_raw_data()

# 🧠 功能 4：自定义指标（占位）
elif menu_option == "🧠 自定义指标（占位）":
    st.subheader("🧠 自定义指标（敬请期待）")
    st.info("如 ROE = 净利润 / 所有者权益，将在后续版本中开放。")


#streamlit run streamlit_app.py