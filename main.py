import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def load_and_process_data():
    from processor.process_excel import process_excel_file

    print("\n📁 请选择包含财务数据文件的文件夹...")

    try:
        root = tk.Tk()
        root.withdraw()
        folder_path = filedialog.askdirectory(title="选择数据集文件夹")
        if not folder_path:
            print("未选择任何文件夹，操作已取消。")
            return

        print(f"你选择的文件夹是：{folder_path}")

        def contains_non_ascii(s):
            return not all(ord(c) < 128 for c in s)

        if contains_non_ascii(folder_path):
            print("选中的路径包含中文或特殊字符，建议使用英文路径。")
            return

        files = [f for f in os.listdir(folder_path) if f.lower().endswith('.xlsx')]

        if not files:
            print("该文件夹中未找到 .xlsx 文件。")
            return

        print(f"\n找到以下文件：")
        for f in files:
            print("   -", f)

        for f in files:
            full_path = os.path.join(folder_path, f)
            output_path = os.path.join(folder_path, f.replace('.xlsx', '_v1.xlsx'))
            print(f"\n正在处理: {f}")
            try:
                process_excel_file(full_path, output_path)
            except Exception as e:
                print(f"处理失败: {f}，原因：{e}")

        print("\n✅ 所有处理完成。")

    except Exception as e:
        print("出现错误：", str(e))


def query_data():
    import os
    import pandas as pd
    import tkinter as tk
    from tkinter import filedialog

    def extract_company_name(file_path):
        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet, header=None)
                name = str(df.iloc[2, 1]).strip()
                if name:
                    return name
        except:
            pass
        return None

    def query_from_file(file_path, company_name, indicator, target_date):
        xls = pd.ExcelFile(file_path)

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            try:
                short_name = str(df.iloc[2, 1]).strip()
            except:
                continue

            if short_name != company_name:
                continue

            date_row = df.iloc[5, 1:]
            col_index = None
            for i, val in enumerate(date_row):
                try:
                    if pd.to_datetime(val).strftime("%Y-%m-%d") == target_date:
                        col_index = i + 1
                        break
                except:
                    continue

            if col_index is None:
                return f"❌ 日期 {target_date} 未找到"

            for row in range(8, df.shape[0]):
                if str(df.iloc[row, 0]).strip() == indicator:
                    return df.iloc[row, col_index]

            return f"❌ 指标 “{indicator}” 未找到"

        return f"❌ 公司 {company_name} 未找到匹配表格"

    try:
        print("\n📁 请选择包含多个财报文件的文件夹...")
        root = tk.Tk()
        root.withdraw()
        folder_path = filedialog.askdirectory(title="选择数据文件夹")

        if not folder_path:
            print("未选择任何文件夹，操作取消。")
            return

        files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
                 if f.lower().endswith(".xlsx")]

        if not files:
            print("❌ 文件夹中未找到 Excel 文件")
            return

        company_map = {}
        for file in files:
            name = extract_company_name(file)
            if name and name not in company_map:
                company_map[name] = file

        if not company_map:
            print("❌ 未提取到任何公司简称")
            return

        company_list = sorted(company_map.keys())
        print("\n📌 可用公司简称：")
        print(" | ".join(company_list))
        user_input = input("请输入公司简称（可输入关键词，例如“三诺”）: ").strip()

        matches = [name for name in company_list if user_input in name]

        if not matches:
            print(f"❌ 未找到包含 “{user_input}” 的公司简称")
            return
        elif len(matches) == 1:
            company_name = matches[0]
            print(f"✅ 匹配到公司：{company_name}")
        else:
            print("⚠️ 匹配到多个公司，请选择：")
            for i, name in enumerate(matches):
                print(f"{i+1}. {name}")
            idx = input("请输入公司编号: ").strip()
            if not idx.isdigit() or int(idx) < 1 or int(idx) > len(matches):
                print("❌ 输入无效")
                return
            company_name = matches[int(idx) - 1]

        indicator = input("请输入指标名称（如：资产总计）: ").strip()
        date_str = input("请输入查询日期（格式：YYYY-MM-DD）: ").strip()

        file_path = company_map[company_name]
        result = query_from_file(file_path, company_name, indicator, date_str)

        print(f"\n🔍 查询结果：{company_name} - {indicator} @ {date_str} = {result}")

    except Exception as e:
        print("❌ 查询失败：", str(e))


def filter_by_indicators():
    pass


def define_custom_indicator():
    pass


def menu():
    while True:
        print("\n📊 Simple Financial Data Bot - 主菜单")
        print("1. 读取原始数据并计算保存指标")
        print("2. 设置筛选条件，输出满足条件的公司及财报时间")
        print("3. 查询某个公司或指标的值")
        print("4. 自定义新指标")
        print("5. 退出程序")

        choice = input("请输入选项 (1-5): ").strip()

        if choice == "1":
            load_and_process_data()
        elif choice == "2":
            filter_by_indicators()
        elif choice == "3":
            query_data()
        elif choice == "4":
            define_custom_indicator()
        elif choice == "5":
            print("👋 程序已退出。")
            break
        else:
            print("无效输入，请输入 1-5 之间的数字。")


if __name__ == "__main__":
    menu()
