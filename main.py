
import os
import tkinter as tk
from tkinter import filedialog
import re


def load_and_process_data():
    print("\n📁 请选择包含财务数据文件的文件夹...")

    try:
        # 初始化 tkinter
        root = tk.Tk()
        root.withdraw()

        folder_path = filedialog.askdirectory(title="选择数据集文件夹")
        
        if not folder_path:
            print("⚠️ 未选择任何文件夹，操作已取消。")
            return

        print(f"📂 你选择的文件夹是：{folder_path}")
        
        def contains_non_ascii(s):
            return not all(ord(c) < 128 for c in s)

        if contains_non_ascii(folder_path):
            print("❌ 选中的路径包含中文或特殊字符，建议使用英文路径。")
            return

        # 遍历并拼接绝对路径
        files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('.csv', '.xlsx'))]

        if not files:
            print("❌ 该文件夹中未找到 .csv 或 .xlsx 文件。")
            return

        print(f"✅ 找到以下数据文件：")
        for f in files:
            print("   -", f)

        # 🔧 这里可以尝试读取第一个文件做测试
        import pandas as pd
        print("\n📖 尝试读取第一个文件进行预览...")
        df = pd.read_excel(files[0]) if files[0].endswith('.xlsx') else pd.read_csv(files[0])
        print(df.head())

    except Exception as e:
        print("❗ 出现错误：", str(e))


def filter_by_indicators():
    pass
def query_data():
    pass
def define_custom_indicator():
    pass

def menu():
    while True:
        print("\n Simple Financial Data Bot - 主菜单")
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
