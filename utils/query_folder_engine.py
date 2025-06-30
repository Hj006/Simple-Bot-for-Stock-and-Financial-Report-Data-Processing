import pandas as pd
import os


def extract_company_name(file_obj):
    """
    支持 Streamlit 上传的 BytesIO 对象或本地路径
    """
    try:
        xls = pd.ExcelFile(file_obj)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            name = str(df.iloc[2, 1]).strip()
            if name:
                return name
    except Exception as e:
        print(f"[DEBUG] 解析失败: {e}")
    return None



def query_from_file(file_obj, company_name, indicator, target_date):
    """
    在上传的 Excel 文件中查询公司简称 + 指标 + 日期 的交叉值。
    支持本地文件或 Streamlit 上传文件。
    """
    try:
        xls = pd.ExcelFile(file_obj)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

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

    except Exception as e:
        return f"❌ 查询出错：{e}"


    return f"❌ 公司 {company_name} 未找到匹配表格"
