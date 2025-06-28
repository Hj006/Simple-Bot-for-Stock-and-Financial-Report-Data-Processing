# utils/profit_logic.py
import pandas as pd
import numpy as np

def compute_profit_metrics(df_profit: pd.DataFrame, df_main: pd.DataFrame) -> dict:
    results = {}
    n_cols = df_profit.shape[1]

    def col(i, row):
        return df_profit.iloc[row, i]
    
    def prev(i, row):
        try:
            if i + 1 >= df_profit.shape[1] or row >= df_profit.shape[0]:
                return np.nan
            return df_profit.iloc[row, i + 1]
        except:
            return np.nan

    def prev4(i, row):
        try:
            if i + 4 >= df_profit.shape[1] or row >= df_profit.shape[0]:
                return np.nan
            return df_profit.iloc[row, i + 4]
        except:
            return np.nan

    # 同理 prev2, prev5 也改成这样

    def prev2(i, row):
        try:
            if i + 2 >= df_profit.shape[1] or row >= df_profit.shape[0]:
                return np.nan
            return df_profit.iloc[row, i + 2]
        except:
            return np.nan
        
    def prev5(i, row):
        try:
            if i + 5 >= df_profit.shape[1] or row >= df_profit.shape[0]:
                return np.nan
            return df_profit.iloc[row, i + 5]
        except:
            return np.nan
    for i in range(n_cols):
        flag = df_profit.iloc[0, i]

        kf = col(i, 54) - col(i, 44) + col(i, 45) - col(i, 31)
        results.setdefault('扣非归母净利润', []).append(kf)

        ebit = col(i, 49) - col(i, 20)
        results.setdefault('已获利息倍数', []).append(ebit / col(i, 20) if col(i, 20) else np.nan)

        results.setdefault('毛利率', []).append(1 - col(i, 15) / col(i, 8) if col(i, 8) else np.nan)
        results.setdefault('销售费用/收入', []).append(col(i, 17) / col(i, 8) if col(i, 8) else np.nan)
        results.setdefault('管理费用/收入', []).append(col(i, 18) / col(i, 8) if col(i, 8) else np.nan)
        results.setdefault('研发费用/收入', []).append(col(i, 19) / col(i, 8) if col(i, 8) else np.nan)
        results.setdefault('财务费用/收入', []).append(col(i, 20) / col(i, 8) if col(i, 8) else np.nan)

        sf_sum = sum([col(i, r) for r in range(17, 21 + 1)])
        income = col(i, 8)
        results.setdefault('费用/收入', []).append(sf_sum / income if income else np.nan)

        sf = sum([col(i, r) for r in range(17, 21 + 1)])
        margin = col(i, 8) - col(i, 15)
        results.setdefault('三费/毛利润', []).append(sf / margin if margin else np.nan)

        results.setdefault('营业利润率', []).append(col(i, 43) / col(i, 8) if col(i, 8) else np.nan)
        results.setdefault('净利率', []).append(col(i, 54) / col(i, 8) if col(i, 8) else np.nan)

        def yoy(row):
            return (col(i, row) / prev4(i, row) - 1) if prev4(i, row) else np.nan

        yoy_rows = {
            '营业收入': 8, '营业总成本': 14, '销售费用': 17,
            '管理费用': 18, '研发费用': 19, '财务费用': 20,
            '资产减值损失': 37, '公允价值变动收益': 36,
            '投资收益': 32, '其中:对联营企业和合营企业的投资收益': 33,
            '营业利润': 43, '归母净利润': 58
        }

        for name, row in yoy_rows.items():
            results.setdefault(name, []).append(yoy(row))

        adjusted = col(i, 58) - col(i, 44) + col(i, 45) - col(i, 31) - col(i, 32)
        results.setdefault('归母扣非净利润 （估算，不含投资收益）', []).append(adjusted)

        try:
            curr = col(i, 58) - col(i, 44) + col(i, 45) - col(i, 31) - col(i, 32)
            base = prev(i, 58) - prev(i, 44) + prev(i, 45) - prev(i, 31) - prev(i, 32)
            value = (curr / base - 1) if base else np.nan
        except IndexError:
            value = np.nan
        results.setdefault('归母扣非净利润 %（估算，不含投资收益）同比', []).append(value)

        def qoq(val_curr, val_prev, val_prev2):
            if flag == 1:
                return val_curr / (val_prev - val_prev2) - 1 if val_prev - val_prev2 else np.nan
            elif flag == 2:
                return (val_curr - val_prev) / val_prev if val_prev else np.nan
            elif flag in (3, 4):
                return (val_curr - val_prev) / (val_prev - val_prev2) - 1 if (val_prev - val_prev2) else np.nan
            return np.nan

        results.setdefault('营业收入（环比）', []).append(qoq(col(i, 8), prev(i, 8), prev2(i, 8)))
        results.setdefault('归母净利润（环比）', []).append(qoq(col(i, 58), prev(i, 58), prev2(i, 58)))

        try:
            curr = col(i, 58) - col(i, 44) + col(i, 45) - col(i, 31) - col(i, 32)
            p1 = prev(i, 58) - prev(i, 44) + prev(i, 45) - prev(i, 31) - prev(i, 32)
            p2 = prev2(i, 58) - prev2(i, 44) + prev2(i, 45) - prev2(i, 31) - prev2(i, 32)

            if flag == 1:
                val = curr / (p1 - p2) - 1 if (p1 - p2) else np.nan
            elif flag == 2:
                val = (curr - p1) / p1 if p1 else np.nan
            elif flag in (3, 4):
                val = (curr - p1) / (p1 - p2) - 1 if (p1 - p2) else np.nan
            else:
                val = np.nan
        except IndexError:
            val = np.nan

        results.setdefault('归母扣非净利润 %（估算，不含投资收益）（环比）', []).append(val)

        if i >= df_main.shape[1]:
            results.setdefault('应收应付率', []).append(np.nan)
        else:
            ar_sum = df_main.iloc[13:16, i].sum()
            ap_sum = df_main.iloc[77:80, i].sum()
            results.setdefault('应收应付率', []).append((ar_sum - ap_sum) / ap_sum if ap_sum else np.nan)

        cp = col(i, 8) - sum([col(i, r) for r in range(15, 21)])
        results.setdefault('核心利润', []).append(cp)

        def quarter_yoy(row):
            if flag == 1:
                return (col(i, row) / prev4(i, row)) - 1 if prev4(i, row) else np.nan
            else:
                curr_delta = col(i, row) - prev(i, row)
                prev_delta = prev4(i, row) - prev5(i, row)
                return (curr_delta / prev_delta) - 1 if prev_delta else np.nan

        results.setdefault('每季度收入（同比）', []).append(quarter_yoy(8))
        results.setdefault('每季度利润（同比）', []).append(quarter_yoy(58))

        try:
            curr = col(i, 8) - sum([col(i, r) for r in range(15, 21)])
            p1 = prev(i, 8) - sum([prev(i, r) for r in range(15, 21)])
            p4 = prev4(i, 8) - sum([prev4(i, r) for r in range(15, 21)])
            p5 = prev2(i, 8) - sum([prev5(i, r) for r in range(15, 21)])

            if flag == 1:
                val = curr / p4 - 1 if p4 else np.nan
            else:
                base = p5
                val = (curr - p1) / (p4 - base) - 1 if (p4 - base) else np.nan
        except IndexError:
            val = np.nan

        results.setdefault('每季度核心利润（同比）', []).append(val)

        try:
            curr = col(i, 58) - col(i, 44) + col(i, 45) - col(i, 31) - col(i, 32)
            p1 = prev(i, 58) - prev(i, 44) + prev(i, 45) - prev(i, 31) - prev(i, 32)
            p5 = prev5(i, 58) - prev5(i, 44) + prev5(i, 45) - prev5(i, 31) - prev5(i, 32)
            p4 = prev4(i, 58) - prev4(i, 44) + prev4(i, 45) - prev4(i, 31) - prev4(i, 32)
            if flag == 1:
                val = curr / p4 - 1 if p1 else np.nan
            else:
                val = (curr - p1) / (p4 - p5) - 1 if (p4 - p5) else np.nan
        except IndexError:
            val = np.nan

        results.setdefault('每季度归母扣非净利润（同比）（估算，不含投资收益）', []).append(val)

        try:
            curr = col(i, 8) - sum([col(i, r) for r in range(15, 21)])
            p1 = prev(i, 8) - sum([prev(i, r) for r in range(15, 21)])
            p2 = prev2(i, 8) - sum([prev2(i, r) for r in range(15, 21)])

            if flag == 1:
                val = curr / (p1 - p2) - 1 if (p1 - p2) else np.nan
            elif flag == 2:
                val = (curr - p1) / p1 - 1 if p1 else np.nan
            else:
                val = (curr - p1) / (p1 - p2) - 1 if (p1 - p2) else np.nan

        except IndexError:
            val = np.nan

        results.setdefault('每季度核心利润（环比）', []).append(val)

    return {k: pd.Series(v) for k, v in results.items()}
