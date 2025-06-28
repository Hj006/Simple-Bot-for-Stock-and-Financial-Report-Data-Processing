import pandas as pd
import numpy as np

def compute_cashflow_metrics(df_cashflow: pd.DataFrame, df_profit: pd.DataFrame, df_main: pd.DataFrame) -> dict:
    results = {}
    n_cols = df_cashflow.shape[1]

    def col(df, i, row):  # 当前列
        return df.iloc[row, i] if i < df.shape[1] else np.nan

    def prev(df, i, row):
        try:
            return df.iloc[row, i + 1]
        except (IndexError, KeyError):
            return np.nan

    def prev2(df, i, row):
        try:
            return df.iloc[row, i + 2]
        except (IndexError, KeyError):
            return np.nan

    def prev4(df, i, row):
        try:
            return df.iloc[row, i + 4]
        except (IndexError, KeyError):
            return np.nan

    def prev5(df, i, row):
        try:
            return df.iloc[row, i + 5]
        except (IndexError, KeyError):
            return np.nan

    for i in range(n_cols):
        flag = df_cashflow.iloc[0, i]

        # 常规指标
        try:
            depreciation = sum([col(df_cashflow, i, r) for r in range(87, 92)])
            val = col(df_cashflow, i, 40) / depreciation if depreciation else np.nan
            results.setdefault('经营活动现金净流量/折旧摊销', []).append(val)
        except Exception:
            results.setdefault('经营活动现金净流量/折旧摊销', []).append(np.nan)

        try:
            results.setdefault('经营现金流净额/净利润', []).append(
                col(df_cashflow, i, 40) / col(df_profit, i, 54)
            )
        except Exception:
            results.setdefault('经营现金流净额/净利润', []).append(np.nan)

        try:
            results.setdefault('销售商品、提供劳务收到的现金/营业收入', []).append(
                col(df_cashflow, i, 9) / col(df_profit, i, 8)
            )
        except Exception:
            results.setdefault('销售商品、提供劳务收到的现金/营业收入', []).append(np.nan)

        try:
            results.setdefault('购建固定资产、无形资产和其他长期资产所支付的现金/经营现金流净额', []).append(
                col(df_cashflow, i, 50) / col(df_cashflow, i, 40)
            )
        except Exception:
            results.setdefault('购建固定资产、无形资产和其他长期资产所支付的现金/经营现金流净额', []).append(np.nan)

        try:
            results.setdefault('购建固定资产、无形资产和其他长期资产所支付的现金/收入', []).append(
                col(df_cashflow, i, 50) / col(df_profit, i, 8)
            )
        except Exception:
            results.setdefault('购建固定资产、无形资产和其他长期资产所支付的现金/收入', []).append(np.nan)

        try:
            invest_dividend_sum = col(df_cashflow, i, 56) + col(df_cashflow, i, 69)
            results.setdefault('现金余额/(投资支出+现金分红)', []).append(
                col(df_main, i, 9) / invest_dividend_sum if invest_dividend_sum else np.nan
            )
        except Exception:
            results.setdefault('现金余额/(投资支出+现金分红)', []).append(np.nan)

        try:
            delta_ar = (
                (col(df_main, i, 13) + col(df_main, i, 14) + col(df_main, i, 15)) -
                (col(df_main, i + 4, 13) + col(df_main, i + 4, 14) + col(df_main, i + 4, 15))
            )
            result_val = (col(df_cashflow, i, 9) + delta_ar) / col(df_profit, i, 8)
            results.setdefault('销售商品、提供劳务收到的现金+应收账款，应收票据增加额/营业收入', []).append(result_val)
        except Exception:
            results.setdefault('销售商品、提供劳务收到的现金+应收账款，应收票据增加额/营业收入', []).append(np.nan)

        # 同比
        try:
            results.setdefault('销售商品、提供劳务收到的现金（同比）', []).append(
                col(df_cashflow, i, 9) / prev4(df_cashflow, i, 9) - 1
                if prev4(df_cashflow, i, 9) else np.nan
            )
        except Exception:
            results.setdefault('销售商品、提供劳务收到的现金（同比）', []).append(np.nan)

        try:
            results.setdefault('经营活动产生的现金流量净额（同比）', []).append(
                col(df_cashflow, i, 40) / prev4(df_cashflow, i, 40) - 1
                if prev4(df_cashflow, i, 40) else np.nan
            )
        except Exception:
            results.setdefault('经营活动产生的现金流量净额（同比）', []).append(np.nan)

        # 环比
        try:
            c = col(df_cashflow, i, 9)
            p1 = prev(df_cashflow, i, 9)
            p2 = prev2(df_cashflow, i, 9)
            if flag == 1:
                results.setdefault('销售商品、提供劳务收到的现金（环比）', []).append(
                    c / (p1 - p2) - 1 if (p1 - p2) else np.nan
                )
            elif flag == 2:
                results.setdefault('销售商品、提供劳务收到的现金（环比）', []).append(
                    (c - p1) / p1 if p1 else np.nan
                )
            else:
                results.setdefault('销售商品、提供劳务收到的现金（环比）', []).append(
                    (c - p1) / (p1 - p2) - 1 if (p1 - p2) else np.nan
                )
        except Exception:
            results.setdefault('销售商品、提供劳务收到的现金（环比）', []).append(np.nan)

        try:
            c = col(df_cashflow, i, 40)
            p1 = prev(df_cashflow, i, 40)
            p2 = prev2(df_cashflow, i, 40)
            if flag == 1:
                results.setdefault('经营活动产生的现金流量净额（环比）', []).append(
                    c / (p1 - p2) - 1 if (p1 - p2) else np.nan
                )
            elif flag == 2:
                results.setdefault('经营活动产生的现金流量净额（环比）', []).append(
                    (c - p1) / p1 if p1 else np.nan
                )
            else:
                results.setdefault('经营活动产生的现金流量净额（环比）', []).append(
                    (c - p1) / (p1 - p2) - 1 if (p1 - p2) else np.nan
                )
        except Exception:
            results.setdefault('经营活动产生的现金流量净额（环比）', []).append(np.nan)

        # 每季度同比
        try:
            c = col(df_cashflow, i, 9)
            f = prev4(df_cashflow, i, 9)
            c2 = prev(df_cashflow, i, 9)
            f2 = prev4(df_cashflow, i, 9)
            g2 = prev5(df_cashflow, i, 9)
            if flag == 1:
                val = c / f - 1 if f else np.nan
            else:
                base = f2 - g2
                val = (c - c2) / base - 1 if base else np.nan
            results.setdefault('每季度销售商品、提供劳务收到的现金（同比)', []).append(val)
        except Exception:
            results.setdefault('每季度销售商品、提供劳务收到的现金（同比)', []).append(np.nan)

        try:
            c = col(df_cashflow, i, 40)
            f = prev4(df_cashflow, i, 40)
            c2 = prev(df_cashflow, i, 40)
            f2 = prev4(df_cashflow, i, 40)
            g2 = prev5(df_cashflow, i, 40)
            if flag == 1:
                val = c / f - 1 if f else np.nan
            else:
                base = f2 - g2
                val = (c - c2) / base - 1 if base else np.nan
            results.setdefault('每季度经营活动产生的现金流量净额（同比）', []).append(val)
        except Exception:
            results.setdefault('每季度经营活动产生的现金流量净额（同比）', []).append(np.nan)

    return {k: pd.Series(v) for k, v in results.items()}
