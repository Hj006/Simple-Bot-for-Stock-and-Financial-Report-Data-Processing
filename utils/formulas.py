#\utils\formulas.py
import pandas as pd
import numpy as np
from utils.profit_logic import compute_profit_metrics
from utils.cashflow_logic import compute_cashflow_metrics

# === YoY / QoQ ===
def yoy_series(row: pd.Series, offset=4):
    result = []
    for i in range(len(row)):
        if i + offset < len(row):
            denom = row[i + offset]
            result.append(row[i] / denom - 1 if denom else np.nan)
        else:
            result.append(np.nan)
    return pd.Series(result)

def qoq_series(row: pd.Series, offset=1):
    result = []
    for i in range(len(row)):
        if i + offset < len(row):
            denom = row[i + offset]
            result.append(row[i] / denom - 1 if denom else np.nan)
        else:
            result.append(np.nan)
    return pd.Series(result)

# === 资产负债表部分 ===
def calc_receivable_change_ratio_multicol(df_main, df_profit):
    result = []
    for col in range(df_main.shape[1] - 1):
        curr_sum = df_main.iloc[13:16, col].sum()
        next_sum = df_main.iloc[13:16, col + 1].sum()
        delta = curr_sum - next_sum

        current_flag = df_main.iloc[0, col]
        income_curr = df_profit.iloc[9, col]
        income_next = df_profit.iloc[9, col + 1]

        base = income_curr if current_flag == 1 else income_curr - income_next
        result.append(delta / base if base else np.nan)
    result.append(np.nan)
    return pd.Series(result)

def calc_roe(df, df_profit):
    result = []
    for i in range(df.shape[1]):
        flag = df.iloc[0, i]
        net_profit = df_profit.iloc[58, i]
        equity = df.iloc[140, i]
        if not equity: result.append(np.nan); continue
        multiplier = 4 / flag if flag in [1, 2, 3, 4] else np.nan
        result.append((net_profit * multiplier) / equity if multiplier else np.nan)
    return pd.Series(result)

def calc_net_margin(df_profit):
    return df_profit.iloc[54, :] / df_profit.iloc[8, :]

def calc_asset_turnover(df, df_profit):
    result = []
    for i in range(df.shape[1]):
        flag = df.iloc[0, i]
        revenue = df_profit.iloc[8, i]
        multiplier = 4 / flag if flag in [1, 2, 3, 4] else np.nan
        revenue *= multiplier if multiplier else 1
        try:
            avg_asset = (df.iloc[71, i] + df.iloc[71, 4]) / 2
        except:
            avg_asset = np.nan
        result.append(revenue / avg_asset if avg_asset else np.nan)
    return pd.Series(result)

def calc_leverage(df):
    result = []
    for i in range(df.shape[1]):
        try:
            avg_asset = (df.iloc[71, i] + df.iloc[71, 4]) / 2
        except:
            avg_asset = np.nan
        equity = df.iloc[140, i]
        result.append(avg_asset / equity if equity else np.nan)
    return pd.Series(result)

def calc_inventory_turnover(df, df_profit):
    result = []
    for i in range(df.shape[1]):
        flag = df.iloc[0, i]
        cogs = df_profit.iloc[15, i]
        multiplier = 4 / flag if flag in [1, 2, 3, 4] else np.nan
        cogs *= multiplier if multiplier else 1
        try:
            avg_inv = (df.iloc[22, i] + df.iloc[22, 4]) / 2
        except:
            avg_inv = np.nan
        result.append(cogs / avg_inv if avg_inv else np.nan)
    return pd.Series(result)

def calc_ar_turnover(df, df_profit):
    result = []
    for i in range(df.shape[1]):
        flag = df.iloc[0, i]
        revenue = df_profit.iloc[8, i]
        multiplier = 4 / flag if flag in [1, 2, 3, 4] else np.nan
        revenue *= multiplier if multiplier else 1
        try:
            avg_ar = (df.iloc[14, i] + df.iloc[14, 4]) / 2
        except:
            avg_ar = np.nan
        result.append(revenue / avg_ar if avg_ar else np.nan)
    return pd.Series(result)


# === 统一封装函数（供主处理器调用） ===
def compute_all_metrics(df_main, df_profit, df_cashflow) -> dict:
    """
    返回合并后的所有指标计算结果 dict: {指标名 -> pd.Series}
    """
    # 主体计算函数映射
    calculated = {
        'ROE(pro forma)': calc_roe(df_main, df_profit),
        '净利率': calc_net_margin(df_profit),
        '总资产周转率': calc_asset_turnover(df_main, df_profit),
        '杠杆': calc_leverage(df_main),
        '存货周转率': calc_inventory_turnover(df_main, df_profit),
        '应收账款周转率': calc_ar_turnover(df_main, df_profit),
        '应收变化/收入变化': calc_receivable_change_ratio_multicol(df_main, df_profit),
        # 其他指标（例如同比/环比）建议在 profit_metrics 中处理
    }

    # 合并利润表指标
    calculated.update(compute_profit_metrics(df_profit, df_main))
    calculated.update(compute_cashflow_metrics(df_cashflow, df_profit, df_main))

    return calculated
