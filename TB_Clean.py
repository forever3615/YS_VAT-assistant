import pandas as pd
import os
import warnings

# 1. 基础配置：忽略 openpyxl 带来的样式警告
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def clean_balance_sheet(file_path, output_name):
    """清洗普通科目余额表 (13列)"""
    df_raw = pd.read_excel(file_path, header=None)
    anchor_index = None
    for i in range(min(10, len(df_raw))):
        row_str = "".join(df_raw.iloc[i].fillna('').astype(str))
        if "科目代码" in row_str:
            anchor_index = i
            break

    if anchor_index is None: return

    # 判断表头深度
    next_row_str = "".join(df_raw.iloc[anchor_index + 1].fillna('').astype(str))
    start_row = anchor_index + 2 if ("借方" in next_row_str or "贷方" in next_row_str) else anchor_index + 1

    df_final = df_raw.iloc[start_row:].copy()
    columns = ['科目代码', '科目名称', '公司', '年初_借方', '年初_贷方', '期初_借方', '期初_贷方',
               '本期_借方', '本期_贷方', '本年_借方', '本年_贷方', '期末_借方', '期末_贷方']
    df_final = df_final.iloc[:, :13]
    df_final.columns = columns
    df_final = df_final.dropna(subset=['科目代码'])
    df_final = df_final[df_final['科目代码'].astype(str).str.contains(r'\d', na=False)]

    save_to_excel(df_final, output_name, "科目余额表")


def clean_profit_center_report(file_path, output_name):
    """清洗总账利润中心报表 (12列)"""
    df_raw = pd.read_excel(file_path, header=None)
    anchor_index = None
    for i in range(min(10, len(df_raw))):
        row_str = "".join(df_raw.iloc[i].fillna('').astype(str))
        if "会计科目" in row_str or "利润中心编码" in row_str:
            anchor_index = i
            break

    if anchor_index is None: return

    # 判断表头深度
    next_row_str = "".join(df_raw.iloc[anchor_index + 1].fillna('').astype(str))
    start_row = anchor_index + 2 if ("借方" in next_row_str or "贷方" in next_row_str) else anchor_index + 1

    df_final = df_raw.iloc[start_row:].copy()
    columns = ['会计科目', '科目名称', '法人组织编码', '法人组织名称', '利润中心编码', '利润中心名称',
               '期初_借方', '期初_贷方', '本期_借方', '本期_贷方', '期末_借方', '期末_贷方']
    df_final = df_final.iloc[:, :12]
    df_final.columns = columns

    # 核心需求：删除利润中心编码为空的行
    df_final = df_final.dropna(subset=['利润中心编码'])
    # 剔除包含“合计”字样的非明细行
    df_final = df_final[df_final['会计科目'].astype(str).str.contains(r'\d', na=False)]

    save_to_excel(df_final, output_name, "利润中心余额表")


def save_to_excel(df, output_name, sheet_name):
    """通用的保存逻辑：不存在则新建，存在则追加/替换 Sheet"""
    if not os.path.exists(output_name):
        df.to_excel(output_name, sheet_name=sheet_name, index=False)
        print(f">>> 已新建文件 {output_name} 并写入 Sheet: {sheet_name}")
    else:
        with pd.ExcelWriter(output_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f">>> 已在 {output_name} 中更新 Sheet: {sheet_name}")


# ==========================================
# 3. Main 函数调用部分
# ==========================================
# cleaner.py 内部结构

def start_cleaning_task(target_excel, balance_files, profit_files):
    """
    将之前的 main 逻辑封装进这个函数
    """
    print("开始清洗流程...")

    # 调用之前的清洗函数
    for f in balance_files:
        if os.path.exists(f):
            clean_balance_sheet(f, target_excel)

    for f in profit_files:
        if os.path.exists(f):
            clean_profit_center_report(f, target_excel)

    print("\n任务处理完毕！")


# 只有直接运行 cleaner.py 时才会执行这里
if __name__ == "__main__":
    # 这里可以放默认的测试逻辑
    default_output = "汇总结果.xlsx"
    start_cleaning_task(default_output, [], [])