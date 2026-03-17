import pandas as pd
import warnings
import tkinter as tk
from tkinter import messagebox

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def process_vat_check(total_file_path, tax_rate_path, water_simple=None, interactive=True):
    root = None
    if interactive:
        root = tk.Tk()
        root.withdraw()

    print("正在读取数据并检查老项目名单...")
    df_profit = pd.read_excel(total_file_path, sheet_name='利润中心余额表')
    df_tax_base = pd.read_excel(tax_rate_path, sheet_name='VAT rate')
    df_tax_base.columns = ['科目代码', '科目名称', '分类', '基础税率', '标注']

    # 统一转换格式
    df_tax_base['科目代码'] = df_tax_base['科目代码'].astype(str).str.strip()
    df_profit['会计科目'] = df_profit['会计科目'].astype(str).str.strip()

    # --- 1. 不动产老项目前置检查 ---
    old_proj_mask = df_tax_base['标注'].astype(str).str.contains('老项目租赁', na=False)
    old_subjects = df_tax_base.loc[old_proj_mask, '科目代码'].tolist()

    lease_income_pcs = df_profit[df_profit['会计科目'].isin(old_subjects)][
        ['利润中心编码', '利润中心名称']].drop_duplicates()

    try:
        df_old_list = pd.read_excel(total_file_path, sheet_name='不动产租赁')
    except Exception:
        df_old_list = pd.DataFrame(columns=['利润中心代码', '利润中心名称', '是否为老项目'])

    check_m = pd.merge(lease_income_pcs, df_old_list,
                       left_on=['利润中心编码', '利润中心名称'],
                       right_on=['利润中心代码', '利润中心名称'], how='left')
    missing = check_m[check_m['是否为老项目'].isna()]

    if not missing.empty:
        new_items = missing[['利润中心编码', '利润中心名称', '是否为老项目']].rename(
            columns={'利润中心编码': '利润中心代码'})
        df_old_list = pd.concat([df_old_list, new_items], ignore_index=True)
        with pd.ExcelWriter(total_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_old_list.to_excel(writer, sheet_name='不动产租赁', index=False)
        msg = "已补充新的租赁利润中心，请在‘不动产租赁’Sheet标注后重新运行。"
        if interactive:
            messagebox.showwarning("停机：补充名单", msg)
            root.destroy()
            return
        raise ValueError(msg)

    # --- 2. 水费简易计税确认 ---
    if water_simple is None:
        if interactive:
            is_water_simple = messagebox.askyesno("税率确认", "自来水费是否适用简易计税（3%）？")
        else:
            is_water_simple = False
            print("[提示] 非交互模式下未指定 water_simple，默认按‘否’处理。")
    else:
        is_water_simple = bool(water_simple)

    # --- 3. 核心计算逻辑 ---
    # 以税率表为基准匹配数据
    df_calc = pd.merge(df_tax_base, df_profit, left_on='科目代码', right_on='会计科目', how='inner')
    df_calc = pd.merge(df_calc, df_old_list[['利润中心代码', '是否为老项目']], left_on='利润中心编码',
                       right_on='利润中心代码', how='left')

    def calculate_tax(row):
        # A. 确定适用税率
        rate = row['基础税率']
        tag = str(row['标注'])
        is_old = str(row['是否为老项目'])

        final_rate = rate
        if '自来水费' in tag and is_water_simple:
            final_rate = 0.03
        elif '老项目租赁' in tag and is_old == '是':
            final_rate = 0.05

        # B. 差异化计算公式
        code = str(row['科目代码'])
        credit_val = pd.to_numeric(row['本期_贷方'], errors='coerce') or 0

        # 2203开头的科目：含税价还原 (贷方 / (1+r) * r)
        if code.startswith('2203'):
            return (credit_val / (1 + final_rate)) * final_rate, final_rate
        # 6001开头的科目：不含税价直乘 (贷方 * r)
        elif code.startswith('6001'):
            return credit_val * final_rate, final_rate
        else:
            return 0, final_rate

    # 应用逻辑并拆分结果
    df_calc[['测算销项税额', '适用税率']] = df_calc.apply(
        lambda x: pd.Series(calculate_tax(x)), axis=1
    )

    # --- 4. 整理并导出 (优化标注显示逻辑) ---
    df_result = df_calc[[
        '利润中心编码', '利润中心名称', '科目代码', '科目名称_x', '分类',
        '适用税率', '是否为老项目', '标注', '本期_贷方', '测算销项税额'
    ]].copy()

    # 优化“是否不动产老项目”：仅当标注包含“老项目租赁”时显示，否则设为空
    df_result['是否不动产老项目'] = df_result.apply(
        lambda x: x['是否为老项目'] if '老项目租赁' in str(x['标注']) else '', axis=1
    )

    # 优化“水费是否简易”：仅当标注包含“自来水费”时显示，否则设为空
    df_result['水费是否简易'] = df_result.apply(
        lambda x: ('是' if is_water_simple else '否') if '自来水费' in str(x['标注']) else '', axis=1
    )

    # 移除临时的中间列
    df_result.drop(columns=['是否为老项目', '标注'], inplace=True)
    df_result.rename(columns={'科目名称_x': '科目名称'}, inplace=True)

    # --- 5. 写回总表 ---
    with pd.ExcelWriter(total_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_result.to_excel(writer, sheet_name='销项税额测算', index=False)

    if interactive and root is not None:
        root.destroy()


if __name__ == "__main__":
    process_vat_check("2024年3月财务清洗总表.xlsx", "适用税率表.xlsx")
