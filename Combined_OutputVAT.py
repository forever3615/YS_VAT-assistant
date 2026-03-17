import difflib
import warnings
import tkinter as tk
from tkinter import messagebox

import pandas as pd

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def ask_confirmation(biz_name, fin_name):
    """弹出确认窗口，返回是否采用建议映射。"""
    root = tk.Tk()
    root.withdraw()
    result = messagebox.askyesno(
        "利润中心匹配确认",
        f"业务项目名：【{biz_name}】\n财务利润中心：【{fin_name}】\n\n相似度较高，是否视为同一个利润中心？"
    )
    root.destroy()
    return result


def _build_project_mapping(df_biz, fin_pc_names, cutoff=0.85, interactive=True):
    """构建项目名称到利润中心名称映射，并返回需人工确认清单。"""
    mapping = {}
    pending_rows = []

    for project_name in sorted(df_biz['项目名称'].dropna().astype(str).unique()):
        if project_name in fin_pc_names:
            mapping[project_name] = project_name
            continue

        matches = difflib.get_close_matches(project_name, fin_pc_names, n=1, cutoff=cutoff)
        if matches:
            suggested = matches[0]
            score = round(difflib.SequenceMatcher(None, project_name, suggested).ratio(), 4)
            use_suggestion = ask_confirmation(project_name, suggested) if interactive else True
            mapping[project_name] = suggested if use_suggestion else project_name
            pending_rows.append({
                '项目名称': project_name,
                '建议利润中心名称': suggested,
                '相似度': score,
                '是否采用建议': '是' if use_suggestion else '否'
            })
        else:
            mapping[project_name] = project_name
            pending_rows.append({
                '项目名称': project_name,
                '建议利润中心名称': project_name,
                '相似度': 0,
                '是否采用建议': '无匹配'
            })

    return mapping, pd.DataFrame(pending_rows)


def process_vat_with_mid_platform(target_file, mid_base_file, interactive=True):
    print("\n>>> 开始生成 [销项测算（含中台）] 并进行税金核对 ...")

    # 1. 读取基础资料
    df_mid_base = pd.read_excel(mid_base_file)
    df_biz_detail = pd.read_excel(target_file, sheet_name='中台明细汇总')
    df_vat_old = pd.read_excel(target_file, sheet_name='销项税额测算')
    df_profit = pd.read_excel(target_file, sheet_name='利润中心余额表')

    # 2. 准备财务利润中心清单用于自动匹配
    fin_pc_names = df_vat_old['利润中心名称'].dropna().astype(str).unique().tolist()

    # 3. 建立科目映射
    sub_mapping = df_mid_base.set_index('中台收费科目名称')['预收账款科目编码'].to_dict()

    # 4. 处理业务端名称对齐
    df_biz = df_biz_detail.copy()
    df_biz['科目代码'] = df_biz['收费科目'].map(sub_mapping)
    df_biz = df_biz.dropna(subset=['科目代码'])

    pc_name_map, df_mapping_check = _build_project_mapping(df_biz, fin_pc_names, interactive=interactive)
    df_biz['利润中心名称'] = df_biz['项目名称'].map(pc_name_map)

    # 5. 聚合与合并逻辑
    df_biz_sum = df_biz.groupby(['利润中心名称', '科目代码']).agg({'增值税申报销售额': 'sum'}).reset_index()
    df_vat_base = df_vat_old.groupby(['利润中心名称', '科目代码']).agg({
        '利润中心编码': 'first',
        '科目名称': 'first',
        '分类': 'first',
        '适用税率': 'first',
        '本期_贷方': 'sum'
    }).reset_index()

    # 6. 执行最终合并
    df_integrated = pd.merge(df_vat_base, df_biz_sum, on=['利润中心名称', '科目代码'], how='outer').fillna(0)

    # 7. 计算调整后数据
    df_integrated['调整后销售额'] = df_integrated['本期_贷方'] + df_integrated['增值税申报销售额']
    df_integrated['适用税率'] = df_integrated['适用税率'].apply(lambda x: x if x > 0 else 0.06)
    df_integrated['调整后测算销项税额'] = (df_integrated['调整后销售额'] * df_integrated['适用税率']).round(2)

    # 8. 核心核对逻辑
    tax_map_codes = {
        0.13: ['2221.13.04.03'],
        0.09: ['2221.13.04.11'],
        0.06: ['2221.13.04.10', '2221.13.04.12'],
        0.05: ['2221.14.01.01', '2221.14.01.02'],
        0.03: ['2221.14.01.03', '2221.14.01.04'],
        0.01: ['2221.14.01.05']
    }

    all_pcs = df_integrated['利润中心名称'].unique()
    check_rows = []

    for pc in all_pcs:
        for rate_val, codes in tax_map_codes.items():
            if rate_val in [0.13, 0.09, 0.06]:
                actual_tax = df_profit[(df_profit['利润中心名称'] == pc) &
                                       (df_profit['会计科目'].isin(codes))]['本期_贷方'].sum()
            else:
                actual_tax = df_profit[(df_profit['利润中心名称'] == pc) &
                                       (df_profit['会计科目'].isin(codes))]['本期_借方'].sum()

            calc_tax = df_integrated[(df_integrated['利润中心名称'] == pc) &
                                     (df_integrated['适用税率'] == rate_val)]['调整后测算销项税额'].sum()

            if abs(actual_tax) > 0.01 or abs(calc_tax) > 0.01:
                check_rows.append({
                    '利润中心': pc,
                    '税率': rate_val,
                    '账面税金': round(actual_tax, 2),
                    '调整后销项测算': round(calc_tax, 2),
                    '差异': round(actual_tax - calc_tax, 2)
                })

    df_check_sheet = pd.DataFrame(check_rows)

    # 9. 写入文件
    with pd.ExcelWriter(target_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_integrated.to_excel(writer, sheet_name='销项测算（含中台）', index=False)
        df_check_sheet.to_excel(writer, sheet_name='中台测算核对表', index=False)
        if not df_mapping_check.empty:
            df_mapping_check.to_excel(writer, sheet_name='利润中心映射待确认', index=False)

    if not df_mapping_check.empty:
        print(">>> [提示] 已生成 Sheet '利润中心映射待确认'，请复核后按需要调整项目命名或映射表。")
    print(">>> [成功] 已生成 '销项测算（含中台）' 与 '中台测算核对表'。")


if __name__ == "__main__":
    process_vat_with_mid_platform("2024年3月财务清洗总表.xlsx", "中台基础物管科目.xlsx")
