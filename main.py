# 从 TB_Clean.py 文件中导入封装好的启动函数
from TB_Clean import start_cleaning_task
from OutputVAT_Check import process_vat_check
from Mid_Platform_Data_Processing import process_revenue_summary
from Combined_OutputVAT import process_vat_with_mid_platform
from finalize_file import finalize_and_beautify

def main():
    # 1. 定义输出文件名
    target_file = "2024年3月财务清洗总表.xlsx"

    # 2. 定义需要处理的文件列表
    tb_file = ["科目余额表20260306170334.xlsx"]
    gl_file = ["总账利润中心科目余额报表20260309172722.xlsx"]
    vat_rate_file = "适用税率表.xlsx"
    collection_file = "收款汇总表_20260310_1611.xlsx"
    carryover_file = "收入结转汇总表_20260310_1612.xlsx"
    provision_file = "物业收入计提表_20260310_1610.xlsx"
    mid_base_file = "中台基础物管科目.xlsx"

    # 3. 调用执行
    start_cleaning_task(target_file, tb_file, gl_file)
    process_vat_check(target_file, vat_rate_file)
    process_revenue_summary(target_file, collection_file, carryover_file, provision_file)
    process_vat_with_mid_platform(target_file, mid_base_file)
    finalize_and_beautify(target_file)


if __name__ == "__main__":
    main()