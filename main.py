import argparse

from Combined_OutputVAT import process_vat_with_mid_platform
from Mid_Platform_Data_Processing import process_revenue_summary
from OutputVAT_Check import process_vat_check
from TB_Clean import start_cleaning_task
from finalize_file import finalize_and_beautify


def build_parser():
    parser = argparse.ArgumentParser(description='增值税申报核对流程')
    parser.add_argument('--target-file', default='2024年3月财务清洗总表.xlsx')
    parser.add_argument('--tb-file', nargs='+', default=['科目余额表20260306170334.xlsx'])
    parser.add_argument('--gl-file', nargs='+', default=['总账利润中心科目余额报表20260309172722.xlsx'])
    parser.add_argument('--vat-rate-file', default='适用税率表.xlsx')
    parser.add_argument('--collection-file', default='收款汇总表_20260310_1611.xlsx')
    parser.add_argument('--carryover-file', default='收入结转汇总表_20260310_1612.xlsx')
    parser.add_argument('--provision-file', default='物业收入计提表_20260310_1610.xlsx')
    parser.add_argument('--mid-base-file', default='中台基础物管科目.xlsx')
    parser.add_argument('--water-simple', action='store_true', help='自来水费按3%简易计税')
    parser.add_argument('--non-interactive', action='store_true', help='关闭 Tk 窗口交互')
    return parser


def main():
    args = build_parser().parse_args()

    start_cleaning_task(args.target_file, args.tb_file, args.gl_file)
    process_vat_check(
        args.target_file,
        args.vat_rate_file,
        water_simple=args.water_simple,
        interactive=not args.non_interactive,
    )
    process_revenue_summary(args.target_file, args.collection_file, args.carryover_file, args.provision_file)
    process_vat_with_mid_platform(args.target_file, args.mid_base_file)
    finalize_and_beautify(args.target_file)


if __name__ == '__main__':
    main()
