from __future__ import annotations

from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
TOTAL_TEMPLATE = ROOT / "assets" / "APP&WEB总数据.xlsx"
PERIODS = ("本周", "上周", "MTD")


def _safe_div(numerator: float, denominator: float) -> float:
    return float(numerator) / float(denominator) if denominator else 0.0


def _change(current: float, previous: float) -> float:
    return current / previous - 1 if previous else 0.0


def generate_total_report(web_summary, app_summary, fixed_costs, start, end):
    """生成APP与WEB合并总表；两端summary必须使用同一统计周期。"""
    if not TOTAL_TEMPLATE.exists():
        raise FileNotFoundError(f"找不到APP&WEB总表模板：{TOTAL_TEMPLATE}")

    combined = {}
    for period in PERIODS:
        web = web_summary[period]
        app = app_summary[period]
        fixed = float(fixed_costs.get(period, 0.0) or 0.0)
        cost_without_fixed = (
            float(web["web_cost"])
            + float(web["kol_cost"])
            + float(app["total_cost"])
        )
        cost_with_fixed = cost_without_fixed + fixed
        new_revenue = float(web["new_user_revenue"]) + float(app["new_user_revenue"])
        non_subscription_revenue = (
            float(web["all_user_non_subscription_revenue"])
            + float(app["all_user_non_subscription_revenue"])
        )
        all_revenue = float(web["all_user_revenue"]) + float(app["all_user_revenue"])
        combined[period] = {
            "新用户综合 ROI (D0)": _safe_div(new_revenue, cost_without_fixed),
            "新老综合 ROI （除订阅）": _safe_div(non_subscription_revenue, cost_with_fixed),
            "新老用户综合ROI（含订阅）": _safe_div(all_revenue, cost_with_fixed),
            "新用户总收入": new_revenue,
            "新老总收入（除订阅）": non_subscription_revenue,
            "新老用户收入（含订阅）": all_revenue,
            "总成本（含固定成本）": cost_with_fixed,
            "总成本（新用户不含固定成本）": cost_without_fixed,
        }

    workbook = load_workbook(TOTAL_TEMPLATE)
    sheet = workbook.active
    row_keys = {
        4: "新用户综合 ROI (D0)",
        5: "新老综合 ROI （除订阅）",
        6: "新老用户综合ROI（含订阅）",
        7: "新用户总收入",
        8: "新老总收入（除订阅）",
        9: "新老用户收入（含订阅）",
        10: "总成本（含固定成本）",
        11: "总成本（新用户不含固定成本）",
    }
    for row, key in row_keys.items():
        current = combined["本周"][key]
        previous = combined["上周"][key]
        mtd = combined["MTD"][key]
        for column, value in zip((2, 3, 4, 5), (current, previous, _change(current, previous), mtd)):
            sheet.cell(row, column, value)

    for row in range(4, 7):
        for column in range(2, 6):
            sheet.cell(row, column).number_format = "0%"
    for row in range(7, 12):
        for column in (2, 3, 5):
            sheet.cell(row, column).number_format = '"$"#,##0'
        sheet.cell(row, 4).number_format = "0%"
    if sheet.max_column >= 6:
        sheet.delete_cols(6, sheet.max_column - 5)
    sheet.title = "APP&WEB总数据"
    sheet.print_area = "A2:E11"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True

    output = BytesIO()
    workbook.save(output)
    filename = f"APP&WEB总数据_{start:%Y.%m.%d}-{end:%Y.%m.%d}.xlsx"
    return filename, output.getvalue()
