from __future__ import annotations

from datetime import date, timedelta
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from report_engine import (
    METRICS,
    TARGET_REGIONS,
    Period,
    aggregate_metrics,
    prepare_data,
    safe_div,
)
from app_report_engine import (
    APPSFLYER_REGION_COUNTRIES,
    REGIONS,
    _sum_appsflyer_revenue,
    _sum_cost,
    prepare_app_costs,
    prepare_appsflyer_revenue,
)


RATE_METRICS = {"注册付费率", "新用户注册付费率", "ROI", "新用户ROI"}
MONEY_METRICS = {"注册单价", "客单价", "营收", "新用户营收", "消耗"}


def _style_sheet(ws):
    dark, lime, pale, line = "111111", "E3F85A", "F2F4F7", "D9DCE1"
    thin = Side(style="thin", color=line)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "C2"
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=dark)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 1).fill = PatternFill("solid", fgColor=pale)
        for col in range(1, ws.max_column + 1):
            ws.cell(row, col).border = Border(bottom=thin)
        if row % 2 == 0:
            ws.cell(row, ws.max_column).fill = PatternFill("solid", fgColor=lime)
    widths = {1: 18, 2: 24, 3: 16, 4: 16, 5: 14}
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.auto_filter.ref = ws.dimensions


def _write_web_sheet(ws, current: pd.DataFrame, previous: pd.DataFrame):
    ws.append(["地区", "指标", "本期", "对比期", "环比"])
    for region in TARGET_REGIONS:
        for metric in METRICS:
            now = float(current.at[region, metric]) if region in current.index else 0.0
            before = float(previous.at[region, metric]) if region in previous.index else 0.0
            ws.append([region, metric, now, before, safe_div(now - before, before) if before else 0.0])
            row = ws.max_row
            value_format = "0.00%" if metric in RATE_METRICS else "$#,##0.00" if metric in MONEY_METRICS else "#,##0"
            ws.cell(row, 3).number_format = value_format
            ws.cell(row, 4).number_format = value_format
            ws.cell(row, 5).number_format = "0.00%"
    _style_sheet(ws)


def _app_region_values(google, appsflyer, start: date, end: date, region: str | None):
    country = APPSFLYER_REGION_COUNTRIES.get(region) if region else None
    cost = _sum_cost(google, start, end, region) if region else _sum_cost(google, start, end)
    revenue = _sum_appsflyer_revenue(
        appsflyer, start, end,
        media_sources={"googleadwords_int"},
        country=country,
    )
    return {"Google消耗": cost, "Google投放收入": revenue, "ROI": safe_div(revenue, cost)}


def _write_app_sheet(ws, google, appsflyer, current_range, previous_range):
    ws.append(["地区", "指标", "本期", "对比期", "环比"])
    for region in ("全部地区", *REGIONS):
        region_key = None if region == "全部地区" else region
        current = _app_region_values(google, appsflyer, *current_range, region_key)
        previous = _app_region_values(google, appsflyer, *previous_range, region_key)
        for metric in ("Google消耗", "Google投放收入", "ROI"):
            now, before = current[metric], previous[metric]
            ws.append([region, metric, now, before, safe_div(now - before, before) if before else 0.0])
            row = ws.max_row
            value_format = "0.00%" if metric == "ROI" else "$#,##0.00"
            ws.cell(row, 3).number_format = value_format
            ws.cell(row, 4).number_format = value_format
            ws.cell(row, 5).number_format = "0.00%"
    _style_sheet(ws)


def generate_monday_friday_brief(
    web_reg, web_orders, web_google, web_bing,
    app_google, app_appsflyer,
    current_start: date, current_end: date,
    previous_start: date, previous_end: date,
):
    if current_start > current_end or previous_start > previous_end:
        raise ValueError("开始日期不能晚于结束日期。")
    if (current_end - current_start).days != (previous_end - previous_start).days:
        raise ValueError("本期与对比期必须保持相同天数。")

    web_reg, web_orders, web_google, web_bing, warnings = prepare_data(
        web_reg, web_orders, web_google, web_bing
    )
    current = aggregate_metrics(
        web_reg, web_orders, web_google, web_bing,
        Period("本期", current_start, current_end), "SEM",
    )
    previous = aggregate_metrics(
        web_reg, web_orders, web_google, web_bing,
        Period("对比期", previous_start, previous_end), "SEM",
    )

    app_google, _, _, app_warnings = prepare_app_costs(
        app_google,
        pd.DataFrame(columns=["date", "region", "cost"]),
        pd.DataFrame(columns=["date", "region", "cost"]),
    )
    app_appsflyer, revenue_warnings = prepare_appsflyer_revenue(app_appsflyer)
    warnings.extend(app_warnings)
    warnings.extend(revenue_warnings)

    workbook = Workbook()
    workbook.remove(workbook.active)
    _write_web_sheet(workbook.create_sheet("Web SEM分地区"), current, previous)
    _write_app_sheet(
        workbook.create_sheet("App SEM"), app_google, app_appsflyer,
        (current_start, current_end), (previous_start, previous_end),
    )
    output = BytesIO()
    workbook.save(output)
    return (
        f"周一周五数据简报_{current_start:%Y%m%d}-{current_end:%Y%m%d}.xlsx",
        output.getvalue(),
        warnings,
    )
