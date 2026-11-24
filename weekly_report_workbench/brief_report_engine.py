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


def _first_column(frame: pd.DataFrame, candidates: tuple[str, ...]) -> str | None:
    return next((column for column in candidates if column in frame.columns), None)


def _normalize_google_campaign(frame: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "date": ("天", "日期", "Day", "Date"),
        "campaign": ("广告系列", "Campaign", "Campaign name"),
        "ad_group": ("广告单元", "广告组", "Ad group", "Ad group name"),
        "cost": ("费用", "花费", "Cost", "Spend"),
        "conversion_value": ("转化价值", "Conversions value", "Conversion value"),
    }
    result = frame.copy()
    for target, candidates in aliases.items():
        if target in result.columns:
            continue
        source = _first_column(result, candidates)
        if source:
            result = result.rename(columns={source: target})
    required = ["date", "campaign", "cost", "conversion_value"]
    missing = [column for column in required if column not in result.columns]
    if missing:
        raise ValueError(
            "Google前台ROI数据缺少字段：" + "、".join(missing)
            + "。API模式会自动读取；手动文件需包含广告系列、费用和转化价值。"
        )
    if "ad_group" not in result.columns:
        result["ad_group"] = ""
    result["date"] = pd.to_datetime(result["date"], errors="coerce")
    for column in ("cost", "conversion_value"):
        result[column] = pd.to_numeric(
            result[column].astype(str).str.replace(",", "", regex=False).str.replace(r"[^0-9.\-]", "", regex=True),
            errors="coerce",
        ).fillna(0.0)
    result["campaign"] = result["campaign"].fillna("").astype(str).str.strip()
    result["ad_group"] = result["ad_group"].fillna("").astype(str).str.strip()
    return result[result["date"].notna() & result["campaign"].ne("")].copy()


def _source_mask(frame: pd.DataFrame, source_field: str | None, ad_groups: set[str], campaign: str) -> pd.Series:
    if source_field is None:
        return pd.Series(False, index=frame.index)
    source = frame[source_field].fillna("").astype(str).str.strip().str.casefold()
    accepted = {value.casefold() for value in ad_groups if value}
    accepted.add(campaign.casefold())
    return source.isin(accepted)


def _campaign_backend_rows(reg, orders, google_campaign, start: date, end: date, new_only: bool):
    period_google = google_campaign[google_campaign["date"].dt.date.between(start, end)]
    grouped = period_google.groupby("campaign", sort=False)
    reg_source = _first_column(reg, ("广告来源", "末次广告来源", "广告单元", "末次广告单元"))
    order_source = _first_column(orders, ("广告来源", "末次广告来源", "广告单元", "末次广告单元"))
    # 订单表通常不直接带广告单元；此时按用户ID回连注册表的广告来源。
    if order_source is None and reg_source is not None:
        reg_uid = _first_column(reg, ("用户ID", "用户id", "uid", "UID"))
        order_uid = _first_column(orders, ("用户ID", "用户id", "uid", "UID"))
        if reg_uid and order_uid:
            reg_link = reg[[reg_uid, reg_source]].dropna(subset=[reg_uid]).copy()
            reg_link["__uid_key"] = reg_link[reg_uid].astype(str).str.strip()
            source_by_uid = reg_link.drop_duplicates("__uid_key", keep="last").set_index("__uid_key")[reg_source]
            orders = orders.copy()
            orders["__brief_ad_source"] = orders[order_uid].astype(str).str.strip().map(source_by_uid).fillna("")
            order_source = "__brief_ad_source"
    rows = []
    for campaign, campaign_google in grouped:
        ad_groups = set(campaign_google["ad_group"].dropna().astype(str).str.strip()) - {""}
        spend = float(campaign_google["cost"].sum())
        conversion_value = float(campaign_google["conversion_value"].sum())
        reg_mask = (
            reg["注册时间"].dt.date.between(start, end)
            & reg["__web"]
            & _source_mask(reg, reg_source, ad_groups, campaign)
        )
        order_mask = (
            orders["订单生成时间"].dt.date.between(start, end)
            & orders["__web"]
            & _source_mask(orders, order_source, ad_groups, campaign)
        )
        if "合作渠道" in orders.columns:
            order_mask &= orders["合作渠道"].fillna("").astype(str).str.strip().eq("")
        if new_only:
            order_mask &= orders["注册转化周期"].eq(0)
        else:
            if "支付位置" in orders.columns:
                order_mask &= ~orders["支付位置"].fillna("").astype(str).str.contains("订阅", na=False)
            if "末次来源类型" in orders.columns:
                order_mask &= ~orders["末次来源类型"].fillna("").astype(str).str.upper().eq("AUTO_RENEWAL")
        backend_orders = orders.loc[order_mask]
        revenue = float(backend_orders["订单金额"].sum())
        rows.append({
            "广告系列": campaign,
            "后台注册": int(reg_mask.sum()),
            "后台订单量": int(order_mask.sum()),
            "消耗": spend,
            "后台营收": revenue,
            "后台ROI": safe_div(revenue, spend),
            "Google前台ROI": safe_div(conversion_value, spend),
            "广告单元": "、".join(sorted(ad_groups)),
        })
    return rows, reg_source, order_source


def _write_front_back_roi_sheet(ws, reg, orders, google_campaign, start: date, end: date):
    google_campaign = _normalize_google_campaign(google_campaign)
    sections = (("除订阅", False), ("新用户", True))
    headers = ["广告系列", "后台注册", "后台订单量", "消耗", "后台营收", "后台ROI", "Google前台ROI", "广告单元"]
    source_fields = (None, None)
    for index, (label, new_only) in enumerate(sections):
        if index:
            ws.append([])
            ws.append([])
        rows, reg_source, order_source = _campaign_backend_rows(reg, orders, google_campaign, start, end, new_only)
        source_fields = (reg_source, order_source)
        ws.append([f"{start:%m.%d}-{end:%m.%d}({label})", *headers[1:]])
        header_row = ws.max_row
        for cell in ws[header_row]:
            cell.fill = PatternFill("solid", fgColor="FFC000" if cell.column == 1 else "F2F2F2")
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        for values in rows:
            ws.append([values[column] for column in headers])
            row = ws.max_row
            for col in (4, 5):
                ws.cell(row, col).number_format = '$#,##0.00'
            for col in (6, 7):
                ws.cell(row, col).number_format = "0%"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    widths = [42, 14, 14, 14, 18, 14, 18, 52]
    for column, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(column)].width = width
    return source_fields


def generate_monday_friday_brief(
    web_reg, web_orders, web_google, web_bing,
    app_google, app_appsflyer, web_google_campaign,
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
    reg_source, order_source = _write_front_back_roi_sheet(
        workbook.create_sheet("Google前后台ROI"),
        web_reg, web_orders, web_google_campaign, current_start, current_end,
    )
    if not reg_source:
        warnings.append("注册表缺少广告来源/广告单元字段，Google前后台ROI的后台注册将显示为0。")
    if not order_source:
        warnings.append("订单表缺少广告来源/广告单元字段，Google前后台ROI的后台订单和营收将显示为0。")
    output = BytesIO()
    workbook.save(output)
    return (
        f"周一周五数据简报_{current_start:%Y%m%d}-{current_end:%Y%m%d}.xlsx",
        output.getvalue(),
        warnings,
    )
