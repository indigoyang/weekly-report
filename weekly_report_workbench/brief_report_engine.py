from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
import re

import pandas as pd
from openpyxl import load_workbook

from report_engine import (
    Period,
    aggregate_metrics,
    prepare_data,
    safe_div,
)
from app_report_engine import (
    APPSFLYER_BUY_CHANNELS,
    APPSFLYER_NATURAL_CHANNELS,
    APPSFLYER_REGION_COUNTRIES,
    REGIONS,
    _sum_appsflyer_revenue,
    _sum_cost,
    prepare_app_costs,
    prepare_appsflyer_revenue,
)


ROOT = Path(__file__).resolve().parent
BRIEF_TEMPLATE = ROOT / "assets" / "数据监控格式.xlsx"


WEB_TEMPLATE_METRICS = {
    "新增注册": "新增注册",
    "注册单价": "注册单价",
    "新用户客单价": "__new_aov",
    "新老用户（除订阅）注册付费率": "注册付费率",
    "新用户注册付费率": "新用户注册付费率",
    "新老用户（除订阅）付费量": "付费量",
    "新老用户（除订阅）营收": "营收",
    "新用户订单量": "新用户订单量",
    "新用户营收": "新用户营收",
    "消耗": "消耗",
    "SEM新老用户ROI（除订阅）": "ROI",
    "SEM新用户ROI": "新用户ROI",
    "全渠道新用户ROI": "__all_channel_new_roi",
}


def _metric_value(frame: pd.DataFrame, all_channel: pd.DataFrame, region: str, key: str) -> float:
    if region not in frame.index:
        return 0.0
    if key == "__new_aov":
        return safe_div(frame.at[region, "新用户营收"], frame.at[region, "新用户订单量"])
    if key == "__all_channel_new_roi":
        revenue = float(all_channel.at[region, "新用户营收"]) if region in all_channel.index else 0.0
        return safe_div(revenue, frame.at[region, "消耗"])
    return float(frame.at[region, key])


def _write_web_template_sheet(ws, current, previous, all_current, all_previous):
    for row in range(1, ws.max_row + 1):
        section = ws.cell(row, 1).value
        if not isinstance(section, str) or not section.startswith("SEM-"):
            continue
        region = section.removeprefix("SEM-").strip()
        next_row = row + 1
        while next_row <= ws.max_row and ws.cell(next_row, 1).value:
            label = str(ws.cell(next_row, 1).value).strip()
            key = WEB_TEMPLATE_METRICS.get(label)
            if key:
                now = _metric_value(current, all_current, region, key)
                before = _metric_value(previous, all_previous, region, key)
                if key == "注册单价":
                    number_format = '"$"#,##0.00'
                elif key in {"__new_aov", "营收", "新用户营收", "消耗"}:
                    number_format = '"$"#,##0'
                elif key in {"注册付费率", "新用户注册付费率"}:
                    number_format = "0.00%"
                elif key in {"ROI", "新用户ROI", "__all_channel_new_roi"}:
                    number_format = "0%"
                else:
                    number_format = "#,##0"
                ws.cell(next_row, 2, now).number_format = number_format
                ws.cell(next_row, 3, before).number_format = number_format
                ws.cell(next_row, 4, safe_div(now - before, before) if before else 0.0).number_format = "0%"
            next_row += 1


def _app_template_values(google, appsflyer, start: date, end: date) -> dict[str, float]:
    paid_sources = set(APPSFLYER_BUY_CHANNELS)
    google_cost = _sum_cost(google, start, end)
    paid_revenue = _sum_appsflyer_revenue(appsflyer, start, end, media_sources=paid_sources)
    natural_revenue = _sum_appsflyer_revenue(
        appsflyer, start, end, media_sources=APPSFLYER_NATURAL_CHANNELS
    )
    google_revenue = _sum_appsflyer_revenue(
        appsflyer, start, end, media_sources={"googleadwords_int"}
    )
    asa_revenue = _sum_appsflyer_revenue(
        appsflyer, start, end, media_sources={"Apple Search Ads"}
    )
    values = {
        "APP投放收入": paid_revenue,
        "APP自然量收入": natural_revenue,
        "总计": paid_revenue + natural_revenue,
        "APP—投放新用户ROI（D7）": safe_div(paid_revenue, google_cost),
        "Google—新用户 ROI（D7）": safe_div(google_revenue, google_cost),
        "ASA—新用户ROI": 0.0,
        "APP投放总收入（D7）": paid_revenue,
        "Google收入（D7）": google_revenue,
        "ASA收入（D7）": asa_revenue,
    }
    region_revenue_total = 0.0
    for region in REGIONS:
        revenue = _sum_appsflyer_revenue(
            appsflyer, start, end,
            media_sources=paid_sources,
            country=APPSFLYER_REGION_COUNTRIES[region],
        )
        region_revenue_total += revenue
        values[f"地区收入_{region}"] = revenue
        values[f"地区ROI_{region}"] = safe_div(revenue, _sum_cost(google, start, end, region))
    values["地区收入_其他"] = paid_revenue - region_revenue_total
    return values


def _write_app_template_sheet(ws, google, appsflyer, current_range, previous_range):
    current = _app_template_values(google, appsflyer, *current_range)
    previous = _app_template_values(google, appsflyer, *previous_range)
    row_keys = {
        2: "APP投放收入", 3: "APP自然量收入", 4: "总计",
        9: "APP—投放新用户ROI（D7）", 10: "Google—新用户 ROI（D7）",
        11: "ASA—新用户ROI", 12: "APP投放总收入（D7）",
        13: "Google收入（D7）", 14: "ASA收入（D7）",
        16: "地区收入_印度", 17: "地区收入_印度尼西亚", 18: "地区收入_台湾",
        19: "地区收入_韩国", 20: "地区收入_其他", 22: "地区ROI_印度",
        23: "地区ROI_印度尼西亚", 24: "地区ROI_台湾", 25: "地区ROI_韩国",
    }
    for row, key in row_keys.items():
        now, before = current[key], previous[key]
        ws.cell(row, 2, now)
        ws.cell(row, 3, before)
        ws.cell(row, 4, safe_div(now - before, before) if before else 0.0)


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


def _source_mask(frame: pd.DataFrame, source_field: str | None, ad_groups: set[str]) -> pd.Series:
    if source_field is None:
        return pd.Series(False, index=frame.index)
    source = frame[source_field].fillna("").astype(str).str.strip().str.casefold()
    accepted = {value.casefold() for value in ad_groups if value}
    if not accepted:
        return pd.Series(False, index=frame.index)
    pattern = "|".join(re.escape(keyword) for keyword in sorted(accepted))
    return source.str.contains(pattern, regex=True)


def _campaign_backend_rows(
    reg, orders, google_campaign, campaign_mappings: list[tuple[str, set[str]]],
    start: date, end: date, new_only: bool,
):
    period_google = google_campaign[google_campaign["date"].dt.date.between(start, end)]
    reg_source = _first_column(reg, ("末次落地页",))
    order_source = _first_column(orders, ("末次落地页",))
    rows = []
    for campaign, ad_groups in campaign_mappings:
        campaign_google = period_google[
            period_google["campaign"].fillna("").astype(str).str.strip().eq(campaign)
        ]
        spend = float(campaign_google["cost"].sum())
        conversion_value = float(campaign_google["conversion_value"].sum())
        reg_mask = (
            reg["注册时间"].dt.date.between(start, end)
            & reg["__web"]
            & _source_mask(reg, reg_source, ad_groups)
        )
        order_mask = (
            orders["订单生成时间"].dt.date.between(start, end)
            & orders["__web"]
            & _source_mask(orders, order_source, ad_groups)
        )
        if "合作渠道" in orders.columns:
            order_mask &= orders["合作渠道"].fillna("").astype(str).str.strip().eq("")
        if new_only:
            order_mask &= orders["注册转化周期"].eq(0)
        else:
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


def _split_ad_groups(value) -> set[str]:
    return {
        item.strip()
        for item in re.split(r"[、,，;/；|\n]+", str(value or ""))
        if item.strip()
    }


def _template_campaign_mappings(ws, start_row: int, end_row: int) -> list[tuple[str, set[str]]]:
    mappings = []
    for row in range(start_row, end_row + 1):
        campaign = str(ws.cell(row, 1).value or "").strip()
        if campaign:
            mappings.append((campaign, _split_ad_groups(ws.cell(row, 8).value)))
    return mappings


def _write_front_back_roi_sheet(ws, reg, orders, google_campaign, bing, start: date, end: date):
    google_campaign = _normalize_google_campaign(google_campaign)
    bing_campaign = bing.rename(columns={"stat_date": "date", "ad_campaign": "campaign"})
    bing_campaign = bing_campaign.assign(
        campaign=bing_campaign["campaign"].fillna("").astype(str).str.strip(),
        conversion_value=0.0,
    )
    sections = (
        (3, 7, "除订阅", False, google_campaign, "Google"),
        (10, 14, "新用户", True, google_campaign, "Google"),
        (20, 23, "除订阅", False, bing_campaign, "Bing"),
        (26, 29, "新用户", True, bing_campaign, "Bing"),
    )
    headers = ["广告系列", "后台注册", "后台订单量", "消耗", "后台营收", "后台ROI", "Google前台ROI", "广告单元"]
    source_fields = (None, None)
    for header_row, last_row, label, new_only, performance, platform in sections:
        first_row = header_row + 1
        mappings = _template_campaign_mappings(ws, first_row, last_row)
        rows, reg_source, order_source = _campaign_backend_rows(
            reg, orders, performance, mappings, start, end, new_only
        )
        source_fields = (reg_source, order_source)
        ws.cell(header_row, 1, f"{start:%m.%d}-{end:%m.%d}({label})")
        ws.cell(header_row, 8, "末次落地页关键词")
        ws.cell(header_row, 7, f"{platform}前台ROI")
        for row, values in zip(range(first_row, last_row + 1), rows):
            for column, key in enumerate(headers[1:7], 2):
                ws.cell(row, column, values[key])
            ws.cell(row, 5).number_format = '"$"#,##0.00'
            if platform == "Bing":
                ws.cell(row, 7).value = None
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
    non_subscription_orders = web_orders[
        ~web_orders["末次来源类型"].fillna("").astype(str).str.upper().eq("AUTO_RENEWAL")
    ].copy()
    current = aggregate_metrics(
        web_reg, non_subscription_orders, web_google, web_bing,
        Period("本期", current_start, current_end), "SEM",
    )
    previous = aggregate_metrics(
        web_reg, non_subscription_orders, web_google, web_bing,
        Period("对比期", previous_start, previous_end), "SEM",
    )
    all_current = aggregate_metrics(
        web_reg, web_orders, web_google, web_bing,
        Period("本期", current_start, current_end), "全渠道",
    )
    all_previous = aggregate_metrics(
        web_reg, web_orders, web_google, web_bing,
        Period("对比期", previous_start, previous_end), "全渠道",
    )

    app_google, _, _, app_warnings = prepare_app_costs(
        app_google,
        pd.DataFrame(columns=["date", "region", "cost"]),
        pd.DataFrame(columns=["date", "region", "cost"]),
    )
    app_appsflyer, revenue_warnings = prepare_appsflyer_revenue(app_appsflyer)
    warnings.extend(app_warnings)
    warnings.extend(revenue_warnings)

    if not BRIEF_TEMPLATE.exists():
        raise FileNotFoundError(f"找不到数据监控模板：{BRIEF_TEMPLATE}")
    # 直接复用业务模板，保留原始工作表结构、字体、颜色、边框和数字格式。
    workbook = load_workbook(BRIEF_TEMPLATE)
    _write_web_template_sheet(
        workbook["Web-分地区"], current, previous, all_current, all_previous
    )
    _write_app_template_sheet(
        workbook["App-SEM"], app_google, app_appsflyer,
        (current_start, current_end), (previous_start, previous_end),
    )
    reg_source, order_source = _write_front_back_roi_sheet(
        workbook["Google前后台ROI对比"],
        web_reg, web_orders, web_google_campaign, web_bing, current_start, current_end,
    )
    warnings.append("Bing 消耗数据未提供前台转化价值，Bing 前台 ROI 留空。")
    if not reg_source:
        warnings.append("注册表缺少“末次落地页”字段，Google前后台ROI的后台注册将显示为0。")
    if not order_source:
        warnings.append("订单表缺少“末次落地页”字段，Google前后台ROI的后台订单和营收将显示为0。")
    output = BytesIO()
    workbook.save(output)
    return (
        f"周一周五数据简报_{current_start:%Y%m%d}-{current_end:%Y%m%d}.xlsx",
        output.getvalue(),
        warnings,
    )
