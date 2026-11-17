from __future__ import annotations

from datetime import date, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "assets" / "UGC数据监控.xlsx"
TRADITIONAL = {"中国香港", "香港", "中国澳门", "澳门", "中国台湾", "台湾", "港澳台"}
WEB_CHANNELS = ("SEM", "KOL", "SEO", "Affiliate", "Unassigned", "教育认证")


def _text(frame, field):
    if field not in frame:
        return pd.Series("", index=frame.index, dtype="object")
    return frame[field].fillna("").astype(str).str.strip().str.upper()


def _dedupe(frame, fields):
    for field in fields:
        if field in frame:
            return frame.drop_duplicates(field, keep="last")
    return frame


def _web_channel(frame):
    """严格复刻用户提供的SQL归因CASE与优先级。"""
    source = _text(frame, "末次来源类型")
    coop = _text(frame, "合作渠道")
    result = pd.Series(None, index=frame.index, dtype="object")
    result[source.eq("SEM")] = "SEM"
    result[source.isin({"KOL", "DIRECT", "SOCIAL_ORGANIC"})] = "KOL"
    result[source.eq("SEO") | source.str.startswith("REFERRAL")] = "SEO"
    result[source.str.startswith("AFFILIATE")] = "Affiliate"
    result[source.eq("UNASSIGNED")] = "Unassigned"
    result[source.eq("")] = "教育认证"
    # 订单合作渠道优先；注册表通常没有该字段，空列不会影响。
    result[coop.str.startswith("KOL")] = "KOL"
    result[coop.str.startswith("AFFILIATE")] = "Affiliate"
    invalid_coop = coop.ne("") & ~coop.str.startswith(("KOL", "AFFILIATE"))
    result[invalid_coop] = None
    return result


def _app_channel(frame):
    source = _text(frame, "APP渠道来源")
    source = source.where(source.ne(""), _text(frame, "末次来源类型"))
    source = source.where(source.ne(""), _text(frame, "渠道来源"))
    source = source.where(source.ne(""), _text(frame, "渠道类型"))
    detail = source + " " + _text(frame, "末次来源") + " " + _text(frame, "广告来源")
    result = pd.Series("自然来源", index=frame.index, dtype="object")
    result[detail.str.contains("ASA|APPLE SEARCH", regex=True)] = "ASA"
    result[detail.str.contains("SEM|GOOGLE|META|FACEBOOK|FB", regex=True)] = "SEM"
    return result


def _prepare(reg, orders, *, app_only=False):
    reg = _dedupe(reg.copy(), ("用户ID", "用户id"))
    orders = _dedupe(orders.copy(), ("订单id", "订单ID"))
    reg_required = ["注册时间", "注册地区"]
    order_required = ["国家", "注册转化周期"]
    if not app_only:
        reg_required.append("注册端口")
        order_required.append("支付端口")
    if "订单生成时间" not in orders and "支付时间" not in orders:
        order_required.append("订单生成时间或支付时间")
    missing_reg = [field for field in reg_required if field not in reg]
    missing_order = [field for field in order_required if field not in orders]
    if "订单金额" not in orders and "兑换美元" in orders:
        orders["订单金额"] = orders["兑换美元"]
    if "订单金额" not in orders:
        missing_order.append("订单金额或兑换美元")
    if missing_reg or missing_order:
        messages = []
        if missing_reg:
            messages.append("注册数据缺少：" + "、".join(missing_reg))
        if missing_order:
            messages.append("订单数据缺少：" + "、".join(missing_order))
        raise ValueError("；".join(messages))
    if app_only and "注册端口" not in reg:
        reg["注册端口"] = "APP"
    if app_only and "支付端口" not in orders:
        orders["支付端口"] = "APP"
    reg["__date"] = pd.to_datetime(reg["注册时间"], errors="coerce")
    reg["__region"] = reg["注册地区"].fillna("").astype(str).str.strip()
    reg["__web_channel"] = _web_channel(reg)
    reg["__app_channel"] = _app_channel(reg)
    order_date_field = "订单生成时间" if "订单生成时间" in orders else "支付时间"
    orders["__date"] = pd.to_datetime(orders[order_date_field], errors="coerce")
    orders["__region"] = orders["国家"].fillna("").astype(str).str.strip()
    orders["订单金额"] = pd.to_numeric(orders["订单金额"], errors="coerce").fillna(0.0)
    orders["注册转化周期"] = pd.to_numeric(orders["注册转化周期"], errors="coerce")
    orders["__web_channel"] = _web_channel(orders)
    orders["__app_channel"] = _app_channel(orders)
    if "支付位置" not in orders:
        orders["支付位置"] = ""
    return reg, orders


def _is_web(series):
    return series.fillna("").astype(str).str.contains("pc|移动|其他", case=False, regex=True)


def _is_app(series):
    return series.fillna("").astype(str).str.contains("APP", case=False, regex=False)


def _range(frame, start, end):
    return frame[frame["__date"].dt.date.between(start, end)]


def _channel_metrics(reg, orders, channel_field, channels):
    output = {}
    for channel in channels:
        channel_reg = reg[reg[channel_field].eq(channel)]
        channel_orders = orders[orders[channel_field].eq(channel)]
        new_orders = channel_orders[channel_orders["注册转化周期"].eq(0)]
        non_subscription = channel_orders[
            ~channel_orders["支付位置"].fillna("").astype(str).str.contains("订阅", na=False)
        ]
        output[channel] = {
            "reg": int(len(channel_reg)),
            "paid": int(len(new_orders)),
            "new_rev": float(new_orders["订单金额"].sum()),
            "all_rev": float(non_subscription["订单金额"].sum()),
        }
    return output


def _sum_channels(channels):
    return {
        key: sum(value[key] for value in channels.values())
        for key in ("reg", "paid", "new_rev", "all_rev")
    }


def _build_metrics(web_reg, web_orders, app_reg, app_orders, start, end):
    period_days = (end - start).days + 1
    prior = (start - timedelta(days=period_days), start - timedelta(days=1))
    output = {}
    for period, (period_start, period_end) in {
        "current": (start, end), "prior": prior
    }.items():
        period_web_reg = _range(web_reg, period_start, period_end)
        period_web_orders = _range(web_orders, period_start, period_end)
        region_web_reg = period_web_reg[period_web_reg["__region"].isin(TRADITIONAL)]
        region_web_orders = period_web_orders[period_web_orders["__region"].isin(TRADITIONAL)]
        web = _channel_metrics(
            region_web_reg[_is_web(region_web_reg["注册端口"])],
            region_web_orders[_is_web(region_web_orders["支付端口"])],
            "__web_channel", WEB_CHANNELS,
        )
        period_app_reg = _range(app_reg, period_start, period_end)
        period_app_orders = _range(app_orders, period_start, period_end)
        region_app_reg = period_app_reg[period_app_reg["__region"].isin(TRADITIONAL)]
        region_app_orders = period_app_orders[period_app_orders["__region"].isin(TRADITIONAL)]
        app = _channel_metrics(
            region_app_reg[_is_app(region_app_reg["注册端口"])],
            region_app_orders[_is_app(region_app_orders["支付端口"])],
            "__app_channel", ("SEM", "自然来源", "ASA"),
        )
        output[period] = {"web": web, "app": app}
    return output


def _rate(value, base):
    return value / base if base else 0


def _change(current, prior):
    return current / prior - 1 if prior else (0 if not current else 1)


def _series(current, prior, metric):
    if metric == "rate":
        current_value = _rate(current["paid"], current["reg"])
        prior_value = _rate(prior["paid"], prior["reg"])
    elif metric in {"cost", "new_roi", "all_roi"}:
        return "/", "/", "/"
    else:
        current_value, prior_value = current[metric], prior[metric]
    return current_value, prior_value, _change(current_value, prior_value)


def _write_three_rows(ws, rows, columns, current, prior, metric):
    for column, key in zip(columns, current):
        values = _series(current[key], prior[key], metric)
        for row, value in zip(rows, values):
            ws.cell(row, column, value)


def generate_ugc_report(web_reg, web_orders, app_reg, app_orders, start: date, end: date):
    web_reg, web_orders = _prepare(web_reg, web_orders)
    app_reg, app_orders = _prepare(app_reg, app_orders, app_only=True)
    metrics = _build_metrics(web_reg, web_orders, app_reg, app_orders, start, end)
    wb = load_workbook(TEMPLATE)
    ws = wb.active

    web_current = metrics["current"]["web"]
    web_prior = metrics["prior"]["web"]
    web_current = {"总计": _sum_channels(web_current), **web_current}
    web_prior = {"总计": _sum_channels(web_prior), **web_prior}
    app_current = metrics["current"]["app"]
    app_prior = metrics["prior"]["app"]
    app_current = {"总计": _sum_channels(app_current), **app_current}
    app_prior = {"总计": _sum_channels(app_prior), **app_prior}
    groups_web = [
        ((3, 4, 5), "reg"), ((6, 7, 8), "paid"), ((9, 10, 11), "rate"),
        ((12, 13, 14), "new_rev"), ((15, 16, 17), "all_rev"),
        ((18, 19, 20), "cost"), ((21, 22, 23), "new_roi"), ((24, 25, 26), "all_roi"),
    ]
    groups_app = [
        ((4, 5, 6), "reg"), ((7, 8, 9), "paid"), ((10, 11, 12), "rate"),
        ((13, 14, 15), "new_rev"), ((16, 17, 18), "all_rev"),
        ((19, 20, 21), "cost"), ((22, 23, 24), "new_roi"), ((25, 26, 27), "all_roi"),
    ]
    for rows, metric in groups_web:
        _write_three_rows(ws, rows, range(3, 10), web_current, web_prior, metric)
    for rows, metric in groups_app:
        _write_three_rows(ws, rows, range(15, 19), app_current, app_prior, metric)

    combined_current = {
        "WEB-KOL": web_current["KOL"],
        "APP-自然来源": app_current["自然来源"],
    }
    combined_prior = {
        "WEB-KOL": web_prior["KOL"],
        "APP-自然来源": app_prior["自然来源"],
    }
    combined_current = {"APP&WEB-总计": _sum_channels(combined_current), **combined_current}
    combined_prior = {"APP&WEB-总计": _sum_channels(combined_prior), **combined_prior}
    combined_groups = [
        ((31, 32, 34), "reg"), ((35, 36, 38), "paid"), ((39, 40, 41), "rate"),
        ((42, 43, 45), "new_rev"), ((46, 47, 49), "all_rev"),
        ((50, 51, 53), "cost"), ((54, 55, 56), "new_roi"), ((57, 58, 60), "all_roi"),
    ]
    for rows, metric in combined_groups:
        _write_three_rows(ws, rows, range(3, 6), combined_current, combined_prior, metric)
    for current_row, prior_row, increment_row in (
        (31, 32, 33), (35, 36, 37), (42, 43, 44), (46, 47, 48), (50, 51, 52)
    ):
        for column in range(3, 6):
            current_value = ws.cell(current_row, column).value
            prior_value = ws.cell(prior_row, column).value
            ws.cell(increment_row, column, "/" if current_value == "/" else current_value - prior_value)

    days = (end - start).days + 1
    prior_start, prior_end = start - timedelta(days=days), start - timedelta(days=1)
    current_label = f"{start.month}.{start.day}-{end.month}.{end.day}"
    prior_label = f"{prior_start.month}.{prior_start.day}-{prior_end.month}.{prior_end.day}"
    for current_row, prior_row in ((3,4),(6,7),(9,10),(12,13),(15,16),(18,19),(21,22),(24,25)):
        ws.cell(current_row, 2, current_label); ws.cell(prior_row, 2, prior_label)
    for current_row, prior_row in ((4,5),(7,8),(10,11),(13,14),(16,17),(19,20),(22,23),(25,26)):
        ws.cell(current_row, 14, current_label); ws.cell(prior_row, 14, prior_label)
    for current_row, prior_row in ((31,32),(35,36),(39,40),(42,43),(46,47),(50,51),(54,55),(57,58)):
        ws.cell(current_row, 2, current_label); ws.cell(prior_row, 2, prior_label)

    output = BytesIO()
    wb.save(output)
    warnings = ["成本和ROI没有订单/注册数据来源，已按“/”保留，未擅自填0。"]
    filename = f"UGC数据监控_{start:%Y%m%d}-{end:%Y%m%d}.xlsx"
    return filename, output.getvalue(), warnings
