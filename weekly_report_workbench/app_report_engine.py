from __future__ import annotations

from datetime import date, timedelta
from io import BytesIO
from pathlib import Path
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
APP_TEMPLATE = ROOT / "assets" / "APP投放表格.xlsx"

PERIODS = ("本周", "上周", "MTD")
REGIONS = ("台湾", "韩国", "印度", "印度尼西亚")
APP_COUNTRY_ROWS = {
    4: "中国台湾",
    5: "印度",
    6: "印度尼西亚",
    7: "美国",
    8: "中国香港",
    9: "马来西亚",
    10: "阿联酋",
    11: "韩国",
    12: "巴西",
    13: "菲律宾",
}

GOOGLE_REGION_MAP = {
    "台湾": "台湾", "中国台湾": "台湾", "taiwan": "台湾",
    "韩国": "韩国", "大韩民国": "韩国", "south korea": "韩国", "republic of korea": "韩国",
    "印度": "印度", "india": "印度",
    "印度尼西亚": "印度尼西亚", "印尼": "印度尼西亚", "indonesia": "印度尼西亚",
}

META_REGION_MAP = {
    "TW": "台湾", "TWN": "台湾", "TAIWAN": "台湾",
    "KR": "韩国", "KOR": "韩国", "SOUTH KOREA": "韩国", "REPUBLIC OF KOREA": "韩国",
    "IN": "印度", "IND": "印度", "INDIA": "印度",
    "ID": "印度尼西亚", "IDN": "印度尼西亚", "INDONESIA": "印度尼西亚",
}

OTHER_ROI_COUNTRIES = {
    "印度", "印度尼西亚", "马来西亚", "菲律宾", "巴西", "日本", "秘鲁", "阿根廷",
    "墨西哥", "葡萄牙", "智利", "厄瓜多尔", "新加坡", "巴拿马", "西班牙", "法国",
    "玻利维亚", "危地马拉", "泰国", "牙买加", "乌拉圭", "多米尼加", "哥伦比亚",
    "韩国", "特立尼达和多巴哥", "萨尔瓦多", "俄罗斯", "委内瑞拉", "洪都拉斯",
    "巴拉圭", "尼加拉瓜",
}

ENGLISH_ROI_COUNTRIES = {
    "英国", "美国", "澳大利亚", "加拿大", "越南", "沙特", "沙特阿拉伯", "阿联酋",
    "荷兰", "缅甸", "土耳其", "埃及", "卡塔尔", "瑞典", "科威特", "柬埔寨",
    "巴基斯坦", "瑞士", "意大利",
}

COUNTRY_ALIASES = {
    "TW": "中国台湾", "TWN": "中国台湾", "TAIWAN": "中国台湾", "台湾": "中国台湾",
    "HK": "中国香港", "HKG": "中国香港", "HONG KONG": "中国香港",
    "HONG KONG SAR": "中国香港", "香港": "中国香港",
    "GB": "英国", "GBR": "英国", "UK": "英国", "UNITED KINGDOM": "英国",
    "US": "美国", "USA": "美国", "UNITED STATES": "美国",
    "AU": "澳大利亚", "AUS": "澳大利亚", "AUSTRALIA": "澳大利亚",
    "CA": "加拿大", "CAN": "加拿大", "CANADA": "加拿大",
    "VN": "越南", "VNM": "越南", "VIETNAM": "越南",
    "SA": "沙特", "SAU": "沙特", "SAUDI ARABIA": "沙特",
    "AE": "阿联酋", "ARE": "阿联酋", "UAE": "阿联酋", "UNITED ARAB EMIRATES": "阿联酋",
    "NL": "荷兰", "NLD": "荷兰", "NETHERLANDS": "荷兰",
    "MM": "缅甸", "MMR": "缅甸", "MYANMAR": "缅甸",
    "TR": "土耳其", "TUR": "土耳其", "TURKEY": "土耳其", "TÜRKIYE": "土耳其",
    "EG": "埃及", "EGY": "埃及", "EGYPT": "埃及",
    "QA": "卡塔尔", "QAT": "卡塔尔", "QATAR": "卡塔尔",
    "SE": "瑞典", "SWE": "瑞典", "SWEDEN": "瑞典",
    "KW": "科威特", "KWT": "科威特", "KUWAIT": "科威特",
    "KH": "柬埔寨", "KHM": "柬埔寨", "CAMBODIA": "柬埔寨",
    "PK": "巴基斯坦", "PAK": "巴基斯坦", "PAKISTAN": "巴基斯坦",
    "CH": "瑞士", "CHE": "瑞士", "SWITZERLAND": "瑞士",
    "IT": "意大利", "ITA": "意大利", "ITALY": "意大利",
    "IN": "印度", "IND": "印度", "INDIA": "印度",
    "ID": "印度尼西亚", "IDN": "印度尼西亚", "INDONESIA": "印度尼西亚",
    "MY": "马来西亚", "MYS": "马来西亚", "MALAYSIA": "马来西亚",
    "PH": "菲律宾", "PHL": "菲律宾", "PHILIPPINES": "菲律宾",
    "BR": "巴西", "BRA": "巴西", "BRAZIL": "巴西",
    "JP": "日本", "JPN": "日本", "JAPAN": "日本",
    "PE": "秘鲁", "PER": "秘鲁", "PERU": "秘鲁",
    "AR": "阿根廷", "ARG": "阿根廷", "ARGENTINA": "阿根廷",
    "MX": "墨西哥", "MEX": "墨西哥", "MEXICO": "墨西哥",
    "PT": "葡萄牙", "PRT": "葡萄牙", "PORTUGAL": "葡萄牙",
    "CL": "智利", "CHL": "智利", "CHILE": "智利",
    "EC": "厄瓜多尔", "ECU": "厄瓜多尔", "ECUADOR": "厄瓜多尔",
    "SG": "新加坡", "SGP": "新加坡", "SINGAPORE": "新加坡",
    "PA": "巴拿马", "PAN": "巴拿马", "PANAMA": "巴拿马",
    "ES": "西班牙", "ESP": "西班牙", "SPAIN": "西班牙",
    "FR": "法国", "FRA": "法国", "FRANCE": "法国",
    "BO": "玻利维亚", "BOL": "玻利维亚", "BOLIVIA": "玻利维亚",
    "GT": "危地马拉", "GTM": "危地马拉", "GUATEMALA": "危地马拉",
    "TH": "泰国", "THA": "泰国", "THAILAND": "泰国",
    "JM": "牙买加", "JAM": "牙买加", "JAMAICA": "牙买加",
    "UY": "乌拉圭", "URY": "乌拉圭", "URUGUAY": "乌拉圭",
    "DO": "多米尼加", "DOM": "多米尼加", "DOMINICAN REPUBLIC": "多米尼加",
    "CO": "哥伦比亚", "COL": "哥伦比亚", "COLOMBIA": "哥伦比亚",
    "KR": "韩国", "KOR": "韩国", "SOUTH KOREA": "韩国", "REPUBLIC OF KOREA": "韩国",
    "TT": "特立尼达和多巴哥", "TTO": "特立尼达和多巴哥", "TRINIDAD AND TOBAGO": "特立尼达和多巴哥",
    "SV": "萨尔瓦多", "SLV": "萨尔瓦多", "EL SALVADOR": "萨尔瓦多",
    "RU": "俄罗斯", "RUS": "俄罗斯", "RUSSIA": "俄罗斯", "RUSSIAN FEDERATION": "俄罗斯",
    "VE": "委内瑞拉", "VEN": "委内瑞拉", "VENEZUELA": "委内瑞拉",
    "HN": "洪都拉斯", "HND": "洪都拉斯", "HONDURAS": "洪都拉斯",
    "PY": "巴拉圭", "PRY": "巴拉圭", "PARAGUAY": "巴拉圭",
    "NI": "尼加拉瓜", "NIC": "尼加拉瓜", "NICARAGUA": "尼加拉瓜",
}


def _canonical_country(value) -> str:
    text = str(value or "").strip()
    return COUNTRY_ALIASES.get(text.upper(), text)


def _roi_group(value) -> str:
    country = _canonical_country(value)
    if country in ENGLISH_ROI_COUNTRIES:
        return "英语地区"
    if country in OTHER_ROI_COUNTRIES:
        return "其他地区"
    return ""


def _read_csv(file) -> pd.DataFrame:
    file.seek(0)
    try:
        return pd.read_csv(file, encoding="utf-8-sig", low_memory=False)
    except UnicodeDecodeError:
        file.seek(0)
        return pd.read_csv(file, encoding="gb18030", low_memory=False)


def _read_google_excel(file) -> pd.DataFrame:
    """兼容 Google Ads 顶部有报告标题、日期范围的导出格式。"""
    file.seek(0)
    raw = pd.read_excel(file, header=None)
    header_index = None
    for index, row in raw.head(20).iterrows():
        values = {str(value).strip() for value in row.tolist() if pd.notna(value)}
        if "天" in values and ("费用" in values or "花费" in values):
            header_index = index
            break
    if header_index is None:
        file.seek(0)
        frame = pd.read_excel(file)
    else:
        frame = raw.iloc[header_index + 1 :].copy()
        frame.columns = [str(value).strip() for value in raw.iloc[header_index].tolist()]
    frame = frame.dropna(how="all")
    frame.columns = [str(column).strip() for column in frame.columns]
    return frame


def read_app_uploads(files: Iterable, source: str) -> pd.DataFrame:
    frames = []
    for file in files or []:
        name = str(file.name).lower()
        if name.endswith(".csv"):
            frame = _read_csv(file)
        elif source == "google":
            frame = _read_google_excel(file)
        else:
            file.seek(0)
            frame = pd.read_excel(file)
        frame.columns = [str(column).strip() for column in frame.columns]
        frame["__来源文件"] = Path(str(file.name)).name
        frames.append(frame)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _rename_first(frame: pd.DataFrame, aliases: dict[str, tuple[str, ...]]) -> pd.DataFrame:
    result = frame.copy()
    for target, candidates in aliases.items():
        if target in result.columns:
            continue
        for candidate in candidates:
            if candidate in result.columns:
                result = result.rename(columns={candidate: target})
                break
    return result


def prepare_app_costs(
    google: pd.DataFrame, meta: pd.DataFrame, asa: pd.DataFrame
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, list[str]]:
    warnings: list[str] = []
    google = _rename_first(
        google,
        {
            "date": ("天", "日期", "Day"),
            "region": ("相符的地理位置", "地理位置", "国家/地区"),
            "cost": ("费用", "花费", "Cost", "Spend"),
        },
    )
    meta = _rename_first(
        meta,
        {
            "date": ("报告开始日期", "日期", "Day"),
            "region": ("国家/地区", "国家/地区（实际位置）", "Country"),
            "cost": ("已花费金额 (USD)", "花费", "Spend", "Amount spent (USD)"),
        },
    )
    asa = _rename_first(
        asa,
        {
            "date": ("日期", "Date", "Day"),
            "region": ("国家或地区", "国家/地区", "Country or Region"),
            "cost": ("支出", "花费", "Spend"),
        },
    )
    for name, frame in (("Google", google), ("Meta", meta), ("ASA", asa)):
        missing = [column for column in ("date", "region", "cost") if column not in frame.columns]
        if missing:
            raise ValueError(f"{name}消耗数据缺少字段：{', '.join(missing)}")

    google["date"] = pd.to_datetime(google["date"], errors="coerce")
    google["cost"] = pd.to_numeric(google["cost"], errors="coerce").fillna(0.0)
    # Google 分天分地区导出同时包含地区明细和多组每日总计行。
    # 地区明细已经覆盖完整消耗，必须排除“总计：”行，否则总消耗会重复累加。
    google_total_rows = google["region"].fillna("").astype(str).str.strip().str.startswith("总计：")
    removed_google_totals = int(google_total_rows.sum())
    if removed_google_totals:
        google = google.loc[~google_total_rows].copy()
        warnings.append(f"Google消耗已排除 {removed_google_totals} 行分天总计，仅按分地区明细汇总。")
    google["region_group"] = (
        google["region"].fillna("").astype(str).str.strip().map(
            lambda value: GOOGLE_REGION_MAP.get(value, GOOGLE_REGION_MAP.get(value.lower(), "其他"))
        )
    )
    google["roi_group"] = google["region"].map(_roi_group)

    meta["date"] = pd.to_datetime(meta["date"], errors="coerce")
    meta["cost"] = pd.to_numeric(meta["cost"], errors="coerce").fillna(0.0)
    meta["region_group"] = (
        meta["region"].fillna("").astype(str).str.strip().str.upper().map(META_REGION_MAP).fillna("其他")
    )
    meta["roi_group"] = meta["region"].map(_roi_group)

    asa["date"] = pd.to_datetime(asa["date"], errors="coerce")
    asa["cost"] = pd.to_numeric(asa["cost"], errors="coerce").fillna(0.0)
    asa["region_group"] = (
        asa["region"].fillna("").astype(str).str.strip().map(
            lambda value: GOOGLE_REGION_MAP.get(value, GOOGLE_REGION_MAP.get(value.lower(), "其他"))
        )
    )
    asa["roi_group"] = asa["region"].map(_roi_group)

    if google["date"].isna().any():
        warnings.append(f"Google有 {int(google['date'].isna().sum())} 行日期无法识别，已忽略。")
    if meta["date"].isna().any():
        warnings.append(f"Meta有 {int(meta['date'].isna().sum())} 行日期无法识别，已忽略。")
    if asa["date"].isna().any():
        warnings.append(f"ASA有 {int(asa['date'].isna().sum())} 行日期无法识别，已忽略。")
    return google, meta, asa, warnings


def prepare_app_orders(orders: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    warnings: list[str] = []
    orders = orders.copy()
    if "订单金额" not in orders.columns and "兑换美元" in orders.columns:
        orders["订单金额"] = orders["兑换美元"]
        warnings.append("APP总数据的订单金额使用订单表“兑换美元”字段。")
    required = ["支付端口", "国家", "订单金额", "注册转化周期"]
    missing = [column for column in required if column not in orders.columns]
    if "支付时间" not in orders.columns and "订单生成时间" not in orders.columns:
        missing.append("支付时间或订单生成时间")
    if missing:
        raise ValueError(f"APP总数据所用订单表缺少字段：{', '.join(missing)}")
    for key in ("订单id", "订单ID"):
        if key in orders.columns:
            before = len(orders)
            orders = orders.drop_duplicates(key, keep="last")
            if before != len(orders):
                warnings.append(f"APP总数据按{key}去除 {before-len(orders)} 条重复订单。")
            break
    date_field = "支付时间" if "支付时间" in orders.columns else "订单生成时间"
    if date_field != "支付时间":
        warnings.append("订单表缺少支付时间，APP总数据暂按订单生成时间统计。")
    orders["app_date"] = pd.to_datetime(orders[date_field], errors="coerce")
    orders["订单金额"] = pd.to_numeric(
        orders["订单金额"].astype(str).str.replace(",", "", regex=False).str.replace(r"[^0-9.\-]", "", regex=True),
        errors="coerce",
    ).fillna(0.0)
    orders["注册转化周期"] = pd.to_numeric(orders["注册转化周期"], errors="coerce")
    orders["roi_group"] = orders["国家"].map(_roi_group)
    app_mask = orders["支付端口"].fillna("").astype(str).str.contains("APP", case=False, regex=False)
    invalid_dates = int(orders.loc[app_mask, "app_date"].isna().sum())
    if invalid_dates:
        warnings.append(f"APP订单有 {invalid_dates} 行支付日期无法识别，已忽略。")
    return orders.loc[app_mask & orders["app_date"].notna()].copy(), warnings


def _period_ranges(start: date, end: date) -> dict[str, tuple[date, date]]:
    days = (end - start).days + 1
    return {
        "本周": (start, end),
        "上周": (start - timedelta(days=days), start - timedelta(days=1)),
        "MTD": (end.replace(day=1), end),
    }


def _sum_cost(frame: pd.DataFrame, start: date, end: date, region: str | None = None) -> float:
    mask = frame["date"].dt.date.between(start, end)
    if region is not None:
        mask &= frame["region_group"].eq(region)
    return float(frame.loc[mask, "cost"].sum())


def _sum_cost_group(frame: pd.DataFrame, start: date, end: date, roi_group: str) -> float:
    mask = frame["date"].dt.date.between(start, end) & frame["roi_group"].eq(roi_group)
    return float(frame.loc[mask, "cost"].sum())


def _sum_country_cost(frame: pd.DataFrame, start: date, end: date, country: str) -> float:
    mask = frame["date"].dt.date.between(start, end)
    canonical_country = frame["region"].map(_canonical_country)
    return float(frame.loc[mask & canonical_country.eq(country), "cost"].sum())


def _sum_order_revenue(
    orders: pd.DataFrame, start: date, end: date, *, new_only: bool = False,
    roi_group: str | None = None, country: str | None = None,
) -> float:
    mask = orders["app_date"].dt.date.between(start, end)
    if new_only:
        mask &= orders["注册转化周期"].eq(0)
    if roi_group is not None:
        mask &= orders["roi_group"].eq(roi_group)
    if country is not None:
        mask &= orders["国家"].map(_canonical_country).eq(country)
    return float(orders.loc[mask, "订单金额"].sum())


def _safe_div(numerator: float, denominator: float) -> float:
    return float(numerator) / float(denominator) if denominator else 0.0


def _change(current: float, previous: float) -> float:
    return current / previous - 1 if previous else 0.0


def generate_app_report(
    google: pd.DataFrame,
    meta: pd.DataFrame,
    asa: pd.DataFrame,
    orders: pd.DataFrame,
    start: date,
    end: date,
    manual: dict[str, float],
    targets: dict[str, float],
) -> tuple[str, bytes, list[str], dict[str, dict[str, float]]]:
    if start > end:
        raise ValueError("开始日期不能晚于结束日期。")
    if not APP_TEMPLATE.exists():
        raise FileNotFoundError(f"找不到APP投放模板：{APP_TEMPLATE}")

    google, meta, asa, warnings = prepare_app_costs(google, meta, asa)
    orders, order_warnings = prepare_app_orders(orders)
    warnings.extend(order_warnings)
    ranges = _period_ranges(start, end)
    values: dict[str, dict[str, float]] = {}
    cost_summary: dict[str, dict[str, float]] = {}
    total_values: dict[str, dict[str, float]] = {}

    for period, (period_start, period_end) in ranges.items():
        google_cost = _sum_cost(google, period_start, period_end)
        meta_cost = _sum_cost(meta, period_start, period_end)
        asa_cost = _sum_cost(asa, period_start, period_end)
        google_revenue = float(manual.get(f"google_revenue_{period}", 0.0))
        meta_revenue = float(manual.get(f"meta_revenue_{period}", 0.0))
        asa_revenue = float(manual.get(f"asa_revenue_{period}", 0.0))
        total_revenue = google_revenue + meta_revenue + asa_revenue
        total_cost = google_cost + meta_cost + asa_cost

        values[period] = {
            "APP—投放新用户ROI（D7）": _safe_div(total_revenue, total_cost),
            "Google—投放新用户 ROI（D7）": _safe_div(google_revenue, google_cost),
            "Meta—投放新用户 ROI（D7）": _safe_div(meta_revenue, meta_cost),
            "ASA—新用户ROI": _safe_div(asa_revenue, asa_cost),
            "APP投放总收入（D7）": total_revenue,
            "Google收入（D7）": google_revenue,
            "Meta收入（D7）": meta_revenue,
            "ASA收入（D7）": asa_revenue,
        }
        cost_summary[period] = {
            "Google消耗": google_cost,
            "Meta消耗": meta_cost,
            "ASA消耗": asa_cost,
            "总消耗": total_cost,
        }

        region_revenue_total = 0.0
        for region in REGIONS:
            revenue = float(manual.get(f"region_revenue_{region}_{period}", 0.0))
            region_revenue_total += revenue
            values[period][f"地区收入_{region}"] = revenue
            region_cost = (
                _sum_cost(google, period_start, period_end, region)
                + _sum_cost(meta, period_start, period_end, region)
                + _sum_cost(asa, period_start, period_end, region)
            )
            values[period][f"地区ROI_{region}"] = _safe_div(revenue, region_cost)
        values[period]["地区收入_其他"] = total_revenue - region_revenue_total

        app_total_revenue = _sum_order_revenue(orders, period_start, period_end)
        app_new_revenue = _sum_order_revenue(orders, period_start, period_end, new_only=True)
        app_natural_revenue = float(manual.get(f"app_natural_revenue_{period}", 0.0))
        total_values[period] = {
            "APP—新老综合ROI": _safe_div(app_total_revenue, total_cost),
            "APP—新用户综合ROI（D0）": _safe_div(app_new_revenue, total_cost),
            "APP—投放新用户ROI（D7）": values[period]["APP—投放新用户ROI（D7）"],
            "APP—应用商店（ASA）新用户ROI（D7）": values[period]["ASA—新用户ROI"],
            "APP—其他地区新老综合ROI": _safe_div(
                _sum_order_revenue(orders, period_start, period_end, roi_group="其他地区"),
                sum(_sum_cost_group(frame, period_start, period_end, "其他地区") for frame in (google, meta, asa)),
            ),
            "APP—英语地区新老综合ROI": _safe_div(
                _sum_order_revenue(orders, period_start, period_end, roi_group="英语地区"),
                sum(_sum_cost_group(frame, period_start, period_end, "英语地区") for frame in (google, meta, asa)),
            ),
            "总收入": app_total_revenue,
            "APP投放收入": total_revenue,
            "APP自然量收入": app_natural_revenue,
            "老用户收入": app_total_revenue - total_revenue - app_natural_revenue,
            "总成本": total_cost,
        }

    workbook = load_workbook(APP_TEMPLATE)
    sheet = workbook[workbook.sheetnames[0]]
    month_label = f"{end.month}月"
    sheet["E2"] = f"{month_label}月MTD"
    sheet["F2"] = f"{month_label}目标"
    sheet["G2"] = f"{month_label}完成率"
    sheet["E11"] = f"{month_label}月MTD"
    sheet["F11"] = f"{month_label}目标"
    sheet["G11"] = f"{month_label}完成率"
    sheet["E17"] = f"{month_label}月MTD"
    sheet["F17"] = f"{month_label}目标"
    sheet["G17"] = f"{month_label}完成率"

    row_keys = {
        3: "APP—投放新用户ROI（D7）",
        4: "Google—投放新用户 ROI（D7）",
        5: "Meta—投放新用户 ROI（D7）",
        6: "ASA—新用户ROI",
        7: "APP投放总收入（D7）",
        8: "Google收入（D7）",
        9: "Meta收入（D7）",
        10: "ASA收入（D7）",
        12: "地区收入_台湾",
        13: "地区收入_韩国",
        14: "地区收入_印度",
        15: "地区收入_印度尼西亚",
        16: "地区收入_其他",
        18: "地区ROI_台湾",
        19: "地区ROI_韩国",
        20: "地区ROI_印度",
        21: "地区ROI_印度尼西亚",
    }
    for row, key in row_keys.items():
        current = values["本周"][key]
        previous = values["上周"][key]
        mtd = values["MTD"][key]
        target = float(targets.get(key, 0.0))
        sheet.cell(row, 2, current)
        sheet.cell(row, 3, previous)
        sheet.cell(row, 4, _change(current, previous))
        sheet.cell(row, 5, mtd)
        sheet.cell(row, 6, target)
        sheet.cell(row, 7, _safe_div(mtd, target))

    for row in list(range(3, 7)) + list(range(18, 22)):
        for column in range(2, 8):
            sheet.cell(row, column).number_format = "0%"
    for row in range(7, 17):
        for column in (2, 3, 5, 6):
            sheet.cell(row, column).number_format = '$#,##0'
        for column in (4, 7):
            sheet.cell(row, column).number_format = "0%"

    # 模板H列是内部统计口径说明，不属于最终对外报表。
    if sheet.max_column >= 8:
        sheet.delete_cols(8, sheet.max_column - 7)

    total_sheet = workbook["Sheet2"]
    total_sheet["B2"] = f"{start:%m.%d}-{end:%m.%d}"
    previous_start, previous_end = ranges["上周"]
    total_sheet["C2"] = f"{previous_start:%m.%d}-{previous_end:%m.%d}"
    total_sheet["E2"] = f"{end.month}月MTD"
    total_rows = {
        3: "APP—新老综合ROI",
        4: "APP—新用户综合ROI（D0）",
        5: "APP—投放新用户ROI（D7）",
        6: "APP—应用商店（ASA）新用户ROI（D7）",
        7: "APP—其他地区新老综合ROI",
        8: "APP—英语地区新老综合ROI",
        9: "总收入",
        10: "APP投放收入",
        11: "APP自然量收入",
        12: "老用户收入",
        13: "总成本",
    }
    for row, key in total_rows.items():
        current = total_values["本周"][key]
        previous = total_values["上周"][key]
        mtd = total_values["MTD"][key]
        target = float(targets.get(key, 0.0))
        total_sheet.cell(row, 2, current)
        total_sheet.cell(row, 3, previous)
        total_sheet.cell(row, 4, _change(current, previous))
        total_sheet.cell(row, 5, mtd)
        total_sheet.cell(row, 6, target)
        total_sheet.cell(row, 7, _safe_div(mtd, target))
    for row in range(3, 9):
        for column in range(2, 8):
            total_sheet.cell(row, column).number_format = "0%"
    for row in range(9, 14):
        for column in (2, 3, 5, 6):
            total_sheet.cell(row, column).number_format = '$#,##0'
        for column in (4, 7):
            total_sheet.cell(row, column).number_format = "0%"
    if total_sheet.max_column >= 8:
        total_sheet.delete_cols(8, total_sheet.max_column - 7)

    # Sheet3：所选月份MTD的APP分国家营收、消耗和ROI。
    country_sheet = workbook["Sheet3"]
    country_sheet["B2"] = "新老用户含订阅营收"
    country_sheet["E3"] = "（含订阅）"
    mtd_start, mtd_end = ranges["MTD"]
    for row, country in APP_COUNTRY_ROWS.items():
        all_revenue = _sum_order_revenue(orders, mtd_start, mtd_end, country=country)
        new_revenue = _sum_order_revenue(
            orders, mtd_start, mtd_end, new_only=True, country=country
        )
        country_cost = sum(
            _sum_country_cost(frame, mtd_start, mtd_end, country)
            for frame in (google, meta, asa)
        )
        country_sheet.cell(row, 2, all_revenue)
        country_sheet.cell(row, 3, new_revenue)
        country_sheet.cell(row, 4, country_cost)
        country_sheet.cell(row, 5, _safe_div(all_revenue, country_cost))
        country_sheet.cell(row, 6, _safe_div(new_revenue, country_cost))

    total_all_revenue = _sum_order_revenue(orders, mtd_start, mtd_end)
    total_new_revenue = _sum_order_revenue(orders, mtd_start, mtd_end, new_only=True)
    total_mtd_cost = sum(_sum_cost(frame, mtd_start, mtd_end) for frame in (google, meta, asa))
    for column, value in enumerate(
        (
            total_all_revenue,
            total_new_revenue,
            total_mtd_cost,
            _safe_div(total_all_revenue, total_mtd_cost),
            _safe_div(total_new_revenue, total_mtd_cost),
        ),
        start=2,
    ):
        country_sheet.cell(14, column, value)
    for row in range(4, 15):
        for column in (2, 3, 4):
            country_sheet.cell(row, column).number_format = '$#,##0'
        for column in (5, 6):
            country_sheet.cell(row, column).number_format = "0%"
    # 第17行是模板维护用SQL说明，不放进最终下载表。
    if country_sheet.max_row > 14:
        country_sheet.delete_rows(15, country_sheet.max_row - 14)
    country_sheet.print_area = "A2:F14"
    country_sheet.page_setup.orientation = "landscape"
    country_sheet.page_setup.fitToWidth = 1
    country_sheet.page_setup.fitToHeight = 1
    country_sheet.sheet_properties.pageSetUpPr.fitToPage = True

    output = BytesIO()
    workbook.save(output)
    filename = f"APP投放营收_{start:%Y.%m.%d}-{end:%Y.%m.%d}.xlsx"
    return filename, output.getvalue(), warnings, cost_summary
