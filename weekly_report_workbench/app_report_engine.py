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


def prepare_app_costs(google: pd.DataFrame, meta: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
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
    for name, frame in (("Google", google), ("Meta", meta)):
        missing = [column for column in ("date", "region", "cost") if column not in frame.columns]
        if missing:
            raise ValueError(f"{name}消耗数据缺少字段：{', '.join(missing)}")

    google["date"] = pd.to_datetime(google["date"], errors="coerce")
    google["cost"] = pd.to_numeric(google["cost"], errors="coerce").fillna(0.0)
    google["region_group"] = (
        google["region"].fillna("").astype(str).str.strip().map(
            lambda value: GOOGLE_REGION_MAP.get(value, GOOGLE_REGION_MAP.get(value.lower(), "其他"))
        )
    )

    meta["date"] = pd.to_datetime(meta["date"], errors="coerce")
    meta["cost"] = pd.to_numeric(meta["cost"], errors="coerce").fillna(0.0)
    meta["region_group"] = (
        meta["region"].fillna("").astype(str).str.strip().str.upper().map(META_REGION_MAP).fillna("其他")
    )

    if google["date"].isna().any():
        warnings.append(f"Google有 {int(google['date'].isna().sum())} 行日期无法识别，已忽略。")
    if meta["date"].isna().any():
        warnings.append(f"Meta有 {int(meta['date'].isna().sum())} 行日期无法识别，已忽略。")
    return google, meta, warnings


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


def _safe_div(numerator: float, denominator: float) -> float:
    return float(numerator) / float(denominator) if denominator else 0.0


def _change(current: float, previous: float) -> float:
    return current / previous - 1 if previous else 0.0


def generate_app_report(
    google: pd.DataFrame,
    meta: pd.DataFrame,
    start: date,
    end: date,
    manual: dict[str, float],
    targets: dict[str, float],
) -> tuple[str, bytes, list[str], dict[str, dict[str, float]]]:
    if start > end:
        raise ValueError("开始日期不能晚于结束日期。")
    if not APP_TEMPLATE.exists():
        raise FileNotFoundError(f"找不到APP投放模板：{APP_TEMPLATE}")

    google, meta, warnings = prepare_app_costs(google, meta)
    ranges = _period_ranges(start, end)
    values: dict[str, dict[str, float]] = {}
    cost_summary: dict[str, dict[str, float]] = {}

    for period, (period_start, period_end) in ranges.items():
        google_cost = _sum_cost(google, period_start, period_end)
        meta_cost = _sum_cost(meta, period_start, period_end)
        asa_cost = float(manual.get(f"asa_cost_{period}", 0.0))
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
                + float(manual.get(f"asa_region_cost_{region}_{period}", 0.0))
            )
            values[period][f"地区ROI_{region}"] = _safe_div(revenue, region_cost)
        values[period]["地区收入_其他"] = total_revenue - region_revenue_total

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

    output = BytesIO()
    workbook.save(output)
    filename = f"APP投放营收_{start:%Y.%m.%d}-{end:%Y.%m.%d}.xlsx"
    return filename, output.getvalue(), warnings, cost_summary
