from __future__ import annotations

from datetime import date
from typing import Mapping
import json

import pandas as pd
import requests
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account


API_VERSION = "v26"
API_ROOT = f"https://googleads.googleapis.com/{API_VERSION}"

# Google Ads API 返回 ISO 国家码；周报引擎沿用现有中文地区口径。
COUNTRY_NAMES = {
    "AE": "阿联酋", "AR": "阿根廷", "AU": "澳大利亚", "BO": "玻利维亚",
    "BR": "巴西", "CA": "加拿大", "CH": "瑞士", "CL": "智利",
    "CO": "哥伦比亚", "DO": "多米尼加", "EC": "厄瓜多尔", "EG": "埃及",
    "ES": "西班牙", "FR": "法国", "GB": "英国", "GT": "危地马拉",
    "HK": "香港", "HN": "洪都拉斯", "ID": "印度尼西亚", "IN": "印度",
    "IT": "意大利", "JM": "牙买加", "JP": "日本", "KH": "柬埔寨",
    "KR": "韩国", "KW": "科威特", "MM": "缅甸", "MO": "澳门",
    "MX": "墨西哥", "MY": "马来西亚", "NI": "尼加拉瓜", "NL": "荷兰",
    "PA": "巴拿马", "PE": "秘鲁", "PH": "菲律宾", "PK": "巴基斯坦",
    "PT": "葡萄牙", "PY": "巴拉圭", "QA": "卡塔尔", "RU": "俄罗斯",
    "SA": "沙特", "SE": "瑞典", "SG": "新加坡", "SV": "萨尔瓦多",
    "TH": "泰国", "TR": "土耳其", "TT": "特立尼达和多巴哥",
    "TW": "台湾", "US": "美国", "UY": "乌拉圭", "VE": "委内瑞拉",
    "VN": "越南",
}


class GoogleAdsAPIError(RuntimeError):
    pass


def google_ads_settings(secrets) -> dict:
    try:
        raw = secrets.get("google_ads", {})
    except Exception:
        return {}
    return dict(raw) if raw else {}


def google_ads_configured(secrets) -> bool:
    settings = google_ads_settings(secrets)
    account_keys = ("manager_customer_id", "web_customer_id", "app_customer_id")
    accounts_ready = all(str(settings.get(key, "")).strip() for key in account_keys)
    service_ready = bool(settings.get("service_account_json"))
    oauth_ready = all(
        str(settings.get(key, "")).strip()
        for key in ("client_id", "client_secret", "refresh_token")
    )
    return accounts_ready and (service_ready or oauth_ready)


def _digits(value) -> str:
    return "".join(character for character in str(value or "") if character.isdigit())


def _access_token(settings: Mapping) -> str:
    scopes = ["https://www.googleapis.com/auth/adwords"]
    service_json = settings.get("service_account_json")
    if service_json:
        info = json.loads(service_json) if isinstance(service_json, str) else dict(service_json)
        credentials = service_account.Credentials.from_service_account_info(info, scopes=scopes)
    else:
        credentials = Credentials(
            token=None,
            refresh_token=str(settings["refresh_token"]),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=str(settings["client_id"]),
            client_secret=str(settings["client_secret"]),
            scopes=scopes,
        )
    credentials.refresh(Request())
    return str(credentials.token)


def _search_stream(customer_id: str, query: str, headers: dict) -> list[dict]:
    response = requests.post(
        f"{API_ROOT}/customers/{customer_id}/googleAds:searchStream",
        headers=headers,
        json={"query": query},
        timeout=90,
    )
    if not response.ok:
        try:
            error = response.json().get("error", {})
            message = error.get("message") or response.text[:500]
            status = error.get("status", "")
        except Exception:
            message, status = response.text[:500], ""
        raise GoogleAdsAPIError(f"Google Ads API读取失败（{status or response.status_code}）：{message}")
    rows: list[dict] = []
    for batch in response.json():
        rows.extend(batch.get("results", []))
    return rows


def _geo_names(customer_id: str, criterion_ids: set[int], headers: dict) -> dict[int, str]:
    result: dict[int, str] = {}
    sorted_ids = sorted(criterion_ids)
    for offset in range(0, len(sorted_ids), 500):
        group = sorted_ids[offset : offset + 500]
        if not group:
            continue
        ids = ",".join(str(value) for value in group)
        query = f"""
            SELECT geo_target_constant.id, geo_target_constant.name,
                   geo_target_constant.country_code
            FROM geo_target_constant
            WHERE geo_target_constant.id IN ({ids})
        """
        for row in _search_stream(customer_id, query, headers):
            geo = row.get("geoTargetConstant", {})
            geo_id = int(geo.get("id", 0) or 0)
            code = str(geo.get("countryCode", "")).upper()
            name = COUNTRY_NAMES.get(code, str(geo.get("name", "")).strip() or "其他")
            result[geo_id] = name
    return result


def fetch_google_costs(
    secrets,
    account: str,
    start: date,
    end: date,
) -> pd.DataFrame:
    """读取与Google Ads“相符的地理位置”分天报告相同口径的费用。"""
    if start > end:
        raise ValueError("Google Ads API开始日期不能晚于结束日期。")
    settings = google_ads_settings(secrets)
    if not google_ads_configured(secrets):
        raise GoogleAdsAPIError("尚未配置Google Ads API的Streamlit Secrets。")
    account_key = f"{account}_customer_id"
    if account_key not in settings:
        raise GoogleAdsAPIError(f"Google Ads API缺少配置：{account_key}")
    customer_id = _digits(settings[account_key])
    manager_id = _digits(settings["manager_customer_id"])
    headers = {
        "Authorization": f"Bearer {_access_token(settings)}",
        "Content-Type": "application/json",
        "login-customer-id": manager_id,
    }
    query = f"""
        SELECT segments.date, campaign.name,
               geographic_view.country_criterion_id,
               geographic_view.location_type, metrics.cost_micros
        FROM geographic_view
        WHERE segments.date BETWEEN '{start:%Y-%m-%d}' AND '{end:%Y-%m-%d}'
    """
    raw_rows = _search_stream(customer_id, query, headers)
    criterion_ids = {
        int(row.get("geographicView", {}).get("countryCriterionId", 0) or 0)
        for row in raw_rows
    }
    criterion_ids.discard(0)
    geo_names = _geo_names(customer_id, criterion_ids, headers)
    records = []
    for row in raw_rows:
        geo = row.get("geographicView", {})
        geo_id = int(geo.get("countryCriterionId", 0) or 0)
        records.append({
            "天": row.get("segments", {}).get("date"),
            "地理位置": geo_names.get(geo_id, "其他"),
            "广告系列": row.get("campaign", {}).get("name", ""),
            "费用": float(row.get("metrics", {}).get("costMicros", 0) or 0) / 1_000_000,
            "位置类型": geo.get("locationType", ""),
            "__来源文件": f"Google Ads API · {customer_id}",
        })
    return pd.DataFrame(
        records,
        columns=["天", "地理位置", "广告系列", "费用", "位置类型", "__来源文件"],
    )
