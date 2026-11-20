from __future__ import annotations

from datetime import date
from io import BytesIO
from typing import Mapping
from zipfile import ZipFile
import time

import pandas as pd
import requests


REPORTING_ROOT = "https://reporting.api.bingads.microsoft.com/Reporting/v13/GenerateReport"
TOKEN_URL = "https://oauth2.googleapis.com/token"


class BingAdsAPIError(RuntimeError):
    pass


def bing_ads_settings(secrets) -> dict:
    try:
        raw = secrets.get("bing_ads", {})
    except Exception:
        return {}
    return dict(raw) if raw else {}


def bing_ads_configured(secrets) -> bool:
    settings = bing_ads_settings(secrets)
    required = (
        "developer_token", "client_id", "client_secret", "refresh_token",
        "customer_id", "account_id",
    )
    return all(str(settings.get(key, "")).strip() for key in required)


def _digits(value) -> str:
    return "".join(character for character in str(value or "") if character.isdigit())


def _access_token(settings: Mapping) -> str:
    response = requests.post(
        TOKEN_URL,
        data={
            "client_id": str(settings["client_id"]),
            "client_secret": str(settings["client_secret"]),
            "refresh_token": str(settings["refresh_token"]),
            "grant_type": "refresh_token",
        },
        timeout=30,
    )
    if not response.ok:
        raise BingAdsAPIError(
            f"Bing授权刷新失败（{response.status_code}）：{response.text[:300]}"
        )
    token = str(response.json().get("access_token", "")).strip()
    if not token:
        raise BingAdsAPIError("Bing授权刷新成功，但未返回访问令牌。")
    return token


def _request_json(url: str, headers: dict, payload: dict) -> dict:
    response = requests.post(url, headers=headers, json=payload, timeout=60)
    if not response.ok:
        try:
            detail = response.json()
        except Exception:
            detail = response.text[:500]
        raise BingAdsAPIError(f"Bing Ads API读取失败（{response.status_code}）：{detail}")
    try:
        return response.json()
    except Exception as exc:
        raise BingAdsAPIError("Bing Ads API返回了无法识别的数据。") from exc


def _report_request(account_id: int, start: date, end: date) -> dict:
    return {
        "ReportRequest": {
            "ExcludeColumnHeaders": False,
            "ExcludeReportFooter": True,
            "ExcludeReportHeader": True,
            "Format": "Csv",
            "FormatVersion": "2.0",
            "ReportName": "Weekly Report Bing User Location",
            "ReturnOnlyCompleteData": False,
            "Type": "UserLocationPerformanceReportRequest",
            "Aggregation": "Daily",
            "Columns": [
                "TimePeriod", "Country", "CampaignName",
                "Impressions", "Clicks", "Spend", "Ctr",
            ],
            "Scope": {"AccountIds": [account_id]},
            "Time": {
                "CustomDateRangeStart": {
                    "Day": start.day, "Month": start.month, "Year": start.year,
                },
                "CustomDateRangeEnd": {
                    "Day": end.day, "Month": end.month, "Year": end.year,
                },
                "ReportTimeZone": "BeijingChongqingHongKongUrumqi",
            },
        }
    }


def _download_report(url: str) -> pd.DataFrame:
    response = requests.get(url, timeout=90)
    if not response.ok:
        raise BingAdsAPIError(f"Bing报表下载失败（{response.status_code}）。")
    content = response.content
    try:
        with ZipFile(BytesIO(content)) as archive:
            names = [name for name in archive.namelist() if name.lower().endswith(".csv")]
            if not names:
                raise BingAdsAPIError("Bing报表压缩包中没有CSV文件。")
            content = archive.read(names[0])
    except BingAdsAPIError:
        raise
    except Exception:
        pass
    try:
        return pd.read_csv(BytesIO(content), encoding="utf-8-sig")
    except Exception as exc:
        raise BingAdsAPIError("Bing报表CSV无法读取。") from exc


def _normalize_report(frame: pd.DataFrame, account_id: str) -> pd.DataFrame:
    aliases = {
        "TimePeriod": "Date",
        "Country": "Country/region (physical location)",
        "CampaignName": "Campaign name",
        "Impressions": "Impressions",
        "Clicks": "Clicks",
        "Spend": "Spend",
        "Ctr": "CTR",
    }
    missing = [column for column in aliases if column not in frame.columns]
    if missing:
        raise BingAdsAPIError(f"Bing API报表缺少字段：{', '.join(missing)}")
    result = frame[list(aliases)].rename(columns=aliases).copy()
    result["Date"] = pd.to_datetime(result["Date"], errors="coerce")
    for column in ("Impressions", "Clicks"):
        result[column] = pd.to_numeric(
            result[column].astype(str).str.replace(",", "", regex=False), errors="coerce"
        ).fillna(0)
    result["Spend"] = pd.to_numeric(
        result["Spend"].astype(str).str.replace(r"[^0-9.\-]", "", regex=True),
        errors="coerce",
    ).fillna(0.0)
    ctr_text = result["CTR"].astype(str)
    ctr_values = pd.to_numeric(ctr_text.str.replace("%", "", regex=False), errors="coerce")
    result["CTR"] = ctr_values.where(~ctr_text.str.contains("%", regex=False), ctr_values / 100)
    result["__来源文件"] = f"Bing Ads API · {account_id}"
    return result[result["Date"].notna()].reset_index(drop=True)


def fetch_bing_costs(secrets, start: date, end: date) -> pd.DataFrame:
    """读取Bing“用户位置”分天分地区报告，与现有上传文件保持同一口径。"""
    if start > end:
        raise ValueError("Bing Ads API开始日期不能晚于结束日期。")
    settings = bing_ads_settings(secrets)
    if not bing_ads_configured(secrets):
        raise BingAdsAPIError("尚未配置Bing Ads API的Streamlit Secrets。")
    account_id = _digits(settings["account_id"])
    customer_id = _digits(settings["customer_id"])
    headers = {
        "Authorization": f"Bearer {_access_token(settings)}",
        "Content-Type": "application/json",
        "DeveloperToken": str(settings["developer_token"]),
        "CustomerId": customer_id,
        "CustomerAccountId": account_id,
        "IdentityProvider": str(settings.get("identity_provider", "Google")),
    }
    submitted = _request_json(
        f"{REPORTING_ROOT}/Submit",
        headers,
        _report_request(int(account_id), start, end),
    )
    request_id = submitted.get("ReportRequestId") or submitted.get("reportRequestId")
    if not request_id:
        raise BingAdsAPIError(f"Bing报表提交后未返回任务编号：{submitted}")
    download_url = ""
    last_status = ""
    for _ in range(24):
        polled = _request_json(
            f"{REPORTING_ROOT}/Poll", headers, {"ReportRequestId": request_id}
        )
        status_block = polled.get("ReportRequestStatus", polled)
        last_status = str(status_block.get("Status", ""))
        download_url = str(status_block.get("ReportDownloadUrl", "") or "")
        if last_status.lower() == "success" and download_url:
            break
        if last_status.lower() == "error":
            raise BingAdsAPIError(f"Bing报表生成失败：{status_block}")
        time.sleep(1.5)
    if not download_url:
        raise BingAdsAPIError(f"Bing报表生成超时（当前状态：{last_status or '未知'}）。")
    return _normalize_report(_download_report(download_url), account_id)
