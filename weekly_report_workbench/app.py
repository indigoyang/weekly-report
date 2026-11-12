from datetime import date, timedelta
from copy import deepcopy
from pathlib import Path
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile
from html import escape

import streamlit as st
import pandas as pd

from report_engine import read_uploads
from app_report_engine import generate_app_report, read_app_uploads
from v3_report_engine import generate_v3_reports
from target_store import (
    ALL_CHANNEL_TARGET_METRICS,
    ALL_CHANNEL_TARGET_ROWS,
    SEM_TARGET_METRICS,
    SEM_TARGET_ROWS,
    NATURAL_TARGET_METRICS,
    NATURAL_TARGET_ROWS,
    load_targets,
    save_targets,
)
from rule_store import load_rules, save_rules, validate_rules


st.set_page_config(page_title="数据自动化", page_icon="A", layout="wide", initial_sidebar_state="expanded")

# 报告结构升级时清理当前浏览器会话里旧版本生成的结果。
# 否则Streamlit热更新后，页面可能继续显示更新前缓存的4份报告。
REPORT_SCHEMA_VERSION = "2026-11-10-app-target-download-v8"
if st.session_state.get("report_schema_version") != REPORT_SCHEMA_VERSION:
    st.session_state.pop("reports", None)
    st.session_state.pop("warnings", None)
    st.session_state.pop("app_report", None)
    st.session_state.pop("app_report_warnings", None)
    st.session_state["report_schema_version"] = REPORT_SCHEMA_VERSION
admin_mode = str(st.query_params.get("admin", "")) == "1"
active_port = str(st.query_params.get("port", "web")).lower()
if active_port not in {"web", "app"}:
    active_port = "web"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@700;900&family=Inter:wght@400;500;600;700&display=swap');
*{box-sizing:border-box}:root{--paper:#f6f5f0;--surface:#fff;--ink:#111;--lime:#e3f85a;--muted:#6b7280;--line:#111;--soft:#ecebe5;--success:#dff8a8}
html{scroll-behavior:smooth}.stApp{background:var(--paper);color:var(--ink);font-family:'Inter','PingFang SC',sans-serif}.block-container{max-width:1480px;padding:34px 42px 80px;min-height:100vh;background:var(--paper)}
[data-testid="stHeader"]{background:transparent}[data-testid="stToolbar"]{right:1rem}[data-testid="stSidebar"]{width:320px!important;min-width:320px!important;background:var(--paper);border-right:1px solid #1112}[data-testid="stSidebar"]>div:first-child{width:320px!important;padding:48px 30px}[data-testid="stSidebarCollapseButton"],[data-testid="collapsedControl"]{display:none}
.brand-mark{width:46px;height:46px;border-radius:14px;background:#111;color:var(--lime);display:grid;place-items:center;font-family:'Noto Serif SC',serif;font-size:21px;font-weight:900;margin:0 auto 10px;box-shadow:3px 3px 0 var(--lime);transform:translate(-1px,-1px)}.brand-caption{text-align:center;font-size:7px;font-weight:800;letter-spacing:1.7px;color:#9ca3af;margin:0 0 25px}.nav-link{width:40px;height:40px;display:grid;place-items:center;margin:10px auto;border-radius:50%;border:1px solid transparent;color:#9ca3af;text-decoration:none!important;font-size:17px;transition:.18s}.nav-link:hover{color:#111;background:#e5e7eb}.nav-link.active{color:#111;background:var(--lime);border-color:#111;box-shadow:2px 2px 0 #111;transform:translate(-1px,-1px)}.port-label{font-size:8px;font-weight:800;letter-spacing:1.2px;text-align:center;color:#9ca3af;margin:13px 0 2px}.port-divider{width:34px;height:1px;background:#1112;margin:15px auto}.nav-space{height:205px}.nav-hint{font-size:9px;color:#9ca3af;text-align:center;line-height:1.35;margin:0 5px}
[data-testid="stSidebar"] .sidebar-section-title{font-family:'Noto Serif SC',serif;font-size:18px;font-weight:700;letter-spacing:1px;color:#a1a7b2;margin:0 10px 28px}
[data-testid="stSidebar"] .stButton{width:100%!important;margin:0 0 14px!important}
[data-testid="stSidebar"] .stButton>button{position:relative!important;justify-content:flex-start!important;width:100%!important;max-width:none!important;height:78px!important;min-height:78px!important;padding:0 30px!important;border:1px solid transparent!important;border-radius:18px!important;background:transparent!important;color:#747b87!important;font-size:21px!important;font-weight:800!important;box-shadow:none!important;letter-spacing:.2px!important}
[data-testid="stSidebar"] .stButton>button p{width:100%!important;text-align:left!important;font-size:21px!important;font-weight:800!important;white-space:nowrap!important}
[data-testid="stSidebar"] .stButton>button:hover{background:#ecece7!important;color:#111!important;transform:none!important;box-shadow:none!important}
.topbar{padding:4px 0 28px;border-bottom:1px solid #1112;margin-bottom:24px}.eyebrow{font-size:11px;color:#6b7280;letter-spacing:2.2px;text-transform:uppercase;margin-bottom:9px}.topbar h1{font-family:'Noto Serif SC',serif;font-size:38px;line-height:1.2;font-weight:900;letter-spacing:-1.5px;margin:0 0 10px;color:#111}.topbar p{font-size:13px;color:#6b7280;margin:0}.steps{display:flex;align-items:center;gap:0;margin-top:28px}.step{display:flex;align-items:center;gap:8px;font-size:12px;color:#9ca3af;white-space:nowrap}.step b{width:25px;height:25px;display:grid;place-items:center;border-radius:50%;border:1px solid #9ca3af;font-size:9px}.step.done{color:#6b7280}.step.on{color:#111;font-weight:700}.step.on b{background:var(--lime);border-color:#111;color:#111}.step-line{width:48px;height:1px;background:#1113;margin:0 14px}
.upload-head{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:13px;border-bottom:1px solid #1112;margin-bottom:18px}.upload-head h2{font-family:'Noto Serif SC',serif;font-size:21px;margin:0 0 5px;color:#111;font-weight:700}.upload-head h2:before{content:'01';font-family:'Inter',sans-serif;color:#9ca3af;font-size:12px;margin-right:10px}.upload-head p{margin:0;font-size:11px;color:#6b7280}.upload-count{font-size:10px;color:#111;padding:5px 8px;background:#f3f4f6;border:1px solid #d1d5db}
[data-testid="stVerticalBlockBorderWrapper"]{border:1px solid #111!important;border-radius:0!important;padding:22px!important;background:#fff!important;box-shadow:none!important;margin-bottom:20px}[data-testid="stFileUploader"]{background:#fff;border:1px solid #111;border-radius:0;padding:13px 14px 7px;min-height:135px;transition:.18s}[data-testid="stFileUploader"]:hover{background:#e3f85a12}[data-testid="stFileUploader"] label{color:#111;font-weight:700;font-size:12px}[data-testid="stFileUploaderDropzone"]{min-height:64px;border:0;background:transparent;border-radius:0;padding:8px 0}[data-testid="stFileUploaderDropzone"] button{color:#111!important;background:var(--lime)!important;border:1px solid #111!important;border-radius:50%!important;width:35px!important;height:35px!important;min-width:35px!important;padding:0!important;font-size:0!important;box-shadow:none!important}[data-testid="stFileUploaderDropzone"] button:after{content:'＋';font-size:20px;color:#111;position:absolute;inset:0;display:grid;place-items:center}[data-testid="stFileUploaderDropzone"] small{font-size:9px;color:#9ca3af}
.file-status{margin:-6px 0 14px;padding:7px 9px;font-size:9px;letter-spacing:.4px;color:#9ca3af;background:#f7f7f5;border:1px solid #1111}.file-status.ok{color:#111;background:var(--success);border-color:#111}.panel{background:#fff;border:1px solid #111;padding:18px;margin-bottom:18px;box-shadow:none}.panel-title{display:flex;justify-content:space-between;align-items:center;font-size:13px;font-weight:700;padding-bottom:10px;border-bottom:1px solid #1112;margin-bottom:14px}.panel-title span{font-size:9px;font-weight:500;color:#9ca3af}.mini-metrics{display:flex;align-items:end;gap:12px;margin-bottom:14px}.mini-metric{border:0;background:transparent;padding:0}.mini-metric:first-child strong{font-family:'Noto Serif SC',serif;font-size:34px;font-weight:900}.mini-metric strong{display:block;font-size:15px}.mini-metric small{color:#9ca3af;font-size:9px}.progress-track{height:8px;background:#e5e7eb;overflow:hidden}.progress-track i{display:block;height:100%;background:var(--lime);border-right:1px solid #111}.task{display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid #1111;font-size:10px}.task:last-child{border:0}.task span{color:#9ca3af}.task .done{color:#111;background:var(--lime);border:1px solid #111;padding:1px 5px}
.timeline{border-left:1px solid #d1d5db;margin:5px 0 0 7px;padding-left:16px}.event{position:relative;margin:0 0 17px}.event:last-child{margin-bottom:0}.event:before{content:'';position:absolute;left:-21px;top:4px;width:8px;height:8px;border-radius:50%;background:#fff;border:1px solid #9ca3af}.event:first-child:before{background:var(--lime);border-color:#111}.event b{font-size:10px;display:block;margin-bottom:3px}.event span{font-size:9px;color:#9ca3af}
[data-testid="stExpander"]{border:1px solid #111!important;border-radius:0!important;background:#fff;box-shadow:none}[data-testid="stExpander"] summary{font-family:'Noto Serif SC',serif;font-weight:700}.stNumberInput input,.stDateInput input{border-radius:0!important;background:#fafafa!important}.stButton>button,.stDownloadButton>button{border-radius:0;min-height:43px;font-size:11px;font-weight:700;border:1px solid #111}.stButton>button[kind="primary"],.stDownloadButton>button[kind="primary"]{background:var(--lime);color:#111;border-color:#111;box-shadow:3px 3px 0 #111}.stButton>button[kind="primary"]:hover,.stDownloadButton>button[kind="primary"]:hover{transform:translate(-1px,-1px);box-shadow:4px 4px 0 #111}.stButton>button[kind="primary"]:disabled{background:#e5e7eb;border-color:#c7c7c7;color:#9ca3af;box-shadow:none}
.download-card{background:#fff;border:1px solid #111;padding:14px;margin:5px 0 8px}.download-card b{font-family:'Noto Serif SC',serif;font-size:12px}.download-card span{color:#9ca3af;font-size:9px}.section-label{font-family:'Noto Serif SC',serif;font-size:18px;font-weight:700;margin:28px 0 5px}.section-note{color:#6b7280;font-size:10px;margin-bottom:12px}.fixed-footer{height:52px;margin-top:28px;border-top:1px solid #1112;display:flex;align-items:center;justify-content:space-between}.fixed-footer span,.fixed-footer b{font-size:9px;color:#6b7280}footer{visibility:hidden}
@media(max-width:900px){[data-testid="stSidebar"]{display:none}.block-container{padding:20px 15px 55px}.topbar h1{font-size:30px}.step-line{width:13px;margin:0 6px}.steps{overflow-x:auto}[data-testid="stVerticalBlockBorderWrapper"]{padding:14px!important}}

/* 可读性修正版：收紧超宽屏、放大文字、修复上传按钮重叠 */
html{font-size:17px}
.block-container{max-width:1280px!important;padding:42px 42px 90px!important;margin:0 auto!important}
[data-testid="stSidebar"]{width:320px!important;min-width:320px!important}
[data-testid="stSidebar"]>div:first-child{width:320px!important;padding:48px 30px!important}
.brand-mark{width:44px;height:44px;font-size:19px;margin-bottom:42px}
.nav-link{width:44px;height:44px;font-size:20px;margin:16px auto}
.nav-space{height:270px}.nav-hint{font-size:10px;color:#71717a;line-height:1.45}
.topbar{padding-bottom:32px;margin-bottom:28px}.eyebrow{font-size:13px;margin-bottom:12px}
.topbar h1{font-size:46px;margin-bottom:14px}.topbar p{font-size:15px}
.steps{margin-top:32px}.step{font-size:14px;gap:9px}.step b{width:29px;height:29px;font-size:11px}.step-line{width:52px;margin:0 15px}
.upload-head{padding-bottom:16px;margin-bottom:20px}.upload-head h2{font-size:25px}.upload-head h2:before{font-size:14px}.upload-head p{font-size:13px}.upload-count{font-size:12px;padding:7px 10px}
[data-testid="stVerticalBlockBorderWrapper"]{padding:24px!important;margin-bottom:22px!important}
[data-testid="stFileUploader"]{padding:16px 16px 10px!important;min-height:150px!important}
[data-testid="stFileUploader"] label{font-size:15px!important;line-height:1.45!important}
[data-testid="stFileUploaderDropzone"]{min-height:74px!important;padding:14px 0!important;display:flex!important;align-items:center!important;justify-content:flex-end!important}
[data-testid="stFileUploaderDropzoneInstructions"]{display:none!important}
[data-testid="stFileUploaderDropzone"] svg{display:none!important}
[data-testid="stFileUploaderDropzone"] small{display:none!important}
[data-testid="stFileUploaderDropzone"] button{color:transparent!important;background:var(--lime)!important;border:1px solid #111!important;border-radius:0!important;width:118px!important;min-width:118px!important;height:42px!important;padding:0!important;font-size:0!important;box-shadow:2px 2px 0 #111!important;position:relative!important}
[data-testid="stFileUploaderDropzone"] button:after{content:'上传文件'!important;font-size:14px!important;font-weight:700!important;color:#111!important;position:absolute!important;inset:0!important;display:grid!important;place-items:center!important}
.file-status{padding:9px 11px;font-size:11px;margin-bottom:16px}
.panel{padding:21px;margin-bottom:20px}.panel-title{font-size:16px;padding-bottom:12px;margin-bottom:16px}.panel-title span{font-size:11px;color:#71717a}
.mini-metrics{gap:16px;margin-bottom:16px}.mini-metric:first-child strong{font-size:40px}.mini-metric strong{font-size:18px}.mini-metric small{font-size:11px;color:#71717a}
.task{padding:10px 0;font-size:12px}.task .done{padding:2px 6px}.event b{font-size:13px}.event span{font-size:11px;color:#71717a}
.section-label{font-size:23px;margin-top:32px}.section-note{font-size:12px;margin-bottom:14px}
.download-card{padding:16px}.download-card b{font-size:15px}.download-card span{font-size:11px}
.fixed-footer span,.fixed-footer b{font-size:11px}
.stButton>button,.stDownloadButton>button{font-size:13px;min-height:45px}
@media(min-width:1600px){.block-container{max-width:1340px!important}.topbar h1{font-size:50px}}
@media(max-width:900px){html{font-size:16px}.block-container{padding:22px 16px 60px!important}.topbar h1{font-size:34px}.topbar p{font-size:13px}.step{font-size:12px}.step-line{width:14px;margin:0 6px}[data-testid="stVerticalBlockBorderWrapper"]{padding:15px!important}[data-testid="stFileUploader"]{min-height:132px!important}.panel{padding:16px}}

/* 上传卡片：整卡展示信息，右侧圆形加号触发文件选择 */
[data-testid="stFileUploader"]{position:relative!important;min-height:210px!important;padding:32px 34px!important;overflow:visible!important}
[data-testid="stFileUploader"] label{max-width:72%!important;margin:0!important;padding:0!important}
[data-testid="stFileUploader"] label p{font-size:16px!important;line-height:1.65!important;color:#6b7280!important;font-weight:500!important;margin:0!important}
[data-testid="stFileUploader"] label strong{display:block!important;font-family:'Noto Serif SC',serif!important;font-size:23px!important;line-height:1.35!important;color:#111!important;font-weight:900!important;margin-bottom:7px!important}
[data-testid="stFileUploader"]:after{content:'XLSX, XLSM, CSV · MAX 1GB';position:absolute;left:34px;bottom:30px;color:#9ca3af;font-size:13px;font-weight:600;letter-spacing:.3px}
[data-testid="stFileUploaderDropzone"]{position:absolute!important;right:30px!important;top:31px!important;width:58px!important;height:58px!important;min-height:58px!important;padding:0!important;display:block!important;background:transparent!important;overflow:visible!important}
[data-testid="stFileUploaderDropzone"] button{display:block!important;position:absolute!important;inset:0!important;width:58px!important;min-width:58px!important;height:58px!important;border-radius:50%!important;background:#fff!important;border:1.5px solid #111!important;box-shadow:none!important;padding:0!important}
[data-testid="stFileUploaderDropzone"] button:after{content:'＋'!important;font-family:'Inter',sans-serif!important;font-size:34px!important;line-height:1!important;font-weight:400!important;color:#111!important;position:absolute!important;inset:0!important;display:grid!important;place-items:center!important}
[data-testid="stFileUploader"]:hover [data-testid="stFileUploaderDropzone"] button{background:var(--lime)!important}
/* 上传后的原生文件名/删除按钮会压到右侧加号；状态由下方绿色提示条统一展示 */
[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"]{display:none!important}
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileData"]{display:none!important}
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"]{display:none!important}
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileSize"]{display:none!important}
[data-testid="stFileUploader"] [data-testid*="FileUploaderFile"]{display:none!important}
[data-testid="stFileUploader"] [data-testid*="UploadedFile"]{display:none!important}
[data-testid="stFileUploader"] [data-testid*="DeleteBtn"]{display:none!important}
[data-testid="stFileUploader"] [role="list"]{display:none!important}
[data-testid="stFileUploader"] ul{display:none!important}
[data-testid="stFileUploaderDropzone"]{z-index:20!important}
[data-testid="stFileUploaderDropzone"] button{z-index:21!important}
.file-status:not(.ok){display:none!important}.file-status.ok{margin:-8px 0 16px!important;font-size:12px!important;line-height:1.45!important;overflow-wrap:anywhere!important}.file-status.ok span{display:block;margin-top:3px;color:#46501f;font-size:10px;font-weight:500}
.app-shell{background:#fff;border:1px solid #111;padding:34px;margin-top:8px}.app-shell h2{font-family:'Noto Serif SC',serif;font-size:28px;margin:0 0 8px}.app-shell>p{color:#6b7280;font-size:14px;margin:0 0 28px}.app-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px}.app-card{min-height:150px;border:1px solid #111;padding:24px;position:relative;background:#fafafa}.app-card b{display:block;font-family:'Noto Serif SC',serif;font-size:20px;margin-bottom:10px}.app-card span{font-size:12px;color:#6b7280}.app-card i{position:absolute;right:18px;top:18px;font-style:normal;font-size:10px;border:1px solid #111;padding:4px 7px;background:var(--lime)}
@media(max-width:900px){[data-testid="stFileUploader"]{min-height:180px!important;padding:25px 22px!important}[data-testid="stFileUploader"] label strong{font-size:19px!important}[data-testid="stFileUploader"]:after{left:22px;bottom:24px;font-size:11px}[data-testid="stFileUploaderDropzone"]{right:20px!important;top:23px!important;width:50px!important;height:50px!important;min-height:50px!important}[data-testid="stFileUploaderDropzone"] button{width:50px!important;min-width:50px!important;height:50px!important}}

/* 2026-11 UI定稿：按Web/App原型统一尺寸，放在最后避免被旧规则覆盖 */
html{font-size:15px!important}
.stApp{background:#f6f5f0!important}
.block-container{max-width:1200px!important;padding:48px 32px 300px!important;margin:0 auto!important}
[data-testid="stSidebar"]{width:224px!important;min-width:224px!important;background:#f6f5f0!important}
[data-testid="stSidebar"]>div:first-child{width:224px!important;padding:40px 20px!important}
[data-testid="stSidebar"] .sidebar-section-title{font-family:'Inter','PingFang SC',sans-serif!important;font-size:11px!important;font-weight:800!important;letter-spacing:1.2px!important;color:#9ca3af!important;margin:0 8px 16px!important}
[data-testid="stSidebar"] .stButton{margin:0 0 12px!important}
[data-testid="stSidebar"] .stButton>button{height:52px!important;min-height:52px!important;padding:0 16px!important;border-radius:8px!important;font-size:15px!important;font-weight:700!important}
[data-testid="stSidebar"] .stButton>button p{font-size:15px!important;font-weight:700!important}

.topbar{padding:0 0 24px!important;margin:0 0 26px!important}
.eyebrow{font-size:10px!important;letter-spacing:2px!important;margin-bottom:8px!important}
.topbar h1{font-size:36px!important;line-height:1.25!important;margin:0 0 8px!important;letter-spacing:-1px!important}
.topbar p{font-size:13px!important;margin:0!important}
.steps{margin-top:26px!important}.step{font-size:12px!important;gap:7px!important}.step b{width:22px!important;height:22px!important;font-size:9px!important}.step-line{width:34px!important;margin:0 10px!important}

[data-testid="stVerticalBlockBorderWrapper"]{padding:24px!important;margin-bottom:22px!important;border:1px solid #111!important;box-shadow:4px 4px 0 #111!important}
.upload-head{padding-bottom:15px!important;margin-bottom:20px!important}.upload-head h2{font-size:22px!important}.upload-head h2:before{font-size:13px!important}.upload-head p{font-size:11px!important}.upload-count{font-size:11px!important;padding:6px 9px!important}
[data-testid="stFileUploader"]{min-height:164px!important;padding:26px 28px!important}
[data-testid="stFileUploader"] label{max-width:72%!important}
[data-testid="stFileUploader"] label p{font-size:13px!important;line-height:1.5!important}
[data-testid="stFileUploader"] label strong{font-size:20px!important;line-height:1.35!important;margin-bottom:6px!important}
[data-testid="stFileUploader"]:after{left:28px!important;bottom:25px!important;font-size:10px!important;letter-spacing:1px!important}
[data-testid="stFileUploaderDropzone"]{right:25px!important;top:24px!important;width:44px!important;height:44px!important;min-height:44px!important}
[data-testid="stFileUploaderDropzone"] button{width:44px!important;min-width:44px!important;height:44px!important;border-width:1.5px!important}
[data-testid="stFileUploaderDropzone"] button:after{font-size:27px!important}
.file-status.ok{margin:-8px 0 14px!important;padding:7px 9px!important;font-size:10px!important}.file-status.ok span{font-size:9px!important}

.section-label{font-size:21px!important;margin:28px 0 4px!important}.section-note{font-size:11px!important;margin-bottom:14px!important}
[data-testid="stExpander"]{margin-bottom:12px!important}[data-testid="stExpander"] summary{font-family:'Inter','PingFang SC',sans-serif!important;font-size:13px!important;padding:14px 18px!important}
.panel{padding:20px!important;margin-bottom:22px!important;box-shadow:3px 3px 0 #111!important}.panel-title{font-size:14px!important}.panel-title span{font-size:10px!important}
.mini-metric:first-child strong{font-size:38px!important}.mini-metric strong{font-size:16px!important}.mini-metric small{font-size:10px!important}.task{font-size:11px!important}.event b{font-size:11px!important}.event span{font-size:10px!important}
.download-card{padding:16px!important}.download-card b{font-size:14px!important}.download-card span{font-size:10px!important}

/* 上传后的删除入口保留，但降级为轻量操作，避免破坏卡片网格 */
[class*="st-key-clear_"] .stButton>button{min-height:30px!important;height:30px!important;background:transparent!important;border:0!important;box-shadow:none!important;color:#6b7280!important;font-size:10px!important;padding:0 8px!important}
[class*="st-key-clear_"] .stButton>button:hover{color:#111!important;background:#ecebe5!important;transform:none!important;box-shadow:none!important}

/* 固定底部操作栏；使用真实占位元素扩展滚动区，避免遮住目标、生成和下载区域 */
.fixed-footer{position:fixed!important;left:224px!important;right:0!important;bottom:0!important;height:72px!important;margin:0!important;padding:0 max(32px,calc((100vw - 224px - 1200px)/2 + 32px))!important;background:rgba(246,245,240,.96)!important;backdrop-filter:blur(10px)!important;border-top:1px solid #1112!important;z-index:998!important;display:flex!important;align-items:center!important}
.fixed-footer span{font-size:11px!important;color:#6b7280!important}.fixed-footer b{display:none!important}
.fixed-footer a{color:#111!important;font-weight:700!important;text-decoration:underline!important;margin-left:12px!important}
.st-key-save_web_manual,.st-key-save_app_manual{position:fixed!important;right:max(32px,calc((100vw - 224px - 1200px)/2 + 32px))!important;bottom:15px!important;width:255px!important;z-index:1000!important}
.st-key-save_web_manual .stButton>button,.st-key-save_app_manual .stButton>button{height:42px!important;min-height:42px!important;background:#111!important;color:#e3f85a!important;border:1px solid #111!important;box-shadow:3px 3px 0 #777!important;font-size:13px!important}
.st-key-save_web_manual .stButton>button:before,.st-key-save_app_manual .stButton>button:before{content:'⇩';font-size:16px;margin-right:8px}
.inline-save-status{position:fixed;right:max(307px,calc((100vw - 224px - 1200px)/2 + 307px));bottom:15px;width:390px;height:42px;z-index:1000;display:flex;align-items:center;padding:0 14px;border:1px solid #badb88;background:#edf9e9;color:#245c47;font-size:12px;line-height:1.35}
.target-anchor{height:1px;scroll-margin-top:90px}
.page-bottom-spacer{display:block!important;width:100%!important;height:240px!important;min-height:240px!important;clear:both!important;pointer-events:none!important}
/* 右侧生成按钮随页面保持在固定栏上方，长表单下也能随时操作 */
.st-key-web_generate,.st-key-app_generate{position:sticky!important;bottom:88px!important;z-index:900!important;background:#f6f5f0!important;padding:8px 0 3px!important}

@media(max-width:900px){
  [data-testid="stSidebar"]{display:none!important}.block-container{max-width:100%!important;padding:24px 16px 240px!important}.topbar h1{font-size:31px!important}.steps{overflow-x:auto!important}
  [data-testid="stVerticalBlockBorderWrapper"]{padding:15px!important;box-shadow:3px 3px 0 #111!important}[data-testid="stFileUploader"]{min-height:150px!important;padding:22px!important}
  .fixed-footer{left:0!important;padding:0 16px!important}.fixed-footer span{display:none!important}.st-key-save_web_manual,.st-key-save_app_manual{right:16px!important;width:210px!important}.inline-save-status{display:none!important}.st-key-web_generate,.st-key-app_generate{bottom:84px!important}.page-bottom-spacer{height:210px!important;min-height:210px!important}
}
</style>
""", unsafe_allow_html=True)

if not admin_mode:
    st.markdown("<style>.st-key-admin_settings{display:none!important}</style>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown(
        '<div class="sidebar-section-title">生成模块</div>',
        unsafe_allow_html=True,
    )
    if st.button("🖥️  Web 周报", key="port_web", help="进入Web周报生成", use_container_width=True):
        st.query_params["port"] = "web"
        st.rerun()
    if st.button("📱  App 周报", key="port_app", help="进入App周报生成", use_container_width=True):
        st.query_params["port"] = "app"
        st.rerun()

active_button_key = "port_web" if active_port == "web" else "port_app"
st.markdown(
    f"<style>.st-key-{active_button_key} button{{background:var(--lime)!important;"
    "color:#111!important;border-color:#111!important;box-shadow:3px 3px 0 #111!important;"
    "transform:translate(-1px,-1px)!important}}"
    f".st-key-{active_button_key} button:after{{content:'';position:absolute;left:-1px;top:22%;"
    "bottom:22%;width:4px;background:#111;border-radius:0 4px 4px 0}}</style>",
    unsafe_allow_html=True,
)

if active_port == "app":
    st.markdown(
        """
        <header class="topbar">
          <div class="eyebrow">APP Report Workbench</div>
          <h1>APP数据自动化</h1>
          <p>上传Google、Meta与ASA消耗，补充投放平台收入，生成APP投放营收表</p>
          <div class="steps">
            <div class="step done"><b>✓</b>选择周期</div><div class="step-line"></div>
            <div class="step on"><b>02</b>上传消耗</div><div class="step-line"></div>
            <div class="step"><b>03</b>补充收入</div><div class="step-line"></div>
            <div class="step"><b>04</b>生成表格</div>
          </div>
        </header>
        """,
        unsafe_allow_html=True,
    )

    today = date.today()
    app_default_end = today - timedelta(days=1)
    app_default_start = app_default_end - timedelta(days=6)
    app_main, app_side = st.columns([3.15, 1], gap="large")

    def app_upload_status(files):
        if not files:
            return '<div class="file-status">待上传</div>'
        names = "、".join(escape(file.name) for file in files)
        return (
            f'<div class="file-status ok">✓ 已上传 {len(files)} 个文件 · 等待格式检查'
            f'<span>{names}</span></div>'
        )

    with app_main:
        with st.container(border=True, key="app_upload_box"):
            app_upload_header = st.empty()
            app_u1, app_u2, app_u3 = st.columns(3)
            with app_u1:
                app_google_upload_version = st.session_state.get("app_google_upload_version", 0)
                app_google_files = st.file_uploader(
                    "**Google消耗**  \nGoogle Ads 投放消耗",
                    type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
                    key=f"app_google_files_{app_google_upload_version}",
                )
                st.markdown(app_upload_status(app_google_files), unsafe_allow_html=True)
                if app_google_files and st.button(
                    "删除已上传的Google文件", key="clear_app_google", use_container_width=True
                ):
                    st.session_state["app_google_upload_version"] = app_google_upload_version + 1
                    st.session_state.pop("app_report", None)
                    st.rerun()
            with app_u2:
                app_meta_upload_version = st.session_state.get("app_meta_upload_version", 0)
                app_meta_files = st.file_uploader(
                    "**Meta消耗**  \nMeta Ads 投放消耗",
                    type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
                    key=f"app_meta_files_{app_meta_upload_version}",
                )
                st.markdown(app_upload_status(app_meta_files), unsafe_allow_html=True)
                if app_meta_files and st.button(
                    "删除已上传的Meta文件", key="clear_app_meta", use_container_width=True
                ):
                    st.session_state["app_meta_upload_version"] = app_meta_upload_version + 1
                    st.session_state.pop("app_report", None)
                    st.rerun()
            with app_u3:
                app_asa_upload_version = st.session_state.get("app_asa_upload_version", 0)
                app_asa_files = st.file_uploader(
                    "**ASA消耗**  \nASA 分地区分天消耗",
                    type=["xlsx", "xlsm", "csv"], accept_multiple_files=True,
                    key=f"app_asa_files_{app_asa_upload_version}",
                )
                st.markdown(app_upload_status(app_asa_files), unsafe_allow_html=True)
                if app_asa_files and st.button(
                    "删除已上传的ASA文件", key="clear_app_asa", use_container_width=True
                ):
                    st.session_state["app_asa_upload_version"] = app_asa_upload_version + 1
                    st.session_state.pop("app_report", None)
                    st.rerun()
            app_uploaded_count = sum(
                bool(value) for value in (app_google_files, app_meta_files, app_asa_files)
            )
            shared_order_files = st.session_state.get("shared_order_files", [])
            app_upload_header.markdown(
                f'<div class="upload-head"><div><h2>上传消耗数据</h2>'
                f'<p>按日期和地区自动汇总</p></div>'
                f'<div class="upload-count">{app_uploaded_count} / 3 已完成</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown(
            '<div class="section-label">02 补充手动数据</div>'
            '<div class="section-note">投放平台收入暂无自动数据源，按本周、上周、MTD填写。</div>',
            unsafe_allow_html=True,
        )
        app_manual = {}
        app_saved_manual = st.session_state.get("app_manual_saved", {})
        app_periods = (("本周", "本周"), ("上周", "上周"), ("MTD", "累计 / MTD"))

        with st.expander("投放平台收入", expanded=True):
            for prefix, title in (
                ("google_revenue", "Google收入（D7）"),
                ("meta_revenue", "Meta收入（D7）"),
                ("asa_revenue", "ASA收入（D7）"),
            ):
                st.markdown(f"**{title}**")
                columns = st.columns(3)
                for column, (period, label) in zip(columns, app_periods):
                    manual_key = f"{prefix}_{period}"
                    app_manual[manual_key] = column.number_input(
                        label, min_value=0.0, value=float(app_saved_manual.get(manual_key, 0.0)),
                        key=f"app_{manual_key}",
                    )

        with st.expander("地区买量收入（手动填写）", expanded=False):
            st.caption("填写台湾、韩国、印度、印度尼西亚的Google + Meta + ASA收入合计；其他地区由总收入减去这四个地区自动得到。")
            for region in ("台湾", "韩国", "印度", "印度尼西亚"):
                st.markdown(f"**{region}收入**")
                columns = st.columns(3)
                for column, (period, label) in zip(columns, app_periods):
                    manual_key = f"region_revenue_{region}_{period}"
                    app_manual[manual_key] = column.number_input(
                        label, min_value=0.0, value=float(app_saved_manual.get(manual_key, 0.0)),
                        key=f"app_region_revenue_{region}_{period}",
                    )

        with st.expander("APP总数据补充", expanded=False):
            st.caption(
                "APP自然量收入按本周、上周和MTD手动填写；总收入、新用户收入和老用户收入"
                "会自动读取Web页面上传的订单表。"
            )
            st.markdown("**APP自然量收入**")
            columns = st.columns(3)
            for column, (period, label) in zip(columns, app_periods):
                manual_key = f"app_natural_revenue_{period}"
                app_manual[manual_key] = column.number_input(
                    label, min_value=0.0,
                    value=float(app_saved_manual.get(manual_key, 0.0)),
                    key=f"app_{manual_key}",
                )

        app_target_defaults = {
            "APP—新老综合ROI": 1.5,
            "APP—新用户综合ROI（D0）": 1.2,
            "APP—投放新用户ROI（D7）": 1.2,
            "APP—应用商店（ASA）新用户ROI（D7）": 0.75,
            "APP—其他地区新老综合ROI": 1.2,
            "APP—英语地区新老综合ROI": 1.2,
            "总收入": 15000.0,
            "APP投放收入": 2000.0,
            "APP自然量收入": 9000.0,
            "老用户收入": 10000.0,
            "总成本": 7500.0,
            "Google—投放新用户 ROI（D7）": 1.2,
            "Meta—投放新用户 ROI（D7）": 1.2,
            "ASA—新用户ROI": 1.2,
            "APP投放总收入（D7）": 4000.0,
            "Google收入（D7）": 2000.0,
            "Meta收入（D7）": 2000.0,
            "ASA收入（D7）": 2000.0,
            "地区收入_台湾": 1000.0,
            "地区收入_韩国": 500.0,
            "地区收入_印度": 500.0,
            "地区收入_印度尼西亚": 500.0,
            "地区收入_其他": 500.0,
            "地区ROI_台湾": 1.0,
            "地区ROI_韩国": 1.0,
            "地区ROI_印度": 1.0,
            "地区ROI_印度尼西亚": 1.0,
        }
        st.markdown('<div id="app-targets" class="target-anchor"></div>', unsafe_allow_html=True)
        with st.expander("APP月度目标", expanded=False):
            st.caption(
                "可以从Excel复制一整列目标，粘贴到下面的批量框后应用；"
                "也可以选中目标列第一格直接粘贴。ROI目标支持120%或1.2两种格式。"
            )
            pasted_app_targets = st.text_area(
                "批量粘贴目标（严格按照下方指标顺序，一行一个）",
                key="app_targets_paste_text", height=110,
            )
            if st.button("应用整列目标", key="apply_app_targets_paste"):
                lines = [line.strip() for line in pasted_app_targets.splitlines() if line.strip()]
                if len(lines) != len(app_target_defaults):
                    st.error(f"需要粘贴 {len(app_target_defaults)} 个目标，当前识别到 {len(lines)} 个。")
                else:
                    parsed = {}
                    try:
                        for key, raw in zip(app_target_defaults, lines):
                            value_text = raw.split("\t")[-1].replace(",", "").replace("$", "").strip()
                            is_percent = value_text.endswith("%")
                            value = float(value_text.rstrip("%"))
                            parsed[key] = value / 100 if is_percent else value
                    except ValueError:
                        st.error("目标列中存在无法识别的内容，请只粘贴数字、金额或百分比。")
                    else:
                        st.session_state["app_pasted_targets"] = parsed
                        st.session_state["app_target_editor_version"] = (
                            st.session_state.get("app_target_editor_version", 0) + 1
                        )
                        st.rerun()
            app_target_values = {
                **app_target_defaults,
                **st.session_state.get("app_pasted_targets", {}),
            }
            target_frame = pd.DataFrame(
                [{"指标": key, "目标": value} for key, value in app_target_values.items()]
            )
            edited_targets = st.data_editor(
                target_frame,
                hide_index=True,
                disabled=["指标"],
                use_container_width=True,
                key=f"app_targets_editor_{st.session_state.get('app_target_editor_version', 0)}",
                column_config={"目标": st.column_config.NumberColumn("目标", min_value=0.0, format="%.2f")},
            )
            app_targets = {
                str(row["指标"]): float(row["目标"] or 0.0)
                for row in edited_targets.to_dict("records")
            }

        # 所有手动数据与月度目标填写完成后，再统一暂存。
        app_save_col, app_save_status_col = st.columns([1, 2.2], gap="small")
        with app_save_col:
            app_saved_now = st.button(
                "暂存当前手动数据", key="save_app_manual", use_container_width=True
            )
        if app_saved_now:
            st.session_state["app_manual_saved"] = deepcopy(app_manual)
        with app_save_status_col:
            if app_saved_now:
                st.markdown(
                    '<div class="inline-save-status">✓ 已暂存到当前浏览器会话；重新上传文件不会丢失。</div>',
                    unsafe_allow_html=True,
                )
            elif st.session_state.get("app_manual_saved"):
                st.markdown(
                    '<div class="inline-save-status">✓ 当前手动数据已暂存。</div>',
                    unsafe_allow_html=True,
                )

    with app_side:
        st.markdown('<div class="panel"><div class="panel-title">APP报告周期<span>自动保存</span></div>', unsafe_allow_html=True)
        app_start = st.date_input("开始日期", app_default_start, key="app_report_start")
        app_end = st.date_input("结束日期", app_default_end, key="app_report_end")
        st.caption(f"当前周期 · {app_start:%Y.%m.%d} – {app_end:%Y.%m.%d}")
        st.markdown('</div>', unsafe_allow_html=True)

        app_completed = 1 + app_uploaded_count + int(bool(shared_order_files))
        app_readiness = int(app_completed / 5 * 100)
        app_tasks = (
            ("报告周期", True),
            ("Web订单数据", bool(shared_order_files)),
            ("Google消耗", bool(app_google_files)),
            ("Meta消耗", bool(app_meta_files)),
            ("ASA消耗", bool(app_asa_files)),
        )
        app_tasks_html = "".join(
            f'<div class="task"><b>{"✓" if done else "○"} {name}</b>'
            f'<span class="{"done" if done else ""}">{"已完成" if done else "待上传"}</span></div>'
            for name, done in app_tasks
        )
        st.markdown(
            f'<div class="panel"><div class="panel-title">报告准备度<span>{app_readiness}%</span></div>'
            f'<div class="mini-metrics"><div class="mini-metric"><strong>{app_readiness}%</strong><small>当前进度</small></div>'
            f'<div class="mini-metric"><strong>{app_completed} / 5</strong><small>已完成项目</small></div></div>'
            f'<div class="progress-track"><i style="width:{app_readiness}%"></i></div><div style="height:10px"></div>'
            f'{app_tasks_html}</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="panel"><div class="panel-title">本次输出<span>APP工作簿</span></div>'
            '<div class="timeline"><div class="event"><b>自动汇总消耗</b><span>Google / Meta / ASA按周期和地区统计</span></div>'
            '<div class="event"><b>生成Excel</b><span>沿用APP投放表格格式</span></div></div></div>',
            unsafe_allow_html=True,
        )
        app_ready = all((shared_order_files, app_google_files, app_meta_files, app_asa_files)) and app_start <= app_end
        if not shared_order_files:
            app_button_label = "请先到Web页面上传订单数据"
        elif not all((app_google_files, app_meta_files, app_asa_files)):
            app_button_label = "请先上传Google、Meta和ASA消耗"
        else:
            app_button_label = "✦ 生成APP工作簿"
        app_generate = st.button(
            app_button_label,
            type="primary", disabled=not app_ready, use_container_width=True,
            key="app_generate",
        )

    if app_start > app_end:
        with app_side:
            st.error("开始日期不能晚于结束日期。")

    if app_generate:
        try:
            with app_main:
                with st.status("正在读取APP消耗并生成表格……", expanded=True) as status:
                    app_google = read_app_uploads(app_google_files, "google")
                    st.write(f"Google消耗：{len(app_google):,} 行")
                    app_meta = read_app_uploads(app_meta_files, "meta")
                    st.write(f"Meta消耗：{len(app_meta):,} 行")
                    app_asa = read_app_uploads(app_asa_files, "asa")
                    st.write(f"ASA消耗：{len(app_asa):,} 行")
                    app_orders = read_uploads(shared_order_files)
                    st.write(f"Web订单数据：{len(app_orders):,} 行")
                    filename, content, warnings, costs = generate_app_report(
                        app_google, app_meta, app_asa, app_orders,
                        app_start, app_end, app_manual, app_targets
                    )
                    st.session_state["app_report"] = (filename, content)
                    st.session_state["app_report_warnings"] = warnings
                    st.session_state["app_cost_summary"] = costs
                    status.update(label="APP工作簿已生成", state="complete", expanded=False)
        except Exception as exc:
            with app_main:
                st.exception(exc)

    if st.session_state.get("app_report"):
        filename, content = st.session_state["app_report"]
        st.markdown(
            "<style>.st-key-save_app_manual,.inline-save-status,.fixed-footer,.page-bottom-spacer{display:none!important}</style>",
            unsafe_allow_html=True,
        )
        with app_main:
            for warning in st.session_state.get("app_report_warnings", []):
                st.warning(warning)
            costs = st.session_state.get("app_cost_summary", {})
            if costs:
                st.markdown('<div class="section-label">03 自动汇总消耗</div>', unsafe_allow_html=True)
                cost_table = pd.DataFrame(costs).T.reset_index(names="周期")
                st.dataframe(cost_table, hide_index=True, use_container_width=True)
            st.markdown(
                '<div class="section-label">04 APP表格已就绪</div>'
                '<div class="section-note">工作簿包含APP投放表、APP总数据表和APP分国家表。</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                f'<div class="download-card"><b>APP投放、总数据与分国家数据</b><br><span>Excel · {len(content)/1024:.1f} KB</span></div>',
                unsafe_allow_html=True,
            )
            st.download_button(
                "下载APP工作簿", content, filename,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True, key="download_app_report",
            )

    st.markdown('<div class="page-bottom-spacer" aria-hidden="true"></div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="fixed-footer"><span>APP与WEB数据相互独立'
        '<a href="#app-targets">跳到APP月度目标</a></span>'
        '<b>DATA AUTOMATION · APP工作台</b></div>',
        unsafe_allow_html=True,
    )
    st.stop()

st.markdown("""<header class="topbar"><div class="eyebrow">WEB Report Workbench</div><h1>WEB数据自动化</h1><p>上传WEB数据，自动完成校验、归因计算并生成五份可编辑周报</p><div class="steps"><div class="step done"><b>✓</b>选择周期</div><div class="step-line"></div><div class="step on"><b>02</b>上传数据</div><div class="step-line"></div><div class="step"><b>03</b>补充信息</div><div class="step-line"></div><div class="step"><b>04</b>生成报告</div></div></header><div id="upload-data"></div>""", unsafe_allow_html=True)

today = date.today()
default_end = today - timedelta(days=1)
default_start = default_end - timedelta(days=6)
target_month_end = st.session_state.get("report_end", default_end)

main_col, side_col = st.columns([3.15, 1], gap="large")

def upload_status(files):
    if not files:
        return '<div class="file-status">待上传</div>'
    names = "、".join(escape(file.name) for file in files)
    return (
        f'<div class="file-status ok">✓ 已上传 {len(files)} 个文件 · 等待格式检查'
        f'<span>{names}</span></div>'
    )

with main_col:
    upload_container = st.container(border=True, key="web_upload_box")
    with upload_container:
        upload_header = st.empty()
        u1, u2 = st.columns(2)
        with u1:
            reg_upload_version = st.session_state.get("reg_upload_version", 0)
            reg_files = st.file_uploader(
                "**注册数据**  \n用户注册与渠道来源", type=["xlsx","xlsm","csv"],
                accept_multiple_files=True, key=f"reg_files_{reg_upload_version}"
            )
            st.markdown(upload_status(reg_files), unsafe_allow_html=True)
            if reg_files and st.button("删除已上传的注册文件", key="clear_reg_files", use_container_width=True):
                st.session_state["reg_upload_version"] = reg_upload_version + 1
                st.session_state.pop("reports", None)
                st.rerun()
            google_upload_version = st.session_state.get("google_upload_version", 0)
            google_files = st.file_uploader(
                "**Google消耗**  \nGoogle Ads 投放消耗", type=["xlsx","xlsm","csv"],
                accept_multiple_files=True, key=f"google_files_{google_upload_version}"
            )
            st.markdown(upload_status(google_files), unsafe_allow_html=True)
            if google_files and st.button("删除已上传的Google文件", key="clear_google_files", use_container_width=True):
                st.session_state["google_upload_version"] = google_upload_version + 1
                st.session_state.pop("reports", None)
                st.rerun()
        with u2:
            order_upload_version = st.session_state.get("order_upload_version", 0)
            order_files = st.file_uploader(
                "**订单数据**  \n订单及收入相关数据", type=["xlsx","xlsm","csv"],
                accept_multiple_files=True, key=f"order_files_{order_upload_version}"
            )
            if order_files:
                st.session_state["shared_order_files"] = list(order_files)
            st.markdown(upload_status(order_files), unsafe_allow_html=True)
            if order_files and st.button("删除已上传的订单文件", key="clear_order_files", use_container_width=True):
                st.session_state["order_upload_version"] = order_upload_version + 1
                st.session_state.pop("reports", None)
                st.session_state.pop("shared_order_files", None)
                st.rerun()
            bing_upload_version = st.session_state.get("bing_upload_version", 0)
            bing_files = st.file_uploader(
                "**Bing消耗**  \nBing Ads 投放消耗", type=["xlsx","xlsm","csv"],
                accept_multiple_files=True, key=f"bing_files_{bing_upload_version}"
            )
            st.markdown(upload_status(bing_files), unsafe_allow_html=True)
            if bing_files and st.button("删除已上传的Bing文件", key="clear_bing_files", use_container_width=True):
                st.session_state["bing_upload_version"] = bing_upload_version + 1
                st.session_state.pop("reports", None)
                st.rerun()
        uploaded_count = sum(bool(x) for x in [reg_files, order_files, google_files, bing_files])
        upload_header.markdown(f'<div class="upload-head"><div><h2>上传源数据</h2><p>上传4类数据，系统将自动完成格式检查和数据预处理</p></div><div class="upload-count">{uploaded_count} / 4 已完成</div></div>', unsafe_allow_html=True)

    st.markdown('<div id="manual-data"></div><div class="section-label">02 补充手动数据</div><div class="section-note">对应全渠道周报中的成本、流量和退款；没有数据时可保持为0。</div>', unsafe_allow_html=True)
    with st.expander("展开填写成本、流量和退款", expanded=False):
        input_values = {}
        web_saved_manual = st.session_state.get("web_manual_saved", {})
        periods = [("本周","本周"),("上周","上周"),("MTD","累计 / MTD")]
        groups = [("commission","affiliate分佣成本"),("seo","SEO—外链成本"),("pc_flow","PC流量"),("h5_flow","H5流量"),("refund","PC端退款金额")]
        for prefix, title in groups:
            st.markdown(f"**{title}**")
            cols = st.columns(3)
            for col, (key, label) in zip(cols, periods):
                manual_key = f"{prefix}_{key}"
                input_values[manual_key] = col.number_input(
                    label, min_value=0.0, value=float(web_saved_manual.get(manual_key, 0.0)),
                    key=manual_key,
                )

    with st.expander("自然来源周报补充数据", expanded=False):
        st.caption(
            "APP数据、达人数据、UGC费用与品牌词搜索量暂时没有数据库来源，按本周、上周和MTD手动填写。"
            "达人合作费 + UGC专项费用 + 投流费会作为KOL表成本总计，只加入全渠道ROI分母，不计入表内展示的总成本。"
        )
        natural_groups = [
            ("natural_creators", "已发布达人数量"),
            ("natural_exposure", "曝光量"),
            ("natural_brand_search", "WEB品牌词搜索量"),
            ("natural_app_reg", "APP自然量注册数"),
            ("natural_app_revenue", "APP自然量新用户收入"),
            ("natural_creator_cost", "达人合作费"),
            ("natural_ugc_cost", "UGC专项费用"),
            ("natural_boost_cost", "投流费"),
        ]
        for prefix, title in natural_groups:
            st.markdown(f"**{title}**")
            cols = st.columns(3)
            for col, (period_key, label) in zip(cols, periods):
                manual_key = f"{prefix}_{period_key}"
                input_values[manual_key] = col.number_input(
                    label, min_value=0.0, value=float(web_saved_manual.get(manual_key, 0.0)),
                    key=manual_key,
                )

    st.markdown('<div id="web-targets" class="target-anchor"></div>', unsafe_allow_html=True)
    with st.expander("月度目标设置（每月填写一次）", expanded=False):
        month_key = target_month_end.strftime("%Y-%m")
        all_targets = load_targets()
        month_targets = all_targets.get(month_key, {})
        st.caption(
            f"当前维护月份：{month_key}。金额、人数直接填数值；百分比目标按百分数填写，"
            "例如75%填写75。保存后所有用户默认使用这组目标。"
        )
        target_tab_all, target_tab_sem, target_tab_natural = st.tabs(
            ["全渠道目标", "SEM目标", "自然来源目标"]
        )

        def target_editor(tab, rows_config, values, key):
            with tab:
                metric_order = [metric_key for metric_key, _ in rows_config if metric_key is not None]
                st.caption(
                    f"共 {len(metric_order)} 个可填写目标；蓝灰色分区行只用于区分模块。"
                    "可从Excel复制一整列，粘贴到批量框后一次应用。"
                )
                paste_text = st.text_area(
                    "批量粘贴目标（严格按照下方指标顺序，一行一个）",
                    key=f"{key}_paste_text", height=95,
                )
                if st.button("应用整列目标", key=f"{key}_apply_paste"):
                    lines = [line.strip() for line in paste_text.splitlines() if line.strip()]
                    if len(lines) != len(metric_order):
                        st.error(f"需要粘贴 {len(metric_order)} 个目标，当前识别到 {len(lines)} 个。")
                    else:
                        parsed = {}
                        try:
                            for metric_key, raw in zip(metric_order, lines):
                                value_text = (
                                    raw.split("\t")[-1].replace(",", "").replace("$", "")
                                    .replace("%", "").strip()
                                )
                                parsed[metric_key] = float(value_text)
                        except ValueError:
                            st.error("目标列中存在无法识别的内容，请只粘贴数字、金额或百分比。")
                        else:
                            st.session_state[f"{key}_pasted_values"] = parsed
                            st.session_state[f"{key}_editor_version"] = (
                                st.session_state.get(f"{key}_editor_version", 0) + 1
                            )
                            st.rerun()
                editor_values = {
                    **values,
                    **st.session_state.get(f"{key}_pasted_values", {}),
                }
                edited_frames = []
                current_rows = []
                section_no = 0

                def render_metrics(metric_rows, editor_key):
                    if not metric_rows:
                        return
                    frame = pd.DataFrame([
                        {
                            "_目标键": metric_key,
                            "指标": label,
                            "目标": editor_values.get(metric_key),
                        }
                        for metric_key, label in metric_rows
                    ])
                    edited_frames.append(st.data_editor(
                        frame,
                        column_config={
                            "_目标键": None,
                            "指标": st.column_config.TextColumn("指标", disabled=True),
                            "目标": st.column_config.NumberColumn(
                                "目标", min_value=0.0, format="%.2f"
                            ),
                        },
                        disabled=["指标"],
                        hide_index=True,
                        use_container_width=True,
                        height=min(360, 38 + 36 * len(metric_rows)),
                        key=f"{editor_key}_{st.session_state.get(f'{key}_editor_version', 0)}",
                    ))

                for metric_key, label in rows_config:
                    if metric_key is None:
                        render_metrics(current_rows, f"{key}_section_{section_no}")
                        current_rows = []
                        section_no += 1
                        st.markdown(
                            "<div style='margin:12px 0 6px;padding:10px 14px;"
                            "background:#d8e8f5;border-left:5px solid #4b78c2;"
                            "font-weight:700;color:#244f75;'>"
                            f"{label}</div>",
                            unsafe_allow_html=True,
                        )
                    else:
                        current_rows.append((metric_key, label))
                render_metrics(current_rows, f"{key}_section_{section_no}")

                if not edited_frames:
                    return pd.DataFrame(columns=["_目标键", "指标", "目标"])
                return pd.concat(edited_frames, ignore_index=True)

        edited_all = target_editor(
            target_tab_all, ALL_CHANNEL_TARGET_ROWS,
            month_targets.get("all_channel", {}), f"all_targets_{month_key}"
        )
        edited_sem = target_editor(
            target_tab_sem, SEM_TARGET_ROWS,
            month_targets.get("sem", {}), f"sem_targets_{month_key}"
        )
        edited_natural = target_editor(
            target_tab_natural, NATURAL_TARGET_ROWS,
            month_targets.get("natural_source", {}), f"natural_targets_{month_key}"
        )

        if st.button("保存本月目标", type="primary", key="save_monthly_targets"):
            def clean(frame):
                result = {}
                for row in frame.to_dict("records"):
                    value = row.get("目标")
                    metric_key = str(row.get("_目标键", "")).strip()
                    if metric_key and value is not None and not pd.isna(value):
                        result[metric_key] = float(value)
                return result

            all_targets[month_key] = {
                "all_channel": clean(edited_all),
                "sem": clean(edited_sem),
                "natural_source": clean(edited_natural),
            }
            try:
                save_targets(all_targets, st.secrets)
                st.success(f"{month_key}目标已保存，之后会默认使用。")
            except Exception as exc:
                st.error(str(exc))

    # 所有Web手动数据与月度目标填写完成后，再统一暂存。
    web_save_col, web_save_status_col = st.columns([1, 2.2], gap="small")
    with web_save_col:
        web_saved_now = st.button(
            "暂存当前手动数据", key="save_web_manual", use_container_width=True
        )
    if web_saved_now:
        st.session_state["web_manual_saved"] = deepcopy(input_values)
    with web_save_status_col:
        if web_saved_now:
            st.markdown(
                '<div class="inline-save-status">✓ 已暂存到当前浏览器会话；重新上传文件不会丢失。</div>',
                unsafe_allow_html=True,
            )
        elif st.session_state.get("web_manual_saved"):
            st.markdown(
                '<div class="inline-save-status">✓ 当前手动数据已暂存。</div>',
                unsafe_allow_html=True,
            )

    admin_settings = st.container(key="admin_settings")
    with admin_settings.expander("统计规则设置（口径变更时修改）", expanded=False):
        rules = load_rules()
        st.caption(
            "这里修改的是团队共用统计口径。保存后将写入 GitHub；建议先写清变更说明，"
            "并用同一周期生成新旧报告核对。"
        )

        def csv_text(values):
            return "，".join(str(v) for v in (values or []))

        def csv_values(value, keep_blank=False):
            values = [v.strip() for v in str(value).replace("，", ",").split(",")]
            return values if keep_blank else [v for v in values if v]

        field_tabs = st.tabs(["字段对应", "用户与渠道", "消耗与地区"])
        field_frames = {}
        field_names = {
            "registration": "注册数据", "orders": "订单数据",
            "google": "Google消耗", "bing": "Bing消耗",
        }
        with field_tabs[0]:
            st.info("左列是系统标准字段，不要修改；右列填写上传文件中的真实表头。")
            for group, title in field_names.items():
                st.markdown(f"**{title}**")
                mapping = rules.get("fields", {}).get(group, {})
                frame = pd.DataFrame([
                    {"系统标准字段": canonical, "上传文件字段": source}
                    for canonical, source in mapping.items()
                ])
                field_frames[group] = st.data_editor(
                    frame,
                    disabled=["系统标准字段"], hide_index=True,
                    use_container_width=True, key=f"rule_fields_{group}",
                )

        with field_tabs[1]:
            new_user_field = st.text_input(
                "新用户判断字段", rules.get("new_user", {}).get("field", "注册转化周期")
            )
            new_user_value = st.number_input(
                "新用户判断值", value=float(rules.get("new_user", {}).get("value", 0)), step=1.0
            )
            subscription_values = st.text_input(
                "续费对应的末次来源类型（多个用逗号隔开）",
                csv_text(rules.get("subscription", {}).get("source_values", ["AUTO_RENEWAL"])),
            )
            priorities = rules.get("channels", {}).get("cooperation_priority", [])
            priority_map = {item.get("channel"): item.get("prefix", "") for item in priorities}
            kol_prefix = st.text_input("KOL合作渠道前缀", priority_map.get("KOL", "KOL"))
            affiliate_prefix = st.text_input("Affiliate合作渠道前缀", priority_map.get("Affiliate", "AFFILIATE"))
            source_rules = rules.get("channels", {}).get("blank_cooperation_sources", {})
            source_inputs = {}
            for channel in ["SEM", "KOL", "Affiliate", "Unassigned"]:
                source_inputs[channel] = st.text_input(
                    f"合作渠道为空时，{channel}对应末次来源类型",
                    csv_text(source_rules.get(channel, [])), key=f"source_rule_{channel}",
                )
            default_channel = st.selectbox(
                "其他未匹配来源默认归类",
                ["SEO", "KOL", "SEM", "Affiliate", "Unassigned", "Blank"],
                index=["SEO", "KOL", "SEM", "Affiliate", "Unassigned", "Blank"].index(
                    rules.get("channels", {}).get("default_channel", "SEO")
                ),
            )

        with field_tabs[2]:
            google_excludes = st.text_input(
                "Google广告系列排除词（多个用逗号隔开）",
                csv_text(rules.get("sem_cost", {}).get("google_exclude_contains", ["KOL", "generic"])),
            )
            region_rules = rules.get("regions", {})
            english_regions = st.text_area("英语地区国家", csv_text(region_rules.get("english", [])))
            traditional_regions = st.text_area("繁中地区国家", csv_text(region_rules.get("traditional", [])))
            north_america_regions = st.text_input("北美地区国家", csv_text(region_rules.get("north_america", ["美国", "加拿大"])))

        editor_name = st.text_input("修改人", value=rules.get("updated_by", ""), key="rules_updated_by")
        change_note = st.text_input("本次修改说明（必填）", key="rules_change_note")
        if st.button("保存统计规则", type="primary", key="save_report_rules"):
            updated = deepcopy(rules)
            updated.setdefault("fields", {})
            for group, frame in field_frames.items():
                updated["fields"][group] = {
                    str(row["系统标准字段"]).strip(): str(row["上传文件字段"]).strip()
                    for row in frame.to_dict("records")
                    if str(row.get("系统标准字段", "")).strip()
                    and str(row.get("上传文件字段", "")).strip()
                }
            updated["new_user"] = {"field": new_user_field.strip(), "value": new_user_value}
            updated["subscription"] = {"source_values": csv_values(subscription_values)}
            updated["channels"] = {
                "cooperation_priority": [
                    {"prefix": kol_prefix.strip(), "channel": "KOL"},
                    {"prefix": affiliate_prefix.strip(), "channel": "Affiliate"},
                ],
                "blank_cooperation_sources": {
                    **{channel: csv_values(value) for channel, value in source_inputs.items()},
                    "Auto_Renewal": csv_values(subscription_values),
                    "Blank": [""],
                },
                "default_channel": default_channel,
            }
            updated["sem_cost"] = {"google_exclude_contains": csv_values(google_excludes)}
            updated["regions"] = {
                "english": csv_values(english_regions),
                "traditional": csv_values(traditional_regions),
                "north_america": csv_values(north_america_regions),
            }
            errors = validate_rules(updated)
            if not change_note.strip():
                errors.append("请填写本次修改说明")
            if errors:
                for error in errors:
                    st.error(error)
            else:
                try:
                    save_rules(updated, st.secrets, editor_name, change_note)
                    st.success("统计规则已保存。后续生成报告将使用新规则。")
                except Exception as exc:
                    st.error(str(exc))

    with admin_settings.expander("当前统计口径说明（团队必读）", expanded=False):
        metric_doc = Path(__file__).resolve().parent / "config" / "统计口径说明.md"
        if metric_doc.exists():
            st.markdown(metric_doc.read_text(encoding="utf-8"))
        else:
            st.warning("未找到统计口径说明文件。")

with side_col:
    st.markdown('<div class="panel"><div class="panel-title">报告信息<span>自动保存</span></div>', unsafe_allow_html=True)
    start = st.date_input("开始日期", default_start, key="report_start")
    end = st.date_input("结束日期", default_end, key="report_end")
    st.caption(f"当前周期 · {start:%Y.%m.%d} – {end:%Y.%m.%d}")
    st.markdown('</div>', unsafe_allow_html=True)

    completed = 1 + uploaded_count; readiness = int(completed / 5 * 100)
    task_rows = [("报告周期",True),("注册数据",bool(reg_files)),("订单数据",bool(order_files)),("Google消耗",bool(google_files)),("Bing消耗",bool(bing_files))]
    tasks_html = "".join(f'<div class="task"><b>{"✓" if done else "○"} {name}</b><span class="{"done" if done else ""}">{"已完成" if done else "待上传"}</span></div>' for name,done in task_rows)
    st.markdown(f'<div class="panel"><div class="panel-title">报告准备度<span>{readiness}%</span></div><div class="mini-metrics"><div class="mini-metric"><strong>{readiness}%</strong><small>当前进度</small></div><div class="mini-metric"><strong>{completed} / 5</strong><small>已完成项目</small></div></div><div class="progress-track"><i style="width:{readiness}%"></i></div><div style="height:10px"></div>{tasks_html}</div>', unsafe_allow_html=True)
    st.markdown('<div class="panel"><div class="panel-title">自动化流程<span>1–3分钟</span></div><div class="timeline"><div class="event"><b>数据文件检查</b><span>上传后自动校验必要字段</span></div><div class="event"><b>字段自动识别</b><span>兼容Google与Bing中文字段</span></div><div class="event"><b>报告生成</b><span>输出全渠道 / SEM / SEM分地区 / 自然来源 / 会员类型</span></div></div></div>', unsafe_allow_html=True)
    ready = all([reg_files, order_files, google_files, bing_files]) and start <= end
    generate_clicked = st.button(
        "✦ 开始生成报告" if ready else "请先上传全部必要数据",
        type="primary", disabled=not ready, use_container_width=True,
        key="web_generate",
    )

if start > end:
    with side_col: st.error("开始日期不能晚于结束日期。")

if generate_clicked:
    try:
        with main_col:
            with st.status("正在读取和检查数据……", expanded=True) as status:
                reg = read_uploads(reg_files); st.write(f"注册数据：{len(reg):,} 行")
                orders = read_uploads(order_files); st.write(f"订单数据：{len(orders):,} 行")
                google = read_uploads(google_files); st.write(f"Google消耗：{len(google):,} 行")
                bing = read_uploads(bing_files); st.write(f"Bing消耗：{len(bing):,} 行")
                # 直接把当前目标配置交给生成引擎。这样刚在网站保存的目标无需等待
                # Streamlit 重新部署，也会立即写入本次生成的 Excel。
                reports, warnings = generate_v3_reports(
                    reg, orders, google, bing, None, start, end, input_values,
                    targets=load_targets(), rules=load_rules(),
                )
                if len(reports) != 5:
                    raise RuntimeError(f"报告生成数量异常：应生成5份，实际生成{len(reports)}份，请刷新页面后重试。")
                st.session_state["reports"] = reports; st.session_state["warnings"] = warnings
                status.update(label="五份周报已生成", state="complete", expanded=False)
    except Exception as exc:
        with main_col: st.exception(exc)

if st.session_state.get("reports"):
    reports = st.session_state["reports"]
    st.markdown(
        "<style>.st-key-save_web_manual,.inline-save-status,.fixed-footer,.page-bottom-spacer{display:none!important}</style>",
        unsafe_allow_html=True,
    )
    with main_col:
        for warning in st.session_state.get("warnings", []): st.warning(warning)
        st.markdown('<div id="report-result"></div><div class="section-label">03 报告已就绪</div><div class="section-note">可分别下载Excel，也可以一次下载完整压缩包。</div>', unsafe_allow_html=True)
        cols = st.columns(2)
        metas = ["全渠道末次归因", "SEM末次归因", "SEM分地区末次归因", "自然来源周报", "会员类型周报"]
        for index, ((filename, content), meta) in enumerate(zip(reports.items(), metas)):
            col = cols[index % 2]
            with col:
                st.markdown(f'<div class="download-card"><b>{meta}</b><br><span>Excel · {len(content)/1024:.1f} KB</span></div>', unsafe_allow_html=True)
                st.download_button("下载周报", content, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=filename)
        bundle = BytesIO()
        with ZipFile(bundle, "w", ZIP_DEFLATED) as archive:
            for filename, content in reports.items(): archive.writestr(filename, content)
        st.download_button("一键下载五份周报（ZIP）", bundle.getvalue(), f"末次归因周报_{start:%Y.%m.%d}-{end:%Y.%m.%d}.zip", "application/zip", type="primary", use_container_width=True)

st.markdown('<div class="page-bottom-spacer" aria-hidden="true"></div>', unsafe_allow_html=True)
st.markdown(
    '<div id="help"></div><div class="fixed-footer"><span>'
    '文件仅用于本次浏览器会话 · 刷新页面后需重新上传'
    '<a href="#web-targets">跳到月度目标</a></span>'
    '<b>DATA AUTOMATION · 周报工作台</b></div>',
    unsafe_allow_html=True,
)
