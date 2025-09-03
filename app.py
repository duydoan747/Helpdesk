# app.py
from __future__ import annotations

import io
import time, random
from functools import wraps
from datetime import datetime, date, time as dtime, timezone
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound
# Debug user info từ Streamlit Cloud
try:
    user_info = st.experimental_user  # chỉ hoạt động khi Viewer authentication bật
    st.sidebar.write("🔍 Debug user_info:", user_info)
except Exception as e:
    st.sidebar.error(f"Lỗi khi lấy user_info: {e}")

# =========================
# Cấu hình chung
# =========================
st.set_page_config(
    page_title="IT Helpdesk → SGDAVH",
    page_icon="🛠️",
    layout="wide",
)

APP_TITLE = "IT Helpdesk → SGDAVH"
VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")

# =========================
# PHÂN QUYỀN THEO EMAIL
# =========================
RAW_ALLOWED = {
    "duydv3@fpt.com",
    "duydoan747@gmail.com"
    "congnv17@fpt.com",
    "vuln3@fpt.com",
    "vinhpt14@fpt.com",
    "phubq2@fpt.com",
    "phuongnam.kietnp@fpt.net",
}
ALLOWED_EMAILS = {e.strip().lower() for e in RAW_ALLOWED}

# lấy email user từ Streamlit Cloud (cần bật Viewer authentication)
user = getattr(st, "experimental_user", None)
email = getattr(user, "email", None)
email_norm = (email or "").strip().lower()

# luôn hiển thị email hiện tại để dễ debug
st.sidebar.info(f"👤 Email đăng nhập hiện tại: {email_norm or 'N/A'}")

if not email_norm:
    st.error("⛔ Chưa nhận được email đăng nhập từ Streamlit Cloud.")
    st.caption("👉 Hãy bật Viewer authentication trong Settings → Sharing của app, sau đó đăng nhập lại bằng Google.")
    st.stop()

if email_norm not in ALLOWED_EMAILS:
    st.error("⛔ Bạn không có quyền truy cập ứng dụng này.")
    st.caption(f"Email hiện tại: {email_norm}")
    st.info("Nếu đúng email công ty mà vẫn bị chặn, hãy kiểm tra allowlist trong code hoặc Settings → Sharing.")
    st.stop()

# =========================
# Kết nối Google Sheets
# =========================
SHEET_ID: str = st.secrets["SHEET_ID"]
SHEET_NAME = "Data"

COLUMNS = [
    "Tên công ty",
    "SHD",
    "Nguyên nhân đầu vào",
    "TT User",
    "Tình trạng",
    "End ticket",
    "Cách xử lý",
    "Thời gian phát sinh (UTC ISO)",
    "Thời gian hoàn thành (UTC ISO)",
    "KTV",
    "CreatedAt (UTC ISO)",
    "CreatedBy",
    "SLA_gio",
]

def get_gspread_client_service():
    sa_info = st.secrets["gcp_service_account"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

def _ensure_header(ws) -> None:
    header = ws.row_values(1)
    if header != COLUMNS:
        ws.update("A1", [COLUMNS])

@st.cache_resource(show_spinner=False)
def open_worksheet():
    gc = get_gspread_client_service()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(COLUMNS))
        ws.update("A1", [COLUMNS])
        return ws
    _ensure_header(ws)
    return ws

# =========================
# Helpers
# =========================
def now_vn_rounded():
    n = datetime.now(VN_TZ)
    return n.replace(second=0, microsecond=0)

def to_utc_iso(dt_local: datetime) -> str:
    return dt_local.astimezone(timezone.utc).isoformat()

def append_ticket(row: list[str]) -> None:
    ws = open_worksheet()
    ws.append_row(row, value_input_option="RAW")

# =========================
# UI
# =========================
st.title(APP_TITLE)
st.caption("Lưu & báo cáo ticket trực tiếp trên Google Sheets (Service Account qua Secrets)")

# (phần form nhập ticket và báo cáo bạn giữ nguyên như bản trước — không thay đổi)
