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

# Lấy từ Secrets
SHEET_ID: str = st.secrets["SHEET_ID"]
SHEET_NAME = "Data"

# Header cố định trên Sheet (thứ tự cột)
COLUMNS = [
    "Tên công ty",
    "SHD",
    "Nguyên nhân đầu vào",
    "TT User",
    "Tình trạng",
    "End ticket",  # NEW
    "Cách xử lý",
    "Thời gian phát sinh (UTC ISO)",
    "Thời gian hoàn thành (UTC ISO)",
    "KTV",
    "CreatedAt (UTC ISO)",
    "SLA_gio",
]

# =========================
# Retry/Backoff utilities
# =========================
def retry(max_attempts=5, base=0.5, cap=8.0):
    """Exponential backoff cho các call tới Google API."""
    def deco(fn):
        @wraps(fn)
        def inner(*args, **kwargs):
            delay = base
            for attempt in range(1, max_attempts + 1):
                try:
                    return fn(*args, **kwargs)
                except Exception as e:
                    if attempt == max_attempts:
                        raise
                    time.sleep(delay + random.random() * 0.2)
                    delay = min(delay * 2, cap)
        return inner
    return deco

# =========================
# Kết nối Google Sheets
# =========================
def get_gspread_client_service():
    """Authorize gspread dùng dict trong secrets['gcp_service_account'] với scope tối thiểu."""
    sa_info = st.secrets["gcp_service_account"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",   # ghi/đọc sheet
        "https://www.googleapis.com/auth/drive.file",     # chỉ các file được share/được app tạo
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

def _ensure_header(ws) -> None:
    """Đảm bảo hàng 1 là header đúng như COLUMNS."""
    header = ws.row_values(1)
    if header != COLUMNS:
        ws.update("A1", [COLUMNS])

@st.cache_resource(show_spinner=False)
def open_worksheet():
    """Mở worksheet; nếu chưa có thì tạo và ghi header."""
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

def _open_or_create_logs(sh):
    try:
        return sh.worksheet("Logs")
    except WorksheetNotFound:
        wslog = sh.add_worksheet("Logs", rows=1000, cols=3)
        wslog.append_row(["UTC", "Action", "Message"])
        return wslog

def log_error(action: str, msg: str):
    """Ghi lỗi nhẹ nhàng vào sheet Logs (im lặng khi thất bại để tránh vòng lặp)."""
    try:
        gc = get_gspread_client_service()
        sh = gc.open_by_key(SHEET_ID)
        wslog = _open_or_create_logs(sh)
        wslog.append_row([datetime.utcnow().isoformat(), action, str(msg)])
    except Exception:
        pass

@retry()
def _safe_append_row(ws, row):
    ws.append_row(row, value_input_option="RAW")

@retry()
def _safe_get_all_values(ws):
    return ws.get_all_values()

@st.cache_data(show_spinner=False, ttl=60)
def read_all_as_dataframe() -> pd.DataFrame:
    """Đọc toàn bộ dữ liệu thành DataFrame; parse thời gian & tính SLA."""
    ws = open_worksheet()
    values = _safe_get_all_values(ws)

    if not values or len(values) == 1:  # chỉ có header hoặc rỗng
        return pd.DataFrame(columns=COLUMNS)

    header = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)

    # Bổ sung cột thiếu (nếu sheet cũ chưa đủ header)
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # Parse thời gian UTC
    for col in ["Thời gian phát sinh (UTC ISO)", "Thời gian hoàn thành (UTC ISO)", "CreatedAt (UTC ISO)"]:
        df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)

    # Tính SLA_gio nếu có đủ 2 mốc
    has_both = df["Thời gian phát sinh (UTC ISO)"].notna() & df["Thời gian hoàn thành (UTC ISO)"].notna()
    df.loc[has_both, "SLA_gio"] = (
        (df.loc[has_both, "Thời gian hoàn thành (UTC ISO)"] - df.loc[has_both, "Thời gian phát sinh (UTC ISO)"])
        .dt.total_seconds() / 3600.0
    )
    df["SLA_gio"] = pd.to_numeric(df["SLA_gio"], errors="coerce")

    # Thêm cột hiển thị theo giờ VN
    df["Phát sinh (VN)"] = df["Thời gian phát sinh (UTC ISO)"].dt.tz_convert(VN_TZ)
    df["Hoàn thành (VN)"] = df["Thời gian hoàn thành (UTC ISO)"].dt.tz_convert(VN_TZ)

    # Sắp xếp mới nhất trước
    df = df.sort_values(by=["Thời gian phát sinh (UTC ISO)"], ascending=False, na_position="last").reset_index(drop=True)
    return df

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    out = io.StringIO()
    df.to_csv(out, index=False, encoding="utf-8")
    return out.getvalue().encode("utf-8")

def local_to_utc_iso(d: date, t: dtime) -> str:
    """Ghép ngày+giờ VN → UTC ISO string."""
    dt_local = datetime(d.year, d.month, d.day, t.hour, t.minute, t.second, tzinfo=VN_TZ)
    return dt_local.astimezone(timezone.utc).isoformat()

def append_ticket(row: list[str]) -> None:
    ws = open_worksheet()
    _safe_append_row(ws, row)

# =========================
# UI
# =========================
st.title(APP_TITLE)
st.caption("Lưu & báo cáo ticket trực tiếp trên Google Sheets (Service Account qua Secrets)")

with st.expander("➕ Nhập ticket mới", expanded=True):
    # ===== Helpers =====
    def _now_vn_rounded():
        n = datetime.now(VN_TZ)
        return n.replace(second=0, microsecond=0)

    # ===== Defaults in session_state (giữ nguyên sau khi rerun) =====
    if "ngay_psinh" not in st.session_state:
        st.session_state["ngay_psinh"] = datetime.now(VN_TZ).date()
    if "gio_psinh" not in st.session_state:
        st.session_state["gio_psinh"] = _now_vn_rounded().time()

    if "co_tg_hoanthanh" not in st.session_state:
        st.session_state["co_tg_hoanthanh"] = False
    if "ngay_done" not in st.session_state:
        st.session_state["ngay_done"] = datetime.now(VN_TZ).date()
    if "gio_done" not in st.session_state:
        st.session_state["gio_done"] = _now_vn_rounded().time()

    # ===== Form inputs =====
    c1, c2 = st.columns(2)

    ten_cty = c1.text_input("Tên công ty *").strip()
    # Ngày/Giờ phát sinh: dùng key để giữ giá trị đã chọn
    ngay_psinh = c2.date_input("Ngày phát sinh *",
                               key="ngay_psinh",
                               format="YYYY/MM/DD")
    shd = c1.text_input("SHD (Số HĐ/Số hồ sơ) *").strip()
    gio_psinh = c2.time_input("Giờ phát sinh *",
                              key="gio_psinh",
                              step=60)

    nguyen_nhan = c1.text_input("Nguyên nhân đầu vào *").strip()
    tt_user = c2.text_input("TT User").strip()

    cach_xl = c1.text_area("Cách xử lý * (mô tả ngắn gọn)").strip()

    tinh_trang = c2.selectbox("Tình trạng *", ["Mới", "Đang xử lý", "Hoàn thành", "Tạm dừng"])

    # End ticket
    end_ticket = c1.selectbox(
        "End ticket",
        ["Remote", "Onsite", "Tạo Checklist cho chi nhánh"],
        index=0,
    )

    ktv = c2.text_input("KTV phụ trách").strip()

    # Có thời gian hoàn thành?
    co_tg_hoanthanh = st.checkbox("Có thời gian hoàn thành?",
                                  key="co_tg_hoanthanh")
    if co_tg_hoanthanh:
        c3, c4 = st.columns(2)
        ngay_done = c3.date_input("Ngày hoàn thành",
                                  key="ngay_done",
                                  format="YYYY/MM/DD")
        gio_done = c4.time_input("Giờ hoàn thành",
                                 key="gio_done",
                                 step=60)
        # Kết quả UTC ISO
        tg_done_utc = local_to_utc_iso(st.session_state["ngay_done"],
                                       st.session_state["gio_done"])
    else:
        # Nếu bỏ chọn, xóa giá trị ghi xuống để không tính SLA
        tg_done_utc = ""

    # Nút reset thời gian về "bây giờ" (tuỳ chọn)
    if st.button("⟲ Đặt lại giờ về hiện tại (VN)"):
        st.session_state["gio_psinh"] = _now_vn_rounded().time()
        if st.session_state.get("co_tg_hoanthanh"):
            st.session_state["gio_done"] = _now_vn_rounded().time()
        st.success("Đã đặt lại giờ về thời điểm hiện tại.")

    # ===== Lưu =====
    if st.button("Lưu vào Google Sheet", type="primary"):
        required = [ten_cty, shd, nguyen_nhan, cach_xl, tinh_trang]
        if any(not x for x in required):
            st.error("Vui lòng điền đầy đủ các trường bắt buộc (*)")
        else:
            try:
                # Lấy từ session_state để đảm bảo đúng giờ người dùng đã chọn
                tg_ps_utc = local_to_utc_iso(st.session_state["ngay_psinh"],
                                             st.session_state["gio_psinh"])
                created_utc = datetime.now(timezone.utc).isoformat()

                # SLA (giờ)
                if tg_done_utc:
                    start = datetime.fromisoformat(tg_ps_utc.replace("Z", "+00:00"))
                    end = datetime.fromisoformat(tg_done_utc.replace("Z", "+00:00"))
                    sla_gio = round((end - start).total_seconds() / 3600.0, 2)
                else:
                    sla_gio = ""

                # Thứ tự row KHỚP 100% COLUMNS (có End ticket)
                row = [
                    ten_cty,                   # "Tên công ty"
                    shd,                       # "SHD"
                    nguyen_nhan,               # "Nguyên nhân đầu vào"
                    tt_user,                   # "TT User"
                    tinh_trang,                # "Tình trạng"
                    end_ticket,                # "End ticket"
                    cach_xl,                   # "Cách xử lý"
                    tg_ps_utc,                 # "Thời gian phát sinh (UTC ISO)"
                    tg_done_utc,               # "Thời gian hoàn thành (UTC ISO)"
                    ktv,                       # "KTV"
                    created_utc,               # "CreatedAt (UTC ISO)"
                    sla_gio,                   # "SLA_gio"
                ]
                append_ticket(row)
                st.cache_data.clear()
                st.success("✅ Đã lưu ticket vào Google Sheet!")
                st.balloons()
            except Exception as e:
                log_error("APPEND", e)
                st.error(f"❌ Lỗi khi ghi Google Sheet: {e}")


st.divider()

# =========================
# Báo cáo & Lọc dữ liệu
# =========================
st.header("📊 Báo cáo & Lọc dữ liệu")

c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
today_vn = datetime.now(VN_TZ).date()
from_day = c1.date_input("Từ ngày", value=today_vn, format="YYYY/MM/DD")
to_day = c2.date_input("Đến ngày", value=today_vn, format="YYYY/MM/DD")
flt_cty = c3.text_input("Lọc theo tên Cty")
flt_ktv = c4.text_input("Lọc theo KTV")

try:
    df_raw = read_all_as_dataframe()
    if df_raw.empty:
        st.info("Chưa có dữ liệu.")
    else:
        # Lọc theo ngày (dựa trên thời gian *VN*)
        m_start = datetime(from_day.year, from_day.month, from_day.day, 0, 0, tzinfo=VN_TZ)
        m_end = datetime(to_day.year, to_day.month, to_day.day, 23, 59, 59, tzinfo=VN_TZ)

        df = df_raw.copy()
        in_range = df["Phát sinh (VN)"].between(m_start, m_end, inclusive="both")
        df = df[in_range]

        if flt_cty.strip():
            df = df[df["Tên công ty"].astype(str).str.contains(flt_cty.strip(), case=False, na=False)]
        if flt_ktv.strip():
            df = df[df["KTV"].astype(str).str.contains(flt_ktv.strip(), case=False, na=False)]

        # Hiển thị gọn gàng
        show_cols = [
            "Tên công ty",
            "SHD",
            "Nguyên nhân đầu vào",
            "TT User",
            "Tình trạng",
            "End ticket",      # NEW
            "Cách xử lý",
            "Phát sinh (VN)",
            "Hoàn thành (VN)",
            "KTV",
            "SLA_gio",
        ]
        cols_view = [c for c in show_cols if c in df.columns]

        if "Phát sinh (VN)" in df.columns:
            df["Phát sinh (VN)"] = df["Phát sinh (VN)"].dt.strftime("%Y-%m-%d %H:%M:%S")
        if "Hoàn thành (VN)" in df.columns:
            df["Hoàn thành (VN)"] = df["Hoàn thành (VN)"].dt.strftime("%Y-%m-%d %H:%M:%S")

        st.dataframe(df[cols_view] if cols_view else df, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Tải CSV đã lọc",
            data=to_csv_bytes(df[cols_view] if cols_view else df),
            file_name=f"helpdesk_{from_day}_{to_day}.csv",
            mime="text/csv",
        )
except Exception as e:
    log_error("REPORT_LOAD", e)
    st.error(f"Đã gặp lỗi khi tải dữ liệu: {e}")
