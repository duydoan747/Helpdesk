# app.py
from __future__ import annotations
import io
from datetime import datetime, date, time, timezone
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound

# =========================
# Config cơ bản
# =========================
st.set_page_config(
    page_title="IT Helpdesk → SGDAVH",
    page_icon="🛠️",
    layout="wide",
)

APP_TITLE = "IT Helpdesk → SGDAVH"
VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")

# =========================
# PHÂN QUYỀN (ADMIN + VIEWER)
# =========================
ADMIN_EMAILS = {
    "duydoan747@gmail.com",   # admin
}
VIEWER_EMAILS = {
    "duydominic33@gmail.com",
}
ADMIN_EMAILS  = {e.strip().lower() for e in ADMIN_EMAILS}
VIEWER_EMAILS = {e.strip().lower() for e in VIEWER_EMAILS}
ALL_ALLOWED   = ADMIN_EMAILS | VIEWER_EMAILS

user_info = getattr(st, "experimental_user", None)
email_norm = (getattr(user_info, "email", None) or "").strip().lower()

st.sidebar.info(f"👤 Email đăng nhập hiện tại: {email_norm or 'N/A'}")

def is_admin(e: str) -> bool:
    return e in ADMIN_EMAILS

def is_allowed(e: str) -> bool:
    return e in ALL_ALLOWED

if not email_norm:
    st.warning("App đang ở chế độ public hoặc chưa nhận được email đăng nhập. "
               "Bạn chỉ có thể xem báo cáo. Hãy bật Viewer authentication để hạn chế quyền.")
    CAN_CREATE = False
else:
    CAN_CREATE = is_allowed(email_norm)

if email_norm and not CAN_CREATE:
    st.error("⛔ Bạn không có quyền tạo ticket. Bạn chỉ có thể xem báo cáo.")
    # Chỉ xem báo cáo, không được tạo ticket

# =========================
# Google Sheets
# =========================
SHEET_ID: str = st.secrets["SHEET_ID"]
SHEET_NAME = "Data"

COLUMNS = [
    "Tên công ty",
    "SHĐ",
    "Nguyên nhân đầu vào",
    "TT User",
    "Tình trạng",
    "Cách xử lý",
    "End ticket",
    "Thời gian phát sinh (UTC ISO)",
    "Thời gian hoàn thành (UTC ISO)",
    "KTV",
    "CreatedAt (UTC ISO)",
    "SLA_gio",
]

def get_gspread_client_service():
    sa_info = st.secrets["gcp_service_account"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def open_worksheet():
    gc = get_gspread_client_service()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=1000, cols=len(COLUMNS))
        ws.append_row(COLUMNS, value_input_option="RAW")
    return ws

@st.cache_data(show_spinner=False, ttl=60)
def read_all_as_dataframe() -> pd.DataFrame:
    ws = open_worksheet()
    values = ws.get_all_values()
    if not values or len(values) == 1:
        return pd.DataFrame(columns=COLUMNS)

    df = pd.DataFrame(values[1:], columns=values[0])
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    for col in ["Thời gian phát sinh (UTC ISO)", "Thời gian hoàn thành (UTC ISO)", "CreatedAt (UTC ISO)"]:
        df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)

    has_both = df["Thời gian phát sinh (UTC ISO)"].notna() & df["Thời gian hoàn thành (UTC ISO)"].notna()
    df.loc[has_both, "SLA_gio"] = (
        (df.loc[has_both, "Thời gian hoàn thành (UTC ISO)"] - df.loc[has_both, "Thời gian phát sinh (UTC ISO)"])
        .dt.total_seconds() / 3600.0
    )
    df["SLA_gio"] = pd.to_numeric(df["SLA_gio"], errors="coerce")

    df["Phát sinh (VN)"] = df["Thời gian phát sinh (UTC ISO)"].dt.tz_convert(VN_TZ)
    df["Hoàn thành (VN)"] = df["Thời gian hoàn thành (UTC ISO)"].dt.tz_convert(VN_TZ)

    return df.sort_values(by="Thời gian phát sinh (UTC ISO)", ascending=False, na_position="last").reset_index(drop=True)

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    out = io.StringIO()
    df.to_csv(out, index=False, encoding="utf-8")
    return out.getvalue().encode("utf-8")

def local_to_utc_iso(d: date, t: time) -> str:
    dt_local = datetime(d.year, d.month, d.day, t.hour, t.minute, t.second, tzinfo=VN_TZ)
    return dt_local.astimezone(timezone.utc).isoformat()

def append_ticket(row: list[str]) -> None:
    ws = open_worksheet()
    ws.append_row(row, value_input_option="RAW")

# =========================
# UI
# =========================
st.title(APP_TITLE)
st.caption("Lưu & báo cáo ticket trực tiếp trên Google Sheets (Service Account qua Secrets)")

# Form tạo ticket (chỉ người có quyền mới thấy)
if CAN_CREATE:
    with st.expander("➕ Nhập ticket mới", expanded=True):
        with st.form("ticket_form", clear_on_submit=True):
            c1, c2 = st.columns(2)

            ten_cty = c1.text_input("Tên công ty *")
            ngay_psinh = c2.date_input("Ngày phát sinh *", value=datetime.now(VN_TZ).date(), format="YYYY/MM/DD")
            shd = c1.text_input("SHĐ (Số HĐ/Số hồ sơ) *")
            gio_psinh = c2.time_input("Giờ phát sinh *", value=datetime.now(VN_TZ).time().replace(second=0), step=60)

            nguyen_nhan = c1.text_input("Nguyên nhân đầu vào *")
            tt_user = c2.text_input("TT User")
            cach_xl = c1.text_area("Cách xử lý * (mô tả ngắn gọn)")
            end_ticket = c2.selectbox("End ticket", ["Remote", "Onsite", "Tạo Checklist cho chi nhánh"])

            tinh_trang = c1.selectbox("Tình trạng *", ["Mới", "Đang xử lý", "Hoàn thành", "Tạm dừng"])
            ktv = c2.text_input("KTV phụ trách")

            co_tg_hoanthanh = st.checkbox("Có thời gian hoàn thành?")
            if co_tg_hoanthanh:
                c3, c4 = st.columns(2)
                ngay_done = c3.date_input("Ngày hoàn thành", value=datetime.now(VN_TZ).date(), format="YYYY/MM/DD")
                gio_done = c4.time_input("Giờ hoàn thành", value=datetime.now(VN_TZ).time().replace(second=0), step=60)
                tg_done_utc = local_to_utc_iso(ngay_done, gio_done)
            else:
                tg_done_utc = ""

            submitted = st.form_submit_button("Lưu vào Google Sheet", type="primary")
            if submitted:
                if not (ten_cty and shd and nguyen_nhan and cach_xl and tinh_trang):
                    st.error("Vui lòng điền đầy đủ các trường bắt buộc (*)")
                else:
                    try:
                        tg_ps_utc = local_to_utc_iso(ngay_psinh, gio_psinh)
                        created_utc = datetime.now(timezone.utc).isoformat()

                        if tg_done_utc:
                            start = datetime.fromisoformat(tg_ps_utc.replace("Z", "+00:00"))
                            end = datetime.fromisoformat(tg_done_utc.replace("Z", "+00:00"))
                            sla_gio = round((end - start).total_seconds() / 3600.0, 2)
                        else:
                            sla_gio = ""

                        row = [
                            ten_cty,
                            shd,
                            nguyen_nhan,
                            tt_user,
                            tinh_trang,
                            cach_xl,
                            end_ticket,
                            tg_ps_utc,
                            tg_done_utc,
                            ktv,
                            created_utc,
                            sla_gio,
                        ]
                        append_ticket(row)
                        st.success("✅ Đã lưu ticket vào Google Sheet!")
                        st.balloons()
                    except Exception as e:
                        st.error(f"❌ Lỗi khi ghi Google Sheet: {e}")

st.divider()

# =========================
# Báo cáo
# =========================
st.header("📊 Báo cáo & Lọc dữ liệu")

c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
today_vn = datetime.now(VN_TZ).date()
from_day = c1.date_input("Từ ngày", value=today_vn.replace(day=max(1, today_vn.day - 7)), format="YYYY/MM/DD")
to_day = c2.date_input("Đến ngày", value=today_vn, format="YYYY/MM/DD")
flt_cty = c3.text_input("Lọc theo tên Cty")
flt_ktv = c4.text_input("Lọc theo KTV")

try:
    df_raw = read_all_as_dataframe()
    if df_raw.empty:
        st.info("Chưa có dữ liệu.")
    else:
        m_start = datetime(from_day.year, from_day.month, from_day.day, 0, 0, tzinfo=VN_TZ)
        m_end = datetime(to_day.year, to_day.month, to_day.day, 23, 59, 59, tzinfo=VN_TZ)

        df = df_raw.copy()
        in_range = df["Phát sinh (VN)"].between(m_start, m_end, inclusive="both")
        df = df[in_range]

        if flt_cty.strip():
            df = df[df["Tên công ty"].str.contains(flt_cty.strip(), case=False, na=False)]
        if flt_ktv.strip():
            df = df[df["KTV"].str.contains(flt_ktv.strip(), case=False, na=False)]

        show_cols = [
            "Tên công ty",
            "SHĐ",
            "Nguyên nhân đầu vào",
            "TT User",
            "Tình trạng",
            "Cách xử lý",
            "End ticket",
            "Phát sinh (VN)",
            "Hoàn thành (VN)",
            "KTV",
            "SLA_gio",
        ]
        st.dataframe(
            df[show_cols].assign(
                **{
                    "Phát sinh (VN)": df["Phát sinh (VN)"].dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "Hoàn thành (VN)": df["Hoàn thành (VN)"].dt.strftime("%Y-%m-%d %H:%M:%S"),
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "⬇️ Tải CSV đã lọc",
            data=to_csv_bytes(df[show_cols]),
            file_name=f"helpdesk_{from_day}_{to_day}.csv",
            mime="text/csv",
        )
except Exception as e:
    st.error(f"Đã gặp lỗi khi tải dữ liệu: {e}")
