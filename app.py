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
# Cấu hình chung
# =========================
st.set_page_config(
    page_title="IT Helpdesk → SGDAVH",
    page_icon="🛠️",
    layout="wide",
)

APP_TITLE = "IT Helpdesk → SGDAVH"
VN_TZ = ZoneInfo("Asia/Ho_Chi_Minh")

def now_vn_rounded():
    n = datetime.now(VN_TZ)
    return n.replace(second=0, microsecond=0)

# =========================
# AUTH bằng email Streamlit (ADMIN + allowlist)
# =========================
ADMIN_EMAIL = "duydoan747@gmail.com"
ALLOWED_EMAILS = {"duydominic3@gmail.com"}

def _extract_email_from_userinfo(user_info) -> str:
    """Trả về email dạng lower-case; hỗ trợ cả dict và object."""
    if not user_info:
        return ""
    if isinstance(user_info, dict):
        return (user_info.get("email") or "").strip().lower()
    return (getattr(user_info, "email", "") or "").strip().lower()

# --- Auth: lấy email đăng nhập từ Streamlit Cloud ---
user_info = getattr(st, "user", None)
email_norm = (getattr(user_info, "email", None) or "").strip().lower()

if not email_norm:
    st.error("⚠️ Chưa nhận được email đăng nhập từ Streamlit Cloud. Hãy bật Viewer authentication trong Settings → Sharing và đăng nhập lại bằng Google.")
    st.stop()

# Quản lý quyền
ADMIN_EMAIL = "duydoan747@gmail.com"
ALLOWED_EMAILS = {"duydominic3@gmail.com"}

if email_norm == ADMIN_EMAIL:
    is_admin = True
elif email_norm in ALLOWED_EMAILS:
    is_admin = False
else:
    st.error("⛔ Bạn không có quyền truy cập app này. Liên hệ admin để được cấp quyền.")
    st.stop()


# Kiểm tra quyền
is_admin = (email_norm == ADMIN_EMAIL)
if (not is_admin) and (email_norm not in ALLOWED_EMAILS):
    st.error("⛔ Bạn không có quyền truy cập app này. Liên hệ admin để được cấp quyền.")
    st.stop()

# =========================
# Google Sheets
# =========================
SHEET_ID = st.secrets["SHEET_ID"]
GCP_SERVICE_ACCOUNT = st.secrets["gcp_service_account"]
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
    creds = Credentials.from_service_account_info(GCP_SERVICE_ACCOUNT, scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
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

    header = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=header)

    # Thêm cột thiếu
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

    for col in ["Thời gian phát sinh (UTC ISO)", "Thời gian hoàn thành (UTC ISO)", "CreatedAt (UTC ISO)"]:
        df[col] = pd.to_datetime(df[col], errors="coerce", utc=True)

    has_both = df["Thời gian phát sinh (UTC ISO)"].notna() & df["Thời gian hoàn thành (UTC ISO)"].notna()
    df.loc[has_both, "SLA_gio"] = (
        (df.loc[has_both, "Thời gian hoàn thành (UTC ISO)"] - df.loc[has_both, "Thời gian phát sinh (UTC ISO)"])
        .dt.total_seconds()
        / 3600.0
    )
    df["SLA_gio"] = pd.to_numeric(df["SLA_gio"], errors="coerce")

    df["Phát sinh (VN)"] = df["Thời gian phát sinh (UTC ISO)"].dt.tz_convert(VN_TZ)
    df["Hoàn thành (VN)"] = df["Thời gian hoàn thành (UTC ISO)"].dt.tz_convert(VN_TZ)

    df = df.sort_values(by=["Thời gian phát sinh (UTC ISO)"], ascending=False, na_position="last").reset_index(drop=True)
    return df

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
# UI nhập ticket
# =========================
st.title(APP_TITLE)
st.caption("Lưu & báo cáo ticket trực tiếp trên Google Sheets (Service Account qua Secrets)")

# Khởi tạo session state mặc định để tiện reset sau khi lưu
if "defaults" not in st.session_state:
    st.session_state.defaults = {
        "ten_cty": "",
        "shd": "",
        "nguyen_nhan": "",
        "tt_user": "",
        "tinh_trang": "Mới",
        "cach_xl": "",
        "ktv": "",
        "end_ticket": "Remote",
        "ngay_psinh": datetime.now(VN_TZ).date(),
        "gio_psinh": now_vn_rounded().time(),
        "co_tg": False,
        "ngay_done": datetime.now(VN_TZ).date(),
        "gio_done": now_vn_rounded().time(),
    }
# Copy defaults → current state (chỉ khi chưa có)
for k, v in st.session_state.defaults.items():
    st.session_state.setdefault(k, v)

with st.expander("➕ Nhập ticket mới", expanded=True):
    c1, c2 = st.columns(2)

    ten_cty = c1.text_input("Tên công ty *", value=st.session_state.ten_cty, key="ten_cty")
    ngay_psinh = c2.date_input("Ngày phát sinh *", value=st.session_state.ngay_psinh, key="ngay_psinh")
    shd = c1.text_input("SHĐ (Số HĐ/Số hồ sơ) *", value=st.session_state.shd, key="shd")
    gio_psinh = c2.time_input("Giờ phát sinh *", value=st.session_state.gio_psinh, step=60, key="gio_psinh")

    nguyen_nhan = c1.text_input("Nguyên nhân đầu vào *", value=st.session_state.nguyen_nhan, key="nguyen_nhan")
    tt_user = c2.text_input("TT User", value=st.session_state.tt_user, key="tt_user")
    cach_xl = c1.text_area("Cách xử lý * (mô tả ngắn gọn)", value=st.session_state.cach_xl, key="cach_xl")

    tinh_trang = c2.selectbox(
        "Tình trạng *",
        ["Mới", "Đang xử lý", "Hoàn thành", "Tạm dừng"],
        index=["Mới", "Đang xử lý", "Hoàn thành", "Tạm dừng"].index(st.session_state.tinh_trang),
        key="tinh_trang",
    )
    ktv = c1.text_input("KTV phụ trách", value=st.session_state.ktv, key="ktv")

    end_ticket = c2.selectbox(
        "End ticket",
        ["Remote", "Onsite", "Tạo Checklist cho chi nhánh"],
        index=["Remote", "Onsite", "Tạo Checklist cho chi nhánh"].index(st.session_state.end_ticket),
        key="end_ticket",
    )

    co_tg_hoanthanh = st.checkbox("Có thời gian hoàn thành?", value=st.session_state.co_tg, key="co_tg")
    if co_tg_hoanthanh:
        c3, c4 = st.columns(2)
        ngay_done = c3.date_input("Ngày hoàn thành", value=st.session_state.ngay_done, key="ngay_done")
        gio_done = c4.time_input("Giờ hoàn thành", value=st.session_state.gio_done, step=60, key="gio_done")
        tg_done_utc = local_to_utc_iso(ngay_done, gio_done)
    else:
        tg_done_utc = ""

    if st.button("Lưu vào Google Sheet", type="primary"):
        required = [ten_cty, shd, nguyen_nhan, cach_xl, tinh_trang]
        if any(not (x or "").strip() for x in required):
            st.error("⚠️ Vui lòng điền đầy đủ các trường bắt buộc (*)")
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
                    tt_user or "",
                    tinh_trang,
                    cach_xl,
                    end_ticket,
                    tg_ps_utc,
                    tg_done_utc,
                    ktv or "",
                    created_utc,
                    sla_gio,
                ]
                append_ticket(row)

                # ✅ Reset toàn bộ input về mặc định sau khi lưu
                for k, v in st.session_state.defaults.items():
                    st.session_state[k] = v

                st.success("✅ Đã lưu ticket vào Google Sheet! (form đã được reset)")
            except Exception as e:
                st.error(f"❌ Lỗi khi ghi Google Sheet: {e}")

st.divider()

# =========================
# Báo cáo & lọc dữ liệu
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
        m_start = datetime(from_day.year, from_day.month, from_day.day, 0, 0, tzinfo=VN_TZ)
        m_end = datetime(to_day.year, to_day.month, to_day.day, 23, 59, 59, tzinfo=VN_TZ)

        df = df_raw.copy()
        df = df[df["Phát sinh (VN)"].between(m_start, m_end, inclusive="both")]

        if flt_cty.strip():
            df = df[df["Tên công ty"].str.contains(flt_cty.strip(), case=False, na=False)]
        if flt_ktv.strip():
            df = df[df["KTV"].str.contains(flt_ktv.strip(), case=False, na=False)]

        show_cols = [
            "Tên công ty", "SHĐ", "Nguyên nhân đầu vào", "TT User", "Tình trạng",
            "Cách xử lý", "End ticket", "Phát sinh (VN)", "Hoàn thành (VN)", "KTV", "SLA_gio"
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

        # Chỉ admin mới tải CSV
        if is_admin:
            st.download_button(
                "⬇️ Tải CSV đã lọc",
                data=to_csv_bytes(df[show_cols]),
                file_name=f"helpdesk_{from_day}_{to_day}.csv",
                mime="text/csv",
            )
        else:
            st.info("Chỉ admin mới có quyền tải báo cáo CSV.")
except Exception as e:
    st.error(f"❌ Đã gặp lỗi khi tải dữ liệu: {e}")
