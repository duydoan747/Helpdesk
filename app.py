# -*- coding: utf-8 -*-
"""
IT Helpdesk → Google Sheets (Service Account via Streamlit Secrets)
- Đọc SHEET_ID và khóa dịch vụ từ st.secrets (không cần file JSON cạnh app)
- Có trường 'Nguyên nhân đầu vào' (trước 'Tình trạng')
- Form nhập dùng date_input + time_input
"""

import json
from google.oauth2.service_account import Credentials
import gspread
import streamlit as st

def get_gspread_client_service():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    sa_info = dict(st.secrets["gcp_service_account"])
creds = Credentials.from_service_account_info(sa_info, scopes=scopes)     #
    return gspread.authorize(creds)

# =========================
# Cấu hình chung
# =========================
st.set_page_config(page_title="IT Helpdesk → Google Sheets", page_icon="🧰", layout="wide")
APP_TITLE = "IT Helpdesk → Google Sheets"
SHEET_ID = st.secrets["SHEET_ID"]                     # <-- lấy từ secrets
TZ = tz.gettz("Asia/Ho_Chi_Minh")
SHEET_NAME = "Data"

# Các cột cố định trên Sheet
COLUMNS = [
    "Tên công ty", "ShĐ", "Nguyên nhân đầu vào", "Tình trạng", "Cách xử lý",
    "Thời gian phát sinh (UTC ISO)", "Thời gian hoàn thành (UTC ISO)",
    "KTV", "CreatedAt (UTC ISO)"
]

# =========================
# Kết nối Google Sheets (Service Account từ secrets)
# =========================
from google.oauth2.service_account import Credentials
import gspread

def get_gspread_client_service():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    sa_info = dict(st.secrets["gcp_service_account"])     # đọc trực tiếp từ secrets
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_resource
def open_worksheet():
    gc = get_gspread_client_service()
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=2000, cols=len(COLUMNS) + 5)
        ws.append_row(COLUMNS)  # Header
    # Đảm bảo header đúng
    header = ws.row_values(1)
    if header != COLUMNS:
        ws.update(f"A1:{chr(64+len(COLUMNS))}1", [COLUMNS])
    return ws

def append_ticket_row(row_values):
    ws = open_worksheet()
    ws.append_row(row_values, value_input_option="USER_ENTERED")

def read_all_as_dataframe():
    ws = open_worksheet()
    values = ws.get_all_values()
    if not values or len(values) < 2:
        return pd.DataFrame(columns=COLUMNS)
    df = pd.DataFrame(values[1:], columns=values[0])  # bỏ header dòng 1
    return df

# =========================
# Tiện ích thời gian & xuất file
# =========================
def local_dt_to_utc_iso(dt_local):
    """Nhận datetime (TZ VN) → trả ISO UTC."""
    if dt_local is None:
        return ""
    if dt_local.tzinfo is None:
        dt_local = dt_local.replace(tzinfo=TZ)
    return dt_local.astimezone(tz.UTC).replace(second=0, microsecond=0).isoformat()

def utc_iso_to_vn_str(iso_str):
    try:
        if not iso_str:
            return ""
        dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00")).astimezone(TZ)
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return iso_str

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="BaoCao")
    return output.getvalue()

# =========================
# UI
# =========================
st.title(APP_TITLE)
st.caption("Lưu & báo cáo ticket trực tiếp trên Google Sheets (Service Account qua Secrets)")

# ====== FORM NHẬP
with st.expander("➕ Nhập ticket mới", expanded=True):
    col1, col2 = st.columns(2)

    # Cột trái: thông tin chung
    with col1:
        ten_cty = st.text_input("Tên công ty *", "")
        so_hd = st.text_input("ShĐ (Số HĐ/Số hồ sơ) *", "")
        nguyen_nhan = st.text_input("Nguyên nhân đầu vào *", "")
        tinh_trang = st.selectbox("Tình trạng *", ["Mới", "Đang xử lý", "Chờ KH phản hồi", "Hoàn thành"])
        ktv = st.text_input("KTV phụ trách", "")

    # Cột phải: thời gian & mô tả
    with col2:
        now_vn = datetime.now(TZ).replace(second=0, microsecond=0)

        # Phát sinh: tách Ngày + Giờ
        ps_date = st.date_input("Ngày phát sinh *", value=now_vn.date(), key="ps_date")
        ps_time = st.time_input("Giờ phát sinh *", value=now_vn.time(), step=900, key="ps_time")  # 15'
        cach_xu_ly = st.text_area("Cách xử lý * (mô tả ngắn gọn)", height=120)

        # Hoàn thành (tuỳ chọn)
        has_done = st.checkbox("Có thời gian hoàn thành?", value=False)
        done_date = None
        done_time = None
        if has_done:
            done_date = st.date_input("Ngày hoàn thành", value=now_vn.date(), key="done_date")
            done_time = st.time_input("Giờ hoàn thành", value=now_vn.time(), step=900, key="done_time")

    # Nút lưu
    b1, _ = st.columns([1,4])
    if b1.button("Lưu vào Google Sheet", type="primary", use_container_width=True):
        # Validate
        missing = []
        if not ten_cty.strip(): missing.append("Tên công ty")
        if not so_hd.strip(): missing.append("ShĐ")
        if not nguyen_nhan.strip(): missing.append("Nguyên nhân đầu vào")
        if not tinh_trang.strip(): missing.append("Tình trạng")
        if not cach_xu_ly.strip(): missing.append("Cách xử lý")
        if not (ps_date and ps_time): missing.append("Thời gian phát sinh")
        if tinh_trang == "Hoàn thành" and not (has_done and done_date and done_time):
            missing.append("Thời gian hoàn thành (khi trạng thái = Hoàn thành)")

        if missing:
            st.error("Thiếu thông tin: " + ", ".join(missing))
        else:
            # Gộp ngày+giờ thành datetime (múi giờ VN) rồi convert UTC ISO
            def combine_to_iso(d, t):
                if not d or not t:
                    return ""
                dt_local = datetime.combine(d, t).replace(tzinfo=TZ, second=0, microsecond=0)
                return local_dt_to_utc_iso(dt_local)

            tg_phat_sinh_iso = combine_to_iso(ps_date, ps_time)
            tg_hoan_thanh_iso = combine_to_iso(done_date, done_time) if has_done else ""

            try:
                row = [
                    ten_cty.strip(),
                    so_hd.strip(),
                    nguyen_nhan.strip(),
                    tinh_trang.strip(),
                    cach_xu_ly.strip(),
                    tg_phat_sinh_iso,
                    tg_hoan_thanh_iso,
                    ktv.strip(),
                    datetime.utcnow().replace(second=0, microsecond=0).isoformat()
                ]
                append_ticket_row(row)
                st.success("Đã lưu lên Google Sheet ✅")
            except Exception as e:
                st.error(f"Lỗi khi ghi Google Sheet: {e}")

st.divider()
st.subheader("📊 Báo cáo & Lọc dữ liệu")

with st.container():
    c1, c2, c3, c4 = st.columns([1.2,1.2,1,1])
    today = datetime.now(TZ).date()
    from_dt = c1.date_input("Từ ngày", value=today - timedelta(days=7))
    to_dt = c2.date_input("Đến ngày", value=today + timedelta(days=1))
    ten_cty_filter = c3.text_input("Lọc theo tên Cty", "")
    ktv_filter = c4.text_input("Lọc theo KTV", "")

    df_raw = read_all_as_dataframe()

    if df_raw.empty:
        st.info("Chưa có dữ liệu trên Google Sheet.")
    else:
        # Lọc theo thời gian phát sinh (UTC ISO)
        def in_range(iso_str):
            try:
                if not iso_str:
                    return False
                dt_utc = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))
                start = datetime(from_dt.year, from_dt.month, from_dt.day, 0, 0, tzinfo=TZ).astimezone(tz.UTC)
                end = datetime(to_dt.year, to_dt.month, to_dt.day, 0, 0, tzinfo=TZ).astimezone(tz.UTC)
                return start <= dt_utc < end
            except Exception:
                return False

        mask = df_raw["Thời gian phát sinh (UTC ISO)"].apply(in_range)
        if ten_cty_filter.strip():
            mask &= df_raw["Tên công ty"].str.contains(ten_cty_filter.strip(), case=False, na=False)
        if ktv_filter.strip():
            mask &= df_raw["KTV"].str.contains(ktv_filter.strip(), case=False, na=False)

        df = df_raw[mask].copy()

        # Tính SLA (giờ)
        def calc_sla(row):
            try:
                s = row["Thời gian phát sinh (UTC ISO)"]
                e = row["Thời gian hoàn thành (UTC ISO)"]
                if s and e:
                    start = datetime.fromisoformat(s.replace("Z", "+00:00"))
                    end = datetime.fromisoformat(e.replace("Z", "+00:00"))
                    return round((end - start).total_seconds()/3600.0, 2)
            except:
                return None
            return None

        df["SLA_gio"] = df.apply(calc_sla, axis=1)

        # Hiển thị cột thời gian theo VN
        df["Phát sinh"] = df["Thời gian phát sinh (UTC ISO)"].apply(utc_iso_to_vn_str)
        df["Hoàn thành"] = df["Thời gian hoàn thành (UTC ISO)"].apply(utc_iso_to_vn_str)

        # Cột hiển thị
        view_cols = [
            "Tên công ty","ShĐ","Nguyên nhân đầu vào","Tình trạng",
            "Cách xử lý","Phát sinh","Hoàn thành","KTV","SLA_gio"
        ]
        df = df.sort_values(by="Phát sinh", ascending=False, na_position="last")

        st.dataframe(df[view_cols], use_container_width=True, hide_index=True)
        st.caption(f"🧮 Tổng ticket: {len(df)} | Đã tính SLA (giờ) cho {df['SLA_gio'].notna().sum()} ticket hoàn thành")

        # Xuất file
        colx1, colx2 = st.columns(2)
        with colx1:
            st.download_button(
                "⬇️ Xuất CSV (UTF-8)",
                data=to_csv_bytes(df[view_cols]),
                file_name="baocao_helpdesk.csv",
                mime="text/csv"
            )
        with colx2:
            st.download_button(
                "⬇️ Xuất Excel",
                data=to_excel_bytes(df[view_cols]),
                file_name="baocao_helpdesk.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

st.caption("💡 Mẹo: Muốn team cùng nhập, chạy `streamlit run app.py --server.address 0.0.0.0` rồi share URL LAN.")
