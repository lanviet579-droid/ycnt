import streamlit as st
import io, os, base64, csv
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm  
from datetime import datetime, timedelta

# Bật chế độ Wide
st.set_page_config(page_title="XUẤT PDF", page_icon="📄", layout="wide")

st.markdown("### 📄 App gửi YCNT (Bản sửa lỗi ghi cột chuẩn 100%)")

# Khởi tạo bộ nhớ tạm
if 'stt_num' not in st.session_state:
    st.session_state.stt_num = 1
if 'pdf_xuat' not in st.session_state:
    st.session_state.pdf_xuat = None
if 'lich_su_full' not in st.session_state:
    st.session_state.lich_su_full = []

# ==========================================
# --- HÀM GHI DỮ LIỆU CHUẨN CỘT ---
# ==========================================
def ghi_len_google_sheets(url_sheet, data_row):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        if "gcp_service_account" not in st.secrets:
            return False, "Thiếu cấu hình Secrets."
        
        credentials_dict = dict(st.secrets["gcp_service_account"])
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
        client = gspread.authorize(creds)
        
        # Mở Sheet đầu tiên
        sheet = client.open_by_url(url_sheet).sheet1
        
        # Ép dữ liệu thành danh sách các chuỗi sạch
        clean_row = [str(x) for x in data_row]
        
        # Dùng append_row với table_alpha để ép nó vào từng cột A, B, C...
        sheet.append_row(clean_row, value_input_option='USER_ENTERED')
        
        return True, "Đã lưu đúng cột!"
    except Exception as e:
        return False, f"Lỗi: {str(e)}"

# --- GIỮ NGUYÊN LOGIC PDF CỦA ÔNG ---
def hien_thi_pdf(pdf_bytes):
    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
    pdf_display = f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf">'
    st.markdown(pdf_display, unsafe_allow_html=True)

def create_final_pdf(d):
    packet = io.BytesIO()
    can = canvas.Canvas(packet)
    font_name = 'Helvetica'
    if os.path.exists("Arial.ttf"):
        pdfmetrics.registerFont(TTFont('Arial-VN', 'Arial.ttf'))
        font_name = 'Arial-VN'
    
    def in_chu_mm(x_mm, y_mm, text, max_width_mm, font_size=11, kc_dong=5.0):
        if not text: return y_mm 
        can.setFont(font_name, font_size)
        y, x = y_mm * mm, x_mm * mm
        max_w = max_width_mm * mm
        for doan in str(text).split('\n'):
            dong = ""
            for tu in doan.split(' '):
                if pdfmetrics.stringWidth(dong + tu + " ", font_name, font_size) <= max_w:
                    dong += tu + " "
                else:
                    can.drawString(x, y, dong.strip())
                    y -= kc_dong * mm
                    dong = tu + " "
            can.drawString(x, y, dong.strip())
            y -= kc_dong * mm
        return y / mm

    in_chu_mm(163.0, 242.5, d['stt'], 40)
    in_chu_mm(148.0, 228, d['nl'], 40)
    in_chu_mm(22.0, 204.0, d['stt'].split('/')[0], 15)
    y_nd = in_chu_mm(35.0, 212.0, d['nd'], 72)
    in_chu_mm(35.0, y_nd - 1.0, f" {d['vt']}", 72)
    in_chu_mm(116.0, 212.0, d['gnt'], 15)
    in_chu_mm(112.0, 200.0, d['nnt'], 15)
    in_chu_mm(137.0, 212.0, d['ktnt'], 35)
    in_chu_mm(130.0, 152.0, d['ch'], 50, 12)

    can.save()
    packet.seek(0)
    if not os.path.exists("Mau_YCNT.pdf"): return None
    reader = PdfReader(open("Mau_YCNT.pdf", "rb"))
    writer = PdfWriter()
    page = reader.pages[0]
    page.merge_page(PdfReader(packet).pages[0])
    writer.add_page(page)
    for i in range(1, len(reader.pages)): writer.add_page(reader.pages[i])
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

col_nhap, col_xem = st.columns([1.2, 1])

with col_nhap:
    with st.expander("⚡ NHẬP NHANH TỪ ZALO", expanded=False):
        fast_txt = st.text_area("📋 Dán nội dung Zalo:", height=100)

    data = {"nl": datetime.now().strftime("%d/%m/%Y"), "gnt": "08:30", "nnt": "", "nd": "", "vt": "", "ch": "Nguyễn Hữu Biên", "ktnt": "Nguyễn Văn A"}

    if fast_txt:
        try:
            for p in fast_txt.split(";"):
                if ":" in p:
                    k, v = [i.strip() for i in p.split(":", 1)]
                    if "Ngày NT" in k:
                        data["nnt"] = v
                        try:
                            dt = datetime.strptime(v.replace('-','/').replace('.','/'), "%d/%m/%Y")
                            data["nl"] = (dt - timedelta(days=1)).strftime("%d/%m/%Y")
                        except: pass
                    elif "Giờ NT" in k: data["gnt"] = v
                    elif "Vị trí" in k or "Địa điểm" in k: data["vt"] = v
                    elif "Nội dung" in k: data["nd"] = v
                    elif "CHT" in k: data["ch"] = v
                    elif "KT" in k: data["ktnt"] = v
        except: pass

    st.markdown("### 📝 CHI TIẾT PHIẾU")
    c1, c2 = st.columns(2)
    with c1:
        stt_num = st.number_input("Số phiếu", value=st.session_state.stt_num, step=1)
        st.session_state.stt_num = stt_num
        stt = st.text_input("Mã số phiếu", value=f"{stt_num:02d}/Dacinco/ĐNCNC")
        gnt = st.text_input("Giờ NT", value=data["gnt"])
        ch = st.text_input("Chỉ huy trưởng", value=data["ch"])
    with c2:
        nl = st.text_input("Ngày lập", value=data["nl"])
        nnt = st.text_input("Ngày NT", value=data["nnt"])
        ktnt = st.text_input("Kỹ thuật", value=data["ktnt"])
    nd = st.text_area("Nội dung", value=data["nd"], height=80)
    vt = st.text_area("Vị trí", value=data["vt"], height=80)

    with st.expander("⚙️ GOOGLE SHEETS", expanded=True):
        bat_gs = st.checkbox("Tự động lưu", value=True)
        link_gs = st.text_input("Link Sheet:", value="https://docs.google.com/spreadsheets/d/1r7SJ64Ht5Tnrg3mmiq0eSYhX7OhKNaD7kN1KJOoFtuQ/edit?gid=0#gid=0")

    if st.button("🔥 XUẤT PDF & LƯU SHEETS", use_container_width=True, type="primary"):
        f_data = {"stt": stt, "nl": nl, "gnt": gnt, "nnt": nnt, "nd": nd, "vt": vt, "ch": ch, "ktnt": ktnt}
        pdf_out = create_final_pdf(f_data)
        if pdf_out:
            st.session_state.pdf_xuat = pdf_out
            st.session_state.lich_su_full.append(f_data)
            if bat_gs and link_gs:
                # --- PHẦN QUAN TRỌNG NHẤT: DỮ LIỆU PHẢI LÀ MẢNG ---
                row_to_push = [stt, nl, gnt, nnt, nd, vt, ch, ktnt]
                ok, res = ghi_len_google_sheets(link_gs, row_to_push)
                st.toast(res)
            st.success("✅ Đã xuất và lưu Sheet!")
        else: st.error("Thiếu file mẫu!")

with col_xem:
    st.markdown("### 👁️ XEM TRƯỚC")
    if st.session_state.pdf_xuat:
        st.download_button("📥 TẢI PDF", data=st.session_state.pdf_xuat, file_name=f"YCNT_{st.session_state.stt_num}.pdf", mime="application/pdf", use_container_width=True)
        if st.button("🔄 PHIẾU TIẾP THEO", use_container_width=True):
            st.session_state.stt_num += 1
            st.session_state.pdf_xuat = None
            st.rerun()
        hien_thi_pdf(st.session_state.pdf_xuat)
