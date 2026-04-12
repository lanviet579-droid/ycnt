import streamlit as st
import io, os, base64, csv
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm  
from datetime import datetime, timedelta

# Cấu hình giao diện Wide
st.set_page_config(page_title="HỆ THỐNG YCNT - MR BIEN", page_icon="📄", layout="wide")

st.markdown("### 📄 gửi YCNT Đa Dự Án-Đổi mẫu LH Bien 0903585579")

# --- KHỞI TẠO BỘ NHỚ (LOGIC GỐC) ---
if 'stt_num' not in st.session_state:
    st.session_state.stt_num = 1
if 'pdf_xuat' not in st.session_state:
    st.session_state.pdf_xuat = None
if 'lich_su_full' not in st.session_state:
    st.session_state.lich_su_full = []

# ==========================================
# --- CẤU HÌNH DỰ ÁN & TỌA ĐỘ CHI TIẾT ---
# ==========================================
with st.sidebar:
    st.markdown("### 🏢 THIẾT LẬP DỰ ÁN")
    du_an_chon = st.selectbox("Chọn dự án:", ["Dacinco - ĐNCNC", "Dự án Quảng Nam", "Dự án Huế"])
    
    if du_an_chon == "Dacinco - ĐNCNC":
        ma_ban = "Dacinco/ĐNCNC"
        file_mau = "Mau_YCNT.pdf"
        cht_mac_dinh = "Nguyễn Hữu Biên"
        kt_mac_dinh = "Bùi Văn Năng 0935538496"
        link_sheet_mac_dinh = "https://docs.google.com/spreadsheets/d/1mvyOY5gTL813M2Uf8AuamfH2A7rrbJj4C5v1EhArnB4/edit?gid=0#gid=0"
        p = {
            "so_phieu": (163.0, 242.5), "ngay_gui": (148.0, 228.0), "stt_bang": (24.0, 204.0),
            "noi_dung": (35.0, 212.0), "dia_diem": (35.0, 207.0), "gio_nt": (116.0, 212.0),
            "ngay_nt": (112.0, 200.0), "ten_kt": (137.0, 212.0), "ten_cht": (130.0, 152.0)
        }
        
    elif du_an_chon == "Dự án Quảng Nam":
        ma_ban = "QN/2026"; file_mau = "Mau_QuangNam.pdf"; cht_mac_dinh = "Trần Văn A"; kt_mac_dinh = "Kỹ thuật QN"; link_sheet_mac_dinh = "LINK_SHEET_QUANG_NAM"
        p = {
            "so_phieu": (160.0, 240.0), "ngay_gui": (145.0, 225.0), "stt_bang": (22.0, 202.0),
            "noi_dung": (38.0, 210.0), "dia_diem": (38.0, 205.0), "gio_nt": (114.0, 210.0),
            "ngay_nt": (110.0, 198.0), "ten_kt": (135.0, 210.0), "ten_cht": (128.0, 150.0)
        }
    
    else:
        ma_ban = "HUE/2026"; file_mau = "Mau_Hue.pdf"; cht_mac_dinh = "Lê Văn C"; kt_mac_dinh = "Kỹ thuật Huế"; link_sheet_mac_dinh = "LINK_SHEET_HUẾ"
        p = { "so_phieu": (163.0, 242.5), "ngay_gui": (148.0, 228.0), "stt_bang": (24.0, 204.0), "noi_dung": (35.0, 212.0), "dia_diem": (35.0, 207.0), "gio_nt": (116.0, 212.0), "ngay_nt": (112.0, 200.0), "ten_kt": (137.0, 212.0), "ten_cht": (130.0, 152.0) }

# ==========================================
# --- LOGIC NHẢY SỐ THÔNG MINH ---
# ==========================================
if st.session_state.lich_su_full:
    danh_sach_so = []
    for item in st.session_state.lich_su_full:
        try:
            so = int(item['stt'].split('/')[0])
            danh_sach_so.append(so)
        except: pass
    if danh_sach_so:
        st.session_state.stt_num = max(danh_sach_so) + 1

# ==========================================
# --- HÀM XỬ LÝ (SỬA LỖI TRUYỀN THAM SỐ) ---
# ==========================================
def ghi_len_google_sheets(url_sheet, data_row):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        if "gcp_service_account" not in st.secrets:
            return False, "Chưa cài mã API Google trong phần Secrets."
        
        credentials_dict = dict(st.secrets["gcp_service_account"])
        # THÊM 'scopes=' VÀO ĐÂY ĐỂ SỬA LỖI ÔNG GỬI
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
        
        gspread.authorize(creds).open_by_url(url_sheet).sheet1.append_row(data_row, value_input_option='RAW')
        return True, "Đồng bộ thành công!"
    except Exception as e: return False, str(e)

def create_final_pdf(d, file_path, pos):
    packet = io.BytesIO()
    can = canvas.Canvas(packet)
    font_name = 'Helvetica'
    if os.path.exists("times.ttf"):
        pdfmetrics.registerFont(TTFont('Times-New-Roman', 'times.ttf'))
        font_name = 'Times-New-Roman'
    
    def in_chu(x, y, text, w, size=10):
        can.setFont(font_name, size)
        y_pt, x_pt = y*mm, x*mm
        for doan in str(text).split('\n'):
            dong = ""
            for tu in doan.split(' '):
                if pdfmetrics.stringWidth(dong + tu + " ", font_name, size) <= w*mm: dong += tu + " "
                else: can.drawString(x_pt, y_pt, dong.strip()); y_pt -= 5*mm; dong = tu + " "
            can.drawString(x_pt, y_pt, dong.strip()); y_pt -= 5*mm

    in_chu(pos["so_phieu"][0], pos["so_phieu"][1], d['stt'], 40)
    in_chu(pos["ngay_gui"][0], pos["ngay_gui"][1], d['nl'], 40)
    so_n = d['stt'].split('/')[0] if '/' in d['stt'] else d['stt']
    in_chu(pos["stt_bang"][0], pos["stt_bang"][1], so_n, 15)
    in_chu(pos["noi_dung"][0], pos["noi_dung"][1], d['nd'], 72)
    in_chu(pos["dia_diem"][0], pos["dia_diem"][1], d['vt'], 72)
    in_chu(pos["gio_nt"][0], pos["gio_nt"][1], d['gnt'], 15)
    in_chu(pos["ngay_nt"][0], pos["ngay_nt"][1], d['nnt'], 15)
    in_chu(pos["ten_kt"][0], pos["ten_kt"][1], d['ktnt'], 35)
    in_chu(pos["ten_cht"][0], pos["ten_cht"][1], d['ch'], 50, size=10)

    can.save(); packet.seek(0)
    if not os.path.exists(file_path): return None
    reader = PdfReader(open(file_path, "rb")); writer = PdfWriter()
    page = reader.pages[0]; page.merge_page(PdfReader(packet).pages[0]); writer.add_page(page)
    for i in range(1, len(reader.pages)): writer.add_page(reader.pages[i])
    out = io.BytesIO(); writer.write(out); return out.getvalue()

# ==========================================
# --- GIAO DIỆN CHÍNH ---
# ==========================================
col_nhap, col_xem = st.columns([1.2, 1])

with col_nhap:
    fast_txt = st.text_area("📋 Dán nội dung Zalo:", height=100)
    data = {"nl": datetime.now().strftime("%d/%m/%Y"), "gnt": "08:30", "nnt": "", "nd": "", "vt": "", "ch": cht_mac_dinh, "ktnt": kt_mac_dinh}
    
    if fast_txt:
        try:
            for p_zalo in fast_txt.split(";"):
                if ":" in p_zalo:
                    k, v = [i.strip() for i in p_zalo.split(":", 1)]
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
        stt_input = st.number_input("Số thứ tự", value=st.session_state.stt_num, step=1)
        if stt_input != st.session_state.stt_num: st.session_state.stt_num = stt_input; st.rerun()
        stt_full = st.text_input("STT đầy đủ", value=f"{st.session_state.stt_num:02d}/{ma_ban}")
        gnt = st.text_input("Giờ NT", value=data["gnt"])
        ch = st.text_input("Chỉ huy trưởng", value=data["ch"])
    with c2:
        nl = st.text_input("Ngày lập", value=data["nl"])
        nnt = st.text_input("Ngày NT", value=data["nnt"])
        ktnt = st.text_input("Kỹ thuật - SĐT", value=data["ktnt"])

    nd = st.text_area("Nội dung", value=data["nd"], height=80)
    vt = st.text_area("Vị trí", value=data["vt"], height=80)
    
    with st.expander("⚙️ CẤU HÌNH GOOGLE SHEETS", expanded=True):
        bat_gs = st.checkbox("Bật tự động lưu", value=True)
        link_hien_tai = st.text_input("Link Sheet đang dùng:", value=link_sheet_mac_dinh)

    if st.button("🔥 XUẤT PDF & LƯU SHEETS", use_container_width=True, type="primary"):
        f_data = {"stt": stt_full, "nl": nl, "gnt": gnt, "nnt": nnt, "nd": nd, "vt": vt, "ch": ch, "ktnt": ktnt}
        pdf_out = create_final_pdf(f_data, file_mau, p)
        if pdf_out:
            st.session_state.pdf_xuat = pdf_out
            st.session_state.lich_su_full.append(f_data)
            if bat_gs:
                ok, msg = ghi_len_google_sheets(link_hien_tai, [stt_full, nl, gnt, nnt, nd, vt, ch, ktnt])
                if ok: st.success(f"✅ Đã lưu vào Sheets!")
                else: st.error(f"Lỗi: {msg}")
        else: st.error(f"❌ Thiếu file {file_mau} trên GitHub!")

with col_xem:
    st.markdown("### 👁️ XEM TRƯỚC")
    if st.session_state.pdf_xuat:
        st.download_button(f"📥 TẢI PDF {stt_full.split('/')[0]}", data=st.session_state.pdf_xuat, file_name=f"YCNT_{stt_full.split('/')[0]}.pdf", mime="application/pdf", use_container_width=True)
        if st.button("🔄 PHIẾU TIẾP THEO", use_container_width=True):
            st.session_state.stt_num += 1; st.session_state.pdf_xuat = None; st.rerun()
        base64_pdf = base64.b64encode(st.session_state.pdf_xuat).decode('utf-8')
        st.markdown(f'<embed src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf">', unsafe_allow_html=True)
