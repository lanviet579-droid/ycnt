import streamlit as st
import io, os, base64, csv
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm  
from datetime import datetime, timedelta

# Bật chế độ Wide để tận dụng tối đa màn hình
st.set_page_config(page_title="XUẤT PDF", page_icon="📄", layout="wide")

st.markdown("### 📄 App gửi YCNT(đổi mẫuLH Bien 0903585579")

# Khởi tạo bộ nhớ tạm
if 'stt_num' not in st.session_state:
    st.session_state.stt_num = 1
if 'pdf_xuat' not in st.session_state:
    st.session_state.pdf_xuat = None
if 'ten_file_pdf' not in st.session_state:
    st.session_state.ten_file_pdf = "YCNT.pdf"
if 'lich_su' not in st.session_state:
    st.session_state.lich_su = []
if 'lich_su_full' not in st.session_state:
    st.session_state.lich_su_full = []

# ==========================================
# --- 2 HÀM XỬ LÝ DỮ LIỆU ---
# ==========================================
def tao_file_excel():
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['Số Phiếu', 'Ngày Lập', 'Giờ NT', 'Ngày NT', 'Nội dung', 'Vị trí', 'Chỉ huy trưởng', 'Kỹ thuật'])
    for item in reversed(st.session_state.lich_su_full):
        writer.writerow([item['stt'], item['nl'], item['gnt'], item['nnt'], item['nd'], item['vt'], item['ch'], item['ktnt']])
    return '\ufeff' + output.getvalue()

def ghi_len_google_sheets(url_sheet, data_row):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        if "gcp_service_account" not in st.secrets:
            return False, "Chưa cài mã API Google trong phần Secrets của Streamlit."
        credentials_dict = dict(st.secrets["gcp_service_account"])
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(credentials_dict, scopes=scopes)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(url_sheet).sheet1
        sheet.append_row(data_row)
        return True, "Đồng bộ thành công!"
    except Exception as e:
        return False, f"Lỗi Google Sheets: {str(e)}"

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
    
    def in_chu_mm(x_mm, y_mm, text, max_width_mm, font_size=11, khoang_cach_dong_mm=5.0):
        if not text: return y_mm 
        can.setFont(font_name, font_size)
        y = y_mm * mm
        x = x_mm * mm
        max_width = max_width_mm * mm
        khoang_cach_dong = khoang_cach_dong_mm * mm
        cac_doan = str(text).split('\n') 
        for doan in cac_doan:
            tu_ngu = doan.split(' ') 
            dong_hien_tai = ""
            for tu in tu_ngu:
                test_line = dong_hien_tai + tu + " "
                do_rong = pdfmetrics.stringWidth(test_line, font_name, font_size)
                if do_rong <= max_width: 
                    dong_hien_tai = test_line
                else: 
                    if dong_hien_tai.strip():
                        can.drawString(x, y, dong_hien_tai.strip())
                        y -= khoang_cach_dong
                    dong_hien_tai = tu + " "
            if dong_hien_tai.strip():
                can.drawString(x, y, dong_hien_tai.strip())
                y -= khoang_cach_dong
        return y / mm

    in_chu_mm(x_mm=163.0, y_mm=242.5, text=d['stt'], max_width_mm=40) 
    in_chu_mm(x_mm=148.0, y_mm=228, text=d['nl'], max_width_mm=40)  
    so_tt_ngan = d['stt'].split('/')[0] if '/' in d['stt'] else d['stt']
    in_chu_mm(x_mm=22.0, y_mm=204.0, text=so_tt_ngan, max_width_mm=15) 
    y_sau_noi_dung = in_chu_mm(x_mm=35.0, y_mm=212.0, text=d['nd'], max_width_mm=72) 
    in_chu_mm(x_mm=35.0, y_mm=y_sau_noi_dung - 1.0, text=f" {d['vt']}", max_width_mm=72) 
    in_chu_mm(x_mm=116.0, y_mm=212.0, text=d['gnt'], max_width_mm=15) 
    in_chu_mm(x_mm=112.0, y_mm=200.0, text=d['nnt'], max_width_mm=15) 
    in_chu_mm(x_mm=137.0, y_mm=212.0, text=d['ktnt'], max_width_mm=35) 
    in_chu_mm(x_mm=130.0, y_mm=152.0, text=d['ch'], max_width_mm=50, font_size=12) 

    can.save()
    packet.seek(0)
    if not os.path.exists("Mau_YCNT.pdf"): return None
    existing_pdf = PdfReader(open("Mau_YCNT.pdf", "rb"))
    output = PdfWriter()
    page = existing_pdf.pages[0]
    page.merge_page(PdfReader(packet).pages[0])
    output.add_page(page)
    for i in range(1, len(existing_pdf.pages)):
        output.add_page(existing_pdf.pages[i])
    pdf_bytes = io.BytesIO()
    output.write(pdf_bytes)
    return pdf_bytes.getvalue()

col_nhap, col_xem = st.columns([1.2, 1])

with col_nhap:
    with st.expander("⚡ NHẬP NHANH TỪ ZALO", expanded=False):
        fast_txt = st.text_area("📋 Dán nội dung Zalo vào đây:", height=100, label_visibility="collapsed")

    data = {"nl": datetime.now().strftime("%d/%m/%Y"), "gnt": "08:30", "nnt": "", "nd": "", "vt": "", "ch": "Nguyễn Hữu Biên", "ktnt": "Nguyễn Văn A"}

    if fast_txt:
        try:
            parts = fast_txt.split(";")
            for p in parts:
                if ":" in p:
                    k, v = p.split(":", 1)
                    k, v = k.strip(), v.strip()
                    if "Ngày NT" in k: 
                        data["nnt"] = v
                        try:
                            v_clean = v.replace('-', '/').replace('.', '/')
                            dt_nnt = datetime.strptime(v_clean, "%d/%m/%Y")
                            data["nl"] = (dt_nnt - timedelta(days=1)).strftime("%d/%m/%Y")
                        except ValueError: pass
                    elif "Giờ NT" in k: data["gnt"] = v
                    elif "Địa điểm NT" in k: data["vt"] = v
                    elif "Vị trí" in k: data["vt"] = v
                    elif "Nội dung NT" in k: data["nd"] = v
                    elif "CHT" in k: data["ch"] = v
                    elif "KT" in k: data["ktnt"] = v
        except: pass

    st.markdown("### 📝 CHI TIẾT PHIẾU")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Số Phiếu**")
        c_tru, c_so, c_cong = st.columns([1, 2, 1])
        with c_tru:
            if st.button("➖"): st.session_state.stt_num = max(1, st.session_state.stt_num - 1); st.rerun()
        with c_cong:
            if st.button("➕"): st.session_state.stt_num += 1; st.rerun()
        with c_so:
            stt_str = f"{st.session_state.stt_num:02d}/Dacinco/ĐNCNC" 
            stt = st.text_input("STT", value=stt_str, label_visibility="collapsed")
        gnt = st.text_input("Giờ NT", value=data["gnt"])
        ch = st.text_input("Chỉ huy trưởng", value=data["ch"])

    with c2:
        nl = st.text_input("Ngày lập (Tự lùi 1 ngày)", value=data["nl"])
        nnt = st.text_input("Ngày NT", value=data["nnt"])
        ktnt = st.text_input("Kỹ thuật - SĐT", value=data["ktnt"])

    nd = st.text_area("Nội dung", value=data["nd"], height=80)
    vt = st.text_area("Vị trí", value=data["vt"], height=80)

    # --- BẢN GHIM CỨNG LINK GOOGLE SHEETS ---
    with st.expander("⚙️ TÍCH HỢP GOOGLE SHEETS (ĐÃ GHIM)", expanded=False):
        bat_gsheets = st.checkbox("Bật tự động lưu lên Google Sheets", value=True) # Mặc định BẬT
        link_co_dinh = "https://docs.google.com/spreadsheets/d/1r7SJ64Ht5Tnrg3mmiq0eSYhX7OhKNaD7kN1KJOoFtuQ/edit?gid=0#gid=0"
        link_gsheets = st.text_input("Link Google Sheets:", value=link_co_dinh)

    st.markdown("---")
    
    if st.button("🔥 XUẤT PDF & XEM TRƯỚC", use_container_width=True, type="primary"):
        final_data = {"stt": stt, "nl": nl, "gnt": gnt, "nnt": nnt, "nd": nd, "vt": vt, "ch": ch, "ktnt": ktnt}
        pdf_out = create_final_pdf(final_data)
        if pdf_out:
            st.session_state.pdf_xuat = pdf_out
            so_phieu_thuc_te = stt.split('/')[0] if '/' in stt else stt
            st.session_state.ten_file_pdf = f"YCNT_{so_phieu_thuc_te}.pdf"
            st.session_state.lich_su.insert(0, {"stt": so_phieu_thuc_te, "ngay": nnt, "nd": nd})
            st.session_state.lich_su_full.append(final_data)
            
            msg_gsheets = ""
            if bat_gsheets and link_gsheets:
                row_data = [so_phieu_thuc_te, nl, gnt, nnt, nd, vt, ch, ktnt]
                ok, msg = ghi_len_google_sheets(link_gsheets, row_data)
                if ok: msg_gsheets = f" - ☁️ {msg}"
                else: st.warning(f"⚠️ {msg}")
            
            st.success(f"✅ Đã tạo xong!{msg_gsheets}")
        else:
            st.error("❌ Thiếu file Mau_YCNT.pdf trong thư mục GitHub!")
            
    st.markdown("### 🕒 NHẬT KÝ ĐÃ GỬI")
    if len(st.session_state.lich_su_full) > 0:
        st.download_button(label="📥 TẢI FILE EXCEL LỊCH SỬ (CSV)", data=tao_file_excel().encode('utf-8'), file_name=f"LichSu_YCNT_{datetime.now().strftime('%d%m%Y')}.csv", mime="text/csv", use_container_width=True)

    with st.container(height=300):
        if len(st.session_state.lich_su) == 0: st.caption("Chưa xuất phiếu nào...")
        else:
            for item in st.session_state.lich_su:
                st.markdown(f"**Phiếu {item['stt']}** - {item['ngay']}")
                st.caption(f"{item['nd']}")
                st.divider()

with col_xem:
    st.markdown("### 👁️ XEM TRƯỚC PDF")
    if st.session_state.pdf_xuat:
        st.download_button(label=f"📥 TẢI XUỐNG {st.session_state.ten_file_pdf}", data=st.session_state.pdf_xuat, file_name=st.session_state.ten_file_pdf, mime="application/pdf", use_container_width=True)
        if st.button("🔄 CHUYỂN SANG PHIẾU TIẾP THEO", use_container_width=True):
            st.session_state.stt_num += 1
            st.session_state.pdf_xuat = None
            st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)
        hien_thi_pdf(st.session_state.pdf_xuat)
    else: st.info("Bấm 'Xuất PDF & Xem trước' để hiển thị nội dung tại đây.")
