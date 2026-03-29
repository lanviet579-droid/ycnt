import streamlit as st
import io, os
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm  
from datetime import datetime

st.set_page_config(page_title="DACICO PDF CHUẨN", page_icon="📄")
st.title("📄 DACICO-Gửi YCNT by Bien 0903585579")

if 'stt_num' not in st.session_state:
    st.session_state.stt_num = 1

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

    # =========================================================
    # TỌA ĐỘ VÀNG (ĐÃ NẮN LẠI DỰA VÀO ẢNH LỖI)
    # =========================================================
    
    # 1. Số phiếu (Nâng lên 1.5mm để bằng chữ in sẵn, kéo trái 1 chút)
    in_chu_mm(x_mm=163.0, y_mm=242.5, text=d['stt'], max_width_mm=40) 
    
    # 2. Ngày gửi (Nâng lên 1.5mm)
    in_chu_mm(x_mm=148.0, y_mm=228, text=d['nl'], max_width_mm=40)  
    
    # 3. STT (Kéo nhẹ sang trái cho cân giữa)
    so_tt_ngan = d['stt'].split('/')[0] if '/' in d['stt'] else d['stt']
    in_chu_mm(x_mm=22.0, y_mm=204.0, text=so_tt_ngan, max_width_mm=15) 
    
    # 4. Nội dung (Thụt vào trong 3mm (thành X=35) để KHÔNG LIẾM VẠCH)
    y_sau_noi_dung = in_chu_mm(x_mm=35.0, y_mm=212.0, text=d['nd'], max_width_mm=72) 
    
    # 5. Vị trí (Tự động nối tiếp)
    in_chu_mm(x_mm=35.0, y_mm=y_sau_noi_dung - 1.0, text=f" {d['vt']}", max_width_mm=72) 
    
    # 6. Thời gian - Giờ 
    in_chu_mm(x_mm=116.0, y_mm=212.0, text=d['gnt'], max_width_mm=15) 
    
    # 7. Thời gian - Ngày (Nhấc lên 1 chút cho đỡ dính lề dưới)
    in_chu_mm(x_mm=112.0, y_mm=200.0, text=d['nnt'], max_width_mm=15) 
    
    # 8. Kỹ thuật (Kéo mạnh sang trái thành 128, bóp max_width còn 28 để không tràn)
    in_chu_mm(x_mm=137.0, y_mm=212.0, text=d['ktnt'], max_width_mm=35) 
    
    # 9. Chỉ huy trưởng (Kéo sang trái 12mm (thành 130) cho chui vào giữa chữ ký)
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

# --- GIAO DIỆN NHẬP LIỆU ---
fast_txt = st.text_area("📋 Dán nội dung Zalo vào đây:", height=100)

data = {
    "nl": datetime.now().strftime("%d/%m/%Y"), 
    "gnt": "08:30", "nnt": "", "nd": "", "vt": "", 
    "ch": "Nguyễn Văn A", "ktnt": "Nguyễn Hữu B"
}

if fast_txt:
    try:
        parts = fast_txt.split(";")
        for p in parts:
            if ":" in p:
                k, v = p.split(":", 1)
                k, v = k.strip(), v.strip()
                if "Ngày NT" in k: data["nnt"] = v
                elif "Giờ NT" in k: data["gnt"] = v
                elif "Địa điểm NT" in k: data["vt"] = v
                elif "Nội dung NT" in k: data["nd"] = v
                elif "CHT" in k: data["ch"] = v
                elif "KT" in k: data["ktnt"] = v
    except: pass

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Số Phiếu**")
    c_tru, c_so, c_cong = st.columns([1, 2, 1])
    with c_tru:
        if st.button("➖"):
            st.session_state.stt_num = max(1, st.session_state.stt_num - 1)
    with c_cong:
        if st.button("➕"):
            st.session_state.stt_num += 1
    with c_so:
        stt_str = f"{st.session_state.stt_num:02d}/Dacinco/ĐNCNC" 
        stt = st.text_area("STT", value=stt_str, label_visibility="collapsed", height=68)
        
    gnt = st.text_area("Giờ NT", value=data["gnt"], height=68)

with col2:
    nl = st.text_area("Ngày lập", value=data["nl"], height=68)
    nnt = st.text_area("Ngày NT", value=data["nnt"], height=68)

nd = st.text_area("Nội dung", value=data["nd"], height=100)
vt = st.text_area("Vị trí", value=data["vt"], height=100)

col3, col4 = st.columns(2)
with col3:
    ch = st.text_area("Chỉ huy trưởng", value=data["ch"], height=68)
with col4:
    ktnt = st.text_area("Kỹ thuật - SĐT", value=data["ktnt"], height=68)

if st.button("🔥 XUẤT PDF CHUẨN DACINCO"):
    final_data = {
        "stt": stt, "nl": nl, "gnt": gnt, "nnt": nnt, 
        "nd": nd, "vt": vt, "ch": ch, "ktnt": ktnt
    }
    pdf_out = create_final_pdf(final_data)
    
    if pdf_out:
        st.download_button(
            label="📥 TẢI PDF CHUẨN & GỬI ZALO",
            data=pdf_out,
            file_name=f"YCNT_{st.session_state.stt_num:02d}.pdf",
            mime="application/pdf"
        )
        st.success("✅ Đã nắn lại toàn bộ xương khớp! Chữ nằm ngay ngắn như lính duyệt binh rồi nhé!")
    else:
        st.error("❌ Thiếu file Mau_YCNT.pdf trong thư mục!")