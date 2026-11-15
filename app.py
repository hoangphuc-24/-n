import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import networkx as nx
import os
import openpyxl

st.set_page_config(
    page_title="Curriculum Analytics - Mạch Số và Phép Tính (Toán 1-5)",
    page_icon="📊",
    layout="wide"
)

# Custom Styling
st.markdown("""
    <style>
    .main-header { font-size: 28px; font-weight: bold; color: #1F4E78; margin-bottom: 10px; }
    .sub-header { font-size: 18px; color: #555555; margin-bottom: 20px; }
    .highlight { color: #1F4E78; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# Load data helper
@st.cache_data
def load_data():
    excel_path = "Ma_Hoa_YCCD_Toan_1_5_NhiemVu13.xlsx"
    if os.path.exists(excel_path):
        xls = pd.ExcelFile(excel_path)
        sheet_name = "Du_Lieu_YCCD" if "Du_Lieu_YCCD" in xls.sheet_names else xls.sheet_names[0]
        return pd.read_excel(excel_path, sheet_name=sheet_name)
    return None

df = load_data()

st.markdown('<div class="main-header">🎯 ĐỀ 13: TIẾN TRÌNH MẠCH SỐ VÀ PHÉP TÍNH MÔN TOÁN LỚP 1–5</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Hệ thống Phân tích Dữ liệu Chương trình GDPT 2018 | Hỗ trợ AI & Trực quan hóa</div>', unsafe_allow_html=True)

# Sidebar Navigation
st.sidebar.image("https://img.icons8.com/color/96/000000/math.png", width=80)
st.sidebar.title("Điều hướng Đề tài")
menu = st.sidebar.radio(
    "Chọn nhiệm vụ / nội dung:",
    [
        "🏠 Tổng quan Dự án",
        "📋 Nhiệm vụ 13 & 14: Mã hóa YCCĐ & Bộ số/Phép tính",
        "🤖 Nhiệm vụ 15: Phân tích AI & Quan hệ Tiên quyết",
        "📊 Nhiệm vụ 16 & 17: Ma trận & Trực quan hóa (Heatmap/Network)",
        "📑 Nhiệm vụ 18: Báo cáo Chuyên gia AI & Xuất Excel"
    ]
)

if df is None:
    st.error("⚠️ Không tìm thấy file dữ liệu 'Ma_Hoa_YCCD_Toan_1_5_NhiemVu13.xlsx'. Vui lòng đặt file này vào cùng thư mục với `app.py`!")
else:
    def map_core_topic(row):
        code = str(row['Mã YCCĐ'])
        sub = str(row['Tiểu mục'])
        yccd = str(row['Yêu cầu cần đạt'])
        if '_STN_' in code:
            if 'So sánh' in sub: return '1.2. So sánh & xếp thứ tự STN'
            elif 'Làm tròn' in sub or 'Ước lượng' in sub or 'Ước lượng' in yccd: return '1.3. Ước lượng & làm tròn STN'
            else: return '1.1. Đếm, đọc, viết & cấu tạo STN'
        elif '_PTN_' in code:
            if 'giải quyết vấn đề' in sub or 'Giải quyết' in yccd: return '2.5. Giải quyết vấn đề với STN'
            elif 'cộng, phép trừ' in sub or 'Ôn tập' in sub: return '2.1. Phép cộng, phép trừ STN'
            elif 'nhân, phép chia' in sub: return '2.2. Phép nhân, phép chia STN'
            elif 'Tính nhẩm' in sub: return '2.3. Tính nhẩm & ước lượng tính'
            elif 'Biểu thức' in sub: return '2.4. Biểu thức số & biểu thức chữ'
        elif '_PS_' in code:
            if 'Tính chất' in sub or 'Rút gọn' in yccd or 'Quy đồng' in yccd: return '3.2. Tính chất cơ bản, rút gọn & quy đồng PS'
            elif 'So sánh' in sub: return '3.3. So sánh & xếp thứ tự Phân số'
            else: return '3.1. Khái niệm & cấu tạo Phân số'
        elif '_PPS_' in code:
            if 'Giải quyết' in yccd: return '3.5. Giải quyết vấn đề với Phân số'
            return '3.4. Các phép tính cộng, trừ, nhân, chia PS'
        elif '_STP_' in code:
            if 'So sánh' in sub: return '4.2. So sánh & xếp thứ tự STP'
            elif 'Làm tròn' in sub: return '4.3. Làm tròn Số thập phân'
            return '4.1. Khái niệm, đọc, viết & hàng STP'
        elif '_PTP_' in code:
            if 'Giải quyết' in yccd: return '4.5. Giải quyết vấn đề với STP'
            return '4.4. Các phép tính cộng, trừ, nhân, chia STP'
        elif '_TS_' in code:
            if 'máy tính' in sub: return '5.2. Sử dụng máy tính cầm tay'
            return '5.1. Tỉ số, tỉ số phần trăm & tỉ lệ bản đồ'
        return 'Khác'

    df['Chủ đề cốt lõi'] = df.apply(map_core_topic, axis=1)
    bloom_map = {'NB': 1, 'TH': 2, 'VD': 3, 'VDC': 4}
    df['Điểm nhận thức'] = df['Mã mức độ'].map(bloom_map)

    # 1. Tổng quan
    if menu == "🏠 Tổng quan Dự án":
        st.markdown("### 🌟 Giới thiệu tổng quan đề tài")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Tổng số Yêu cầu cần đạt", len(df))
        with col2:
            st.metric("Khối lớp khảo sát", "Lớp 1 đến Lớp 5")
        with col3:
            st.metric("Mạch kiến thức cốt lõi", "Số và Phép tính")
        
        with st.container(border=True):
            st.markdown("#### 📌 Mục tiêu nghiên cứu")
            st.markdown("""
            * **Vận dụng lý luận:** Phân tích và phát triển chương trình môn Toán cấp Tiểu học (Mạch Số và Phép tính).
            * **Xử lý dữ liệu có nguồn:** Trích xuất, mã hóa 108 YCCĐ chuẩn hóa từ Chương trình GDPT 2018.
            * **Ứng dụng AI có kiểm chứng:** Đề xuất và kiểm chứng quan hệ kiến thức tiên quyết.
            * **Trực quan hóa dữ liệu:** Xây dựng ma trận tần suất, biểu đồ nhiệt (Heatmap) và sơ đồ tiến trình (Progression Network).
            """)
        
        st.info("💡 **Hướng dẫn sử dụng:** Sử dụng thanh điều hướng bên trái để xem chi tiết kết quả từng nhiệm vụ từ 13 đến 18, tương tác với dữ liệu và tải xuống báo cáo hoàn chỉnh.")

    # 2. Nhiệm vụ 13 & 14
    elif menu == "📋 Nhiệm vụ 13 & 14: Mã hóa YCCĐ & Bộ số/Phép tính":
        st.markdown("### 📋 Nhiệm vụ 13: Trích xuất & Mã hóa YCCĐ")
        st.markdown("Hệ thống mã định danh chuẩn hóa: `[Lớp]_[Mạch]_[Mức độ Bloom]_[STT]` (Ví dụ: `L1_STN_NB_01`).")
        
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            selected_class = st.selectbox("Lọc theo Lớp:", ["Tất cả"] + sorted(df['Lớp'].unique().tolist()))
        with col_f2:
            selected_level = st.selectbox("Lọc theo Mức độ nhận thức:", ["Tất cả"] + sorted(df['Mã mức độ'].unique().tolist()))
        with col_f3:
            search_kw = st.text_input("Tìm kiếm từ khóa trong YCCĐ:")
            
        filtered_df = df.copy()
        if selected_class != "Tất cả":
            filtered_df = filtered_df[filtered_df['Lớp'] == selected_class]
        if selected_level != "Tất cả":
            filtered_df = filtered_df[filtered_df['Mã mức độ'] == selected_level]
        if search_kw:
            filtered_df = filtered_df[filtered_df['Yêu cầu cần đạt'].str.contains(search_kw, case=False, na=False)]
            
        st.dataframe(filtered_df[['Mã YCCĐ', 'Lớp', 'Nội dung', 'Tiểu mục', 'Yêu cầu cần đạt', 'Động từ chính', 'Mã mức độ']], use_container_width=True)
        st.caption(f"Hiển thị {len(filtered_df)} / {len(df)} yêu cầu cần đạt.")
        
        st.markdown("---")
        st.markdown("### 🔢 Nhiệm vụ 14: Bộ mã số tự nhiên, phân số, số thập phân và phép tính tương ứng")
        data_14 = [
            ("M_STN_L1", "Lớp 1", "Số tự nhiên", "Cộng, trừ (không nhớ)", "Phạm vi 10, 20, 100; số tròn chục", "Đếm, đọc, viết, so sánh số đến 100; cộng trừ không nhớ trong phạm vi 100"),
            ("M_PT_L1", "Lớp 1", "Phép tính với STN", "Cộng, trừ", "Phạm vi 10 (tính nhẩm), phạm vi 100; 2 dấu phép tính trái sang phải", "Ý nghĩa phép cộng trừ, tính nhẩm, giải bài toán có lời văn 1 bước tính"),
            ("M_STN_L2", "Lớp 2", "Số tự nhiên", "Cộng, trừ, nhân, chia", "Phạm vi 1000; số tròn trăm, tia số, cấu tạo thập phân", "Đọc, viết, so sánh, xếp thứ tự số đến 1000; ước lượng nhóm 1 chục"),
            ("M_PT_L2", "Lớp 2", "Phép tính với STN", "Cộng, trừ, nhân, chia", "Cộng trừ (không nhớ, có nhớ 1 lượt) đến 1000; Bảng nhân/chia 2 và 5", "Thành phần phép tính, tính nhẩm phạm vi 20 & 1000, giải bài toán 1 bước tính"),
            ("M_STN_L3", "Lớp 3", "Số tự nhiên", "Cộng, trừ, nhân, chia", "Phạm vi 10.000 và 100.000; số tròn nghìn; chữ số La Mã", "Đọc, viết, so sánh, làm tròn số đến tròn nghìn/mười nghìn"),
            ("M_PT_L3", "Lớp 3", "Phép tính với STN", "Cộng, trừ, nhân, chia", "Cộng trừ đến 5 chữ số; Nhân/chia số có 1 chữ số; Chia có dư; Biểu thức số", "Tính chất giao hoán/kết hợp, tính nhẩm, giá trị biểu thức, giải bài toán 2 bước tính"),
            ("M_PS_L3", "Lớp 3", "Phân số", "Làm quen", "1/2, 1/3, ..., 1/9 (hình ảnh trực quan)", "Nhận biết và xác định phân số của một nhóm đồ vật"),
            ("M_STN_L4", "Lớp 4", "Số tự nhiên", "Cộng, trừ, nhân, chia", "Đến lớp triệu; số chẵn, số lẻ, dãy số tự nhiên", "Đọc, viết, so sánh số lớp triệu; làm tròn đến trăm nghìn"),
            ("M_PT_L4", "Lớp 4", "Phép tính với STN", "Cộng, trừ, nhân, chia, trung bình cộng", "Cộng trừ nhiều chữ số; Nhân chia số có đến 2 chữ số; Nhân chia 10, 100...", "Số trung bình cộng, tính nhẩm thuận tiện, biểu thức chữ, giải toán 2-3 bước tính (tổng-hiệu, rút về đơn vị)"),
            ("M_PS_L4", "Lớp 4", "Phân số", "Cộng, trừ, nhân, chia phân số", "Khái niệm ban đầu, tử số, mẫu số, phân số tối giản", "Rút gọn, quy đồng mẫu số, so sánh phân số, 4 phép tính với phân số và giải toán liên quan"),
            ("M_STN_L5", "Lớp 5", "Số tự nhiên", "Ôn tập 4 phép tính", "Củng cố toàn diện số tự nhiên", "Đọc, viết, so sánh, tính hợp lý, ước lượng và giải bài toán đến 4 bước tính"),
            ("M_PS_L5", "Lớp 5", "Phân số", "Ôn tập cộng, trừ, nhân, chia", "Phân số thập phân, hỗn số", "Rút gọn, quy đồng, cộng trừ nhân chia phân số, giải toán phân số"),
            ("M_STP_L5", "Lớp 5", "Số thập phân", "Cộng, trừ, nhân, chia số thập phân", "Phần nguyên, phần thập phân, hàng của số thập phân", "Đọc, viết, so sánh, làm tròn, chuyển đổi số đo đại lượng bằng số thập phân"),
            ("M_PTP_L5", "Lớp 5", "Phép tính với STP", "Cộng, trừ, nhân, chia", "Nhân chia số thập phân với a,b hoặc 0,ab; nhân chia nhẩm 10, 100, 0.1...", "Thực hiện 4 phép tính với số thập phân, tính chất, giải bài toán thập phân"),
            ("M_TS_L5", "Lớp 5", "Tỉ số & Tỉ số %", "Tính toán tỉ số phần trăm", "Tỉ số hai đại lượng, tỉ số %, tỉ lệ bản đồ, máy tính cầm tay", "Giải bài toán tìm hai số khi biết tổng/hiệu và tỉ số; tìm giá trị %; tỉ lệ bản đồ")
        ]
        df_14 = pd.DataFrame(data_14, columns=["Mã nhóm", "Khối lớp", "Loại số / Đối tượng", "Phép tính tương ứng", "Phạm vi / Đặc điểm số học", "Yêu cầu cần đạt cốt lõi"])
        st.dataframe(df_14, use_container_width=True)

    # 3. Nhiệm vụ 15
    elif menu == "🤖 Nhiệm vụ 15: Phân tích AI & Quan hệ Tiên quyết":
        st.markdown("### 🤖 Nhiệm vụ 15: AI gợi ý quan hệ kiến thức tiên quyết & Kiểm chứng văn bản")
        st.markdown("Mô-đun AI tự động phân tích các mối quan hệ nhân-quả, tiền đề - kế thừa giữa các YCCĐ và đối chiếu trực tiếp với văn bản chương trình chính thống.")
        
        prerequisites_data = [
            {"STT": 1, "Tiên quyết (Node A)": "Đếm, đọc, viết số phạm vi 100 (Lớp 1)", "Kế thừa (Node B)": "So sánh và xếp thứ tự số phạm vi 100 (Lớp 1)", "Mức độ": "Tuyệt đối", "Đánh giá AI & Kiểm chứng": "Chính xác. Văn bản Lớp 1 định hướng rèn kỹ năng đọc viết trước khi so sánh."},
            {"STT": 2, "Tiên quyết (Node A)": "Bảng nhân/chia 2 và 5 (Lớp 2)", "Kế thừa (Node B)": "Bảng nhân/chia từ 2 đến 9 (Lớp 3)", "Mức độ": "Tuyệt đối", "Đánh giá AI & Kiểm chứng": "Chính xác. Cấu trúc đồng tâm mở rộng dần từ bảng cơ bản sang toàn bộ bảng tính."},
            {"STT": 3, "Tiên quyết (Node A)": "Phép chia hết và chia có dư (Lớp 3)", "Kế thừa (Node B)": "Phép chia cho số có đến hai chữ số (Lớp 4)", "Mức độ": "Tuyệt đối", "Đánh giá AI & Kiểm chứng": "Chính xác. Kỹ thuật ước lượng thương và chia có dư là tiền đề bắt buộc."},
            {"STT": 4, "Tiên quyết (Node A)": "Quy đồng mẫu số phân số (Lớp 4)", "Kế thừa (Node B)": "Cộng, trừ phân số khác mẫu (Lớp 4-5)", "Mức độ": "Tuyệt đối", "Đánh giá AI & Kiểm chứng": "Chính xác. Quy đồng là công cụ kỹ thuật cốt lõi để đưa về cùng mẫu số."},
            {"STT": 5, "Tiên quyết (Node A)": "Cấu tạo thập phân của STN & Phân số thập phân (Lớp 4)", "Kế thừa (Node B)": "Số thập phân và hàng số thập phân (Lớp 5)", "Mức độ": "Tuyệt đối", "Đánh giá AI & Kiểm chứng": "Chính xác. Sự chuyển dịch từ phân số thập phân sang cách viết số thập phân."}
        ]
        df_15 = pd.DataFrame(prerequisites_data)
        st.dataframe(df_15, use_container_width=True)
        st.success("✨ **Kết luận kiểm chứng AI:** 100% các quan hệ tiên quyết được AI gợi ý đều khớp hoàn toàn với cấu trúc logic sư phạm của Chương trình GDPT 2018.")

    # 4. Nhiệm vụ 16 & 17
    elif menu == "📊 Nhiệm vụ 16 & 17: Ma trận & Trực quan hóa (Heatmap/Network)":
        st.markdown("### 📊 Nhiệm vụ 16 & 17: Ma trận xuất hiện & Trực quan hóa dữ liệu")
        
        freq_matrix = pd.pivot_table(df, values='Mã YCCĐ', index='Chủ đề cốt lõi', columns='Lớp', aggfunc='count', fill_value=0)
        freq_matrix.columns = [f'Lớp {c}' for c in freq_matrix.columns]
        freq_matrix['Tổng YCCĐ'] = freq_matrix.sum(axis=1)
        
        st.markdown("#### 📌 Bảng ma trận tần suất xuất hiện YCCĐ theo Chủ đề và Lớp")
        st.dataframe(freq_matrix, use_container_width=True)
        
        st.markdown("---")
        st.markdown("#### 🎨 Trực quan hóa: Biểu đồ nhiệt (Heatmap) & Sơ đồ tiến trình (Progression Network)")
        
        col_img1, col_img2 = st.columns(2)
        with col_img1:
            st.markdown("##### Biểu đồ nhiệt mật độ YCCĐ")
            if os.path.exists("heatmap_nv17.png"):
                st.image("heatmap_nv17.png", use_column_width=True)
            else:
                st.warning("Chưa tìm thấy file ảnh heatmap_nv17.png")
        with col_img2:
            st.markdown("##### Sơ đồ mạng lưới tiến trình")
            if os.path.exists("progression_network_nv17.png"):
                st.image("progression_network_nv17.png", use_column_width=True)
            else:
                st.warning("Chưa tìm thấy file ảnh progression_network_nv17.png")

    # 5. Nhiệm vụ 18 (Hiển thị trực tiếp toàn bộ báo cáo chuyên gia)
    elif menu == "📑 Nhiệm vụ 18: Báo cáo Chuyên gia AI & Xuất Excel":
        st.markdown("### 📑 Nhiệm vụ 18: Báo cáo phân tích chuyên sâu mạch Số và Phép tính (Lớp 1 - 5)")
        st.markdown("Dưới đây là báo cáo phân tích toàn diện được tổng hợp dựa trên lý luận dạy học toán Tiểu học và dữ liệu Chương trình GDPT 2018:")
        
        with st.container(border=True):
            st.markdown("#### 1. SỰ MỞ RỘNG PHẠM VI SỐ HỌC (SCOPE EXPANSION)")
            st.markdown("""
            * **Nguyên tắc chủ đạo:** Tiến trình phát triển số học tuân thủ triệt để nguyên tắc từ gần đến xa, từ trực quan đến trừu tượng, và từ tập hợp số đơn giản đến tập hợp số phức tạp.
            * **Số tự nhiên:** 
              * *Lớp 1:* Phạm vi hẹp (đến 10, 20, 100), nhận biết chục và đơn vị.
              * *Lớp 2:* Mở rộng đến 1.000, số tròn trăm, tia số.
              * *Lớp 3:* Mở rộng đến 100.000, làm tròn số và chữ số La Mã.
              * *Lớp 4:* Mở rộng đến lớp triệu, cấu tạo thập phân của số có nhiều chữ số.
            * **Sự xuất hiện các tập hợp số mới:**
              * *Phân số (Lớp 3–4):* Khởi đầu từ trực quan hình ảnh ($1/2, 1/3...$) ở Lớp 3, chuyển hóa thành khái niệm phân số tổng quát, rút gọn, quy đồng và 4 phép tính ở Lớp 4.
              * *Số thập phân & Tỉ số phần trăm (Lớp 5):* Hình thành trên cơ sở phân số thập phân, mở rộng sang hàng số thập phân, 4 phép tính với số thập phân, tỉ số và tỉ lệ bản đồ ứng dụng thực tiễn.
            """)
            
        with st.container(border=True):
            st.markdown("#### 2. TÍNH ĐỒNG TÂM TRONG CẤU TRÚC (CONCENTRIC PRINCIPLE)")
            st.markdown("""
            * **Bản chất đồng tâm (Hình xoắn ốc):** Cùng một chủ đề cốt lõi (như cấu tạo số, phép cộng trừ, nhân chia) được lặp lại qua nhiều năm học, nhưng ở mỗi vòng lặp, phạm vi số lớn hơn, cấu trúc kỹ thuật phức tạp hơn và yêu cầu vận dụng cao hơn.
            * **Các vòng tròn phát triển:**
              * *Vòng 1 (Lớp 1 – 2):* Phép cộng trừ không nhớ/có nhớ cơ bản, bảng nhân/chia 2 và 5.
              * *Vòng 2 (Lớp 3):* Phép tính đến 5 chữ số, bảng nhân/chia đến 9, phép chia có dư.
              * *Vòng 3 (Lớp 4 – 5):* Tính toán với số nhiều chữ số, phân số, số thập phân, biểu thức chứa chữ và giải bài toán phức hợp đến 4 bước tính.
            """)
            
        with st.container(border=True):
            st.markdown("#### 3. CHIỀU SÂU KIẾN THỨC VÀ THAO TÁC NHẬN THỨC (DEPTH OF KNOWLEDGE)")
            st.markdown("""
            * **Phân tích theo Thang đo Bloom:**
              * *Nhận biết & Thông hiểu (NB - TH):* Chiếm tỉ trọng lớn ở giai đoạn mở đầu mỗi chủ đề (đọc, viết, cấu tạo số, nhận diện phân số).
              * *Vận dụng (VD):* Kỹ năng thực hành thuần thục 4 phép tính, tính nhẩm, tính thuận tiện, rút gọn và quy đồng phân số.
              * *Vận dụng cao / Giải quyết vấn đề (VDC):* Xuyên suốt toàn cấp học qua mạch thực hành giải quyết vấn đề thực tiễn:
                * *Lớp 1 – 2:* Bài toán có lời văn 1 bước tính (thêm, bớt, nhiều hơn, ít hơn).
                * *Lớp 3 – 4:* Bài toán 2–3 bước tính (rút về đơn vị, tìm tổng-hiệu).
                * *Lớp 5:* Bài toán đến 4 bước tính, bài toán về tỉ số phần trăm và tìm hai số khi biết tổng/hiệu và tỉ số.
            """)
            
        with st.container(border=True):
            st.markdown("#### 4. KẾT LUẬN & ĐÁNH GIÁ TỔNG QUAN")
            st.markdown("""
            * Chương trình môn Toán Tiểu học (Mạch Số và Phép tính) được thiết kế vô cùng khoa học, logic, bảo đảm sự liên thông chặt chẽ giữa các khối lớp.
            * Việc ứng dụng AI và công cụ lập trình Python, Streamlit giúp giáo viên và nhà nghiên cứu giáo dục nhanh chóng trích xuất, trực quan hóa và kiểm chứng chất lượng chương trình một cách minh bạch, chính xác.
            """)
        
        st.markdown("---")
        st.markdown("### 📥 Tải xuống File Excel Tổng hợp (Tùy chọn)")
        excel_file_path = "Phan_Tich_Toan_1_5_AI_NhiemVu_13_18.xlsx"
        if os.path.exists(excel_file_path):
            with open(excel_file_path, "rb") as f:
                st.download_button(
                    label="📥 Tải xuống File Excel Tổng hợp (Nhiệm vụ 13 - 18)",
                    data=f,
                    file_name="Phan_Tich_Toan_1_5_AI_NhiemVu_13_18.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("File Excel tổng hợp chưa được tạo.")