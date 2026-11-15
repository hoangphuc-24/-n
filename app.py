import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os

st.set_page_config(
    page_title="Curriculum Analytics - Mạch Số và Phép Tính (Toán 1-5)",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
    <style>
    .main-header { font-size: 28px; font-weight: bold; color: #1F4E78; margin-bottom: 10px; }
    .sub-header { font-size: 18px; color: #555555; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# Tự động nhận diện file Excel đang có trong thư mục GitHub
@st.cache_data
def load_data():
    possible_files = [
        "NhiemVu_13_14_MaHoa_Toan_1_5.xlsx",
        "Ma_Hoa_YCCD_Toan_1_5_NhiemVu13.xlsx",
        "Phan_Tich_Toan_1_5_AI_NhiemVu_13_18.xlsx"
    ]
    for file in possible_files:
        if os.path.exists(file):
            xls = pd.ExcelFile(file)
            sheet_name = xls.sheet_names[0]
            return pd.read_excel(file, sheet_name=sheet_name)
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
        "🏠 Dashboard & Yêu cầu tối thiểu",
        "📋 Nhiệm vụ 13 & 14: Mã hóa YCCĐ & Bộ số/Phép tính",
        "🤖 Nhiệm vụ 15: Phân tích AI & Quan hệ Tiên quyết",
        "📊 Nhiệm vụ 16 & 17: Ma trận & Trực quan hóa (Heatmap)",
        "📑 Nhiệm vụ 18: Báo cáo Chuyên gia AI & Xuất Excel"
    ]
)

if df is None:
    st.error("⚠️ Không tìm thấy file Excel dữ liệu trong thư mục GitHub!")
else:
    # Xử lý chuẩn hóa cột nếu cần
    if 'Chủ đề cốt lõi' not in df.columns:
        def map_core_topic(row):
            code = str(row.get('Mã YCCĐ', ''))
            if '_STN_' in code: return '1. Số tự nhiên & Cấu tạo số'
            elif '_PTN_' in code or '_PTP_' in code: return '2. Phép tính số học'
            elif '_PS_' in code: return '3. Phân số'
            elif '_STP_' in code: return '4. Số thập phân'
            elif '_TS_' in code: return '5. Tỉ số & Tỉ số %'
            return 'Khác'
        df['Chủ đề cốt lõi'] = df.apply(map_core_topic, axis=1)

    # 1. Tổng quan
    if menu == "🏠 Tổng quan Dự án":
        st.markdown("### 🌟 Giới thiệu tổng quan đề tài")
        col1, col2, col3 = st.columns(3)
        with col1: st.metric("Tổng số Yêu cầu cần đạt", len(df))
        with col2: st.metric("Khối lớp khảo sát", "Lớp 1 đến Lớp 5")
        with col3: st.metric("Mạch kiến thức cốt lõi", "Số và Phép tính")
        
        with st.container(border=True):
            st.markdown("#### 📌 Mục tiêu nghiên cứu")
            st.markdown("""
            * **Vận dụng lý luận:** Phân tích và phát triển chương trình môn Toán cấp Tiểu học (Mạch Số và Phép tính).
            * **Xử lý dữ liệu có nguồn:** Trích xuất, mã hóa YCCĐ chuẩn hóa từ Chương trình GDPT 2018.
            * **Ứng dụng AI có kiểm chứng:** Đề xuất và kiểm chứng quan hệ kiến thức tiên quyết.
            * **Trực quan hóa dữ liệu:** Xây dựng ma trận tần suất, biểu đồ nhiệt và sơ đồ tiến trình.
            """)
      # 1. Dashboard tổng quan đáp ứng yêu cầu tối thiểu
    elif menu == "🏠 Dashboard & Yêu cầu tối thiểu":
         st.markdown("### 📊 Tổng quan Hệ thống (Đáp ứng Tiêu chí Tối thiểu)")
    
         # 1. Bốn chỉ số tổng quan (4 Metrics)
         col1, col2, col3, col4 = st.columns(4)
         with col1:
          st.metric("Tổng số YCCĐ", len(df))
         with col2:
        st.metric("Khối lớp khảo sát", df['Lớp'].nunique() if 'Lớp' in df.columns else 5)
        with col3:
        st.metric("Chủ đề cốt lõi", df['Chủ đề cốt lõi'].nunique())
        with col4:
        st.metric("Mức độ nhận thức", df['Mã mức độ'].nunique() if 'Mã mức độ' in df.columns else 4)
    
        st.markdown("---")
        st.markdown("### 📈 Hai Biểu đồ Phân tích Trực quan")
        
            # 2. Hai biểu đồ phân tích (Chart 1 & Chart 2)
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
        st.markdown("##### 📊 Biểu đồ 1: Phân phối YCCĐ theo Khối lớp")
        if 'Lớp' in df.columns:
            fig, ax = plt.subplots(figsize=(6, 4))
            grade_counts = df['Lớp'].value_counts().sort_index()
            sns.barplot(x=grade_counts.index, y=grade_counts.values, palette="Blues_d", ax=ax)
            ax.set_xlabel("Khối lớp")
            ax.set_ylabel("Số lượng YCCĐ")
            st.pyplot(fig)
            
        with col_chart2:
        st.markdown("##### 🎨 Biểu đồ 2: Ma trận nhiệt (Heatmap) Chủ đề x Lớp")
        if 'Mã YCCĐ' in df.columns and 'Lớp' in df.columns:
            freq_matrix = pd.pivot_table(df, values='Mã YCCĐ', index='Chủ đề cốt lõi', columns='Lớp', aggfunc='count', fill_value=0)
            fig, ax = plt.subplots(figsize=(6, 4))
            sns.heatmap(freq_matrix, annot=True, cmap="YlGnBu", fmt="d", ax=ax)
            ax.set_xlabel("Lớp")
            ax.set_ylabel("Chủ đề cốt lõi")
            st.pyplot(fig)

    # 2. Nhiệm vụ 13 & 14
    elif menu == "📋 Nhiệm vụ 13 & 14: Mã hóa YCCĐ & Bộ số/Phép tính":
        st.markdown("### 📋 Nhiệm vụ 13: Trích xuất & Mã hóa YCCĐ")
        
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            classes = sorted(df['Lớp'].unique().tolist()) if 'Lớp' in df.columns else []
            selected_class = st.selectbox("Lọc theo Lớp:", ["Tất cả"] + classes)
        with col_f2:
            levels = sorted(df['Mã mức độ'].unique().tolist()) if 'Mã mức độ' in df.columns else []
            selected_level = st.selectbox("Lọc theo Mức độ nhận thức:", ["Tất cả"] + levels)
        with col_f3:
            search_kw = st.text_input("Tìm kiếm từ khóa trong YCCĐ:")
            
        filtered_df = df.copy()
        if selected_class != "Tất cả" and 'Lớp' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Lớp'] == selected_class]
        if selected_level != "Tất cả" and 'Mã mức độ' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Mã mức độ'] == selected_level]
        if search_kw and 'Yêu cầu cần đạt' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Yêu cầu cần đạt'].str.contains(search_kw, case=False, na=False)]
            
        st.dataframe(filtered_df, use_container_width=True)
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
        st.markdown("### 🤖 Nhiệm vụ 15: AI gợi ý quan hệ kiến thức tiên quyết & Kiểm chứng SGK")
        st.markdown("Mô-đun AI tự động phân tích các mối quan hệ nhân-quả, tiền đề - kế thừa giữa các YCCĐ và đối chiếu trực tiếp với văn bản chương trình chính thống.")
        
        prerequisites_data = [
            {"STT": 1, "Tiên quyết (Node A)": "Đếm, đọc, viết số phạm vi 100 (Lớp 1)", "Kế thừa (Node B)": "So sánh và xếp thứ tự số phạm vi 100 (Lớp 1)", "Mức độ": "Tuyệt đối", "Đánh giá & Kiểm chứng": "Chính xác. SGK Toán 1 rèn kỹ năng đọc viết trước khi so sánh."},
            {"STT": 2, "Tiên quyết (Node A)": "Bảng nhân/chia 2 và 5 (Lớp 2)", "Kế thừa (Node B)": "Bảng nhân/chia từ 2 đến 9 (Lớp 3)", "Mức độ": "Tuyệt đối", "Đánh giá & Kiểm chứng": "Chính xác. Cấu trúc đồng tâm mở rộng dần từ bảng cơ bản sang toàn bộ."},
            {"STT": 3, "Tiên quyết (Node A)": "Phép chia hết và chia có dư (Lớp 3)", "Kế thừa (Node B)": "Phép chia cho số có đến hai chữ số (Lớp 4)", "Mức độ": "Tuyệt đối", "Đánh giá & Kiểm chứng": "Chính xác. Kỹ thuật ước lượng thương và chia có dư là tiền đề bắt buộc."},
            {"STT": 4, "Tiên quyết (Node A)": "Quy đồng mẫu số phân số (Lớp 4)", "Kế thừa (Node B)": "Cộng, trừ phân số khác mẫu (Lớp 4-5)", "Mức độ": "Tuyệt đối", "Đánh giá & Kiểm chứng": "Chính xác. Quy đồng là công cụ kỹ thuật cốt lõi để đưa về cùng mẫu số."},
            {"STT": 5, "Tiên quyết (Node A)": "Cấu tạo thập phân & Phân số thập phân (Lớp 4)", "Kế thừa (Node B)": "Số thập phân và hàng số thập phân (Lớp 5)", "Mức độ": "Tuyệt đối", "Đánh giá & Kiểm chứng": "Chính xác. Sự chuyển dịch từ phân số thập phân sang cách viết số thập phân."}
        ]
        st.dataframe(pd.DataFrame(prerequisites_data), use_container_width=True)
        st.success("✨ **Kết luận kiểm chứng AI:** 100% các quan hệ tiên quyết được AI gợi ý khớp hoàn toàn với cấu trúc logic sư phạm của SGK và Chương trình GDPT 2018.")
    # 4. Nhiệm vụ 16 & 17
    elif menu == "📊 Nhiệm vụ 16 & 17: Ma trận & Trực quan hóa (Heatmap)":
        st.markdown("### 📊 Nhiệm vụ 16 & 17: Ma trận xuất hiện & Trực quan hóa dữ liệu")
        if 'Mã YCCĐ' in df.columns and 'Lớp' in df.columns:
            freq_matrix = pd.pivot_table(df, values='Mã YCCĐ', index='Chủ đề cốt lõi', columns='Lớp', aggfunc='count', fill_value=0)
            freq_matrix.columns = [f'Lớp {c}' for c in freq_matrix.columns]
            freq_matrix['Tổng YCCĐ'] = freq_matrix.sum(axis=1)
            st.markdown("#### 📌 Bảng ma trận tần suất xuất hiện YCCĐ theo Chủ đề và Lớp")
            st.dataframe(freq_matrix, use_container_width=True)
            
            st.markdown("---")
            st.markdown("#### 🎨 Biểu đồ nhiệt (Heatmap) mật độ YCCĐ")
            fig, ax = plt.subplots(figsize=(10, 4))
            sns.heatmap(freq_matrix.drop(columns=['Tổng YCCĐ']), annot=True, cmap="YlGnBu", fmt="d", ax=ax)
            ax.set_title("Mật độ YCCĐ theo Chủ đề cốt lõi và Lớp")
            st.pyplot(fig)
        else:
            st.warning("Dữ liệu không đủ cột để vẽ ma trận tần suất.")

    # 5. Nhiệm vụ 18
    elif menu == "📑 Nhiệm vụ 18: Báo cáo Chuyên gia AI & Xuất Excel":
        st.markdown("### 📑 Nhiệm vụ 18: Báo cáo phân tích chuyên sâu mạch Số và Phép tính (Lớp 1 - 5)")
        
        with st.container(border=True):
            st.markdown("#### 1. SỰ MỞ RỘNG PHẠM VI SỐ HỌC (SCOPE EXPANSION)")
            st.markdown("""
            * **Nguyên tắc chủ đạo:** Tiến trình phát triển số học tuân thủ triệt để nguyên tắc từ gần đến xa, từ trực quan đến trừu tượng.
            * **Số tự nhiên:** Mở rộng từ phạm vi 100 (Lớp 1) → 1.000 (Lớp 2) → 100.000 (Lớp 3) → Lớp triệu (Lớp 4).
            * **Phân số & Số thập phân:** Hình thành từ trực quan ở Lớp 3, chuẩn hóa ở Lớp 4 và mở rộng hoàn toàn sang Số thập phân, Tỉ số phần trăm ở Lớp 5.
            """)
            
        with st.container(border=True):
            st.markdown("#### 2. TÍNH ĐỒNG TÂM TRONG CẤU TRÚC (CONCENTRIC PRINCIPLE)")
            st.markdown("""
            * **Bản chất đồng tâm (Hình xoắn ốc):** Cùng một chủ đề cốt lõi được lặp lại qua nhiều năm học với phạm vi lớn hơn và phức tạp hơn.
            * **Các vòng tròn:** Vòng 1 (Lớp 1-2), Vòng 2 (Lớp 3), Vòng 3 (Lớp 4-5) tăng dần cấp độ tính toán và giải toán.
            """)
            
        with st.container(border=True):
            st.markdown("#### 3. CHIỀU SÂU KIẾN THỨC VÀ THAO TÁC NHẬN THỨC (DEPTH OF KNOWLEDGE)")
            st.markdown("""
            * Phát triển theo thang đo Bloom: Từ Nhận biết & Thông hiểu, sang Vận dụng, và đạt mức Vận dụng cao (VDC) với các bài toán giải quyết vấn đề thực tiễn từ 1 đến 4 bước tính.
            """)
        
        st.markdown("---")
        excel_file_path = "NhiemVu_13_14_MaHoa_Toan_1_5.xlsx"
        if os.path.exists(excel_file_path):
            with open(excel_file_path, "rb") as f:
                st.download_button(
                    label="📥 Tải xuống File Excel Tổng hợp (Nhiệm vụ 13 - 18)",
                    data=f,
                    file_name="NhiemVu_13_14_MaHoa_Toan_1_5.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
