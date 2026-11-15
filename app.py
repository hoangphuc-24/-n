from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import streamlit as st


# ------------------------------ Cấu hình trang ------------------------------
st.set_page_config(
    page_title="Curriculum Analytics - Mạch Số và Phép Tính (Toán 1-5)",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    .main-header { font-size: 28px; font-weight: bold; color: #1F4E78; margin-bottom: 10px; }
    .sub-header { font-size: 18px; color: #555555; margin-bottom: 20px; }
    </style>
    """,
    unsafe_allow_html=True,
)

APP_DIR = Path(__file__).resolve().parent
EXCEL_FILES = (
    "NhiemVu_13_14_MaHoa_Toan_1_5.xlsx",
    "Ma_Hoa_YCCD_Toan_1_5_NhiemVu13.xlsx",
    "Phan_Tich_Toan_1_5_AI_NhiemVu_13_18.xlsx",
)

MENU_OVERVIEW = "🏠 Tổng quan Dự án"
MENU_DASHBOARD = "🏠 Dashboard & Yêu cầu tối thiểu"
MENU_TASK_13_14 = "📋 Nhiệm vụ 13 & 14: Mã hóa YCCĐ & Bộ số/Phép tính"
MENU_TASK_15 = "🤖 Nhiệm vụ 15: Phân tích AI & Quan hệ Tiên quyết"
MENU_TASK_16_17 = "📊 Nhiệm vụ 16 & 17: Ma trận & Trực quan hóa (Heatmap)"
MENU_TASK_18 = "📑 Nhiệm vụ 18: Báo cáo Chuyên gia AI & Xuất Excel"


@st.cache_data(show_spinner=False)
def load_data():
    """Đọc trang tính đầu tiên của file Excel khả dụng trong thư mục ứng dụng."""
    for filename in EXCEL_FILES:
        filepath = APP_DIR / filename
        if filepath.is_file():
            try:
                return pd.read_excel(filepath, sheet_name=0)
            except (OSError, ValueError, ImportError) as exc:
                st.error(f"Không thể đọc file dữ liệu `{filename}`: {exc}")
                return None
    return None


def infer_core_topic(row):
    code = str(row.get("Mã YCCĐ", "")).upper()
    if "_STN_" in code:
        return "1. Số tự nhiên & Cấu tạo số"
    if "_PTN_" in code or "_PTP_" in code:
        return "2. Phép tính số học"
    if "_PS_" in code:
        return "3. Phân số"
    if "_STP_" in code:
        return "4. Số thập phân"
    if "_TS_" in code:
        return "5. Tỉ số & Tỉ số %"
    return "Khác"


def make_frequency_matrix(data):
    """Tạo ma trận chủ đề x lớp; trả về None nếu thiếu cột cần thiết."""
    if not {"Lớp", "Chủ đề cốt lõi"}.issubset(data.columns):
        return None
    value_column = "Mã YCCĐ" if "Mã YCCĐ" in data.columns else "Chủ đề cốt lõi"
    return pd.pivot_table(
        data,
        values=value_column,
        index="Chủ đề cốt lõi",
        columns="Lớp",
        aggfunc="count",
        fill_value=0,
    )


# ------------------------------ Tiêu đề/sidebar ------------------------------
st.markdown(
    '<div class="main-header">🎯 ĐỀ 13: TIẾN TRÌNH MẠCH SỐ VÀ PHÉP TÍNH MÔN TOÁN LỚP 1–5</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-header">Hệ thống Phân tích Dữ liệu Chương trình GDPT 2018 | Hỗ trợ AI & Trực quan hóa</div>',
    unsafe_allow_html=True,
)

st.sidebar.image("https://img.icons8.com/color/96/000000/math.png", width=80)
st.sidebar.title("Điều hướng Đề tài")
menu = st.sidebar.radio(
    "Chọn nhiệm vụ / nội dung:",
    [MENU_OVERVIEW, MENU_DASHBOARD, MENU_TASK_13_14, MENU_TASK_15, MENU_TASK_16_17, MENU_TASK_18],
)

# ------------------------------ Nạp/chuẩn hóa dữ liệu ------------------------------
df = load_data()
if df is None:
    st.error(
        "⚠️ Không tìm thấy hoặc không đọc được file Excel dữ liệu. "
        "Hãy đặt một trong các file dữ liệu được hỗ trợ cùng thư mục với `app.py`."
    )
    st.caption("Tên file được hỗ trợ: " + ", ".join(EXCEL_FILES))
    st.stop()

if "Chủ đề cốt lõi" not in df.columns:
    df["Chủ đề cốt lõi"] = df.apply(infer_core_topic, axis=1)

# Tạo bản sao chỉ số để mọi nhánh menu cùng dùng dữ liệu đã chuẩn hóa.

# ------------------------------ Tổng quan ------------------------------
if menu == MENU_OVERVIEW:
    st.markdown("### 🌟 Giới thiệu tổng quan đề tài")
    col1, col2, col3 = st.columns(3)
    col1.metric("Tổng số Yêu cầu cần đạt", len(df))
    col2.metric("Khối lớp khảo sát", "Lớp 1 đến Lớp 5")
    col3.metric("Mạch kiến thức cốt lõi", "Số và Phép tính")

    with st.container(border=True):
        st.markdown("#### 📌 Mục tiêu nghiên cứu")
        st.markdown(
            """
            * **Vận dụng lý luận:** Phân tích và phát triển chương trình môn Toán cấp Tiểu học (mạch Số và Phép tính).
            * **Xử lý dữ liệu có nguồn:** Trích xuất, mã hóa YCCĐ chuẩn hóa từ Chương trình GDPT 2018.
            * **Ứng dụng AI có kiểm chứng:** Đề xuất và kiểm chứng quan hệ kiến thức tiên quyết.
            * **Trực quan hóa dữ liệu:** Xây dựng ma trận tần suất, biểu đồ nhiệt và sơ đồ tiến trình.
            """
        )

# ------------------------------ Dashboard ------------------------------
elif menu == MENU_DASHBOARD:
    st.markdown("### 📊 Tổng quan hệ thống (đáp ứng tiêu chí tối thiểu)")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Tổng số YCCĐ", len(df))
    col2.metric("Khối lớp khảo sát", df["Lớp"].nunique() if "Lớp" in df.columns else 5)
    col3.metric("Chủ đề cốt lõi", df["Chủ đề cốt lõi"].nunique())
    col4.metric("Mức độ nhận thức", df["Mã mức độ"].nunique() if "Mã mức độ" in df.columns else "—")

    st.markdown("---")
    st.markdown("### 📈 Hai biểu đồ phân tích trực quan")
    chart1, chart2 = st.columns(2)

    with chart1:
        st.markdown("##### 📊 Phân phối YCCĐ theo khối lớp")
        if "Lớp" in df.columns and not df.empty:
            counts = df["Lớp"].value_counts().sort_index()
            fig, ax = plt.subplots(figsize=(6, 4))
            sns.barplot(x=counts.index.astype(str), y=counts.values, color="#4C78A8", ax=ax)
            ax.set_xlabel("Khối lớp")
            ax.set_ylabel("Số lượng YCCĐ")
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
        else:
            st.info("Không có cột 'Lớp' để vẽ biểu đồ.")

    with chart2:
        st.markdown("##### 🎨 Heatmap chủ đề cốt lõi × lớp")
        matrix = make_frequency_matrix(df)
        if matrix is not None and not matrix.empty:
            fig, ax = plt.subplots(figsize=(6, 4))
            sns.heatmap(matrix, annot=True, cmap="YlGnBu", fmt="d", ax=ax)
            ax.set_xlabel("Lớp")
            ax.set_ylabel("Chủ đề cốt lõi")
            fig.tight_layout()
            st.pyplot(fig)
            plt.close(fig)
        else:
            st.info("Không đủ dữ liệu (cần cột 'Lớp') để vẽ heatmap.")

# ------------------------------ Nhiệm vụ 13 & 14 ------------------------------
elif menu == MENU_TASK_13_14:
    st.markdown("### 📋 Nhiệm vụ 13: Trích xuất & mã hóa YCCĐ")
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        classes = sorted(df["Lớp"].dropna().unique().tolist(), key=str) if "Lớp" in df.columns else []
        selected_class = st.selectbox("Lọc theo lớp:", ["Tất cả"] + classes)
    with col_f2:
        levels = sorted(df["Mã mức độ"].dropna().unique().tolist(), key=str) if "Mã mức độ" in df.columns else []
        selected_level = st.selectbox("Lọc theo mức độ nhận thức:", ["Tất cả"] + levels)
    with col_f3:
        search_kw = st.text_input("Tìm kiếm từ khóa trong YCCĐ:")

    filtered_df = df.copy()
    if selected_class != "Tất cả" and "Lớp" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["Lớp"] == selected_class]
    if selected_level != "Tất cả" and "Mã mức độ" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["Mã mức độ"] == selected_level]
    if search_kw and "Yêu cầu cần đạt" in filtered_df.columns:
        filtered_df = filtered_df[
            filtered_df["Yêu cầu cần đạt"].astype("string").str.contains(
                search_kw, case=False, na=False, regex=False
            )
        ]

    st.dataframe(filtered_df, use_container_width=True)
    st.caption(f"Hiển thị {len(filtered_df)} / {len(df)} yêu cầu cần đạt.")

    st.markdown("---")
    st.markdown("### 🔢 Nhiệm vụ 14: Bộ số và phép tính tương ứng")
    data_14 = [
        ("M_STN_L1", "Lớp 1", "Số tự nhiên", "Cộng, trừ", "Phạm vi 10, 20, 100; số tròn chục", "Đếm, đọc, viết, so sánh số đến 100; cộng, trừ trong phạm vi phù hợp"),
        ("M_PT_L1", "Lớp 1", "Phép tính với số tự nhiên", "Cộng, trừ", "Tính nhẩm trong phạm vi 10; phép tính trong phạm vi 100", "Hiểu ý nghĩa phép cộng, phép trừ và thực hiện tính toán cơ bản"),
        ("M_STN_L2", "Lớp 2", "Số tự nhiên", "Cộng, trừ, nhân, chia", "Phạm vi 1.000; số tròn trăm; cấu tạo thập phân", "Đọc, viết, so sánh, xếp thứ tự số đến 1.000"),
        ("M_PT_L2", "Lớp 2", "Phép tính với số tự nhiên", "Cộng, trừ, nhân, chia", "Cộng trừ đến 1.000; bảng nhân/chia 2 và 5", "Tính toán và giải quyết tình huống thực tế đơn giản"),
        ("M_STN_L3", "Lớp 3", "Số tự nhiên", "Cộng, trừ, nhân, chia", "Phạm vi 10.000 và 100.000; làm quen số La Mã", "Đọc, viết, so sánh, làm tròn số tự nhiên"),
        ("M_PT_L3", "Lớp 3", "Phép tính với số tự nhiên", "Cộng, trừ, nhân, chia", "Tính với số nhiều chữ số; nhân/chia số có một chữ số; biểu thức", "Thực hiện phép tính, chia có dư và giải toán"),
        ("M_PS_L3", "Lớp 3", "Phân số", "Làm quen phân số", "Các phân số đơn vị qua hình ảnh trực quan", "Nhận biết phân số và phân số của một nhóm đồ vật"),
        ("M_STN_L4", "Lớp 4", "Số tự nhiên", "Cộng, trừ, nhân, chia", "Đến lớp triệu; số chẵn, số lẻ; dãy số tự nhiên", "Đọc, viết, so sánh số có nhiều chữ số"),
        ("M_PT_L4", "Lớp 4", "Phép tính với số tự nhiên", "Cộng, trừ, nhân, chia", "Phép tính nhiều chữ số; nhân/chia với số có nhiều chữ số", "Tính toán, tìm thành phần chưa biết và giải toán"),
        ("M_PS_L4", "Lớp 4", "Phân số", "Cộng, trừ, nhân, chia phân số", "Tử số, mẫu số; phân số bằng nhau và tối giản", "Rút gọn, quy đồng và thực hiện các phép tính phân số cơ bản"),
        ("M_STN_L5", "Lớp 5", "Số tự nhiên", "Ôn tập bốn phép tính", "Củng cố số tự nhiên và các phép tính", "Tính hợp lý, ước lượng và giải toán"),
        ("M_PS_L5", "Lớp 5", "Phân số", "Cộng, trừ, nhân, chia", "Phân số thập phân; hỗn số", "Ôn tập các phép tính và vận dụng giải toán"),
        ("M_STP_L5", "Lớp 5", "Số thập phân", "Cộng, trừ, nhân, chia", "Phần nguyên, phần thập phân và các hàng", "Đọc, viết, so sánh và làm tròn số thập phân"),
        ("M_PTP_L5", "Lớp 5", "Phép tính với số thập phân", "Cộng, trừ, nhân, chia", "Tính với số thập phân; nhân/chia nhẩm với 10, 100, 0,1...", "Thực hiện phép tính và giải toán với số thập phân"),
        ("M_TS_L5", "Lớp 5", "Tỉ số và tỉ số phần trăm", "Tính tỉ số, tỉ số phần trăm", "Tỉ số hai đại lượng; tỉ lệ bản đồ", "Giải toán về tỉ số phần trăm và vận dụng trong thực tế"),
    ]
    columns_14 = ["Mã nhóm", "Khối lớp", "Loại số / Đối tượng", "Phép tính tương ứng", "Phạm vi / Đặc điểm số học", "Yêu cầu cần đạt"]
    st.dataframe(pd.DataFrame(data_14, columns=columns_14), use_container_width=True)

# ------------------------------ Nhiệm vụ 15 ------------------------------
elif menu == MENU_TASK_15:
    st.markdown("### 🤖 Nhiệm vụ 15: Gợi ý quan hệ kiến thức tiên quyết")
    st.markdown(
        "Bảng dưới đây trình bày các quan hệ kiến thức tham khảo. Cần đối chiếu với chương trình và sách giáo khoa trước khi sử dụng làm kết luận chính thức."
    )
    prerequisites_data = [
        {"STT": 1, "Tiên quyết (Node A)": "Đếm, đọc, viết số trong phạm vi 100 (Lớp 1)", "Kế thừa (Node B)": "So sánh và xếp thứ tự số trong phạm vi 100 (Lớp 1)", "Mức độ": "Cao", "Căn cứ": "Cần hiểu số và giá trị số trước khi so sánh."},
        {"STT": 2, "Tiên quyết (Node A)": "Bảng nhân/chia 2 và 5 (Lớp 2)", "Kế thừa (Node B)": "Mở rộng bảng nhân/chia (Lớp 3)", "Mức độ": "Cao", "Căn cứ": "Kiến thức bảng nhân/chia được mở rộng và vận dụng."},
        {"STT": 3, "Tiên quyết (Node A)": "Phép chia và chia có dư (Lớp 3)", "Kế thừa (Node B)": "Thực hiện phép chia với số lớn hơn (Lớp 4)", "Mức độ": "Cao", "Căn cứ": "Kế thừa ý nghĩa và kỹ thuật phép chia đã học."},
        {"STT": 4, "Tiên quyết (Node A)": "Nhận biết phân số và phân số bằng nhau", "Kế thừa (Node B)": "Quy đồng, cộng và trừ phân số", "Mức độ": "Cao", "Căn cứ": "Cần hiểu cấu tạo và tính chất phân số."},
        {"STT": 5, "Tiên quyết (Node A)": "Hàng của số và phân số thập phân", "Kế thừa (Node B)": "Đọc, viết và tính toán với số thập phân (Lớp 5)", "Mức độ": "Cao", "Căn cứ": "Số thập phân dựa trên cấu tạo thập phân."},
    ]
    st.dataframe(pd.DataFrame(prerequisites_data), use_container_width=True)
    st.info("Các quan hệ trên là gợi ý phân tích, không nên xem là kết quả kiểm chứng tự động nếu chưa có nguồn đối chiếu cụ thể.")

# ------------------------------ Nhiệm vụ 16 & 17 ------------------------------
elif menu == MENU_TASK_16_17:
    st.markdown("### 📊 Nhiệm vụ 16 & 17: Ma trận tần suất và trực quan hóa")
    matrix = make_frequency_matrix(df)
    if matrix is None:
        st.warning("Dữ liệu cần có cột 'Lớp' để tạo ma trận tần suất.")
    elif matrix.empty:
        st.info("Không có dữ liệu để lập ma trận.")
    else:
        matrix.columns = [f"Lớp {column}" for column in matrix.columns]
        matrix["Tổng YCCĐ"] = matrix.sum(axis=1)
        st.markdown("#### 📌 Ma trận tần suất YCCĐ theo chủ đề và lớp")
        st.dataframe(matrix, use_container_width=True)
        st.markdown("---")
        st.markdown("#### 🎨 Heatmap mật độ YCCĐ")
        fig, ax = plt.subplots(figsize=(10, 4))
        sns.heatmap(matrix.drop(columns=["Tổng YCCĐ"]), annot=True, cmap="YlGnBu", fmt="g", ax=ax)
        ax.set_title("Mật độ YCCĐ theo chủ đề cốt lõi và lớp")
        fig.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

# ------------------------------ Nhiệm vụ 18 ------------------------------
elif menu == MENU_TASK_18:
    st.markdown("### 📑 Nhiệm vụ 18: Báo cáo phân tích mạch Số và Phép tính (Lớp 1–5)")
    with st.container(border=True):
        st.markdown("#### 1. Sự mở rộng phạm vi số học")
        st.markdown(
            """
            * Tiến trình học tập phát triển từ trực quan đến trừu tượng và mở rộng dần phạm vi số.
            * Số tự nhiên được mở rộng qua các lớp; học sinh củng cố giá trị hàng và cấu tạo thập phân.
            * Phân số được làm quen trước, sau đó phát triển các phép tính; số thập phân và tỉ số phần trăm được học ở lớp trên.
            """
        )
    with st.container(border=True):
        st.markdown("#### 2. Tính đồng tâm trong cấu trúc")
        st.markdown(
            "Các chủ đề cốt lõi được củng cố và mở rộng qua nhiều lớp, với độ phức tạp tăng dần về phạm vi tính toán và tình huống vận dụng."
        )
    with st.container(border=True):
        st.markdown("#### 3. Chiều sâu kiến thức và thao tác nhận thức")
        st.markdown(
            "Quá trình học tập tiến từ nhận biết, thông hiểu đến vận dụng kiến thức để giải quyết bài toán và tình huống thực tế."
        )

    st.markdown("---")
    excel_path = next((APP_DIR / name for name in EXCEL_FILES if (APP_DIR / name).is_file()), None)
    if excel_path is not None:
        st.download_button(
            label="📥 Tải xuống file Excel dữ liệu",
            data=excel_path.read_bytes(),
            file_name=excel_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Không tìm thấy file Excel để tải xuống.")
