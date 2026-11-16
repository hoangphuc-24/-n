import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==============================================================================
# CẤU HÌNH TRANG & GIAO DIỆN
# ==============================================================================
st.set_page_config(
    page_title="Curriculum Analytics | Mạch Số & Phép tính (Lớp 1-5)",
    page_icon="📊",
    layout="wide",
)

st.title("📊 CURRICULUM ANALYTICS: MẠCH SỐ VÀ PHÉP TÍNH (LỚP 1–5)")
st.caption(
    "Phân tích tiến trình phát triển chương trình GDPT 2018 môn Toán Tiểu học"
    " | Dữ liệu chuẩn hóa: 108 YCCĐ"
)


# ==============================================================================
# HÀM TẢI VÀ XỬ LÝ DỮ LIỆU TỪ 4 FILE EXCEL
# ==============================================================================
@st.cache_data
def load_all_data():
    # 1. File NV 13,14.xlsx
    df_yccd = pd.read_excel("NV 13,14.xlsx", sheet_name="NV13_Ma_Hoa_YCCD")
    df_nv14 = pd.read_excel(
        "NV 13,14.xlsx", sheet_name="NV14_Bo_Ma_So_Phep_Tinh", skiprows=2
    )
    df_quyuoc = pd.read_excel(
        "NV 13,14.xlsx", sheet_name="Quy_Uoc_Ma_Hoa", skiprows=2
    )

    # 2. File NV15.xlsx
    df_nv15_tq = pd.read_excel(
        "NV15.xlsx", sheet_name="Tong Quan & Tien Quyet", skiprows=4
    )
    df_nv15_pvs = pd.read_excel(
        "NV15.xlsx", sheet_name="Tuyen Pham Vi So", skiprows=2
    )
    df_nv15_pt = pd.read_excel(
        "NV15.xlsx", sheet_name="Tuyen Phep Tinh & Bai Toan", skiprows=2
    )
    df_nv15_tc = pd.read_excel(
        "NV15.xlsx", sheet_name="Thông Tin Tra Cứu", skiprows=2
    )

    # 3. File NV16.xlsx
    df_nv16_raw = pd.read_excel(
        "NV16.xlsx", sheet_name="2_DinhLuong_TanSuat_YCCD", skiprows=3
    )
    df_nv16_mach = df_nv16_raw.iloc[:6].copy()
    df_nv16_nhan_thuc = pd.read_excel(
        "NV16.xlsx", sheet_name="2_DinhLuong_TanSuat_YCCD", skiprows=12
    )
    df_nv16_truyxuat = pd.read_excel(
        "NV16.xlsx", sheet_name="3_TruyXuat_108_MaYCCD", skiprows=2
    )

    # 4. File NV17.xlsx
    df_nv17_raw = pd.read_excel(
        "NV17.xlsx", sheet_name="3_DuLieu_Heatmap_Excel", skiprows=3
    )
    df_heatmap_18 = df_nv17_raw.iloc[:18].copy()
    for col in ["Lớp 1", "Lớp 2", "Lớp 3", "Lớp 4", "Lớp 5", "Tổng YCCĐ"]:
        df_heatmap_18[col] = pd.to_numeric(
            df_heatmap_18[col], errors="coerce"
        ).fillna(0)

    df_heatmap_cog = pd.read_excel(
        "NV17.xlsx", sheet_name="3_DuLieu_Heatmap_Excel", skiprows=25
    ).iloc[:4]
    for col in ["Lớp 1", "Lớp 2", "Lớp 3", "Lớp 4", "Lớp 5", "Tổng YCCĐ"]:
        df_heatmap_cog[col] = pd.to_numeric(
            df_heatmap_cog[col], errors="coerce"
        ).fillna(0)

    return (
        df_yccd,
        df_nv14,
        df_quyuoc,
        df_nv15_tq,
        df_nv15_pvs,
        df_nv15_pt,
        df_nv15_tc,
        df_nv16_mach,
        df_nv16_nhan_thuc,
        df_nv16_truyxuat,
        df_heatmap_18,
        df_heatmap_cog,
    )


try:
    (
        df_yccd,
        df_nv14,
        df_quyuoc,
        df_nv15_tq,
        df_nv15_pvs,
        df_nv15_pt,
        df_nv15_tc,
        df_nv16_mach,
        df_nv16_nhan_thuc,
        df_nv16_truyxuat,
        df_heatmap_18,
        df_heatmap_cog,
    ) = load_all_data()
except Exception as e:
    st.error(
        f"Lỗi đọc dữ liệu Excel: {e}. Hãy kiểm tra lại tên các file Excel trên"
        " GitHub."
    )
    st.stop()

weight_map = {"NB": 1, "TH": 2, "VD": 3, "VDC": 4}
df_yccd["Điểm nhận thức"] = df_yccd["Mã mức độ"].map(weight_map)

# ==============================================================================
# THANH BÊN (SIDEBAR): BỘ LỌC TƯƠNG TÁC
# ==============================================================================
st.sidebar.header("🔍 BỘ LỌC TRUY XUẤT DỮ LIỆU")

all_grades = sorted(df_yccd["Lớp"].unique().tolist())
selected_grades = st.sidebar.multiselect(
    "1. Chọn Khối lớp:",
    options=all_grades,
    default=all_grades,
    format_func=lambda x: f"Lớp {x}",
)

all_topics = df_yccd["Nội dung"].unique().tolist()
selected_topics = st.sidebar.multiselect(
    "2. Chọn Mạch nội dung:", options=all_topics, default=all_topics
)

all_levels = ["NB", "TH", "VD", "VDC"]
level_labels = {
    "NB": "NB - Nhận biết (1)",
    "TH": "TH - Thông hiểu (2)",
    "VD": "VD - Vận dụng (3)",
    "VDC": "VDC - Vận dụng cao (4)",
}
selected_levels = st.sidebar.multiselect(
    "3. Chọn Thao tác nhận thức:",
    options=all_levels,
    default=all_levels,
    format_func=lambda x: level_labels.get(x, x),
)

search_keyword = st.sidebar.text_input(
    "4. Tìm kiếm từ khóa trong YCCĐ / Mã:",
    placeholder="Ví dụ: phân số, làm tròn, L4_...",
)

# Áp dụng bộ lọc
df_filtered = df_yccd[
    (df_yccd["Lớp"].isin(selected_grades))
    & (df_yccd["Nội dung"].isin(selected_topics))
    & (df_yccd["Mã mức độ"].isin(selected_levels))
]

if search_keyword.strip():
    kw = search_keyword.strip().lower()
    df_filtered = df_filtered[
        df_filtered["Yêu cầu cần đạt"].str.lower().str.contains(kw, na=False)
        | df_filtered["Mã YCCĐ"].str.lower().str.contains(kw, na=False)
        | df_filtered["Tiểu mục"].str.lower().str.contains(kw, na=False)
    ]

# ==============================================================================
# PHẦN 1: 4 CHỈ SỐ TỔNG QUAN (KPI CARDS)
# ==============================================================================
st.markdown("### Chỉ số tổng quan toàn mạch")
kpi1, kpi2, kpi3, kpi4 = st.columns(4)

total_records = len(df_filtered)
total_routes = len(df_heatmap_18)
high_cog_count = len(df_filtered[df_filtered["Mã mức độ"].isin(["VD", "VDC"])])
high_cog_pct = (
    (high_cog_count / total_records * 100) if total_records > 0 else 0
)
avg_depth = (
    df_filtered["Điểm nhận thức"].mean() if total_records > 0 else 0
)

with kpi1:
    st.metric(
        "📌 Tổng số YCCĐ",
        f"{total_records} / 108",
        f"Chiếm {total_records/108*100:.1f}% chương trình",
    )
with kpi2:
    st.metric(
        "🗂️ Tuyến nội dung chuẩn hóa",
        f"{total_routes} tuyến",
        "5 nhóm mạch kiến thức",
    )
with kpi3:
    st.metric(
        "🧠 Chỉ số độ sâu tư duy TB",
        f"{avg_depth:.2f} / 4.00",
        "Tăng từ 2.00 (L1) → 2.54 (L5)",
    )
with kpi4:
    st.metric(
        "🚀 Tỷ lệ Vận dụng & VDC",
        f"{high_cog_pct:.1f}%",
        f"{high_cog_count} YCCĐ mức cao",
    )

st.divider()

# ==============================================================================
# PHẦN 2: CÁC TAB MENU (ĐÃ XÓA GHI CHÚ NHIỆM VỤ & THÊM TAB KHO FILE EXCEL)
# ==============================================================================
tab_1, tab_2, tab_3, tab_4, tab_5, tab_files = st.tabs([
    "📋 Mã hóa YCCĐ & Bộ mã Số",
    "🔗 Quan hệ kiến thức tiên quyết",
    "📊 Ma trận xuất hiện & Mức phát triển",
    "🗺️ Sơ đồ tiến trình & Heatmap",
    "📝 Phân tích đồng tâm & Báo cáo",
    "📂 Kho dữ liệu Excel gốc",
])

# ------------------------------------------------------------------------------
# TAB 1: MÃ HÓA YCCĐ & BỘ MÃ SỐ
# ------------------------------------------------------------------------------
with tab_1:
    st.subheader(
        "Trích xuất & Mã hóa 108 Yêu cầu cần đạt (Bảng truy xuất minh chứng)"
    )

    col_dl, col_info = st.columns([1, 3])
    with col_dl:
        csv_data = df_filtered.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📥 Tải bảng dữ liệu đang lọc (CSV)",
            data=csv_data,
            file_name="Ma_Hoa_108_YCCD.csv",
            mime="text/csv",
        )
    with col_info:
        st.info(
            f"Đang hiển thị **{len(df_filtered)} / 108** mã YCCĐ theo bộ lọc"
            " bên trái."
        )

    st.dataframe(
        df_filtered[[
            "Mã YCCĐ",
            "Lớp",
            "Nội dung",
            "Tiểu mục",
            "Yêu cầu cần đạt",
            "Động từ chính",
            "Mã mức độ",
            "Thao tác nhận thức",
            "Điểm nhận thức",
        ]],
        use_container_width=True,
        hide_index=True,
    )

    st.divider()
    st.subheader(
        "Bộ mã hóa Số tự nhiên, Phân số, Số thập phân & Phép tính tương ứng"
    )
    st.dataframe(df_nv14, use_container_width=True, hide_index=True)

    with st.expander("📖 Xem Quy ước hệ thống mã định danh"):
        st.dataframe(df_quyuoc, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# TAB 2: QUAN HỆ KIẾN THỨC TIÊN QUYẾT
# ------------------------------------------------------------------------------
with tab_2:
    st.subheader("Phân tích quan hệ kiến thức tiên quyết qua các khối lớp")
    st.markdown(
        "**1. Bảng tổng hợp mối quan hệ tiên quyết trực tiếp từ Lớp 1 đến"
        " Lớp 5:**"
    )
    st.dataframe(df_nv15_tq, use_container_width=True, hide_index=True)

    col_15a, col_15b = st.columns(2)
    with col_15a:
        st.markdown("**2. Tuyến mở rộng Phạm vi số & Cấu tạo thập phân:**")
        st.dataframe(df_nv15_pvs, use_container_width=True, hide_index=True)
    with col_15b:
        st.markdown("**3. Tuyến Kỹ thuật phép tính, Biểu thức & Giải toán:**")
        st.dataframe(df_nv15_pt, use_container_width=True, hide_index=True)

    with st.expander("📚 Văn bản căn cứ & Tài liệu kiểm chứng"):
        st.dataframe(df_nv15_tc, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# TAB 3: MA TRẬN XUẤT HIỆN & MỨC PHÁT TRIỂN
# ------------------------------------------------------------------------------
with tab_3:
    st.subheader("Ma trận xuất hiện và Mức độ phát triển theo khối lớp")

    col_16_chart1, col_16_chart2 = st.columns(2)

    with col_16_chart1:
        st.markdown("**Biểu đồ 1: Phân bố YCCĐ theo Khối lớp & Mức độ nhận thức**")
        if not df_filtered.empty:
            dist_df = (
                df_filtered.groupby(["Lớp", "Mã mức độ"])
                .size()
                .reset_index(name="Số lượng")
            )
            dist_df["Lớp_Label"] = "Lớp " + dist_df["Lớp"].astype(str)
            color_map = {
                "NB": "#636EFA",
                "TH": "#00CC96",
                "VD": "#FFA15A",
                "VDC": "#EF553B",
            }
            fig_bar = px.bar(
                dist_df,
                x="Lớp_Label",
                y="Số lượng",
                color="Mã mức độ",
                color_discrete_map=color_map,
                category_orders={"Mã mức độ": ["NB", "TH", "VD", "VDC"]},
                text="Số lượng",
                labels={
                    "Lớp_Label": "Khối lớp",
                    "Số lượng": "Số lượng YCCĐ",
                    "Mã mức độ": "Mức độ",
                },
            )
            fig_bar.update_layout(
                barmode="stack",
                height=380,
                legend=dict(orientation="h", y=1.12),
            )
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.warning("Không có dữ liệu phù hợp với bộ lọc.")

    with col_16_chart2:
        st.markdown(
            "**Biểu đồ 2: Chỉ số độ sâu tư duy trung bình qua các lớp (Thang"
            " 1–4)**"
        )
        if not df_filtered.empty:
            depth_df = (
                df_filtered.groupby("Lớp")["Điểm nhận thức"]
                .mean()
                .reset_index()
            )
            depth_df["Lớp_Label"] = "Lớp " + depth_df["Lớp"].astype(str)
            fig_line = px.line(
                depth_df,
                x="Lớp_Label",
                y="Điểm nhận thức",
                markers=True,
                text=depth_df["Điểm nhận thức"].round(2),
                labels={
                    "Lớp_Label": "Khối lớp",
                    "Điểm nhận thức": "Chỉ số độ sâu TB",
                },
            )
            fig_line.update_traces(
                textposition="top center",
                line=dict(color="#AB63FA", width=3),
                marker=dict(size=10),
            )
            fig_line.update_yaxes(range=[1.5, 3.0])
            fig_line.update_layout(height=380)
            st.plotly_chart(fig_line, use_container_width=True)

    st.markdown(
        "**Bảng 1: Ma trận tần suất xuất hiện YCCĐ theo 5 nhóm mạch kiến"
        " thức:**"
    )
    st.dataframe(df_nv16_mach, use_container_width=True, hide_index=True)

    st.markdown(
        "**Bảng 2: Ma trận phát triển thao tác nhận thức & Chỉ số độ sâu tư"
        " duy:**"
    )
    st.dataframe(df_nv16_nhan_thuc, use_container_width=True, hide_index=True)

    with st.expander("🔍 Xem Bảng ánh xạ 108 Mã YCCĐ vào Ma trận phát triển"):
        st.dataframe(df_nv16_truyxuat, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------------------
# TAB 4: SƠ ĐỒ TIẾN TRÌNH & HEATMAP
# ------------------------------------------------------------------------------
with tab_4:
    st.subheader("Sơ đồ tiến trình và Bản đồ nhiệt mật độ kiến thức")

    st.markdown(
        "**Phần 1: Bản đồ tiến trình phát triển mạch Số và phép tính (Lớp 1 →"
        " Lớp 5)**"
    )
    progression_dot = """
    digraph G {
        rankdir=LR;
        node [shape=box, style="rounded,filled", fontname="Arial", fontsize=10];
        edge [color="#555555", penwidth=1.5];

        subgraph cluster_1 {
            label="LỚP 1 (10 YCCĐ)"; style=filled; color="#EBF5FB";
            L1_STN [label="Số tự nhiên phạm vi 100\\n(Chục, Đơn vị)", fillcolor="#AED6F1"];
            L1_PT [label="Cộng, trừ không nhớ\\nToán 1 bước tính", fillcolor="#D6EAF8"];
        }
        subgraph cluster_2 {
            label="LỚP 2 (20 YCCĐ)"; style=filled; color="#E8F8F5";
            L2_STN [label="Số tự nhiên đến 1.000\\n(Trăm, Chục, Đơn vị)", fillcolor="#A3E4D7"];
            L2_PT [label="Cộng, trừ có nhớ 1 lượt\\nBảng nhân/chia 2 và 5", fillcolor="#D1F2EB"];
        }
        subgraph cluster_3 {
            label="LỚP 3 (23 YCCĐ)"; style=filled; color="#FEF9E7";
            L3_STN [label="Số tự nhiên đến 100.000\\nSố La Mã, Làm tròn số", fillcolor="#F9E79F"];
            L3_PT [label="Bảng nhân/chia 2-9\\nNhân/chia 1 chữ số, Chia dư", fillcolor="#FCF3CF"];
            L3_PS [label="Làm quen Phân số\\n(1/2 đến 1/9 trực quan)", fillcolor="#FAD7A0"];
        }
        subgraph cluster_4 {
            label="LỚP 4 (29 YCCĐ)"; style=filled; color="#FDEDEC";
            L4_STN [label="Số tự nhiên đến Lớp triệu\\nDãy số, Trung bình cộng", fillcolor="#F5B7B1"];
            L4_PT [label="4 phép tính đa chữ số\\nBiểu thức chữ, Toán 2-3 bước", fillcolor="#FADBD8"];
            L4_PS [label="Phân số hoàn chỉnh\\nRút gọn, Quy đồng, 4 phép tính PS", fillcolor="#F5CBA7"];
        }
        subgraph cluster_5 {
            label="LỚP 5 (26 YCCĐ)"; style=filled; color="#F4ECF7";
            L5_STN_PS [label="Hệ thống hóa STN & PS\\nHỗn số, Phân số thập phân", fillcolor="#D7BDE2"];
            L5_STP [label="Số thập phân & 4 phép tính\\nChuyển đổi số đo đại lượng", fillcolor="#EBDEF0"];
            L5_TS [label="Tỉ số & Tỉ số phần trăm\\nTỉ lệ bản đồ, Toán 4 bước", fillcolor="#E8DAEF"];
        }

        L1_STN -> L2_STN -> L3_STN -> L4_STN -> L5_STN_PS;
        L1_PT -> L2_PT -> L3_PT -> L4_PT -> L5_STN_PS;
        L3_PT -> L3_PS -> L4_PS -> L5_STN_PS;
        L4_PS -> L5_STP;
        L4_PS -> L5_TS;
        L4_PT -> L5_STP;
    }
    """
    st.graphviz_chart(progression_dot, use_container_width=True)

    st.divider()
    st.markdown("**Phần 2: Heatmap mật độ YCCĐ qua các khối lớp**")
    col_h1, col_h2 = st.columns([1.4, 1])

    with col_h1:
        st.caption("Heatmap 1: Mật độ 18 Tuyến nội dung chuẩn hóa (Lớp 1–5)")
        z_data_18 = df_heatmap_18[
            ["Lớp 1", "Lớp 2", "Lớp 3", "Lớp 4", "Lớp 5"]
        ].values
        y_labels_18 = df_heatmap_18["Tuyến nội dung chuẩn hóa"].tolist()
        x_labels = ["Lớp 1", "Lớp 2", "Lớp 3", "Lớp 4", "Lớp 5"]

        fig_heat18 = go.Figure(
            data=go.Heatmap(
                z=z_data_18,
                x=x_labels,
                y=y_labels_18,
                colorscale="YlGnBu",
                text=z_data_18,
                texttemplate="%{text}",
                textfont={"size": 11},
                colorbar=dict(title="YCCĐ"),
            )
        )
        fig_heat18.update_layout(
            height=680,
            yaxis=dict(autorange="reversed"),
            margin=dict(l=10, r=10, t=20, b=10),
        )
        st.plotly_chart(fig_heat18, use_container_width=True)

    with col_h2:
        st.caption(
            "Heatmap 2: Mật độ 4 Thao tác nhận thức (NB, TH, VD, VDC)"
        )
        z_data_cog = df_heatmap_cog[
            ["Lớp 1", "Lớp 2", "Lớp 3", "Lớp 4", "Lớp 5"]
        ].values
        y_labels_cog = df_heatmap_cog["Mã mức độ"].tolist()

        fig_heat_cog = go.Figure(
            data=go.Heatmap(
                z=z_data_cog,
                x=x_labels,
                y=y_labels_cog,
                colorscale="OrRd",
                text=z_data_cog,
                texttemplate="%{text}",
                textfont={"size": 14},
                colorbar=dict(title="YCCĐ"),
            )
        )
        fig_heat_cog.update_layout(
            height=380,
            yaxis=dict(autorange="reversed"),
            margin=dict(l=10, r=10, t=20, b=10),
        )
        st.plotly_chart(fig_heat_cog, use_container_width=True)
        st.dataframe(
            df_heatmap_cog[[
                "Mã mức độ",
                "Lớp 1",
                "Lớp 2",
                "Lớp 3",
                "Lớp 4",
                "Lớp 5",
                "Tổng YCCĐ",
            ]],
            use_container_width=True,
            hide_index=True,
        )

# ------------------------------------------------------------------------------
# TAB 5: PHÂN TÍCH ĐỒNG TÂM & BÁO CÁO BẢO VỆ
# ------------------------------------------------------------------------------
with tab_5:
    st.subheader(
        "Phân tích sự mở rộng phạm vi, tính đồng tâm và chiều sâu kiến thức"
    )
    st.success(
        "**Câu hỏi bảo vệ:** *Mạch Số và phép tính thể hiện nguyên tắc đồng"
        " tâm ra sao?*"
    )

    st.markdown("""
    Trong Chương trình GDPT 2018 môn Toán Tiểu học (Lớp 1–5), nguyên tắc **đồng tâm phát triển (xoắn ốc)** được thể hiện rõ nét qua **3 trục cốt lõi** được chứng minh bằng dữ liệu định lượng từ 108 YCCĐ:

    1. **Đồng tâm về mở rộng vòng số (Phạm vi số học):**
       - **Lớp 1:** Số tự nhiên trong phạm vi **10, 20, 100** (cấu tạo Chục – Đơn vị).
       - **Lớp 2:** Mở rộng vòng số lên **1.000** (bổ sung hàng Trăm).
       - **Lớp 3:** Mở rộng lên **10.000 và 100.000** (hàng Nghìn, Chục nghìn) và **làm quen phân số** ($1/2$ đến $1/9$).
       - **Lớp 4:** Mở rộng đến **Lớp triệu** và hoàn thiện **tập hợp Phân số** (tử số, mẫu số, quy đồng, rút gọn).
       - **Lớp 5:** Hệ thống hóa Số tự nhiên – Phân số và mở rộng sang **Số thập phân, Tỉ số phần trăm**.
       *(Mỗi vòng số mới đều lặp lại chu trình nhận thức: Đếm/Đọc/Viết $\rightarrow$ Cấu tạo thập phân $\rightarrow$ So sánh/Xếp thứ tự $\rightarrow$ Làm tròn số).*

    2. **Đồng tâm về cấu trúc phép tính & tính chất số học:**
       - Bốn phép tính (`Cộng, Trừ, Nhân, Chia`) cùng các tính chất (*giao hoán, kết hợp, phân phối*) không học dứt điểm một lần mà lặp lại trên từng tập số mới:
         - **Giai đoạn 1 (Lớp 1–3):** Thực hiện trên **Số tự nhiên** (từ không nhớ $\rightarrow$ có nhớ 1 lượt $\rightarrow$ nhớ nhiều lượt $\rightarrow$ chia có dư).
         - **Giai đoạn 2 (Lớp 4):** Tái cấu trúc 4 phép tính trên **Phân số**.
         - **Giai đoạn 3 (Lớp 5):** Tái cấu trúc 4 phép tính trên **Số thập phân**.

    3. **Đồng tâm nâng cao về độ sâu tư duy & giải quyết vấn đề:**
       - **Chỉ số độ sâu tư duy trung bình** tăng tịnh tiến qua các lớp: **2.00 (Lớp 1) $\rightarrow$ 1.95 (Lớp 2) $\rightarrow$ 2.13 (Lớp 3) $\rightarrow$ 2.31 (Lớp 4) $\rightarrow$ 2.54 (Lớp 5)**.
       - Tỷ trọng nhóm mức độ **Vận dụng & Vận dụng cao (VD + VDC)** tăng từ **40.0%** ở Lớp 1 lên **53.8%** ở Lớp 5.
       - Độ phức tạp của **bài toán có lời văn** phát triển đồng tâm theo số bước giải: **1 bước tính** (Lớp 1–2) $\rightarrow$ **2 bước tính** (Lớp 3) $\rightarrow$ **2–3 bước tính** (Lớp 4) $\rightarrow$ **3–4 bước tính** (Lớp 5).
    """)

    st.divider()
    st.subheader("🛠️ Báo cáo giải thích phương pháp & Kiểm chứng AI")
    st.markdown("""
    - **Bước 1 – Trích xuất & Mã hóa dữ liệu:** Rà soát toàn văn Thông tư 32/2018/TT-BGDĐT, bóc tách thành **108 YCCĐ** độc lập và gắn mã định danh theo cấu trúc `[Lớp]_[Mạch]_[Mức độ]_[STT]`.
    - **Bước 2 – Phân tích quan hệ tiên quyết với AI có kiểm chứng:** Sử dụng AI gợi ý cây phụ thuộc kiến thức giữa các lớp, sau đó đối chiếu thủ công với SGK *Kết nối tri thức với cuộc sống* để loại bỏ các suy diễn vượt chuẩn chương trình.
    - **Bước 3 – Định lượng & Trực quan hóa:** Gán trọng số nhận thức ($NB=1, TH=2, VD=3, VDC=4$), xây dựng ma trận tần suất 18 tuyến nội dung và biểu diễn trên dashboard Streamlit.
    """)

# ------------------------------------------------------------------------------
# TAB 6: KHO DỮ LIỆU EXCEL GỐC (TẢI XUỐNG & XEM TRƯỚC)
# ------------------------------------------------------------------------------
with tab_files:
    st.subheader("📂 Kho lưu trữ & Tải xuống các file dữ liệu Excel gốc")
    st.markdown(
        "Bạn có thể tải trực tiếp từng file Excel (`.xlsx`) về máy hoặc xem"
        " nhanh nội dung các sheet bên dưới:"
    )

    # Thay đổi URL bên dưới thành link GitHub thực tế của bạn nếu muốn dẫn trực tiếp tới repo
    GITHUB_REPO_URL = "https://github.com/hoangphuc-24"

    excel_catalog = [
        {
            "filename": "NV 13,14.xlsx",
            "title": "Dữ liệu Mã hóa YCCĐ & Bộ mã Số – Phép tính",
            "desc": (
                "Chứa 108 mã YCCĐ chi tiết, bộ mã phân loại nhóm số và quy ước"
                " mã định danh."
            ),
        },
        {
            "filename": "NV15.xlsx",
            "title": "Dữ liệu Phân tích Quan hệ Kiến thức Tiên quyết",
            "desc": (
                "Chứa ma trận quan hệ tiên quyết, tuyến phạm vi số và giới hạn"
                " kỹ thuật phép tính."
            ),
        },
        {
            "filename": "NV16.xlsx",
            "title": "Dữ liệu Định lượng Tần suất & Mức độ Phát triển",
            "desc": (
                "Chứa ma trận tần suất xuất hiện theo mạch kiến thức và chỉ số"
                " độ sâu tư duy."
            ),
        },
        {
            "filename": "NV17.xlsx",
            "title": "Dữ liệu Bản đồ Tiến trình & Heatmap",
            "desc": (
                "Chứa dữ liệu heatmap 18 tuyến nội dung chuẩn hóa và 4 thao tác"
                " nhận thức."
            ),
        },
    ]

    cols_file = st.columns(2)
    for idx, item in enumerate(excel_catalog):
        fname = item["filename"]
        with cols_file[idx % 2]:
            with st.container(border=True):
                st.markdown(f"#### 📄 {item['title']}")
                st.caption(f"Tên file gốc: `{fname}` — {item['desc']}")

                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if os.path.exists(fname):
                        with open(fname, "rb") as f:
                            st.download_button(
                                label=f"📥 Tải file {fname}",
                                data=f,
                                file_name=fname,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{fname}",
                                use_container_width=True,
                            )
                    else:
                        st.warning(f"Không tìm thấy {fname}")

                with btn_col2:
                    # Link mở trực tiếp trên GitHub (bạn có thể điền tên repo cụ thể vào URL nếu cần)
                    st.link_button(
                        label="🔗 Mở trên GitHub",
                        url=f"{GITHUB_REPO_URL}",
                        use_container_width=True,
                    )

    st.divider()
    st.markdown("### 👁️ Xem nhanh nội dung chi tiết từng file Excel")
    selected_preview_file = st.selectbox(
        "Chọn file Excel muốn kiểm tra cấu trúc các sheet:",
        options=[x["filename"] for x in excel_catalog],
    )

    if os.path.exists(selected_preview_file):
        xls_preview = pd.ExcelFile(selected_preview_file)
        sheet_tabs = st.tabs(xls_preview.sheet_names)
        for i, s_name in enumerate(xls_preview.sheet_names):
            with sheet_tabs[i]:
                df_raw = pd.read_excel(selected_preview_file, sheet_name=s_name)
                st.dataframe(df_raw, use_container_width=True)
