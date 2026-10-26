import os
import json
import random
import pandas as pd
import streamlit as st

# ==========================================
# 0. CẤU HÌNH TRANG & THƯ MỤC LƯU TRỮ
# ==========================================
st.set_page_config(
    page_title="EduClass-AI: Trợ lý Sư phạm số Lớp 6/6",
    page_icon="🎓",
    layout="wide"
)

DATA_DIR = "du_lieu_lop_hoc_moi"
LOG_FILE = os.path.join(DATA_DIR, "luot_tra_cuu_moi.json")

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

# Mật khẩu bảo mật phân hệ
GVBM_PASSWORD = "221219"
GVCN_PASSWORD = "221220"


# ==========================================
# 1. HÀM XỬ LÝ LÀM SẠCH VÀ HỖ TRỢ DỮ LIỆU
# ==========================================
def clean_id(val) -> str:
    """
    Làm sạch chuỗi mã định danh:
    Ép kiểu chuỗi, xóa khoảng trắng, cắt bỏ đuôi `.0` nếu pandas đọc số nguyên dạng float,
    lọc bỏ các giá trị rác hoặc None/Nan.
    """
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", ""]:
        return ""
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str


def load_search_logs() -> dict:
    """Tải lịch sử tra cứu của học sinh từ file JSON."""
    if not os.path.exists(LOG_FILE):
        return {}
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_search_logs(logs: dict):
    """Ghi lịch sử tra cứu vào file JSON."""
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(logs, f, ensure_ascii=False, indent=4)


def record_search(student_id: str, subject_name: str):
    """Cộng thêm 1 lượt tra cứu cho môn học của học sinh cụ thể."""
    logs = load_search_logs()
    if student_id not in logs:
        logs[student_id] = {}
    logs[student_id][subject_name] = logs[student_id].get(subject_name, 0) + 1
    save_search_logs(logs)


def get_subject_files() -> list:
    """Lấy danh sách các môn học đã tạo file trong thư mục dữ liệu."""
    if not os.path.exists(DATA_DIR):
        return []
    files = [
        f for f in os.listdir(DATA_DIR)
        if f.endswith("_DiemTongHop.xlsx") or f.endswith("_DiemTongHop.csv")
    ]
    return files


def load_subject_df(file_path: str) -> pd.DataFrame:
    """Tải và làm sạch dữ liệu bảng điểm môn học."""
    if file_path.endswith(".xlsx"):
        df = pd.read_excel(file_path, dtype=str)
    else:
        df = pd.read_csv(file_path, dtype=str)

    # Nhận diện và chuẩn hóa cột mã định danh linh hoạt
    id_col = None
    for col in df.columns:
        if "mã định danh" in str(col).lower() or "id" in str(col).lower():
            id_col = col
            break
    if not id_col and len(df.columns) > 0:
        id_col = df.columns[0]

    if id_col:
        df.rename(columns={id_col: "Mã định danh"}, inplace=True)
        df["Mã định danh"] = df["Mã định danh"].apply(clean_id)
        df = df[df["Mã định danh"] != ""]

    # Nhận diện và chuẩn hóa cột họ tên linh hoạt
    name_col = None
    for col in df.columns:
        if "họ" in str(col).lower() or "tên" in str(col).lower() or "name" in str(col).lower():
            name_col = col
            break
    if name_col and name_col != "Họ và Tên":
        df.rename(columns={name_col: "Họ và Tên"}, inplace=True)

    return df


# CSS chỉ áp dụng cho các bảng điểm HTML, không thay đổi dữ liệu hay xử lý nghiệp vụ.
# Đặt HTML ở đầu dòng (không thụt 4 khoảng trắng) để Streamlit không hiển thị nó như code.
EDU_TABLE_STYLE = """<style>
.edu-table-wrap {
    width: 100%;
    max-width: 100%;
    max-height: 420px;                  /* Không hiển thị toàn bộ các dòng cùng lúc */
    overflow: auto;                     /* Cuộn dọc và ngang khi bảng dài / rộng */
    border: 1px solid #dbe4ef;
    border-radius: 12px;
    background: var(--background-color, #ffffff);
    box-shadow: 0 2px 12px rgba(15, 23, 42, 0.05);
    scrollbar-width: thin;
}
.edu-table-wrap .edu-html-table {
    width: 100%;
    min-width: max-content;             /* Nhiều cột thì cuộn ngang, không bị ép chữ */
    margin: 0;
    border: 0;
    border-collapse: separate;
    border-spacing: 0;
    font-size: 14px;
    line-height: 1.45;
    text-align: left;
    color: var(--text-color, #1f2937);
}
.edu-table-wrap .edu-html-table thead th {
    position: sticky;
    top: 0;
    z-index: 1;
    background: #1e3a5f;
    color: #ffffff;
    font-weight: 650;
    white-space: nowrap;
    border-bottom: 1px solid #24466f;
    padding: 12px 14px;
}
.edu-table-wrap .edu-html-table tbody td {
    padding: 10px 14px;
    white-space: nowrap;
    border-bottom: 1px solid #e2e8f0;
    background: var(--background-color, #ffffff);
}
.edu-table-wrap .edu-html-table tbody tr:nth-child(even) td {
    background: color-mix(in srgb, var(--background-color, #ffffff) 96%, #64748b 4%);
}
.edu-table-wrap .edu-html-table tbody tr:hover td {
    background: color-mix(in srgb, var(--background-color, #ffffff) 88%, #3b82f6 12%);
}
.edu-table-wrap .edu-html-table th + th,
.edu-table-wrap .edu-html-table td + td {
    border-left: 1px solid #e2e8f030;
}
</style>"""


def render_html_table(df: pd.DataFrame):
    """
    Render bảng HTML thật có CSS và vùng cuộn giới hạn chiều cao.
    Các ô là HTML <th>/<td> thật để sửa text bằng DevTools > Elements.
    Không bật contenteditable để tránh chỉnh sửa trực tiếp trên giao diện.
    Thay đổi chỉ nằm ở DOM, không cập nhật DataFrame hay file Excel/CSV.
    """
    table_html = df.to_html(
        index=False,
        escape=True,       # Text dữ liệu được escape để không chèn HTML/JS.
        na_rep="",
        border=0,
        classes="edu-html-table",
        justify="left",
    )

    # Giữ nguyên th/td do pandas.to_html() tạo ra, không thêm contenteditable.

    # Cố ý tránh f-string HTML bị thụt lề: Markdown có thể hiểu thành code block.
    st.markdown(EDU_TABLE_STYLE, unsafe_allow_html=True)
    st.markdown(
        '<div class="edu-table-wrap" role="region" aria-label="Bảng điểm">'
        + table_html
        + '</div>',
        unsafe_allow_html=True,
    )


# ==========================================
# 2. CÔNG THỨC TÍNH ĐTBMÔN CHUẨN BỘ GD&ĐT
# ==========================================
def calculate_subject_dtb(row, score_cols: list) -> float:
    """
    Tính ĐTB Môn theo đúng quy định hệ số Bộ GD&ĐT:
    - Cột Cuối kỳ (CK, CUỐI KỲ, CUOI KY) -> Hệ số 3
    - Cột Giữa kỳ (GK, GIỮA KỲ, GIUA KY) -> Hệ số 2
    - Các cột điểm còn lại -> Hệ số 1
    """
    total_weighted_score = 0.0
    total_weight = 0

    for col in score_cols:
        val = row.get(col)
        if pd.isna(val) or val is None or str(val).strip() == "":
            continue
        try:
            score = float(str(val).replace(",", "."))
        except ValueError:
            continue

        col_upper = str(col).upper()
        if "CK" in col_upper or "CUỐI KỲ" in col_upper or "CUOI KY" in col_upper:
            weight = 3
        elif "GK" in col_upper or "GIỮA KỲ" in col_upper or "GIUA KY" in col_upper:
            weight = 2
        else:
            weight = 1

        total_weighted_score += score * weight
        total_weight += weight

    if total_weight == 0:
        return None
    return round(total_weighted_score / total_weight, 2)


# ==========================================
# 3. GIAO DIỆN CHÍNH & CHỌN PHÂN HỆ
# ==========================================
st.title("🎓 EduClass-AI: Trợ lý Sư phạm số Lớp 6/6")
st.caption("Mô hình Trợ lý Sư phạm số dành cho Lớp Chủ nhiệm (Micro-SaaS)")

st.sidebar.title("📌 Danh mục Phân hệ")
role = st.sidebar.radio(
    "Chọn vai trò sử dụng:",
    ["👨‍👩‍👦 Học sinh / Phụ huynh", "👨‍🏫 Giáo viên Bộ môn", "👩‍🏫 Giáo viên Chủ nhiệm"]
)

st.sidebar.markdown("---")
st.sidebar.info("Hệ thống quản lý điểm số và tư vấn học tập cá nhân hóa Lớp 6/6.")


# ==========================================
# PHÂN HỆ 1: GIÁO VIÊN BỘ MÔN (GVBM - BẢO MẬT 221219)
# ==========================================
if role == "👨‍🏫 Giáo viên Bộ môn":
    st.header("👨‍🏫 Phân hệ Giáo viên Bộ môn (GVBM)")
    st.markdown("🔒 **Bảo mật phân hệ**: Vui lòng nhập mật khẩu giáo viên bộ môn để tiếp tục.")

    entered_gvbm_pass = st.text_input("Nhập Mật khẩu Giáo viên Bộ môn:", type="password")

    if not entered_gvbm_pass:
        st.warning("Vui lòng nhập mật khẩu giáo viên bộ môn để xác thực.")
    elif entered_gvbm_pass != GVBM_PASSWORD:
        st.error("Mật khẩu giáo viên bộ môn không chính xác! Vui lòng thử lại.")
    else:
        st.success("Xác thực thành công! Chào mừng Giáo viên Bộ môn.")
        st.markdown("---")
        st.markdown("Tải lên và cập nhật điểm số môn học theo cơ chế ghép dồn (Incremental Merging).")

        subject_input = st.text_input("Nhập tên môn học phụ trách (Ví dụ: Toán, Ngữ văn, Tiếng Anh...):")
        uploaded_file = st.file_uploader("Tải lên file Excel/CSV điểm mới:", type=["xlsx", "xls", "csv"])

        if subject_input and uploaded_file:
            try:
                if uploaded_file.name.endswith(".csv"):
                    df_upload = pd.read_csv(uploaded_file, dtype=str)
                else:
                    df_upload = pd.read_excel(uploaded_file, dtype=str)

                st.success("Tải file thành công! Hãy chọn các cột thông tin tương ứng dưới đây:")

                cols = list(df_upload.columns)
                col_id = st.selectbox("Chọn cột chứa 'Mã định danh Bộ GD&ĐT':", cols, index=0)
                col_name = st.selectbox("Chọn cột chứa 'Họ và Tên':", cols, index=min(1, len(cols)-1))

                remaining_cols = [c for c in cols if c not in [col_id, col_name]]
                selected_score_cols = st.multiselect(
                    "Chọn các cột điểm cần đưa vào hệ thống (tự động loại bỏ cột rác):",
                    options=remaining_cols,
                    default=remaining_cols
                )

                if st.button("💾 Gộp và Lưu vào Hệ thống", type="primary"):
                    if not selected_score_cols:
                        st.error("Vui lòng chọn ít nhất 1 cột điểm!")
                    else:
                        df_upload[col_id] = df_upload[col_id].apply(clean_id)
                        df_upload = df_upload[df_upload[col_id] != ""]

                        df_new = df_upload[[col_id, col_name] + selected_score_cols].copy()
                        df_new.rename(columns={col_id: "Mã định danh", col_name: "Họ và Tên"}, inplace=True)

                        file_name = f"{subject_input.strip()}_DiemTongHop.xlsx"
                        save_path = os.path.join(DATA_DIR, file_name)

                        # Cơ chế Ghép dồn dữ liệu an toàn chống lỗi KeyError
                        if os.path.exists(save_path):
                            df_existing = load_subject_df(save_path)
                            merge_key_old = "Mã định danh" if "Mã định danh" in df_existing.columns else df_existing.columns[0]
                            merge_key_new = "Mã định danh"

                            df_merged = pd.merge(
                                df_existing, df_new,
                                left_on=[merge_key_old],
                                right_on=[merge_key_new],
                                how="outer",
                                suffixes=('', '_new')
                            )
                            if merge_key_old != merge_key_new and merge_key_old in df_merged.columns:
                                df_merged.drop(columns=[merge_key_old], inplace=True)

                            if "Họ và Tên_new" in df_merged.columns:
                                df_merged["Họ và Tên"] = df_merged["Họ và Tên"].fillna(df_merged["Họ và Tên_new"])
                                df_merged.drop(columns=["Họ và Tên_new"], inplace=True)

                            for col in selected_score_cols:
                                if f"{col}_new" in df_merged.columns:
                                    df_merged[col] = df_merged[f"{col}_new"].fillna(df_merged[col])
                                    df_merged.drop(columns=[f"{col}_new"], inplace=True)
                        else:
                            df_merged = df_new

                        df_merged.to_excel(save_path, index=False)
                        st.success(f"Đã lưu bảng điểm môn **{subject_input}** thành công tại `{save_path}`!")
                        render_html_table(df_merged)

            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi xử lý file: {e}")

        st.markdown("---")
        st.subheader("📋 Xem lại các bảng điểm đã lưu trên hệ thống")
        existing_files = get_subject_files()
        if existing_files:
            selected_view_file = st.selectbox("Chọn môn để xem lại bảng điểm:", existing_files)
            if selected_view_file:
                df_view = load_subject_df(os.path.join(DATA_DIR, selected_view_file))
                render_html_table(df_view)
        else:
            st.info("Chưa có dữ liệu môn học nào được lưu.")


# ==========================================
# PHÂN HỆ 2: GIÁO VIÊN CHỦ NHIỆM (GVCN - BẢO MẬT 221220)
# ==========================================
elif role == "👩‍🏫 Giáo viên Chủ nhiệm":
    st.header("👩‍🏫 Phân hệ Giáo viên Chủ nhiệm (GVCN)")
    st.markdown("🔒 **Bảo mật phân hệ**: Vui lòng nhập mật khẩu chủ nhiệm để truy cập số liệu tổng hợp lớp.")

    entered_password = st.text_input("Nhập Mật khẩu Chủ nhiệm:", type="password")

    if not entered_password:
        st.warning("Vui lòng nhập mật khẩu chủ nhiệm để tiếp tục.")
    elif entered_password != GVCN_PASSWORD:
        st.error("Mật khẩu chủ nhiệm không chính xác! Vui lòng thử lại.")
    else:
        st.success("Xác thực thành công! Chào mừng Giáo viên Chủ nhiệm.")
        st.markdown("---")

        existing_files = get_subject_files()
        if not existing_files:
            st.warning("Hiện chưa có dữ liệu bảng điểm nào. GVBM cần tải file lên trước.")
        else:
            selected_file = st.selectbox("Chọn môn học cần theo dõi chi tiết:", existing_files)
            subject_name = selected_file.replace("_DiemTongHop.xlsx", "").replace("_DiemTongHop.csv", "")

            df_subject = load_subject_df(os.path.join(DATA_DIR, selected_file))
            score_cols = [c for c in df_subject.columns if c not in ["Mã định danh", "Họ và Tên"]]

            df_subject["ĐTB Môn"] = df_subject.apply(lambda row: calculate_subject_dtb(row, score_cols), axis=1)

            logs = load_search_logs()
            df_subject["Số lượt tra cứu"] = df_subject["Mã định danh"].apply(
                lambda sid: logs.get(sid, {}).get(subject_name, 0)
            )

            total_students = len(df_subject)
            avg_score = df_subject["ĐTB Môn"].dropna().mean()
            avg_score_str = f"{avg_score:.2f}" if pd.notna(avg_score) else "N/A"
            gioi_count = (df_subject["ĐTB Môn"] >= 8.0).sum()
            total_searches = df_subject["Số lượt tra cứu"].sum()

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("👥 Tổng số Học sinh", f"{total_students} HS")
            m2.metric("📊 ĐTB Môn Cả Lớp", avg_score_str)
            m3.metric("🌟 Số HS Đạt Loại Giỏi (>=8.0)", f"{gioi_count} HS")
            m4.metric("🔍 Tổng Lượt Tra Cứu", f"{total_searches} lượt")

            st.subheader(f"📊 Bảng điểm chi tiết môn: {subject_name}")
            render_html_table(df_subject)


# ==========================================
# PHÂN HỆ 3: HỌC SINH / PHỤ HUYNH
# ==========================================
elif role == "👨‍👩‍👦 Học sinh / Phụ huynh":
    st.header("👨‍👩‍👦 Phân hệ Tra cứu Cá nhân (Học sinh & Phụ huynh)")
    st.info("🔒 **Bảo mật tuyệt đối**: Nhập chính xác Mã định danh và Mã PIN (4 chữ số cuối của Mã định danh) để tra cứu.")

    c1, c2 = st.columns(2)
    with c1:
        input_id = st.text_input("Nhập Mã định danh Bộ GD&ĐT:", type="default").strip()
    with c2:
        input_pin = st.text_input("Nhập Mã PIN bảo mật (4 chữ số cuối):", type="password").strip()

    if st.button("🔍 Tra cứu Kết quả Học tập", type="primary"):
        cleaned_input_id = clean_id(input_id)

        if not cleaned_input_id or not input_pin:
            st.error("Vui lòng nhập đầy đủ Mã định danh và Mã PIN!")
        elif len(cleaned_input_id) < 4 or input_pin != cleaned_input_id[-4:]:
            st.error("Mã PIN không hợp lệ! Mã PIN chính là 4 chữ số cuối của Mã định danh.")
        else:
            subject_files = get_subject_files()
            if not subject_files:
                st.warning("Hệ thống chưa có dữ liệu điểm môn học nào.")
            else:
                student_name = None
                combined_records = []
                subject_scores = {}

                for f in subject_files:
                    sub_name = f.replace("_DiemTongHop.xlsx", "").replace("_DiemTongHop.csv", "")
                    df_sub = load_subject_df(os.path.join(DATA_DIR, f))

                    row = df_sub[df_sub["Mã định danh"] == cleaned_input_id]
                    if not row.empty:
                        record_search(cleaned_input_id, sub_name)

                        r = row.iloc[0]
                        if not student_name and "Họ và Tên" in r and pd.notna(r["Họ và Tên"]):
                            student_name = r["Họ và Tên"]

                        s_cols = [c for c in df_sub.columns if c not in ["Mã định danh", "Họ và Tên"]]
                        dtb = calculate_subject_dtb(r, s_cols)

                        rec = {"Môn học": sub_name}
                        for sc in s_cols:
                            rec[sc] = r[sc] if pd.notna(r[sc]) else "-"
                        rec["ĐTB Môn"] = dtb if dtb is not None else "Chưa đủ điểm"

                        combined_records.append(rec)
                        if dtb is not None:
                            subject_scores[sub_name] = dtb

                if not combined_records:
                    st.error("Không tìm thấy thông tin điểm ứng với Mã định danh này.")
                else:
                    st.success(f"👋 Xin chào Học sinh: **{student_name if student_name else 'Học sinh'}**")

                    st.subheader("📋 Bảng Điểm Tổng Hợp Các Môn Học")
                    df_student_summary = pd.DataFrame(combined_records)
                    render_html_table(df_student_summary)

                    st.markdown("---")
                    st.subheader("🤖 AI Feedback & Lời Khuyên Sư Phạm Cá Nhân Hóa")

                    display_name = student_name if student_name else "em"

                    if not subject_scores:
                        st.info(f"Hiện tại chưa có đủ dữ liệu điểm để đưa ra đánh giá cho {display_name}.")
                    else:
                        best_subject = max(subject_scores, key=subject_scores.get)
                        best_score = subject_scores[best_subject]

                        # Đa dạng hóa lời khen động bằng random.choice
                        compliment_templates = [
                            f"Thầy/Cô vô cùng ấn tượng khi thấy **{display_name}** đã bứt phá ngoạn mục và đạt thành tích dẫn đầu ở môn **{best_subject}** với điểm số tuyệt vời **{best_score}**!",
                            f"Thật tự hào về sự tập trung và năng lực tuyệt vời của **{display_name}**. Môn **{best_subject}** chính là điểm sáng rực rỡ nhất trong bảng điểm kỳ này với ĐTB đạt **{best_score}**.",
                            f"Tuyệt vời quá **{display_name}** ơi! Thầy/Cô rất biểu dương tinh thần học tập nghiêm túc giúp em chinh phục thành công vị trí dẫn đầu môn **{best_subject}** (ĐTB: **{best_score}**).",
                            f"**{display_name}** đã chứng minh sự thông minh và nỗ lực đáng nể của mình khi xuất sắc vươn lên đứng đầu môn **{best_subject}** với mức điểm **{best_score}**!"
                        ]
                        selected_compliment = random.choice(compliment_templates)

                        gioi_subjects = [s for s, score in subject_scores.items() if score >= 8.0]

                        feedback_msg = f"🌟 **Lời nhận xét sư phạm dành cho {display_name}:**\n\n"
                        feedback_msg += f"- {selected_compliment}\n"

                        if gioi_subjects:
                            gioi_str = ", ".join(gioi_subjects)
                            feedback_msg += f"- Các môn học đạt kết quả loại Giỏi (ĐTB >= 8.0) của **{display_name}** bao gồm: **{gioi_str}**.\n"
                        else:
                            feedback_msg += f"- **{display_name}** đang có sự nỗ lực đều đặn ở các môn, hãy tiếp tục phát huy để bứt phá đạt kết quả loại Giỏi trong thời gian tới nhé.\n"

                        feedback_msg += (
                            f"\n💡 **Định hướng học tập tiếp theo:**\n"
                            f"Để gặt hái thêm nhiều thành công hơn nữa trong học tập, **{display_name}** cần đặc biệt ghi nhớ "
                            f"cần cố gắng **chăm chỉ hơn, hoàn thành đầy đủ bài tập về nhà và ôn bài cũ trước khi đến lớp**."
                        )

                        feedback_html = feedback_msg.replace('\n', '<br>')

                        st.markdown(
                            f"""
                            <div style="background-color: #f0f8ff; padding: 18px; border-radius: 10px; border-left: 5px solid #1e90ff;">
                                {feedback_html}
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
