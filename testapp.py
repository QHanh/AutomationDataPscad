import streamlit as st
import pandas as pd
import re, os, tempfile, time
import xlsxwriter
from PIL import ImageGrab
import win32com.client, pythoncom
import matplotlib.pyplot as plt

# Mảng màu cho các đường biểu đồ
COLORS = ["#0072BD", "#D95319", "#EDB120", "#7E2F8E", "#77AC30", "#4DBEEE", "#A2142F"]

def parse_inf(inf_text):
    pgb_map = {}
    for line in inf_text.splitlines():
        line = line.strip()
        if line.startswith("PGB("):
            idx = int(line.split("(")[1].split(")")[0])
            desc = re.search(r'Desc="([^"]+)"', line)
            pgb_map[idx] = desc.group(1) if desc else f"PGB{idx}"
    return pgb_map

def extract_num(filename):
    match = re.search(r"(\d+)", filename)
    return int(match.group(1)) if match else 9999

# --- HÀM TẠO EXCEL VỚI BIỂU ĐỒ ---
def generate_excel_with_chart(df, selected_cols, temp_dir):
    xl_path = os.path.join(temp_dir, "AllData.xlsx")
    workbook = xlsxwriter.Workbook(xl_path)
    worksheet = workbook.add_worksheet("Sheet1")
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})

    # Ghi tên cột và dữ liệu
    worksheet.write(0, 0, "Time")
    worksheet.write_column(1, 0, df["Time"])
    
    for col_idx, col_name in enumerate(selected_cols, start=1):
        worksheet.write(0, col_idx, col_name)
        worksheet.write_column(1, col_idx, df[col_name])

    # Thêm series
    num_rows = len(df)
    for i, col_name in enumerate(selected_cols):
        col_letter = chr(66 + i) 
        chart.add_series({
            'name':       f"=Sheet1!${col_letter}$1",
            'categories': f"=Sheet1!$A$2:$A${num_rows + 1}",
            'values':     f"=Sheet1!${col_letter}$2:${col_letter}${num_rows + 1}",
            'line':       {'color': COLORS[i % len(COLORS)], 'width': 1.5}
        })

    # x_max = min(df["Time"].max(), 2)

    chart.set_x_axis({
        'name': 'Frequency',
        'name_font': {'name': 'Times New Roman', 'size': 9, 'bold': True},
        'num_font': {'name': 'Times New Roman', 'size': 9},
        # 'min': 0,
        # 'max': x_max
    })
    chart.set_y_axis({
        'name': 'Index',
        'name_font': {'name': 'Times New Roman', 'size': 9, 'bold': True},
        'num_font': {'name': 'Times New Roman', 'size': 9},
    })
    chart.set_legend({'position': 'top', 'font': {'name': 'Times New Roman', 'size': 9}})
    chart.set_style(15)

    worksheet.insert_chart('E2', chart, {'x_scale': 2, 'y_scale': 2})
    workbook.close()
    return xl_path

# --- HÀM LƯU BIỂU ĐỒ EXCEL RA PNG ---
def save_excel_graph_as_png(input_excel_path, output_png_path):
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        abs_path = os.path.abspath(input_excel_path)
        wb = excel.Workbooks.Open(abs_path)
        sheet = wb.Sheets(1)
        time.sleep(2)
        for shape in sheet.Shapes:
            if "Chart" in shape.Name:
                shape.Copy()
                image = ImageGrab.grabclipboard()
                if image:
                    image.save(output_png_path, 'PNG')
                break
        wb.Close(SaveChanges=False)
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()

# --- Giao diện Streamlit ---
st.set_page_config(page_title="HVRT Data Viewer", layout="wide")
st.title("📊 Data Processing Visualization")

# Người dùng chọn loại file OUT
has_header = st.radio("File OUT:", ["Có tên cột (dòng đầu là header)", "Không có tên cột (dùng file INF)"])

# Nếu cần file INF
inf_file = None
if has_header == "Không có tên cột (dùng file INF)":
    st.subheader("1. Upload file INF")
    inf_file = st.file_uploader("Chọn file .inf", type=["inf"])

st.subheader("2. Upload các file OUT")
out_files = st.file_uploader("Chọn nhiều file .out", type=["out"], accept_multiple_files=True)

if (has_header.startswith("Có tên cột") and out_files) or (has_header.startswith("Không có") and inf_file and out_files):
    if st.button("✅ Xác nhận", type="primary"):
        with st.spinner("Đang xử lý dữ liệu..."):
            out_files_sorted = sorted(out_files, key=lambda f: extract_num(f.name))
            df_all, start_idx = None, 1

            if has_header.startswith("Không có"):
                # Dùng file INF
                pgb_map = parse_inf(inf_file.read().decode("utf-8", errors="ignore"))
                for f in out_files_sorted:
                    df = pd.read_csv(f, sep=r"\s+", header=None)
                    col_names = ["Time"] + [pgb_map.get(i, f"PGB{i}") for i in range(start_idx, start_idx + df.shape[1] - 1)]
                    df.columns = col_names
                    if df_all is None:
                        df_all = df
                    else:
                        df_all = df_all.merge(df, on="Time", how="outer")
                    start_idx += df.shape[1] - 1
            else:
                # Có header trong file OUT
                for f in out_files_sorted:
                    df = pd.read_csv(f, sep=r"\s+", header=0)
                    df.rename(columns={df.columns[0]: "Time"}, inplace=True)
                    if df_all is None:
                        df_all = df
                    else:
                        df_all = df_all.merge(df, on="Time", how="outer")

            if "Time" in df_all.columns:
                df_all["Time"] = df_all["Time"] / 60

            st.session_state["df_all"] = df_all
            st.success("Đọc và ghép dữ liệu thành công!")

# --- Vẽ biểu đồ ---
if "df_all" in st.session_state:
    df_all = st.session_state["df_all"]
    options = [c for c in df_all.columns if c != "Time"]
    selected_cols = st.multiselect("Chọn các cột để hiển thị", options, default=options[:3])

    chart_method = st.radio("Chọn phương thức vẽ biểu đồ:", ["Excel (xuất file)", "Matplotlib (nhanh)"])

    if st.button("📊 Vẽ biểu đồ", type="primary"):
        if selected_cols:
            if chart_method == "Matplotlib (nhanh)":
                with st.spinner("Đang vẽ bằng Matplotlib..."):
                    fig, ax = plt.subplots(figsize=(10, 4))
                    for i, col in enumerate(selected_cols):
                        ax.plot(df_all["Time"], df_all[col], label=col, color=COLORS[i % len(COLORS)], linewidth=1.2)
                    ax.set_xlabel("Frequency", fontname="Times New Roman", fontsize=9, fontweight="bold")
                    ax.set_ylabel("Index", fontname="Times New Roman", fontsize=9, fontweight="bold")
                    ax.legend(fontsize=8, loc="upper center", ncol=3)
                    ax.grid(True, linestyle="--", alpha=0.6)
                    st.pyplot(fig)

            else:  # Excel
                with st.spinner("Đang tạo file Excel và trích xuất biểu đồ..."):
                    with tempfile.TemporaryDirectory() as temp_dir:
                        xl_path = generate_excel_with_chart(df_all, selected_cols, temp_dir)
                        png_path = os.path.join(temp_dir, "Chart.png")
                        save_excel_graph_as_png(xl_path, png_path)

                        if os.path.exists(xl_path):
                            with open(xl_path, "rb") as f:
                                excel_bytes = f.read()
                            st.download_button("📥 Tải file Excel (có biểu đồ)", excel_bytes, file_name="AllData_with_Chart.xlsx")

                        if os.path.exists(png_path):
                            st.image(png_path, caption="Biểu đồ từ Excel")
                            with open(png_path, "rb") as f:
                                st.download_button("🖼 Tải file ảnh (.png)", f, file_name="DataChart.png")
        else:
            st.warning("Hãy chọn ít nhất một cột để vẽ.")
