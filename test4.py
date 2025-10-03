import streamlit as st
import pandas as pd
import re, os, tempfile, time
import xlsxwriter
from scipy.signal import find_peaks
from PIL import ImageGrab
import win32com.client, pythoncom

# Mảng màu cho các đường biểu đồ (giữ nguyên)
COLORS = ["#0072BD", "#D95319", "#EDB120", "#7E2F8E", "#77AC30", "#4DBEEE", "#A2142F"]

# --- Các hàm xử lý file (giữ nguyên) ---
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

# --- HÀM TẠO EXCEL VỚI BIỂU ĐỒ NHÚNG ---
def generate_excel_with_chart(df, selected_cols, temp_dir):
    """
    Tạo file Excel chứa cả dữ liệu và biểu đồ nhúng.
    Trả về đường dẫn đến file Excel đã tạo.
    """
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

    # Thêm các series dữ liệu vào biểu đồ
    num_rows = len(df)
    for i, col_name in enumerate(selected_cols):
        # Lấy chữ cái của cột trong Excel (B, C, D, ...)
        col_letter = chr(66 + i) 
        chart.add_series({
            'name':       f"=Sheet1!${col_letter}$1",
            'categories': f"=Sheet1!$A$2:$A${num_rows + 1}",
            'values':     f"=Sheet1!${col_letter}$2:${col_letter}${num_rows + 1}",
            'line':       {'color': COLORS[i % len(COLORS)], 'width': 1.5}
        })

    # Tùy chỉnh giao diện biểu đồ (giống với file app_auto_process_out_pscad.py)
    chart.set_x_axis({
        'name': 'Frequency Order',
        'name_font': {'name': 'Times New Roman', 'size': 9, 'bold': True},
        'num_font': {'name': 'Times New Roman', 'size': 9},
        'min': 0,
        'max': 50
    })
    chart.set_y_axis({
        'name': 'Impedance (Ohms)',
        'name_font': {'name': 'Times New Roman', 'size': 9, 'bold': True},
        'num_font': {'name': 'Times New Roman', 'size': 9}
    })
    chart.set_legend({'position': 'top', 'font': {'name': 'Times New Roman', 'size': 9}})
    chart.set_style(15)

    # Chèn biểu đồ vào worksheet và phóng to
    worksheet.insert_chart('E2', chart, {'x_scale': 2, 'y_scale': 2})
    workbook.close()

    return xl_path

# --- HÀM LƯU BIỂU ĐỒ TỪ EXCEL RA PNG (DÙNG WIN32COM) ---
def save_excel_graph_as_png(input_excel_path, output_png_path):
    """
    Mở file Excel, sao chép biểu đồ và lưu thành file PNG.
    Đây là bước chậm và phụ thuộc vào môi trường.
    """
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        abs_path = os.path.abspath(input_excel_path)
        wb = excel.Workbooks.Open(abs_path)
        sheet = wb.Sheets(1)
        
        # Đợi một chút để Excel có thời gian render biểu đồ
        time.sleep(2)

        # Tìm biểu đồ trong sheet và sao chép
        for shape in sheet.Shapes:
            if "Chart" in shape.Name:
                shape.Copy()
                image = ImageGrab.grabclipboard()
                if image:
                    image.save(output_png_path, 'PNG')
                break
        wb.Close(SaveChanges=False)
    except Exception as e:
        st.error(f"Lỗi khi tương tác với Excel: {e}")
        st.warning("Hãy đảm bảo bạn đang chạy trên Windows và đã cài đặt Microsoft Excel.")
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()

# --- Giao diện Streamlit ---
st.set_page_config(page_title="HVRT Data Viewer", layout="wide")
st.title("📊 HVRT Data Visualization")

# ... (Phần upload file giữ nguyên) ...
st.subheader("1. Upload file INF")
inf_file = st.file_uploader("Chọn file .inf", type=["inf"])

st.subheader("2. Upload các file OUT")
st.info("⚠️ Hãy upload file .out theo thứ tự, hoặc hệ thống sẽ dựa vào số thứ tự trong tên file.")
out_files = st.file_uploader("Chọn nhiều file .out", type=["out"], accept_multiple_files=True)

if inf_file and out_files:
    if st.button("✅ Xác nhận", type="primary"):
        with st.spinner("Đang xử lý dữ liệu..."):
            pgb_map = parse_inf(inf_file.read().decode("utf-8", errors="ignore"))
            out_files_sorted = sorted(out_files, key=lambda f: extract_num(f.name))
            df_all, start_idx = None, 1
            for f in out_files_sorted:
                df = pd.read_csv(f, sep=r"\s+", header=None)
                col_names = ["Time"] + [pgb_map.get(i, f"PGB{i}") for i in range(start_idx, start_idx + df.shape[1] - 1)]
                df.columns = col_names
                if df_all is None:
                    df_all = df
                else:
                    df_all = df_all.merge(df, on="Time", how="outer")
                start_idx += df.shape[1] - 1
            st.session_state["df_all"] = df_all
            st.success("Đọc và ghép dữ liệu thành công!")

if "df_all" in st.session_state:
    df_all = st.session_state["df_all"]
    options = [c for c in df_all.columns if c != "Time"]
    selected_cols = st.multiselect("Chọn các cột để hiển thị", options, default=options[:3] if len(options) > 2 else options)

    if st.button("📊 Vẽ biểu đồ và xuất file", type="primary"):
        if selected_cols:
            with st.spinner("Đang tạo file Excel và trích xuất biểu đồ... Vui lòng đợi, quá trình này có thể mất một lúc."):
                # Sử dụng thư mục tạm để lưu file
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 1. Tạo file Excel với biểu đồ nhúng
                    xl_path = generate_excel_with_chart(df_all, selected_cols, temp_dir)
                    
                    # 2. Mở file Excel đó để trích xuất biểu đồ ra PNG
                    png_path = os.path.join(temp_dir, "Chart.png")
                    save_excel_graph_as_png(xl_path, png_path)

                    # 3. Đọc dữ liệu từ các file đã tạo để cung cấp cho việc tải xuống
                    if os.path.exists(xl_path):
                        with open(xl_path, "rb") as f:
                            excel_bytes = f.read()
                        st.session_state['excel_bytes'] = excel_bytes
                    else:
                        st.session_state['excel_bytes'] = None
                        
                    if os.path.exists(png_path):
                        with open(png_path, "rb") as f:
                            png_bytes = f.read()
                        st.session_state['png_bytes'] = png_bytes
                    else:
                        st.session_state['png_bytes'] = None

            # Hiển thị kết quả sau khi xử lý xong
            if st.session_state.get('png_bytes'):
                st.image(st.session_state['png_bytes'], caption="Biểu đồ được trích xuất từ file Excel")

                col1, col2 = st.columns(2)
                if st.session_state.get('excel_bytes'):
                    col1.download_button("📥 Tải file Excel (có biểu đồ)", st.session_state['excel_bytes'], file_name="AllData_with_Chart.xlsx")
                col2.download_button("🖼 Tải file ảnh (.png)", st.session_state['png_bytes'], file_name="DataChart.png")
            else:
                 st.error("Không thể tạo file ảnh. Vui lòng kiểm tra lại môi trường và đảm bảo Excel đã được cài đặt.")

            # Phân tích và hiển thị peaks (giữ nguyên)
            st.subheader("🔎 Phân tích Peaks")
            for c in selected_cols:
                peaks, _ = find_peaks(df_all[c].dropna(), height=1)
                if len(peaks) > 0:
                    st.write(f"**{c}** có {len(peaks)} peaks tại các điểm Time: {df_all['Time'].iloc[peaks].round(4).tolist()}")
                else:
                    st.write(f"**{c}** không tìm thấy peak nào.")
        else:
            st.warning("Hãy chọn ít nhất một cột để vẽ.")