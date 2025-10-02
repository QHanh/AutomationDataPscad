import streamlit as st
import pandas as pd
import os
import xlsxwriter
from scipy.signal import find_peaks
import win32com.client
import pythoncom
from PIL import ImageGrab
import tempfile
import time

SESSION_RESULT_KEY = "processing_result"

COLORS = ["#0072BD", "#D95319", "#EDB120", "#7E2F8E", "#77AC30", "#4DBEEE", "#A2142F"]

def convert_out_to_csv(out_file_path, csv_file_path):
    """Chuyển nội dung file .out -> .csv (giống controller.py)."""
    with open(out_file_path, 'r') as out_f, open(csv_file_path, 'w') as csv_f:
        for line in out_f:
            csv_f.write(",".join(line.split()) + "\n")

def process_and_generate_files(uploaded_files, temp_dir):
    """Hàm chính để xử lý các file được tải lên và tạo ra kết quả."""
    xlsx_files_info = []

    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        base_name = os.path.splitext(file_name)[0]
        out_path = os.path.join(temp_dir, file_name)
        
        with open(out_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        csv_path = os.path.join(temp_dir, base_name + '.csv')
        convert_out_to_csv(out_path, csv_path)

        raw_df = pd.read_csv(csv_path)
        freq = pd.to_numeric(raw_df['F(Hz)'], errors='coerce') / 60
        imp = pd.to_numeric(raw_df['|Z+|(ohms)'], errors='coerce')
        df = pd.DataFrame({'Frequency': freq, 'Impedance': imp}).dropna()

        xl_path = os.path.join(temp_dir, base_name + '.xlsx')
        df.to_excel(xl_path, index=False)
        xlsx_files_info.append({'path': xl_path, 'name': f"{base_name}.xlsx"})

    all_xlfile_path = os.path.join(temp_dir, "AllData.xlsx")

    series_data = []
    max_peaks = 0
    for info in xlsx_files_info:
        df = pd.read_excel(info['path'])
        freq, imped = df['Frequency'], df['Impedance']
        peaks, _ = find_peaks(imped, height=1)
        series_data.append({"name": info['name'], "freq": freq, "imp": imped, "peaks": peaks})
        max_peaks = max(max_peaks, len(peaks))

    workbook = xlsxwriter.Workbook(all_xlfile_path)
    worksheet = workbook.add_worksheet()
    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})

    for col, s in enumerate(series_data, start=1):
        clean_name = os.path.splitext(s["name"])[0]
        worksheet.write(0, col, clean_name)

    for j in range(max_peaks):
        worksheet.write(1 + 2*j, 0, f"peak{j+1}")
        worksheet.write(2 + 2*j, 0, "freq")
    for col, s in enumerate(series_data, start=1):
        for j in range(max_peaks):
            if j < len(s["peaks"]):
                idx = s["peaks"][j]
                worksheet.write(1 + 2*j, col, float(s["imp"].iloc[idx]))
                worksheet.write(2 + 2*j, col, float(s["freq"].iloc[idx]))

    start_row = 2 * max_peaks + 3
    worksheet.write(start_row, 0, "Frequency")
    worksheet.write_column(start_row + 1, 0, series_data[0]["freq"])
    for col, s in enumerate(series_data, start=1):
        worksheet.write_column(start_row + 1, col, s["imp"])

    n_rows = len(series_data[0]["freq"])
    cat_range = f'=Sheet1!$A${start_row + 2}:$A${start_row + 1 + n_rows}'
    for i, s in enumerate(series_data):
        col_letter = chr(65 + i + 1)
        val_range = f'=Sheet1!${col_letter}${start_row + 2}:${col_letter}${start_row + 1 + n_rows}'
        chart.add_series({
            'name': f'=Sheet1!${col_letter}$1',
            'categories': cat_range,
            'values': val_range,
            'line': {
                'color': COLORS[i % len(COLORS)],
                'width': 1.5,
            },
        })

    max_impedance = max(s['imp'].max() for s in series_data)
    chart.set_x_axis({'min': 0, 'max': 50, 'name': 'Frequency Order', 'name_font': {'name': 'Times New Roman', 'size': 9, 'bold': True}, 'num_font': {'name': 'Times New Roman', 'size': 9}})
    chart.set_y_axis({'name': 'Impedance (Ohms)', 'name_font': {'name': 'Times New Roman', 'size': 9, 'bold': True}, 'num_font': {'name': 'Times New Roman', 'size': 9}})
    chart.set_legend({'position': 'top', 'font': {'name': 'Times New Roman', 'size': 9}})
    chart.set_style(15)
    worksheet.insert_chart('E2', chart, {'x_scale': 2, 'y_scale': 2})
    workbook.close()

    all_png_path = os.path.join(temp_dir, "AllData.png")
    save_excel_graph_as_png(all_xlfile_path, all_png_path)
    
    return all_xlfile_path, all_png_path

def save_excel_graph_as_png(input_excel_path, output_png_path):
    """Lấy chart từ Excel -> PNG"""
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        abs_path = os.path.abspath(input_excel_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Không tìm thấy file Excel: {abs_path}")
        retries = 3
        wb = None
        last_err = None
        for _ in range(retries):
            try:
                wb = excel.Workbooks.Open(abs_path)
                break
            except Exception as err:
                last_err = err
                time.sleep(1)
        if wb is None:
            raise RuntimeError(f"Excel không thể mở workbook: {abs_path}. Chi tiết: {last_err}")
        sheet = wb.Sheets(1)
        
        time.sleep(2)

        for i, shape in enumerate(sheet.Shapes):
            if "Chart" in shape.Name:
                shape.Copy()
                image = ImageGrab.grabclipboard()
                if image:
                    image.save(output_png_path, 'PNG')
                    break
        # for shape in sheet.Shapes:
        #     if "Chart" in shape.Name:
        #         chart_obj = shape.Chart
        #         chart_obj.Export(output_png_path, FilterName='PNG')
        #         break

        # # Option: upscale ảnh
        # from PIL import Image
        # img = Image.open(output_png_path)
        # new_size = (img.width*2, img.height*2)
        # img = img.resize(new_size, Image.LANCZOS)
        # img.save(output_png_path, dpi=(300, 300))
        wb.Close(SaveChanges=False)
    except Exception as e:
        st.error(f"Lỗi khi xử lý Excel: {e}")
        st.warning("Hãy đảm bảo bạn đang chạy trên Windows và đã cài đặt Microsoft Excel.")
    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()

# --- Giao diện Streamlit ---
st.set_page_config(page_title="Automation Data Processing", layout="wide")
st.title("📊 Automation Data Processing")

st.info("Tải lên các file .out để tự động tạo báo cáo Excel và biểu đồ.")

uploaded_files = st.file_uploader(
    "Chọn file .out", 
    accept_multiple_files=True, 
    type=['out']
)

if SESSION_RESULT_KEY not in st.session_state:
    st.session_state[SESSION_RESULT_KEY] = None

if uploaded_files:
    if st.button("Bắt đầu xử lý", type="primary"):
        with st.spinner('Vui lòng đợi, đang xử lý dữ liệu...'):
            with tempfile.TemporaryDirectory() as temp_dir:
                try:
                    all_xlfile_path, all_png_path = process_and_generate_files(uploaded_files, temp_dir)
                    if not os.path.exists(all_png_path):
                        st.error("Không thể tạo file ảnh PNG. Vui lòng kiểm tra lại.")
                        st.session_state[SESSION_RESULT_KEY] = None
                    else:
                        with open(all_xlfile_path, "rb") as f:
                            excel_bytes = f.read()
                        with open(all_png_path, "rb") as f:
                            png_bytes = f.read()
                        st.session_state[SESSION_RESULT_KEY] = {
                            "excel_bytes": excel_bytes,
                            "png_bytes": png_bytes,
                            "excel_name": "AllDataFinal.xlsx",
                            "png_name": "DataVisualFinal.png"
                        }
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi: {e}")
                    st.session_state[SESSION_RESULT_KEY] = None
else:
    st.session_state[SESSION_RESULT_KEY] = None

result = st.session_state.get(SESSION_RESULT_KEY)
if result:
    st.success("Xử lý hoàn tất!")
    st.subheader("Biểu đồ tổng hợp")
    st.image(result["png_bytes"])

    col1, col2 = st.columns(2)
    col1.download_button(
        label="📥 Tải file Excel",
        data=result["excel_bytes"],
        file_name=result["excel_name"],
        mime="application/vnd.ms-excel"
    )
    col2.download_button(
        label="📥 Tải file PNG",
        data=result["png_bytes"],
        file_name=result["png_name"],
        mime="image/png"
    )
