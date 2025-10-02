import os
import pandas as pd
import xlsxwriter
from scipy.signal import find_peaks
import win32com.client
from PIL import ImageGrab

# --- Hàm tiện ích ---
def convert_out_to_csv(out_file):
    """Chuyển file .out -> .csv"""
    csv_file = out_file.split('.')[0] + '.csv'
    with open(out_file, 'r') as out, open(csv_file, 'w') as csv:
        csv.writelines(",".join(line.split()) + "\n" for line in out)
    return csv_file

def get_all_file_names(working_dir, file_ext):
    """Lấy danh sách file có phần mở rộng chỉ định"""
    return [f.split('.')[0] for f in os.listdir(working_dir) if f.endswith(file_ext)]

def remove_files_with_extensions(src, *exts):
    """Xóa file tạm"""
    for file in os.listdir(src):
        if file.endswith(tuple(exts)):
            os.remove(os.path.join(src, file))

def saveExcelGraphAsPNG(inputExcelFilePath, outputPNGImagePath):
    """Lấy chart từ Excel -> PNG"""
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = 0
    o.DisplayAlerts = 0
    wb = o.Workbooks.Open(inputExcelFilePath)
    sheet = o.Sheets(1)
    for shape in sheet.Shapes:
        shape.Copy()
        image = ImageGrab.grabclipboard()
        image.save(outputPNGImagePath, 'PNG')
    wb.Close(True)
    o.Quit()

# --- Main ---
work_dir = os.getcwd()
out_files = get_all_file_names(work_dir, ".out")
xlsx_files = []

# B1: Xử lý từng file out -> Excel riêng
for fileName in out_files:
    convert_out_to_csv(fileName + ".out")
    data = pd.read_csv(fileName + ".csv")
    Freq = data['F(Hz)'] / 60
    Imped = data['|Z+|(ohms)']
    
    xlfile_name = fileName + ".xlsx"
    workbook = xlsxwriter.Workbook(xlfile_name)
    worksheet = workbook.add_worksheet()
    
    worksheet.write_row('A1', ['Frequency', 'Impedance'])
    worksheet.write_column('A2', Freq)
    worksheet.write_column('B2', Imped)

    chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
    chart.add_series({
        'name': fileName,
        'categories': f'=Sheet1!$A$2:$A${len(Freq)+1}',
        'values': f'=Sheet1!$B$2:$B${len(Imped)+1}',
    })
    chart.set_title({'name': 'Frequency Scan'})
    chart.set_x_axis({'name': 'Frequency Order', 'min': 0, 'max': 50})
    chart.set_y_axis({'name': 'Impedance (Ohm)'})
    chart.set_style(15)
    worksheet.insert_chart('E2', chart)
    workbook.close()
    
    xlsx_files.append(xlfile_name)

# B2: Tạo file tổng hợp AllData.xlsx
all_xlfile = "AllData.xlsx"

# ==== Pass 1: đọc dữ liệu và peak ====
series = []
max_peaks = 0
for xlfile in xlsx_files:
    df = pd.read_excel(xlfile)
    Freq, Imped = df['Frequency'], df['Impedance']
    peaks, props = find_peaks(Imped, height=1)   # có thể chỉnh height/distance nếu cần
    series.append({"name": xlfile, "freq": Freq, "imp": Imped, "peaks": peaks})
    max_peaks = max(max_peaks, len(peaks))

# ==== Pass 2: ghi Excel ====
workbook = xlsxwriter.Workbook(all_xlfile)
worksheet = workbook.add_worksheet()
chart = workbook.add_chart({'type': 'scatter', 'subtype': 'smooth'})
colors = ["#0072BD", "#D95319", "#EDB120", "#7E2F8E", "#77AC30", "#4DBEEE", "#A2142F"]

# Header hàng 1
for col, s in enumerate(series, start=1):
    clean_name = os.path.splitext(s["name"])[0]
    worksheet.write(0, col, clean_name)

# Ghi nhãn peak ở cột A (theo max_peaks)
for j in range(max_peaks):
    worksheet.write(1 + 2*j, 0, f"peak{j+1}")
    worksheet.write(2 + 2*j, 0, "freq")

# Ghi giá trị peak cho từng file
for col, s in enumerate(series, start=1):
    for j in range(max_peaks):
        if j < len(s["peaks"]):
            idx = s["peaks"][j]
            worksheet.write(1 + 2*j, col, float(s["imp"].iloc[idx]))   # peakN
            worksheet.write(2 + 2*j, col, float(s["freq"].iloc[idx]))  # freq
        # else: để trống nếu file này ít peak hơn

# Bảng dữ liệu gốc bắt đầu từ một hàng cố định sau block peak
start_row = 2*max_peaks + 3
worksheet.write(start_row, 0, "Frequency")
worksheet.write_column(start_row+1, 0, series[0]["freq"])  # giả định các file cùng trục Frequency
for col, s in enumerate(series, start=1):
    worksheet.write_column(start_row+1, col, s["imp"])

# Chart
n_rows = len(series[0]["freq"])
A_range = f'=Sheet1!$A${start_row+2}:$A${start_row+1+n_rows}'
for col, s in enumerate(series, start=1):
    col_letter = chr(65+col)  # B, C, ...
    B_range = f'=Sheet1!${col_letter}${start_row+2}:${col_letter}${start_row+1+n_rows}'
    chart.add_series({
        'name': f'=Sheet1!${col_letter}$1',
        'categories': A_range,
        'values': B_range,
        'line': {
            'color': colors[(col-1) % len(colors)],
            'width': 1.5,
        },
    })

chart.set_x_axis({'min': 0, 'max': 50, 'name': 'Frequency Order', 'name_font': {'name': 'Times New Roman', 'size': 9, 'bold': True}, 'num_font': {'name': 'Times New Roman', 'size': 9}})
chart.set_y_axis({'name': 'Impedance (Ohm)', 'name_font': {'name': 'Times New Roman', 'size': 9, 'bold': True}, 'num_font': {'name': 'Times New Roman', 'size': 9}})
chart.set_legend({'position': 'top', 'font': {'name': 'Times New Roman', 'size': 9}})
chart.set_style(15)
worksheet.insert_chart('E2', chart, {'x_scale': 1.2, 'y_scale': 1.2})
workbook.close()

# B3: Xuất PNG tổng hợp
saveExcelGraphAsPNG(os.path.join(work_dir, all_xlfile), os.path.join(work_dir, "AllData.png"))

# B4: Xóa file csv tạm
remove_files_with_extensions(work_dir, ".csv")
