import mhi.pscad
import pandas as pd
import matplotlib.pyplot as plt
import os

# --- Đường dẫn file PSCAD ---
file_path = os.path.abspath('') + "\\"
file_name = "PSCAD_Automation"

# --- Kết nối với PSCAD ---
with mhi.pscad.application() as pscad:
    # Mở project .pscx
    pscad.load(file_path + file_name + ".pscx")
    pscad_project = pscad.project(file_name)

    # Cấu hình tham số mô phỏng
    pscad_project.parameters(time_duration="0.1")
    pscad_project.parameters(time_step="50")
    pscad_project.parameters(sample_step="50")

    # Lấy component theo ID (ví dụ: resistor)
    resistor = pscad_project.component(776720729)

    # Chạy 5 lần mô phỏng với các giá trị khác nhau
    for i in range(5):
        # Đặt tên file output
        pscad_project.parameters(PlotType="1", output_filename=f"Output{i+1}")

        # Gán giá trị điện trở thay đổi theo vòng lặp
        resistor.parameters(Name="R", R=f"{2*(i+1)} [ohm]")

        # Chạy mô phỏng
        pscad_project.run()

# --- Đọc dữ liệu và vẽ ---
plt.figure(figsize=(8, 5))

for i in range(5):
    # Đọc file output (đã convert sang dạng text/CSV)
    temp = pd.read_csv(
        f"{file_path}{file_name}.if12\\Output{i+1}_01.out",
        delimiter=r"\s+",   # tách theo khoảng trắng/tab
        header=None,
        skiprows=1
    )

    time = temp.iloc[:, 0]    # cột thời gian
    current = temp.iloc[:, 1] # cột giá trị (ví dụ dòng điện)

    plt.plot(time, current, label=f"Run {i+1}")

plt.xlabel("Time (s)")
plt.ylabel("Current (A)")
plt.title("Kết quả mô phỏng PSCAD")
plt.legend()
plt.grid(True)
plt.show()
