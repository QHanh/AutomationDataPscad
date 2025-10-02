import streamlit as st
import mhi.pscad
import pandas as pd
import matplotlib.pyplot as plt
import os

# --- Đường dẫn chứa project PSCAD ---
BASE_PATH = os.path.abspath('')
PROJECT_FILES = [f for f in os.listdir(BASE_PATH) if f.endswith(".pscx")]

st.title("PSCAD Automation Dashboard")

# --- Chọn file project ---
file_name = st.selectbox("Chọn file PSCAD project (.pscx):", PROJECT_FILES)

if file_name:
    project_name = os.path.splitext(file_name)[0]
    project_path = os.path.join(BASE_PATH, file_name)

    with mhi.pscad.application() as pscad:
        # Load project
        pscad.load(project_path)
        pscad_project = pscad.project(project_name)

        # --- Cấu hình tham số mô phỏng ---
        st.subheader("Cấu hình mô phỏng")
        time_duration = st.number_input("Thời gian mô phỏng (s)", value=0.1, step=0.1)
        time_step = st.number_input("Time step", value=50, step=10)
        sample_step = st.number_input("Sample step", value=50, step=10)

        pscad_project.parameters(
            time_duration=str(time_duration),
            time_step=str(time_step),
            sample_step=str(sample_step),
        )

        # --- Lấy tất cả component ---
        components = pscad_project.find_all()
        comp_options = {f"{c.label} ({c.definition}) [IID={c.iid}]": c for c in components}

        st.subheader("Chọn component để chỉnh tham số")
        selected_label = st.selectbox("Component:", list(comp_options.keys()))
        selected_comp = comp_options[selected_label]

        # Hiển thị và chỉnh sửa tham số
        comp_params = selected_comp.parameters()
        new_params = {}
        st.write("Thông số hiện tại:")
        for k, v in comp_params.items():
            new_params[k] = st.text_input(f"{k}", value=str(v))

        # Cập nhật tham số component
        if st.button("Cập nhật component"):
            selected_comp.parameters(**new_params)
            st.success("Đã cập nhật tham số component!")

        # --- Chạy mô phỏng nhiều lần ---
        st.subheader("Chạy mô phỏng")
        num_runs = st.number_input("Số lần chạy", value=3, step=1, min_value=1)

        if st.button("Bắt đầu mô phỏng"):
            results = {}
            for i in range(1, num_runs + 1):
                out_file = f"Run_{i}"
                pscad_project.parameters(PlotType="OUT", output_filename=out_file)
                pscad_project.run()

                # Đọc dữ liệu
                csv_path = os.path.join(BASE_PATH, f"{project_name}.if12", f"{out_file}_01.out")
                df = pd.read_csv(csv_path, delimiter=r"\s+", header=None, skiprows=1)
                time = df.iloc[:, 0]
                current = df.iloc[:, 1]
                results[i] = (time, current)

            # Hiển thị kết quả
            st.subheader("Kết quả mô phỏng")
            fig, ax = plt.subplots()
            for i, (t, y) in results.items():
                ax.plot(t, y, label=f"Run {i}")
            ax.set_xlabel("Time (s)")
            ax.set_ylabel("Current (A)")
            ax.legend()
            ax.grid(True)
            st.pyplot(fig)
