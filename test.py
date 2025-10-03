import mhi.pscad
import os

file_path = os.path.abspath('C:\\Users\\hqh14\\OneDrive\\Desktop\\08_19_2025_PSCAD_Model_CN_rev1') + "\\"
file_name = "main_3LG"

with mhi.pscad.connect() as pscad:
    pscad.load(file_path + file_name + ".pscx")
    proj = pscad.project(file_name)
    
    # Lấy danh sách tên parameters
    try:
        params = proj.parameters()
        print("Parameters trong project:")
        for param_name in params:
            print(f"  - {param_name}")
            # Lấy giá trị của từng parameter
            try:
                value = proj.parameter(param_name).value
                print(f"    Value: {value}")
            except:
                pass
    except Exception as e:
        print(f"Error getting parameters: {e}")
    
    # Hoặc đơn giản hơn:
    print("\n=== Danh sách parameters ===")
    try:
        params = proj.parameters()
        print(params)
    except Exception as e:
        print(f"Error: {e}")