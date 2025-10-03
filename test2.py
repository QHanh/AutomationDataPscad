import mhi.pscad
import os
import pandas as pd
from datetime import datetime

# ======================== CẤU HÌNH ========================
file_path = os.path.abspath('C:\\Users\\hqh14\\OneDrive\\Desktop\\08_19_2025_PSCAD_Model_CN_rev1') + "\\"
file_name = "main_3LG"
canvas_name = "Main"  # Hoặc "main" tùy project

# ======================== 1. EXPORT RA EXCEL ========================
def export_to_excel(output_file="pscad_components.xlsx"):
    """Export tất cả components và parameters ra Excel"""
    
    with mhi.pscad.connect() as pscad:
        pscad.load(file_path + file_name + ".pscx")
        proj = pscad.project(file_name)
        canvas = proj.canvas(canvas_name)
        components = canvas.components()
        
        print(f"📦 Đang export {len(components)} components...")
        
        # Chuẩn bị dữ liệu
        data = []
        
        for i, comp in enumerate(components, 1):
            # Lấy Component ID (IID - unique identifier)
            try:
                comp_iid = comp.iid
            except:
                comp_iid = f"unknown_{i}"
            
            # Lấy vị trí component
            try:
                comp_bounds = str(comp.bounds)
            except:
                comp_bounds = "N/A"
            
            # Lấy definition name (loại component)
            try:
                defn = comp.definition
                defn_name = str(defn).split('[')[-1].split(']')[0] if '[' in str(defn) else str(defn)
            except:
                defn_name = "Unknown"
            
            # Lấy parameters
            try:
                all_params = comp.parameters()  # Dùng parameters() thay vì get_parameters()
                
                # Lấy tên component từ parameter Name nếu có
                comp_name = ""
                if all_params and 'Name' in all_params:
                    comp_name = str(all_params['Name'])
                
                if all_params:
                    for param_name, param_value in all_params.items():
                        data.append({
                            'Component_Index': i,
                            'Component_IID': str(comp_iid),
                            'Component_Name': comp_name,
                            'Component_Type': defn_name,
                            'Component_Location': comp_bounds,
                            'Parameter_Name': param_name,
                            'Current_Value': str(param_value),
                            'New_Value': '',  # Cột để user nhập giá trị mới
                            'Notes': ''  # Cột ghi chú
                        })
                else:
                    # Component không có parameters
                    data.append({
                        'Component_Index': i,
                        'Component_IID': str(comp_iid),
                        'Component_Name': comp_name,
                        'Component_Type': defn_name,
                        'Component_Location': comp_bounds,
                        'Parameter_Name': 'N/A',
                        'Current_Value': 'N/A',
                        'New_Value': '',
                        'Notes': 'No parameters'
                    })
            except Exception as e:
                print(f"  ⚠️  Component {i} (IID: {comp_iid}): Lỗi - {e}")
            
            if i % 20 == 0:
                print(f"  ✓ Đã xử lý {i}/{len(components)} components")
        
        # Tạo DataFrame và export
        df = pd.DataFrame(data)
        
        # Kiểm tra file có đang mở không
        try:
            # Thử mở file để kiểm tra
            with open(output_file, 'a'):
                pass
        except PermissionError:
            # File đang được mở, tạo tên file mới
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = output_file.rsplit('.', 1)[0]
            output_file = f"{base_name}_{timestamp}.xlsx"
            print(f"⚠️  File gốc đang mở, lưu vào: {output_file}")
        
        # Tạo file Excel với formatting
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Components', index=False)
            
            # Format Excel
            worksheet = writer.sheets['Components']
            
            # Auto-fit columns
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Freeze header row
            worksheet.freeze_panes = 'A2'
            
            # Tạo sheet Instructions
            instructions = pd.DataFrame({
                'Step': [1, 2, 3, 4],
                'Instruction': [
                    'Mở sheet "Components"',
                    'Tìm parameters cần thay đổi',
                    'Nhập giá trị mới vào cột "New_Value"',
                    'Lưu file và chạy import_from_excel() trong Python'
                ],
                'Example': [
                    '',
                    'Tìm component có tên "Vrms" hoặc type "master:pgb"',
                    'Ví dụ: thay đổi "Max" từ "2.0" thành "5.0"',
                    'import_from_excel("pscad_components.xlsx", dry_run=False)'
                ]
            })
            instructions.to_excel(writer, sheet_name='Instructions', index=False)
        
        print(f"\n✅ Export thành công!")
        print(f"📁 File: {output_file}")
        print(f"📊 Tổng số parameters: {len(data)}")
        print(f"📊 Tổng số components: {len(df['Component_Index'].unique())}")
        print(f"\n💡 Hướng dẫn:")
        print(f"   1. Mở file Excel")
        print(f"   2. Xem sheet 'Instructions' để biết cách sử dụng")
        print(f"   3. Nhập giá trị mới vào cột 'New_Value' trong sheet 'Components'")
        print(f"   4. Chạy import_from_excel() để cập nhật vào PSCAD")


# ======================== 2. IMPORT TỪ EXCEL ========================
def import_from_excel(excel_file="pscad_components.xlsx", dry_run=True):
    """
    Import và update parameters từ Excel vào PSCAD
    
    Args:
        excel_file: Đường dẫn file Excel
        dry_run: True = chỉ hiển thị thay đổi, False = thực sự update
    """
    
    # Đọc file Excel
    print(f"📖 Đang đọc file: {excel_file}")
    df = pd.read_excel(excel_file, sheet_name='Components')
    
    # Lọc các dòng có New_Value không rỗng
    df_changes = df[df['New_Value'].notna() & (df['New_Value'] != '')]
    
    if len(df_changes) == 0:
        print("⚠️  Không tìm thấy giá trị mới nào trong cột 'New_Value'")
        return
    
    print(f"📝 Tìm thấy {len(df_changes)} thay đổi")
    print("\n" + "="*70)
    
    # Nhóm theo Component_IID (unique identifier)
    changes_by_iid = df_changes.groupby('Component_IID')
    
    with mhi.pscad.connect() as pscad:
        pscad.load(file_path + file_name + ".pscx")
        proj = pscad.project(file_name)
        canvas = proj.canvas(canvas_name)
        components = canvas.components()
        
        # Tạo mapping từ IID sang component object
        comp_map = {}
        for comp in components:
            try:
                comp_map[str(comp.iid)] = comp
            except:
                pass
        
        update_count = 0
        error_count = 0
        
        for comp_iid, changes in changes_by_iid:
            comp_iid_str = str(comp_iid)
            
            # Tìm component theo IID
            if comp_iid_str not in comp_map:
                print(f"❌ Component IID {comp_iid}: Không tìm thấy")
                error_count += 1
                continue
            
            comp = comp_map[comp_iid_str]
            
            # Lấy thông tin component
            try:
                comp_name = changes.iloc[0]['Component_Name']
                comp_type = changes.iloc[0]['Component_Type']
                comp_idx = changes.iloc[0]['Component_Index']
            except:
                comp_name = ""
                comp_type = "Unknown"
                comp_idx = "?"
            
            display_name = f"[{comp_idx}] {comp_type}"
            if comp_name:
                display_name += f" ({comp_name})"
            
            print(f"\n🔧 {display_name}")
            print(f"   IID: {comp_iid}")
            
            # Chuẩn bị dictionary parameters để update
            params_to_update = {}
            
            for _, row in changes.iterrows():
                param_name = row['Parameter_Name']
                old_value = row['Current_Value']
                new_value = row['New_Value']
                
                print(f"   • {param_name}: '{old_value}' → '{new_value}'")
                params_to_update[param_name] = str(new_value)
            
            # Update parameters
            if not dry_run:
                try:
                    # Dùng cách mới: comp.parameters(**params) thay vì comp.set_parameters()
                    comp.parameters(**params_to_update)
                    print(f"   ✅ Đã update thành công!")
                    
                    # Verify ngay sau khi update
                    updated_params = comp.parameters()  # Dùng parameters() thay vì get_parameters()
                    print(f"   🔍 Verify:")
                    for param_name in params_to_update.keys():
                        actual = updated_params.get(param_name, 'N/A')
                        expected = params_to_update[param_name]
                        if str(actual) == str(expected):
                            print(f"      ✓ {param_name} = {actual}")
                        else:
                            print(f"      ✗ {param_name}: expected '{expected}', got '{actual}'")
                    
                    update_count += 1
                except Exception as e:
                    print(f"   ❌ Lỗi update: {e}")
                    error_count += 1
            else:
                print(f"   ℹ️  DRY RUN - không thực sự update")
                update_count += 1
        
        # Save project nếu không phải dry run
        if not dry_run:
            try:
                print(f"\n💾 Đang lưu project...")
                proj.save()
                print(f"✅ Đã lưu project thành công!")
                
                # Unload và load lại để force PSCAD GUI refresh
                print(f"🔄 Đang unload project...")
                proj.unload()
                print(f"✅ Đã unload!")
                
                print(f"📥 Đang load lại project...")
                pscad.load(file_path + file_name + ".pscx")
                print(f"✅ Đã load lại project!")
                print(f"\n💡 PSCAD GUI đã được refresh. Kiểm tra lại component trong GUI.")
                
            except Exception as e:
                print(f"❌ Lỗi: {e}")
        
        print("\n" + "="*70)
        print(f"📊 Tóm tắt:")
        print(f"   • Thành công: {update_count}")
        print(f"   • Lỗi: {error_count}")
        
        if dry_run:
            print(f"\n💡 Đây là DRY RUN mode. Để thực sự update, chạy:")
            print(f"   import_from_excel('{excel_file}', dry_run=False)")


# ======================== 3. TIỆN ÍCH ========================
def export_by_type(component_type, output_file=None):
    """
    Export chỉ một loại component cụ thể
    
    Args:
        component_type: Tên loại component, ví dụ 'master:multimeter'
        output_file: Tên file output (tự động nếu None)
    """
    
    if output_file is None:
        safe_type = component_type.replace(':', '_').replace('/', '_')
        output_file = f"components_{safe_type}.xlsx"
    
    with mhi.pscad.connect() as pscad:
        pscad.load(file_path + file_name + ".pscx")
        proj = pscad.project(file_name)
        canvas = proj.canvas(canvas_name)
        components = canvas.components()
        
        data = []
        
        for i, comp in enumerate(components, 1):
            try:
                comp_iid = comp.iid
                comp_bounds = str(comp.bounds)
                defn = comp.definition
                defn_name = str(defn).split('[')[-1].split(']')[0] if '[' in str(defn) else str(defn)
            except:
                continue
            
            # Filter theo type
            if defn_name != component_type:
                continue
            
            try:
                all_params = comp.get_parameters()
                comp_name = ""
                if all_params and 'Name' in all_params:
                    comp_name = str(all_params['Name'])
                
                if all_params:
                    for param_name, param_value in all_params.items():
                        data.append({
                            'Component_Index': i,
                            'Component_IID': str(comp_iid),
                            'Component_Name': comp_name,
                            'Component_Type': defn_name,
                            'Component_Location': comp_bounds,
                            'Parameter_Name': param_name,
                            'Current_Value': str(param_value),
                            'New_Value': '',
                            'Notes': ''
                        })
            except:
                pass
        
        if len(data) == 0:
            print(f"⚠️  Không tìm thấy component nào có type: {component_type}")
            return
        
        df = pd.DataFrame(data)
        df.to_excel(output_file, index=False)
        
        print(f"✅ Export {len(data)} parameters từ {len(df['Component_Index'].unique())} components")
        print(f"📁 File: {output_file}")


def list_component_types():
    """Liệt kê tất cả các loại components trong project"""
    
    with mhi.pscad.connect() as pscad:
        pscad.load(file_path + file_name + ".pscx")
        proj = pscad.project(file_name)
        canvas = proj.canvas(canvas_name)
        components = canvas.components()
        
        types = {}
        
        for comp in components:
            try:
                defn = comp.definition
                defn_name = str(defn).split('[')[-1].split(']')[0] if '[' in str(defn) else str(defn)
                types[defn_name] = types.get(defn_name, 0) + 1
            except:
                pass
        
        print(f"📊 Tổng số: {len(components)} components\n")
        print("Component Types:")
        print("-" * 50)
        
        for comp_type, count in sorted(types.items(), key=lambda x: x[1], reverse=True):
            print(f"  {comp_type}: {count}")


# ======================== WORKFLOW AN TOÀN ========================
def safe_import(excel_file="pscad_components.xlsx", dry_run=False):
    """
    Import an toàn - cảnh báo nếu PSCAD GUI đang chạy
    """
    import subprocess
    
    # Kiểm tra PSCAD có đang chạy không
    try:
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq PSCAD*.exe'], 
                              capture_output=True, text=True)
        
        if 'PSCAD' in result.stdout:
            print("\n" + "="*70)
            print("⚠️  CẢNH BÁO: PSCAD GUI ĐANG CHẠY!")
            print("="*70)
            print("\nĐể thay đổi hiển thị trong GUI, bạn CẦN:")
            print("  1. Đóng PSCAD GUI hoàn toàn")
            print("  2. Chạy lại hàm này")
            print("  3. Sau khi update xong, mở PSCAD GUI và load project")
            print("\n💡 Giá trị VẪN SẼ ĐƯỢC LƯU vào file, nhưng GUI không refresh!")
            print("="*70)
            
            response = input("\nBạn có muốn tiếp tục? (y/n): ")
            if response.lower() != 'y':
                print("Đã hủy.")
                return
    except:
        pass  # Không check được cũng không sao
    
    # Thực hiện import
    import_from_excel(excel_file, dry_run)


# ======================== SỬ DỤNG ========================
if __name__ == "__main__":
    print("""
    ╔══════════════════════════════════════════════════════════════╗
    ║           PSCAD PARAMETER MANAGEMENT TOOL                    ║
    ╚══════════════════════════════════════════════════════════════╝
    
    WORKFLOW KHUYẾN NGHỊ:
    
    1. Export parameters:
       export_to_excel("params.xlsx")
    
    2. Sửa file Excel (cột "New_Value")
    
    3. ĐÓNG PSCAD GUI hoàn toàn
    
    4. Import và update:
       safe_import("params.xlsx", dry_run=False)
    
    5. Mở PSCAD GUI và load project để xem thay đổi
    
    ──────────────────────────────────────────────────────────────
    """)
    
    # 1. Xem danh sách các loại components
    # list_component_types()
    
    # 2. Export tất cả ra Excel
    # export_to_excel("pscad_components.xlsx")
    
    # 3. Export chỉ một loại component
    # export_by_type("master:multimeter", "multimeters.xlsx")
    
    # 4. Sau khi sửa Excel, xem trước thay đổi (DRY RUN)
    # safe_import("pscad_components.xlsx", dry_run=True)
    
    # 5. Update thực sự vào PSCAD (ĐÓNG PSCAD GUI TRƯỚC)
    safe_import("pscad_components.xlsx", dry_run=False)