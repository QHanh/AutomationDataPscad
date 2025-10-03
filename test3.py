import mhi.pscad
import os
import pandas as pd
from datetime import datetime
import subprocess
from typing import Optional, List, Dict, Any
from pathlib import Path

class PscadManager:
    """
    A class to manage interactions with a PSCAD project, including exporting
    and importing component parameters via Excel.
    """

    def __init__(self, project_path: str, project_name: str, canvas_name: str = "Main"):
        """
        Initializes the PscadManager.

        Args:
            project_path (str): The absolute path to the directory containing the .pscx file.
            project_name (str): The name of the project (without the .pscx extension).
            canvas_name (str): The name of the canvas to work on.
        """
        self.project_path = Path(project_path)
        self.project_name = project_name
        self.canvas_name = canvas_name
        self.pscx_file = self.project_path / f"{self.project_name}.pscx"

        if not self.pscx_file.exists():
            raise FileNotFoundError(f"PSCAD project file not found at: {self.pscx_file}")

        self.pscad = None
        self.project = None
        self.canvas = None

    def __enter__(self):
        """Connect to PSCAD and load the project."""
        print(f"🔌 Connecting to PSCAD and loading project '{self.project_name}'...")
        self.pscad = mhi.pscad.application()
        self.pscad.load(str(self.pscx_file))
        self.project = self.pscad.project(self.project_name)
        self.canvas = self.project.canvas(self.canvas_name)
        print("✅ Connection successful.")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Clean up resources."""
        if self.project and not exc_type:
             # Only unload if no error occurred during the process
            print("🔄 Unloading project...")
            self.project.unload()
        print("🔌 Disconnected from PSCAD.")

    def _get_component_data(self, components: List[Any]) -> List[Dict[str, Any]]:
        """Extracts detailed parameter data from a list of components."""
        data = []
        total = len(components)
        print(f"📦 Processing {total} components...")

        for i, comp in enumerate(components, 1):
            try:
                comp_iid = str(comp.iid)
                defn_str = str(comp.definition)
                comp_type = defn_str.split('[')[-1].split(']')[0] if '[' in defn_str else defn_str
                
                params = comp.parameters()
                comp_name = str(params.get('Name', ''))

                if params:
                    for param_name, param_value in params.items():
                        data.append({
                            'Component_Index': i, 'Component_IID': comp_iid,
                            'Component_Name': comp_name, 'Component_Type': comp_type,
                            'Parameter_Name': param_name, 'Current_Value': str(param_value),
                        })
                else: # Component has no parameters
                    data.append({
                        'Component_Index': i, 'Component_IID': comp_iid,
                        'Component_Name': comp_name, 'Component_Type': comp_type,
                        'Parameter_Name': 'N/A', 'Current_Value': 'N/A',
                    })
            except AttributeError as e:
                print(f"⚠️ Skipping component {i} due to an attribute error: {e}")
            
            if i % 50 == 0 or i == total:
                print(f"   ...processed {i}/{total} components.")
        return data

    def export_to_excel(self, output_file: str = "pscad_parameters.xlsx"):
        """Exports all component parameters to a formatted Excel file."""
        components = self.canvas.components()
        data = self._get_component_data(components)

        if not data:
            print("No data to export.")
            return

        df = pd.DataFrame(data)
        df['New_Value'] = ''
        df['Notes'] = ''

        # Handle case where file is already open
        final_output_file = self._get_writable_filepath(output_file)
        
        print(f"\nWriting to Excel file: {final_output_file}...")
        with pd.ExcelWriter(final_output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Components', index=False)
            self._format_excel_sheet(writer, 'Components', df)
            self._create_instructions_sheet(writer)

        print("\n✅ Export complete!")
        print(f"   📁 File saved: {final_output_file}")
        print(f"   📊 Total components: {len(df['Component_IID'].unique())}")
        print(f"   📊 Total parameters: {len(df)}")

    def import_from_excel(self, excel_file: str, dry_run: bool = True):
        """Imports and updates parameters from an Excel file."""
        print(f"📖 Reading from: {excel_file}")
        try:
            df = pd.read_excel(excel_file, sheet_name='Components')
        except FileNotFoundError:
            print(f"❌ Error: Excel file not found at '{excel_file}'")
            return

        df_changes = df[df['New_Value'].notna() & (df['New_Value'] != '')].copy()
        df_changes['New_Value'] = df_changes['New_Value'].astype(str)


        if df_changes.empty:
            print("⚠️ No changes found in 'New_Value' column.")
            return

        print(f"📝 Found {len(df_changes)} parameter changes to apply.")
        
        # Create a map of IID to component for quick lookup
        comp_map = {str(comp.iid): comp for comp in self.canvas.components()}
        
        update_count = 0
        error_count = 0

        for comp_iid, group in df_changes.groupby('Component_IID'):
            comp = comp_map.get(comp_iid)
            if not comp:
                print(f"\n❌ Component IID '{comp_iid}': Not found in project. Skipping.")
                error_count += len(group)
                continue

            first_row = group.iloc[0]
            display_name = f"[{first_row['Component_Index']}] {first_row['Component_Type']} ({first_row['Component_Name']})"
            print(f"\n🔧 Processing Component: {display_name} (IID: {comp_iid})")

            params_to_update = {}
            for _, row in group.iterrows():
                param_name = row['Parameter_Name']
                old_val = row['Current_Value']
                new_val = row['New_Value']
                print(f"   • {param_name}: '{old_val}' → '{new_val}'")
                params_to_update[param_name] = new_val

            if not dry_run:
                try:
                    comp.parameters(**params_to_update)
                    print("   ✅ Update successful.")
                    update_count += len(group)
                except Exception as e:
                    print(f"   ❌ Update FAILED: {e}")
                    error_count += len(group)
            else:
                update_count += len(group)
        
        if not dry_run and error_count == 0:
            print("\n💾 Saving project...")
            self.project.save()
            print("✅ Project saved successfully.")
        
        print("\n" + "="*50)
        print("📊 Import Summary:")
        if dry_run:
            print("   Mode: DRY RUN (No changes were actually made)")
        print(f"   - Successful changes planned/applied: {update_count}")
        print(f"   - Errors encountered: {error_count}")
        print("="*50)
        
        if dry_run:
            print(f"\n💡 To apply these changes, run the command again with `dry_run=False`")


    def _get_writable_filepath(self, filepath: str) -> str:
        """Checks if a file is writable, appending a timestamp if not."""
        try:
            with open(filepath, 'a'):
                pass
            return filepath
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base, ext = os.path.splitext(filepath)
            new_filepath = f"{base}_{timestamp}{ext}"
            print(f"⚠️ Original file '{filepath}' is busy. Saving to '{new_filepath}' instead.")
            return new_filepath

    def _format_excel_sheet(self, writer, sheet_name, df):
        """Applies auto-width and freezes panes for an Excel sheet."""
        worksheet = writer.sheets[sheet_name]
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
                series.astype(str).map(len).max(),
                len(str(series.name))
            )) + 2
            worksheet.column_dimensions[chr(65 + idx)].width = min(max_len, 50)
        worksheet.freeze_panes = 'A2'

    def _create_instructions_sheet(self, writer):
        """Creates a helpful 'Instructions' sheet in the Excel file."""
        instructions_df = pd.DataFrame({
            'Step': [1, 2, 3, 4],
            'Instruction': [
                "Go to the 'Components' sheet.",
                "Find the parameter you want to change.",
                "Enter the new desired value in the 'New_Value' column for that row.",
                "Save this Excel file and run the Python import script."
            ]
        })
        instructions_df.to_excel(writer, sheet_name='Instructions', index=False)
        self._format_excel_sheet(writer, 'Instructions', instructions_df)


def check_pscad_running() -> bool:
    """Checks if a PSCAD process is currently running."""
    try:
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq PSCAD*.exe'],
            capture_output=True, text=True, check=True
        )
        return 'PSCAD' in result.stdout
    except (subprocess.CalledProcessError, FileNotFoundError):
        return False # tasklist might not be available or no process found

# ======================== CONFIGURATION & USAGE ========================
if __name__ == "__main__":
    # --- PLEASE CONFIGURE THESE VALUES ---
    # Use absolute path for reliability
    PROJECT_DIRECTORY = r'C:\Users\hqh14\OneDrive\Desktop\08_19_2025_PSCAD_Model_CN_rev1'
    PROJECT_NAME = "main_3LG"
    CANVAS_NAME = "Main"
    EXCEL_FILE = "pscad_parameters_output.xlsx"
    
    print("="*60)
    print("       PSCAD PARAMETER MANAGEMENT TOOL")
    print("="*60)
    print("\nRecommended Workflow:")
    print("1. Run this script to EXPORT parameters to Excel.")
    print("2. Open the generated .xlsx file and fill in the 'New_Value' column.")
    print("3. IMPORTANT: Close the PSCAD application completely.")
    print("4. Run this script again to IMPORT the changes.")
    print("\n" + "-"*60 + "\n")

    # --- CHOOSE ACTION ---
    
    # ACTION 1: EXPORT PARAMETERS TO EXCEL
    # try:
    #     with PscadManager(PROJECT_DIRECTORY, PROJECT_NAME, CANVAS_NAME) as manager:
    #         manager.export_to_excel(EXCEL_FILE)
    # except Exception as e:
    #     print(f"\n❌ An error occurred during export: {e}")

    # ACTION 2: IMPORT PARAMETERS FROM EXCEL
    if check_pscad_running():
        print("⚠️ WARNING: PSCAD application is currently running.")
        print("   For changes to appear correctly in the GUI, please CLOSE PSCAD first.")
        print("   The script will still save changes to the project file, but the GUI won't refresh.")
        if input("   Do you want to continue anyway? (y/n): ").lower() != 'y':
            print("   Import cancelled by user.")
            exit()
            
    try:
        # Step 2.1: Dry Run (highly recommended)
        print("\n--- Starting DRY RUN (Previewing changes) ---")
        with PscadManager(PROJECT_DIRECTORY, PROJECT_NAME, CANVAS_NAME) as manager:
            manager.import_from_excel(EXCEL_FILE, dry_run=True)
            
        # Step 2.2: Actual Import
        print("\n--- Preparing for ACTUAL IMPORT ---")
        if input("Proceed with applying these changes to PSCAD? (y/n): ").lower() == 'y':
            with PscadManager(PROJECT_DIRECTORY, PROJECT_NAME, CANVAS_NAME) as manager:
                manager.import_from_excel(EXCEL_FILE, dry_run=False)
        else:
            print("Actual import cancelled by user.")

    except Exception as e:
        print(f"\n❌ An error occurred during import: {e}")