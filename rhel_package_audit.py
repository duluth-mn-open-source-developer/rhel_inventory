"""
================================================================================
                       CORPORATE INFRASTRUCTURE AUDIT SERVICES
================================================================================
System Identifier : SYSTEM-PACKAGE-AUDIT-ENGINE
Component Name    : rhel_package_audit.py
Security Domain   : Security Administration & Inventory Assurance
Description       : Programmatically interfaces with the RHEL native DNF/RPM 
                    database and local Python pip ecosystems using parameters
                    provided by an external JSON configuration file.
Target OS         : Red Hat Enterprise Linux 8.x / 9.x
================================================================================

================================================================================
                               CHANGE CONTROL LOG
================================================================================
Date        | Version | Author             | Description of Change
------------|---------|--------------------|------------------------------------
2026-03-15  | 1.0.0   | Core Systems Team  | Initial baseline build using DNF 
            |         |                    | standard library subprocess layers.
2026-04-10  | 1.1.0   | Infrastructure Dev | Integrated openpyxl formatting suite 
            |         |                    | and zebra-striping layout updates.
2026-06-02  | 2.0.0   | Lead SecOps Arch   | Shifted from `dnf info` to raw 
            |         |                    | `dnf repoquery --queryformat` tags.
2026-06-02  | 2.1.0   | Systems Integrator | Appended local system context flags 
            |         |                    | (--cacheonly) to fix missing/empty 
            |         |                    | %{installtime} metadata strings.
2026-06-02  | 2.2.0   | DevOps Engineer    | Added PIP package parsing pipeline.
2026-06-02  | 2.3.0   | Python Architect   | Upgraded PIP parser to use standard
            |         |                    | library importlib.metadata.
2026-06-02  | 2.3.1   | Systems Engineer   | Fixed |%| typo and added *extra logic.
2026-06-02  | 3.0.0   | Core DevOps Eng    | Externalized script properties into 
            |         |                    | an isolated audit_config.json block.
================================================================================

================================================================================
                             INSTRUCTIONS FOR USAGE
================================================================================
1. PREREQUISITES:
   - Ensure 'audit_config.json' is located in the execution directory.
   - Install openpyxl library: `pip3 install openpyxl` or `dnf install python3-openpyxl`
   - Execution requires root or sudo privileges to safely scan local DNF caching layers:
     `sudo python3 rhel_package_audit.py`
================================================================================
"""

import os
import json
import datetime
import subprocess
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import importlib.metadata as metadata
except ImportError:
    import importlib_metadata as metadata

CONFIG_PATH = "audit_config.json"

def load_system_config():
    """Loads and returns configurations from the centralized workspace json map."""
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(f"[CRITICAL] Configuration profile mapping '{CONFIG_PATH}' missing from root context.")
    with open(CONFIG_PATH, "r") as f:
        return json.load(f)

def build_rhel_audit_report():
    """Queries package systems via parameterized config schemas to compile an excel audit report."""
    try:
        cfg = load_system_config()
    except Exception as e:
        print(f"[FATAL] System configuration parser aborted: {e}")
        return False

    os_cfg = cfg["os_audit"]
    output_filename = os_cfg.get("output_filename", "RHEL_Package_Audit_Report.xlsx")

    # --------------------------------------------------------------------------
    # 1. DATA PIPELINE ACQUISITION: OS PACKAGES
    # --------------------------------------------------------------------------
    dnf_stdout = ""
    dnf_cmd = os_cfg["dnf_command"] + [f"--queryformat={os_cfg['dnf_format']}"]
    
    try:
        result_dnf = subprocess.run(dnf_cmd, capture_output=True, text=True, check=True)
        dnf_stdout = result_dnf.stdout
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("[WARNING] DNF local cache incomplete. Executing native RPM fallback pipeline...")
        rpm_cmd = os_cfg["rpm_fallback_command"] + [f"--queryformat={os_cfg['rpm_format']}"]
        try:
            result_rpm = subprocess.run(rpm_cmd, capture_output=True, text=True, check=True)
            dnf_stdout = result_rpm.stdout
        except Exception as e:
            print(f"[CRITICAL] OS pipeline failure: {e}")
            return False

    # --------------------------------------------------------------------------
    # 2. DATA PIPELINE ACQUISITION: Python PIP MODULES
    # --------------------------------------------------------------------------
    pip_data_list = []
    try:
        dists = sorted(metadata.distributions(), key=lambda d: d.metadata['Name'].lower())
        for dist in dists:
            p_name = dist.metadata['Name']
            p_ver = dist.metadata['Version']
            p_desc = dist.metadata.get('Summary', 'No description provided by author.')
            
            p_date = "Unknown"
            if hasattr(dist, 'locate_file'):
                meta_dir_path = dist.locate_file('')
                if os.path.exists(meta_dir_path):
                    stat_info = os.stat(meta_dir_path)
                    p_date = datetime.datetime.fromtimestamp(stat_info.st_mtime).strftime('%Y-%m-%d %H:%M:%S')

            pip_data_list.append([p_name, p_ver, p_date, p_desc])
    except Exception as e:
        print(f"[WARNING] PIP context mapping encountered issues: {e}")

    # --------------------------------------------------------------------------
    # 3. WORKBOOK STYLING SCHEMA INITIALIZATION
    # --------------------------------------------------------------------------
    wb = openpyxl.Workbook()
    
    ws_dash = wb.active
    ws_dash.title = "Dashboard Overview"
    ws_dash.views.sheetView[0].showGridLines = True
    
    ws_dnf = wb.create_sheet(title="Installed Packages Data")
    ws_dnf.views.sheetView[0].showGridLines = True
    
    ws_pip = wb.create_sheet(title="PIP Packages Data")
    ws_pip.views.sheetView[0].showGridLines = True

    DARK_BLUE, WHITE = "1B365D", "FFFFFF"
    font_title = Font(name="Segoe UI", size=18, bold=True, color=DARK_BLUE)
    font_subtitle = Font(name="Segoe UI", size=11, italic=True, color="555555")
    font_header = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    font_data = Font(name="Segoe UI", size=10)
    font_bold_data = Font(name="Segoe UI", size=10, bold=True)
    
    fill_header = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
    fill_zebra = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
    fill_accent = PatternFill(start_color="E6F0FA", end_color="E6F0FA", fill_type="solid")
    
    border_thin = Border(left=Side(style="thin", color="E0E0E0"), right=Side(style="thin", color="E0E0E0"),
                         top=Side(style="thin", color="E0E0E0"), bottom=Side(style="thin", color="E0E0E0"))
    border_total = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))

    # --------------------------------------------------------------------------
    # 4. POPULATE TAB 2: SYSTEM DNF DATA
    # --------------------------------------------------------------------------
    dnf_headers = ["Package Name", "Version", "Release", "Architecture", "Size (MB)", "Repository", "Installation Date", "Vendor", "Source RPM"]
    for col_num, h_text in enumerate(dnf_headers, 1):
        c = ws_dnf.cell(row=1, column=col_num, value=h_text)
        c.font = font_header; c.fill = fill_header; c.border = border_thin
        c.alignment = Alignment(horizontal="center" if col_num in [4,6,7] else ("right" if col_num == 5 else "left"), vertical="center")

    dnf_idx = 2
    for line in dnf_stdout.strip().split("\n"):
        if not line or "|" not in line:
            continue
        parts = line.split("|")
        if len(parts) < 9:
            continue
        
        name, version, release, arch, raw_size, repo, raw_time, vendor, source, *extra = parts
        size_mb = round(int(raw_size) / (1024 * 1024), 2) if raw_size.isdigit() else 0.0
        
        try:
            if not raw_time or raw_time == "0" or "none" in raw_time.lower():
                inst_date = "System Image Baseline"
            else:
                inst_date = datetime.datetime.fromtimestamp(int(raw_time)).strftime('%Y-%m-%d %H:%M:%S')
        except (ValueError, TypeError):
            inst_date = "System Image Baseline"

        row_values = [name, version, release, arch, size_mb, repo, inst_date, vendor, source]
        
        for col_idx, val in enumerate(row_values, 1):
            cell = ws_dnf.cell(row=dnf_idx, column=col_idx, value=val)
            cell.font = font_data; cell.border = border_thin
            if dnf_idx % 2 == 0: cell.fill = fill_zebra
                
            if col_idx == 5:
                cell.number_format = '#,##0.0'
                cell.alignment = Alignment(horizontal="right")
            elif col_idx in [4, 6, 7]:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="left")
        dnf_idx += 1

    ws_dnf.cell(row=dnf_idx, column=1, value="Total Installed Packages Size").font = font_bold_data
    ws_dnf.cell(row=dnf_idx, column=1).border = border_total
    total_size_cell = ws_dnf.cell(row=dnf_idx, column=5, value=f"=SUM(E2:E{dnf_idx-1})")
    total_size_cell.font = font_bold_data; total_size_cell.number_format = '#,##0.0 "MB"'; total_size_cell.border = border_total
    
    for col in range(2, 10):
        if col != 5: ws_dnf.cell(row=dnf_idx, column=col).border = border_total

    ws_dnf.auto_filter.ref = f"A1:I{dnf_idx-1}"
    ws_dnf.freeze_panes = "A2"

    # --------------------------------------------------------------------------
    # 5. POPULATE TAB 3: ENHANCED PYTHON PIP DATA
    # --------------------------------------------------------------------------
    pip_headers = ["Extension Module Name", "Active Core Version", "Inferred Installation Date", "Package Description / Summary"]
    for col_num, h_text in enumerate(pip_headers, 1):
        c = ws_pip.cell(row=1, column=col_num, value=h_text)
        c.font = font_header; c.fill = fill_header; c.border = border_thin
        c.alignment = Alignment(horizontal="center" if col_num in [2, 3] else "left", vertical="center")

    pip_idx = 2
    if pip_data_list:
        for row_values in pip_data_list:
            for col_idx, val in enumerate(row_values, 1):
                cell = ws_pip.cell(row=pip_idx, column=col_idx, value=val)
                cell.font = font_data; cell.border = border_thin
                if pip_idx % 2 == 0: cell.fill = fill_zebra
                if col_idx in [2, 3]: cell.alignment = Alignment(horizontal="center")
                else: cell.alignment = Alignment(horizontal="left")
            pip_idx += 1
    else:
        ws_pip.cell(row=2, column=1, value="No PIP data found or dependencies unindexed.").font = font_data

    ws_pip.auto_filter.ref = f"A1:D{max(2, pip_idx-1)}"
    ws_pip.freeze_panes = "A2"

    # --------------------------------------------------------------------------
    # 6. GENERATE EXECUTIVE KPI DASHBOARD
    # --------------------------------------------------------------------------
    ws_dash.cell(row=2, column=2, value="RHEL Package Infrastructure Audit Report").font = font_title
    ws_dash.cell(row=3, column=2, value="Automated system landscape generation using Python & DNF").font = font_subtitle
    
    ws_dash.merge_cells("B5:C5"); ws_dash.cell(row=5, column=2, value="OS Audited Packages").font = font_header; ws_dash.cell(row=5, column=2).fill = fill_header
    ws_dash.merge_cells("B6:C7"); c_cell = ws_dash.cell(row=6, column=2, value=f"=COUNTA('Installed Packages Data'!A2:A{dnf_idx-1})")
    c_cell.font = Font(name="Segoe UI", size=20, bold=True, color=DARK_BLUE); c_cell.alignment = Alignment(horizontal="center", vertical="center"); c_cell.fill = fill_accent
    
    ws_dash.merge_cells("E5:F5"); ws_dash.cell(row=5, column=5, value="Total Disk Footprint").font = font_header; ws_dash.cell(row=5, column=5).fill = fill_header
    ws_dash.merge_cells("E6:F7"); f_cell = ws_dash.cell(row=6, column=5, value=f"='Installed Packages Data'!E{dnf_idx}")
    f_cell.font = Font(name="Segoe UI", size=20, bold=True, color=DARK_BLUE); f_cell.alignment = Alignment(horizontal="center", vertical="center"); f_cell.fill = fill_accent

    ws_dash.merge_cells("H5:I5"); ws_dash.cell(row=5, column=8, value="PIP Audited Packages").font = font_header; ws_dash.cell(row=5, column=8).fill = fill_header
    ws_dash.merge_cells("H6:I7"); p_cell = ws_dash.cell(row=6, column=8, value=f"=COUNTA('PIP Packages Data'!A2:A{max(2, pip_idx-1)})")
    p_cell.font = Font(name="Segoe UI", size=20, bold=True, color=DARK_BLUE); p_cell.alignment = Alignment(horizontal="center", vertical="center"); p_cell.fill = fill_accent

    for r in range(5, 8):
        for c in [2, 3, 5, 6, 8, 9]: ws_dash.cell(row=r, column=c).border = border_thin

    # --------------------------------------------------------------------------
    # 7. AUTO-FIT CELL DIMENSIONS
    # --------------------------------------------------------------------------
    for sheet in wb.worksheets:
        for col in sheet.columns:
            if sheet.title == "Dashboard Overview" and col[0].column > 9: continue
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or '')) for cell in col)
            if sheet.title == "PIP Packages Data" and col_letter == 'D':
                sheet.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 60)
            else:
                sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output_filename)
    print(f"[SUCCESS] Multi-Tab Corporate Audit Report Generated: {output_filename}")
    return True

if __name__ == "__main__":
    build_rhel_audit_report()
