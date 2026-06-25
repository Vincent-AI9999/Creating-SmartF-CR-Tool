import streamlit as st
import pandas as pd
import openpyxl
import os
import glob
import re
import copy
from datetime import datetime

# Set page config
st.set_page_config(
    page_title="Mobifone CR Generator Tool",
    page_icon="📶",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (Rich Aesthetics)
st.markdown("""
<style>
    .reportview-container {
        background: #0e1117;
    }
    .main-title {
        font-size: 38px;
        font-weight: 800;
        color: #1F4E79;
        text-align: center;
        margin-bottom: 20px;
        font-family: 'Outfit', 'Inter', sans-serif;
    }
    .subtitle {
        font-size: 16px;
        color: #8a8d93;
        text-align: center;
        margin-bottom: 40px;
    }
    .section-header {
        font-size: 22px;
        font-weight: 600;
        color: #1F4E79;
        border-bottom: 2px solid #1F4E79;
        padding-bottom: 5px;
        margin-top: 20px;
        margin-bottom: 15px;
    }
    .success-box {
        padding: 20px;
        border-radius: 10px;
        background-color: #d4edda;
        color: #155724;
        border: 1px solid #c3e6cb;
        margin-top: 15px;
    }
    .stButton>button {
        background-color: #1F4E79;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px 25px;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #153552;
        box-shadow: 0px 4px 10px rgba(31, 78, 121, 0.4);
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# CONSTANTS & CONFIGURATIONS
# -----------------------------------------------------------------------------
BASE_WORKSPACE_PATH = r"F:\OneDrive - Mobifone\F_WORKING\Python_Coding\Create CR"
TEMPLATE_DIR = os.path.join(BASE_WORKSPACE_PATH, "CR mẫu")
SOURCE_CELL_DIR = r"F:\OneDrive - Mobifone\F_WORKING\Python_Coding\MTCL 2026\Source_cell"
DATABASE_DIR = r"F:\OneDrive - Mobifone\F_WORKING\Python_Coding\Database trạm"

# -----------------------------------------------------------------------------
# HELPER FUNCTIONS
# -----------------------------------------------------------------------------
def get_latest_cell_analysis():
    """Finds the latest Cell_Analysis_*.xlsx file > 10KB in Source_cell."""
    if not os.path.exists(SOURCE_CELL_DIR):
        return None
    files = glob.glob(os.path.join(SOURCE_CELL_DIR, "Cell_Analysis_*.xlsx"))
    valid_files = [f for f in files if os.path.getsize(f) > 10240] # > 10KB
    if not valid_files:
        return None
    return max(valid_files, key=os.path.getmtime)

@st.cache_data
def load_bad_cells_from_analysis(filepath):
    """Loads poor KPI cells from the latest Cell Analysis file."""
    bad_cells = {'3G': set(), '4G': set(), '5G': set()}
    if not filepath or not os.path.exists(filepath):
        return bad_cells
    
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        # 3G Top 10
        if '3G_Top10' in wb.sheetnames:
            sheet = wb['3G_Top10']
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) > 1 and row[1]: # CELLBNAME NEW
                    bad_cells['3G'].add(str(row[1]).strip())
        
        # 4G Top 10
        if '4G_Top10' in wb.sheetnames:
            sheet = wb['4G_Top10']
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) > 1 and row[1]: # eUTraCell New
                    bad_cells['4G'].add(str(row[1]).strip())
        
        # PRB Top 10 (Also 4G cells)
        if 'PRB_Top10' in wb.sheetnames:
            sheet = wb['PRB_Top10']
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) > 1 and row[1]: # EUTRACELL
                    bad_cells['4G'].add(str(row[1]).strip())
        
        # 5G Top 10
        if '5G_Top10' in wb.sheetnames:
            sheet = wb['5G_Top10']
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) > 1 and row[1]: # NRCell Name
                    bad_cells['5G'].add(str(row[1]).strip())
                    
    except Exception as e:
        st.warning(f"Không thể đọc file Cell Analysis gợi ý: {e}")
        
    return bad_cells

def extract_site_name(cell_name, tech):
    """Extracts site name from cell name based on technology."""
    s = str(cell_name).strip()
    if not s or s.lower() == 'nan':
        return ''
    if tech == '3G':
        if len(s) == 13:
            return s[:-5]  # e.g., LDGBLO07BM3GB -> LDGBLO07
        elif len(s) == 7:
            return s[:6]
        elif len(s) > 7:
            return s[:-4]
        return s[:6]
    else: # 4G & 5G
        if len(s) == 13:
            return s[:-5]  # e.g., DNIDXO09CM4CA -> DNIDXO09
        return s[:-4]

@st.cache_data
def load_cell_database(vendor, tech):
    """Loads or constructs cell name to distName and siteName mappings from consolidated Excel DBs."""
    mapping = {}
    try:
        # Load from consolidated database files
        filename = f"{tech}_{'NSN' if vendor == 'Nokia' else 'ERA'}.xlsx"
        db_path = os.path.join(DATABASE_DIR, filename)
        
        if os.path.exists(db_path):
            wb = openpyxl.load_workbook(db_path, read_only=True)
            if 'usual_report' in wb.sheetnames:
                df = pd.read_excel(db_path, sheet_name='usual_report')
                df.columns = [str(c).strip() for c in df.columns]
                
                # Nokia 4G mapping
                if vendor == 'Nokia' and tech == '4G':
                    # Has MRBTS, LNCEL, cellName, LNBTS
                    for _, row in df.iterrows():
                        c_name = str(row.get('cellName', '')).strip()
                        mrbts = str(row.get('MRBTS', '')).strip()
                        lnbts = str(row.get('LNBTS', '')).strip()
                        lncel = str(row.get('LNCEL', '')).strip()
                        if c_name and mrbts and lncel:
                            # Construct distName: PLMN-PLMN/MRBTS-XXXXXX/LNBTS-XXXXXX/LNCEL-XX
                            dist = f"PLMN-PLMN/{mrbts}/{lnbts}/{lncel}"
                            site = extract_site_name(c_name, '4G')
                            mapping[c_name] = {
                                'distName': dist,
                                'siteName': site,
                                'cellName': c_name,
                                'tech': '4G',
                                'vendor': 'Nokia'
                            }
                # Nokia 3G mapping
                elif vendor == 'Nokia' and tech == '3G':
                    # Has distName, cellName
                    for _, row in df.iterrows():
                        c_name = str(row.get('cellName', '')).strip()
                        dist = str(row.get('distName', '')).strip()
                        if c_name and dist:
                            site = extract_site_name(c_name, '3G')
                            mapping[c_name] = {
                                'distName': dist,
                                'siteName': site,
                                'cellName': c_name,
                                'tech': '3G',
                                'vendor': 'Nokia'
                            }
                # Ericsson 3G mapping
                elif vendor == 'Ericsson' and tech == '3G':
                    # Has ManagedElement_id, CellName
                    for _, row in df.iterrows():
                        c_name = str(row.get('CellName', '')).strip()
                        me_id = str(row.get('ManagedElement_id', '')).strip()
                        if c_name and me_id:
                            # Construct DN
                            dist = f"SubNetwork=MobiFone_DongNai,MeContext={me_id},ManagedElement={me_id},UtranNetwork=1,UtranCell={c_name}"
                            mapping[c_name] = {
                                'distName': dist,
                                'siteName': me_id,
                                'cellName': c_name,
                                'tech': '3G',
                                'vendor': 'Ericsson'
                            }
                # Ericsson 4G mapping
                elif vendor == 'Ericsson' and tech == '4G':
                    # Has ManagedElement_id, CellName
                    for _, row in df.iterrows():
                        c_name = str(row.get('CellName', '')).strip()
                        me_id = str(row.get('ManagedElement_id', '')).strip()
                        if c_name and me_id:
                            dist = f"SubNetwork=MobiFone_DongNai,MeContext={me_id},ManagedElement={me_id},ENodeBFunction=1,EUtranCellFDD={c_name}"
                            mapping[c_name] = {
                                'distName': dist,
                                'siteName': me_id,
                                'cellName': c_name,
                                'tech': '4G',
                                'vendor': 'Ericsson'
                            }
                # Ericsson 5G mapping
                elif vendor == 'Ericsson' and tech == '5G':
                    # Has ManagedElement_id, CellName
                    for _, row in df.iterrows():
                        c_name = str(row.get('CellName', '')).strip()
                        me_id = str(row.get('ManagedElement_id', '')).strip()
                        if c_name and me_id:
                            dist = f"SubNetwork=MobiFone_DongNai,MeContext={me_id},ManagedElement={me_id},GNBCUCPFunction=1,NRCellCU={c_name}"
                            mapping[c_name] = {
                                'distName': dist,
                                'siteName': me_id,
                                'cellName': c_name,
                                'tech': '5G',
                                'vendor': 'Ericsson'
                            }
            
            # Special case for Nokia 5G (where 5G_NSN.xlsx only has hardware_report)
            elif vendor == 'Nokia' and tech == '5G':
                # We can query Master_Cell_File_*.xlsx sheet '5G Cells'!
                master_files = glob.glob(os.path.join(DATABASE_DIR, "Master_Cell_File_*.xlsx"))
                if master_files:
                    latest_master = max(master_files, key=os.path.getmtime)
                    m_wb = openpyxl.load_workbook(latest_master, read_only=True)
                    if '5G Cells' in m_wb.sheetnames:
                        df_5g = pd.read_excel(latest_master, sheet_name='5G Cells')
                        df_5g.columns = [str(c).strip() for c in df_5g.columns]
                        for _, row in df_5g.iterrows():
                            c_name = str(row.get('Cell Name', '')).strip()
                            vdr = str(row.get('Vendor', '')).strip()
                            if c_name and vdr == 'NSN':
                                cell_id = int(row.get('Cell ID', 0))
                                site = extract_site_name(c_name, '5G')
                                # Construct Nokia 5G distName: PLMN-PLMN/MRBTS-XXXX/NRBTS-XXXX/NRCELL-XX
                                # Let's fetch gNodeB ID from cell info
                                gnb = str(row.get('gNodeB ID', '')).strip()
                                # e.g. MRBTS-5856010
                                if 'MRBTS-' in gnb:
                                    mrbts = gnb
                                else:
                                    mrbts = f"MRBTS-{gnb}"
                                dist = f"PLMN-PLMN/{mrbts}/NRBTS-{gnb.replace('MRBTS-', '')}/NRCELL-{cell_id}"
                                mapping[c_name] = {
                                    'distName': dist,
                                    'siteName': site,
                                    'cellName': c_name,
                                    'tech': '5G',
                                    'vendor': 'Nokia'
                                }
        else:
            # Fallback to scanning raw cell exports if consolidated file is missing
            st.warning(f"Không tìm thấy database tổng hợp {filename}. Sẽ quét file raw trong dump...")
            
    except Exception as e:
        st.error(f"Lỗi khi xây dựng bản đồ Cell Database: {e}")
        
    return mapping

# Cache for loaded raw dumps to prevent reading large text files repeatedly
@st.cache_resource
def load_dump_file(filepath):
    """Loads a raw tab-separated text dump or CSV file and caches it."""
    try:
        if filepath.endswith('.csv'):
            return pd.read_csv(filepath, low_memory=False)
        else:
            return pd.read_csv(filepath, sep='\t', low_memory=False)
    except Exception as e:
        st.error(f"Lỗi khi đọc file dump {os.path.basename(filepath)}: {e}")
        return pd.DataFrame()

def query_nokia_parameter_value(tech, sheet_name, cell_dist_name, param_name, key_cols_values):
    """Queries the parameter value from Nokia raw export dump files."""
    tech_folder = '3G' if tech == '3G' else ('4G' if tech == '4G' else '5G')
    sub = 'DUMP' if tech == '3G' else 'Dump'
    dump_dir = os.path.join(DATABASE_DIR, tech_folder, sub, 'NSN')
    
    if not os.path.exists(dump_dir):
        return None
    
    # Locate files like Export_{sheet_name}__*.txt
    pattern = os.path.join(dump_dir, f"Export_{sheet_name}__*.txt")
    files = glob.glob(pattern)
    if not files:
        # Try case insensitive search
        all_files = os.listdir(dump_dir)
        files = [os.path.join(dump_dir, f) for f in all_files if f.lower().startswith(f"export_{sheet_name.lower()}__")]
        
    if not files:
        return None
        
    latest_file = max(files, key=os.path.getmtime)
    df = load_dump_file(latest_file)
    if df.empty:
        return None
        
    # Map MO to distName if needed
    if 'MO' in df.columns:
        df = df.rename(columns={'MO': 'distName'})
        
    if 'distName' not in df.columns:
        return None
        
    # Filter rows by cell_dist_name
    # Associated if MO == cell_dist_name or MO starts with cell_dist_name + '/'
    df_filtered = df[df['distName'].astype(str).str.startswith(cell_dist_name)]
    if df_filtered.empty:
        return None
        
    # If other key columns exist (like utraCarrierFreq)
    for col, val in key_cols_values.items():
        if col != 'distName' and col != 'DN' and col in df_filtered.columns:
            # Try to match the key value (e.g. utraCarrierFreq = 10587)
            # Match using numeric or string depending on column type
            df_filtered = df_filtered[df_filtered[col].astype(str) == str(val)]
            
    if not df_filtered.empty and param_name in df_filtered.columns:
        return df_filtered.iloc[0][param_name]
        
    return None

def query_ericsson_parameter_value(tech, sheet_name, cell_name, cell_dist_name, param_name, key_cols_values):
    """Queries the parameter value from Ericsson raw dumps or consolidated Excel databases."""
    tech_folder = '3G' if tech == '3G' else ('4G' if tech == '4G' else '5G')
    sub = 'DUMP' if tech == '3G' else 'Dump'
    dump_dir = os.path.join(DATABASE_DIR, tech_folder, sub, 'ERA')
    
    # Try raw CSV dumps first
    if os.path.exists(dump_dir):
        # Look for vsData{sheet_name}.csv or {sheet_name}.csv
        csv_files = []
        for name in [sheet_name, f"vsData{sheet_name}"]:
            csv_files.extend(glob.glob(os.path.join(dump_dir, f"{name}.csv")))
            csv_files.extend(glob.glob(os.path.join(dump_dir, f"{name.lower()}.csv")))
            
        if csv_files:
            latest_file = max(csv_files, key=os.path.getmtime)
            df = load_dump_file(latest_file)
            if not df.empty:
                # Ericsson CSVs usually have DN as first column or specific IDs
                # Let's search by cell name or distName
                dn_col = None
                for c in df.columns:
                    if str(c).lower() in ['dn', 'distname', 'managedelement', 'utrancell', 'eutrancellfdd', 'nrcellcu']:
                        dn_col = c
                        break
                if not dn_col:
                    dn_col = df.columns[0]
                    
                df_filtered = df[df[dn_col].astype(str).str.contains(cell_name, case=False) | df[dn_col].astype(str).str.contains(cell_dist_name, case=False)]
                
                # Filter by other keys if any
                for col, val in key_cols_values.items():
                    if col not in ['distName', 'DN', 'cellName', 'CellName'] and col in df_filtered.columns:
                        df_filtered = df_filtered[df_filtered[col].astype(str) == str(val)]
                        
                if not df_filtered.empty and param_name in df_filtered.columns:
                    return df_filtered.iloc[0][param_name]
                    
    # Fallback to consolidated Excel DB
    db_file = os.path.join(DATABASE_DIR, f"{tech}_ERA.xlsx")
    if os.path.exists(db_file):
        try:
            wb = openpyxl.load_workbook(db_file, read_only=True)
            for sheet in ['usual_report', 'cell_report']:
                if sheet in wb.sheetnames:
                    df = pd.read_excel(db_file, sheet_name=sheet)
                    df.columns = [str(c).strip() for c in df.columns]
                    
                    # Find by CellName
                    cell_col = 'CellName' if 'CellName' in df.columns else ('cellName' if 'cellName' in df.columns else None)
                    if cell_col:
                        df_filtered = df[df[cell_col].astype(str) == cell_name]
                        if not df_filtered.empty and param_name in df_filtered.columns:
                            return df_filtered.iloc[0][param_name]
        except Exception as e:
            pass
            
    return None

def copy_cell_style(src_cell, dst_cell):
    """Copies all formatting styles from a source cell to a destination cell."""
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format

# -----------------------------------------------------------------------------
# APPLICATION MAIN INTERFACE
# -----------------------------------------------------------------------------
st.markdown("<div class='main-title'>MOBIFONE CR AUTOMATION TOOL</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Tối ưu hóa quy trình cấu hình tham số mạng vô tuyến Nokia & Ericsson</div>", unsafe_allow_html=True)

# Layout: 2 Columns
col_left, col_right = st.columns([1, 2])

with col_left:
    st.markdown("<div class='section-header'>1. Tham số đầu vào</div>", unsafe_allow_html=True)
    
    vendor = st.selectbox("Chọn Nhà cung cấp (Vendor):", ["Nokia", "Ericsson"])
    tech = st.selectbox("Chọn Công nghệ (RAN):", ["4G", "3G", "5G"])
    
    # Input cell list or site list
    input_method = st.radio(
        "Phương thức chọn Cell/Trạm:",
        ["Gợi ý từ KPI kém SmartF", "Nhập danh sách Cell thủ công", "Nhập danh sách Trạm thủ công"]
    )
    
    selected_cell_names = []
    
    # 1. KPI kém SmartF
    if input_method == "Gợi ý từ KPI kém SmartF":
        latest_analysis = get_latest_cell_analysis()
        if latest_analysis:
            st.info(f"Đọc file KPI SmartF: **{os.path.basename(latest_analysis)}**")
            bad_cells_dict = load_bad_cells_from_analysis(latest_analysis)
            tech_bad_cells = sorted(list(bad_cells_dict.get(tech, [])))
            
            if tech_bad_cells:
                selected_cell_names = st.multiselect(
                    f"Chọn Cell lỗi {tech} ({len(tech_bad_cells)} cells gợi ý):",
                    options=tech_bad_cells,
                    default=tech_bad_cells[:5] # Default pre-select first 5
                )
            else:
                st.warning(f"Không tìm thấy Cell lỗi {tech} trong file phân tích.")
        else:
            st.warning("Không tìm thấy file Cell_Analysis_*.xlsx hợp lệ trong Source_cell.")
            
    # 2. Nhập Cell thủ công
    elif input_method == "Nhập danh sách Cell thủ công":
        cell_input = st.text_area(
            "Nhập danh sách Cell (mỗi dòng 1 cell hoặc cách nhau bởi dấu phẩy):",
            placeholder="Ví dụ: DNIDXO09CM4CA, DNIDXO09CM4CB, TGMTP1C4BA"
        )
        if cell_input:
            selected_cell_names = [c.strip() for c in re.split(r'[\n,]+', cell_input) if c.strip()]
            
    # 3. Nhập Trạm thủ công
    else:
        site_input = st.text_area(
            "Nhập danh sách Mã Trạm (mỗi dòng 1 trạm hoặc cách nhau bởi dấu phẩy):",
            placeholder="Ví dụ: DNIDXO09, TGMTP1"
        )
        if site_input:
            input_sites = [s.strip() for s in re.split(r'[\n,]+', site_input) if s.strip()]
            selected_cell_names = []
            
    # CR mẫu template selection
    st.markdown("<div class='section-header'>2. Chọn CR mẫu (Template)</div>", unsafe_allow_html=True)
    if os.path.exists(TEMPLATE_DIR):
        templates = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith('.xlsx')]
        if templates:
            recommended = []
            for t in templates:
                v_match = (vendor == 'Nokia' and any(k in t.lower() for k in ['nsn', 'nokia', 'nok'])) or \
                          (vendor == 'Ericsson' and any(k in t.lower() for k in ['era', 'ericsson', 'eri']))
                t_match = any(k in t.lower() for k in [tech.lower()])
                if v_match and t_match:
                    recommended.append(t)
            
            sorted_templates = recommended + [t for t in templates if t not in recommended]
            
            def format_template_option(option):
                if option in recommended:
                    return f"🌟 (Khuyên dùng) {option}"
                return option
                
            selected_template = st.selectbox(
                "Chọn file CR mẫu:",
                options=sorted_templates,
                format_func=format_template_option
            )
        else:
            st.error("Không tìm thấy file Excel mẫu nào trong thư mục CR mẫu.")
            selected_template = None
    else:
        st.error(f"Thư mục {TEMPLATE_DIR} không tồn tại.")
        selected_template = None

# Right Column - Data preview and generation
with col_right:
    st.markdown("<div class='section-header'>3. Kiểm tra thông tin & Tạo CR</div>", unsafe_allow_html=True)
    
    with st.spinner("Đang kết nối cơ sở dữ liệu trạm..."):
        cell_db = load_cell_database(vendor, tech)
    
    target_cells = []
    
    if input_method == "Nhập danh sách Trạm thủ công" and 'input_sites' in locals() and input_sites:
        for c_name, info in cell_db.items():
            if info['siteName'] in input_sites:
                target_cells.append(info)
    else:
        for c_name in selected_cell_names:
            if c_name in cell_db:
                target_cells.append(cell_db[c_name])
            else:
                st.warning(f"Cell **{c_name}** không tìm thấy trong Database. Sẽ bỏ qua hoặc điền trống.")
                
    if target_cells:
        st.write(f"Đã tìm thấy **{len(target_cells)}** cells phù hợp trong cơ sở dữ liệu:")
        preview_df = pd.DataFrame(target_cells)
        st.dataframe(preview_df[['cellName', 'siteName', 'distName', 'vendor', 'tech']], use_container_width=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_filename = f"CR_{vendor}_{tech}_{timestamp}.xlsx"
        
        if st.button("Bắt đầu tạo Change Request (CR)"):
            with st.spinner("Đang đồng bộ tham số và tạo file Excel..."):
                try:
                    template_path = os.path.join(TEMPLATE_DIR, selected_template)
                    wb = openpyxl.load_workbook(template_path)
                    
                    # 1. Re-populate 'Trạm ảnh hưởng' sheet
                    if 'Trạm ảnh hưởng' in wb.sheetnames:
                        sheet = wb['Trạm ảnh hưởng']
                        max_r = sheet.max_row
                        if max_r >= 3:
                            sheet.delete_rows(3, max_r - 2)
                            
                        unique_sites = sorted(list(set([c['siteName'] for c in target_cells])))
                        
                        for idx, site in enumerate(unique_sites):
                            row_idx = 3 + idx
                            sheet.cell(row=row_idx, column=1, value=idx + 1)
                            sheet.cell(row=row_idx, column=2, value=site)
                            sheet.cell(row=row_idx, column=3, value=tech)
                            
                            for col in range(1, 4):
                                c = sheet.cell(row=row_idx, column=col)
                                c.font = openpyxl.styles.Font(name='Segoe UI', size=11)
                                c.alignment = openpyxl.styles.Alignment(horizontal='center' if col != 2 else 'left')
                                c.border = openpyxl.styles.Border(
                                    left=openpyxl.styles.Side(style='thin', color='D3D3D3'),
                                    right=openpyxl.styles.Side(style='thin', color='D3D3D3'),
                                    top=openpyxl.styles.Side(style='thin', color='D3D3D3'),
                                    bottom=openpyxl.styles.Side(style='thin', color='D3D3D3')
                                )
                                
                    # 2. Re-populate 'LNCELL' sheet
                    if 'LNCELL' in wb.sheetnames:
                        sheet = wb['LNCELL']
                        max_r = sheet.max_row
                        if max_r >= 2:
                            sheet.delete_rows(2, max_r - 1)
                            
                        for idx, cell in enumerate(target_cells):
                            row_idx = 2 + idx
                            sheet.cell(row=row_idx, column=1, value=cell['distName'])
                            sheet.cell(row=row_idx, column=2, value=cell['cellName'])
                            sheet.cell(row=row_idx, column=3, value=cell['siteName'])
                            
                            for col in range(1, 4):
                                c = sheet.cell(row=row_idx, column=col)
                                c.font = openpyxl.styles.Font(name='Segoe UI', size=10)
                                c.border = openpyxl.styles.Border(
                                    left=openpyxl.styles.Side(style='thin', color='D3D3D3'),
                                    right=openpyxl.styles.Side(style='thin', color='D3D3D3')
                                )
                                
                    # 3. Process Parameter Sheets
                    for sheetname in wb.sheetnames:
                        if sheetname in ['1. Thông tin chung', '2. Chi tiết tác động', 'Trạm ảnh hưởng', 'LNCELL']:
                            continue
                            
                        sheet = wb[sheetname]
                        max_r = sheet.max_row
                        if max_r < 2:
                            continue
                            
                        row_1 = [sheet.cell(row=1, column=c).value for c in range(1, sheet.max_column + 1)]
                        row_2 = [sheet.cell(row=2, column=c).value for c in range(1, sheet.max_column + 1)]
                        
                        start_old = -1
                        start_new = -1
                        for idx, val in enumerate(row_1):
                            if val and 'old' in str(val).lower():
                                start_old = idx
                                break
                        for idx, val in enumerate(row_1):
                            if val and 'new' in str(val).lower():
                                start_new = idx
                                break
                                
                        if start_old == -1 or start_new == -1:
                            st.warning(f"Bỏ qua sheet **{sheetname}**: không nhận diện được cột Old/New.")
                            continue
                            
                        key_cols = []
                        for col_idx in range(start_old):
                            col_name = row_2[col_idx]
                            key_cols.append((col_name, col_idx))
                            
                        params = []
                        for col_idx in range(start_old, start_new):
                            params.append(row_2[col_idx])
                            
                        target_new_values = {}
                        for idx, p_name in enumerate(params):
                            new_col_idx = start_new + idx + 1
                            target_new_values[p_name] = sheet.cell(row=3, column=new_col_idx).value
                            
                        styles_row = 3
                        styles_by_col = {}
                        for col_idx in range(1, sheet.max_column + 1):
                            styles_by_col[col_idx] = sheet.cell(row=styles_row, column=col_idx)
                            
                        if max_r >= 3:
                            sheet.delete_rows(3, max_r - 2)
                            
                        row_count = 0
                        for cell in target_cells:
                            instances = []
                            
                            if vendor == 'Nokia':
                                tech_folder = '3G' if tech == '3G' else ('4G' if tech == '4G' else '5G')
                                sub_dir = 'DUMP' if tech == '3G' else 'Dump'
                                dump_path = os.path.join(DATABASE_DIR, tech_folder, sub_dir, 'NSN')
                                
                                pattern = os.path.join(dump_path, f"Export_{sheetname}__*.txt")
                                files = glob.glob(pattern)
                                if not files:
                                    all_files = os.listdir(dump_path) if os.path.exists(dump_path) else []
                                    files = [os.path.join(dump_path, f) for f in all_files if f.lower().startswith(f"export_{sheetname.lower()}__")]
                                    
                                if files:
                                    latest_f = max(files, key=os.path.getmtime)
                                    df_dump = load_dump_file(latest_f)
                                    if not df_dump.empty:
                                        if 'MO' in df_dump.columns:
                                            df_dump = df_dump.rename(columns={'MO': 'distName'})
                                        df_filtered = df_dump[df_dump['distName'].astype(str).str.startswith(cell['distName'])]
                                        for _, d_row in df_filtered.iterrows():
                                            instances.append(dict(d_row))
                                            
                            else: # Ericsson
                                tech_folder = '3G' if tech == '3G' else ('4G' if tech == '4G' else '5G')
                                sub_dir = 'DUMP' if tech == '3G' else 'Dump'
                                dump_path = os.path.join(DATABASE_DIR, tech_folder, sub_dir, 'ERA')
                                
                                csv_files = []
                                if os.path.exists(dump_path):
                                    for name in [sheetname, f"vsData{sheetname}"]:
                                        csv_files.extend(glob.glob(os.path.join(dump_path, f"{name}.csv")))
                                        csv_files.extend(glob.glob(os.path.join(dump_path, f"{name.lower()}.csv")))
                                        
                                if csv_files:
                                    latest_f = max(csv_files, key=os.path.getmtime)
                                    df_dump = load_dump_file(latest_f)
                                    if not df_dump.empty:
                                        dn_col = None
                                        for c in df_dump.columns:
                                            if str(c).lower() in ['dn', 'distname', 'managedelement', 'utrancell', 'eutrancellfdd', 'nrcellcu']:
                                                dn_col = c
                                                break
                                        if not dn_col:
                                            dn_col = df_dump.columns[0]
                                        df_filtered = df_dump[df_dump[dn_col].astype(str).str.contains(cell['cellName'], case=False) | df_dump[dn_col].astype(str).str.contains(cell['distName'], case=False)]
                                        for _, d_row in df_filtered.iterrows():
                                            instances.append(dict(d_row))
                                            
                            if not instances:
                                dummy = {'distName': cell['distName'], 'DN': cell['distName'], 'cellName': cell['cellName'], 'CellName': cell['cellName']}
                                for p_name in params:
                                    if vendor == 'Nokia':
                                        val = query_nokia_parameter_value(tech, sheetname, cell['distName'], p_name, {})
                                    else:
                                        val = query_ericsson_parameter_value(tech, sheetname, cell['cellName'], cell['distName'], p_name, {})
                                    dummy[p_name] = val
                                instances.append(dummy)
                                
                            for inst in instances:
                                row_idx = 3 + row_count
                                row_count += 1
                                
                                for col_name, col_offset in key_cols:
                                    cell_obj = sheet.cell(row=row_idx, column=col_offset + 1)
                                    copy_cell_style(styles_by_col[col_offset + 1], cell_obj)
                                    
                                    if str(col_name).lower() in ['distname', 'dn']:
                                        val = inst.get('distName', inst.get('DN', cell['distName']))
                                    elif str(col_name).lower() in ['cellname', 'name']:
                                        val = inst.get('cellName', inst.get('CellName', inst.get('name', cell['cellName'])))
                                    else:
                                        val = inst.get(col_name, '')
                                    cell_obj.value = val
                                    
                                for p_idx, p_name in enumerate(params):
                                    old_col_idx = start_old + p_idx + 1
                                    new_col_idx = start_new + p_idx + 1
                                    
                                    cell_old = sheet.cell(row=row_idx, column=old_col_idx)
                                    cell_new = sheet.cell(row=row_idx, column=new_col_idx)
                                    
                                    copy_cell_style(styles_by_col[old_col_idx], cell_old)
                                    copy_cell_style(styles_by_col[new_col_idx], cell_new)
                                    
                                    cell_old.value = inst.get(p_name, None)
                                    cell_new.value = target_new_values[p_name]
                                    
                        st.write(f"Đã xuất **{row_count}** cấu hình cho sheet **{sheetname}**")
                        
                    out_path = os.path.join(BASE_WORKSPACE_PATH, out_filename)
                    wb.save(out_path)
                    
                    st.markdown(f"""
                    <div class='success-box'>
                        🎉 <b>Tạo CR thành công!</b><br>
                        File kết quả đã được lưu tại:<br>
                        <code>{out_path}</code>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    with open(out_path, "rb") as f:
                        st.download_button(
                            label="📥 Tải xuống file CR kết quả",
                            data=f,
                            file_name=out_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình tạo CR: {e}")
                    st.exception(e)
    else:
        st.info("Vui lòng chọn hoặc nhập Cell/Trạm đầu vào để kiểm tra.")
