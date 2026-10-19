import base64
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import io
import zipfile
import os
import pathlib
import time
import json
from urllib.parse import quote_plus

# Thư viện kết nối CSDL SQL Server
try:
    from sqlalchemy import create_engine
    HAS_SQL_LIB = True
except Exception:
    HAS_SQL_LIB = False

# Thư viện hỗ trợ mở cửa sổ chọn thư mục Windows Explorer
try:
    import tkinter as tk
    from tkinter import filedialog
    HAS_TKINTER = True
except Exception:
    HAS_TKINTER = False

# --- CẤU HÌNH TRANG ---
st.set_page_config(
    page_title="Rà soát & Kiểm định XML BHYT - Bệnh Viện Quân Y 5",
    page_icon="🏥",
    layout="wide"
)

# --- TUỲ CHỈNH GIAO DIỆN & CSS ---
st.markdown("""
    <style>
        .main {
            background-color: #0f172a;
            color: #f8fafc;
        }
        .stSidebar {
            background-color: #1e293b !important;
            border-right: 1px solid #334155;
        }
        .stSidebar p, .stSidebar span, .stSidebar label, .stSidebar div, .stSidebar h3, .stSidebar h5 {
            color: #f8fafc !important;
        }
        .stSidebar input, .stSidebar textarea {
            color: #0f172a !important;
            background-color: #ffffff !important;
            font-weight: 600;
        }
        
        @keyframes ledGlow {
            0% { box-shadow: 0 0 10px rgba(16, 185, 129, 0.4), inset 0 0 10px rgba(16, 185, 129, 0.2); }
            50% { box-shadow: 0 0 25px rgba(16, 185, 129, 0.8), inset 0 0 15px rgba(16, 185, 129, 0.4); }
            100% { box-shadow: 0 0 10px rgba(16, 185, 129, 0.4), inset 0 0 10px rgba(16, 185, 129, 0.2); }
        }
        
        .hospital-header {
            background: linear-gradient(135deg, #064e3b 0%, #047857 50%, #059669 100%);
            color: #ffffff;
            padding: 25px 35px;
            border-radius: 12px;
            margin-bottom: 25px;
            border: 1px solid #10b981;
            animation: ledGlow 3s infinite alternate;
        }
        .hospital-header h1 {
            color: #ffffff !important;
            font-size: 2.2rem;
            margin: 0;
            font-weight: 800;
            letter-spacing: 1px;
            text-shadow: 0 0 10px rgba(255,255,255,0.5);
        }
        .hospital-header p {
            color: #a7f3d0 !important;
            font-size: 1.05rem;
            margin-top: 8px;
            margin-bottom: 0;
            font-weight: 600;
        }

        .glow-card {
            background-color: #1e293b;
            border: 1px solid #334155;
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 20px;
        }

        .stButton>button {
            background: linear-gradient(135deg, #059669 0%, #047857 100%);
            color: white !important;
            border-radius: 8px;
            font-weight: 600;
            border: 1px solid #10b981;
            padding: 0.6rem 1.2rem;
            box-shadow: 0 0 10px rgba(16, 185, 129, 0.2);
        }
        .stButton>button:hover {
            background: linear-gradient(135deg, #10b981 0%, #059669 100%);
            box-shadow: 0 0 20px rgba(16, 185, 129, 0.6);
        }
    </style>
""", unsafe_allow_html=True)

# --- THƯ MỤC VÀ CẤU HÌNH LƯU TRỮ CỐ ĐỊNH ---
CATEGORY_DIR = "saved_categories"
os.makedirs(CATEGORY_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(CATEGORY_DIR, "config_folder.txt")
CONFIG_AUTO_SCAN_FILE = os.path.join(CATEGORY_DIR, "config_auto_scan.txt")
CONFIG_MACHINE_SERVICES_FILE = os.path.join(CATEGORY_DIR, "config_machine_services.json")
CONFIG_DB_CREDENTIALS_FILE = os.path.join(CATEGORY_DIR, "config_db_credentials.json")
CONFIG_CUSTOM_COMPARES_FILE = os.path.join(CATEGORY_DIR, "config_custom_compares.json")

DEFAULT_DB_CONFIG = {
    "server": "192.168.0.98",
    "database": "BVQY5",
    "username": "sa",
    "password": "123@"
}

def load_db_credentials():
    if os.path.exists(CONFIG_DB_CREDENTIALS_FILE):
        try:
            with open(CONFIG_DB_CREDENTIALS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return DEFAULT_DB_CONFIG

def save_db_credentials(config_dict):
    try:
        with open(CONFIG_DB_CREDENTIALS_FILE, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

if 'db_config' not in st.session_state:
    st.session_state['db_config'] = load_db_credentials()

def fetch_catalog_from_his(query_string):
    if not HAS_SQL_LIB:
        return pd.DataFrame(), False, "Thiếu thư viện kết nối CSDL (pyodbc hoặc sqlalchemy chưa được cài đặt)."
    
    cfg = st.session_state['db_config']
    server = cfg.get("server", "192.168.0.98")
    database = cfg.get("database", "BVQY5")
    username = cfg.get("username", "sa")
    password = cfg.get("password", "123@")
    
    encoded_password = quote_plus(password)
    drivers_to_try = ["ODBC Driver 17 for SQL Server", "ODBC Driver 18 for SQL Server", "SQL Server"]
    
    for drv in drivers_to_try:
        try:
            conn_str = f"mssql+pyodbc://{username}:{encoded_password}@{server}/{database}?driver={quote_plus(drv)}&TrustServerCertificate=yes"
            engine = create_engine(conn_str)
            df_his = pd.read_sql(query_string, con=engine)
            return df_his, True, ""
        except Exception:
            continue
    return pd.DataFrame(), False, "Không thể kết nối SQL Server."

def load_saved_folder_path():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            pass
    return ""

def save_folder_path_to_disk(path):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            f.write(path)
    except Exception:
        pass

def load_auto_scan_config():
    if os.path.exists(CONFIG_AUTO_SCAN_FILE):
        try:
            with open(CONFIG_AUTO_SCAN_FILE, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
                if len(lines) >= 2:
                    return lines[0].strip() == "True", int(lines[1].strip())
        except Exception:
            pass
    return False, 5

def save_auto_scan_config_to_disk(is_enabled, interval):
    try:
        with open(CONFIG_AUTO_SCAN_FILE, "w", encoding="utf-8") as f:
            f.write(f"{is_enabled}\n{interval}")
    except Exception:
        pass

def load_machine_services_config():
    if os.path.exists(CONFIG_MACHINE_SERVICES_FILE):
        try:
            with open(CONFIG_MACHINE_SERVICES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_machine_services_config_to_disk(config_dict):
    try:
        with open(CONFIG_MACHINE_SERVICES_FILE, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

def load_custom_compares_config():
    if os.path.exists(CONFIG_CUSTOM_COMPARES_FILE):
        try:
            with open(CONFIG_CUSTOM_COMPARES_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "Đối chiếu Tên Dịch vụ kỹ thuật": {
            "src_cat": "Danh mục Trang thiết bị CLS_Cổng BHYT",
            "target_cats": ["Danh mục máy CLS"],
            "mappings": {}
        }
    }

def save_custom_compares_config_to_disk(config_dict):
    try:
        with open(CONFIG_CUSTOM_COMPARES_FILE, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=4)
    except Exception:
        pass

def load_categories_from_disk():
    cats = {
        "Danh mục DVKT_Cổng BHYT": pd.DataFrame(),
        "Danh mục Dịch vụ kỹ thuật (XML3)": pd.DataFrame(),
        "Danh mục Thuốc (XML2)": pd.DataFrame(),
        "Danh mục Vật tư y tế": pd.DataFrame(),
        "Danh mục Trang thiết bị CLS_Cổng BHYT": pd.DataFrame(),
        "Danh mục máy CLS": pd.DataFrame()
    }
    for file_path in pathlib.Path(CATEGORY_DIR).glob("*.*"):
        if file_path.name in ["config_folder.txt", "config_auto_scan.txt", "config_machine_services.json", "config_db_credentials.json", "config_custom_compares.json"]:
            continue
        if file_path.name.startswith("result_"):
            continue
        cat_name = file_path.stem
        try:
            if file_path.suffix == '.csv':
                cats[cat_name] = pd.read_csv(file_path)
            elif file_path.suffix in ['.xlsx', '.xls']:
                cats[cat_name] = pd.read_excel(file_path)
        except Exception:
            pass
    return cats

# Tải các kết quả đối chiếu đã lưu bền vững trên ổ cứng khi khởi động
def load_saved_compare_results():
    results = {}
    for file_path in pathlib.Path(CATEGORY_DIR).glob("result_*.csv"):
        # Tên file dạng result_<mapping_key_encoded>.csv
        try:
            mapping_key = file_path.stem[7:].replace("___", "__").replace("_VS_", "__VS__")
            df_res = pd.read_csv(file_path)
            results[mapping_key] = df_res
        except Exception:
            pass
    return results

def save_compare_result_to_disk(mapping_key, df_res):
    try:
        safe_key = mapping_key.replace("__VS__", "_VS_").replace("__", "___")
        file_path = os.path.join(CATEGORY_DIR, f"result_{safe_key}.csv")
        df_res.to_csv(file_path, index=False, encoding='utf-8-sig')
    except Exception:
        pass

if 'categories' not in st.session_state:
    st.session_state['categories'] = load_categories_from_disk()

if 'saved_folder_path' not in st.session_state:
    st.session_state['saved_folder_path'] = load_saved_folder_path()

if 'machine_services_config' not in st.session_state:
    st.session_state['machine_services_config'] = load_machine_services_config()

if 'custom_compares' not in st.session_state:
    st.session_state['custom_compares'] = load_custom_compares_config()

if 'multi_compare_results' not in st.session_state:
    st.session_state['multi_compare_results'] = load_saved_compare_results()

saved_auto, saved_interval = load_auto_scan_config()
if 'enable_auto_scan' not in st.session_state:
    st.session_state['enable_auto_scan'] = saved_auto
if 'scan_interval_mins' not in st.session_state:
    st.session_state['scan_interval_mins'] = saved_interval

def save_category_to_disk(cat_name, df):
    safe_name = "".join(c for c in cat_name if c.isalnum() or c in (' ', '_', '-')).strip()
    file_path = os.path.join(CATEGORY_DIR, f"{safe_name}.csv")
    df.to_csv(file_path, index=False, encoding='utf-8-sig')

def delete_category_from_disk(cat_name):
    safe_name = "".join(c for c in cat_name if c.isalnum() or c in (' ', '_', '-')).strip()
    file_path = os.path.join(CATEGORY_DIR, f"{safe_name}.csv")
    if os.path.exists(file_path):
        os.remove(file_path)

def decode_base64_xml(base64_str):
    try:
        decoded_bytes = base64.b64decode(base64_str.strip())
        try:
            decoded_text = decoded_bytes.decode('utf-8')
        except UnicodeDecodeError:
            decoded_text = decoded_bytes.decode('windows-1258', errors='ignore')
        return ET.fromstring(decoded_text)
    except Exception:
        return None

def deep_parse_element(element, parent_tag=""):
    row_data = {}
    def recursive_traverse(node, prefix=""):
        has_sub = False
        for child in node:
            child_tag = child.tag.split('}')[-1]
            tag_name = f"{prefix}_{child_tag}" if prefix else child_tag
            if len(list(child)) > 0:
                recursive_traverse(child, tag_name)
                has_sub = True
            else:
                val = child.text.strip() if child.text else ""
                if val:
                    row_data[tag_name] = val
        return has_sub
    recursive_traverse(element, parent_tag)
    return row_data

def extract_xml_records(root_element):
    rows = []
    candidates = []
    for elem in root_element.iter():
        if len(list(elem)) > 0 and all(len(list(sub)) == 0 for sub in elem):
            candidates.append(elem)
            
    if candidates:
        for elem in candidates:
            row_data = {child.tag.split('}')[-1]: (child.text.strip() if child.text else "") for child in elem if len(list(child)) == 0}
            if any(row_data.values()):
                rows.append(row_data)
    
    if not rows:
        for child in root_element:
            res = deep_parse_element(child)
            if res:
                rows.append(res)
    return rows

def parse_xml1_to_dict(element):
    result = {}
    for child in element.iter():
        if len(list(child)) == 0:
            if child.text and child.text.strip():
                tag_name = child.tag.split('}')[-1]
                result[tag_name] = child.text.strip()
    return result

def parse_xml_safely(uploaded_file_or_path):
    content_bytes = None
    if hasattr(uploaded_file_or_path, "read"):
        uploaded_file_or_path.seek(0)
        content_bytes = uploaded_file_or_path.read()
    else:
        with open(uploaded_file_or_path, "rb") as f:
            content_bytes = f.read()
            
    for enc in ['utf-8', 'windows-1258', 'latin-1']:
        try:
            text_content = content_bytes.decode(enc)
            return ET.fromstring(text_content), True, ""
        except Exception:
            continue
            
    return None, False, "Không thể giải mã định dạng file (lỗi mã hóa font)."

def process_xml_package(uploaded_file_or_path):
    sub_files_data = {}
    is_valid_format = True
    error_message = ""
    
    root, success, err_msg = parse_xml_safely(uploaded_file_or_path)
    if not success:
        return sub_files_data, False, err_msg
        
    try:
        hoso_list = root.findall('.//FILEHOSO')
        if not hoso_list:
            root_tag = root.tag.split('}')[-1].upper()
            if not root_tag.startswith('XML'):
                is_valid_format = False
                error_message = "Không tìm thấy thẻ <FILEHOSO> hoặc cấu trúc gốc không khớp chuẩn XML BHYT."
            sub_files_data[root_tag] = extract_xml_records(root)
            return sub_files_data, is_valid_format, error_message

        for fh in hoso_list:
            loai_hoso = fh.find('LOAIHOSO')
            noi_dung = fh.find('NOIDUNGFILE')
            
            if loai_hoso is not None and noi_dung is not None:
                l_type = loai_hoso.text.strip().upper()
                if noi_dung.text:
                    sub_root = decode_base64_xml(noi_dung.text)
                    if sub_root is not None:
                        if 'XML1' in l_type:
                            sub_files_data[l_type] = parse_xml1_to_dict(sub_root)
                        else:
                            extracted_rows = extract_xml_records(sub_root)
                            clean_rows = []
                            for r in extracted_rows:
                                clean_r = {str(k): str(v) for k, v in r.items()}
                                if clean_r and clean_r not in clean_rows:
                                    clean_rows.append(clean_r)
                            sub_files_data[l_type] = clean_rows
                    else:
                        is_valid_format = False
                        error_message = f"Không thể giải mã Base64 nội dung của loại hồ sơ: {l_type}"
            else:
                is_valid_format = False
                error_message = "Thẻ <FILEHOSO> thiếu <LOAIHOSO> hoặc <NOIDUNGFILE>."
    except Exception as e:
        is_valid_format = False
        error_message = f"Lỗi cấu trúc cú pháp XML (Syntax Error): {str(e)}"
        
    return sub_files_data, is_valid_format, error_message

@st.cache_data(show_spinner=False)
def get_optimized_bhyt_lookup(df_bhyt):
    bhyt_lookup_by_code = {}
    bhyt_lookup_by_name = {}
    if df_bhyt.empty:
        return bhyt_lookup_by_code, bhyt_lookup_by_name
        
    code_col = next((c for c in df_bhyt.columns if 'MA' in c.upper()), None)
    name_col = next((c for c in df_bhyt.columns if 'TEN_DVKT_PHEDUYET' in c.upper() or 'TEN' in c.upper()), None)
    machine_col = next((c for c in df_bhyt.columns if 'MAY' in c.upper()), None)
    
    for _, row in df_bhyt.iterrows():
        c_code = str(row[code_col]).strip() if code_col and pd.notna(row[code_col]) else ""
        c_name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
        c_mach = str(row[machine_col]).strip() if machine_col and pd.notna(row[machine_col]) else ""
        
        if c_code:
            bhyt_lookup_by_code[c_code.lower()] = {"name": c_name, "machine": c_mach}
        if c_name:
            bhyt_lookup_by_name[c_name.lower()] = {"code": c_code, "machine": c_mach}
            
    return bhyt_lookup_by_code, bhyt_lookup_by_name

def audit_cls_machine_code(all_tables):
    has_cls_error = False
    details = []
    
    target_cat = "Danh mục DVKT_Cổng BHYT"
    df_bhyt_cat = st.session_state['categories'].get(target_cat, pd.DataFrame())
    
    bhyt_lookup_by_code, bhyt_lookup_by_name = get_optimized_bhyt_lookup(df_bhyt_cat)

    machine_config_dict = st.session_state.get('machine_services_config', {})
    configured_services = machine_config_dict.get(target_cat, [])

    for t_name, t_data in all_tables.items():
        if any(keyword in t_name for keyword in ['XML3', 'DVKTV', 'CANLAMSANG', 'CLS']):
            if isinstance(t_data, list):
                for idx, row in enumerate(t_data):
                    ma_may_xml = row.get('MA_MAY', row.get('ma_may', '')).strip()
                    ma_dv_xml = row.get('MA_DICH_VU', row.get('MA_VАТ_TU', row.get('MA_DVKT', ''))).strip()
                    ten_dv_xml = row.get('TEN_DICH_VU', row.get('TEN_DVKT', row.get('TEN_VAT_TU', row.get('TEN_VАТ_TU', '')))).strip()
                    
                    if not ma_dv_xml and not ten_dv_xml:
                        continue
                    
                    if configured_services and ten_dv_xml not in configured_services:
                        continue
                    
                    match_info = bhyt_lookup_by_code.get(ma_dv_xml.lower()) or bhyt_lookup_by_name.get(ten_dv_xml.lower())
                    
                    error_msg = ""
                    if not match_info:
                        error_msg = "❌ Dịch vụ không tồn tại trên Cổng BHYT"
                    else:
                        standard_name = match_info.get("name", "")
                        if standard_name and ten_dv_xml.lower() != standard_name.lower():
                            error_msg = f"⚠️ Tên dịch vụ không khớp với Cổng BHYT ({standard_name})"
                        
                        standard_machine = match_info.get("machine", "")
                        if not ma_may_xml:
                            error_msg = "🚨 Thiếu mã máy (MA_MAY trống)"
                        elif standard_machine and ma_may_xml.lower() != standard_machine.lower():
                            error_msg = f"⚠️ Mã máy ({ma_may_xml}) không khớp với Cổng BHYT ({standard_machine})"

                    if error_msg:
                        has_cls_error = True
                        details.append({
                            "Bảng": t_name,
                            "Dòng số": idx + 1,
                            "Mã dịch vụ": ma_dv_xml,
                            "Tên dịch vụ XML": ten_dv_xml,
                            "Mã máy XML": ma_may_xml if ma_may_xml else "(Trống)",
                            "Tên BHYT chuẩn": match_info["name"] if match_info else "(Không tìm thấy)",
                            "Mã máy BHYT chuẩn": match_info["machine"] if match_info else "(Không có)",
                            "Lỗi / Trạng thái": error_msg
                        })

    return has_cls_error, details

def select_folder_dialog():
    if HAS_TKINTER:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder_selected = filedialog.askdirectory(master=root)
        root.destroy()
        return folder_selected
    return ""

def run_folder_scan(folder_path):
    xml_files = list(pathlib.Path(folder_path).glob("*.xml"))
    folder_summary = []
    all_folder_results = {}
    
    for x_file in xml_files:
        file_results, is_valid, err_msg = process_xml_package(x_file)
        all_folder_results[x_file.name] = file_results
        
        missing_machine_status = "Không có"
        cls_details = []
        if is_valid:
            has_missing, cls_details = audit_cls_machine_code(file_results)
            if has_missing:
                missing_machine_status = f"Có lỗi ({len(cls_details)} dịch vụ)"
            else:
                missing_machine_status = "Đầy đủ & Khớp chuẩn"
        else:
            missing_machine_status = "Không kiểm tra (Sai định dạng)"
        
        folder_summary.append({
            "Tên tệp": x_file.name,
            "Trạng thái định dạng": "✅ Đúng định dạng (QĐ 4750/3176)" if is_valid else f"❌ Sai định dạng ({err_msg})",
            "Kiểm tra mã máy CLS": missing_machine_status,
            "Chi tiết lỗi CLS": cls_details
        })
    return pd.DataFrame(folder_summary), all_folder_results, len(xml_files)

# --- TIÊU ĐỀ GIAO DIỆN CHÍNH ---
st.markdown("""
    <div class="hospital-header" translate="no">
        <h1>🏥 BỆNH VIỆN QUÂN Y 5</h1>
        <p>⚡ HỆ THỐNG RÀ SOÁT & KIỂM ĐỊNH DỮ LIỆU XML BHYT (QUYẾT ĐỊNH 4750 & 3176)</p>
    </div>
""", unsafe_allow_html=True)

# --- THANH SIDEBAR ---
with st.sidebar:
    st.markdown("### ⚙️ QUẢN TRỊ HỆ THỐNG")
    source_mode = st.radio(
        "Chọn chức năng chính:", 
        [
            "Rà soát tệp/thư mục XML", 
            "📦 Quản lý & Xem chi tiết Danh mục",
            "🔄 Đối chiếu Đa Danh mục & Ghim kết quả",
            "🔌 Kết nối & Đồng bộ trực tiếp từ HIS"
        ]
    )
    
    st.markdown("---")
    st.markdown("##### 📚 Quản lý Danh mục chuẩn")
    
    new_cat_name = st.text_input("Tên danh mục mới:")
    if st.button("➕ Thêm danh mục mới"):
        if new_cat_name and new_cat_name not in st.session_state['categories']:
            st.session_state['categories'][new_cat_name] = pd.DataFrame()
            save_category_to_disk(new_cat_name, pd.DataFrame())
            st.success(f"Đã thêm: {new_cat_name}")
            st.rerun()

    existing_cats = list(st.session_state['categories'].keys())
    if existing_cats:
        st.markdown("---")
        cat_action_target = st.selectbox("Chọn danh mục thao tác:", existing_cats, key="cat_action_select")
        
        ren_col1, ren_col2 = st.columns([2, 1])
        with ren_col1:
            new_renamed_val = st.text_input("Tên mới:", value=cat_action_target, key="input_rename_cat", label_visibility="collapsed")
        with ren_col2:
            if st.button("✏️ Sửa tên"):
                if new_renamed_val and new_renamed_val != cat_action_target:
                    if new_renamed_val in st.session_state['categories']:
                        st.error("Tên danh mục đã tồn tại!")
                    else:
                        df_old = st.session_state['categories'].pop(cat_action_target)
                        st.session_state['categories'][new_renamed_val] = df_old
                        delete_category_from_disk(cat_action_target)
                        save_category_to_disk(new_renamed_val, df_old)
                        st.success(f"Đã đổi tên thành: {new_renamed_val}")
                        st.rerun()
        
        if st.button("🗑️ Xóa danh mục này"):
            if len(existing_cats) > 1:
                del st.session_state['categories'][cat_action_target]
                delete_category_from_disk(cat_action_target)
                st.success(f"Đã xóa: {cat_action_target}")
                st.rerun()
            else:
                st.warning("Cần giữ lại ít nhất 1 danh mục trên hệ thống!")

    st.markdown("---")
    st.markdown("##### 📥 Đẩy file Excel/CSV vào danh mục")
    target_cat = st.selectbox("Chọn danh mục đích:", existing_cats, key="target_cat_select")
    ref_catalog_file = st.file_uploader(f"File nguồn cho [{target_cat}]:", type=['xlsx', 'xls', 'csv'], key="upload_cat_file")
    
    if ref_catalog_file is not None:
        if st.button("🚀 Cập nhật dữ liệu"):
            try:
                if ref_catalog_file.name.endswith('.csv'):
                    df_uploaded = pd.read_csv(ref_catalog_file)
                else:
                    df_uploaded = pd.read_excel(ref_catalog_file)
                
                st.session_state['categories'][target_cat] = df_uploaded
                save_category_to_disk(target_cat, df_uploaded)
                st.success(f"✅ Đã lưu {len(df_uploaded)} dòng!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi đọc file: {e}")

    st.markdown("---")
    st.markdown("##### 📦 Danh sách danh mục hiện có:")
    for c_name, c_df in st.session_state['categories'].items():
        st.write(f"- **{c_name}**: {len(c_df)} bản ghi")

# --- XỬ LÝ GIAO DIỆN CHÍNH ---
if source_mode == "🔌 Kết nối & Đồng bộ trực tiếp từ HIS":
    st.markdown("""
        <div class="glow-card">
            <h3>🔌 KẾT NỐI VÀ ĐỒNG BỘ TRỰC TIẾP CƠ SỞ DỮ LIỆU HIS</h3>
            <p style="color: #cbd5e1;">Cấu hình kết nối bảo mật đến Database SQL Server của Bệnh viện Quân y 5 để truy vấn và đồng bộ danh mục trực tiếp.</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.expander("🛠️ Cấu hình thông số kết nối Database (Đã bảo mật)", expanded=False):
        cfg = st.session_state['db_config']
        db_server = st.text_input("Server IP / Tên máy chủ:", value=cfg.get("server", "192.168.0.98"))
        db_name = st.text_input("Tên Database:", value=cfg.get("database", "BVQY5"))
        db_user = st.text_input("Tài khoản (Username):", value=cfg.get("username", "sa"))
        db_pass = st.text_input("Mật khẩu (Password):", type="password", value=cfg.get("password", "123@"))
        
        if st.button("💾 Lưu cấu hình kết nối"):
            new_cfg = {
                "server": db_server,
                "database": db_name,
                "username": db_user,
                "password": db_pass
            }
            st.session_state['db_config'] = new_cfg
            save_db_credentials(new_cfg)
            st.success("✅ Đã lưu cấu hình kết nối bảo mật lên hệ thống!")

    st.markdown("---")
    st.markdown("#### 📋 Duyệt & Xem chi tiết dữ liệu trực tiếp từ các bảng SQL Server")
    
    col_fetch_tbl, col_search_tbl = st.columns([1, 2])
    with col_fetch_tbl:
        btn_get_tables = st.button("🔄 Lấy danh sách toàn bộ bảng từ SQL Server")
    
    if btn_get_tables or 'his_all_tables_list' in st.session_state:
        if btn_get_tables:
            with st.spinner("Đang tải danh sách bảng từ SQL Server..."):
                query_tables = "SELECT TABLE_SCHEMA + '.' + TABLE_NAME AS FULL_TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
                df_tbls, success, err_msg = fetch_catalog_from_his(query_tables)
                if success and not df_tbls.empty:
                    st.session_state['his_all_tables_list'] = df_tbls['FULL_TABLE_NAME'].tolist()
                    st.success(f"✅ Tìm thấy **{len(st.session_state['his_all_tables_list'])}** bảng trong cơ sở dữ liệu `BVQY5`!")
                else:
                    st.error(f"❌ Lỗi lấy danh sách bảng: {err_msg}")

        if 'his_all_tables_list' in st.session_state and st.session_state['his_all_tables_list']:
            all_tables_found = st.session_state['his_all_tables_list']
            
            search_table_keyword = st.text_input("🔍 Nhập từ khóa tìm kiếm tên bảng (ví dụ: thuoc, bhyt, kcb...):", value="")
            
            filtered_tables = [
                tbl for tbl in all_tables_found 
                if search_table_keyword.lower() in tbl.lower()
            ] if search_table_keyword else all_tables_found
            
            st.caption(f"Hiển thị {len(filtered_tables)} / {len(all_tables_found)} bảng phù hợp.")

            if filtered_tables:
                selected_table_to_view = st.selectbox(
                    "🎯 Chọn bảng muốn xem dữ liệu chi tiết:", 
                    options=filtered_tables,
                    key="select_his_table_to_view"
                )
                
                col_v1, col_v2 = st.columns([1, 3])
                with col_v1:
                    btn_preview_table = st.button("👁️ Xem chi tiết bảng này", key="btn_view_his_table_now")
                with col_v2:
                    row_limit_view = st.number_input("Số bản ghi tối đa cần tải:", min_value=10, max_value=10000, value=100, step=50)

                if btn_preview_table:
                    with st.spinner(f"Đang tải {row_limit_view} bản ghi đầu tiên từ bảng [{selected_table_to_view}]..."):
                        preview_sql = f"SELECT TOP {row_limit_view} * FROM {selected_table_to_view}"
                        df_preview, preview_ok, preview_err = fetch_catalog_from_his(preview_sql)
                        if preview_ok and not df_preview.empty:
                            st.session_state['preview_his_table_name'] = selected_table_to_view
                            st.session_state['preview_his_table_df'] = df_preview
                            st.success(f"🎉 Đã tải thành công **{len(df_preview)}** dòng dữ liệu từ bảng `[{selected_table_to_view}]`!")
                        else:
                            st.error(f"❌ Không thể xem chi tiết bảng [{selected_table_to_view}]: {preview_err}")

                if 'preview_his_table_df' in st.session_state:
                    st.markdown(f"##### 📊 Dữ liệu chi tiết bảng: `{st.session_state.get('preview_his_table_name', '')}`")
                    st.dataframe(st.session_state['preview_his_table_df'], use_container_width=True)
                    
                    cats_for_sync = list(st.session_state['categories'].keys())
                    col_save_cat, col_save_btn = st.columns([2, 1])
                    with col_save_cat:
                        save_target_cat = st.selectbox("Chọn danh mục ứng dụng để lưu trữ bảng này:", options=cats_for_sync, key="save_preview_to_app_cat")
                    with col_save_btn:
                        st.write("")
                        st.write("")
                        if st.button("💾 Lưu vào Danh mục ứng dụng"):
                            st.session_state['categories'][save_target_cat] = st.session_state['preview_his_table_df']
                            save_category_to_disk(save_target_cat, st.session_state['preview_his_table_df'])
                            st.success(f"✅ Đã lưu dữ liệu bảng vào [{save_target_cat}] thành công!")
            else:
                st.warning("⚠️ Không tìm thấy bảng nào phù hợp với từ khóa tìm kiếm của bạn.")

    st.markdown("---")
    st.markdown("#### ⚡ Hoặc Tự nhập truy vấn SQL tùy chỉnh")
    
    default_sql = "SELECT TOP 100 * FROM dbo.DANHMUCKCB"
    his_query = st.text_area("Nhập câu lệnh SQL SELECT tùy chỉnh:", value=default_sql, height=100)
    
    if st.button("🚀 Thực thi câu lệnh SQL"):
        if not his_query.strip():
            st.warning("⚠️ Vui lòng nhập câu lệnh SQL truy vấn.")
        else:
            with st.spinner("Đang kết nối cơ sở dữ liệu SQL Server..."):
                df_synced, success, err_msg = fetch_catalog_from_his(his_query)
                if success and not df_synced.empty:
                    st.success(f"🎉 TRUY VẤN THÀNH CÔNG: Trả về **{len(df_synced)}** bản ghi!")
                    st.dataframe(df_synced, use_container_width=True)
                elif success and df_synced.empty:
                    st.warning("⚠️ Kết nối thành công nhưng câu lệnh không trả về bản ghi nào.")
                else:
                    st.error(f"❌ Lỗi kết nối CSDL HIS: {err_msg}")

elif source_mode == "🔄 Đối chiếu Đa Danh mục & Ghim kết quả":
    st.markdown("""
        <div class="glow-card">
            <h3>🔄 QUẢN LÝ MỤC NỘI DUNG ĐỐI CHIẾU & GHIM CỐ ĐỊNH KẾT QUẢ</h3>
            <p style="color: #cbd5e1;">Quản lý và tùy chỉnh các mục nội dung đối chiếu độc lập. Chọn một mục bên dưới để cấu hình hoặc xem lại kết quả đã ghim.</p>
        </div>
    """, unsafe_allow_html=True)
    
    custom_compares = st.session_state['custom_compares']
    compare_item_names = list(custom_compares.keys())
    
    # --- KHU VỰC QUẢN LÝ GỌN GÀNG (2 CỘT CÂN ĐỐI) ---
    col_add, col_manage = st.columns([1, 1])
    
    with col_add:
        st.markdown("##### ➕ Thêm mục nội dung mới")
        new_item_title = st.text_input("Tên mục đối chiếu mới:", placeholder="Ví dụ: Đối chiếu Mã máy CLS", label_visibility="collapsed", key="input_new_compare_item")
        if st.button("➕ Tạo mục đối chiếu mới", use_container_width=True):
            if new_item_title and new_item_title not in custom_compares:
                cats_list_all = list(st.session_state['categories'].keys())
                custom_compares[new_item_title] = {
                    "src_cat": cats_list_all[0] if cats_list_all else "",
                    "target_cats": [cats_list_all[1]] if len(cats_list_all) > 1 else [],
                    "mappings": {}
                }
                save_custom_compares_config_to_disk(custom_compares)
                st.success(f"✅ Đã tạo: {new_item_title}")
                st.rerun()
            elif new_item_title in custom_compares:
                st.error("Tên mục này đã tồn tại!")

    with col_manage:
        st.markdown("##### 🎯 Chọn & Quản lý Mục đối chiếu")
        if compare_item_names:
            selected_item_title = st.selectbox(
                "Chọn mục đối chiếu:",
                options=compare_item_names,
                label_visibility="collapsed",
                key="selectbox_active_compare_item"
            )
        else:
            selected_item_title = None
            st.info("Chưa có mục nào.")

    if selected_item_title and selected_item_title in custom_compares:
        active_item = custom_compares[selected_item_title]
        
        st.markdown("---")
        col_r1, col_r2, col_r3 = st.columns([3, 1, 1])
        with col_r1:
            renamed_title = st.text_input("Đổi tên mục hiện tại:", value=selected_item_title, label_visibility="collapsed", key=f"rename_input_{selected_item_title}")
        with col_r2:
            if st.button("✏️ Đổi tên", use_container_width=True, key=f"btn_rename_{selected_item_title}"):
                if renamed_title and renamed_title != selected_item_title:
                    if renamed_title in custom_compares:
                        st.error("Tên đã tồn tại!")
                    else:
                        custom_compares[renamed_title] = custom_compares.pop(selected_item_title)
                        save_custom_compares_config_to_disk(custom_compares)
                        st.success("Đã đổi tên thành công!")
                        st.rerun()
        with col_r3:
            if st.button("🗑️ Xóa mục", use_container_width=True, key=f"btn_del_{selected_item_title}"):
                if len(custom_compares) > 1:
                    del custom_compares[selected_item_title]
                    save_custom_compares_config_to_disk(custom_compares)
                    st.success("Đã xóa mục thành công!")
                    st.rerun()
                else:
                    st.warning("Cần giữ lại ít nhất 1 mục đối chiếu!")

        st.markdown("---")
        cats_list = list(st.session_state['categories'].keys())
        
        if len(cats_list) < 2:
            st.warning("⚠️ Bạn cần ít nhất 2 danh mục trên hệ thống để thực hiện đối chiếu.")
        else:
            col_cat_src, col_cat_targets = st.columns(2)
            
            default_src_idx = 0
            if active_item.get("src_cat") in cats_list:
                default_src_idx = cats_list.index(active_item["src_cat"])
                
            with col_cat_src:
                src_cat_name = st.selectbox(
                    "🏛️ Chọn Danh mục Nguồn Chuẩn:", 
                    options=cats_list, 
                    index=default_src_idx,
                    key=f"sel_src_cat_{selected_item_title}"
                )
            
            target_cat_options = [c for c in cats_list if c != src_cat_name]
            saved_targets = [t for t in active_item.get("target_cats", []) if t in target_cat_options]
            
            with col_cat_targets:
                selected_target_cats = st.multiselect(
                    "💻 Chọn các Danh mục cần đối chiếu (Có thể chọn nhiều):",
                    options=target_cat_options,
                    default=saved_targets,
                    key=f"sel_target_cats_{selected_item_title}"
                )
                
            if src_cat_name != active_item.get("src_cat") or selected_target_cats != active_item.get("target_cats"):
                active_item["src_cat"] = src_cat_name
                active_item["target_cats"] = selected_target_cats
                save_custom_compares_config_to_disk(custom_compares)
            
            df_src = st.session_state['categories'][src_cat_name]
            
            if df_src.empty:
                st.error(f"❌ Danh mục nguồn [{src_cat_name}] hiện đang trống dữ liệu.")
            elif not selected_target_cats:
                st.warning("⚠️ Vui lòng chọn ít nhất một danh mục đối chiếu ở mục bên phải.")
            else:
                st.markdown("#### ⚙️ Cấu hình cột khóa và so sánh chi tiết")
                src_cols = list(df_src.columns)
                
                for t_cat in selected_target_cats:
                    df_target = st.session_state['categories'][t_cat]
                    target_cols = list(df_target.columns)
                    
                    mapping_key = f"{selected_item_title}__{src_cat_name}__VS__{t_cat}"
                    saved_mapping = active_item.get("mappings", {}).get(t_cat, {})
                    
                    with st.expander(f"📦 Cấu hình & Kết quả đối chiếu với: [{t_cat}]", expanded=True):
                        col_m1, col_m2 = st.columns(2)
                        
                        default_src_match = saved_mapping.get('src_match_cols', [c for c in src_cols if 'MA' in c.upper() or 'CODE' in c.upper()][:1])
                        default_target_match = saved_mapping.get('target_match_cols', [c for c in target_cols if 'MA' in c.upper() or 'CODE' in c.upper()][:1])
                        
                        with col_m1:
                            sel_src_cols = st.multiselect(f"Cột mã định danh ở Nguồn [{src_cat_name}]:", options=src_cols, default=default_src_match, key=f"src_m_{mapping_key}")
                        with col_m2:
                            sel_target_cols = st.multiselect(f"Cột mã định danh tương ứng ở Đích [{t_cat}]:", options=target_cols, default=default_target_match, key=f"target_m_{mapping_key}")
                            
                        col_v1, col_v2 = st.columns(2)
                        default_src_val = saved_mapping.get('src_val_cols', [c for c in src_cols if 'TEN' in c.upper() or 'GIA' in c.upper()][:1])
                        default_target_val = saved_mapping.get('target_val_cols', [c for c in target_cols if 'TEN' in c.upper() or 'GIA' in c.upper()][:1])
                        
                        with col_v1:
                            sel_src_vals = st.multiselect(f"Cột so sánh nội dung ở Nguồn:", options=src_cols, default=default_src_val, key=f"src_v_{mapping_key}")
                        with col_v2:
                            sel_target_vals = st.multiselect(f"Cột so sánh nội dung ở Đích:", options=target_cols, default=default_target_val, key=f"target_v_{mapping_key}")
                        
                        if st.button(f"🚀 Tiến hành Đối chiếu & Ghim kết quả cho [{t_cat}]", key=f"btn_run_{mapping_key}"):
                            if len(sel_src_cols) != len(sel_target_cols) or len(sel_src_vals) != len(sel_target_vals):
                                st.error("⚠️ Số lượng cột khóa định danh hoặc cột so sánh giữa hai danh mục phải bằng nhau.")
                            elif not sel_src_cols:
                                st.warning("⚠️ Vui lòng chọn ít nhất 1 cột khóa định danh.")
                            else:
                                if "mappings" not in active_item:
                                    active_item["mappings"] = {}
                                active_item["mappings"][t_cat] = {
                                    "src_match_cols": sel_src_cols,
                                    "target_match_cols": sel_target_cols,
                                    "src_val_cols": sel_src_vals,
                                    "target_val_cols": sel_target_vals
                                }
                                save_custom_compares_config_to_disk(custom_compares)
                                
                                with st.spinner(f"Đang đối chiếu dữ liệu với [{t_cat}]..."):
                                    df_t_copy = df_target.copy()
                                    df_s_copy = df_src.copy()
                                    
                                    df_t_copy['_TARGET_KEY'] = df_t_copy[sel_target_cols].apply(
                                        lambda row: "_".join([str(val).strip() if pd.notna(val) else "" for val in row]), axis=1
                                    )
                                    target_dict = {row['_TARGET_KEY']: row for _, row in df_t_copy.iterrows()}
                                    
                                    compare_results = []
                                    for idx, s_row in df_s_copy.iterrows():
                                        s_key = "_".join([str(s_row[c]).strip() if pd.notna(s_row[c]) else "" for c in sel_src_cols])
                                        
                                        status = ""
                                        detail_diff = []
                                        if s_key not in target_dict:
                                            status = f"❌ Không có trong [{t_cat}]"
                                        else:
                                            t_row = target_dict[s_key]
                                            is_different = False
                                            for c_s, c_t in zip(sel_src_vals, sel_target_vals):
                                                v_s = str(s_row.get(c_s, '')).strip()
                                                v_t = str(t_row.get(c_t, '')).strip()
                                                if v_s.lower() != v_t.lower():
                                                    is_different = True
                                                    detail_diff.append(f"{c_s} ({v_s}) ≠ {c_t} ({v_t})")
                                            
                                            if is_different:
                                                status = f"⚠️ Sai lệch nội dung: {'; '.join(detail_diff)}"
                                            else:
                                                status = "✅ Khớp hoàn toàn"
                                                
                                        row_res = {"STT": idx + 1, "Mã khóa": s_key, "Trạng thái đối chiếu": status}
                                        for c in sel_src_cols + sel_src_vals:
                                            row_res[f"Nguồn_{c}"] = s_row.get(c, '')
                                        if s_key in target_dict:
                                            t_row = target_dict[s_key]
                                            for c in sel_target_cols + sel_target_vals:
                                                row_res[f"Đích_{c}"] = t_row.get(c, '')
                                        compare_results.append(row_res)
                                        
                                    df_compare_res = pd.DataFrame(compare_results)
                                    st.session_state['multi_compare_results'][mapping_key] = df_compare_res
                                    # Lưu kết quả xuống ổ cứng bền vững
                                    save_compare_result_to_disk(mapping_key, df_compare_res)
                                    st.success(f"🎉 Đã hoàn tất đối chiếu, ghim và lưu kết quả bền vững với [{t_cat}]!")

                        if mapping_key in st.session_state['multi_compare_results']:
                            df_res = st.session_state['multi_compare_results'][mapping_key]
                            st.markdown(f"##### 📊 KẾT QUẢ ĐỐI CHIẾU ĐÃ GHIM: `{src_cat_name}` ⇄ `{t_cat}`")
                            
                            m_cnt = len(df_res[df_res['Trạng thái đối chiếu'].str.contains('✅')])
                            d_cnt = len(df_res[df_res['Trạng thái đối chiếu'].str.contains('⚠️')])
                            ms_cnt = len(df_res[df_res['Trạng thái đối chiếu'].str.contains('❌')])
                            
                            cs1, cs2, cs3, cs4 = st.columns(4)
                            cs1.metric("Tổng bản ghi", len(df_res))
                            cs2.metric("✅ Khớp hoàn toàn", m_cnt)
                            cs3.metric("⚠️ Sai lệch", d_cnt)
                            cs4.metric("❌ Chưa có", ms_cnt)
                            
                            flt = st.selectbox("Lọc kết quả:", options=["Tất cả", "Chỉ xem lỗi/sai lệch (⚠️ & ❌)", "✅ Khớp hoàn toàn", "⚠️ Sai lệch nội dung", "❌ Chưa có"], key=f"flt_{mapping_key}")
                            df_disp = df_res.copy()
                            if flt == "Chỉ xem lỗi/sai lệch (⚠️ & ❌)":
                                df_disp = df_disp[~df_disp['Trạng thái đối chiếu'].str.contains('✅')]
                            elif flt != "Tất cả":
                                df_disp = df_disp[df_disp['Trạng thái đối chiếu'].str.contains(flt[:1])]
                                
                            st.dataframe(df_disp, use_container_width=True)
                            
                            csv_data = df_res.to_csv(index=False).encode('utf-8-sig')
                            st.download_button(
                                label=f"📥 Tải xuống báo cáo đối chiếu [{t_cat}] (CSV)",
                                data=csv_data,
                                file_name=f"Bao_cao_{selected_item_title}_{src_cat_name}_vs_{t_cat}.csv",
                                mime="text/csv",
                                key=f"dl_{mapping_key}"
                            )

elif source_mode == "📦 Quản lý & Xem chi tiết Danh mục":
    st.markdown("""
        <div class="glow-card">
            <h3>📦 QUẢN LÝ VÀ XEM CHI TIẾT DANH MỤC CHUẨN BHYT</h3>
            <p style="color: #cbd5e1;">Tích chọn danh mục <b>Danh mục DVKT_Cổng BHYT</b> để cấu hình các dịch vụ cần rà soát kiểm tra mã máy.</p>
        </div>
    """, unsafe_allow_html=True)
    
    cats_list = list(st.session_state['categories'].keys())
    
    selected_view_cat = st.radio(
        "📌 TÍCH CHỌN DANH MỤC CẦN XEM & CẤU HÌNH:",
        options=cats_list,
        format_func=lambda name: f"📄 {name}  —  ({len(st.session_state['categories'][name])} bản ghi)"
    )
    
    st.markdown("---")
    
    if selected_view_cat:
        cat_df = st.session_state['categories'][selected_view_cat]
        
        if not cat_df.empty:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.success(f"📊 **{selected_view_cat}** | **{len(cat_df)}** dòng | **{len(cat_df.columns)}** cột")
            with col2:
                csv_cat_data = cat_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label=f"📥 Tải xuống (CSV)",
                    data=csv_cat_data,
                    file_name=f"{selected_view_cat}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            st.markdown(f"#### ⚙️ Cấu hình chọn dịch vụ kiểm tra Mã máy cho: [{selected_view_cat}]")
            
            cat_name_cols = [c for c in cat_df.columns if 'TEN_DVKT_PHEDUYET' in c.upper() or 'TEN' in c.upper()]
            if cat_name_cols:
                col_for_names = cat_name_cols[0]
                all_services = cat_df[col_for_names].dropna().astype(str).unique().tolist()
                
                if selected_view_cat not in st.session_state['machine_services_config']:
                    st.session_state['machine_services_config'][selected_view_cat] = []
                
                saved_config = st.session_state['machine_services_config'][selected_view_cat]

                search_query = st.text_input("🔍 Nhập tên dịch vụ để lọc:", value=st.session_state.get(f"kw_{selected_view_cat}", ""), key=f"input_kw_{selected_view_cat}")
                st.session_state[f"kw_{selected_view_cat}"] = search_query
                
                filtered_services = [s for s in all_services if search_query.lower() in s.lower()] if search_query else all_services
                
                st.write(f"Hiển thị **{len(filtered_services)}** / {len(all_services)} dịch vụ phù hợp:")

                col_btn_all, col_btn_none = st.columns(2)
                with col_btn_all:
                    if st.button("☑️ Tích chọn tất cả kết quả đang hiển thị", key=f"btn_select_all_{selected_view_cat}"):
                        for s in filtered_services:
                            if s not in saved_config:
                                saved_config.append(s)
                            st.session_state[f"chk_svc_{selected_view_cat}_{s}"] = True
                        st.session_state['machine_services_config'][selected_view_cat] = saved_config
                        save_machine_services_config_to_disk(st.session_state['machine_services_config'])
                        st.success("Đã chọn tất cả dịch vụ hiển thị!")
                        st.rerun()
                with col_btn_none:
                    if st.button("❌ Bỏ chọn tất cả", key=f"btn_clear_all_{selected_view_cat}"):
                        st.session_state['machine_services_config'][selected_view_cat] = []
                        for s in all_services:
                            st.session_state[f"chk_svc_{selected_view_cat}_{s}"] = False
                        save_machine_services_config_to_disk(st.session_state['machine_services_config'])
                        st.success("Đã bỏ chọn tất cả!")
                        st.rerun()

                updated_selection = list(saved_config)
                container_box = st.container(height=350)
                with container_box:
                    for s in filtered_services:
                        is_checked = s in saved_config
                        checkbox_val = st.checkbox(s, value=is_checked, key=f"chk_svc_{selected_view_cat}_{s}")
                        if checkbox_val and s not in updated_selection:
                            updated_selection.append(s)
                        elif not checkbox_val and s in updated_selection:
                            updated_selection.remove(s)
                
                if updated_selection != saved_config:
                    st.session_state['machine_services_config'][selected_view_cat] = updated_selection
                    save_machine_services_config_to_disk(st.session_state['machine_services_config'])
                
                st.info(f"💾 Đã lưu cấu hình: **{len(updated_selection)}** dịch vụ được chọn kiểm tra mã máy trong danh mục này.")
            else:
                st.warning("⚠️ Không tìm thấy cột tên dịch vụ trong danh mục này để tạo bộ lọc.")

            st.markdown("---")
            st.dataframe(cat_df, use_container_width=True)
        else:
            st.warning(f"⚠️ Danh mục **{selected_view_cat}** hiện đang trống. Hãy tải file lên từ thanh bên trái.")

else:
    sub_source_mode = st.radio("Chọn phương thức rà soát:", ["Quét tự động theo thư mục (Folder Scan)", "Chọn tệp trực tiếp từ máy tính (File Picker)"], horizontal=True)

    if sub_source_mode == "Quét tự động theo thư mục (Folder Scan)":
        st.markdown("#### ⏱️ Cấu hình tự động quét theo thời gian")
        col_auto1, col_auto2 = st.columns([1, 2])
        
        def update_auto_scan_settings():
            save_auto_scan_config_to_disk(st.session_state['enable_auto_scan'], st.session_state['scan_interval_mins'])
            st.session_state['target_scan_time'] = time.time() + (st.session_state['scan_interval_mins'] * 60)

        with col_auto1:
            st.checkbox("Bật tự động quét định kỳ", key='enable_auto_scan', on_change=update_auto_scan_settings)
        with col_auto2:
            st.number_input("Khoảng thời gian tự động quét lại (phút):", min_value=1, max_value=120, key='scan_interval_mins', on_change=update_auto_scan_settings)

        st.markdown("---")

        col_btn_action, col_timer_display = st.columns([2, 3])
        with col_btn_action:
            trigger_scan = st.button("🚀 Bắt đầu rà soát toàn bộ thư mục")

        with col_timer_display:
            if st.session_state['enable_auto_scan']:
                current_time = time.time()
                interval_secs = st.session_state['scan_interval_mins'] * 60
                
                if 'target_scan_time' not in st.session_state:
                    st.session_state['target_scan_time'] = current_time + interval_secs
                
                target_epoch_ms = int(st.session_state['target_scan_time'] * 1000)
                
                timer_html = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <style>
                        body {{
                            background-color: transparent;
                            color: #38bdf8;
                            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
                            font-weight: 600;
                            font-size: 1.05rem;
                            margin: 0;
                            padding-top: 5px;
                        }}
                    </style>
                </head>
                <body>
                    ⏳ Tự động quét lại sau: <span id="timer-val">--:--</span>
                    <script>
                        var targetTime = {target_epoch_ms};
                        function updateTimer() {{
                            var now = new Date().getTime();
                            var distance = targetTime - now;
                            var timerElement = document.getElementById("timer-val");
                            if (!timerElement) return;
                            
                            if (distance <= 0) {{
                                timerElement.innerHTML = "Đang quét...";
                                clearInterval(window.timerInterval);
                                
                                var buttons = window.parent.document.querySelectorAll('button');
                                for (var i = 0; i < buttons.length; i++) {{
                                    if (buttons[i].innerText.includes("Bắt đầu rà soát toàn bộ thư mục")) {{
                                        buttons[i].click();
                                        break;
                                    }}
                                }}
                            }} else {{
                                var totalSeconds = Math.floor(distance / 1000);
                                var m = Math.floor(totalSeconds / 60);
                                var s = totalSeconds % 60;
                                timerElement.innerHTML = (m < 10 ? "0" + m : m) + ":" + (s < 10 ? "0" + s : s);
                            }}
                        }}
                        window.timerInterval = setInterval(updateTimer, 1000);
                        updateTimer();
                    </script>
                </body>
                </html>
                """
                components.html(timer_html, height=35)
            else:
                st.markdown("<div style='padding-top: 5px; color: #94a3b8; font-style: italic;'>Trạng thái: Tự động quét đang tắt</div>", unsafe_allow_html=True)

        st.markdown("### 📁 Quét và Kiểm định hàng loạt tệp XML theo thư mục")
        
        col_input, col_btn = st.columns([4, 1])
        
        with col_btn:
            st.write("") 
            st.write("") 
            if st.button("📂 Chọn thư mục..."):
                picked_dir = select_folder_dialog()
                if picked_dir:
                    st.session_state['saved_folder_path'] = picked_dir
                    save_folder_path_to_disk(picked_dir)
                    if 'folder_summary' in st.session_state:
                        del st.session_state['folder_summary']
                    st.session_state['target_scan_time'] = time.time() + (st.session_state['scan_interval_mins'] * 60)
                    st.rerun()

        with col_input:
            folder_path = st.text_input(
                "Đường dẫn thư mục chứa tệp XML (Được lưu cố định):",
                value=st.session_state['saved_folder_path']
            )
            if folder_path != st.session_state['saved_folder_path']:
                st.session_state['saved_folder_path'] = folder_path
                save_folder_path_to_disk(folder_path)
                if 'folder_summary' in st.session_state:
                    del st.session_state['folder_summary']
                st.session_state['target_scan_time'] = time.time() + (st.session_state['scan_interval_mins'] * 60)
        
        if folder_path and os.path.exists(folder_path):
            xml_files = list(pathlib.Path(folder_path).glob("*.xml"))
            
            if xml_files:
                st.info(f"🔍 Tìm thấy **{len(xml_files)}** tệp XML trong thư mục cố định này.")

                current_time = time.time()
                interval_secs = st.session_state['scan_interval_mins'] * 60

                if 'target_scan_time' not in st.session_state:
                    st.session_state['target_scan_time'] = current_time + interval_secs

                if trigger_scan:
                    st.session_state['target_scan_time'] = time.time() + interval_secs

                should_run_scan = False
                if 'folder_summary' not in st.session_state:
                    should_run_scan = True
                elif trigger_scan:
                    should_run_scan = True
                elif st.session_state['enable_auto_scan']:
                    if current_time >= st.session_state['target_scan_time']:
                        should_run_scan = True

                if should_run_scan:
                    with st.spinner("Đang tiến hành rà soát thư mục siêu tốc..."):
                        df_sum, res_dict, total_f = run_folder_scan(folder_path)
                        st.session_state['folder_summary'] = df_sum
                        st.session_state['folder_results'] = res_dict
                        st.session_state['target_scan_time'] = time.time() + interval_secs
                    if trigger_scan:
                        st.success("🎉 Đã hoàn tất rà soát toàn bộ thư mục!")
                        st.rerun()

            else:
                st.warning("⚠️ Không tìm thấy tệp XML nào trong thư mục này.")
        elif folder_path:
            st.error("❌ Đường dẫn thư mục hiện tại không tồn tại trên hệ thống.")

    else:
        st.markdown("### 💻 Chọn tệp XML trực tiếp từ máy tính")
        uploaded_files = st.file_uploader(
            "Chọn một hoặc nhiều tệp XML hồ sơ BHYT từ máy tính:",
            type=['xml'],
            accept_multiple_files=True
        )
        
        if uploaded_files:
            folder_summary = []
            all_folder_results = {}
            
            for uploaded_file in uploaded_files:
                file_results, is_valid, err_msg = process_xml_package(uploaded_file)
                all_folder_results[uploaded_file.name] = file_results
                
                missing_machine_status = "Không có"
                cls_details = []
                if is_valid:
                    has_missing, cls_details = audit_cls_machine_code(file_results)
                    if has_missing:
                        missing_machine_status = f"Có lỗi ({len(cls_details)} dịch vụ)"
                    else:
                        missing_machine_status = "Đầy đủ & Khớp chuẩn"
                else:
                    missing_machine_status = "Không kiểm tra (Sai định dạng)"
                
                folder_summary.append({
                    "Tên tệp": uploaded_file.name,
                    "Trạng thái định dạng": "✅ Đúng định dạng (QĐ 4750/3176)" if is_valid else f"❌ Sai định dạng ({err_msg})",
                    "Kiểm tra mã máy CLS": missing_machine_status,
                    "Chi tiết lỗi CLS": cls_details
                })
            
            st.session_state['folder_summary'] = pd.DataFrame(folder_summary)
            st.session_state['folder_results'] = all_folder_results

    # --- HIỂN THỊ BẢNG TỔNG HỢP VÀ CHI TIẾT LỖI ---
    if 'folder_summary' in st.session_state:
        st.markdown("---")
        st.markdown("### 📊 Bảng tổng hợp kết quả rà soát")
        df_summary = st.session_state['folder_summary']
        
        display_df = df_summary[['Tên tệp', 'Trạng thái định dạng', 'Kiểm tra mã máy CLS']].copy()
        st.dataframe(display_df, use_container_width=True)
        
        valid_count = len(df_summary[df_summary['Trạng thái định dạng'].str.contains('✅')])
        invalid_count = len(df_summary[df_summary['Trạng thái định dạng'].str.contains('❌')])
        missing_machine_file_count = len(df_summary[df_summary['Kiểm tra mã máy CLS'].str.contains('Có lỗi')])
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Tổng số tệp quét", len(df_summary))
        col2.metric("Số tệp Đúng định dạng", valid_count)
        col3.metric("Số tệp Sai định dạng", invalid_count)
        col4.metric("🚨 Tệp có lỗi Mã máy/Cổng BHYT", missing_machine_file_count)
        
        if missing_machine_file_count > 0:
            st.markdown("---")
            st.markdown("#### 🚨 Danh sách các tệp có lỗi Mã máy / Sai lệch với Cổng BHYT")
            
            df_missing_files = df_summary[df_summary['Kiểm tra mã máy CLS'].str.contains('Có lỗi')]
            
            for idx, m_row in df_missing_files.iterrows():
                f_name = m_row['Tên tệp']
                f_details = m_row['Chi tiết lỗi CLS']
                
                col_info, col_eye = st.columns([5, 1])
                with col_info:
                    st.write(f"📄 **{f_name}** — ⚠️ Có **{len(f_details)}** dịch vụ cần xử lý.")
                with col_eye:
                    eye_key = f"show_detail_{f_name}"
                    if eye_key not in st.session_state:
                        st.session_state[eye_key] = False
                        
                    btn_label = "👁️ Xem" if not st.session_state[eye_key] else "👁️‍🗨️ Đóng"
                    if st.button(btn_label, key=f"btn_eye_{f_name}__"):
                        st.session_state[eye_key] = not st.session_state[eye_key]
                        st.rerun()
                
                if st.session_state.get(eye_key, False):
                    with st.container():
                        st.info(f"Chi tiết đối chiếu dịch vụ trong tệp **{f_name}** với **Danh mục DVKT_Cổng BHYT**:")
                        st.dataframe(pd.DataFrame(f_details), use_container_width=True)
                        
                        if st.button(f"🎯 Mở bảng XML và trỏ đánh dấu vị trí lỗi trong tệp [{f_name}]", key=f"jump_btn_{f_name}"):
                            st.session_state['jump_target_file'] = f_name
                            st.session_state['jump_target_table'] = f_details[0]['Bảng'] if f_details else ""
                            st.session_state['jump_target_rows'] = [d['Dòng số'] for d in f_details]
                            st.success("Đã kích hoạt định vị! Vui lòng xem phần Chi tiết từng tệp bên dưới.")
                st.markdown("---")

        st.markdown("### 🔍 Chi tiết từng tệp & Đối chiếu danh mục chuẩn")
        
        file_list_names = df_summary['Tên tệp'].tolist()
        default_file_idx = 0
        target_jump_file = st.session_state.get('jump_target_file', None)
        if target_jump_file in file_list_names:
            default_file_idx = file_list_names.index(target_jump_file)
            
        selected_file_name = st.selectbox("Chọn tệp XML kiểm tra:", file_list_names, index=default_file_idx)
        
        if selected_file_name:
            row_info = df_summary[df_summary['Tên tệp'] == selected_file_name].iloc[0]
            st.info(f"**Trạng thái tệp:** {row_info['Trạng thái định dạng']} | **Mã máy CLS:** {row_info['Kiểm tra mã máy CLS']}")
            
            if row_info['Chi tiết lỗi CLS']:
                st.warning(f"⚠️ Phát hiện **{len(row_info['Chi tiết lỗi CLS'])}** dịch vụ cận lâm sàng có lỗi hoặc sai lệch thông tin mã máy!")
                st.dataframe(pd.DataFrame(row_info['Chi tiết lỗi CLS']), use_container_width=True)
            else:
                st.success("✨ Khớp hoàn toàn và đầy đủ mã máy theo Cổng BHYT.")
                
            all_tables = st.session_state['folder_results'].get(selected_file_name, {})
            if all_tables:
                sorted_tables = sorted(list(all_tables.keys()))
                
                default_table_idx = 0
                target_jump_table = st.session_state.get('jump_target_table', None)
                if selected_file_name == target_jump_file and target_jump_table in sorted_tables:
                    default_table_idx = sorted_tables.index(target_jump_table)
                    
                selected_table = st.selectbox("Xem cấu trúc bảng dữ liệu của tệp:", sorted_tables, index=default_table_idx, key="sel_tbl_folder")
                if selected_table:
                    data = all_tables[selected_table]
                    if isinstance(data, list) and len(data) > 0:
                        df_data = pd.DataFrame(data)
                        
                        target_rows_to_highlight = []
                        if selected_file_name == target_jump_file and selected_table == target_jump_table:
                            target_rows_to_highlight = st.session_state.get('jump_target_rows', [])

                        def highlight_missing_machine(row):
                            row_idx = row.name + 1
                            if row_idx in target_rows_to_highlight:
                                return ['background-color: rgba(239, 68, 68, 0.4); font-weight: bold;'] * len(row)
                            return [''] * len(row)

                        compare_cat_name = st.selectbox("Chọn danh mục chuẩn để đối chiếu:", ["-- Không đối chiếu --"] + list(st.session_state['categories'].keys()))
                        
                        if compare_cat_name != "-- Không đối chiếu --":
                            standard_catalog = st.session_state['categories'][compare_cat_name]
                            if not standard_catalog.empty:
                                st.markdown(f"#### 🔄 Cấu hình chọn trường đối chiếu với [{compare_cat_name}]:")
                                
                                col_xml_field, col_cat_field = st.columns(2)
                                
                                xml_cols = list(df_data.columns)
                                default_xml_cols = [c for c in xml_cols if 'MA_' in c.upper()]
                                if not default_xml_cols and xml_cols:
                                    default_xml_cols = [xml_cols[0]]
                                        
                                cat_cols_list = list(standard_catalog.columns)
                                default_cat_cols = [c for c in cat_cols_list if 'TEN_DVKT_PHEDUYET' in c.upper() or 'TEN' in c.upper()]
                                if not default_cat_cols and cat_cols_list:
                                    default_cat_cols = [cat_cols_list[0]]
                                        
                                with col_xml_field:
                                    match_cols = st.multiselect("Chọn các cột mã ở tệp XML:", xml_cols, default=default_xml_cols, key=f"xml_cols_{selected_file_name}_{selected_table}")
                                with col_cat_field:
                                    cat_cols = st.multiselect("Chọn các cột tương ứng trong danh mục chuẩn:", cat_cols_list, default=default_cat_cols, key=f"cat_cols_{compare_cat_name}_{selected_file_name}_{selected_table}")
                                
                                if match_cols and cat_cols and len(match_cols) == len(cat_cols):
                                    def check_row_match(row):
                                        is_empty_all = True
                                        for m_col in match_cols:
                                            val = str(row.get(m_col, '')).strip()
                                            if val and val.lower() != 'nan':
                                                is_empty_all = False
                                                break
                                        if is_empty_all:
                                            return "⏭️ Bỏ qua (Trống mã)"

                                        for m_col, c_col in zip(match_cols, cat_cols):
                                            val = str(row.get(m_col, '')).strip()
                                            if not val or val.lower() == 'nan':
                                                continue
                                            standard_codes = set(standard_catalog[c_col].astype(str).str.strip().values)
                                            if val not in standard_codes:
                                                return "❌ Không có trong danh mục chuẩn"
                                        return "✅ Khớp danh mục"

                                    df_data['So_sanh_danh_muc'] = df_data.apply(check_row_match, axis=1)

                                    if target_rows_to_highlight:
                                        st.error(f"🎯 **Hệ thống đang trỏ đánh dấu tô đỏ** tại các dòng số: **{target_rows_to_highlight}** bị lỗi mã máy.")

                                    st.markdown("##### 📋 Bảng dữ liệu chi tiết sau đối chiếu (Được tô màu đánh dấu vị trí lỗi):")
                                    st.dataframe(df_data.style.apply(highlight_missing_machine, axis=1), use_container_width=True)
                                else:
                                    st.warning("⚠️ Số lượng cột chọn ở XML và Danh mục chuẩn phải bằng nhau để đối chiếu tương ứng.")
                                    st.dataframe(df_data.style.apply(highlight_missing_machine, axis=1), use_container_width=True)
                            else:
                                st.warning(f"Danh mục [{compare_cat_name}] hiện đang trống.")
                                st.dataframe(df_data.style.apply(highlight_missing_machine, axis=1), use_container_width=True)
                        else:
                            st.dataframe(df_data.style.apply(highlight_missing_machine, axis=1), use_container_width=True)
                            
                    elif isinstance(data, dict):
                        st.dataframe(pd.DataFrame([{"Trường dữ liệu": k, "Giá trị": v} for k, v in data.items()]), use_container_width=True)