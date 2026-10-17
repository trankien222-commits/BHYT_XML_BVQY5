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

# --- TUỲ CHỈNH GIAO DIỆN & ĐỘ TƯƠNG PHẢN SIDEBAR (CSS) ---
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
        .stSidebar input {
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

def load_categories_from_disk():
    cats = {
        "Danh mục Dịch vụ kỹ thuật (XML3)": pd.DataFrame(),
        "Danh mục Thuốc (XML2)": pd.DataFrame(),
        "Danh mục Vật tư y tế": pd.DataFrame()
    }
    for file_path in pathlib.Path(CATEGORY_DIR).glob("*.*"):
        if file_path.name in ["config_folder.txt", "config_auto_scan.txt", "config_machine_services.json"]:
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

if 'categories' not in st.session_state:
    st.session_state['categories'] = load_categories_from_disk()

if 'saved_folder_path' not in st.session_state:
    st.session_state['saved_folder_path'] = load_saved_folder_path()

if 'machine_services_config' not in st.session_state:
    st.session_state['machine_services_config'] = load_machine_services_config()

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

def audit_cls_machine_code(all_tables):
    has_cls_missing_machine = False
    details = []
    
    all_configured_services = []
    machine_config_dict = st.session_state.get('machine_services_config', {})
    for cat_name, s_list in machine_config_dict.items():
        for s in s_list:
            if s not in all_configured_services:
                all_configured_services.append(s)
                
    for t_name, t_data in all_tables.items():
        if any(keyword in t_name for keyword in ['XML3', 'DVKTV', 'CANLAMSANG', 'CLS']):
            if isinstance(t_data, list):
                for idx, row in enumerate(t_data):
                    ma_may = row.get('MA_MAY', row.get('ma_may', ''))
                    ma_dich_vu = row.get('MA_DICH_VU', row.get('MA_VАТ_TU', row.get('MA_DVKT', '')))
                    ten_dich_vu = row.get('TEN_DICH_VU', row.get('TEN_DVKT', row.get('TEN_VAT_TU', row.get('TEN_VАТ_TU', ''))))
                    
                    if (not ma_dich_vu or ma_dich_vu.strip() == "") and (not ten_dich_vu or ten_dich_vu.strip() == ""):
                        continue
                    
                    if all_configured_services and ten_dich_vu not in all_configured_services:
                        continue
                    
                    if not ma_may or ma_may.strip() == "":
                        has_cls_missing_machine = True
                        details.append({
                            "Bảng": t_name,
                            "Dòng số": idx + 1,
                            "Mã dịch vụ/VTYT": ma_dich_vu,
                            "Tên dịch vụ": ten_dich_vu,
                            "Lỗi": "Thiếu mã máy (MA_MAY trống)"
                        })
    return has_cls_missing_machine, details

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
                missing_machine_status = f"Có thiếu ({len(cls_details)} dịch vụ)"
            else:
                missing_machine_status = "Đầy đủ mã máy"
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
    source_mode = st.radio("Chọn chức năng chính:", ["Rà soát tệp/thư mục XML", "📦 Quản lý & Xem chi tiết Danh mục"])
    
    st.markdown("---")
    st.markdown("##### 📚 Thêm/Xóa Danh mục chuẩn")
    
    new_cat_name = st.text_input("Tên danh mục mới:")
    if st.button("➕ Thêm danh mục mới"):
        if new_cat_name and new_cat_name not in st.session_state['categories']:
            st.session_state['categories'][new_cat_name] = pd.DataFrame()
            save_category_to_disk(new_cat_name, pd.DataFrame())
            st.success(f"Đã thêm: {new_cat_name}")
            st.rerun()

    existing_cats = list(st.session_state['categories'].keys())
    if existing_cats:
        cat_to_delete = st.selectbox("Chọn danh mục cần xóa:", existing_cats, key="del_cat_select")
        if st.button("🗑️ Xóa danh mục"):
            if len(existing_cats) > 1:
                del st.session_state['categories'][cat_to_delete]
                delete_category_from_disk(cat_to_delete)
                st.success(f"Đã xóa: {cat_to_delete}")
                st.rerun()

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
if source_mode == "📦 Quản lý & Xem chi tiết Danh mục":
    st.markdown("""
        <div class="glow-card">
            <h3>📦 QUẢN LÝ VÀ XEM CHI TIẾT DANH MỤC CHUẨN BHYT</h3>
            <p style="color: #cbd5e1;">Nhập từ khóa tìm kiếm và bấm nút Tích chọn/Bỏ chọn để cấu hình dịch vụ kiểm tra mã máy. Cấu hình được lưu tự động lên hệ thống.</p>
        </div>
    """, unsafe_allow_html=True)
    
    cats_list = list(st.session_state['categories'].keys())
    
    selected_view_cat = st.radio(
        "📌 TÍCH CHỌN DANH MỤC CẦN XEM:",
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
            
            cat_name_cols = [c for c in cat_df.columns if 'TEN' in c.upper()]
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
                    if st.button("❌ Bỏ chọn tất cả danh mục", key=f"btn_clear_all_{selected_view_cat}"):
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
                st.warning("⚠️ Không tìm thấy cột tên dịch vụ (TEN...) trong danh mục này để tạo bộ lọc.")

            st.markdown("---")
            st.dataframe(cat_df, use_container_width=True)
        else:
            st.warning(f"⚠️ Danh mục **{selected_view_cat}** hiện đang trống. Hãy tải file lên từ thanh bên trái.")

else:
    sub_source_mode = st.radio("Chọn phương thức rà soát:", ["Quét tự động theo thư mục (Folder Scan)", "Tải lên từng tệp (File Uploader)"], horizontal=True)

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
                    with st.spinner("Đang tiến hành rà soát thư mục..."):
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
        uploaded_file = st.file_uploader("Tải lên tệp XML tổng hợp hoặc tệp XML thành phần", type=['xml'])
        if uploaded_file:
            file_results, is_valid, err_msg = process_xml_package(uploaded_file)
            
            single_summary = [{
                "Tên tệp": uploaded_file.name,
                "Trạng thái định dạng": "✅ Đúng định dạng (QĐ 4750/3176)" if is_valid else f"❌ Sai định dạng ({err_msg})",
                "Kiểm tra mã máy CLS": "Không có" if not is_valid else ("Đầy đủ mã máy" if not audit_cls_machine_code(file_results)[0] else f"Có thiếu ({len(audit_cls_machine_code(file_results)[1])} dịch vụ)"),
                "Chi tiết lỗi CLS": audit_cls_machine_code(file_results)[1] if is_valid else []
            }]
            st.session_state['folder_summary'] = pd.DataFrame(single_summary)
            st.session_state['folder_results'] = {uploaded_file.name: file_results}

    # --- HIỂN THỊ ĐẦY ĐỦ BẢNG TỔNG HỢP VÀ CHI TIẾT ---
    if 'folder_summary' in st.session_state:
        st.markdown("---")
        st.markdown("### 📊 Bảng tổng hợp kết quả rà soát")
        df_summary = st.session_state['folder_summary']
        
        display_df = df_summary[['Tên tệp', 'Trạng thái định dạng', 'Kiểm tra mã máy CLS']].copy()
        st.dataframe(display_df, use_container_width=True)
        
        valid_count = len(df_summary[df_summary['Trạng thái định dạng'].str.contains('✅')])
        invalid_count = len(df_summary[df_summary['Trạng thái định dạng'].str.contains('❌')])
        missing_machine_file_count = len(df_summary[df_summary['Kiểm tra mã máy CLS'].str.contains('Có thiếu')])
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Tổng số tệp quét", len(df_summary))
        col2.metric("Số tệp Đúng định dạng", valid_count)
        col3.metric("Số tệp Sai định dạng", invalid_count)
        col4.metric("🚨 Tệp thiếu mã máy", missing_machine_file_count)
        
        # --- THỐNG KÊ TỆP THIẾU MÃ MÁY KÈM NÚT CON MẮT (👁️) ĐỂ XEM CHI TIẾT VÀ TRỎ VỊ TRÍ ---
        if missing_machine_file_count > 0:
            st.markdown("---")
            st.markdown("#### 🚨 Danh sách các tệp thiếu mã máy (`MA_MAY`)")
            
            df_missing_files = df_summary[df_summary['Kiểm tra mã máy CLS'].str.contains('Có thiếu')]
            
            for idx, m_row in df_missing_files.iterrows():
                f_name = m_row['Tên tệp']
                f_details = m_row['Chi tiết lỗi CLS']
                
                col_info, col_eye = st.columns([5, 1])
                with col_info:
                    st.write(f"📄 **{f_name}** — ⚠️ Có **{len(f_details)}** dịch vụ thiếu mã máy.")
                with col_eye:
                    # Trạng thái ẩn/hiện chi tiết lưu trong session state
                    eye_key = f"show_detail_{f_name}"
                    if eye_key not in st.session_state:
                        st.session_state[eye_key] = False
                        
                    btn_label = "👁️ Xem" if not st.session_state[eye_key] else "👁️‍🗨️ Đóng"
                    if st.button(btn_label, key=f"btn_eye_{f_name}__"):
                        st.session_state[eye_key] = not st.session_state[eye_key]
                        st.rerun()
                
                # Nếu nút con mắt được kích hoạt (True), hiển thị bảng chi tiết và nút định vị
                if st.session_state.get(eye_key, False):
                    with st.container():
                        st.info(f"Chi tiết các dịch vụ thiếu mã máy trong tệp **{f_name}**:")
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
                st.warning(f"⚠️ Phát hiện **{len(row_info['Chi tiết lỗi CLS'])}** dịch vụ cận lâm sàng bị **thiếu mã máy (`MA_MAY`)**!")
                st.dataframe(pd.DataFrame(row_info['Chi tiết lỗi CLS']), use_container_width=True)
            else:
                st.success("✨ Không phát hiện lỗi thiếu mã máy CLS.")
                
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
                                st.markdown(f"#### 🔄 Cấu hình chọn **nhiều trường** đối chiếu với [{compare_cat_name}]:")
                                
                                col_xml_field, col_cat_field = st.columns(2)
                                
                                xml_cols = list(df_data.columns)
                                default_xml_cols = [c for c in xml_cols if 'MA_' in c.upper()]
                                if not default_xml_cols and xml_cols:
                                    default_xml_cols = [xml_cols[0]]
                                        
                                cat_cols_list = list(standard_catalog.columns)
                                default_cat_cols = [c for c in cat_cols_list if 'MA_' in c.upper()]
                                if not default_cat_cols and cat_cols_list:
                                    default_cat_cols = [cat_cols_list[0]]
                                        
                                with col_xml_field:
                                    match_cols = st.multiselect("Chọn các cột mã ở tệp XML:", xml_cols, default=default_xml_cols, key=f"xml_cols_{selected_file_name}_{selected_table}")
                                with col_cat_field:
                                    cat_cols = st.multiselect("Chọn các cột mã tương ứng trong danh mục chuẩn:", cat_cols_list, default=default_cat_cols, key=f"cat_cols_{compare_cat_name}_{selected_file_name}_{selected_table}")
                                
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
                                    
                                    target_services_config = st.session_state['machine_services_config'].get(compare_cat_name, [])
                                    
                                    if target_services_config:
                                        st.info(f"ℹ️ Đang áp dụng cấu hình chọn lọc kiểm tra mã máy từ danh mục **{compare_cat_name}** ({len(target_services_config)} dịch vụ được tích chọn).")
                                    
                                    missing_machine_filtered = []
                                    for idx, row in df_data.iterrows():
                                        ma_may = str(row.get('MA_MAY', row.get('ma_may', ''))).strip()
                                        ma_dv = str(row.get('MA_DICH_VU', row.get('MA_DVKT', row.get('MA_VАТ_TU', '')))).strip()
                                        ten_dv = str(row.get('TEN_DICH_VU', row.get('TEN_DVKT', row.get('TEN_VAT_TU', row.get('TEN_VАТ_TU', ''))))).strip()
                                        
                                        if not ma_dv and not ten_dv:
                                            continue
                                            
                                        if target_services_config and ten_dv not in target_services_config:
                                            continue
                                            
                                        if not ma_may or ma_may.lower() == 'nan':
                                            missing_machine_filtered.append({
                                                "Dòng số": idx + 1,
                                                "Mã dịch vụ": ma_dv,
                                                "Tên dịch vụ": ten_dv,
                                                "Trạng thái": row['So_sanh_danh_muc'],
                                                "Lỗi": "Thiếu mã máy (MA_MAY trống)"
                                            })

                                    if missing_machine_filtered:
                                        st.warning(f"⚠️ Tìm thấy **{len(missing_machine_filtered)}** dịch vụ theo cấu hình danh mục bị thiếu mã máy (`MA_MAY`)!")
                                        st.dataframe(pd.DataFrame(missing_machine_filtered), use_container_width=True)
                                    else:
                                        st.success("✨ Không phát hiện dịch vụ nào bị thiếu mã máy theo cấu hình của danh mục.")

                                    if target_rows_to_highlight:
                                        st.error(f"🎯 **Hệ thống đang trỏ đánh dấu tô đỏ** tại các dòng số: **{target_rows_to_highlight}** bị thiếu mã máy.")

                                    st.markdown("##### 📋 Bảng dữ liệu chi tiết sau đối chiếu (Được tô màu đánh dấu vị trí lỗi):")
                                    st.dataframe(df_data.style.apply(highlight_missing_machine, axis=1), use_container_width=True)
                                    
                                    df_invalid = df_data[df_data['So_sanh_danh_muc'].str.contains('❌')]
                                    invalid_cat_count = len(df_invalid)
                                    
                                    if invalid_cat_count > 0:
                                        st.warning(f"⚠️ Phát hiện {invalid_cat_count} bản ghi không khớp danh mục chuẩn.")
                                        
                                        csv_error_data = df_invalid.to_csv(index=False).encode('utf-8-sig')
                                        st.download_button(
                                            label="📥 Tải xuống danh sách bản ghi không khớp (CSV)",
                                            data=csv_error_data,
                                            file_name=f"Danh_sach_khong_khop_{selected_table}.csv",
                                            mime="text/csv"
                                        )
                                else:
                                    st.warning("⚠️ Số lượng cột chọn ở XML và Danh mục chuẩn phải bằng nhau để đối chiếu tương ứng.")
                                    if target_rows_to_highlight:
                                        st.error(f"🎯 **Hệ thống đang trỏ đánh dấu tô đỏ** tại các dòng số: **{target_rows_to_highlight}** bị thiếu mã máy.")
                                    st.dataframe(df_data.style.apply(highlight_missing_machine, axis=1), use_container_width=True)
                            else:
                                st.warning(f"Danh mục [{compare_cat_name}] hiện đang trống.")
                                if target_rows_to_highlight:
                                    st.error(f"🎯 **Hệ thống đang trỏ đánh dấu tô đỏ** tại các dòng số: **{target_rows_to_highlight}** bị thiếu mã máy.")
                                st.dataframe(df_data.style.apply(highlight_missing_machine, axis=1), use_container_width=True)
                        else:
                            if target_rows_to_highlight:
                                st.error(f"🎯 **Hệ thống đang trỏ đánh dấu tô đỏ** tại các dòng số: **{target_rows_to_highlight}** bị thiếu mã máy.")
                            st.dataframe(df_data.style.apply(highlight_missing_machine, axis=1), use_container_width=True)
                            
                    elif isinstance(data, dict):
                        st.dataframe(pd.DataFrame([{"Trường dữ liệu": k, "Giá trị": v} for k, v in data.items()]), use_container_width=True)