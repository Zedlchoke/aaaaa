import pandas as pd
import unidecode
import re
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from PIL import Image
import pdfplumber

# ==========================================
# CẤU HÌNH CỘT VÀ TỪ ĐIỂN
# ==========================================
COL_TEN_CTY = 'TENDTPN' 
COL_MA = 'MADTPN'     
COL_DIENGIAI = 'DIENGIAI'

ALIAS_WORDS = {r'\bbetong\b': 'be tong'}

ACTION_VERBS = ['chuyen tien', 'chuyen khoan', 'ck', 'thanh toan', 'tt', 'tra tien', 'ct']

BASE_STOP_WORDS = [
    r'\bchuyen khoan\b', r'\bck\b', r'\bthanh toan\b', r'\btt\b', 
    r'\brut sec\b', r'\bsec\b', r'\brut tien\b', r'\bnop tien\b', r'\bvao tk\b',
    r'\bphi dv\b', r'\bphi quan ly\b', r'\bsms banking\b', r'\bhoa don\b', r'\bhd\b',
    r'\bdot\b', r'\bcuoi\b', r'\btam ung\b', r'\bquyet toan\b', r'\bhdtc\b', r'\bhdkt\b',
    r'\bbe tong\b', r'\bbetong\b', r'\bly tam\b', r'\bthep\b', r'\bauto\b',
    r'\bcong ty\b', r'\bcty\b', r'\bctcp\b', r'\bco phan\b', r'\bcp\b',
    r'\btrach nhiem huu han\b', r'\btnhh\b', r'\bmtv\b', r'\bm t v\b', r'\bmot thanh vien\b',
    r'\bchi nhanh\b', r'\btap doan\b', r'\blien hop\b', r'\bco so\b', r'\bdoanh nghiep\b',
    r'\bviet nam\b', r'\bvn\b', r'\bva\b', r'\bgroup\b', r'\bholdings\b',
    r'\bthuong mai\b', r'\btm\b', r'\bdich vu\b', r'\bdv\b', r'\btmdv\b', r'\btmcp\b',
    r'\bsan xuat\b', r'\bsx\b', r'\bxuat nhap khau\b', r'\bxnk\b', r'\bdau tu\b',
    r'\bxay dung\b', r'\bxd\b', r'\bcong nghiep\b', r'\bcn\b', r'\bco khi\b',
    r'\bkim loai\b', r'\bngu kim\b', r'\bvat lieu\b', r'\bdanh bong\b',
    r'\bkhuon mau\b', r'\bgia cong\b', r'\bkho bai\b', r'\bbao ve\b',
    r'\bkhoa hoc\b', r'\bcong nghe\b', r'\bmoi truong\b', r'\bmachinery\b', r'\bmetal\b',
    r'\bnhom hop kim\b', r'\bnhom\b', r'\bhop kim\b', r'\bthiet bi dien\b', r'\bdien\b', r'\bthiet bi\b', 
    r'\bphat trien\b', r'\bky thuat\b', r'\btong hop\b', r'\bquoc te\b', r'\bche tao\b', r'\bvan tai\b',
    r'\bvat tu\b', r'\bphu lieu\b', r'\bnhua\b', r'\bbao bi\b', r'\btrang tri\b'
]

def apply_alias(text):
    for pattern, replacement in ALIAS_WORDS.items():
        text = re.sub(pattern, replacement, text)
    return text

def normalize_basic(text):
    if pd.isna(text): return ""
    t = unidecode.unidecode(str(text)).lower().strip()
    t = apply_alias(t) 
    t = re.sub(r'[^a-z0-9\s]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

def get_core_name(text, stop_words_list):
    t = normalize_basic(text)
    for _ in range(2):
        for w in stop_words_list:
            t = re.sub(w, ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

def format_excel_sheet(ws):
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            if cell.column in [2, 3, 4]: cell.alignment = Alignment(wrap_text=True, vertical="center")
            else: cell.alignment = Alignment(vertical="center", horizontal="center")

def parse_amt_to_float(val):
    if pd.isna(val): return 0.0
    # Xóa dấu phẩy và khoảng trắng, nhưng giữ lại dấu chấm thập phân và dấu trừ
    s = str(val).strip().replace(',', '').replace(' ', '')
    if not s: return 0.0
    try:
        return float(s)
    except:
        # Nếu có dính chữ, trích xuất con số hợp lệ đầu tiên
        m = re.search(r'-?\d+(\.\d+)?', s)
        return float(m.group(0)) if m else 0.0

def load_smart_ktsc(file_path):
    if not file_path or not os.path.exists(file_path): return []
    try:
        try:
            df = pd.read_excel(file_path, sheet_name='Smart_KTSC_OK', dtype=str)
        except:
            df = pd.read_excel(file_path, dtype=str)
            
        items = df.to_dict('records')
        clean_items = []
        for row in items:
            clean_items.append({
                'SO_HD': str(row.get('SO_HD', '')).strip(),
                'MATHANG': str(row.get('MATHANG', '')).strip(),
                'TTVND': parse_amt_to_float(row.get('TTVND', 0)),
                'TTVND_TT': parse_amt_to_float(row.get('TTVND_TT', 0)), # BỔ SUNG CỘT SAU THUẾ
                COL_MA: str(row.get('MAKH', '')).strip(),
                COL_TEN_CTY: str(row.get('TENKH', '')).strip()
            })
        return clean_items
    except Exception as e:
        print(f"Lỗi đọc file Mua/Bán: {e}")
        return []

# ==========================================
# LÕI XỬ LÝ ĐỐI SOÁT NGÂN HÀNG TOÀN DIỆN
# ==========================================
def process_bank_data(file_saoke, file_master, path_save_doichieu, path_save_saokemoi, user_stop_str, file_muavao=None, file_banra=None, progress_callback=None):
    dynamic_stops = BASE_STOP_WORDS.copy()
    user_stops = [w.strip() for w in user_stop_str.split(',') if w.strip()]
    for w in user_stops:
        w_norm = normalize_basic(w)
        if w_norm: dynamic_stops.append(r'\b' + re.escape(w_norm) + r'\b')

    if progress_callback: progress_callback(5, "Đang đọc file Excel...")
    df_bank = pd.read_excel(file_saoke)
    df_master = pd.read_excel(file_master)
    master_list = df_master.to_dict('records')
    
    if progress_callback: progress_callback(10, "Đang chuẩn hóa danh sách Master chính...")
    for item in master_list:
        item['norm_core'] = get_core_name(item.get(COL_TEN_CTY, ""), dynamic_stops)

    # Nạp file Mua Vào / Bán Ra
    if progress_callback: progress_callback(12, "Đang nạp file Mua Vào / Bán Ra...")
    list_muavao = load_smart_ktsc(file_muavao)
    list_banra = load_smart_ktsc(file_banra)

    all_matches = []
    total_rows = len(df_bank)
    
    for idx, row in df_bank.iterrows():
        if progress_callback and idx % 20 == 0:
            progress = 15 + (idx / total_rows) * 65
            progress_callback(progress, f"Đang đối soát dòng {idx+1}/{total_rows}...")
            
        thu_tu_dong = idx + 2
        diengiai_goc = str(row.get(COL_DIENGIAI, ""))
        if pd.isna(diengiai_goc) or diengiai_goc.strip() == "" or diengiai_goc.lower().strip() == "nan":
            all_matches.append({
                "THỨ TỰ DÒNG GỐC": thu_tu_dong, "DIỄN GIẢI GỐC": "", 
                "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN": "", 
                "MÃ ĐỐI TƯỢNG PHÁP NHÂN": "", "CÁCH MATCH": "", "SCORE_NUM": 0
            })
            continue

        diengiai_norm = normalize_basic(diengiai_goc)
        diengiai_cleaned = get_core_name(diengiai_goc, dynamic_stops)
        diengiai_nospace = diengiai_norm.replace(" ", "") 
        matches_for_row = []
        
        # --- BƯỚC 1: SO KHỚP TRÊN FILE MASTER CHÍNH (P1, P2, P3, P4) ---
        for master_item in master_list:
            core = master_item['norm_core']
            if len(core) < 3: continue
            core_words = core.split()
            core_nospace = core.replace(" ", "")
            
            pattern1 = r'\b' + r'\s*'.join(map(re.escape, core_words)) + r'\b'
            if re.search(pattern1, diengiai_norm):
                matches_for_row.append((1, len(core), master_item, core.upper()))
                continue
                
            if len(diengiai_cleaned) >= 4:
                dg_words = diengiai_cleaned.split()
                if dg_words:
                    pattern2 = r'\b' + r'\s+'.join(map(re.escape, dg_words)) + r'\b'
                    if re.search(pattern2, core):
                        matches_for_row.append((2, -len(core), master_item, diengiai_cleaned.upper()))
                        continue
                        
            if len(core_words) >= 4:
                acronym = "".join(w[0] for w in core_words)
                if len(acronym) >= 4:
                    verb_group = '|'.join(ACTION_VERBS)
                    pattern3 = r'\b(?:' + verb_group + r')\b.*?\b' + re.escape(acronym) + r'\b'
                    if re.search(pattern3, diengiai_norm):
                        matches_for_row.append((3, len(acronym), master_item, acronym.upper()))
                        continue

            # P4: Khớp dính chữ hoặc Khớp Cụm Từ Lõi (Chunk Matching)
            p4_match = False
            for i in range(len(core_words)):
                for j in range(i+1, len(core_words)+1):
                    chunk = "".join(core_words[i:j])
                    # Chấp nhận cụm ghép >= 5 ký tự (nếu gồm từ 2 chữ trở lên) hoặc cụm bất kỳ >= 8 ký tự
                    if (j - i >= 2 and len(chunk) >= 5) or len(chunk) >= 8:
                        if chunk in diengiai_nospace:
                            matches_for_row.append((4, len(chunk), master_item, chunk.upper()))
                            p4_match = True
                            break
                if p4_match: break
            if p4_match: continue

        # --- BƯỚC 2: NẾU THẤT BẠI, QUÉT QUA HÓA ĐƠN/HỢP ĐỒNG/SỐ TIỀN (P5, P6, P7) ---
        if not matches_for_row and (list_muavao or list_banra):
            
            # --- KIỂM TRA QUY TẮC VÀNG TTVND >= 5 TRIỆU ---
            amt_val = parse_amt_to_float(row.get('TTVND', 0))
            if amt_val == 0: 
                amt_val = parse_amt_to_float(row.get('TTVND_TT', 0))
                
            if amt_val >= 5000000:
                target_list = list_banra + list_muavao
                    
                if target_list:
                    invoices_found = []
                    contracts_found = []
                    
                    # Chuẩn hóa diễn giải: Xóa dấu, đưa về chữ thường, chỉ giữ lại chữ, số, khoảng trắng, / và -
                    text_ext = unidecode.unidecode(str(diengiai_goc)).lower()
                    text_ext = re.sub(r'[^a-z0-9\s/\-]', ' ', text_ext)
                    text_ext = re.sub(r'\s+', ' ', text_ext).strip()
                    
                    # Quy tắc 1: Ghi rõ "Hóa đơn"
                    invoices_found.extend(re.findall(r'\bhoa don\s*(?:so\s*)?(\d+)\b', text_ext))
                    
                    # Quy tắc 2: Ghi rõ "Hợp đồng"
                    explicit_contracts = re.findall(r'\bhop dong\s*(?:so\s*)?([a-z0-9/\-]+)\b', text_ext)
                    contracts_found.extend([c for c in explicit_contracts if len(c) >= 4])
                    # Xử lý thêm ca khách gõ sai có khoảng trắng (VD: "hop dong 0 8042026/HDTM")
                    explicit_contracts_spaced = re.findall(r'\bhop dong\s*(?:so\s*)?\d+\s+([a-z0-9/\-]+)\b', text_ext)
                    contracts_found.extend([c for c in explicit_contracts_spaced if '/' in c or '-' in c])
                    
                    # Quy tắc 3: Ghi tắt "HD" - Bộ lọc phân luồng thông minh
                    hd_matches = re.findall(r'\bhd\s*(?:so\s*)?([a-z0-9/\-]+)\b', text_ext)
                    for val in hd_matches:
                        if '/' in val or '-' in val:
                            if len(val) >= 4:
                                contracts_found.append(val)
                        elif val.isdigit():
                            invoices_found.append(val)
                            
                    # Xóa trùng lặp
                    invoices_found = list(set(invoices_found))
                    contracts_found = list(set(contracts_found))
                    
                    # P5: Đối chiếu Số Hóa Đơn
                    for inv in invoices_found:
                        inv_clean = inv.lstrip('0')
                        if not inv_clean: inv_clean = '0'
                        
                        for item in target_list:
                            item_hd_clean = str(item.get('SO_HD', '')).strip().lstrip('0')
                            if not item_hd_clean: item_hd_clean = '0'
                            
                            if item_hd_clean == inv_clean and inv_clean != '0' and item_hd_clean != '0':
                                matches_for_row.append((5, len(inv), item, f"SỐ HĐ: {inv}"))
                                break
                        if matches_for_row: break
                    
                    # P6: Đối chiếu Hợp Đồng (Tìm chuỗi nằm trong cột mặt hàng)
                    if not matches_for_row:
                        for contract in contracts_found:
                            contract_upper = contract.upper()
                            contract_clean = re.sub(r'[^A-Z0-9]', '', contract_upper) 
                            parts = re.split(r'[/_.-]', contract_upper)
                            valid_parts = [p for p in parts if len(p) >= 5] 
                            
                            for item in target_list:
                                mathang_upper = str(item.get('MATHANG', '')).upper()
                                mathang_clean = re.sub(r'[^A-Z0-9]', '', mathang_upper)
                                
                                matched = False
                                if len(contract_clean) >= 5 and contract_clean in mathang_clean:
                                    matched = True
                                elif any(vp in mathang_upper for vp in valid_parts):
                                    matched = True
                                    
                                if matched:
                                    matches_for_row.append((6, len(contract), item, f"HỢP ĐỒNG: {contract}"))
                                    break
                            if matches_for_row: break
                    
                    # P7: Đối chiếu Số Tiền (Duy nhất & Chống sai số làm tròn)
                    if not matches_for_row:
                        matched_companies = set()
                        best_item = None
                        for item in target_list:
                            item_ttvnd = parse_amt_to_float(item.get('TTVND', 0))
                            item_ttvnd_tt = parse_amt_to_float(item.get('TTVND_TT', 0))
                            
                            if abs(item_ttvnd - amt_val) < 1.0 or abs(item_ttvnd_tt - amt_val) < 1.0:
                                matched_companies.add(item.get(COL_MA))
                                best_item = item
                                
                        if len(matched_companies) == 1:
                            matches_for_row.append((7, 0, best_item, f"SỐ TIỀN: {amt_val:,.0f}"))
        # ==========================================
        # GHI NHẬN KẾT QUẢ ĐỐI CHIẾU
        # ==========================================
        if matches_for_row:
            matches_for_row.sort(key=lambda x: (x[0], -x[1]))
            best_match = matches_for_row[0]
            
            p_level = best_match[0]
            match_value = str(best_match[3])
            
            # Phân loại giá trị quét được vào đúng cột dựa trên cấp độ P
            ten_quet = match_value if p_level <= 4 else ""
            hd_quet = match_value.replace("SỐ HĐ: ", "") if p_level == 5 else ""
            hopdong_quet = match_value.replace("HỢP ĐỒNG: ", "") if p_level == 6 else ""
            sotien_quet = match_value.replace("SỐ TIỀN: ", "") if p_level == 7 else ""

            all_matches.append({
                "THỨ TỰ DÒNG GỐC": thu_tu_dong, 
                "DIỄN GIẢI GỐC": diengiai_goc,
                "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI": ten_quet,
                "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI": hd_quet,
                "HỢP ĐỒNG QUÉT ĐƯỢC TRONG DIỄN GIẢI": hopdong_quet,
                "SỐ TIỀN QUÉT ĐƯỢC TRONG DIỄN GIẢI": sotien_quet,
                "TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN": best_match[2].get(COL_TEN_CTY, ""),
                "MÃ ĐỐI TƯỢNG PHÁP NHÂN": best_match[2].get(COL_MA, ""),
                "CÁCH MATCH": f"P{p_level}", 
                "SCORE_NUM": 100 
            })
        else:
            all_matches.append({
                "THỨ TỰ DÒNG GỐC": thu_tu_dong, 
                "DIỄN GIẢI GỐC": diengiai_goc, 
                "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", 
                "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "",
                "HỢP ĐỒNG QUÉT ĐƯỢC TRONG DIỄN GIẢI": "",
                "SỐ TIỀN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "",
                "TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN": "", 
                "MÃ ĐỐI TƯỢNG PHÁP NHÂN": "", 
                "CÁCH MATCH": "", 
                "SCORE_NUM": 0
            })

    # ==========================================
    # TẠO FILE EXCEL KẾT QUẢ ĐỐI CHIẾU
    # ==========================================
    if progress_callback: progress_callback(85, "Đang tạo file Kết Quả Đối Chiếu...")
    wb_doichieu = Workbook()
    ws_doichieu = wb_doichieu.active
    ws_doichieu.title = "Ket Qua Match"
    
    # Cập nhật Header mới với các cột được bổ sung
    headers = [
        "THỨ TỰ DÒNG GỐC", 
        "DIỄN GIẢI GỐC", 
        "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI", 
        "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI",
        "HỢP ĐỒNG QUÉT ĐƯỢC TRONG DIỄN GIẢI",
        "SỐ TIỀN QUÉT ĐƯỢC TRONG DIỄN GIẢI",
        "TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN", 
        "MÃ ĐỐI TƯỢNG PHÁP NHÂN", 
        "CÁCH MATCH"
    ]
    ws_doichieu.append(headers)
    
    best_match_dict = {} 
    for match in all_matches:
        ws_doichieu.append([match.get(h, "") for h in headers])
        if match["SCORE_NUM"] == 100: 
            best_match_dict[match["THỨ TỰ DÒNG GỐC"]] = {
                "MA": match["MÃ ĐỐI TƯỢNG PHÁP NHÂN"], 
                "TEN": match["TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN"],
                "CACH_MATCH": match["CÁCH MATCH"]
            }
            
    format_excel_sheet(ws_doichieu)
    wb_doichieu.save(path_save_doichieu)

    if progress_callback: progress_callback(90, "Đang cập nhật lại file Sao Kê gốc...")
    wb_saoke = load_workbook(file_saoke)
    ws_saoke = wb_saoke.active
    yellow_fill = PatternFill(start_color="FFFF00", fill_type="solid")

    headers_dict = {str(cell.value).strip().upper(): cell.column for cell in ws_saoke[1] if cell.value}
    col_tkno = headers_dict.get('TKNO', 5)
    col_tkco = headers_dict.get('TKCO', 7)
    col_madtpnno = headers_dict.get('MADTPNNO', 6)
    col_madtpnco = headers_dict.get('MADTPNCO', 8)
    col_tenkh = headers_dict.get('TENKH', 11)
    col_ghichu = headers_dict.get('GHICHU')

    def is_valid_cell(val):
        return pd.notna(val) and str(val).strip() != "" and str(val).lower().strip() != "nan"

    for r in range(2, ws_saoke.max_row + 1):
        if r in best_match_dict:
            ma = best_match_dict[r]["MA"]
            ten = best_match_dict[r]["TEN"]
            cach_match = best_match_dict[r]["CACH_MATCH"]

            tkco_val = str(ws_saoke.cell(row=r, column=col_tkco).value).strip() if is_valid_cell(ws_saoke.cell(row=r, column=col_tkco).value) else ""
            tkno_val = str(ws_saoke.cell(row=r, column=col_tkno).value).strip() if is_valid_cell(ws_saoke.cell(row=r, column=col_tkno).value) else ""

            if tkco_val == "1121" or (tkco_val != "" and tkno_val == ""):
                ws_saoke.cell(row=r, column=col_madtpnno).value = ma
            elif tkno_val == "1121" or (tkno_val != "" and tkco_val == ""):
                ws_saoke.cell(row=r, column=col_madtpnco).value = ma
            else:
                ws_saoke.cell(row=r, column=col_madtpnno).value = ma
            
            ws_saoke.cell(row=r, column=col_tenkh).value = ten
            
            if col_ghichu: 
                ws_saoke.cell(row=r, column=col_ghichu).value = cach_match
        else:
            for c in range(1, ws_saoke.max_column + 1): 
                ws_saoke.cell(row=r, column=c).fill = yellow_fill

    if progress_callback: progress_callback(98, "Đang lưu file...")
    wb_saoke.save(path_save_saokemoi)
    if progress_callback: progress_callback(100, "Hoàn tất thành công!")

# ==========================================
# HÀM CẬP NHẬT FILE SỬA TAY (TAB 3)
# ==========================================
def process_update_saoke(file_sk_old, file_dc_edited, path_save, progress_callback=None):
    if progress_callback: progress_callback(10, "Đang đọc file đối chiếu chỉnh sửa...")
    df_dc = pd.read_excel(file_dc_edited)
    update_map = {int(row["THỨ TỰ DÒNG GỐC"]): {"MA": str(row["MÃ ĐỐI TƯỢNG PHÁP NHÂN"]).strip(), "TEN": str(row.get("TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN", "")).strip()} 
                  for _, row in df_dc.iterrows() if pd.notna(row.get("THỨ TỰ DÒNG GỐC")) and pd.notna(row.get("MÃ ĐỐI TƯỢNG PHÁP NHÂN")) and str(row.get("MÃ ĐỐI TƯỢNG PHÁP NHÂN")).strip()}
    
    if progress_callback: progress_callback(40, "Đang mở file Sao kê gốc...")
    wb = load_workbook(file_sk_old)
    ws = wb.active
    no_fill = PatternFill(fill_type=None)

    total_rows = ws.max_row
    for r in range(2, total_rows + 1):
        if r in update_map:
            ma, ten = update_map[r]["MA"], update_map[r]["TEN"]
            if pd.notna(ws.cell(row=r, column=7).value) and str(ws.cell(row=r, column=7).value).strip() != "": ws.cell(row=r, column=6).value = ma
            if pd.notna(ws.cell(row=r, column=5).value) and str(ws.cell(row=r, column=5).value).strip() != "": ws.cell(row=r, column=8).value = ma
            ws.cell(row=r, column=11).value = ten
            for c in range(1, ws.max_column + 1): ws.cell(row=r, column=c).fill = no_fill
            
    if progress_callback: progress_callback(90, "Đang lưu file cuối cùng...")
    wb.save(path_save)
    if progress_callback: progress_callback(100, "Cập nhật thành công!")

# ==========================================
# GIAO DIỆN HỆ THỐNG ĐA TÍNH NĂNG (Tkinter)
# ==========================================
class AppGomNghiepVu:
    def __init__(self, root):
        self.root = root
        self.root.title("Hệ Thống Xử Lý Kế Toán & Ngân Hàng Toàn Diện ETAX Hồ Chí Minh")
        self.root.geometry("780x670")
        
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.tab4 = ttk.Frame(self.notebook) 
        self.tab3 = ttk.Frame(self.notebook) 
        self.tab2 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text="  1. Đối Soát (Excel)  ")
        self.notebook.add(self.tab4, text="  2. Đọc PDF & Xử Lý Liên Hoàn  ")
        self.notebook.add(self.tab3, text="  3. Cập Nhật Bổ Sung  ")
        self.notebook.add(self.tab2, text="  4. Scan Hàng Loạt Ảnh  ")
        
        self.setup_tab1_interface()
        self.setup_tab4_interface()
        self.setup_tab3_interface()
        self.setup_tab2_interface()

    # --------------- TAB 1: ĐỐI SOÁT NGÂN HÀNG ---------------
    def setup_tab1_interface(self):
        self.file_saoke_path = self.file_master_path = self.file_muavao_path = self.file_banra_path = None
        tk.Label(self.tab1, text="PHẦN MỀM ĐỐI CHIẾU MÃ PHÁP NHÂN", font=("Arial", 14, "bold"), fg="#4F81BD").pack(pady=10)
        
        f_files = tk.Frame(self.tab1)
        f_files.pack(fill="x", padx=20, pady=5)
        
        tk.Button(f_files, text="1. Chọn File Sao Kê (Excel)", command=lambda: self.chon_file("sk1"), width=25).grid(row=0, column=0, pady=5, padx=5)
        self.lbl_saoke = tk.Label(f_files, text="Chưa chọn file (Bắt buộc)...", fg="gray")
        self.lbl_saoke.grid(row=0, column=1, sticky="w")
        
        tk.Button(f_files, text="2. Chọn Master MADTPN", command=lambda: self.chon_file("mst"), width=25).grid(row=1, column=0, pady=5, padx=5)
        self.lbl_master = tk.Label(f_files, text="Chưa chọn file (Bắt buộc)...", fg="gray")
        self.lbl_master.grid(row=1, column=1, sticky="w")
        
        tk.Button(f_files, text="3. Chọn File MUA VÀO", command=lambda: self.chon_file("muavao"), width=25).grid(row=2, column=0, pady=5, padx=5)
        self.lbl_muavao = tk.Label(f_files, text="Chưa chọn file (Không bắt buộc)...", fg="gray")
        self.lbl_muavao.grid(row=2, column=1, sticky="w")

        tk.Button(f_files, text="4. Chọn File BÁN RA", command=lambda: self.chon_file("banra"), width=25).grid(row=3, column=0, pady=5, padx=5)
        self.lbl_banra = tk.Label(f_files, text="Chưa chọn file (Không bắt buộc)...", fg="gray")
        self.lbl_banra.grid(row=3, column=1, sticky="w")
        
        f_stop = tk.Frame(self.tab1)
        f_stop.pack(fill="x", padx=20, pady=10)
        tk.Label(f_stop, text="Nhập Tên Chủ TK cần loại trừ khỏi diễn giải:", font=("Arial", 10, "bold")).pack(anchor="w")
        self.entry_stop_words = tk.Entry(f_stop, width=70, font=("Arial", 11))
        self.entry_stop_words.pack(anchor="w", ipady=3)

        self.btn_run_t1 = tk.Button(self.tab1, text="BẮT ĐẦU ĐỐI SOÁT", bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), command=self.t1_chay, height=2)
        self.btn_run_t1.pack(fill="x", padx=25, pady=10)
        
        self.progress_var_t1 = tk.DoubleVar()
        self.progressbar_t1 = ttk.Progressbar(self.tab1, variable=self.progress_var_t1, maximum=100)
        self.progressbar_t1.pack(fill="x", padx=25, pady=5)
        self.lbl_status_t1 = tk.Label(self.tab1, text="Sẵn sàng...", font=("Arial", 9, "italic"))
        self.lbl_status_t1.pack()

    def t1_update_progress(self, percent, text):
        self.root.after(0, lambda: self.progress_var_t1.set(percent))
        self.root.after(0, lambda: self.lbl_status_t1.config(text=text))

    def t1_chay(self):
        if not self.file_saoke_path or not self.file_master_path: 
            return messagebox.showerror("Lỗi", "Vui lòng chọn đủ File Sao kê và File data MADTPN!")
            
        p_dc = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="KetQua_DoiChieu.xlsx")
        if not p_dc: return
        p_sk = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="SaoKe_DaCapNhat.xlsx")
        if not p_sk: return
        
        user_stops = self.entry_stop_words.get()
        self.btn_run_t1.config(state="disabled")
        self.progress_var_t1.set(0)
        
        def task():
            try:
                process_bank_data(
                    file_saoke=self.file_saoke_path, 
                    file_master=self.file_master_path, 
                    path_save_doichieu=p_dc, 
                    path_save_saokemoi=p_sk, 
                    user_stop_str=user_stops, 
                    file_muavao=self.file_muavao_path, 
                    file_banra=self.file_banra_path,
                    progress_callback=self.t1_update_progress
                )
                self.root.after(0, lambda: messagebox.showinfo("Xong", "Đã đối soát thành công!"))
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Lỗi", f"Có lỗi xảy ra:\n{str(e)}"))
            finally:
                self.root.after(0, lambda: self.btn_run_t1.config(state="normal"))
                
        threading.Thread(target=task, daemon=True).start()

    # --------------- TAB 4: ĐỌC PDF & XỬ LÝ LIÊN HOÀN ---------------
    def setup_tab4_interface(self):
        self.file_pdf_paths = [] 
        self.file_master_t4_path = self.file_template_path = None
        tk.Label(self.tab4, text="CHUYỂN ĐỔI BATCH PDF VÀ ĐỐI SOÁT TỰ ĐỘNG", font=("Arial", 14, "bold"), fg="#D32F2F").pack(pady=10)
        
        f_files = tk.Frame(self.tab4); f_files.pack(fill="x", padx=20, pady=5)
        
        tk.Button(f_files, text="1. Chọn Nhiều File PDF", command=lambda: self.chon_file("pdf"), width=25).grid(row=0, column=0, pady=5, padx=5)
        self.lbl_pdf = tk.Label(f_files, text="Chưa chọn file PDF...", fg="gray"); self.lbl_pdf.grid(row=0, column=1, sticky="w")

        tk.Button(f_files, text="2. Chọn File Excel Mẫu", command=lambda: self.chon_file("tmpl"), width=25).grid(row=1, column=0, pady=5, padx=5)
        self.lbl_tmpl = tk.Label(f_files, text="Chưa chọn file mẫu...", fg="gray"); self.lbl_tmpl.grid(row=1, column=1, sticky="w")
        
        tk.Button(f_files, text="3. Chọn Master Data (Excel)", command=lambda: self.chon_file("mst4"), width=25).grid(row=2, column=0, pady=5, padx=5)
        self.lbl_master_t4 = tk.Label(f_files, text="Chưa chọn file Master...", fg="gray"); self.lbl_master_t4.grid(row=2, column=1, sticky="w")

        f_opts = tk.Frame(self.tab4); f_opts.pack(fill="x", padx=20, pady=5)
        tk.Label(f_opts, text="Ngân Hàng PDF:").grid(row=0, column=0, sticky="w")
        self.combo_bank = ttk.Combobox(f_opts, values=["MB Bank", "ACB Bank"], state="readonly")
        self.combo_bank.current(0)
        self.combo_bank.grid(row=0, column=1, padx=10)

        self.btn_run_t4 = tk.Button(self.tab4, text="BẮT ĐẦU CHẠY LIÊN HOÀN", bg="#D32F2F", fg="white", font=("Arial", 11, "bold"), command=self.t4_chay, height=2)
        self.btn_run_t4.pack(fill="x", padx=25, pady=10)
        
        self.progress_var_t4 = tk.DoubleVar()
        self.progressbar_t4 = ttk.Progressbar(self.tab4, variable=self.progress_var_t4, maximum=100)
        self.progressbar_t4.pack(fill="x", padx=25, pady=5)
        self.lbl_status_t4 = tk.Label(self.tab4, text="Sẵn sàng...", font=("Arial", 9, "italic"))
        self.lbl_status_t4.pack()

    def t4_update_progress(self, percent, text):
        self.root.after(0, lambda: self.progress_var_t4.set(percent))
        self.root.after(0, lambda: self.lbl_status_t4.config(text=text))

    def is_garbage_or_header(self, row_text):
        t = str(row_text).upper().strip()
        if not t: return True
        if "NGÀY GIAO DỊCH" in t or "PHÁT SINH NỢ" in t or "PHÁT SINH CÓ" in t or "GHI NỢ" in t or "GHI CÓ" in t: return True
        if "STT" in t and "NỘI DUNG" in t: return True
        if "SỐ DƯ ĐẦU" in t or "TỔNG TIỀN RÚT" in t or "TỔNG TIỀN GỬI" in t or "SỐ DƯ CUỐI" in t or "THỜI GIAN SAO KÊ" in t or "CHỦ TÀI KHOẢN" in t: return True
        return False

    def extract_pdf_to_template(self, pdf_paths, template_path, temp_excel_path):
        bank_name = self.combo_bank.get()
        extracted_data = []
        
        total_pdfs = len(pdf_paths)
        for p_idx, pdf_path in enumerate(pdf_paths):
            self.t4_update_progress(5 + (p_idx / total_pdfs) * 30, f"Đang bóc tách PDF {p_idx+1}/{total_pdfs} ({bank_name})...")
            
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    if bank_name == "MB Bank":
                        tables = page.extract_tables()
                    else:
                        tables = page.extract_tables({
                            "vertical_strategy": "text", 
                            "horizontal_strategy": "text",
                            "intersection_y_tolerance": 15
                        })
                        
                    for table in tables:
                        for row in table:
                            cleaned_row = [re.sub(r'\s+', ' ', str(cell)).strip() if cell else "" for cell in row]
                            row_text_str = " ".join(cleaned_row)
                            if any(cleaned_row) and not self.is_garbage_or_header(row_text_str):
                                extracted_data.append(cleaned_row)
        
        self.t4_update_progress(45, "Đang xử lý dồn dòng thông minh...")
        merged_rows = []
        current_row_data = []
        for row in extracted_data:
            text_str = " ".join(row).strip()
            is_new_transaction = False
            
            if re.search(r'\b\d{2}/\d{2}/\d{4}\b', text_str):
                is_new_transaction = True
            elif bank_name == "MB Bank" and re.match(r'^\d+$', str(row[0]).strip()):
                is_new_transaction = True
                
            if is_new_transaction:
                if current_row_data: merged_rows.append(current_row_data)
                current_row_data = list(row)
            else:
                if current_row_data:
                    for i in range(max(len(row), len(current_row_data))):
                        val_add = str(row[i]).strip() if i < len(row) and row[i] else ""
                        if i < len(current_row_data):
                            if val_add: current_row_data[i] = str(current_row_data[i]) + " " + val_add
                        else: current_row_data.append(val_add)
                else: current_row_data = list(row)
        if current_row_data: merged_rows.append(current_row_data)

        self.t4_update_progress(60, "Đang ánh xạ dữ liệu vào Excel mẫu...")
        wb = load_workbook(template_path)
        ws = wb.active
        
        headers_dict = {str(cell.value).strip().upper(): cell.column for cell in ws[1] if cell.value}
        col_lctg = headers_dict.get('LCTG')
        col_ngayct = headers_dict.get('NGAYCT')
        col_soct = headers_dict.get('SOCT')
        col_diengiai = headers_dict.get('DIENGIAI')
        col_tkno = headers_dict.get('TKNO')
        col_tkco = headers_dict.get('TKCO')
        col_ttvnd = headers_dict.get('TTVND')
        col_ttvnd_tt = headers_dict.get('TTVND_TT')
        col_nghiepvu = headers_dict.get('ID_NGHIEPVU')
        
        if not col_diengiai:
            col_diengiai = ws.max_column + 1
            ws.cell(row=1, column=col_diengiai).value = 'DIENGIAI'

        start_row = ws.max_row + 1
        for row_data in merged_rows:
            full_text = " ".join([str(x) for x in row_data if x]).strip()
            full_text = re.sub(r'\s+', ' ', full_text)
            
            ngay_ct = ""
            soct = ""
            ps_no = 0.0
            ps_co = 0.0
            diengiai_text = ""
            
            def parse_amt(val):
                v = re.sub(r'[^\d]', '', str(val))
                return float(v) if v else 0.0

            if bank_name == "MB Bank":
                dates = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', full_text)
                target_date_str = None
                if len(dates) >= 2:
                    ngay_ct = dates[1]
                    target_date_str = dates[1]
                elif len(dates) == 1:
                    ngay_ct = dates[0]
                    target_date_str = dates[0]

                if target_date_str:
                    parts = full_text.split(target_date_str)
                    if len(dates) >= 2 and dates[0] == target_date_str and len(parts) >= 3:
                        after_date_str = parts[2]
                    elif len(parts) >= 2:
                        after_date_str = parts[1]
                    else:
                        after_date_str = parts[-1]
                        
                    amounts = re.findall(r'(?<!\S)\d+(?:,\d{3})*(?!\S)', after_date_str)
                    if len(amounts) >= 2:
                        ps_no = parse_amt(amounts[0])
                        ps_co = parse_amt(amounts[1])
                    else:
                        ps_no = parse_amt(row_data[3]) if len(row_data) > 3 else 0.0
                        ps_co = parse_amt(row_data[4]) if len(row_data) > 4 else 0.0
                else:
                    ps_no = parse_amt(row_data[3]) if len(row_data) > 3 else 0.0
                    ps_co = parse_amt(row_data[4]) if len(row_data) > 4 else 0.0
                
                if len(row_data) > 7: diengiai_text = str(row_data[-2]).strip()
                if len(row_data) > 8: soct = str(row_data[-1]).strip()
                    
            elif bank_name == "ACB Bank":
                if len(row_data) > 0:
                    date_match = re.search(r'\b\d{2}/\d{2}/\d{4}\b', str(row_data[0]))
                    if date_match: ngay_ct = date_match.group(0)
                    else:
                        date_fallback = re.search(r'\b\d{2}/\d{2}/\d{4}\b', full_text)
                        if date_fallback: ngay_ct = date_fallback.group(0)

                if len(row_data) >= 6:
                    soct = str(row_data[2]).strip()
                    diengiai_text = str(row_data[3]).strip()
                    ps_no = parse_amt(row_data[-3]) if len(row_data) >= 3 else 0.0
                    ps_co = parse_amt(row_data[-2]) if len(row_data) >= 2 else 0.0
                else:
                    diengiai_text = full_text

            if not diengiai_text: diengiai_text = full_text

            if col_lctg: ws.cell(row=start_row, column=col_lctg).value = "CTNH"
            if col_ngayct: ws.cell(row=start_row, column=col_ngayct).value = ngay_ct
            if col_soct: ws.cell(row=start_row, column=col_soct).value = soct
            if col_diengiai: ws.cell(row=start_row, column=col_diengiai).value = diengiai_text
            if col_nghiepvu: ws.cell(row=start_row, column=col_nghiepvu).value = "TIENHANG"
            
            tien = ps_no if ps_no > 0 else ps_co
            if tien > 0:
                if col_ttvnd: ws.cell(row=start_row, column=col_ttvnd).value = tien
                if col_ttvnd_tt: ws.cell(row=start_row, column=col_ttvnd_tt).value = tien
            
            if ps_no > 0: 
                if col_tkco: ws.cell(row=start_row, column=col_tkco).value = 1121
            elif ps_co > 0: 
                if col_tkno: ws.cell(row=start_row, column=col_tkno).value = 1121

            ws.row_dimensions[start_row].height = 35
            for c in range(1, ws.max_column + 1):
                ws.cell(row=start_row, column=c).alignment = Alignment(wrap_text=True, vertical="center")
            
            start_row += 1
            
        wb.save(temp_excel_path)
        return temp_excel_path

    def t4_chay(self):
        if not self.file_pdf_paths or not self.file_master_t4_path or not self.file_template_path: 
            return messagebox.showerror("Lỗi", "Vui lòng chọn đủ File PDF, File Excel Mẫu và Master Data!")
            
        p_dc = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="KetQua_DoiChieu_Auto.xlsx", title="Lưu Kết Quả Đối Chiếu")
        if not p_dc: return
        p_sk = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="SaoKe_PDF_DaCapNhat.xlsx", title="Lưu Sao Kê Cuối")
        if not p_sk: return
        
        self.btn_run_t4.config(state="disabled")
        self.progress_var_t4.set(0)
        
        def task():
            try:
                temp_excel = "Temp_Extracted_SaoKe.xlsx"
                self.extract_pdf_to_template(self.file_pdf_paths, self.file_template_path, temp_excel)
                self.t4_update_progress(80, "Bắt đầu đối soát pháp nhân...")
                
                process_bank_data(
                    file_saoke=temp_excel, 
                    file_master=self.file_master_t4_path, 
                    path_save_doichieu=p_dc, 
                    path_save_saokemoi=p_sk, 
                    user_stop_str="", 
                    file_muavao=None, 
                    file_banra=None, 
                    progress_callback=self.t4_update_progress
                )
                
                if os.path.exists(temp_excel): os.remove(temp_excel)
                self.root.after(0, lambda: messagebox.showinfo("Thành công", "Đã chạy liên hoàn thành công PDF -> Excel Mẫu -> Đối Soát!"))
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Lỗi", f"Có lỗi xảy ra:\n{str(e)}"))
            finally:
                self.root.after(0, lambda: self.btn_run_t4.config(state="normal"))
                
        threading.Thread(target=task, daemon=True).start()

    # --------------- TAB 3: CẬP NHẬT BỔ SUNG ---------------
    def setup_tab3_interface(self):
        self.file_sk_cu = self.file_dc_sua = None
        tk.Label(self.tab3, text="CẬP NHẬT FILE SAO KÊ TỪ FILE SỬA TAY", font=("Arial", 12, "bold"), fg="#FF8C00").pack(pady=15)
        f_files = tk.Frame(self.tab3); f_files.pack(fill="x", padx=20, pady=10)
        tk.Button(f_files, text="1. File Sao Kê (Cần cập nhật)", command=lambda: self.chon_file("sk3"), width=25).grid(row=0, column=0, pady=5, padx=5)
        self.lbl_sk_cu = tk.Label(f_files, text="Chưa chọn...", fg="gray"); self.lbl_sk_cu.grid(row=0, column=1, sticky="w")
        tk.Button(f_files, text="2. File Đối Chiếu (Đã sửa)", command=lambda: self.chon_file("dc3"), width=25).grid(row=1, column=0, pady=5, padx=5)
        self.lbl_dc_sua = tk.Label(f_files, text="Chưa chọn...", fg="gray"); self.lbl_dc_sua.grid(row=1, column=1, sticky="w")
        
        self.btn_run_t3 = tk.Button(self.tab3, text="TIẾN HÀNH CẬP NHẬT", bg="#FF8C00", fg="white", font=("Arial", 11, "bold"), command=self.t3_chay, height=2)
        self.btn_run_t3.pack(fill="x", padx=25, pady=15)
        self.progress_var_t3 = tk.DoubleVar()
        self.progressbar_t3 = ttk.Progressbar(self.tab3, variable=self.progress_var_t3, maximum=100)
        self.progressbar_t3.pack(fill="x", padx=25, pady=5)
        self.lbl_status_t3 = tk.Label(self.tab3, text="Sẵn sàng...", font=("Arial", 9, "italic"))
        self.lbl_status_t3.pack()

    def t3_update_progress(self, percent, text):
        self.root.after(0, lambda: self.progress_var_t3.set(percent))
        self.root.after(0, lambda: self.lbl_status_t3.config(text=text))

    def t3_chay(self):
        if not self.file_sk_cu or not self.file_dc_sua: 
            return messagebox.showerror("Lỗi", "Chọn đủ 2 file!")
        p_save = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="SaoKe_Final.xlsx")
        if not p_save: return
        self.btn_run_t3.config(state="disabled")
        self.progress_var_t3.set(0)
        def task():
            try:
                process_update_saoke(self.file_sk_cu, self.file_dc_sua, p_save, self.t3_update_progress)
                self.root.after(0, lambda: messagebox.showinfo("Thành công", "Đã cập nhật dữ liệu sửa tay!"))
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Lỗi", str(e)))
            finally:
                self.root.after(0, lambda: self.btn_run_t3.config(state="normal"))
        threading.Thread(target=task, daemon=True).start()

    # --------------- TAB 2: SCAN HÌNH ẢNH ---------------
    def setup_tab2_interface(self):
        self.selected_images = []
        tk.Label(self.tab2, text="ỨNG DỤNG GOM HÌNH ẢNH THÀNH PDF", font=("Arial", 13, "bold"), fg="#2E7D32").pack(pady=15)
        f_btns = tk.Frame(self.tab2); f_btns.pack()
        tk.Button(f_btns, text="+ Chọn Thêm Ảnh", command=self.t2_chon_anh, width=20).grid(row=0, column=0, padx=5)
        tk.Button(f_btns, text="🗑 Xóa Danh Sách", command=self.t2_xoa_anh, width=20).grid(row=0, column=1, padx=5)
        self.txt_img_list = tk.Text(self.tab2, height=12, width=65, state="disabled", bg="#F5F5F5")
        self.txt_img_list.pack(pady=15)
        self.btn_run_t2 = tk.Button(self.tab2, text="XUẤT RA PDF", bg="#2E7D32", fg="white", font=("Arial", 11, "bold"), command=self.t2_chay, state="disabled", height=2)
        self.btn_run_t2.pack(fill="x", padx=25, pady=10)

    def t2_chon_anh(self):
        files = filedialog.askopenfilenames(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if files: 
            self.selected_images.extend(files)
            self.t2_cap_nhat_ui()

    def t2_xoa_anh(self): 
        self.selected_images = []
        self.t2_cap_nhat_ui()

    def t2_cap_nhat_ui(self):
        self.txt_img_list.config(state="normal")
        self.txt_img_list.delete("1.0", tk.END)
        for i, f in enumerate(self.selected_images): 
            self.txt_img_list.insert(tk.END, f"{i+1}. {os.path.basename(f)}\n")
        self.txt_img_list.config(state="disabled")
        self.btn_run_t2.config(state="normal" if self.selected_images else "disabled")

    def t2_chay(self):
        p_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="TaiLieu_Scan.pdf")
        if p_pdf:
            self.btn_run_t2.config(state="disabled", text="ĐANG GHÉP PDF...")
            self.root.update()
            try:
                imgs = [Image.open(p).convert('RGB') for p in self.selected_images]
                if imgs: 
                    imgs[0].save(p_pdf, save_all=True, append_images=imgs[1:])
                    messagebox.showinfo("Thành công", f"Đã lưu tại:\n{p_pdf}")
                    self.t2_xoa_anh()
            except Exception as e: 
                messagebox.showerror("Lỗi", str(e))
            finally:
                self.btn_run_t2.config(state="normal", text="XUẤT RA PDF")

    # --------------- HÀM DÙNG CHUNG ---------------
    def chon_file(self, loai):
        if loai == "pdf":
            paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
            if paths: 
                self.file_pdf_paths = paths
                self.lbl_pdf.config(text=f"Đã chọn {len(paths)} file PDF", fg="red")
            return
            
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        if path:
            if loai == "sk1": 
                self.file_saoke_path = path
                self.lbl_saoke.config(text=os.path.basename(path), fg="blue")
            elif loai == "mst": 
                self.file_master_path = path
                self.lbl_master.config(text=os.path.basename(path), fg="blue")
            elif loai == "muavao":
                self.file_muavao_path = path
                self.lbl_muavao.config(text=os.path.basename(path), fg="green")
            elif loai == "banra":
                self.file_banra_path = path
                self.lbl_banra.config(text=os.path.basename(path), fg="green")
            elif loai == "mst4": 
                self.file_master_t4_path = path
                self.lbl_master_t4.config(text=os.path.basename(path), fg="red")
            elif loai == "tmpl": 
                self.file_template_path = path
                self.lbl_tmpl.config(text=os.path.basename(path), fg="green")
            elif loai == "sk3": 
                self.file_sk_cu = path
                self.lbl_sk_cu.config(text=os.path.basename(path), fg="blue")
            elif loai == "dc3": 
                self.file_dc_sua = path
                self.lbl_dc_sua.config(text=os.path.basename(path), fg="blue")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGomNghiepVu(root)
    root.mainloop()