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

# ==========================================
# PHIÊN BẢN ĐÃ SửA LỖI (17/07/2026) - Update 2
# - Tên Tab2 và Tab3 đã làm rõ ràng, khớp với tên người dùng gọi (chỉnh sửa tay, gom ảnh PDF)
# - Tab2 (chỉnh sửa tay): Sửa logic cột động (như Tab1), cải thiện if-else gán Mã đúng bên TKNO/TKCO, parse an toàn, thêm progress + set GHICHU='SửA TAY'
# - Tab3 (gom PDF): Sửa lỗi \n (hiển thị sai line break), thêm label đếm số ảnh, feedback rõ ràng hơn, xử lý PDF 1 hoặc nhiều ảnh đúng
# Yêu cầu chạy: pip install pandas openpyxl pillow unidecode
# ==========================================

# ==========================================
# CẤU HÌNH CỐT & TỪ ĐIỂN
# ==========================================
COL_TEN_CTY = 'TENDTPN' 
COL_MA = 'MADTPN'     
COL_DIENGIAI = 'DIENGIAI'
ALIAS_WORDS = {r'\bbetong\b': 'be tong'}
ACTION_VERBS = ['chuyen tien', 'chuyen khoan', 'ck', 'thanh toan', 'tt', 'tra tien', 'ct']
BASE_STOP_WORDS = [r'\bchuyen khoan\b', r'\bck\b', r'\bthanh toan\b', r'\btt\b', r'\brut sec\b', r'\bsec\b', r'\brut tien\b', r'\bnop tien\b', r'\bvao tk\b', r'\bphi dv\b', r'\bphi quan ly\b', r'\bsms banking\b', r'\bhoa don\b', r'\bhd\b', r'\bdot\b', r'\bcuoi\b', r'\btam ung\b', r'\bquyet toan\b', r'\bhdtc\b', r'\bhdkt\b', r'\bbe tong\b', r'\bbetong\b', r'\bly tam\b', r'\bthep\b', r'\bauto\b', r'\bcong ty\b', r'\bcty\b', r'\bctcp\b', r'\bco phan\b', r'\bcp\b', r'\btrach nhiem huu han\b', r'\btnhh\b', r'\bmtv\b', r'\bm t v\b', r'\bmot thanh vien\b', r'\bchi nhanh\b', r'\btap doan\b', r'\blien hop\b', r'\bco so\b', r'\bdoanh nghiep\b', r'\bviet nam\b', r'\bvn\b', r'\bva\b', r'\bgroup\b', r'\bholdings\b', r'\bthuong mai\b', r'\btm\b', r'\bdich vu\b', r'\bdv\b', r'\btmdv\b', r'\btmcp\b', r'\bsan xuat\b', r'\bsx\b', r'\bxuat nhap khau\b', r'\bxnk\b', r'\bdau tu\b', r'\bxay dung\b', r'\bxd\b', r'\bcong nghiep\b', r'\bcn\b', r'\bco khi\b', r'\bkim loai\b', r'\bngu kim\b', r'\bvat lieu\b', r'\bdanh bong\b', r'\bkhuon mau\b', r'\bgia cong\b', r'\bkho bai\b', r'\bbao ve\b', r'\bkhoa hoc\b', r'\bcong nghe\b', r'\bmoi truong\b', r'\bmachinery\b', r'\bmetal\b', r'\bnhom hop kim\b', r'\bnhom\b', r'\bhop kim\b', r'\bthiet bi dien\b', r'\bdien\b', r'\bthiet bi\b', r'\bphat trien\b', r'\bky thuat\b', r'\btong hop\b', r'\bquoc te\b', r'\bche tao\b', r'\bvan tai\b', r'\bvat tu\b', r'\bphu lieu\b', r'\bnhua\b', r'\bbao bi\b', r'\btrang tri\b']

def apply_alias(text):
    for pattern, replacement in ALIAS_WORDS.items(): text = re.sub(pattern, replacement, text)
    return text

def normalize_basic(text):
    if pd.isna(text): return ""
    t = unidecode.unidecode(str(text)).lower().strip()
    t = apply_alias(t); t = re.sub(r'[^a-z0-9\s]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

def get_core_name(text, stop_words_list):
    t = normalize_basic(text)
    for _ in range(2):
        for w in stop_words_list: t = re.sub(w, ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

# Thuật toán Vector/Fuzzy 100% Thuần Python (Mới thêm cho P5)
def get_bigrams(string):
    s = string.replace(' ', '')
    return [s[i:i+2] for i in range(len(s)-1)] if len(s) > 1 else [s]

def nlp_similarity(s1, s2):
    if not s1 or not s2: return 0.0
    bg1, bg2 = get_bigrams(s1), get_bigrams(s2)
    set1, set2 = set(bg1), set(bg2)
    if not set1 or not set2: return 0.0
    return 2.0 * len(set1.intersection(set2)) / (len(set1) + len(set2))

def format_excel_sheet(ws):
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill(start_color="4F81BD", fill_type="solid"); cell.alignment = Alignment(horizontal="center", vertical="center"); cell.border = thin_border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            if cell.column in [2, 3, 4, 5, 6, 7]: cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="left")
            else: cell.alignment = Alignment(vertical="center", horizontal="center")

def parse_amt_to_float(val):
    if pd.isna(val): return 0.0
    s = str(val).strip().replace(',', '').replace(' ', '')
    if not s: return 0.0
    try: return float(s)
    except:
        m = re.search(r'-?\d+(\.\d+)?', s)
        return float(m.group(0)) if m else 0.0

def load_smart_ktsc(file_path):
    if not file_path or not os.path.exists(file_path): return []
    try:
        df = pd.read_excel(file_path, sheet_name='Smart_KTSC_OK', dtype=str) if 'Smart_KTSC_OK' in pd.ExcelFile(file_path).sheet_names else pd.read_excel(file_path, dtype=str)
        items = df.to_dict('records')
        clean_items = []
        for row in items:
            clean_items.append({'SO_HD': str(row.get('SO_HD', '')).strip(), 'MATHANG': str(row.get('MATHANG', '')).strip(), 'TTVND': parse_amt_to_float(row.get('TTVND', 0)), 'TTVND_TT': parse_amt_to_float(row.get('TTVND_TT', 0)), COL_MA: str(row.get('MAKH', '')).strip(), COL_TEN_CTY: str(row.get('TENKH', '')).strip()})
        return clean_items
    except Exception as e:
        print(f"Lỗi đọc file: {e}"); return []

# ==========================================
# LÕI ĐỐI SOÁT SIÊU VIỆT P1 - P8
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
    for item in master_list: item['norm_core'] = get_core_name(item.get(COL_TEN_CTY, ""), dynamic_stops)
    
    list_muavao = load_smart_ktsc(file_muavao)
    list_banra = load_smart_ktsc(file_banra)

    all_matches = []
    total_rows = len(df_bank)
    for idx, row in df_bank.iterrows():
        if progress_callback and idx % 20 == 0: progress_callback(15 + (idx / total_rows) * 65, f"Đối soát dòng {idx+1}...")
        thu_tu_dong = idx + 2
        diengiai_goc = str(row.get(COL_DIENGIAI, ""))
        if pd.isna(diengiai_goc) or diengiai_goc.strip() == "" or diengiai_goc.lower().strip() == "nan":
            all_matches.append({"THỨ TỰ DÒNG GỐC": thu_tu_dong, "DIỄN GIẢI GỐC": "", "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "HỢP ĐỒNG QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "SỐ TIỀN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN": "", "MÃ ĐỐI TƯỢNG PHÁP NHÂN": "", "CÁCH MATCH": "", "SCORE_NUM": 0})
            continue

        diengiai_norm = normalize_basic(diengiai_goc)
        diengiai_cleaned = get_core_name(diengiai_goc, dynamic_stops)
        diengiai_nospace = diengiai_norm.replace(" ", "") 
        matches_for_row = []
        
        # P1, P2, P3, P4
        for master_item in master_list:
            core = master_item['norm_core']; core_words = core.split(); core_nospace = core.replace(" ", "")
            if len(core) < 3: continue
            if re.search(r'\b' + r'\s*'.join(map(re.escape, core_words)) + r'\b', diengiai_norm): matches_for_row.append((1, len(core), master_item, core.upper())); continue
            if len(diengiai_cleaned) >= 4 and re.search(r'\b' + r'\s+'.join(map(re.escape, diengiai_cleaned.split())) + r'\b', core): matches_for_row.append((2, -len(core), master_item, diengiai_cleaned.upper())); continue
            if len(core_words) >= 4:
                acronym = "".join(w[0] for w in core_words)
                if len(acronym) >= 4 and re.search(r'\b(?:' + '|'.join(ACTION_VERBS) + r')\b.*?\b' + re.escape(acronym) + r'\b', diengiai_norm): matches_for_row.append((3, len(acronym), master_item, acronym.upper())); continue
            
            p4_match = False
            for i in range(len(core_words)):
                for j in range(i+1, len(core_words)+1):
                    chunk = "".join(core_words[i:j])
                    if ((j - i >= 2 and len(chunk) >= 5) or len(chunk) >= 8) and chunk in diengiai_nospace: 
                        matches_for_row.append((4, len(chunk), master_item, chunk.upper()))
                        p4_match = True; break
                if p4_match: break
            if p4_match: continue

        # P5: Thuật toán AI Vector (NLP Similarity)
        if not matches_for_row and len(diengiai_cleaned) >= 5:
            best_sim = 0
            best_master = None
            for master_item in master_list:
                sim = nlp_similarity(diengiai_cleaned, master_item['norm_core'])
                if sim > best_sim:
                    best_sim = sim; best_master = master_item
            if best_sim >= 0.85: # Lấy ngưỡng tin cậy 85%
                matches_for_row.append((5, int(best_sim*100), best_master, f"TÊN AI: {best_master['norm_core'].upper()}"))

        # --- KIỂM TRA MUA BÁN (P6, P7, P8) ---
        if not matches_for_row and (list_muavao or list_banra):
            amt_val = parse_amt_to_float(row.get('TTVND', 0))
            if amt_val == 0: amt_val = parse_amt_to_float(row.get('TTVND_TT', 0))
            
            # Chỉ cho phép P6, P7, P8 hoạt động nếu >= 5,000,000 VND
            if amt_val >= 5000000:
                target_list = list_banra + list_muavao
                if target_list:
                    # Chuẩn hóa văn bản giữ lại các ký tự phục vụ Hóa đơn gộp (+, ,, và)
                    text_ext = unidecode.unidecode(str(diengiai_goc)).lower()
                    text_ext = re.sub(r'[^a-z0-9\s/\-\+\,]', ' ', text_ext)
                    text_ext = re.sub(r'\s+', ' ', text_ext).strip()

                    # P6: HÓA ĐƠN DUY NHẤT & GỘP HÓA ĐƠN
                    # Regex bắt cụm hóa đơn: "hoa don so 8, 9 va 10" hoặc "hd 8+9"
                    hd_pattern = r'\b(?:hoa don|hd)\s*(?:so\s*)?((?:\d+(?:\s*(?:,|\+|va\b|-)\s*)*)+)'
                    raw_hd_groups = re.findall(hd_pattern, text_ext)
                    
                    for hd_group in raw_hd_groups:
                        nums = re.findall(r'\d+', hd_group) # Tách ra ['8', '9', '10']
                        if not nums: continue
                        
                        candidate_sets = []
                        best_item_map = {}
                        
                        for num in set(nums):
                            num_clean = num.lstrip('0') or '0'
                            candidates_for_num = set()
                            for item in target_list:
                                item_hd_clean = str(item.get('SO_HD', '')).strip().lstrip('0') or '0'
                                if item_hd_clean == num_clean and num_clean != '0':
                                    ma = item.get(COL_MA)
                                    candidates_for_num.add(ma)
                                    best_item_map[ma] = item
                            
                            if candidates_for_num:
                                candidate_sets.append(candidates_for_num)
                        
                        if candidate_sets:
                            # TÌM GIAO ĐIỂM (INTERSECTION): Khách hàng duy nhất sở hữu TẤT CẢ các hóa đơn này
                            common_candidates = set.intersection(*candidate_sets)
                            # Đảm bảo rủi ro bằng 0: Chỉ lấy khi có đúng 1 Công ty khớp
                            if len(common_candidates) == 1:
                                matched_ma = list(common_candidates)[0]
                                matches_for_row.append((6, len(nums), best_item_map[matched_ma], f"SỐ HĐ: {', '.join(nums)}"))
                                break
                    
                    # P7: ĐỐI CHIẾU HỢP ĐỒNG (Bắt chuỗi có / hoặc -)
                    if not matches_for_row:
                        contracts_found = []
                        explicit_contracts = re.findall(r'\bhop dong\s*(?:so\s*)?([a-z0-9/\-]+)\b', text_ext)
                        contracts_found.extend([c for c in explicit_contracts if len(c) >= 4])
                        explicit_contracts_spaced = re.findall(r'\bhop dong\s*(?:so\s*)?\d+\s+([a-z0-9/\-]+)\b', text_ext)
                        contracts_found.extend([c for c in explicit_contracts_spaced if '/' in c or '-' in c])
                        
                        # Fallback nếu viết tắt hd nhưng có / hoặc -
                        hd_matches = re.findall(r'\bhd\s*(?:so\s*)?([a-z0-9/\-]+)\b', text_ext)
                        for val in hd_matches:
                            if '/' in val or '-' in val: 
                                if len(val) >= 4: contracts_found.append(val)
                                
                        for contract in set(contracts_found):
                            c_clean = re.sub(r'[^A-Z0-9]', '', contract.upper())
                            parts = re.split(r'[/_.-]', contract.upper())
                            valid_parts = [p for p in parts if len(p) >= 5]
                            
                            for item in target_list:
                                mathang_upper = str(item.get('MATHANG', '')).upper()
                                mathang_clean = re.sub(r'[^A-Z0-9]', '', mathang_upper)
                                matched = False
                                if len(c_clean) >= 5 and c_clean in mathang_clean: matched = True
                                elif any(vp in mathang_upper for vp in valid_parts): matched = True
                                
                                if matched: 
                                    matches_for_row.append((7, len(contract), item, f"HỢP ĐỒNG: {contract}"))
                                    break
                            if matches_for_row: break
                    
                    # P8: SỐ TIỀN DUY NHẤT
                    if not matches_for_row:
                        matched_companies = set(); best_item = None
                        for item in target_list:
                            if abs(parse_amt_to_float(item.get('TTVND', 0)) - amt_val) < 1.0 or abs(parse_amt_to_float(item.get('TTVND_TT', 0)) - amt_val) < 1.0: 
                                matched_companies.add(item.get(COL_MA))
                                best_item = item
                        if len(matched_companies) == 1: 
                            matches_for_row.append((8, 0, best_item, f"SỐ TIỀN: {amt_val:,.0f}"))

        # GHI NHậN KẾT QUẢ ĐỐI CHIẾU
        if matches_for_row:
            matches_for_row.sort(key=lambda x: (x[0], -x[1]))
            best_match = matches_for_row[0]
            p_level = best_match[0]
            match_value = str(best_match[3])
            
            ten_quet = match_value if p_level <= 5 else ""
            hd_quet = match_value.replace("SỐ HĐ: ", "") if p_level == 6 else ""
            hopdong_quet = match_value.replace("HỢP ĐỒNG: ", "") if p_level == 7 else ""
            sotien_quet = match_value.replace("SỐ TIỀN: ", "") if p_level == 8 else ""

            all_matches.append({"THỨ TỰ DÒNG GỐC": thu_tu_dong, "DIỄN GIẢI GỐC": diengiai_goc, "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI": ten_quet, "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI": hd_quet, "HỢP ĐỒNG QUÉT ĐƯỢC TRONG DIỄN GIẢI": hopdong_quet, "SỐ TIỀN QUÉT ĐƯỢC TRONG DIỄN GIẢI": sotien_quet, "TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN": best_match[2].get(COL_TEN_CTY, ""), "MÃ ĐỐI TƯỢNG PHÁP NHÂN": best_match[2].get(COL_MA, ""), "CÁCH MATCH": f"P{p_level}", "SCORE_NUM": 100})
        else:
            all_matches.append({"THỨ TỰ DÒNG GỐC": thu_tu_dong, "DIỄN GIẢI GỐC": diengiai_goc, "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "HỢP ĐỒNG QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "SỐ TIỀN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN": "", "MÃ ĐỐI TƯỢNG PHÁP NHÂN": "", "CÁCH MATCH": "", "SCORE_NUM": 0})
    
    wb_doichieu = Workbook(); ws = wb_doichieu.active; ws.title = "Ket Qua Match"
    headers = ["THỨ TỰ DÒNG GỐC", "DIỄN GIẢI GỐC", "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI", "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI", "HỢP ĐỒNG QUÉT ĐƯỢC TRONG DIỄN GIẢI", "SỐ TIỀN QUÉT ĐƯỢC TRONG DIỄN GIẢI", "TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN", "MÃ ĐỐI TƯỢNG PHÁP NHÂN", "CÁCH MATCH"]
    ws.append(headers)
    
    for m in all_matches: ws.append([m.get(h, "") for h in headers])
    format_excel_sheet(ws); wb_doichieu.save(path_save_doichieu)
    
    wb_saoke = load_workbook(file_saoke); ws = wb_saoke.active
    h_d = {str(c.value).strip().upper(): c.column for c in ws[1] if c.value}
    
    col_tkno, col_tkco = h_d.get('TKNO', 5), h_d.get('TKCO', 7)
    col_madtpnno, col_madtpnco = h_d.get('MADTPNNO', 6), h_d.get('MADTPNCO', 8)
    col_tenkh, col_ghichu = h_d.get('TENKH', 11), h_d.get('GHICHU')
    
    def is_valid_cell(val): return pd.notna(val) and str(val).strip() != "" and str(val).lower().strip() != "nan"

    for r in range(2, ws.max_row + 1):
        matched = [m for m in all_matches if m["THỨ TỰ DÒNG GỐC"] == r and m["SCORE_NUM"] == 100]
        if matched:
            m = matched[0]
            tkco_val = str(ws.cell(row=r, column=col_tkco).value).strip() if is_valid_cell(ws.cell(row=r, column=col_tkco).value) else ""
            tkno_val = str(ws.cell(row=r, column=col_tkno).value).strip() if is_valid_cell(ws.cell(row=r, column=col_tkno).value) else ""

            if tkco_val == "1121" or (tkco_val != "" and tkno_val == ""): ws.cell(row=r, column=col_madtpnno).value = m["MÃ ĐỐI TƯỢNG PHÁP NHÂN"]
            elif tkno_val == "1121" or (tkno_val != "" and tkco_val == ""): ws.cell(row=r, column=col_madtpnco).value = m["MÃ ĐỐI TƯỢNG PHÁP NHÂN"]
            else: ws.cell(row=r, column=col_madtpnno).value = m["MÃ ĐỐI TƯỢNG PHÁP NHÂN"]
            
            ws.cell(row=r, column=col_tenkh).value = m["TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN"]
            if col_ghichu: ws.cell(row=r, column=col_ghichu).value = m["CÁCH MATCH"]
        else:
            for c in range(1, ws.max_column + 1): ws.cell(row=r, column=c).fill = PatternFill(start_color="FFFF00", fill_type="solid")

    if progress_callback: progress_callback(98, "Đang lưu file...")
    wb_saoke.save(path_save_saokemoi)
    if progress_callback: progress_callback(100, "Hoàn tất thành công!")

# ==========================================
# CậP NHậT TỪ FILE SửA TAY (ĐÃ SửA LỖI: cột động + logic đúng)
# ==========================================
def process_update_saoke(file_sk_old, file_dc_edited, path_save, progress_callback=None):
    if progress_callback: progress_callback(5, "Đang đọc file đối chiếu đã chỉnh sửa tay...")
    try:
        df_dc = pd.read_excel(file_dc_edited)
    except Exception as e:
        raise Exception(f"Không đọc được file đối chiếu: {e}")

    # An toàn parse update_map - hỗ trợ cả float row num
    update_map = {}
    for _, row in df_dc.iterrows():
        thu_tu = row.get("THỨ TỰ DÒNG GỐC")
        ma_val = row.get("MÃ ĐỐI TƯỢNG PHÁP NHÂN")
        ten_val = row.get("TÊN MATCH ĐƯỢC TRONG FILE ĐỐI TƯỢNG PHÁP NHÂN", "")
        if pd.notna(thu_tu) and pd.notna(ma_val) and str(ma_val).strip():
            try:
                key = int(float(thu_tu))
                update_map[key] = {
                    "MA": str(ma_val).strip(),
                    "TEN": str(ten_val).strip()
                }
            except (ValueError, TypeError):
                continue

    if progress_callback: progress_callback(25, f"Tìm thấy {len(update_map)} dòng cần cập nhật...")

    if progress_callback: progress_callback(35, "Đang mở file Sao kê gốc...")
    try:
        wb = load_workbook(file_sk_old)
        ws = wb.active
    except Exception as e:
        raise Exception(f"Không mở được file Sao kê: {e}")

    # === DYNAMIC COLUMN DETECTION (giống Tab1 - fix lỗi cột cứng) ===
    h_d = {str(c.value).strip().upper(): c.column for c in ws[1] if c.value}
    col_tkno = h_d.get('TKNO', 5)
    col_tkco = h_d.get('TKCO', 7)
    col_madtpnno = h_d.get('MADTPNNO', 6)
    col_madtpnco = h_d.get('MADTPNCO', 8)
    col_tenkh = h_d.get('TENKH', 11)
    col_ghichu = h_d.get('GHICHU', None)

    no_fill = PatternFill(fill_type=None)
    updated_count = 0

    for r in range(2, ws.max_row + 1):
        if r in update_map:
            ma = update_map[r]["MA"]
            ten = update_map[r]["TEN"]

            # Lấy giá trị TK để quyết định bên nào (NO/CO) - logic tư᨜ tự Tab1
            tkno_val = str(ws.cell(row=r, column=col_tkno).value or "").strip().upper()
            tkco_val = str(ws.cell(row=r, column=col_tkco).value or "").strip().upper()

            if tkco_val == "1121" or (tkco_val and not tkno_val):
                ws.cell(row=r, column=col_madtpnno).value = ma
            elif tkno_val == "1121" or (tkno_val and not tkco_val):
                ws.cell(row=r, column=col_madtpnco).value = ma
            else:
                ws.cell(row=r, column=col_madtpnno).value = ma  # mặc định NO

            ws.cell(row=r, column=col_tenkh).value = ten

            # Ghi chú ngắn (phù hợp với yêu cầu của bạn)
            if col_ghichu and col_ghichu <= ws.max_column:
                ws.cell(row=r, column=col_ghichu).value = "SửA TAY"

            # Xóa fill vàng (nếu có)
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).fill = no_fill

            updated_count += 1

    if progress_callback: progress_callback(90, f"Đã cập nhật {updated_count} dòng...")
    wb.save(path_save)
    if progress_callback: progress_callback(100, "Cập nhật thành công! (SửA TAY)")

# ==========================================
# GIAO DIỆN HỆ THỐNG
# ==========================================
class AppGomNghiepVu:
    def __init__(self, root):
        self.root = root
        self.root.title("Hệ Thống Xử Lý Kế Toán & Ngân Hàng Toàn Diện ETAX Hồ Chí Minh")
        self.root.geometry("780x670")
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook) 
        self.tab3 = ttk.Frame(self.notebook) 
        self.notebook.add(self.tab1, text="  1. Đối Soát Ngân Hàng (Excel)  ")
        self.notebook.add(self.tab2, text="  2. Chỉnh Sửa Tay - Cập Nhật Mã & Tên từ File Đối Chiếu  ")
        self.notebook.add(self.tab3, text="  3. Gom Ảnh Hàng Loạt Thành PDF Chứng Từ  ")
        
        self.setup_tab1_interface()
        self.setup_tab2_interface() 
        self.setup_tab3_interface() 

    def setup_tab1_interface(self):
        self.file_saoke_path = self.file_master_path = self.file_muavao_path = self.file_banra_path = None
        tk.Label(self.tab1, text="PHẦN MỀM ĐỐI CHIẾU MÃ PHÁP NHÂN TRÊN EXCEL", font=("Arial", 14, "bold"), fg="#4F81BD").pack(pady=10)
        f_files = tk.Frame(self.tab1); f_files.pack(fill="x", padx=20, pady=5)
        tk.Button(f_files, text="1. Chọn File Sao Kê (Excel)", command=lambda: self.chon_file("sk1"), width=25).grid(row=0, column=0, pady=5, padx=5)
        self.lbl_saoke = tk.Label(f_files, text="Chưa chọn file (Bắt buộc)...", fg="gray"); self.lbl_saoke.grid(row=0, column=1, sticky="w")
        tk.Button(f_files, text="2. Chọn Master MADTPN", command=lambda: self.chon_file("mst"), width=25).grid(row=1, column=0, pady=5, padx=5)
        self.lbl_master = tk.Label(f_files, text="Chưa chọn file (Bắt buộc)...", fg="gray"); self.lbl_master.grid(row=1, column=1, sticky="w")
        tk.Button(f_files, text="3. Chọn File MUA VÀO Năm", command=lambda: self.chon_file("muavao"), width=25).grid(row=2, column=0, pady=5, padx=5)
        self.lbl_muavao = tk.Label(f_files, text="Chưa chọn file (Không bắt buộc)...", fg="gray"); self.lbl_muavao.grid(row=2, column=1, sticky="w")
        tk.Button(f_files, text="4. Chọn File BÁN RA Năm", command=lambda: self.chon_file("banra"), width=25).grid(row=3, column=0, pady=5, padx=5)
        self.lbl_banra = tk.Label(f_files, text="Chưa chọn file (Không bắt buộc)...", fg="gray"); self.lbl_banra.grid(row=3, column=1, sticky="w")
        f_stop = tk.Frame(self.tab1); f_stop.pack(fill="x", padx=20, pady=10)
        tk.Label(f_stop, text="Nhập Tên Chủ TK cần loại trừ khỏi diễn giải:", font=("Arial", 10, "bold")).pack(anchor="w")
        self.entry_stop_words = tk.Entry(f_stop, width=70, font=("Arial", 11)); self.entry_stop_words.pack(anchor="w", ipady=3)
        self.btn_run_t1 = tk.Button(self.tab1, text="BẨT ĐẦU ĐỐI SOÁT EXCEL", bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), command=self.t1_chay, height=2)
        self.btn_run_t1.pack(fill="x", padx=25, pady=10)
        self.progress_var_t1 = tk.DoubleVar()
        self.progressbar_t1 = ttk.Progressbar(self.tab1, variable=self.progress_var_t1, maximum=100); self.progressbar_t1.pack(fill="x", padx=25, pady=5)
        self.lbl_status_t1 = tk.Label(self.tab1, text="Sẵn sàng...", font=("Arial", 9, "italic")); self.lbl_status_t1.pack()

    def t1_chay(self):
        if not self.file_saoke_path or not self.file_master_path: return messagebox.showerror("Lỗi", "Vui lòng chọn đủ File Sao kê và File data MADTPN!")
        p_dc = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="KetQua_DoiChieu.xlsx")
        if not p_dc: return
        p_sk = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="SaoKe_DaCapNhat.xlsx")
        if not p_sk: return
        user_stops = self.entry_stop_words.get()
        self.btn_run_t1.config(state="disabled"); self.progress_var_t1.set(0)
        def task():
            try:
                process_bank_data(file_saoke=self.file_saoke_path, file_master=self.file_master_path, path_save_doichieu=p_dc, path_save_saokemoi=p_sk, user_stop_str=user_stops, file_muavao=self.file_muavao_path, file_banra=self.file_banra_path, progress_callback=self.t1_update_progress)
                self.root.after(0, lambda: messagebox.showinfo("Xong", "Đã đối soát thành công!"))
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Lỗi", f"Có lỗi xảy ra:\n{str(e)}"))
            finally:
                self.root.after(0, lambda: self.btn_run_t1.config(state="normal"))
        threading.Thread(target=task, daemon=True).start()

    def setup_tab2_interface(self):
        self.file_sk_cu = self.file_dc_sua = None
        tk.Label(self.tab2, text="CậP NHậT FILE SAO KÊ TỪ FILE CHỈNH SửA TAY (Từ KetQua_DoiChieu đã sửa)", font=("Arial", 12, "bold"), fg="#FF8C00").pack(pady=15)
        f_files = tk.Frame(self.tab2); f_files.pack(fill="x", padx=20, pady=10)
        tk.Button(f_files, text="1. File Sao Kê gốc (Cần cập nhật)", command=lambda: self.chon_file("sk3"), width=28).grid(row=0, column=0, pady=5, padx=5)
        self.lbl_sk_cu = tk.Label(f_files, text="Chưa chọn...", fg="gray"); self.lbl_sk_cu.grid(row=0, column=1, sticky="w")
        tk.Button(f_files, text="2. File Đối Chiếu đã Sửa Tay", command=lambda: self.chon_file("dc3"), width=28).grid(row=1, column=0, pady=5, padx=5)
        self.lbl_dc_sua = tk.Label(f_files, text="Chưa chọn...", fg="gray"); self.lbl_dc_sua.grid(row=1, column=1, sticky="w")
        self.btn_run_t2 = tk.Button(self.tab2, text="TIẾN HÀNH CậP NH᫒T ĐÈ MÃ (SửA TAY)", bg="#FF8C00", fg="white", font=("Arial", 11, "bold"), command=self.t2_chay, height=2)
        self.btn_run_t2.pack(fill="x", padx=25, pady=15)
        self.progress_var_t2 = tk.DoubleVar()
        self.progressbar_t2 = ttk.Progressbar(self.tab2, variable=self.progress_var_t2, maximum=100); self.progressbar_t2.pack(fill="x", padx=25, pady=5)
        self.lbl_status_t2 = tk.Label(self.tab2, text="Sẵn sàng...", font=("Arial", 9, "italic")); self.lbl_status_t2.pack()

    def t2_chay(self):
        if not self.file_sk_cu or not self.file_dc_sua: return messagebox.showerror("Lỗi", "Chọn đủ 2 file!")
        p_save = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="SaoKe_Final_SuaTay.xlsx")
        if not p_save: return
        self.btn_run_t2.config(state="disabled"); self.progress_var_t2.set(0)
        def task():
            try:
                process_update_saoke(self.file_sk_cu, self.file_dc_sua, p_save, self.t2_update_progress)
                self.root.after(0, lambda: messagebox.showinfo("Thành công", "Đã cập nhật dữ liệu sửa tay thành công!\nFile đã lưu: " + p_save))
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Lỗi", str(e)))
            finally:
                self.root.after(0, lambda: self.btn_run_t2.config(state="normal"))
        threading.Thread(target=task, daemon=True).start()

    def setup_tab3_interface(self):
        self.selected_images = []
        tk.Label(self.tab3, text="ỨNG DỤNG GOM HÌNH ẢNH THÀNH PDF CHỨNG TỪ", font=("Arial", 13, "bold"), fg="#2E7D32").pack(pady=15)
        f_btns = tk.Frame(self.tab3); f_btns.pack()
        tk.Button(f_btns, text="+ Chọn Thêm Ảnh", command=self.t3_chon_anh, width=20).grid(row=0, column=0, padx=5)
        tk.Button(f_btns, text="🗑 Xóa Danh Sách", command=self.t3_xoa_anh, width=20).grid(row=0, column=1, padx=5)
        self.txt_img_list = tk.Text(self.tab3, height=10, width=65, state="disabled", bg="#F5F5F5"); self.txt_img_list.pack(pady=10)
        self.lbl_so_anh = tk.Label(self.tab3, text="Chưa chọn ảnh nào. Chọn ảnh theo thứ tự mong muốn (trang 1, 2, ...)", fg="gray", font=("Arial", 9)); self.lbl_so_anh.pack()
        self.btn_run_t3 = tk.Button(self.tab3, text="XUẤT RA PDF", bg="#2E7D32", fg="white", font=("Arial", 11, "bold"), command=self.t3_chay, state="disabled", height=2)
        self.btn_run_t3.pack(fill="x", padx=25, pady=10)

    def t3_chon_anh(self):
        files = filedialog.askopenfilenames(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if files: self.selected_images.extend(files); self.t3_cap_nhat_ui()

    def t3_xoa_anh(self): 
        self.selected_images = []; self.t3_cap_nhat_ui()

    def t3_cap_nhat_ui(self):
        self.txt_img_list.config(state="normal"); self.txt_img_list.delete("1.0", tk.END)
        for i, f in enumerate(self.selected_images): 
            self.txt_img_list.insert(tk.END, f"{i+1}. {os.path.basename(f)}\n")  # FIX: dùng \n thực (Python f-string newline)
        self.txt_img_list.config(state="disabled")
        self.btn_run_t3.config(state="normal" if self.selected_images else "disabled")
        self.lbl_so_anh.config(text=f"Đã chọn {len(self.selected_images)} ảnh (thứ tự = thứ tự trang trong PDF)")

    def t3_chay(self):
        p_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="TaiLieu_ChungTu_Scan.pdf")
        if p_pdf:
            self.btn_run_t3.config(state="disabled", text="ĐANG GHÉP PDF... (có thể mất thời gian)"); self.root.update()
            try:
                imgs = [Image.open(p).convert('RGB') for p in self.selected_images]
                if imgs:
                    # Xử lý đúng cho cả 1 và nhiều ảnh
                    if len(imgs) == 1:
                        imgs[0].save(p_pdf, "PDF", resolution=100.0)
                    else:
                        imgs[0].save(p_pdf, save_all=True, append_images=imgs[1:], resolution=100.0)
                    messagebox.showinfo("Thành công", f"Đã tạo PDF thành công!\n{p_pdf}\nTổng {len(imgs)} trang.")
                    self.t3_xoa_anh()
                else:
                    messagebox.showwarning("Cảnh báo", "Không có ảnh nào để tạo PDF.")
            except Exception as e: 
                messagebox.showerror("Lỗi PDF", f"Lỗi khi tạo PDF:\n{str(e)}\nKiểm tra Pillow đã cài đặt chưa (pip install Pillow)")
            finally: 
                self.btn_run_t3.config(state="normal", text="XUẤT RA PDF")

    def t1_update_progress(self, percent, text):
        self.root.after(0, lambda: self.progress_var_t1.set(percent))
        self.root.after(0, lambda: self.lbl_status_t1.config(text=text))

    def t2_update_progress(self, percent, text):
        self.root.after(0, lambda: self.progress_var_t2.set(percent))
        self.root.after(0, lambda: self.lbl_status_t2.config(text=text))

    def chon_file(self, loai):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        if path:
            if loai == "sk1": self.file_saoke_path = path; self.lbl_saoke.config(text=os.path.basename(path), fg="blue")
            elif loai == "mst": self.file_master_path = path; self.lbl_master.config(text=os.path.basename(path), fg="blue")
            elif loai == "muavao": self.file_muavao_path = path; self.lbl_muavao.config(text=os.path.basename(path), fg="green")
            elif loai == "banra": self.file_banra_path = path; self.lbl_banra.config(text=os.path.basename(path), fg="green")
            elif loai == "sk3": self.file_sk_cu = path; self.lbl_sk_cu.config(text=os.path.basename(path), fg="blue")
            elif loai == "dc3": self.file_dc_sua = path; self.lbl_dc_sua.config(text=os.path.basename(path), fg="blue")

if __name__ == "__main__":
    root = tk.Tk(); app = AppGomNghiepVu(root); root.mainloop()