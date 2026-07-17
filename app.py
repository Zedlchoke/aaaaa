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
from openpyxl.utils import get_column_letter
from PIL import Image

COL_TEN_CTY = 'TENDTPN'
COL_MA = 'MADTPN'
COL_DIENGIAI = 'DIENGIAI'

ALIAS_WORDS = {r'\bbetong\b': 'be tong'}
ACTION_VERBS = ['chuyen tien', 'chuyen khoan', 'ck', 'thanh toan', 'tt', 'tra tien', 'ct']
BASE_STOP_WORDS = [r'\bchuyen khoan\b', r'\bck\b', r'\bthanh toan\b', r'\btt\b', r'\brut sec\b', r'\bsec\b', r'\brut tien\b', r'\bnop tien\b', r'\bvao tk\b', r'\bphi dv\b', r'\bphi quan ly\b', r'\bsms banking\b', r'\bhoa don\b', r'\bhd\b', r'\bdot\b', r'\bcuoi\b', r'\btam ung\b', r'\bquyet toan\b', r'\bhdtc\b', r'\bhdkt\b', r'\bbe tong\b', r'\bbetong\b', r'\bly tam\b', r'\bthep\b', r'\bauto\b', r'\bcong ty\b', r'\bcty\b', r'\bctcp\b', r'\bco phan\b', r'\bcp\b', r'\btrach nhiem huu han\b', r'\btnhh\b', r'\bmtv\b', r'\bm t v\b', r'\bmot thanh vien\b', r'\bchi nhanh\b', r'\btap doan\b', r'\blien hop\b', r'\bco so\b', r'\bdoanh nghiep\b', r'\bviet nam\b', r'\bvn\b', r'\bva\b', r'\bgroup\b', r'\bholdings\b', r'\bthuong mai\b', r'\btm\b', r'\bdich vu\b', r'\bdv\b', r'\btmdv\b', r'\btmcp\b', r'\bsan xuat\b', r'\bsx\b', r'\bxuat nhap khau\b', r'\bxnk\b', r'\bdau tu\b', r'\bxay dung\b', r'\bxd\b', r'\bcong nghiep\b', r'\bcn\b', r'\bco khi\b', r'\bkim loai\b', r'\bngu kim\b', r'\bvat lieu\b', r'\bdanh bong\b', r'\bkhuon mau\b', r'\bgia gia cong\b', r'\bkho bai\b', r'\bbao ve\b', r'\bkhoa hoc\b', r'\bcong nghe\b', r'\bmoi truong\b', r'\bmachinery\b', r'\bmetal\b', r'\bnhom hop kim\b', r'\bnhom\b', r'\bhop kim\b', r'\bthiet bi dien\b', r'\bdien\b', r'\bthiet bi\b', r'\bphat trien\b', r'\bky thuat\b', r'\btong hop\b', r'\bquoc te\b', r'\bche tao\b', r'\bvan tai\b', r'\bvat tu\b', r'\bphu lieu\b', r'\bnhua\b', r'\bbao bi\b', r'\btrang tri\b']

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
        for w in stop_words_list: t = re.sub(w, ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

def get_bigrams(string):
    s = string.replace(' ', '')
    return [s[i:i+2] for i in range(len(s)-1)] if len(s) > 1 else [s]

def format_excel_sheet(ws, is_doichieu=False):
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    if is_doichieu:
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 55
        for col in ['C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 22
        ws.column_dimensions['G'].width = 40
        ws.column_dimensions['H'].width = 18
        ws.column_dimensions['I'].width = 14
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for col_idx in [2, 3, 4, 5, 6, 7]:
                cell = row[col_idx - 1]
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    else:
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 18

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

def process_bank_data(file_saoke, file_master, path_save_doichieu, path_save_saokemoi, user_stop_str, file_muavao=None, file_banra=None, progress_callback=None, enabled_ps=None, custom_order=None):
    dynamic_stops = BASE_STOP_WORDS.copy()
    user_stops = [w.strip() for w in user_stop_str.split(',') if w.strip()]
    for w in user_stops:
        w_norm = normalize_basic(w)
        if w_norm: dynamic_stops.append(r'\b' + re.escape(w_norm) + r'\b')
        
    if progress_callback: progress_callback(5, "Đang đọc file Excel...")
    df_bank = pd.read_excel(file_saoke)
    df_master = pd.read_excel(file_master)
    master_list = df_master.to_dict('records')
    
    for item in master_list:
        item['norm_core'] = get_core_name(item.get(COL_TEN_CTY, ""), dynamic_stops)
        core = item['norm_core']
        if len(core) >= 3:
            core_words = core.split()
            p1_pat = r'\b' + r'\s*'.join(map(re.escape, core_words)) + r'\b'
            item['p1_regex'] = re.compile(p1_pat)
        else:
            item['p1_regex'] = None
        item['bigram_set'] = set(get_bigrams(core))
        
    list_muavao = load_smart_ktsc(file_muavao)
    list_banra = load_smart_ktsc(file_banra)
    hd_index_muavao = {}
    hd_index_banra = {}
    
    for item in list_muavao:
        hd = str(item.get('SO_HD', '')).strip().lstrip('0') or '0'
        if hd != '0':
            if hd not in hd_index_muavao: hd_index_muavao[hd] = []
            hd_index_muavao[hd].append(item)
    for item in list_banra:
        hd = str(item.get('SO_HD', '')).strip().lstrip('0') or '0'
        if hd != '0':
            if hd not in hd_index_banra: hd_index_banra[hd] = []
            hd_index_banra[hd].append(item)
            
    hd_pattern_comp = re.compile(r'\b(?:hoa don|hd)\s*(?:so\s*)?((?:\d+(?:\s*(?:,|\+|va\b|-)\s*)*)+)')
    
    if enabled_ps is None:
        enabled_ps = {"P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9"}
    else:
        enabled_ps = set([p.upper() for p in enabled_ps])
        
    if custom_order:
        try: order_list = [p.strip().upper() for p in str(custom_order).split(',') if p.strip()]
        except: order_list = None
    else: 
        order_list = None
        
    all_matches = []
    total_rows = len(df_bank)
    h_d_bank = {str(c).strip().upper(): c for c in df_bank.columns}
    col_tkno_name = h_d_bank.get('TKNO', None)
    col_tkco_name = h_d_bank.get('TKCO', None)
    col_tenkh_name_in_df = h_d_bank.get('TENKH', None)
    
    for idx, row in df_bank.iterrows():
        if progress_callback and idx % 20 == 0: progress_callback(15 + (idx / total_rows) * 65, f"Đối soát dòng {idx+1}...")
        thu_tu_dong = idx + 2
        diengiai_goc = str(row.get(COL_DIENGIAI, ""))
        
        if pd.isna(diengiai_goc) or diengiai_goc.strip() == "" or diengiai_goc.lower().strip() == "nan":
            all_matches.append({"THỨ TỰ DÒNG GỐC": thu_tu_dong, "DIỄN GIẢI GỐC": "", "TÊN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "HÓA ĐƠN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "HỢP ĐỒNG QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "SỐ TIỀN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN": "", "MÃ ĐỐI TƯỢNG PHÁP NHÂN": "", "CÁCH MATCH": "", "SCORE_NUM": 0})
            continue
            
        diengiai_norm = normalize_basic(diengiai_goc)
        diengiai_cleaned = get_core_name(diengiai_goc, dynamic_stops)
        diengiai_nospace = diengiai_norm.replace(" ", "")
        amt_val = parse_amt_to_float(row.get('TTVND', 0))
        if amt_val == 0: amt_val = parse_amt_to_float(row.get('TTVND_TT', 0))
        
        tkco_val = str(row.get(col_tkco_name, '')).strip() if col_tkco_name else ""
        tkno_val = str(row.get(col_tkno_name, '')).strip() if col_tkno_name else ""
        is_mua_vao = tkco_val.startswith('112')
        is_ban_ra = tkno_val.startswith('112')
        
        if is_mua_vao and list_muavao:
            current_target_list = list_muavao
            current_hd_index = hd_index_muavao
        elif is_ban_ra and list_banra:
            current_target_list = list_banra
            current_hd_index = hd_index_banra
        else:
            current_target_list = list_banra + list_muavao
            current_hd_index = {**hd_index_banra, **hd_index_muavao}
            
        matches_for_row = []
        quet_ten = quet_hd = quet_hopdong = quet_sotien = ""
        
        def try_p1():
            nonlocal quet_ten
            if "P1" not in enabled_ps: return
            for master_item in master_list:
                core = master_item['norm_core']
                if len(core) < 3: continue
                p1_re = master_item.get('p1_regex')
                if p1_re and p1_re.search(diengiai_norm):
                    quet_ten = core.upper()
                    matches_for_row.append((1, len(core), master_item, core.upper()))
                    return
        def try_p2():
            nonlocal quet_ten
            if "P2" not in enabled_ps: return
            for master_item in master_list:
                core = master_item['norm_core']
                if len(diengiai_cleaned) >= 4 and re.search(r'\b' + r'\s+'.join(map(re.escape, diengiai_cleaned.split())) + r'\b', core):
                    quet_ten = diengiai_cleaned.upper()
                    matches_for_row.append((2, -len(core), master_item, diengiai_cleaned.upper()))
                    return
        def try_p3():
            nonlocal quet_ten
            if "P3" not in enabled_ps: return
            for master_item in master_list:
                core = master_item['norm_core']
                core_words = core.split()
                if len(core_words) >= 4:
                    acronym = "".join(w[0] for w in core_words)
                    if len(acronym) >= 4 and re.search(r'\b(?:' + '|'.join(ACTION_VERBS) + r')\b.*?\b' + re.escape(acronym) + r'\b', diengiai_norm):
                        quet_ten = acronym.upper()
                        matches_for_row.append((3, len(acronym), master_item, acronym.upper()))
                        return
        def try_p4():
            nonlocal quet_ten
            if "P4" not in enabled_ps: return
            for master_item in master_list:
                core = master_item['norm_core']
                core_words = core.split()
                p4_match = False
                for i in range(len(core_words)):
                    for j in range(i+1, len(core_words)+1):
                        chunk = "".join(core_words[i:j])
                        if ((j - i >= 2 and len(chunk) >= 5) or len(chunk) >= 8) and chunk in diengiai_nospace:
                            quet_ten = chunk.upper()
                            matches_for_row.append((4, len(chunk), master_item, chunk.upper()))
                            p4_match = True
                            break
                    if p4_match: break
                if p4_match: return
        def try_p5():
            nonlocal quet_ten
            if "P5" not in enabled_ps: return
            if len(diengiai_cleaned) >= 5:
                best_sim = 0.0
                best_master = None
                bg1 = get_bigrams(diengiai_cleaned)
                set1 = set(bg1)
                len_set1 = len(set1)
                if len_set1 > 0:
                    for master_item in master_list:
                        set2 = master_item['bigram_set']
                        inter_len = len(set1.intersection(set2))
                        if inter_len > 0:
                            sim = 2.0 * inter_len / (len_set1 + len(set2))
                            if sim > best_sim:
                                best_sim = sim
                                best_master = master_item
                if best_sim >= 0.85 and best_master is not None:
                    quet_ten = best_master['norm_core'].upper()
                    matches_for_row.append((5, int(best_sim*100), best_master, f"TÊN AI: {best_master['norm_core'].upper()}"))
        def try_p6_p7_p8():
            nonlocal quet_hd, quet_hopdong, quet_sotien
            if not ("P6" in enabled_ps or "P7" in enabled_ps or "P8" in enabled_ps): return
            if not current_target_list or amt_val < 5000000: return
            text_ext = unidecode.unidecode(str(diengiai_goc)).lower()
            text_ext = re.sub(r'[^a-z0-9\s/\-\+\,]', ' ', text_ext)
            text_ext = re.sub(r'\s+', ' ', text_ext).strip()
            if "P6" in enabled_ps:
                raw_hd_groups = hd_pattern_comp.findall(text_ext)
                for hd_group in raw_hd_groups:
                    nums = re.findall(r'\d+', hd_group)
                    if not nums: continue
                    candidate_sets = []
                    best_item_map = {}
                    for num in set(nums):
                        num_clean = num.lstrip('0') or '0'
                        if num_clean == '0': continue
                        cands = current_hd_index.get(num_clean, [])
                        if cands:
                            mas_set = set()
                            for item in cands:
                                ma = item.get(COL_MA)
                                if ma:
                                    mas_set.add(ma)
                                    if ma not in best_item_map: best_item_map[ma] = item
                            if mas_set: candidate_sets.append(mas_set)
                    if candidate_sets:
                        try:
                            common = set.intersection(*candidate_sets)
                            if len(common) == 1:
                                matched_ma = list(common)[0]
                                quet_hd = ', '.join(nums)
                                matches_for_row.append((6, len(nums), best_item_map[matched_ma], f"SỐ HĐ: {', '.join(nums)}"))
                                return
                        except: pass
            if "P7" in enabled_ps and not matches_for_row:
                contracts_found = []
                explicit_contracts = re.findall(r'\bhop dong\s*(?:so\s*)?([a-z0-9/\-]+)\b', text_ext)
                contracts_found.extend([c for c in explicit_contracts if len(c) >= 4])
                explicit_contracts_spaced = re.findall(r'\bhop dong\s*(?:so\s*)?\d+\s+([a-z0-9/\-]+)\b', text_ext)
                contracts_found.extend([c for c in explicit_contracts_spaced if '/' in c or '-' in c])
                hd_matches = re.findall(r'\bhd\s*(?:so\s*)?([a-z0-9/\-]+)\b', text_ext)
                for val in hd_matches:
                    if ('/' in val or '-' in val) and len(val) >= 4: contracts_found.append(val)
                for contract in set(contracts_found):
                    c_clean = re.sub(r'[^A-Z0-9]', '', contract.upper())
                    parts = re.split(r'[/_.-]', contract.upper())
                    valid_parts = [p for p in parts if len(p) >= 5]
                    for item in current_target_list:
                        mathang_upper = str(item.get('MATHANG', '')).upper()
                        mathang_clean = re.sub(r'[^A-Z0-9]', '', mathang_upper)
                        matched = False
                        if len(c_clean) >= 5 and c_clean in mathang_clean: matched = True
                        elif any(vp in mathang_upper for vp in valid_parts): matched = True
                        if matched:
                            quet_hopdong = contract
                            matches_for_row.append((7, len(contract), item, f"HỢP ĐỒNG: {contract}"))
                            return
                    if matches_for_row: break
            if "P8" in enabled_ps and not matches_for_row:
                matched_companies = set()
                best_item = None
                for item in current_target_list:
                    if abs(parse_amt_to_float(item.get('TTVND', 0)) - amt_val) < 1.0 or abs(parse_amt_to_float(item.get('TTVND_TT', 0)) - amt_val) < 1.0:
                        matched_companies.add(item.get(COL_MA))
                        best_item = item
                if len(matched_companies) == 1:
                    quet_sotien = f"{amt_val:,.0f}"
                    matches_for_row.append((8, 0, best_item, f"SỐ TIỀN: {amt_val:,.0f}"))
        
        def try_p9():
            nonlocal quet_ten
            if "P9" not in enabled_ps: return
            tenkh_goc = str(row.get(col_tenkh_name_in_df, '')) if col_tenkh_name_in_df else ""
            if not tenkh_goc or str(tenkh_goc).strip().lower() in ["", "nan"]: return
            tenkh_goc_norm = unidecode.unidecode(tenkh_goc).lower().strip()
            if not current_target_list: return
            for item in current_target_list:
                item_tenkh = str(item.get(COL_TEN_CTY, '')).strip()
                item_tenkh_norm = unidecode.unidecode(item_tenkh).lower().strip()
                if tenkh_goc_norm == item_tenkh_norm and tenkh_goc_norm != "":
                    quet_ten = tenkh_goc
                    matches_for_row.append((9, 100, item, f"P9 - TRÙNG TÊN KH GỐC"))
                    return

        if order_list:
            for p in order_list:
                if p == "P1": try_p1()
                elif p == "P2": try_p2()
                elif p == "P3": try_p3()
                elif p == "P4": try_p4()
                elif p == "P5": try_p5()
                elif p in ("P6", "P7", "P8"): try_p6_p7_p8()
                elif p == "P9": try_p9()
                if matches_for_row: break
        else:
            # THAY ĐỔI: Chạy P9 trước tiên làm ưu tiên cao nhất, sau đó mới đến các thuật toán khác
            try_p9()
            if not matches_for_row:
                try_p1()
                try_p2()
                try_p3()
                try_p4()
                try_p5()
                if not matches_for_row:
                    try_p6_p7_p8()

        if matches_for_row:
            matches_for_row.sort(key=lambda x: (x[0], -x[1]))
            best_match = matches_for_row[0]
            p_level = best_match[0]
            ten_quet = quet_ten if p_level in (1, 2, 3, 4, 5, 9) else ""
            hd_quet = quet_hd if p_level == 6 else ""
            hopdong_quet = quet_hopdong if p_level == 7 else ""
            sotien_quet = quet_sotien if p_level == 8 else ""
            all_matches.append({"THỨ TỰ DÒNG GỐC": thu_tu_dong, "DIỄN GIẢI GỐC": diengiai_goc, "TÊN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": ten_quet, "HÓA ĐƠN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": hd_quet, "HỢP ĐỒNG QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": hopdong_quet, "SỐ TIỀN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": sotien_quet, "TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN": best_match[2].get(COL_TEN_CTY, ""), "MÃ ĐỐI TƯỢNG PHÁP NHÂN": best_match[2].get(COL_MA, ""), "CÁCH MATCH": f"P{p_level}", "SCORE_NUM": 100})
        else:
            all_matches.append({"THỨ TỰ DÒNG GỐC": thu_tu_dong, "DIỄN GIẢI GỐC": diengiai_goc, "TÊN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "HÓA ĐƠN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "HỢP ĐỒNG QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "SỐ TIỀN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN": "", "MÃ ĐỐI TƯỢNG PHÁP NHÂN": "", "CÁCH MATCH": "", "SCORE_NUM": 0})
    
    wb_doichieu = Workbook()
    ws = wb_doichieu.active
    ws.title = "Ket Qua Match"
    headers = ["THỨ TỰ DÒNG GỐC", "DIỄN GIẢI GỐC", "TÊN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI", "HÓA ĐƠN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI", "HỢP ĐỒNG QUÉT ĐƯỢC TRẠNG DIỄN GIẢI", "SỐ TIỀN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI", "TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN", "MÃ ĐỐI TƯỢNG PHÁP NHÂN", "CÁCH MATCH"]
    ws.append(headers)
    for m in all_matches: ws.append([m.get(h, "") for h in headers])
    format_excel_sheet(ws, is_doichieu=True)
    
    ws_legend = wb_doichieu.create_sheet("Giai Thich P1-P9")
    legend_data = [
        ["CÁCH MATCH", "Ý NGHĨA", "MÔ TẢ CHI TIẾT"],
        ["P1", "Khớp tên nguyên vẹn", "Tìm chính xác tên công ty theo từ khóa (word boundary) trong diễn giải"],
        ["P2", "Khớp ngược", "Tên công ty nằm trong diễn giải (ngược lại)"],
        ["P3", "Khớp viết tắt (acronym)", "Dùng chữ cái đầu của từ trong tên công ty kết hợp động từ (chuyển, tt, ck...)"],
        ["P4", "Khớp cụm từ (chunk)", "Tìm cụm từ dài (từ 2 từ trở lên) trong tên công ty xuất hiện trong diễn giải"],
        ["P5", "Fuzzy AI (bigram)", "So sánh độ tương đồng tên bằng thuật toán AI (bigram similarity ≥ 85%)"],
        ["P6", "Hóa đơn (Mua/Bán)", "Tìm số HÓA ĐƠN trong diễn giải → đối chiếu với file Mua VÀO hoặc BÁN RA"],
        ["P7", "Hợp đồng / MATHANG", "Tìm số HỢP ĐỒNG hoặc tên hàng trong diễn giải → đối chiếu với file Mua VÀO hoặc BÁN RA"],
        ["P8", "Số tiền duy nhất", "Khớp duy nhất theo số tiền khi dòng ≥ 5 triệu và có file Mua/Bán"],
        ["P9", "Khớp Tên KH Gốc", "Dùng TENKH sẵn có trên file sao kê gốc để đối chiếu chéo (không dấu) với file Mua/Bán"],
    ]
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for r_idx, row_data in enumerate(legend_data, 1):
        for c_idx, value in enumerate(row_data, 1):
            cell = ws_legend.cell(row=r_idx, column=c_idx, value=value)
            cell.border = thin
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            if r_idx == 1: cell.font = Font(bold=True)
    ws_legend.column_dimensions['A'].width = 14
    ws_legend.column_dimensions['B'].width = 25
    ws_legend.column_dimensions['C'].width = 80
    wb_doichieu.save(path_save_doichieu)
    
    wb_saoke = load_workbook(file_saoke)
    ws = wb_saoke.active
    h_d = {str(c.value).strip().upper(): c.column for c in ws[1] if c.value}
    col_tkno = h_d.get('TKNO', 5)
    col_tkco = h_d.get('TKCO', 7)
    col_madtpnno = h_d.get('MADTPNNO', 6)
    col_madtpnco = h_d.get('MADTPNCO', 8)
    col_tenkh = h_d.get('TENKH', 11)
    col_ghichu = h_d.get('GHICHU')
    def is_valid_cell(val): return pd.notna(val) and str(val).strip() != "" and str(val).lower().strip() != "nan"
    
    yellow_fill = PatternFill(start_color="FFFFFF00", fill_type="solid")
    for r in range(2, ws.max_row + 1):
        matched = [m for m in all_matches if m["THỨ TỰ DÒNG GỐC"] == r and m["SCORE_NUM"] == 100]
        if matched:
            m = matched[0]
            ma_dtpn = str(m.get("MÃ ĐỐI TƯỢNG PHÁP NHÂN", "")).strip()
            if ma_dtpn.lower() == "nan": ma_dtpn = ""
            
            tkco_val = str(ws.cell(row=r, column=col_tkco).value).strip() if is_valid_cell(ws.cell(row=r, column=col_tkco).value) else ""
            tkno_val = str(ws.cell(row=r, column=col_tkno).value).strip() if is_valid_cell(ws.cell(row=r, column=col_tkno).value) else ""
            
            if tkco_val == "1121" or (tkco_val != "" and tkno_val == ""): ws.cell(row=r, column=col_madtpnno).value = ma_dtpn
            elif tkno_val == "1121" or (tkno_val != "" and tkco_val == ""): ws.cell(row=r, column=col_madtpnco).value = ma_dtpn
            else: ws.cell(row=r, column=col_madtpnno).value = ma_dtpn
            
            ws.cell(row=r, column=col_tenkh).value = m["TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN"]
            if col_ghichu: ws.cell(row=r, column=col_ghichu).value = m["CÁCH MATCH"]
            
            if ma_dtpn == "":
                for c in range(1, ws.max_column + 1): ws.cell(row=r, column=c).fill = yellow_fill
        else:
            for c in range(1, ws.max_column + 1): ws.cell(row=r, column=c).fill = yellow_fill
            
    if progress_callback: progress_callback(98, "Đang lưu file...")
    wb_saoke.save(path_save_saokemoi)
    if progress_callback: progress_callback(100, "Hoàn tất thành công!")


class AppGomNghiepVu:
    def __init__(self, root):
        self.root = root
        self.root.title("Hệ Thống Xử Lý Kế Toán & Ngân Hàng Toàn Diện ETAX Hồ Chí Minh")
        self.root.geometry("780x670")
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text="  1. Đối Soát Ngân Hàng (Excel)  ")
        self.notebook.add(self.tab2, text="  2. Gom Ảnh Hàng Loạt Thành PDF Chứng Từ  ")
        
        self.setup_tab1_interface()
        self.setup_tab2_interface()

    def setup_tab1_interface(self):
        self.file_saoke_path = self.file_master_path = self.file_muavao_path = self.file_banra_path = None
        tk.Label(self.tab1, text="PHẦN MỀM ĐỐI CHIẾU MÃ PHÁP NHÂN TRÊN EXCEL", font=("Arial", 14, "bold"), fg="#4F81BD").pack(pady=10)
        
        f_files = tk.Frame(self.tab1)
        f_files.pack(fill="x", padx=20, pady=3)
        tk.Button(f_files, text="1. Chọn File Sao Kê (Excel)", command=lambda: self.chon_file("sk1"), width=25).grid(row=0, column=0, pady=3, padx=5)
        self.lbl_saoke = tk.Label(f_files, text="Chưa chọn file (Bắt buộc)...", fg="gray")
        self.lbl_saoke.grid(row=0, column=1, sticky="w")
        tk.Button(f_files, text="2. Chọn Master MADTPN", command=lambda: self.chon_file("mst"), width=25).grid(row=1, column=0, pady=3, padx=5)
        self.lbl_master = tk.Label(f_files, text="Chưa chọn file (Bắt buộc)...", fg="gray")
        self.lbl_master.grid(row=1, column=1, sticky="w")
        tk.Button(f_files, text="3. Chọn File MUA VÀO Năm", command=lambda: self.chon_file("muavao"), width=25).grid(row=2, column=0, pady=3, padx=5)
        self.lbl_muavao = tk.Label(f_files, text="Chưa chọn file (Không bắt buộc)...", fg="gray")
        self.lbl_muavao.grid(row=2, column=1, sticky="w")
        tk.Button(f_files, text="4. Chọn File BÁN RA Năm", command=lambda: self.chon_file("banra"), width=25).grid(row=3, column=0, pady=3, padx=5)
        self.lbl_banra = tk.Label(f_files, text="Chưa chọn file (Không bắt buộc)...", fg="gray")
        self.lbl_banra.grid(row=3, column=1, sticky="w")
        
        f_match = tk.LabelFrame(self.tab1, text=" Tùy chọn kiểu match (P1-P9) - Tick để kích hoạt ", padx=10, pady=5)
        f_match.pack(fill="x", padx=20, pady=5)
        
        self.p_vars = {}
        p_labels = {
            "P1": "P1 - Khớp tên nguyên vẹn",
            "P2": "P2 - Khớp ngược",
            "P3": "P3 - Khớp viết tắt (acronym)",
            "P4": "P4 - Khớp cụm từ (chunk)",
            "P5": "P5 - Fuzzy AI",
            "P6": "P6 - Hóa đơn gộp",
            "P7": "P7 - Hợp đồng / MATHANG",
            "P8": "P8 - Số tiền duy nhất",
            "P9": "P9 - Dùng TENKH gốc trên sao kê"
        }
        for i, (p, label) in enumerate(p_labels.items()):
            row = i // 2
            col = i % 2
            var = tk.BooleanVar(value=True)
            self.p_vars[p] = var
            cb = ttk.Checkbutton(f_match, text=label, variable=var)
            cb.grid(row=row, column=col, sticky="w", padx=5, pady=2)
            
        f_order = tk.Frame(f_match)
        f_order.grid(row=5, column=0, columnspan=2, sticky="w", pady=5)
        tk.Label(f_order, text="Thứ tự tùy chỉnh (trống):", font=("Arial", 9)).pack(side="left")
        self.custom_order_entry = tk.Entry(f_order, width=35, font=("Arial", 9))
        self.custom_order_entry.pack(side="left", padx=5)
        
        f_stop = tk.Frame(self.tab1)
        f_stop.pack(fill="x", padx=20, pady=5)
        tk.Label(f_stop, text="Nhập Tên Chủ TK cần loại trừ:", font=("Arial", 10, "bold")).pack(anchor="w")
        self.entry_stop_words = tk.Entry(f_stop, width=70, font=("Arial", 11))
        self.entry_stop_words.pack(anchor="w", ipady=3)
        
        self.btn_run_t1 = tk.Button(self.tab1, text="BẤT ĐẦU ĐỐI SOÁT EXCEL", bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), command=self.t1_chay, height=2)
        self.btn_run_t1.pack(fill="x", padx=25, pady=10)
        self.progress_var_t1 = tk.DoubleVar()
        self.progressbar_t1 = ttk.Progressbar(self.tab1, variable=self.progress_var_t1, maximum=100)
        self.progressbar_t1.pack(fill="x", padx=25, pady=5)
        self.lbl_status_t1 = tk.Label(self.tab1, text="Sẵn sàng...", font=("Arial", 9, "italic"))
        self.lbl_status_t1.pack()

    def t1_chay(self):
        if not self.file_saoke_path or not self.file_master_path: 
            return messagebox.showerror("Lỗi", "Vui lòng chọn đủ File Sao kê và File data MADTPN!")
        
        p_dc = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="KetQua_DoiChieu.xlsx")
        if not p_dc: return
        p_sk = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="SaoKe_DaCapNhat.xlsx")
        if not p_sk: return
        
        user_stops = self.entry_stop_words.get()
        enabled = [p for p, var in self.p_vars.items() if var.get()]
        custom_order = self.custom_order_entry.get().strip() or None
        
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
                    progress_callback=self.t1_update_progress, 
                    enabled_ps=enabled, 
                    custom_order=custom_order
                )
                self.root.after(0, lambda: messagebox.showinfo("Xong", "Đã đối soát thành công!"))
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Lỗi", f"Có lỗi xảy ra:\n{str(e)}"))
            finally:
                self.root.after(0, lambda: self.btn_run_t1.config(state="normal"))
                
        threading.Thread(target=task, daemon=True).start()

    def setup_tab2_interface(self):
        self.selected_images = []
        tk.Label(self.tab2, text="ỨNG DỤNG GOM HÌNH ẢNH THÀNH PDF CHỨNG TỪ", font=("Arial", 13, "bold"), fg="#2E7D32").pack(pady=15)
        f_btns = tk.Frame(self.tab2)
        f_btns.pack()
        tk.Button(f_btns, text="+ Chọn Thêm Ảnh", command=self.t3_chon_anh, width=32).grid(row=0, column=0, padx=5)
        tk.Button(f_btns, text="🗑 Xóa Toàn Bộ", command=self.t3_xoa_anh, width=25).grid(row=0, column=1, padx=5)
        self.txt_img_list = tk.Text(self.tab2, height=10, width=65, state="disabled", bg="#F5F5F5")
        self.txt_img_list.pack(pady=10)
        self.lbl_so_anh = tk.Label(self.tab2, text="Chưa chọn ảnh.", fg="gray", font=("Arial", 9))
        self.lbl_so_anh.pack()
        self.btn_run_t3 = tk.Button(self.tab2, text="XUẤT RA PDF CHỨNG TỪ", bg="#2E7D32", fg="white", font=("Arial", 11, "bold"), command=self.t3_chay, state="disabled", height=2)
        self.btn_run_t3.pack(fill="x", padx=25, pady=10)

    def t3_chon_anh(self):
        files = filedialog.askopenfilenames(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if files: self.selected_images.extend(files); self.t3_cap_nhat_ui()

    def t3_xoa_anh(self):
        self.selected_images = []
        self.t3_cap_nhat_ui()

    def t3_cap_nhat_ui(self):
        self.txt_img_list.config(state="normal")
        self.txt_img_list.delete("1.0", tk.END)
        for i, f_path in enumerate(self.selected_images):
            self.txt_img_list.insert(tk.END, f"{i+1}. {os.path.basename(f_path)}\n")
        self.txt_img_list.config(state="disabled")
        self.btn_run_t3.config(state="normal" if self.selected_images else "disabled")
        self.lbl_so_anh.config(text=f"Đã chọn {len(self.selected_images)} ảnh")

    def t3_chay(self):
        p_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="TaiLieu_ChungTu_ETAX.pdf")
        if p_pdf:
            self.btn_run_t3.config(state="disabled", text="ĐANG GHÉP VÀ LƯU PDF...")
            self.root.update()
            try:
                imgs = [Image.open(p).convert('RGB') for p in self.selected_images]
                if imgs:
                    if len(imgs) == 1: imgs[0].save(p_pdf, "PDF", resolution=100.0)
                    else: imgs[0].save(p_pdf, save_all=True, append_images=imgs[1:], resolution=100.0)
                    messagebox.showinfo("Thành công", f"Đã tạo PDF thành công!\n\n{p_pdf}")
                    self.t3_xoa_anh()
            except Exception as e:
                messagebox.showerror("Lỗi", str(e))
            finally:
                self.btn_run_t3.config(state="normal", text="XUẤT RA PDF CHỨNG TỪ")

    def t1_update_progress(self, percent, text):
        self.root.after(0, lambda: self.progress_var_t1.set(percent))
        self.root.after(0, lambda: self.lbl_status_t1.config(text=text))

    def chon_file(self, loai):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.csv")])
        if path:
            if loai == "sk1": self.file_saoke_path = path; self.lbl_saoke.config(text=os.path.basename(path), fg="blue")
            elif loai == "mst": self.file_master_path = path; self.lbl_master.config(text=os.path.basename(path), fg="blue")
            elif loai == "muavao": self.file_muavao_path = path; self.lbl_muavao.config(text=os.path.basename(path), fg="green")
            elif loai == "banra": self.file_banra_path = path; self.lbl_banra.config(text=os.path.basename(path), fg="green")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppGomNghiepVu(root)
    root.mainloop()