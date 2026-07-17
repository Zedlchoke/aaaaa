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

# ==========================================
# PHIÊN BẢN SửA GHI CHÚ + FORMAT EXCEL (17/07/2026)
# - Khi đã match (P1-P8) thì cột "CÁCH MATCH" để trống (không ghi Px)
# - File KetQua_DoiChieu.xlsx được format đẹp hơn: cột rộng, wrap_text tự động, dễ đọc
# - Giữ nguyên toàn bộ tính năng tùy chọn P1-P8 + thứ tự
# Cài đặt: pip install pandas openpyxl pillow unidecode
# Chạy: python app.py
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

def format_excel_sheet(ws, is_doichieu=False):
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    # Header
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Data rows
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    # === TỐI ƯU WIDTH + WRAP CHO KET QUA DOI CHIEU ===
    if is_doichieu:
        # Cột 1: THỨ TỰ DÒNG GỐC (nhỏ)
        ws.column_dimensions['A'].width = 8
        # Cột 2: DIỄN GIẢI GỐC (rất rộng + wrap)
        ws.column_dimensions['B'].width = 55
        # Cột 3-6: TÊN/HÓA ĐƠN/HỢP ĐỒNG/SỐ TIỀN QUÉT (vừa)
        for col in ['C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 22
        # Cột 7: TÊN MATCH (rộng)
        ws.column_dimensions['G'].width = 40
        # Cột 8: MÃ (vừa)
        ws.column_dimensions['H'].width = 18
        # Cột 9: CÁCH MATCH (nhỏ)
        ws.column_dimensions['I'].width = 12

        # Bật wrap_text rõ ràng cho các cột chính
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for col_idx in [2, 3, 4, 5, 6, 7]:  # B, C, D, E, F, G
                cell = row[col_idx - 1]
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    else:
        # Format bình thường cho các file khác
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 18

# ==========================================
# LÕI ĐỐI SOÁT SIÊU VIỆT P1 - P8 (TÙY CHỌN KIỂU + THỨ TỰ)
# ==========================================
def process_bank_data(file_saoke, file_master, path_save_doichieu, path_save_saokemoi, user_stop_str, file_muavao=None, file_banra=None, progress_callback=None, priority="name_first", enabled_ps=None, custom_order=None):
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

    target_list = []
    hd_index = {}
    if list_muavao or list_banra:
        target_list = list_banra + list_muavao
        for item in target_list:
            hd = str(item.get('SO_HD', '')).strip().lstrip('0') or '0'
            if hd != '0':
                if hd not in hd_index:
                    hd_index[hd] = []
                hd_index[hd].append(item)

    hd_pattern_comp = re.compile(r'\b(?:hoa don|hd)\s*(?:so\s*)?((?:\d+(?:\s*(?:,|\+|va\b|-)\s*)*)+)')

    if enabled_ps is None:
        enabled_ps = {"P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8"}
    else:
        enabled_ps = set([p.upper() for p in enabled_ps])

    if custom_order:
        try:
            order_list = [p.strip().upper() for p in str(custom_order).split(',') if p.strip()]
        except:
            order_list = None
    else:
        order_list = None

    all_matches = []
    total_rows = len(df_bank)

    for idx, row in df_bank.iterrows():
        if progress_callback and idx % 20 == 0: progress_callback(15 + (idx / total_rows) * 65, f"Đối soát dòng {idx+1}...")
        thu_tu_dong = idx + 2
        diengiai_goc = str(row.get(COL_DIENGIAI, ""))
        if pd.isna(diengiai_goc) or diengiai_goc.strip() == "" or diengiai_goc.lower().strip() == "nan":
            all_matches.append({"THỨ TỰ DÒNG GỐC": thu_tu_dong, "DIỄN GIẢI GỐC": "", "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI": "", "HỢP ĐỒNG QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "SỐ TIỀN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "", "TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN": "", "MÃ ĐỐI TƯỢNG PHÁP NHÂN": "", "CÁCH MATCH": "", "SCORE_NUM": 0})
            continue

        diengiai_norm = normalize_basic(diengiai_goc)
        diengiai_cleaned = get_core_name(diengiai_goc, dynamic_stops)
        diengiai_nospace = diengiai_norm.replace(" ", "") 
        amt_val = parse_amt_to_float(row.get('TTVND', 0))
        if amt_val == 0: amt_val = parse_amt_to_float(row.get('TTVND_TT', 0))

        matches_for_row = []

        def try_p1():
            if "P1" not in enabled_ps: return
            for master_item in master_list:
                core = master_item['norm_core']
                if len(core) < 3: continue
                p1_re = master_item.get('p1_regex')
                if p1_re and p1_re.search(diengiai_norm):
                    matches_for_row.append((1, len(core), master_item, core.upper()))
                    return

        def try_p2():
            if "P2" not in enabled_ps: return
            for master_item in master_list:
                core = master_item['norm_core']
                if len(diengiai_cleaned) >= 4 and re.search(r'\b' + r'\s+'.join(map(re.escape, diengiai_cleaned.split())) + r'\b', core):
                    matches_for_row.append((2, -len(core), master_item, diengiai_cleaned.upper()))
                    return

        def try_p3():
            if "P3" not in enabled_ps: return
            for master_item in master_list:
                core = master_item['norm_core']
                core_words = core.split()
                if len(core_words) >= 4:
                    acronym = "".join(w[0] for w in core_words)
                    if len(acronym) >= 4 and re.search(r'\b(?:' + '|'.join(ACTION_VERBS) + r')\b.*?\b' + re.escape(acronym) + r'\b', diengiai_norm):
                        matches_for_row.append((3, len(acronym), master_item, acronym.upper()))
                        return

        def try_p4():
            if "P4" not in enabled_ps: return
            for master_item in master_list:
                core = master_item['norm_core']
                core_words = core.split()
                p4_match = False
                for i in range(len(core_words)):
                    for j in range(i+1, len(core_words)+1):
                        chunk = "".join(core_words[i:j])
                        if ((j - i >= 2 and len(chunk) >= 5) or len(chunk) >= 8) and chunk in diengiai_nospace:
                            matches_for_row.append((4, len(chunk), master_item, chunk.upper()))
                            p4_match = True
                            break
                    if p4_match: break
                if p4_match: return

        def try_p5():
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
                    matches_for_row.append((5, int(best_sim*100), best_master, f"TÊN AI: {best_master['norm_core'].upper()}"))

        def try_p6_p7_p8():
            if not ("P6" in enabled_ps or "P7" in enabled_ps or "P8" in enabled_ps):
                return
            if not target_list or amt_val < 5000000:
                return
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
                        cands = hd_index.get(num_clean, [])
                        if cands:
                            mas_set = set()
                            for item in cands:
                                ma = item.get(COL_MA)
                                if ma:
                                    mas_set.add(ma)
                                    if ma not in best_item_map:
                                        best_item_map[ma] = item
                            if mas_set:
                                candidate_sets.append(mas_set)
                    if candidate_sets:
                        try:
                            common = set.intersection(*candidate_sets)
                            if len(common) == 1:
                                matched_ma = list(common)[0]
                                matches_for_row.append((6, len(nums), best_item_map[matched_ma], f"SỐ HĐ: {', '.join(nums)}"))
                                return
                        except:
                            pass

            if "P7" in enabled_ps and not matches_for_row:
                contracts_found = []
                explicit_contracts = re.findall(r'\bhop dong\s*(?:so\s*)?([a-z0-9/\-]+)\b', text_ext)
                contracts_found.extend([c for c in explicit_contracts if len(c) >= 4])
                explicit_contracts_spaced = re.findall(r'\bhop dong\s*(?:so\s*)?\d+\s+([a-z0-9/\-]+)\b', text_ext)
                contracts_found.extend([c for c in explicit_contracts_spaced if '/' in c or '-' in c])
                hd_matches = re.findall(r'\bhd\s*(?:so\s*)?([a-z0-9/\-]+)\b', text_ext)
                for val in hd_matches:
                    if '/' in val or '-' in val and len(val) >= 4:
                        contracts_found.append(val)
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
                            return
                    if matches_for_row: break

            if "P8" in enabled_ps and not matches_for_row:
                matched_companies = set()
                best_item = None
                for item in target_list:
                    if abs(parse_amt_to_float(item.get('TTVND', 0)) - amt_val) < 1.0 or abs(parse_amt_to_float(item.get('TTVND_TT', 0)) - amt_val) < 1.0:
                        matched_companies.add(item.get(COL_MA))
                        best_item = item
                if len(matched_companies) == 1:
                    matches_for_row.append((8, 0, best_item, f"SỐ TIỀN: {amt_val:,.0f}"))

        if order_list:
            for p in order_list:
                if p == "P1": try_p1()
                elif p == "P2": try_p2()
                elif p == "P3": try_p3()
                elif p == "P4": try_p4()
                elif p == "P5": try_p5()
                elif p in ("P6", "P7", "P8"): try_p6_p7_p8()
                if matches_for_row: break
        else:
            if priority == "smart_first":
                try_p6_p7_p8()
                if not matches_for_row:
                    try_p1()
                    try_p2()
                    try_p3()
                    try_p4()
                    try_p5()
            else:
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
            match_value = str(best_match[3])

            ten_quet = match_value if p_level <= 5 else ""
            hd_quet = match_value.replace("SỐ HĐ: ", "") if p_level == 6 else ""
            hopdong_quet = match_value.replace("HỢP ĐỒNG: ", "") if p_level == 7 else ""
            sotien_quet = match_value.replace("SỐ TIỀN: ", "") if p_level == 8 else ""

            # === SửA: Khi đã match thì CÁCH MATCH để trống ===
            all_matches.append({
                "THỨ TỰ DÒNG GỐC": thu_tu_dong,
                "DIỄN GIẢI GỐC": diengiai_goc,
                "TÊN QUÉT ĐƯỢC TRONG DIỄN GIẢI": ten_quet,
                "HÓA ĐƠN QUÉT ĐƯỢC TRONG DIỄN GIẢI": hd_quet,
                "HỢP ĐỒNG QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": hopdong_quet,
                "SỐ TIỀN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": sotien_quet,
                "TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN": best_match[2].get(COL_TEN_CTY, ""),
                "MÃ ĐỐI TƯỢNG PHÁP NHÂN": best_match[2].get(COL_MA, ""),
                "CÁCH MATCH": "",   # <--- Để trống khi đã match
                "SCORE_NUM": 100
            })
        else:
            all_matches.append({
                "THỨ TỰ DÒNG GỐC": thu_tu_dong,
                "DIỄN GIẢI GỐC": diengiai_goc,
                "TÊN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "",
                "HÓA ĐƠN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "",
                "HỢP ĐỒNG QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "",
                "SỐ TIỀN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI": "",
                "TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN": "",
                "MÃ ĐỐI TƯỢNG PHÁP NHÂN": "",
                "CÁCH MATCH": "",   # Không match
                "SCORE_NUM": 0
            })

    wb_doichieu = Workbook()
    ws = wb_doichieu.active
    ws.title = "Ket Qua Match"
    headers = ["THỨ TỰ DÒNG GỐC", "DIỄN GIẢI GỐC", "TÊN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI", "HÓA ĐƠN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI", "HỢP ĐỒNG QUÉT ĐƯỢC TRẠNG DIỄN GIẢI", "SỐ TIỀN QUÉT ĐƯỢC TRẠNG DIỄN GIẢI", "TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN", "MÃ ĐỐI TƯỢNG PHÁP NHÂN", "CÁCH MATCH"]
    ws.append(headers)

    for m in all_matches:
        ws.append([m.get(h, "") for h in headers])

    format_excel_sheet(ws, is_doichieu=True)   # <--- Truyền is_doichieu=True
    wb_doichieu.save(path_save_doichieu)

    # Cập nhật file SaoKê
    wb_saoke = load_workbook(file_saoke)
    ws = wb_saoke.active
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

            if tkco_val == "1121" or (tkco_val != "" and tkno_val == ""):
                ws.cell(row=r, column=col_madtpnno).value = m["MÃ ĐỐI TƯỢNG PHÁP NHÂN"]
            elif tkno_val == "1121" or (tkno_val != "" and tkco_val == ""):
                ws.cell(row=r, column=col_madtpnco).value = m["MÃ ĐỐI TƯỢNG PHÁP NHÂN"]
            else:
                ws.cell(row=r, column=col_madtpnno).value = m["MÃ ĐỐI TƯỢNG PHÁP NHÂN"]

            ws.cell(row=r, column=col_tenkh).value = m["TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN"]
            if col_ghichu:
                ws.cell(row=r, column=col_ghichu).value = m["CÁCH MATCH"]
        else:
            for c in range(1, ws.max_column + 1):
                ws.cell(row=r, column=c).fill = PatternFill(start_color="FFFF00", fill_type="solid")

    if progress_callback: progress_callback(98, "Đang lưu file...")
    wb_saoke.save(path_save_saokemoi)
    if progress_callback: progress_callback(100, "Hoàn tất thành công!")

# ==========================================
# CậP NH᫐T TỪ FILE SửA TAY
# ==========================================
def process_update_saoke(file_sk_old, file_dc_edited, path_save, progress_callback=None):
    if progress_callback: progress_callback(5, "Đang đọc file đối chiếu đã chỉnh sửa tay...")
    try:
        df_dc = pd.read_excel(file_dc_edited)
    except Exception as e:
        raise Exception(f"Không đọc được file đối chiếu: {e}")

    update_map = {}
    for _, row in df_dc.iterrows():
        thu_tu = row.get("THỨ TỰ DÒNG GỐC")
        ma_val = row.get("MÃ ĐỐI TƯỢNG PHÁP NHÂN")
        ten_val = row.get("TÊN MATCH ĐƯỢC TRẠNG FILE ĐỐI TƯỢNG PHÁP NHÂN", "")
        if pd.notna(thu_tu) and pd.notna(ma_val) and str(ma_val).strip():
            try:
                key = int(float(thu_tu))
                update_map[key] = {"MA": str(ma_val).strip(), "TEN": str(ten_val).strip()}
            except (ValueError, TypeError):
                continue

    if progress_callback: progress_callback(25, f"Tìm thấy {len(update_map)} dòng cần cập nhật...")

    if progress_callback: progress_callback(35, "Đang mở file Sao kê gốc...")
    try:
        wb = load_workbook(file_sk_old)
        ws = wb.active
    except Exception as e:
        raise Exception(f"Không mở được file Sao kê: {e}")

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

            tkno_val = str(ws.cell(row=r, column=col_tkno).value or "").strip().upper()
            tkco_val = str(ws.cell(row=r, column=col_tkco).value or "").strip().upper()

            if tkco_val == "1121" or (tkco_val and not tkno_val):
                ws.cell(row=r, column=col_madtpnno).value = ma
            elif tkno_val == "1121" or (tkno_val and not tkco_val):
                ws.cell(row=r, column=col_madtpnco).value = ma
            else:
                ws.cell(row=r, column=col_madtpnno).value = ma

            ws.cell(row=r, column=col_tenkh).value = ten

            if col_ghichu and col_ghichu <= ws.max_column:
                ws.cell(row=r, column=col_ghichu).value = "SửA TAY"

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

        f_files = tk.Frame(self.tab1); f_files.pack(fill="x", padx=20, pady=3)
        tk.Button(f_files, text="1. Chọn File Sao Kê (Excel)", command=lambda: self.chon_file("sk1"), width=25).grid(row=0, column=0, pady=3, padx=5)
        self.lbl_saoke = tk.Label(f_files, text="Chưa chọn file (Bắt buộc)...", fg="gray"); self.lbl_saoke.grid(row=0, column=1, sticky="w")
        tk.Button(f_files, text="2. Chọn Master MADTPN", command=lambda: self.chon_file("mst"), width=25).grid(row=1, column=0, pady=3, padx=5)
        self.lbl_master = tk.Label(f_files, text="Chưa chọn file (Bắt buộc)...", fg="gray"); self.lbl_master.grid(row=1, column=1, sticky="w")
        tk.Button(f_files, text="3. Chọn File MUA VÀO Năm", command=lambda: self.chon_file("muavao"), width=25).grid(row=2, column=0, pady=3, padx=5)
        self.lbl_muavao = tk.Label(f_files, text="Chưa chọn file (Không bắt buộc)...", fg="gray"); self.lbl_muavao.grid(row=2, column=1, sticky="w")
        tk.Button(f_files, text="4. Chọn File BÁN RA Năm", command=lambda: self.chon_file("banra"), width=25).grid(row=3, column=0, pady=3, padx=5)
        self.lbl_banra = tk.Label(f_files, text="Chưa chọn file (Không bắt buộc)...", fg="gray"); self.lbl_banra.grid(row=3, column=1, sticky="w")

        f_match = tk.LabelFrame(self.tab1, text=" TÙy chọn kiểu match (P1-P8) - Tick để kích hoạt ", padx=10, pady=5)
        f_match.pack(fill="x", padx=20, pady=5)

        self.p_vars = {}
        p_labels = {
            "P1": "P1 - Khớp tên nguyên vẹn (word boundary)",
            "P2": "P2 - Khớp ngược (tên nằm trong diễn giải)",
            "P3": "P3 - Khớp viết tắt (acronym)",
            "P4": "P4 - Khớp cụm từ (chunk)",
            "P5": "P5 - Fuzzy AI (bigram similarity ≥ 85%)",
            "P6": "P6 - Hóa đơn gộp (intersection unique MA)",
            "P7": "P7 - Hợp đỒng / MATHANG",
            "P8": "P8 - Số tiền duy nhất (khoảng < 1đ)"
        }

        for i, (p, label) in enumerate(p_labels.items()):
            row = i // 2
            col = i % 2
            var = tk.BooleanVar(value=True)
            self.p_vars[p] = var
            cb = ttk.Checkbutton(f_match, text=label, variable=var)
            cb.grid(row=row, column=col, sticky="w", padx=5, pady=2)

        f_order = tk.Frame(f_match)
        f_order.grid(row=4, column=0, columnspan=2, sticky="w", pady=5)
        tk.Label(f_order, text="Thứ tự tùy chỉnh (có thể để trống):", font=("Arial", 9)).pack(side="left")
        self.custom_order_entry = tk.Entry(f_order, width=35, font=("Arial", 9))
        self.custom_order_entry.pack(side="left", padx=5)
        tk.Label(f_order, text="(ví dụ: P6,P7,P8,P1,P5)", font=("Arial", 8), fg="gray").pack(side="left")

        f_priority = tk.Frame(self.tab1)
        f_priority.pack(fill="x", padx=20, pady=5)
        tk.Label(f_priority, text="Nếu không nhập thứ tự tùy chỉnh, dùng:", font=("Arial", 9, "bold")).pack(anchor="w")
        self.priority_var = tk.StringVar(value="name_first")
        ttk.Radiobutton(f_priority, text="✓ Ưu tiên P1-P5 (Tên) trước", variable=self.priority_var, value="name_first").pack(anchor="w")
        ttk.Radiobutton(f_priority, text="✓ Ưu tiên P6-P8 (HÓA ĐƠN/HỢP ĐỒNG/TIỀN) trước", variable=self.priority_var, value="smart_first").pack(anchor="w")

        f_stop = tk.Frame(self.tab1); f_stop.pack(fill="x", padx=20, pady=5)
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
        priority = self.priority_var.get()
        enabled = [p for p, var in self.p_vars.items() if var.get()]
        custom_order = self.custom_order_entry.get().strip() or None

        self.btn_run_t1.config(state="disabled"); self.progress_var_t1.set(0)
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
                    priority=priority,
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
        self.file_sk_cu = self.file_dc_sua = None
        tk.Label(self.tab2, text="CậP NH᫐T FILE SAO KÊ TỪ FILE CHỈNH SửA TAY (Từ KetQua_DoiChieu đã sửa)", font=("Arial", 12, "bold"), fg="#FF8C00").pack(pady=15)
        f_files = tk.Frame(self.tab2); f_files.pack(fill="x", padx=20, pady=10)
        tk.Button(f_files, text="1. File Sao Kê gốc (Cần cập nhật)", command=lambda: self.chon_file("sk3"), width=28).grid(row=0, column=0, pady=5, padx=5)
        self.lbl_sk_cu = tk.Label(f_files, text="Chưa chọn...", fg="gray"); self.lbl_sk_cu.grid(row=0, column=1, sticky="w")
        tk.Button(f_files, text="2. File Đối Chiếu đã Sửa Tay", command=lambda: self.chon_file("dc3"), width=28).grid(row=1, column=0, pady=5, padx=5)
        self.lbl_dc_sua = tk.Label(f_files, text="Chưa chọn...", fg="gray"); self.lbl_dc_sua.grid(row=1, column=1, sticky="w")
        self.btn_run_t2 = tk.Button(self.tab2, text="TIẾN HÀNH CậP NH᫐T ĐÈ MÃ (CHỈNH SửA TAY)", bg="#FF8C00", fg="white", font=("Arial", 11, "bold"), command=self.t2_chay, height=2)
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
                self.root.after(0, lambda: messagebox.showinfo("Thành công", "Đã cập nhật dữ liệu sửa tay thành công!\n\u0110ã lưu file: " + p_save))
            except Exception as e:
                self.root.after(0, lambda e=e: messagebox.showerror("Lỗi", str(e)))
            finally:
                self.root.after(0, lambda: self.btn_run_t2.config(state="normal"))
        threading.Thread(target=task, daemon=True).start()

    def setup_tab3_interface(self):
        self.selected_images = []
        tk.Label(self.tab3, text="ỨNG DỤNG GOM HÌNH ẢNH THÀNH PDF CHỨNG TỪ", font=("Arial", 13, "bold"), fg="#2E7D32").pack(pady=15)
        f_btns = tk.Frame(self.tab3); f_btns.pack()
        tk.Button(f_btns, text="+ Chọn Thêm Ảnh (Chọn theo thứ tự trang muốn)", command=self.t3_chon_anh, width=32).grid(row=0, column=0, padx=5)
        tk.Button(f_btns, text="🗑 Xóa Toàn Bộ Danh Sách", command=self.t3_xoa_anh, width=25).grid(row=0, column=1, padx=5)
        self.txt_img_list = tk.Text(self.tab3, height=10, width=65, state="disabled", bg="#F5F5F5"); self.txt_img_list.pack(pady=10)
        self.lbl_so_anh = tk.Label(self.tab3, text="Chưa chọn ảnh. Sau khi chọn, thứ tự trong list = thứ tự trang trong PDF.", fg="gray", font=("Arial", 9)); self.lbl_so_anh.pack()
        self.btn_run_t3 = tk.Button(self.tab3, text="XUẤT RA PDF CHứNG TỪ", bg="#2E7D32", fg="white", font=("Arial", 11, "bold"), command=self.t3_chay, state="disabled", height=2)
        self.btn_run_t3.pack(fill="x", padx=25, pady=10)

    def t3_chon_anh(self):
        files = filedialog.askopenfilenames(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if files: self.selected_images.extend(files); self.t3_cap_nhat_ui()

    def t3_xoa_anh(self): 
        self.selected_images = []; self.t3_cap_nhat_ui()

    def t3_cap_nhat_ui(self):
        self.txt_img_list.config(state="normal"); self.txt_img_list.delete("1.0", tk.END)
        for i, f_path in enumerate(self.selected_images): 
            self.txt_img_list.insert(tk.END, f"{i+1}. {os.path.basename(f_path)}\n")
        self.txt_img_list.config(state="disabled")
        self.btn_run_t3.config(state="normal" if self.selected_images else "disabled")
        self.lbl_so_anh.config(text=f"Đã chọn {len(self.selected_images)} ảnh | Thứ tự list = thứ tự trang trong PDF")

    def t3_chay(self):
        p_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile="TaiLieu_ChungTu_ETAX.pdf")
        if p_pdf:
            self.btn_run_t3.config(state="disabled", text="ĐANG GHÉP VÀ LƯU PDF..."); self.root.update()
            try:
                imgs = [Image.open(p).convert('RGB') for p in self.selected_images]
                if imgs:
                    if len(imgs) == 1:
                        imgs[0].save(p_pdf, "PDF", resolution=100.0)
                    else:
                        imgs[0].save(p_pdf, save_all=True, append_images=imgs[1:], resolution=100.0)
                    messagebox.showinfo("Thành công", f"Đã tạo file PDF chứng từ thành công!\n\n{p_pdf}\nTổng số trang: {len(imgs)}")
                    self.t3_xoa_anh()
                else:
                    messagebox.showwarning("Thông báo", "Danh sách ảnh trống.")
            except Exception as e: 
                messagebox.showerror("Lỗi tạo PDF", f"{str(e)}\n\nKhuyến nghị: pip install Pillow")
            finally: 
                self.btn_run_t3.config(state="normal", text="XUẤT RA PDF CHứNG TỪ")

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