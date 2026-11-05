import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import io
import re
from copy import copy
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import matplotlib.pyplot as plt
from difflib import SequenceMatcher

# Pengaman untuk Library PowerPoint
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="RAB vs LRA", layout="wide", page_icon="📊")

def normalize_text(text):
    t = str(text)
    t = re.sub(r'^[\w\.]+\s*-\s*', '', t)
    t = re.sub(r'^[a-zA-Z0-9]+\.\s*', '', t)
    t = t.split('|')[0]
    t = t.replace('-', '').strip().lower().replace(',', '')
    return t

def safe_float(val):
    if pd.isna(val): return 0.0
    try:
        if isinstance(val, str):
            val = val.replace(',', '').strip()
            if val == '-' or val == '': return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def get_similarity_score(t1, t2):
    if not t1 or not t2: return 0.0
    clean_t1 = normalize_text(t1)
    clean_t2 = normalize_text(t2)
    
    if clean_t1 == clean_t2: return 1.0
    
    if len(clean_t1) < 3 or len(clean_t2) < 3:
        return 1.0 if clean_t1 == clean_t2 else 0.0
        
    sim = SequenceMatcher(None, clean_t1, clean_t2).ratio()
    
    w1 = set(clean_t1.split())
    w2 = set(clean_t2.split())
    if len(w1) > 0 and len(w2) > 0:
        shorter = w1 if len(w1) < len(w2) else w2
        longer = w2 if len(w1) < len(w2) else w1
        overlap = len(shorter.intersection(longer))
        
        overlap_ratio = overlap / len(shorter)
        length_penalty = len(shorter) / len(longer) 
        
        combined_score = (overlap_ratio * 0.75) + (length_penalty * 0.25)
        
        sim = max(sim, combined_score)
        
    return sim

@st.cache_data
def parse_lra_files_cached(lra_files_input, list_semua_bulan):
    data_realisasi = {}
    monthly_totals = {b: 0 for b in list_semua_bulan}
    component_metrics = {}
    sub_component_realisasi = {}
    sub_component_sisa = {} 
    
    for fname, fbytes in lra_files_input:
        file_obj = io.BytesIO(fbytes)
        df_raw = pd.read_excel(file_obj, header=None, nrows=15)
        file_obj.seek(0)
        
        bulan_file = None
        is_outstanding_file = False
        
        fname_upper = fname.upper()
        for b in list_semua_bulan:
            if b in fname_upper:
                bulan_file = b
                break

        if not bulan_file:
            for r in range(len(df_raw)):
                row_text_raw = " ".join(str(val) for val in df_raw.iloc[r].values if pd.notna(val)).upper()
                for b in list_semua_bulan:
                    if b in row_text_raw:
                        bulan_file = b
                        break
                if bulan_file: break

        fname_clean = fname_upper.replace(" ", "")
        keyword_out = ["SEMUA", "LEVEL", "ALL", "PERIODE", "REKAP", "GUP"]
        if any(kw in fname_clean for kw in keyword_out) and not bulan_file:
            is_outstanding_file = True

        for r in range(len(df_raw)):
            row_text_raw = " ".join(str(val) for val in df_raw.iloc[r].values if pd.notna(val)).upper()
            if any(k in row_text_raw for k in ["SEMUA", "LEVEL", "ALL", "PERIODE"]):
                if not bulan_file:
                    is_outstanding_file = True
                    break

        df_lra = pd.read_excel(file_obj, skiprows=5)
        
        kolom_lra = [str(col).upper() for col in df_lra.columns]
        ada_gup_spm = any('GUP' in col or 'SPM' in col or 'VERIFIKASI' in col for col in kolom_lra)
        if ada_gup_spm and not bulan_file:
            is_outstanding_file = True

        if not bulan_file and not is_outstanding_file: continue
            
        cur_komp, cur_sub, cur_akun = "GLOBAL", "GLOBAL", "GLOBAL"
        
        for index, row in df_lra.iterrows():
            lvl = str(row.get('Level')).strip()
            uraian = str(row.get('Kode / Uraian', '')).strip()
            
            realisasi = safe_float(row.get('Total Realisasi'))
            outstanding_val = safe_float(row.get('GUP')) + safe_float(row.get('SPM')) + safe_float(row.get('Verifikasi'))
                
            if lvl == 'Komponen':
                match = re.search(r'(\d{3})\s*-', uraian)
                if match: cur_komp = match.group(1)
                cur_sub, cur_akun = "GLOBAL", "GLOBAL"
            elif lvl == 'Sub Komponen':
                match = re.search(r'([A-Z])\s*-', uraian)
                if match: cur_sub = match.group(1)
                cur_akun = "GLOBAL"
            elif lvl == 'Akun':
                match = re.search(r'(\d{6})\s*-', uraian)
                if match: cur_akun = match.group(1)
            
            kamar_unik = f"{cur_komp}_{cur_sub}_{cur_akun}"
            if kamar_unik not in data_realisasi: data_realisasi[kamar_unik] = {}
                
            if pd.notna(uraian) and uraian not in ['nan', '']:
                norm_lra = normalize_text(uraian)
                
                if norm_lra not in data_realisasi[kamar_unik]:
                    data_realisasi[kamar_unik][norm_lra] = {b: 0 for b in list_semua_bulan}
                    data_realisasi[kamar_unik][norm_lra]['OUTSTANDING'] = 0
                
                if bulan_file:
                    data_realisasi[kamar_unik][norm_lra][bulan_file] += realisasi
                    monthly_totals[bulan_file] += realisasi
                if is_outstanding_file:
                    data_realisasi[kamar_unik][norm_lra]['OUTSTANDING'] += outstanding_val
                
    return data_realisasi, monthly_totals

@st.cache_data
def process_rab_lra_cached(rab_bytes, data_realisasi_tuple, list_semua_bulan):
    file_rab_io = io.BytesIO(rab_bytes)
    wb = load_workbook(file_rab_io)
    ws = wb.active 
    
    # -----------------------------------------------------
    # MENGUBAH NAMA SHEET MENJADI "DATA"
    # -----------------------------------------------------
    ws.title = "DATA"
    
    # Header yang diharapkan ada di baris 13 (termasuk S: Jumlah, T: Blokir, U: Pagu setelah Blokir)
    baris_header = 13 
    baris_mulai_data = 14
    max_row = ws.max_row
    
    # Kolom V (22) dan seterusnya
    kolom_baru = ["TOTAL Realisasi", "SISA", "OUT STANDING", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER", "KETERANGAN"]
    
    warna_hijau = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    warna_kuning = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    warna_orange = PatternFill(start_color="FCD5B4", end_color="FCD5B4", fill_type="solid") 
    header_ref = ws.cell(row=baris_header, column=21) # Referensi format dari kolom U (Pagu setelah blokir)
    
    start_col = 22
    for i, nama_kolom in enumerate(kolom_baru):
        cell = ws.cell(row=baris_header, column=start_col + i)
        cell.value = nama_kolom
        if header_ref.has_style:
            cell.font, cell.border, cell.alignment = copy(header_ref.font), copy(header_ref.border), copy(header_ref.alignment)
        if nama_kolom == "TOTAL Realisasi": cell.fill = warna_hijau
        elif nama_kolom == "SISA": cell.fill = warna_kuning
        elif nama_kolom == "OUT STANDING": cell.fill = warna_orange
        else:
            if header_ref.has_style: cell.fill = copy(header_ref.fill)

    # Lebar kolom
    ws.column_dimensions[get_column_letter(22)].width = 20.5 # V: Total Realisasi
    ws.column_dimensions[get_column_letter(23)].width = 18.2 # W: Sisa
    ws.column_dimensions[get_column_letter(24)].width = 18.2 # X: Outstanding
    for c_idx in range(25, 37): ws.column_dimensions[get_column_letter(c_idx)].width = 15.0 # Y s/d AJ
    ws.column_dimensions[get_column_letter(37)].width = 35.0 # AK: Keterangan

    data_ditemukan = 0
    rab_komp, rab_sub, rab_akun = "GLOBAL", "GLOBAL", "GLOBAL"
    baris_terpakai = set()
    
    import copy as cp
    data_realisasi = cp.deepcopy(data_realisasi_tuple)
    
    # -----------------------------------------------------
    # STRUKTUR PENYIMPANAN DATA UNTUK PPT & DASHBOARD
    # -----------------------------------------------------
    satker_summary = {"pagu": 0, "blokir": 0, "realisasi": 0, "sisa": 0, "outstanding": 0}
    component_metrics = {}
    sub_component_realisasi = {}
    sub_component_sisa = {} 
    
    # Ambil nilai total Satker (Diasumsikan ada di baris 14, di mana level Satker berada)
    # Anda perlu menyesuaikan ini jika letak grand total RAB tidak selalu di baris 14
    satker_pagu = safe_float(ws.cell(row=14, column=19).value)
    satker_blokir = safe_float(ws.cell(row=14, column=20).value)
    
    # Variabel pelacakan level
    current_komponen_name = ""
    current_komponen_kode = ""
    current_subkomponen_name = ""

    for row_idx in range(baris_mulai_data, max_row + 1):
        kode_col = str(ws.cell(row=row_idx, column=2).value).strip()
        sub_col = str(ws.cell(row=row_idx, column=4).value).strip()
        
        # Ekstrak data Pagu & Blokir per baris (dari file RAB)
        row_pagu = safe_float(ws.cell(row=row_idx, column=19).value)
        row_blokir = safe_float(ws.cell(row=row_idx, column=20).value)
        
        # Deteksi Komponen (Level: 3 digit angka)
        if kode_col.isdigit() and len(kode_col) == 3: 
            rab_komp = kode_col
            rab_sub, rab_akun = "GLOBAL", "GLOBAL"
            current_komponen_kode = kode_col
            
            uraian_komp = " ".join([str(ws.cell(row=row_idx, column=c).value).strip() for c in range(3, 7) if ws.cell(row=row_idx, column=c).value and str(ws.cell(row=row_idx, column=c).value).strip() not in ['-', '']])
            current_komponen_name = f"{kode_col} - {uraian_komp}"
            
            if current_komponen_name not in component_metrics:
                component_metrics[current_komponen_name] = {"pagu": row_pagu, "blokir": row_blokir, "realisasi": 0, "sisa": 0, "persen": 0}

        # Deteksi Sub Komponen (Level: 1 huruf besar)
        elif re.match(r'^[A-Z]\.?\s*$', sub_col):
            rab_sub = sub_col.replace('.', '').strip()
            rab_akun = "GLOBAL" 
            
            uraian_sub = " ".join([str(ws.cell(row=row_idx, column=c).value).strip() for c in range(5, 7) if ws.cell(row=row_idx, column=c).value and str(ws.cell(row=row_idx, column=c).value).strip() not in ['-', '']])
            current_subkomponen_name = f"{rab_sub} - {uraian_sub}"
            
            if current_komponen_name:
                if current_komponen_name not in sub_component_realisasi:
                    sub_component_realisasi[current_komponen_name] = {}
                    sub_component_sisa[current_komponen_name] = {}
                sub_component_realisasi[current_komponen_name][current_subkomponen_name] = 0
                sub_component_sisa[current_komponen_name][current_subkomponen_name] = row_pagu - row_blokir # Sisa awal = Pagu - Blokir
                
        elif kode_col.isdigit() and len(kode_col) == 6: 
            rab_akun = kode_col
            
        kamar_rab_saat_ini = f"{rab_komp}_{rab_sub}_{rab_akun}"
        
        bagian_teks = []
        for col_idx in range(3, 7):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val and isinstance(val, str) and str(val).strip() not in ['-', '']: bagian_teks.append(str(val).strip())
        
        # Proses pencocokan LRA
        if bagian_teks:
            uraian_rab = " ".join(bagian_teks)
            best_match = None
            highest_score = 0
            kamar_opsi = [kamar_rab_saat_ini, f"{rab_komp}_{rab_sub}_GLOBAL", f"{rab_komp}_GLOBAL_GLOBAL"]
            
            for kamar in kamar_opsi:
                if kamar in data_realisasi:
                    for key, dict_bulanan in data_realisasi[kamar].items():
                        score = get_similarity_score(key, uraian_rab)
                        if score > highest_score and score >= 0.80:
                            highest_score = score
                            best_match = (kamar, key)
                        if highest_score == 1.0: break
                if highest_score == 1.0: break
            
            # Jika Cocok, Tulis Realisasi Bulanan ke Excel
            if best_match:
                kamar_ketemu, key_ketemu = best_match
                dict_bulanan = data_realisasi[kamar_ketemu].pop(key_ketemu) 
                nilai_outstanding = dict_bulanan.pop('OUTSTANDING', 0)
                
                target_row = row_idx
                for r_cek in range(row_idx, min(row_idx + 6, max_row + 1)):
                    if ws.cell(row=r_cek, column=19).value is not None and r_cek not in baris_terpakai:
                        target_row = r_cek
                        break
                baris_terpakai.add(target_row)
                
                ws.cell(row=target_row, column=24).value = nilai_outstanding # X: Outstanding
                total_realisasi_baris = nilai_outstanding
                for idx_b, b_name in enumerate(list_semua_bulan):
                    val_b = dict_bulanan[b_name]
                    ws.cell(row=target_row, column=25 + idx_b).value = val_b # Y sd AJ
                    total_realisasi_baris += val_b
                
                # Menulis formula Excel untuk rekap per baris
                ws.cell(row=target_row, column=22).value = f"=SUM(Y{target_row}:AJ{target_row})+X{target_row}" # V: Total Realisasi
                ws.cell(row=target_row, column=23).value = f"=U{target_row}-V{target_row}" # W: Sisa (Pagu Setelah Blokir - Total Realisasi)

                data_ref = ws.cell(row=target_row, column=21)
                for c_idx in range(22, 38):
                    c = ws.cell(row=target_row, column=c_idx)
                    if data_ref.has_style:
                        c.font, c.border, c.alignment, c.number_format = copy(data_ref.font), copy(data_ref.border), copy(data_ref.alignment), copy(data_ref.number_format)
                data_ditemukan += 1
                
                # --- AKUMULASI UNTUK PPT & DASHBOARD ---
                satker_summary["realisasi"] += total_realisasi_baris
                satker_summary["outstanding"] += nilai_outstanding
                
                if current_komponen_name in component_metrics:
                    component_metrics[current_komponen_name]["realisasi"] += total_realisasi_baris
                    
                if current_komponen_name in sub_component_realisasi and current_subkomponen_name:
                    if current_subkomponen_name in sub_component_realisasi[current_komponen_name]:
                        sub_component_realisasi[current_komponen_name][current_subkomponen_name] += total_realisasi_baris
                        
    # --- FINALISASI PERHITUNGAN SUMMARY PPT ---
    satker_summary["pagu"] = satker_pagu
    satker_summary["blokir"] = satker_blokir
    satker_summary["sisa"] = satker_pagu - satker_blokir - satker_summary["realisasi"]
    
    for komp_name, m in component_metrics.items():
        m["sisa"] = m["pagu"] - m["blokir"] - m["realisasi"]
        m["persen"] = (m["realisasi"] / (m["pagu"] - m["blokir"]) * 100) if (m["pagu"] - m["blokir"]) > 0 else 0
        
    for komp_name, subs in sub_component_sisa.items():
        for sub_name in subs:
            realisasi_sub = sub_component_realisasi[komp_name].get(sub_name, 0)
            sub_component_sisa[komp_name][sub_name] -= realisasi_sub # Sisa Akhir = Sisa Awal (Pagu-Blokir) - Realisasi

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue(), data_ditemukan, satker_summary, component_metrics, sub_component_realisasi, sub_component_sisa

# ==========================================
# FUNGSI PEMBUATAN FILE POWERPOINT (PPTX) 
# ==========================================
def create_powerpoint_report(satker_summary, component_metrics, monthly_totals, sub_component_realisasi, sub_component_sisa):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    DEEP_TEAL = RGBColor(0, 64, 64)       
    GOLD_ACCENT = RGBColor(218, 165, 32)  
    GREY_TEXT = RGBColor(80, 90, 100)     
    WHITE = RGBColor(255, 255, 255)       
    LIGHT_BG = RGBColor(245, 247, 250)    

    def add_slide_header(slide, title_text):
        t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.8))
        p = t_box.text_frame.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.name = "Arial"
        p.font.color.rgb = DEEP_TEAL
        
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.2), Inches(11.7), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = GOLD_ACCENT
        line.line.fill.background()
    
    # 1. SLIDE COVER (JUDUL)
    slide1 = prs.slides.add_slide(blank_layout)
    bg = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid(); bg.fill.fore_color.rgb = DEEP_TEAL; bg.line.fill.background()
    
    accent = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(4.2), Inches(1.5), Inches(0.08))
    accent.fill.solid(); accent.fill.fore_color.rgb = GOLD_ACCENT; accent.line.fill.background()

    tf1 = slide1.shapes.add_textbox(Inches(0.9), Inches(2.2), Inches(11), Inches(2)).text_frame
    tf1.word_wrap = True
    
    p1 = tf1.paragraphs[0]
    p1.text = "LAPORAN EKSEKUTIF"
    p1.font.size, p1.font.bold, p1.font.name, p1.font.color.rgb = Pt(48), True, "Arial", WHITE
    
    p2 = tf1.add_paragraph()
    p2.text = "Konsolidasi Anggaran (RAB vs LRA)"
    p2.font.size, p2.font.name, p2.font.color.rgb = Pt(24), "Arial", RGBColor(180, 205, 205)
    
    # 2. SLIDE RINGKASAN EKSEKUTIF 
    slide2 = prs.slides.add_slide(blank_layout)
    add_slide_header(slide2, "Ringkasan Kinerja Anggaran")
    
    total_pagu = satker_summary["pagu"]
    total_blokir = satker_summary["blokir"]
    total_realisasi = satker_summary["realisasi"]
    total_sisa = satker_summary["sisa"]
    pagu_efektif = total_pagu - total_blokir
    persen_nasional = (total_realisasi / pagu_efektif * 100) if pagu_efektif > 0 else 0
    
    metrics = [
        ("Total Pagu", f"Rp {total_pagu:,.0f}"),
        ("Diblokir", f"Rp {total_blokir:,.0f}"),
        ("Total Realisasi", f"Rp {total_realisasi:,.0f}"),
        ("Sisa Efektif", f"Rp {total_sisa:,.0f}"),
        ("Penyerapan", f"{persen_nasional:.2f}%")
    ]
    
    card_width = 2.2
    for i, (label, value) in enumerate(metrics):
        card = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8 + (i * 2.35)), Inches(1.5), Inches(card_width), Inches(1.1))
        card.fill.solid(); card.fill.fore_color.rgb = LIGHT_BG; card.line.color.rgb = RGBColor(220, 225, 230)
        c_tf = card.text_frame; c_tf.clear()
        
        c_p1 = c_tf.paragraphs[0]
        c_p1.text, c_p1.font.size, c_p1.font.color.rgb = label, Pt(11), GREY_TEXT
        
        c_p2 = c_tf.add_paragraph()
        c_p2.text, c_p2.font.size, c_p2.font.bold, c_p2.font.color.rgb = value, Pt(16), True, DEEP_TEAL
    
    if component_metrics:
        p_c = slide2.shapes.add_textbox(Inches(0.8), Inches(2.9), Inches(11.7), Inches(0.5)).text_frame.paragraphs[0]
        p_c.text, p_c.font.size, p_c.font.bold, p_c.font.color.rgb = "Rincian per Komponen Utama:", Pt(16), True, DEEP_TEAL
        
        rows = len(component_metrics) + 1
        cols = 5
        c_table = slide2.shapes.add_table(rows, cols, Inches(0.8), Inches(3.4), Inches(11.7), Inches(0.4 * rows)).table
        c_table.columns[0].width, c_table.columns[1].width, c_table.columns[2].width, c_table.columns[3].width, c_table.columns[4].width = Inches(4.5), Inches(2.0), Inches(2.0), Inches(2.0), Inches(1.2)
        
        headers = ["Komponen Utama", "Pagu Efektif (Rp)", "Total Realisasi (Rp)", "Sisa Anggaran (Rp)", "% Realisasi"]
        for col_idx, h_text in enumerate(headers):
            cell = c_table.cell(0, col_idx); cell.text = h_text
            cell.fill.solid(); cell.fill.fore_color.rgb = DEEP_TEAL
            p = cell.text_frame.paragraphs[0]
            p.font.bold, p.font.color.rgb, p.font.size = True, WHITE, Pt(12)
            
        for r_idx, (komp_name, m) in enumerate(component_metrics.items(), start=1):
            row_color = WHITE if r_idx % 2 != 0 else LIGHT_BG
            pagu_efektif_komp = m["pagu"] - m["blokir"]
            data = [komp_name, f"Rp {pagu_efektif_komp:,.0f}", f"Rp {m['realisasi']:,.0f}", f"Rp {m['sisa']:,.0f}", f"{m['persen']:.2f}%"]
            for c_idx, val in enumerate(data):
                cell = c_table.cell(r_idx, c_idx); cell.text = val
                cell.fill.solid(); cell.fill.fore_color.rgb = row_color
                p = cell.text_frame.paragraphs[0]
                p.font.color.rgb, p.font.size = GREY_TEXT, Pt(11)

    # 3. SLIDE GRAFIK TREN BULANAN
    slide3 = prs.slides.add_slide(blank_layout)
    add_slide_header(slide3, "Tren Penyerapan Anggaran Bulanan")
    
    plt.rcParams['font.family'] = 'sans-serif'
    fig_m, ax_m = plt.subplots(figsize=(12, 5))
    
    months = list(monthly_totals.keys())
    values = list(monthly_totals.values())
    ax_m.bar(months, [v / 1e9 for v in values], color='#008080', width=0.6) 
    
    ax_m.spines['top'].set_visible(False)
    ax_m.spines['right'].set_visible(False)
    ax_m.spines['left'].set_color('#DDDDDD')
    ax_m.spines['bottom'].set_color('#DDDDDD')
    ax_m.tick_params(bottom=False, left=False)
    
    ax_m.set_ylabel('Realisasi (Miliar Rp)', fontsize=12, fontweight='bold', color='#333333')
    plt.xticks(rotation=30, ha='right', fontsize=11, color='#555555')
    plt.yticks(fontsize=11, color='#555555')
    ax_m.grid(axis='y', linestyle='--', alpha=0.4)
    
    plt.tight_layout()
    img_buf = io.BytesIO(); fig_m.savefig(img_buf, format='png', dpi=300, bbox_inches='tight', transparent=True); plt.close(fig_m); img_buf.seek(0)
    slide3.shapes.add_picture(img_buf, Inches(0.5), Inches(1.5), width=Inches(12.0))
    
    # 4. SLIDE RINCIAN SUB KOMPONEN
    semua_komponen_unik = sorted(list(set(list(sub_component_realisasi.keys()) + list(sub_component_sisa.keys()))))
    
    for komp_name in semua_komponen_unik:
        if komp_name in sub_component_realisasi:
            sub_dict_real_table = {k: float(v) for k, v in sub_component_realisasi[komp_name].items() if pd.notna(v) and float(v) >= 0}
            if sub_dict_real_table:
                slide_sub = prs.slides.add_slide(blank_layout)
                add_slide_header(slide_sub, f"Realisasi Sub Komponen: {komp_name}")
                
                sub_labels_chart = [k for k, v in sub_dict_real_table.items() if v > 0]
                sub_vals_chart = [v for v in sub_dict_real_table.values() if v > 0]
                total_komp_chart = sum(sub_vals_chart)
                
                if total_komp_chart > 0:
                    fig_p, ax_p = plt.subplots(figsize=(6, 5))
                    colors_real = ['#004040', '#006666', '#008080', '#20B2AA', '#48D1CC', '#5F9EA0', '#8FBC8F']
                    labels_with_pct = [f"Sub {l.split('-')[0].strip()}\n({val/total_komp_chart*100:.1f}%)" for l, val in zip(sub_labels_chart, sub_vals_chart)]
                    wedges, texts = ax_p.pie(sub_vals_chart, labels=labels_with_pct, startangle=90, colors=colors_real, wedgeprops=dict(width=0.4, edgecolor='w'), labeldistance=1.05)
                    plt.setp(texts, size=10, weight="bold", color="#333333")
                    ax_p.axis('equal'); plt.tight_layout()
                    pie_buf = io.BytesIO(); fig_p.savefig(pie_buf, format='png', dpi=300, bbox_inches='tight', transparent=True); plt.close(fig_p); pie_buf.seek(0)
                    slide_sub.shapes.add_picture(pie_buf, Inches(0.5), Inches(1.8), width=Inches(5.0))
                else:
                    t_box_empty = slide_sub.shapes.add_textbox(Inches(1.5), Inches(3.0), Inches(4.0), Inches(1.0)).text_frame
                    t_box_empty.text = "Seluruh sub komponen belum memiliki realisasi (Rp 0)"
                
                sub_labels_table = list(sub_dict_real_table.keys())
                sub_vals_table = list(sub_dict_real_table.values())
                total_komp_table = sum(sub_vals_table)
                
                rows = len(sub_dict_real_table) + 1
                table = slide_sub.shapes.add_table(rows, 3, Inches(5.8), Inches(1.8), Inches(7.0), Inches(0.4 * rows)).table
                table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(3.5), Inches(2.0), Inches(1.5)
                
                for col_idx, h_text in enumerate(["Sub Komponen", "Total Realisasi (Rp)", "Persentase"]):
                    cell = table.cell(0, col_idx); cell.text = h_text
                    cell.fill.solid(); cell.fill.fore_color.rgb = DEEP_TEAL
                    p = cell.text_frame.paragraphs[0]
                    p.font.bold, p.font.color.rgb, p.font.size = True, WHITE, Pt(12)
                
                for r_idx, (lbl, val) in enumerate(zip(sub_labels_table, sub_vals_table), start=1):
                    pct = (val / total_komp_table * 100) if total_komp_table > 0 else 0
                    row_color = WHITE if r_idx % 2 != 0 else LIGHT_BG
                    for c_idx, text_val in enumerate([lbl, f"Rp {val:,.0f}", f"{pct:.2f}%"]):
                        cell = table.cell(r_idx, c_idx); cell.text = text_val
                        cell.fill.solid(); cell.fill.fore_color.rgb = row_color
                        p = cell.text_frame.paragraphs[0]
                        p.font.color.rgb, p.font.size = GREY_TEXT, Pt(11)

        if komp_name in sub_component_sisa:
            sub_dict_sisa_table = {k: float(v) for k, v in sub_component_sisa[komp_name].items() if pd.notna(v) and float(v) >= 0}
            if sub_dict_sisa_table:
                slide_sisa = prs.slides.add_slide(blank_layout)
                add_slide_header(slide_sisa, f"Sisa Anggaran Sub Komponen: {komp_name}")
                
                sub_labels_s_chart = [k for k, v in sub_dict_sisa_table.items() if v > 0]
                sub_vals_s_chart = [v for v in sub_dict_sisa_table.values() if v > 0]
                total_komp_s_chart = sum(sub_vals_s_chart)
                
                if total_komp_s_chart > 0:
                    fig_s, ax_s = plt.subplots(figsize=(6, 5))
                    colors_sisa = ['#DAA520', '#CD853F', '#D2691E', '#B8860B', '#8B4513', '#A0522D', '#D2B48C']
                    labels_s_with_pct = [f"Sub {l.split('-')[0].strip()}\n({val/total_komp_s_chart*100:.1f}%)" for l, val in zip(sub_labels_s_chart, sub_vals_s_chart)]
                    wedges_s, texts_s = ax_s.pie(sub_vals_s_chart, labels=labels_s_with_pct, startangle=90, colors=colors_sisa, wedgeprops=dict(width=0.4, edgecolor='w'), labeldistance=1.05)
                    plt.setp(texts_s, size=10, weight="bold", color="#333333")
                    ax_s.axis('equal'); plt.tight_layout()
                    pie_buf_s = io.BytesIO(); fig_s.savefig(pie_buf_s, format='png', dpi=300, bbox_inches='tight', transparent=True); plt.close(fig_s); pie_buf_s.seek(0)
                    slide_sisa.shapes.add_picture(pie_buf_s, Inches(0.5), Inches(1.8), width=Inches(5.0))
                else:
                    t_box_empty_s = slide_sisa.shapes.add_textbox(Inches(1.5), Inches(3.0), Inches(4.0), Inches(1.0)).text_frame
                    t_box_empty_s.text = "Seluruh anggaran telah terealisasi (Sisa Rp 0)"
                
                sub_labels_s_table = list(sub_dict_sisa_table.keys())
                sub_vals_s_table = list(sub_dict_sisa_table.values())
                total_komp_s_table = sum(sub_vals_s_table)
                
                rows_s = len(sub_dict_sisa_table) + 1
                table_s = slide_sisa.shapes.add_table(rows_s, 3, Inches(5.8), Inches(1.8), Inches(7.0), Inches(0.4 * rows_s)).table
                table_s.columns[0].width, table_s.columns[1].width, table_s.columns[2].width = Inches(3.5), Inches(2.0), Inches(1.5)
                
                for col_idx, h_text in enumerate(["Sub Komponen", "Total Sisa (Rp)", "Persentase Sisa"]):
                    cell = table_s.cell(0, col_idx); cell.text = h_text
                    cell.fill.solid(); cell.fill.fore_color.rgb = GOLD_ACCENT 
                    p = cell.text_frame.paragraphs[0]
                    p.font.bold, p.font.color.rgb, p.font.size = True, WHITE, Pt(12)
                
                for r_idx, (lbl, val) in enumerate(zip(sub_labels_s_table, sub_vals_s_table), start=1):
                    pct = (val / total_komp_s_table * 100) if total_komp_s_table > 0 else 0
                    row_color = WHITE if r_idx % 2 != 0 else LIGHT_BG
                    for c_idx, text_val in enumerate([lbl, f"Rp {val:,.0f}", f"{pct:.2f}%"]):
                        cell = table_s.cell(r_idx, c_idx); cell.text = text_val
                        cell.fill.solid(); cell.fill.fore_color.rgb = row_color
                        p = cell.text_frame.paragraphs[0]
                        p.font.color.rgb, p.font.size = GREY_TEXT, Pt(11)
                
    ppt_output = io.BytesIO(); prs.save(ppt_output); ppt_output.seek(0)
    return ppt_output.getvalue()

# ==========================================
# ANTARMUKA PENGGUNA (UI STREAMLIT)
# ==========================================
st.title("📊 LAPORAN REALISASI")
st.markdown("**Executive Summary**")
st.divider()

col1, col2 = st.columns(2)
with col1:
    st.info("File RAB")
    file_rab = st.file_uploader("Upload Excel RAB", type=['xlsx', 'xls'], key="rab")
with col2:
    st.info("File LRA")
    file_lra_list = st.file_uploader("Upload Excel LRA (Bisa pilih banyak file)", type=['xlsx', 'xls'], accept_multiple_files=True, key="lra")

if file_rab and file_lra_list:
    st.divider()
    if st.button("🚀 PROSES", type="primary", use_container_width=True):
        
        list_semua_bulan = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']
        
        try:
            with st.status("Sedang memproses dokumen dan menyusun ringkasan...", expanded=True) as status:
                st.write("Mengekstrak data dari seluruh LRA...")
                lra_files_tuple = tuple((f.name, f.getvalue()) for f in file_lra_list)
                list_bulan_tuple = tuple(list_semua_bulan)
                
                data_realisasi, monthly_totals = parse_lra_files_cached(lra_files_tuple, list_bulan_tuple)
                
                st.write("Menyelaraskan dengan RAB & Mengekstrak Data untuk Summary...")
                rab_bytes = file_rab.getvalue()
                output_excel_bytes, data_ditemukan, satker_summary, component_metrics, sub_component_realisasi, sub_component_sisa = process_rab_lra_cached(rab_bytes, data_realisasi, list_bulan_tuple)
                
                st.write("Menyiapkan dokumen presentasi PowerPoint (.pptx)...")
                ppt_output_bytes = create_powerpoint_report(satker_summary, component_metrics, monthly_totals, sub_component_realisasi, sub_component_sisa)
                
                status.update(label="Proses Selesai!", state="complete", expanded=False)

            st.success(f"🎉 SUKSES! Data diselaraskan dan Sheet diubah menjadi 'DATA'.")
            st.divider()

            # ==========================================
            # DASHBOARD EXECUTIVE SUMMARY METRICS
            # ==========================================
            st.subheader("📈 Executive Summary")
            total_pagu_all = satker_summary["pagu"]
            total_blokir_all = satker_summary["blokir"]
            total_realisasi_all = satker_summary["realisasi"]
            total_sisa_all = satker_summary["sisa"]
            pagu_efektif_all = total_pagu_all - total_blokir_all
            persen_nasional = (total_realisasi_all / pagu_efektif_all * 100) if pagu_efektif_all > 0 else 0

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("💰 Total Pagu", f"Rp {total_pagu_all:,.0f}")
            m2.metric("🛑 Diblokir", f"Rp {total_blokir_all:,.0f}")
            m3.metric("📉 Total Realisasi", f"Rp {total_realisasi_all:,.0f}")
            m4.metric("🟡 Sisa Efektif", f"Rp {total_sisa_all:,.0f}")
            m5.metric("📊 Penyerapan", f"{persen_nasional:.2f}%")

            if component_metrics:
                st.markdown("---")
                st.markdown("### 🏷️ Summary Berdasarkan Komponen")
                comp_cols = st.columns(2)
                for idx, (komp_name, metrics) in enumerate(component_metrics.items()):
                    with comp_cols[idx % 2]:
                        st.markdown(f"**{komp_name}**")
                        sub_c1, sub_c2 = st.columns(2)
                        pagu_efektif_komp = metrics['pagu'] - metrics['blokir']
                        sub_c1.metric("Pagu Efektif", f"Rp {pagu_efektif_komp:,.0f}"); sub_c1.metric("Realisasi", f"Rp {metrics['realisasi']:,.0f}")
                        sub_c2.metric("Sisa", f"Rp {metrics['sisa']:,.0f}"); sub_c2.metric("Persentase", f"{metrics['persen']:.2f}%")

            st.markdown("---")
            
            # ==========================================
            # GRAFIK BULANAN
            # ==========================================
            st.markdown("### 📊 REALISASI Bulanan")
            
            realisasi_per_bulan = [monthly_totals.get(b, 0) for b in list_semua_bulan]
            df_monthly_chart = pd.DataFrame({"Bulan": list_semua_bulan, "Realisasi": realisasi_per_bulan})
            df_monthly_chart["Realisasi_Juta"] = df_monthly_chart["Realisasi"] / 1_000_000
            
            def format_rupiah(val): return f"Rp {val:,.0f}".replace(",", ".") if val > 0 else "Rp 0"

            fig_3d_bar = px.bar(df_monthly_chart, x="Bulan", y="Realisasi_Juta", text=df_monthly_chart["Realisasi"].apply(format_rupiah), title="Realisasi Anggaran per Bulan", color="Realisasi_Juta", color_continuous_scale="Tealgrn")
            fig_3d_bar.update_traces(textposition='outside', textfont_size=11, marker_line_color='rgb(8,48,107)', marker_line_width=1.5, opacity=0.9)
            fig_3d_bar.update_layout(plot_bgcolor="rgba(245,247,250,0.8)", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="black", size=12), xaxis_title="Bulan", yaxis_title="Total Realisasi (Juta Rp)", height=480, uniformtext_minsize=8, uniformtext_mode='hide')
            st.plotly_chart(fig_3d_bar, use_container_width=True)

            # ==========================================
            # DOWNLOAD BUTTONS
            # ==========================================
            st.divider()
            st.subheader("📥 DOWNLOAD LAPORAN")
            
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button(
                    label="⬇️ Download Laporan Excel (.xlsx)",
                    data=output_excel_bytes,
                    file_name="LAPORAN_RAB_LRA_LENGKAP_JAN_DES.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with dl_col2:
                st.download_button(
                    label="⬇️ Download Presentasi PowerPoint (.pptx)",
                    data=ppt_output_bytes,
                    file_name="PRESENTASI_EKSEKUTIF_ANGGARAN.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"⚠️ Terjadi kesalahan pada saat pemrosesan: {e}")
