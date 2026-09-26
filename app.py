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

# Pengaman untuk Library PowerPoint
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# ==========================================
# KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(page_title="RAB vs LRA Executive Generator", layout="wide", page_icon="📊")

def normalize_text(text):
    t = str(text)
    t = re.sub(r'^[\w\.]+\s*-\s*', '', t)
    t = re.sub(r'^[a-zA-Z0-9]+\.\s*', '', t)
    t = t.split('|')[0]
    t = t.replace('-', '').strip().lower().replace(',', '')
    return t

def match_texts_smart(t1, t2):
    if len(t1) < 4 or len(t2) < 4: return False
    if t1 == t2 or t1 in t2 or t2 in t1: return True
    
    t1_nospace = t1.replace(' ', '')
    t2_nospace = t2.replace(' ', '')
    if t1_nospace in t2_nospace or t2_nospace in t1_nospace:
        return True
    
    w1 = set(t1.split())
    w2 = set(t2.split())
    shorter = w1 if len(w1) < len(w2) else w2
    longer = w2 if len(w1) < len(w2) else w1
    if len(shorter) == 0: return False
    
    overlap = len(shorter.intersection(longer))
    ratio = overlap / len(shorter)
    if ratio >= 0.70 and (len(longer) / len(shorter) <= 2.5):
        return True
    return False

def get_row_pagu(ws, row_idx):
    """Menghitung nilai pagu per baris detail secara akurat dari volume & biaya satuan"""
    try:
        f_val = ws.cell(row=row_idx, column=6).value
        h_val = ws.cell(row=row_idx, column=8).value
        r_val = ws.cell(row=row_idx, column=18).value
        
        def parse_val(v):
            if isinstance(v, (int, float)): return float(v)
            if isinstance(v, str) and v.startswith('='):
                parts = v.lstrip('=').split('*')
                nums = []
                for p in parts:
                    p = p.strip()
                    try:
                        nums.append(float(p))
                    except:
                        col_letter = ''.join([c for c in p if c.isalpha()])
                        row_num = int(''.join([c for c in p if c.isdigit()]))
                        col_idx = ord(col_letter.upper()) - 64
                        cell_v = ws.cell(row=row_num, column=col_idx).value
                        if isinstance(cell_v, (int, float)):
                            nums.append(float(cell_v))
                        else:
                            nums.append(1.0)
                res = 1.0
                for n in nums: res *= n
                return res
            return 0.0
            
        vol1 = parse_val(f_val) if f_val is not None else 1.0
        vol2 = parse_val(h_val) if h_val is not None else 1.0
        biaya = parse_val(r_val) if r_val is not None else 0.0
        
        h_raw = ws.cell(row=row_idx, column=7).value
        if h_raw is not None and isinstance(h_raw, str) and 'x' in h_raw.lower():
            return vol1 * vol2 * biaya
        else:
            return vol1 * biaya
    except:
        return 0.0

# ==========================================
# FUNGSI PEMROSESAN DATA LRA & RAB
# ==========================================
def parse_lra_files(file_lra_list, list_semua_bulan):
    data_realisasi = {}
    satker_summary = {"pagu": 0, "realisasi": 0, "sisa": 0, "outstanding": 0}
    monthly_totals = {b: 0.0 for b in list_semua_bulan}
    component_summary = {}
    component_metrics = {}
    sub_component_realisasi = {}
    
    for uploaded_lra in file_lra_list:
        uploaded_lra.seek(0)
        try:
            df_lra = pd.read_excel(uploaded_lra, skiprows=5)
        except:
            uploaded_lra.seek(0)
            df_lra = pd.read_excel(uploaded_lra)
            
        df_lra.columns = [str(c).strip() for c in df_lra.columns]
        col_upper_map = {str(c).upper(): c for c in df_lra.columns}
        
        has_monthly_cols = all(b in col_upper_map for b in list_semua_bulan)
        
        level_col = col_upper_map.get('LEVEL', None)
        if not level_col:
            for c in df_lra.columns:
                if 'level' in str(c).lower():
                    level_col = c
                    break

        current_komp = ""
        cur_komp, cur_sub, cur_akun = "GLOBAL", "GLOBAL", "GLOBAL"
        
        for _, row in df_lra.iterrows():
            lvl = str(row.get(level_col, '')).strip() if level_col else ''
            
            uraian_val = ""
            for c in df_lra.columns:
                c_up = str(c).upper()
                if 'URAIAN' in c_up or 'KODE' in c_up or 'DESKRIPSI' in c_up:
                    val_c = str(row.get(c, '')).strip()
                    if val_c and val_c != 'nan':
                        uraian_val = val_c
                        break
            if not uraian_val:
                for c in df_lra.columns[:4]:
                    val_c = str(row.get(c, '')).strip()
                    if val_c and val_c != 'nan':
                        uraian_val = val_c
                        break
            
            def safe_float(val):
                try:
                    if pd.isna(val): return 0.0
                    return float(val)
                except:
                    return 0.0

            pagu = safe_float(row.get(col_upper_map.get('PAGU', ''), 0))
            realisasi_sub = safe_float(row.get(col_upper_map.get('TOTAL REALISASI', col_upper_map.get('REALISASI', '')), 0))
            sisa_sub = safe_float(row.get(col_upper_map.get('SISA', ''), 0))
            
            if 'komponen' in lvl.lower() or (uraian_val and re.search(r'^\d{3}\s*-', uraian_val)):
                match = re.search(r'(\d{3})\s*-', uraian_val)
                if match: cur_komp = match.group(1)
                current_komp = uraian_val
                component_summary[current_komp] = pagu
                if realisasi_sub > component_metrics.get(current_komp, {}).get('realisasi', 0):
                    component_metrics[current_komp] = {
                        "pagu": pagu,
                        "realisasi": realisasi_sub,
                        "sisa": sisa_sub,
                        "persen": (realisasi_sub / pagu * 100) if pagu > 0 else 0
                    }
                cur_sub, cur_akun = "GLOBAL", "GLOBAL"
            elif 'sub komponen' in lvl.lower() or (uraian_val and re.match(r'^[A-Z]\.?\s*-', uraian_val)):
                match = re.search(r'([A-Z])\s*-', uraian_val)
                if match: cur_sub = match.group(1)
                if current_komp:
                    if current_komp not in sub_component_realisasi:
                        sub_component_realisasi[current_komp] = {}
                    if realisasi_sub > sub_component_realisasi[current_komp].get(uraian_val, 0):
                        if realisasi_sub > 0:
                            sub_component_realisasi[current_komp][uraian_val] = realisasi_sub
                cur_akun = "GLOBAL"
            elif 'akun' in lvl.lower() or (uraian_val and re.search(r'^\d{6}\s*-', uraian_val)):
                match = re.search(r'(\d{6})\s*-', uraian_val)
                if match: cur_akun = match.group(1)
            
            if 'satker' in lvl.lower() or ('satker' in str(uraian_val).lower()):
                satker_summary["pagu"] = pagu
                satker_summary["realisasi"] = realisasi_sub
                satker_summary["sisa"] = sisa_sub

            val_gup = safe_float(row.get(col_upper_map.get('GUP', ''), 0))
            val_spm = safe_float(row.get(col_upper_map.get('SPM', ''), 0))
            val_verifikasi = safe_float(row.get(col_upper_map.get('VERIFIKASI', ''), 0))
            outstanding_val = val_gup + val_spm + val_verifikasi

            kamar_unik = f"{cur_komp}_{cur_sub}_{cur_akun}"
            if kamar_unik not in data_realisasi: 
                data_realisasi[kamar_unik] = {}
                
            if uraian_val and uraian_val not in ['nan', '']:
                norm_lra = normalize_text(uraian_val)
                
                if norm_lra not in data_realisasi[kamar_unik]:
                    data_realisasi[kamar_unik][norm_lra] = {b: 0.0 for b in list_semua_bulan}
                    data_realisasi[kamar_unik][norm_lra]['OUTSTANDING'] = 0.0
                
                if has_monthly_cols:
                    for b in list_semua_bulan:
                        col_name = col_upper_map.get(b)
                        if col_name:
                            m_val = safe_float(row.get(col_name, 0))
                            data_realisasi[kamar_unik][norm_lra][b] += m_val
                
                data_realisasi[kamar_unik][norm_lra]['OUTSTANDING'] += outstanding_val
                satker_summary["outstanding"] += outstanding_val
                
    return data_realisasi, satker_summary, monthly_totals, component_summary, component_metrics, sub_component_realisasi

def process_rab_lra(file_rab, data_realisasi, list_semua_bulan):
    wb = load_workbook(file_rab)
    ws = wb.active 
    
    baris_header = 13 
    baris_mulai_data = 14
    
    kolom_baru = [
        "TOTAL Realisasi", "SISA", "OUT STANDING", 
        "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
        "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER", "KETERANGAN"
    ]
    
    header_ref = ws.cell(row=baris_header, column=19)
    start_col = 20
    
    warna_hijau = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    warna_kuning = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    warna_orange = PatternFill(start_color="FCD5B4", end_color="FCD5B4", fill_type="solid") 
    
    for i, nama_kolom in enumerate(kolom_baru):
        cell = ws.cell(row=baris_header, column=start_col + i)
        cell.value = nama_kolom
        if header_ref.has_style:
            cell.font = copy(header_ref.font)
            cell.border = copy(header_ref.border)
            cell.alignment = copy(header_ref.alignment)
        
        if nama_kolom == "TOTAL Realisasi": cell.fill = warna_hijau
        elif nama_kolom == "SISA": cell.fill = warna_kuning
        elif nama_kolom == "OUT STANDING": cell.fill = warna_orange
        else:
            if header_ref.has_style: cell.fill = copy(header_ref.fill)

    ws.column_dimensions[get_column_letter(20)].width = 20.5 
    ws.column_dimensions[get_column_letter(21)].width = 18.2 
    ws.column_dimensions[get_column_letter(22)].width = 18.2 
    for c_idx in range(23, 35): ws.column_dimensions[get_column_letter(c_idx)].width = 15.0 
    ws.column_dimensions[get_column_letter(35)].width = 35.0 

    max_row = ws.max_row
    data_ditemukan = 0
    rab_komp, rab_sub, rab_akun = "GLOBAL", "GLOBAL", "GLOBAL"
    baris_terpakai = set()
    summary_preview = []

    for row_idx in range(baris_mulai_data, max_row + 1):
        kode_col = str(ws.cell(row=row_idx, column=2).value).strip()
        if kode_col.isdigit() and len(kode_col) == 3: 
            rab_komp = kode_col
            rab_sub, rab_akun = "GLOBAL", "GLOBAL"
        elif kode_col.isdigit() and len(kode_col) == 6:
            rab_akun = kode_col

        sub_col = str(ws.cell(row=row_idx, column=4).value).strip()
        if re.match(r'^[A-Z]\.?\s*$', sub_col):
            rab_sub = sub_col.replace('.', '').strip()
            rab_akun = "GLOBAL" 
            
        kamar_rab_saat_ini = f"{rab_komp}_{rab_sub}_{rab_akun}"
        
        bagian_teks = []
        for col_idx in range(3, 7):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val and isinstance(val, str) and str(val).strip() not in ['-', '']:
                bagian_teks.append(str(val).strip())
        
        if bagian_teks:
            uraian_rab = " ".join(bagian_teks)
            norm_rab = normalize_text(uraian_rab)
            
            matched_key = None
            kamar_opsi = [kamar_rab_saat_ini, f"{rab_komp}_{rab_sub}_GLOBAL", f"{rab_komp}_GLOBAL_GLOBAL"]
            
            for kamar in kamar_opsi:
                if len(norm_rab) > 2 and kamar in data_realisasi:
                    for key, dict_bulanan in data_realisasi[kamar].items():
                        if match_texts_smart(key, norm_rab):
                            matched_key = (kamar, key)
                            break
                if matched_key: break
            
            if matched_key:
                kamar_ketemu, key_ketemu = matched_key
                dict_bulanan = data_realisasi[kamar_ketemu].pop(key_ketemu) 
                nilai_outstanding = dict_bulanan.pop('OUTSTANDING', 0)
                
                target_row = row_idx
                for r_cek in range(row_idx, min(row_idx + 6, max_row + 1)):
                    val_s = ws.cell(row=r_cek, column=19).value
                    if val_s is not None and r_cek not in baris_terpakai:
                        target_row = r_cek
                        break
                
                baris_terpakai.add(target_row)
                
                cell_realisasi = ws.cell(row=target_row, column=20)   
                cell_sisa = ws.cell(row=target_row, column=21)        
                cell_outstanding = ws.cell(row=target_row, column=22) 
                
                cell_outstanding.value = nilai_outstanding
                
                row_bulanan_val = {}
                for idx_b, b_name in enumerate(list_semua_bulan):
                    col_target_bulan = 23 + idx_b 
                    cell_bulan = ws.cell(row=target_row, column=col_target_bulan)
                    val_b = dict_bulanan[b_name]
                    cell_bulan.value = val_b
                    row_bulanan_val[b_name] = val_b
                
                cell_realisasi.value = f"=SUM(W{target_row}:AH{target_row})+V{target_row}"
                cell_sisa.value = f"=S{target_row}-T{target_row}"

                pagu_val = get_row_pagu(ws, target_row)
                total_realisasi_row = sum(row_bulanan_val.values()) + nilai_outstanding

                summary_preview.append({
                    "Komponen": rab_komp,
                    "Uraian": uraian_rab,
                    "Pagu": pagu_val,
                    "Realisasi": total_realisasi_row,
                    "Outstanding": nilai_outstanding,
                    **row_bulanan_val
                })

                data_ref = ws.cell(row=target_row, column=19)
                for c_idx in range(20, 36):
                    c = ws.cell(row=target_row, column=c_idx)
                    if data_ref.has_style:
                        c.font = copy(data_ref.font)
                        c.border = copy(data_ref.border)
                        c.alignment = copy(data_ref.alignment)
                        c.number_format = copy(data_ref.number_format)
                
                data_ditemukan += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, data_ditemukan, pd.DataFrame(summary_preview)

# ==========================================
# FUNGSI PEMBUATAN FILE POWERPOINT (PPTX)
# ==========================================
def create_powerpoint_report(satker_summary, component_metrics, monthly_totals, sub_component_realisasi):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # 1. Slide Judul (Cover)
    slide1 = prs.slides.add_slide(blank_layout)
    title_box = slide1.shapes.add_textbox(Inches(1), Inches(2.2), Inches(11.333), Inches(3))
    tf1 = title_box.text_frame
    tf1.word_wrap = True
    
    p1 = tf1.paragraphs[0]
    p1.text = "LAPORAN EKSEKUTIF KONSOLIDASI ANGGARAN"
    p1.font.size = Pt(32)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(24, 43, 73)
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf1.add_paragraph()
    p2.text = "RAB vs LRA (Kementerian Koperasi dan UKM)"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(100, 110, 120)
    p2.alignment = PP_ALIGN.CENTER
    
    # 2. Slide Ringkasan Eksekutif Nasional & Komponen
    slide2 = prs.slides.add_slide(blank_layout)
    t_box2 = slide2.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.8))
    tf2 = t_box2.text_frame
    p_h2 = tf2.paragraphs[0]
    p_h2.text = "Ringkasan Eksekutif & Kinerja Anggaran"
    p_h2.font.size = Pt(24)
    p_h2.font.bold = True
    p_h2.font.color.rgb = RGBColor(24, 43, 73)
    
    total_pagu = satker_summary["pagu"]
    total_realisasi = satker_summary["realisasi"]
    total_sisa = satker_summary["sisa"]
    persen_nasional = (total_realisasi / total_pagu * 100) if total_pagu > 0 else 0
    
    metrics_text = (
        f"• Total Pagu Anggaran : Rp {total_pagu:,.0f}\n"
        f"• Total Realisasi     : Rp {total_realisasi:,.0f}\n"
        f"• Sisa Anggaran       : Rp {total_sisa:,.0f}\n"
        f"• Tingkat Penyerapan  : {persen_nasional:.2f}%\n\n"
    )
    if component_metrics:
        metrics_text += "Detail per Komponen Utama:\n"
        for komp_name, m in component_metrics.items():
            metrics_text += f" - {komp_name} | Pagu: Rp {m['pagu']:,.0f} | Realisasi: Rp {m['realisasi']:,.0f} ({m['persen']:.2f}%)\n"
            
    m_box = slide2.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.7), Inches(5))
    mtf = m_box.text_frame
    mtf.word_wrap = True
    mp = mtf.paragraphs[0]
    mp.text = metrics_text
    mp.font.size = Pt(16)
    mp.font.color.rgb = RGBColor(40, 40, 40)
    
    # 3. Slide Grafik Tren Bulanan
    slide3 = prs.slides.add_slide(blank_layout)
    t_box3 = slide3.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.8))
    tf3 = t_box3.text_frame
    p_h3 = tf3.paragraphs[0]
    p_h3.text = "Grafik Tren Penyerapan Anggaran Bulanan"
    p_h3.font.size = Pt(24)
    p_h3.font.bold = True
    p_h3.font.color.rgb = RGBColor(24, 43, 73)
    
    fig_m, ax_m = plt.subplots(figsize=(10, 4.5))
    months = list(monthly_totals.keys())
    values = list(monthly_totals.values())
    bars = ax_m.bar(months, [v / 1e9 for v in values], color='#2a9d8f')
    ax_m.set_ylabel('Realisasi (Miliar Rp)', fontsize=11, fontweight='bold')
    ax_m.set_title('Akumulasi Realisasi Anggaran per Bulan', fontsize=13, fontweight='bold', pad=15)
    plt.xticks(rotation=35, ha='right', fontsize=10)
    ax_m.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    
    img_buf = io.BytesIO()
    fig_m.savefig(img_buf, format='png', dpi=200, bbox_inches='tight')
    plt.close(fig_m)
    img_buf.seek(0)
    
    slide3.shapes.add_picture(img_buf, Inches(0.8), Inches(1.5), width=Inches(11.7))
    
    # 4. Slide Rincian Sub Komponen (Pie Chart & Tabel)
    if sub_component_realisasi:
        for komp_name, sub_dict in sub_component_realisasi.items():
            slide_sub = prs.slides.add_slide(blank_layout)
            t_box_sub = slide_sub.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.7), Inches(0.8))
            tf_sub = t_box_sub.text_frame
            p_hs = tf_sub.paragraphs[0]
            p_hs.text = f"Proporsi & Rincian Sub Komponen: {komp_name}"
            p_hs.font.size = Pt(22)
            p_hs.font.bold = True
            p_hs.font.color.rgb = RGBColor(24, 43, 73)
            
            sub_labels = list(sub_dict.keys())
            sub_vals = list(sub_dict.values())
            total_komp = sum(sub_vals)
            
            fig_p, ax_p = plt.subplots(figsize=(5.5, 4.5))
            ax_p.pie(
                sub_vals, 
                labels=[f"Sub {l.split('-')[0].strip()}" for l in sub_labels], 
                autopct='%1.1f%%', 
                startangle=90, 
                colors=plt.cm.Paired.colors
            )
            ax_p.axis('equal')
            plt.tight_layout()
            
            pie_buf = io.BytesIO()
            fig_p.savefig(pie_buf, format='png', dpi=200, bbox_inches='tight')
            plt.close(fig_p)
            pie_buf.seek(0)
            
            slide_sub.shapes.add_picture(pie_buf, Inches(0.8), Inches(1.5), width=Inches(5.0))
            
            rows = len(sub_dict) + 1
            cols = 3
            left = Inches(6.2)
            top = Inches(1.8)
            width = Inches(6.3)
            height = Inches(0.5 * rows)
            
            table_shape = slide_sub.shapes.add_table(rows, cols, left, top, width, height)
            table = table_shape.table
            table.columns[0].width = Inches(2.3)
            table.columns[1].width = Inches(2.3)
            table.columns[2].width = Inches(1.7)
            
            table.cell(0, 0).text = "Sub Komponen"
            table.cell(0, 1).text = "Total Realisasi (Rp)"
            table.cell(0, 2).text = "Persentase"
            
            for r_idx, (lbl, val) in enumerate(zip(sub_labels, sub_vals), start=1):
                pct = (val / total_komp * 100) if total_komp > 0 else 0
                table.cell(r_idx, 0).text = lbl
                table.cell(r_idx, 1).text = f"Rp {val:,.0f}"
                table.cell(r_idx, 2).text = f"{pct:.2f}%"
                
    ppt_output = io.BytesIO()
    prs.save(ppt_output)
    ppt_output.seek(0)
    return ppt_output

# ==========================================
# ANTARMUKA PENGGUNA (UI)
# ==========================================
st.title("📊 EXECUTIVE RAB VS LRA GENERATOR")
st.markdown("**Konsolidasi Laporan Anggaran Multi-Bulan & Executive Summary (Kementerian Koperasi dan UKM)**")
st.divider()

col1, col2 = st.columns(2)
with col1:
    st.info("Langkah 1: Masukkan File RAB")
    file_rab = st.file_uploader("Upload Excel RAB", type=['xlsx', 'xls'], key="rab")
with col2:
    st.info("Langkah 2: Masukkan File LRA")
    file_lra_list = st.file_uploader("Upload Excel LRA (Bisa pilih banyak file)", type=['xlsx', 'xls'], accept_multiple_files=True, key="lra")

if file_rab and file_lra_list:
    st.divider()
    if st.button("🚀 Proses & Buat Laporan", type="primary", use_container_width=True):
        
        list_semua_bulan = [
            'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 
            'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
        ]
        
        try:
            with st.status("Sedang memproses dokumen dan menyusun ringkasan...", expanded=True) as status:
                st.write("Mengekstrak data dari seluruh LRA (Realisasi & Outstanding)...")
                data_realisasi, satker_summary, monthly_totals, component_summary, component_metrics, sub_component_realisasi = parse_lra_files(file_lra_list, list_semua_bulan)
                
                st.write("Menyelaraskan dan memodifikasi template RAB...")
                output_excel, data_ditemukan, df_preview = process_rab_lra(file_rab, data_realisasi, list_semua_bulan)
                
                if not df_preview.empty:
                    monthly_totals = {b: df_preview[b].sum() if b in df_preview.columns else 0.0 for b in list_semua_bulan}
                
                st.write("Menyiapkan dokumen presentasi PowerPoint (.pptx)...")
                ppt_output = create_powerpoint_report(satker_summary, component_metrics, monthly_totals, sub_component_realisasi)
                
                status.update(label="Proses Selesai!", state="complete", expanded=False)

            st.success(f"🎉 SUKSES! Berhasil menyelaraskan **{data_ditemukan} baris** data RAB dengan data LRA bulanan.")
            st.divider()

            # ==========================================
            # DASHBOARD EXECUTIVE SUMMARY METRICS
            # ==========================================
            st.subheader("📈 Dashboard Ringkasan Eksekutif (Executive Summary)")
            
            total_pagu_all = satker_summary["pagu"]
            total_realisasi_incl_out = satker_summary["realisasi"]
            total_sisa_all = satker_summary["sisa"]
            persen_nasional = (total_realisasi_incl_out / total_pagu_all * 100) if total_pagu_all > 0 else 0

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("💰 Total Pagu Anggaran", f"Rp {total_pagu_all:,.0f}")
            m2.metric("📉 Total Realisasi", f"Rp {total_realisasi_incl_out:,.0f}")
            m3.metric("🟡 Sisa Anggaran", f"Rp {total_sisa_all:,.0f}")
            m4.metric("📊 Tingkat Penyerapan", f"{persen_nasional:.2f}%")

            if component_metrics:
                st.markdown("---")
                st.markdown("### 🏷️ Ringkasan Per Komponen")
                
                # PERBAIKAN: Render per komponen secara dinamis (2 komponen per baris) agar aman berapapun jumlah komponennya
                komp_metric_items = list(component_metrics.items())
                for i in range(0, len(komp_metric_items), 2):
                    comp_cols = st.columns(2)
                    for j in range(2):
                        if i + j < len(komp_metric_items):
                            komp_name, metrics = komp_metric_items[i + j]
                            with comp_cols[j]:
                                st.markdown(f"**{komp_name}**")
                                sub_c1, sub_c2 = st.columns(2)
                                sub_c1.metric("Pagu", f"Rp {metrics['pagu']:,.0f}")
                                sub_c1.metric("Realisasi", f"Rp {metrics['realisasi']:,.0f}")
                                sub_c2.metric("Sisa", f"Rp {metrics['sisa']:,.0f}")
                                sub_c2.metric("Penyerapan", f"{metrics['persen']:.2f}%")

            st.markdown("---")
            st.markdown("### 📊 Tren Penyerapan Bulanan")
            
            df_monthly_chart = pd.DataFrame({
                "Bulan": list(monthly_totals.keys()),
                "Realisasi": list(monthly_totals.values())
            })
            
            fig_3d_bar = px.bar(
                df_monthly_chart,
                x="Bulan",
                y="Realisasi",
                text=df_monthly_chart["Realisasi"].apply(lambda x: f"Rp {x:,.0f}" if x > 0 else "Rp 0"),
                title="Realisasi Anggaran per Bulan",
                color="Realisasi",
                color_continuous_scale="Tealgrn"
            )
            fig_3d_bar.update_traces(
                textposition='outside', 
                textfont_size=11,
                marker_line_color='rgb(8,48,107)',
                marker_line_width=1.5,
                opacity=0.9
            )
            fig_3d_bar.update_layout(
                plot_bgcolor="rgba(245,247,250,0.8)",
                paper_bgcolor="rgba(0,0,0,0)",
                font=dict(color="black", size=12),
                xaxis_title="Bulan",
                yaxis_title="Total Realisasi (Rp)",
                height=480,
                uniformtext_minsize=8, 
                uniformtext_mode='hide'
            )
            st.plotly_chart(fig_3d_bar, use_container_width=True)

            st.markdown("---")
            st.markdown("### 📋 Rincian Proporsi Total Realisasi Sub Komponen per Komponen")
            
            chart_colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2']
            
            if sub_component_realisasi:
                komp_keys = list(sub_component_realisasi.keys())
                
                for i in range(0, len(komp_keys), 2):
                    row_cols = st.columns(2)
                    for j in range(2):
                        if i + j < len(komp_keys):
                            komp_name = komp_keys[i + j]
                            with row_cols[j]:
                                st.markdown(f"**{komp_name}**")
                                sub_dict = sub_component_realisasi[komp_name]
                                sub_labels = list(sub_dict.keys())
                                sub_values = list(sub_dict.values())
                                total_komp_val = sum(sub_values)
                                
                                if total_komp_val > 0:
                                    fig_donut = go.Figure(data=[go.Pie(
                                        labels=[f"Sub {l.split('-')[0].strip()}" for l in sub_labels],
                                        values=sub_values,
                                        hole=0.45,
                                        textinfo='percent+label',
                                        hoverinfo='none',
                                        textfont_size=11,
                                        marker=dict(colors=chart_colors[:len(sub_labels)], line=dict(color='#FFFFFF', width=2))
                                    )])
                                    fig_donut.update_layout(
                                        plot_bgcolor="rgba(0,0,0,0)",
                                        paper_bgcolor="rgba(0,0,0,0)",
                                        margin=dict(t=20, b=20, l=20, r=20),
                                        showlegend=False,
                                        height=280
                                    )
                                    st.plotly_chart(fig_donut, use_container_width=True, key=f"donut_{i}_{j}")
                                    
                                    head_c = st.columns([0.5, 4.5, 2.5, 2])
                                    with head_c[0]: st.markdown("")
                                    with head_c[1]: st.markdown("**Sub Komponen**")
                                    with head_c[2]: st.markdown("**Total Realisasi (Rp)**")
                                    with head_c[3]: st.markdown("**Persentase**")
                                    st.markdown("<hr style='margin: 4px 0px 8px 0px;'>", unsafe_allow_html=True)

                                    for k, (label, val) in enumerate(zip(sub_labels, sub_values)):
                                        color_hex = chart_colors[k % len(chart_colors)]
                                        pct = (val / total_komp_val) * 100
                                        
                                        row_c = st.columns([0.5, 4.5, 2.5, 2])
                                        with row_c[0]:
                                            st.markdown(f"<div style='width:14px; height:14px; background-color:{color_hex}; border-radius:3px; margin-top:5px;'></div>", unsafe_allow_html=True)
                                        with row_c[1]:
                                            st.markdown(f"<span style='font-size:12px; color:#212529;'>{label}</span>", unsafe_allow_html=True)
                                        with row_c[2]:
                                            st.markdown(f"<span style='font-size:12px; font-weight:500; color:#212529;'>Rp {val:,.0f}</span>", unsafe_allow_html=True)
                                        with row_c[3]:
                                            st.markdown(f"<span style='font-size:12px; font-weight:600; color:#495057;'>{pct:.2f}%</span>", unsafe_allow_html=True)
                                else:
                                    st.info("Belum ada realisasi anggaran pada komponen ini.")
            else:
                st.info("Data realisasi sub komponen belum tersedia.")

            st.divider()
            st.subheader("📥 Download Berkas Laporan & Presentasi")
            
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button(
                    label="⬇️ Download Laporan Excel (.xlsx)",
                    data=output_excel,
                    file_name="LAPORAN_RAB_LRA_LENGKAP_JAN_DES.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with dl_col2:
                st.download_button(
                    label="⬇️ Download Presentasi PowerPoint (.pptx)",
                    data=ppt_output,
                    file_name="PRESENTASI_EKSEKUTIF_ANGGARAN.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"⚠️ Terjadi kesalahan pada saat pemrosesan: {e}")
