import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import io
import re
from copy import copy

st.set_page_config(page_title="RAB vs LRA Generator", layout="wide")
st.title("📊LAPORAN")
st.markdown("RAB VS LRA")

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

col1, col2 = st.columns(2)
with col1:
    file_rab = st.file_uploader("1. Upload Excel RAB", type=['xlsx', 'xls'], key="rab")
with col2:
    file_lra = st.file_uploader("2. Upload Excel LRA", type=['xlsx', 'xls'], key="lra")

if file_rab and file_lra:
    if st.button("🚀 Proses & Buat Laporan", type="primary"):
        with st.spinner("Memproses sistem booking baris..."):
            try:
                # ==========================================
                # 1. BACA LRA
                # ==========================================
                df_lra = pd.read_excel(file_lra, skiprows=5)
                data_realisasi = {}
                
                cur_komp = "GLOBAL"
                cur_sub = "GLOBAL"
                cur_akun = "GLOBAL"
                
                for index, row in df_lra.iterrows():
                    lvl = str(row.get('Level')).strip()
                    uraian = str(row.get('Kode / Uraian', '')).strip()
                    realisasi = row.get('Total Realisasi', 0)
                    if pd.isna(realisasi): realisasi = 0
                        
                    if lvl == 'Komponen':
                        match = re.search(r'(\d{3})\s*-', uraian)
                        if match: cur_komp = match.group(1)
                        cur_sub = "GLOBAL"
                        cur_akun = "GLOBAL"
                    elif lvl == 'Sub Komponen':
                        match = re.search(r'([A-Z])\s*-', uraian)
                        if match: cur_sub = match.group(1)
                        cur_akun = "GLOBAL"
                    elif lvl == 'Akun':
                        match = re.search(r'(\d{6})\s*-', uraian)
                        if match: cur_akun = match.group(1)
                    
                    kamar_unik = f"{cur_komp}_{cur_sub}_{cur_akun}"
                    if kamar_unik not in data_realisasi: data_realisasi[kamar_unik] = {}
                        
                    if pd.notna(uraian) and uraian != 'nan' and uraian != '':
                        norm_lra = normalize_text(uraian)
                        if norm_lra not in data_realisasi[kamar_unik]:
                            data_realisasi[kamar_unik][norm_lra] = []
                        data_realisasi[kamar_unik][norm_lra].append(realisasi)

                # ==========================================
                # 2. BACA & MODIFIKASI RAB 
                # ==========================================
                wb = load_workbook(file_rab)
                ws = wb.active 
                
                baris_header = 13 
                baris_mulai_data = 14
                
                # URAIAN BULAN DIUBAH MENJADI NAMA LENGKAP
                kolom_baru = [
                    "TOTAL Realisasi", "SISA", 
                    "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
                    "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER", "KETERANGAN"
                ]
                
                header_ref = ws.cell(row=baris_header, column=19)
                start_col = 20
                
                warna_hijau = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
                warna_kuning = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                
                for i, nama_kolom in enumerate(kolom_baru):
                    cell = ws.cell(row=baris_header, column=start_col + i)
                    cell.value = nama_kolom
                    if header_ref.has_style:
                        cell.font = copy(header_ref.font)
                        cell.border = copy(header_ref.border)
                        cell.alignment = copy(header_ref.alignment)
                    
                    if nama_kolom == "TOTAL Realisasi": cell.fill = warna_hijau
                    elif nama_kolom == "SISA": cell.fill = warna_kuning
                    else:
                        if header_ref.has_style: cell.fill = copy(header_ref.fill)

                ws.column_dimensions[get_column_letter(20)].width = 20.5
                ws.column_dimensions[get_column_letter(21)].width = 18.2
                # Lebar kolom bulan (22 sampai 33) disesuaikan agar nama bulan yang panjang muat dengan baik
                for c_idx in range(22, 34): ws.column_dimensions[get_column_letter(c_idx)].width = 15.0
                ws.column_dimensions[get_column_letter(34)].width = 35.0

                max_row = ws.max_row
                data_ditemukan = 0
                
                rab_komp = "GLOBAL"
                rab_sub = "GLOBAL"
                rab_akun = "GLOBAL"
                
                # LIST BARIS YANG SUDAH TERPAKAI (Di-booking)
                baris_terpakai = set()

                for row_idx in range(baris_mulai_data, max_row + 1):
                    kode_col = str(ws.cell(row=row_idx, column=2).value).strip()
                    if kode_col.isdigit() and len(kode_col) == 3: 
                        rab_komp = kode_col
                        rab_sub = "GLOBAL"
                        rab_akun = "GLOBAL"
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
                                for key, q in data_realisasi[kamar].items():
                                    if len(q) > 0 and match_texts_smart(key, norm_rab):
                                        matched_key = (kamar, key)
                                        break
                            if matched_key: break
                        
                        if matched_key:
                            kamar_ketemu, key_ketemu = matched_key
                            nilai_realisasi = data_realisasi[kamar_ketemu][key_ketemu].pop(0)
                            
                            # ============================================================
                            # SISTEM BOOKING BARIS: Cari baris berangka S yang belum dipakai
                            # ============================================================
                            target_row = row_idx
                            for r_cek in range(row_idx, min(row_idx + 6, max_row + 1)):
                                val_s = ws.cell(row=r_cek, column=19).value
                                # Jika baris S ada isinya DAN baris ini belum diklaim baris lain
                                if val_s is not None and r_cek not in baris_terpakai:
                                    target_row = r_cek
                                    break
                            
                            baris_terpakai.add(target_row) # Booking baris ini!
                            # ============================================================
                            
                            cell_realisasi = ws.cell(row=target_row, column=20)
                            cell_sisa = ws.cell(row=target_row, column=21)
                            
                            cell_realisasi.value = nilai_realisasi
                            cell_sisa.value = f"=S{target_row}-T{target_row}"
                            
                            data_ref = ws.cell(row=target_row, column=19)
                            for c in [cell_realisasi, cell_sisa]:
                                if data_ref.has_style:
                                    c.font = copy(data_ref.font)
                                    c.border = copy(data_ref.border)
                                    c.alignment = copy(data_ref.alignment)
                                    c.number_format = copy(data_ref.number_format)
                            
                            data_ditemukan += 1

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"🎉 SUKSES! Berhasil menyelaraskan {data_ditemukan} baris menggunakan Sistem Booking.")
                
                st.download_button(
                    label="⬇️ Download Laporan Akhir (.xlsx)",
                    data=output,
                    file_name="LAPORAN_RAB_LRA_SEJAJAR_SEMPURNA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"⚠️ Terjadi kesalahan: {e}")
