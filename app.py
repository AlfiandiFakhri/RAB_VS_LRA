import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import io
import re
from copy import copy

# Pengaman untuk Library PowerPoint
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
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

# ==========================================
# FUNGSI PEMROSESAN DATA LRA & RAB
# ==========================================
def parse_lra_files(file_lra_list, list_semua_bulan):
    data_realisasi = {}
    
    for uploaded_lra in file_lra_list:
        df_raw = pd.read_excel(uploaded_lra, header=None, nrows=15)
        bulan_file = None
        is_outstanding_file = False
        
        # 1. Cek Bulan dari NAMA FILE
        fname_upper = uploaded_lra.name.upper()
        for b in list_semua_bulan:
            if b in fname_upper:
                bulan_file = b
                break

        # 2. Cek Bulan dari ISI FILE (Baris 1-15)
        if not bulan_file:
            for r in range(len(df_raw)):
                row_text_raw = " ".join(str(val) for val in df_raw.iloc[r].values if pd.notna(val)).upper()
                for b in list_semua_bulan:
                    if b in row_text_raw:
                        bulan_file = b
                        break
                if bulan_file: break

        # 3. Deteksi File Outstanding (Berdasarkan Nama File atau Isi File)
        fname_clean = fname_upper.replace(" ", "")
        keyword_out = ["SEMUA", "LEVEL", "ALL", "PERIODE", "REKAP", "GUP"]
        if any(kw in fname_clean for kw in keyword_out) and not bulan_file:
            is_outstanding_file = True

        for r in range(len(df_raw)):
            row_text_raw = " ".join(str(val) for val in df_raw.iloc[r].values if pd.notna(val)).upper()
            if "SEMUA" in row_text_raw or "LEVEL" in row_text_raw or "ALL" in row_text_raw or "PERIODE" in row_text_raw:
                if not bulan_file:
                    is_outstanding_file = True
                    break

        # Baca Data LRA
        df_lra = pd.read_excel(uploaded_lra, skiprows=5)
        
        # 4. Failsafe Utama: Cek Kolom GUP / SPM / Verifikasi
        kolom_lra = [str(col).upper() for col in df_lra.columns]
        ada_gup_spm = any('GUP' in col or 'SPM' in col or 'VERIFIKASI' in col for col in kolom_lra)
        
        if ada_gup_spm and not bulan_file:
            is_outstanding_file = True

        # Jika bukan file bulanan DAN bukan file outstanding, baru di-skip
        if not bulan_file and not is_outstanding_file:
            continue
            
        cur_komp, cur_sub, cur_akun = "GLOBAL", "GLOBAL", "GLOBAL"
        
        for index, row in df_lra.iterrows():
            lvl = str(row.get('Level')).strip()
            uraian = str(row.get('Kode / Uraian', '')).strip()
            
            realisasi = row.get('Total Realisasi', 0)
            try: realisasi = float(realisasi)
            except: realisasi = 0
            if pd.isna(realisasi): realisasi = 0
                
            outstanding_val = 0
            if is_outstanding_file:
                val_gup = row.get('GUP', 0)
                val_spm = row.get('SPM', 0)
                val_verifikasi = row.get('Verifikasi', 0)
                
                try: val_gup = float(val_gup)
                except: val_gup = 0
                try: val_spm = float(val_spm)
                except: val_spm = 0
                try: val_verifikasi = float(val_verifikasi)
                except: val_verifikasi = 0
                
                outstanding_val = val_gup + val_spm + val_verifikasi
                
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
            if kamar_unik not in data_realisasi: 
                data_realisasi[kamar_unik] = {}
                
            if pd.notna(uraian) and uraian not in ['nan', '']:
                norm_lra = normalize_text(uraian)
                
                if norm_lra not in data_realisasi[kamar_unik]:
                    data_realisasi[kamar_unik][norm_lra] = {b: 0 for b in list_semua_bulan}
                    data_realisasi[kamar_unik][norm_lra]['OUTSTANDING'] = 0
                
                if bulan_file:
                    data_realisasi[kamar_unik][norm_lra][bulan_file] += realisasi
                
                if is_outstanding_file:
                    data_realisasi[kamar_unik][norm_lra]['OUTSTANDING'] += outstanding_val
                
    return data_realisasi

def process_rab_lra(file_rab, data_realisasi, list_semua_bulan):
    wb = load_workbook(file_rab)
    ws = wb.active 
    
    baris_header = 13 
    baris_mulai_data = 14
    
    kolom_baru = [
        "TOTAL Realisasi", "SISA", "OUT STANDING", 
        "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
        "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER", 
        "PENYERAPAN (%)", "KETERANGAN"
    ]
    
    header_ref = ws.cell(row=baris_header, column=19)
    start_col = 20
    
    warna_hijau = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    warna_kuning = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    warna_orange = PatternFill(start_color="FCD5B4", end_color="FCD5B4", fill_type="solid") 
    warna_biru_muda = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
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
        elif nama_kolom == "PENYERAPAN (%)": cell.fill = warna_biru_muda
        else:
            if header_ref.has_style: cell.fill = copy(header_ref.fill)

    ws.column_dimensions[get_column_letter(20)].width = 20.5 
    ws.column_dimensions[get_column_letter(21)].width = 18.2 
    ws.column_dimensions[get_column_letter(22)].width = 18.2 
    for c_idx in range(23, 35): ws.column_dimensions[get_column_letter(c_idx)].width = 15.0 
    ws.column_dimensions[get_column_letter(35)].width = 16.0 
    ws.column_dimensions[get_column_letter(36)].width = 35.0 

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
                cell_persen = ws.cell(row=target_row, column=35)      
                
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
                cell_persen.value = f"=T{target_row}/S{target_row}"
                cell_persen.number_format = '0.0%'

                pagu_val = ws.cell(row=target_row, column=19).value or 0
                try: pagu_val = float(pagu_val)
                except: pagu_val = 0
                
                summary_preview.append({
                    "Komponen": rab_komp,
                    "Uraian": uraian_rab,
                    "Pagu": pagu_val,
                    "Outstanding": nilai_outstanding,
                    **row_bulanan_val
                })

                data_ref = ws.cell(row=target_row, column=19)
                for c_idx in range(20, 37):
                    c = ws.cell(row=target_row, column=c_idx)
                    if data_ref.has_style:
                        c.font = copy(data_ref.font)
                        c.border = copy(data_ref.border)
                        c.alignment = copy(data_ref.alignment)
                        if c_idx != 35: 
                            c.number_format = copy(data_ref.number_format)
                
                data_ditemukan += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, data_ditemukan, pd.DataFrame(summary_preview)

def generate_pptx_presentation(metrics, monthly_totals):
    if not HAS_PPTX:
        return None
    prs = Presentation()
    
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "LAPORAN KONSOLIDASI ANGGARAN"
    slide.placeholders[1].text = "Executive Summary RAB vs LRA Tahun 2026\nAsdep PIMEN - Kementerian Koperasi dan UKM"

    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Ringkasan Kinerja Anggaran (Executive Metrics)"
    
    tf = slide.placeholders[1].text_frame
    tf.text = f"• Total Pagu Anggaran : Rp {metrics['pagu']:,.0f}"
    tf.add_paragraph().text = f"• Total Realisasi (incl. Outstanding) : Rp {metrics['realisasi']:,.0f}"
    tf.add_paragraph().text = f"• Total Sisa Anggaran : Rp {metrics['sisa']:,.0f}"
    tf.add_paragraph().text = f"• Rata-rata Tingkat Penyerapan : {metrics['persentase']:.2f}%"

    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Catatan Strategis Pimpinan"
    tf2 = slide.placeholders[1].text_frame
    tf2.text = "• Konsolidasi data bersumber dari LRA bulanan (Januari - Desember) dan rekapitulasi All Periode."
    tf2.add_paragraph().text = "• Nilai Outstanding (GUP + SPM + Verifikasi) telah diintegrasikan langsung ke dalam perhitungan penyerapan."
    tf2.add_paragraph().text = "• Laporan siap digunakan sebagai bahan monitoring dan evaluasi penyerapan anggaran Triwulanan."

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ==========================================
# ANTARMUKA PENGGUNA (UI) STREAMLIT
# ==========================================
st.title("📊 EXECUTIVE RAB VS LRA GENERATOR")
st.markdown("**Sistem Konsolidasi & Analisis Laporan Anggaran Multi-Bulan (Kementerian Koperasi dan UKM)**")
st.divider()

col1, col2 = st.columns(2)
with col1:
    st.info("Langkah 1: Upload File RAB Utama")
    file_rab = st.file_uploader("Pilih file Excel RAB", type=['xlsx', 'xls'], key="rab")
with col2:
    st.info("Langkah 2: Upload Seluruh File LRA (Jan - Des + All Periode)")
    file_lra_list = st.file_uploader("Pilih banyak file LRA sekaligus", type=['xlsx', 'xls'], accept_multiple_files=True, key="lra")

if file_rab and file_l_list_checked := file_lra_list:
    st.divider()
    if st.button("🚀 Proses Konsolidasi & Buat Dashboard", type="primary", use_container_width=True):
        
        list_semua_bulan = [
            'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 
            'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
        ]
        
        try:
            with st.status("Sedang memproses dokumen dan menyelaraskan data...", expanded=True) as status:
                st.write("Mengekstrak data realisasi bulanan dan outstanding...")
                data_realisasi = parse_lra_files(file_lra_list, list_semua_bulan)
                
                st.write("Memodifikasi dan memasukkan rumus ke template RAB...")
                output_excel, data_ditemukan, df_preview = process_rab_lra(file_rab, data_realisasi, list_semua_bulan)
                
                status.update(label="Konsolidasi Selesai!", state="complete", expanded=False)

            st.success(f"🎉 Berhasil menyelaraskan **{data_ditemukan} baris** data anggaran!")
            st.divider()

            st.subheader("📈 Dashboard Ringkasan Eksekutif")
            
            total_pagu_all = df_preview['Pagu'].sum() if not df_preview.empty else 0
            total_out_all = df_preview['Outstanding'].sum() if not df_preview.empty else 0
            
            monthly_sums = {}
            for b in list_semua_bulan:
                if b in df_preview.columns:
                    monthly_sums[b] = df_preview[b].sum()
                else:
                    monthly_sums[b] = 0
            
            total_realisasi_bulanan = sum(monthly_sums.values())
            total_realisasi_incl_out = total_realisasi_bulanan + total_out_all
            total_sisa_all = total_pagu_all - total_realisasi_incl_out
            persen_nasional = (total_realisasi_incl_out / total_pagu_all * 100) if total_pagu_all > 0 else 0

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("💰 Total Pagu Anggaran", f"Rp {total_pagu_all:,.0f}")
            m2.metric("📉 Total Realisasi (+ Outstd)", f"Rp {total_realisasi_incl_out:,.0f}")
            m3.metric("🟡 Sisa Anggaran", f"Rp {total_sisa_all:,.0f}")
            m4.metric("📊 Tingkat Penyerapan", f"{persen_nasional:.2f}%")

            st.markdown("---")
            st.markdown("### 📊 Grafik Tren Penyerapan Anggaran Bulanan")
            s_bulan = pd.Series(monthly_sums)
            st.bar_chart(s_bulan)

            with st.expander("🔍 Pratinjau & Filter Data Konsolidasi (Preview Tabel)", expanded=False):
                komponen_list = df_preview['Komponen'].unique() if not df_preview.empty else []
                selected_komp = st.multiselect("Filter Berdasarkan Komponen:", options=komponen_list, default=komponen_list)
                
                if not df_preview.empty:
                    df_filtered = df_preview[df_preview['Komponen'].isin(selected_komp)]
                    st.dataframe(df_filtered, use_container_width=True)

            st.divider()

            st.subheader("📥 Download Berkas Laporan Akhir")
            
            metrics_dict = {
                "pagu": total_pagu_all,
                "realisasi": total_realisasi_incl_out,
                "sisa": total_sisa_all,
                "persentase": persen_nasional
            }
            
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
                if HAS_PPTX:
                    pptx_file = generate_pptx_presentation(metrics_dict, monthly_sums)
                    st.download_button(
                        label="📊 Download Slide Presentasi Pimpinan (.pptx)",
                        data=pptx_file,
                        file_name="PRESENTASI_KINERJA_ANGGARAN_2026.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="secondary",
                        use_container_width=True
                    )
                else:
                    st.warning("⚠️ Fitur PowerPoint (.pptx) belum aktif. Tambahkan `python-pptx` ke file `requirements.txt` jika ingin mengaktifkannya.")

        except Exception as e:
            st.error(f"⚠️ Terjadi kesalahan pada saat pemrosesan: {e}")
