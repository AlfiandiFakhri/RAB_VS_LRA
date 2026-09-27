# TAMPILAN GRAFIK BULANAN
            st.markdown("### 📊 Tren Penyerapan Bulanan")
            
            # REVISI BARU: Mengambil nilai maksimal (baris hierarki tertinggi) untuk menghindari double-counting
            # Karena pada LRA dan RAB, baris Program/Kegiatan mencakup total keseluruhan dari sub-sub di bawahnya
            realisasi_per_bulan = [df_preview[bulan].max() if (bulan in df_preview.columns and not df_preview[bulan].empty) else 0 for bulan in list_semua_bulan]
            
            # Memastikan tidak ada nilai NaN/kosong
            realisasi_per_bulan = [val if pd.notna(val) else 0 for val in realisasi_per_bulan]

            df_monthly_chart = pd.DataFrame({
                "Bulan": list_semua_bulan,
                "Realisasi": realisasi_per_bulan
            })
            
            # Fungsi untuk format angka standar Indonesia (Ribuan menggunakan titik)
            def format_rupiah(val):
                if val > 0:
                    return f"Rp {val:,.0f}".replace(",", ".")
                return "Rp 0"

            fig_3d_bar = px.bar(
                df_monthly_chart,
                x="Bulan",
                y="Realisasi",
                text=df_monthly_chart["Realisasi"].apply(format_rupiah),
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
