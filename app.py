# ==========================================
            # FITUR BARU: RINCIAN TOTAL SISA SUB KOMPONEN
            # ==========================================
            st.markdown("---")
            st.markdown("### 📋 Rincian Proporsi Total Sisa Anggaran Sub Komponen per Komponen")
            
            # Kita menggunakan sub_component_sisa yang ditarik langsung secara persis dari file LRA
            if sub_component_sisa:
                komp_keys_sisa = list(sub_component_sisa.keys())
                has_any_sisa = False
                
                sisa_cols = st.columns(2)
                
                for idx, komp_name in enumerate(komp_keys_sisa):
                    sisa_dict = sub_component_sisa[komp_name]
                    
                    # Pembersihan nilai: Pastikan val berupa angka nyata dan > 0, hindari NaN atau Infinite
                    sisa_dict_filtered = {}
                    for k, v in sisa_dict.items():
                        try:
                            # Mengecek apakah nilainya merupakan angka numerik yang valid dan lebih dari nol
                            if pd.notna(v) and float(v) > 0 and np.isfinite(float(v)):
                                sisa_dict_filtered[k] = float(v)
                        except (ValueError, TypeError):
                            continue
                    
                    with sisa_cols[idx % 2]:
                        st.markdown(f"**{komp_name}**")
                        
                        # Hanya coba membuat grafik jika masih ada data bernilai positif
                        if len(sisa_dict_filtered) > 0 and sum(sisa_dict_filtered.values()) > 0:
                            has_any_sisa = True
                            sub_labels_sisa = list(sisa_dict_filtered.keys())
                            sub_values_sisa = list(sisa_dict_filtered.values())
                            total_komp_sisa_val = sum(sub_values_sisa)
                            
                            # 1. Grafik Donut Sisa
                            fig_donut_sisa = go.Figure(data=[go.Pie(
                                labels=[f"Sub {l.split('-')[0].strip()}" for l in sub_labels_sisa],
                                values=sub_values_sisa,
                                hole=0.45,
                                textinfo='percent+label',
                                hoverinfo='none',
                                textfont_size=11,
                                marker=dict(colors=chart_colors[:len(sub_labels_sisa)], line=dict(color='#FFFFFF', width=2))
                            )])
                            fig_donut_sisa.update_layout(
                                plot_bgcolor="rgba(0,0,0,0)",
                                paper_bgcolor="rgba(0,0,0,0)",
                                margin=dict(t=20, b=20, l=20, r=20),
                                showlegend=False,
                                height=280
                            )
                            st.plotly_chart(fig_donut_sisa, use_container_width=True)
                            
                            # 2. Header Tabel Sisa
                            head_s = st.columns([0.5, 4.5, 2.5, 2])
                            with head_s[0]: st.markdown("")
                            with head_s[1]: st.markdown("**Sub Komponen**")
                            with head_s[2]: st.markdown("**Total Sisa (Rp)**")
                            with head_s[3]: st.markdown("**Persentase**")
                            st.markdown("<hr style='margin: 4px 0px 8px 0px;'>", unsafe_allow_html=True)

                            # 3. Baris Data Sisa
                            for i, (label_sisa, val_sisa) in enumerate(zip(sub_labels_sisa, sub_values_sisa)):
                                color_hex_sisa = chart_colors[i % len(chart_colors)]
                                pct_sisa = (val_sisa / total_komp_sisa_val) * 100
                                
                                row_s = st.columns([0.5, 4.5, 2.5, 2])
                                with row_s[0]:
                                    st.markdown(f"<div style='width:14px; height:14px; background-color:{color_hex_sisa}; border-radius:3px; margin-top:5px;'></div>", unsafe_allow_html=True)
                                with row_s[1]:
                                    st.markdown(f"<span style='font-size:12px; color:#212529;'>{label_sisa}</span>", unsafe_allow_html=True)
                                with row_s[2]:
                                    st.markdown(f"<span style='font-size:12px; font-weight:500; color:#212529;'>Rp {val_sisa:,.0f}</span>", unsafe_allow_html=True)
                                with row_s[3]:
                                    st.markdown(f"<span style='font-size:12px; font-weight:600; color:#495057;'>{pct_sisa:.2f}%</span>", unsafe_allow_html=True)
                        else:
                            st.info("Seluruh anggaran sub komponen telah terealisasi secara penuh (Sisa Rp 0).")
                            
                if not has_any_sisa and len(komp_keys_sisa) == 0:
                     st.info("Data sisa anggaran per sub komponen belum tersedia.")
            else:
                st.info("Data sisa anggaran per sub komponen belum tersedia.")
