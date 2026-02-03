import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io

# 1. Konfigurasi Halaman
st.set_page_config(page_title="Smart Data Insights", layout="centered")

# 2. CSS Aesthetic
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;800&display=swap');
    html, body, [class*="st-"] { font-family: 'Inter', sans-serif; }
    .main-title {
        font-size: 50px !important; font-weight: 800; text-align: center;
        background: -webkit-linear-gradient(#e5322d, #ff8f8f);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        letter-spacing: -2px; margin-top: 10px;
    }
    .footer {
        position: fixed; bottom: 0; left: 0; width: 100%;
        padding: 10px 40px; font-size: 12px; border-top: 1px solid rgba(128, 128, 128, 0.2);
        backdrop-filter: blur(10px); background-color: rgba(255,255,255,0.05); z-index: 100;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<h1 class="main-title">Smart Data Insights</h1>', unsafe_allow_html=True)
st.write("<p style='text-align: center; opacity: 0.8;'>Sebuah solusi cerdas hehe.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Data", type=["xlsx", "csv"], label_visibility="collapsed")

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('xlsx') else pd.read_csv(uploaded_file)
        
        col1, col2 = st.columns([2, 1])
        with col1:
            keyword = st.text_input("🔍 Filter Data", placeholder="Ketik kata kunci...")
        with col2:
            val_col = st.selectbox("Pilih Kolom Nilai", df.columns)

        if keyword:
            mask = df.apply(lambda row: row.astype(str).str.contains(keyword, case=False, na=False).any(), axis=1)
            filtered_df = df[mask]

            if not filtered_df.empty:
                # Perhitungan Ringkasan
                total_val = filtered_df[val_col].sum()
                avg_val = filtered_df[val_col].mean()
                max_val = filtered_df[val_col].max()

                # Metrik di Web
                st.divider()
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Kumulatif", f"Rp {total_val:,.0f}")
                m2.metric("Rata-rata", f"Rp {avg_val:,.0f}")
                m3.metric("Tertinggi", f"Rp {max_val:,.0f}")

                # Grafik di Web
                fig = px.area(filtered_df, y=val_col, template="plotly_dark")
                fig.update_traces(line_color='#e5322d', fillcolor='rgba(229, 50, 45, 0.2)')
                st.plotly_chart(fig, use_container_width=True)
                
                # --- FITUR EXPORT PRO KE EXCEL ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    filtered_df.to_excel(writer, sheet_name='Analisis', index=False, startrow=5)
                    workbook = writer.book
                    worksheet = writer.sheets['Analisis']
                    
                    # Format Excel
                    header_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#e5322d'})
                    num_fmt = workbook.add_format({'num_format': '#,##0'})
                    
                    # Tulis Ringkasan ke Excel
                    worksheet.write('A1', f'LAPORAN: {keyword.upper()}', header_fmt)
                    worksheet.write('A2', 'Total Kumulatif:'); worksheet.write('B2', total_val, num_fmt)
                    worksheet.write('A3', 'Rata-rata:'); worksheet.write('B3', avg_val, num_fmt)
                    worksheet.write('A4', 'Tertinggi:'); worksheet.write('B4', max_val, num_fmt)

                    # Buat Grafik di Excel
                    chart = workbook.add_chart({'type': 'area'})
                    max_row = len(filtered_df) + 5
                    col_idx = filtered_df.columns.get_loc(val_col)
                    chart.add_series({
                        'name': f'Tren {keyword}',
                        'categories': ['Analisis', 6, 0, max_row, 0],
                        'values': ['Analisis', 6, col_idx, max_row, col_idx],
                        'fill': {'color': '#e5322d', 'transparency': 70}
                    })
                    worksheet.insert_chart('E1', chart)

                st.download_button(
                    label="💾 Download Laporan Excel Lengkap",
                    data=output.getvalue(),
                    file_name=f"Laporan_{keyword}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.dataframe(filtered_df, use_container_width=True)
            else:
                st.warning("Data tidak ditemukan.")
    except Exception as e:
        st.error(f"Error: {e}")

st.markdown(f'<div class="footer">© {datetime.now().year} Smart Data Insights ® - Powered by Faizal Idofi</div>', unsafe_allow_html=True)