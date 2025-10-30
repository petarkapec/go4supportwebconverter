import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(
    page_title="CSV to Excel Converter",
    page_icon="📊",
    layout="centered"
)

# Naslov i opis
st.title("📊 CSV to Excel Converter")
st.markdown("""
Ova aplikacija pretvara CSV datoteku u Excel pivot tablicu.
**Upute:**
1. Odaberi CSV datoteku
2. Aplikacija će automatski generirati Excel
3. Preuzmi rezultat
""")

# Upload datoteke
uploaded_file = st.file_uploader(
    "Odaberi CSV datoteku",
    type="csv",
    help="Odaberi CSV datoteku sa stupcima 'agent_name' i 'Lead_status'"
)

if uploaded_file is not None:
    try:
        # Progress bar
        with st.spinner("Obrađujem podatke..."):
            # Učitavanje CSV-a
            data = pd.read_csv(uploaded_file, encoding='iso-8859-1')
            
            # Provjera potrebnih stupaca
            required_columns = ['agent_name', 'Lead_status']
            missing_columns = [col for col in required_columns if col not in data.columns]
            
            if missing_columns:
                st.error(f"❌ Nedostaju potrebni stupci: {', '.join(missing_columns)}")
                st.info("💡 Provjeri da li CSV sadrži stupce 'agent_name' i 'Lead_status'")
            else:
                st.success(f"✅ Učitano {len(data)} redova podataka")
                
                # Prikaz prvih redova
                with st.expander("📋 Pregled podataka (prvih 5 redova)"):
                    st.dataframe(data.head())
                
                # EKSPANZIJA STATUSA
                rows = []
                for _, row in data.iterrows():
                    statuses = [s.strip() for s in str(row['Lead_status']).split(';')]
                    for status in statuses:
                        rows.append({
                            'agent_name': row['agent_name'].strip(),
                            'Lead_status': status.strip()
                        })

                df_expanded = pd.DataFrame(rows)

                # PIVOT TABLICA
                pivot_table = pd.pivot_table(
                    df_expanded,
                    index='agent_name',
                    columns='Lead_status',
                    aggfunc='size',
                    fill_value=0
                )

                # Sortiranje
                pivot_table = pivot_table.sort_index(axis=0).sort_index(axis=1)
                
                # Prikaz pivot tablice
                with st.expander("📊 Pregled pivot tablice"):
                    st.dataframe(pivot_table)

                # Kreiranje Excel datoteke u memoriji
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    pivot_table.to_excel(writer, sheet_name='Statistika', index=True)
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Statistika']
                    
                    # Formatiranje
                    header_fmt = workbook.add_format({
                        'bold': True, 
                        'bg_color': '#D9E1F2', 
                        'border': 1, 
                        'align': 'center'
                    })
                    cell_fmt = workbook.add_format({'border': 1, 'align': 'center'})
                    
                    # Primjena formata
                    for col_num, value in enumerate(pivot_table.reset_index().columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                    
                    worksheet.set_column(0, len(pivot_table.columns), 20, cell_fmt)
                
                # Postavi pokazivač na početak
                output.seek(0)
                
                # Download button
                timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                filename = f"statistika_po_agentima_{timestamp}.xlsx"
                
                st.download_button(
                    label="📥 Preuzmi Excel datoteku",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Klikni za preuzimanje Excel datoteke"
                )
                
                st.success("✅ Excel datoteka je spremna za preuzimanje!")
                
    except Exception as e:
        st.error(f"❌ Došlo je do greške: {str(e)}")
        st.info("💡 Provjeri format CSV datoteke i encoding")

else:
    st.info("👆 Odaberi CSV datoteku za početak")

# Sidebar s informacijama
with st.sidebar:
    st.header("ℹ️ Informacije")
    st.markdown("""
    **Potrebni stupci u CSV:**
    - `agent_name`
    - `Lead_status`
    
    **Format Lead_status:**
    - Vrijednosti odvojene točka-zarezom
    - Primjer: `Status1; Status2; Status3`
    
    **Podržani encoding:**
    - ISO-8859-1
    """)