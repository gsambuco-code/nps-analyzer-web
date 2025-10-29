import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Configurazione della pagina
st.set_page_config(
    page_title="NPS Analyzer Web",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Titolo principale
st.title("📊 NPS Analyzer - Web Application")
st.markdown("Analizza i dati NPS e genera report automatici - Accessibile da qualsiasi dispositivo")

# Inizializzazione session state
if 'data' not in st.session_state:
    st.session_state.data = None
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False
if 'nps_column' not in st.session_state:
    st.session_state.nps_column = None

# Sidebar per upload file
with st.sidebar:
    st.header("📁 Caricamento Dati")
    
    uploaded_file = st.file_uploader(
        "Carica il tuo file Excel",
        type=['xlsx', 'xls'],
        help="Carica un file Excel contenente i dati NPS"
    )
    
    if uploaded_file is not None:
        try:
            st.session_state.data = pd.read_excel(uploaded_file, sheet_name=0)
            st.success(f"✅ File caricato con successo! ({len(st.session_state.data)} righe)")
            
            # Mostra anteprima dati
            with st.expander("Anteprima dati"):
                st.dataframe(st.session_state.data.head(), use_container_width=True)
                
        except Exception as e:
            st.error(f"Errore nel caricamento del file: {str(e)}")

# Funzione per trovare colonne
def find_column(possible_names, data_columns):
    """Cerca una colonna tra vari nomi possibili"""
    for col in data_columns:
        col_upper = str(col).upper()
        for name in possible_names:
            if name.upper() in col_upper:
                return col
    return None

# Sezione Analisi Dati
st.header("🔍 Analisi Dati")

if st.session_state.data is not None:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if st.button("🚀 Analizza Dati NPS", type="primary", use_container_width=True):
            with st.spinner("Analisi in corso..."):
                try:
                    data = st.session_state.data
                    
                    # Log colonne disponibili
                    st.info(f"Colonne trovate: {list(data.columns)}")
                    
                    # Cerca colonna NPS
                    nps_column = find_column([
                        'NPS SCORE', 'NPS', 'SCORE NPS', 
                        'VALUTAZIONE NPS', 'NPS_VALUE'
                    ], data.columns)
                    
                    if nps_column:
                        st.session_state.nps_column = nps_column
                        nps_scores = data[nps_column].dropna()
                        total_records = len(data)
                        total_responses = len(nps_scores)
                        
                        # Calcola distribuzione NPS
                        detractors = len(nps_scores[nps_scores <= 6])
                        neutrals = len(nps_scores[(nps_scores >= 7) & (nps_scores <= 8)])
                        promoters = len(nps_scores[nps_scores >= 9])
                        
                        if total_responses > 0:
                            nps_score = ((promoters - detractors) / total_responses) * 100
                        else:
                            nps_score = 0
                        
                        # Salva risultati in session state
                        st.session_state.analysis_results = {
                            'total_records': total_records,
                            'total_responses': total_responses,
                            'detractors': detractors,
                            'neutrals': neutrals,
                            'promoters': promoters,
                            'nps_score': nps_score
                        }
                        
                        st.session_state.analysis_done = True
                        st.success("Analisi completata con successo!")
                        
                    else:
                        st.error("Colonna NPS non trovata nel dataset")
                        
                except Exception as e:
                    st.error(f"Errore durante l'analisi: {str(e)}")
    
    with col2:
        st.metric("Righe Caricate", len(st.session_state.data))

# Visualizzazione Risultati Analisi
if st.session_state.analysis_done:
    results = st.session_state.analysis_results
    
    st.subheader("📈 Risultati Analisi NPS")
    
    # Metriche principali
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Record Totali", results['total_records'])
    with col2:
        st.metric("Risposte NPS", results['total_responses'])
    with col3:
        st.metric("Tasso di Risposta", 
                 f"{(results['total_responses']/results['total_records']*100):.1f}%")
    with col4:
        st.metric("NPS Score", f"{results['nps_score']:.1f}%")
    
    # Grafico distribuzione NPS
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Dati per il grafico
        categories = ['Detrattori (0-6)', 'Neutri (7-8)', 'Promotori (9-10)']
        values = [results['detractors'], results['neutrals'], results['promoters']]
        colors = ['#FF4B4B', '#FFA500', '#00D4AA']
        
        fig = go.Figure(data=[go.Pie(
            labels=categories, 
            values=values,
            hole=.3,
            marker_colors=colors
        )])
        fig.update_layout(
            title="Distribuzione NPS",
            showlegend=True
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Tabella riassuntiva
        summary_data = {
            'Categoria': ['Detrattori', 'Neutri', 'Promotori'],
            'Conteggio': [results['detractors'], results['neutrals'], results['promoters']],
            'Percentuale': [
                f"{(results['detractors']/results['total_responses']*100):.1f}%",
                f"{(results['neutrals']/results['total_responses']*100):.1f}%",
                f"{(results['promoters']/results['total_responses']*100):.1f}%"
            ]
        }
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True)

# Sezione Report
st.header("📊 Generazione Report")

if st.session_state.analysis_done:
    tab1, tab2 = st.tabs(["📋 Report Detrattori", "🏆 Classifica Zone"])
    
    with tab1:
        st.subheader("Report DETRATTORI (NPS < 7)")
        
        if st.button("Genera Report Detrattori", type="secondary"):
            with st.spinner("Generazione report in corso..."):
                try:
                    data = st.session_state.data
                    nps_col = st.session_state.nps_column
                    
                    # Filtra detrattori
                    filtered_data = data[data[nps_col] < 7].copy()
                    
                    if len(filtered_data) == 0:
                        st.warning("Nessun detrattore trovato nel dataset")
                    else:
                        # Cerca le colonne necessarie
                        ldv_col = find_column(['LDV'], data.columns)
                        zona_col = find_column(['ZONA_PTL', 'ZONA', 'AREA'], data.columns)
                        verbatim_col = find_column([
                            'NPS - VERBATIM DETRATTORI', 
                            'VERBATIM', 
                            'COMMENTI'
                        ], data.columns)
                        condizione_col = find_column([
                            'CONDIZIONE DEL PACCO', 
                            'CONDIZIONE'
                        ], data.columns)
                        luogo_col = find_column([
                            'LUOGO DI CONSEGNA DEL PACCO',
                            'LUOGO CONSEGNA'
                        ], data.columns)
                        cliente_col = find_column([
                            'RAGIONE SOCIALE CLIENTE',
                            'CLIENTE'
                        ], data.columns)
                        
                        # Crea report
                        report_data = pd.DataFrame()
                        
                        # Popola le colonne del report
                        if ldv_col:
                            report_data['LDV'] = filtered_data[ldv_col]
                        else:
                            report_data['LDV'] = "N/D"
                        
                        report_data['NPS SCORE'] = filtered_data[nps_col]
                        
                        if zona_col:
                            report_data['ZONA_PTL'] = filtered_data[zona_col]
                        else:
                            report_data['ZONA_PTL'] = "N/D"
                        
                        if verbatim_col:
                            report_data['NPS - Verbatim Detrattori'] = filtered_data[verbatim_col].fillna('')
                        else:
                            report_data['NPS - Verbatim Detrattori'] = "N/D"
                        
                        # Combinazione condizioni
                        condizione_pacco = filtered_data[condizione_col].fillna('') if condizione_col else pd.Series([''] * len(filtered_data))
                        luogo_consegna = filtered_data[luogo_col].fillna('') if luogo_col else pd.Series([''] * len(filtered_data))
                        
                        report_data['CONDIZIONE'] = condizione_pacco.astype(str) + " - " + luogo_consegna.astype(str)
                        report_data['CONDIZIONE'] = report_data['CONDIZIONE'].str.replace('^ - $', '', regex=True)
                        
                        if cliente_col:
                            report_data['CLIENTE'] = filtered_data[cliente_col]
                        else:
                            report_data['CLIENTE'] = "N/D"
                        
                        # Ordina per zona (se disponibile)
                        if zona_col:
                            conteggio_per_zona = filtered_data.groupby(zona_col).size().reset_index(name='CONTEggio_DETRATTORI')
                            conteggio_per_zona = conteggio_per_zona.sort_values('CONTEggio_DETRATTORI', ascending=True)
                            
                            ordine_zone = {}
                            for idx, (index, row) in enumerate(conteggio_per_zona.iterrows(), 1):
                                ordine_zone[row[zona_col]] = idx
                            
                            report_data['ORDINE_ZONA'] = report_data['ZONA_PTL'].map(ordine_zone)
                            report_data = report_data.sort_values('ORDINE_ZONA', ascending=True)
                            report_data = report_data.drop('ORDINE_ZONA', axis=1)
                        
                        # Mostra anteprima
                        st.success(f"Report generato: {len(report_data)} record di detrattori")
                        st.dataframe(report_data, use_container_width=True)
                        
                        # Download
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            report_data.to_excel(writer, sheet_name='DETRATTORI', index=False)
                            
                            # Formattazione
                            workbook = writer.book
                            worksheet = writer.sheets['DETRATTORI']
                            
                            # Auto-adjust column widths
                            for i, col in enumerate(report_data.columns):
                                max_len = max(report_data[col].astype(str).str.len().max(), len(col)) + 2
                                worksheet.set_column(i, i, min(max_len, 50))
                        
                        st.download_button(
                            label="📥 Scarica Report Detrattori (Excel)",
                            data=output.getvalue(),
                            file_name=f"report_detrattori_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.ms-excel",
                            type="primary"
                        )
                        
                except Exception as e:
                    st.error(f"Errore nella generazione del report: {str(e)}")
    
    with tab2:
        st.subheader("Classifica Zone per Performance NPS")
        
        if st.button("Genera Report Classifica", type="secondary"):
            with st.spinner("Generazione classifica in corso..."):
                try:
                    data = st.session_state.data
                    nps_col = st.session_state.nps_column
                    
                    # Trova colonna zona
                    zona_col = find_column(['ZONA_PTL', 'ZONA', 'AREA'], data.columns)
                    
                    if not zona_col:
                        st.error("Colonna Zona non trovata nel dataset")
                    else:
                        # Calcola statistiche per zona
                        zone_data = []
                        for zona in data[zona_col].unique():
                            if pd.isna(zona):
                                continue
                                
                            zone_records = data[data[zona_col] == zona]
                            scores = zone_records[nps_col].dropna()
                            
                            if len(scores) == 0:
                                continue
                            
                            detrattori = len(scores[scores <= 6])
                            neutri = len(scores[(scores >= 7) & (scores <= 8)])
                            promotori = len(scores[scores >= 9])
                            totale_risposte = len(scores)
                            
                            nps_percent = ((promotori - detrattori) / totale_risposte) * 100
                            nps_riclassificato = (((promotori - detrattori) / totale_risposte) + 
                                                ((promotori + neutri) / 40)) * 100
                            
                            zone_data.append({
                                'ZONA_PTL': zona,
                                'Detrattori (0-6)': detrattori,
                                'Neutri (7-8)': neutri,
                                'Promotori (9-10)': promotori,
                                'Totale Risposte': totale_risposte,
                                'NPS %': nps_percent,
                                'NPS RICLASSIFICATO%': nps_riclassificato
                            })
                        
                        if zone_data:
                            classifica_df = pd.DataFrame(zone_data)
                            classifica_df = classifica_df.sort_values('NPS RICLASSIFICATO%', ascending=False)
                            
                            # Reset index per la classifica
                            classifica_df.reset_index(drop=True, inplace=True)
                            classifica_df.index += 1
                            
                            st.success(f"Classifica generata per {len(classifica_df)} zone")
                            st.dataframe(classifica_df, use_container_width=True)
                            
                            # Grafico classifica
                            fig = px.bar(
                                classifica_df.head(10),
                                x='ZONA_PTL',
                                y='NPS RICLASSIFICATO%',
                                title="Top 10 Zone per NPS Riclassificato",
                                color='NPS RICLASSIFICATO%',
                                color_continuous_scale='Viridis'
                            )
                            fig.update_layout(xaxis_tickangle=-45)
                            st.plotly_chart(fig, use_container_width=True)
                            
                            # Download
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                classifica_df.to_excel(writer, sheet_name='CLASSIFICA', index=True, index_label='Posizione')
                                
                                # Formattazione
                                workbook = writer.book
                                worksheet = writer.sheets['CLASSIFICA']
                                
                                # Formattazione percentuali
                                percent_format = workbook.add_format({'num_format': '0.00%'})
                                worksheet.set_column('F:G', 15, percent_format)
                            
                            st.download_button(
                                label="📥 Scarica Report Classifica (Excel)",
                                data=output.getvalue(),
                                file_name=f"report_classifica_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                mime="application/vnd.ms-excel",
                                type="primary"
                            )
                        else:
                            st.warning("Nessuna zona con dati NPS sufficienti trovata")
                            
                except Exception as e:
                    st.error(f"Errore nella generazione della classifica: {str(e)}")

else:
    st.info("👈 Carica un file Excel e analizza i dati per generare i report")

# Footer
st.markdown("---")
st.markdown(
    "**NPS Analyzer Web App** - "
    "Accessibile da qualsiasi dispositivo connesso a internet | "
    f"Ultimo aggiornamento: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
)