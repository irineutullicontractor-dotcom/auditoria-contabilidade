import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador Contábil", layout="wide")
st.title("📊 Integração Folha ADM -> CONT")

def converter_valor(valor):
    """Converte valores no formato brasileiro para decimal."""
    if pd.isna(valor): return 0.0
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

col1, col2 = st.columns(2)
with col1:
    file_adm = st.file_uploader("1. Folha - 02-2026 - ADM", type=["xlsx", "csv"])
with col2:
    file_cont = st.file_uploader("2. FOLHA 02 - CONT", type=["xlsx", "csv"])

if file_adm and file_cont:
    try:
        # --- 1. MAPEAMENTO DE EVENTOS (ADM) ---
        df_adm = pd.read_excel(file_adm, skiprows=1) if file_adm.name.endswith('xlsx') else pd.read_csv(file_adm, skiprows=1)
        
        mapa_eventos = {}
        for _, row in df_adm.iterrows():
            # Tenta ler Colunas A/D e F/I
            for ev_idx, val_idx in [(0, 3), (5, 8)]:
                try:
                    ev_bruto = str(row.iloc[ev_idx]).split('.')[0].strip()
                    if ev_bruto.isdigit():
                        mapa_eventos[int(ev_bruto)] = mapa_eventos.get(int(ev_bruto), 0) + converter_valor(row.iloc[val_idx])
                except: pass

        # --- 2. PROCESSAMENTO DA PLANILHA CONT ---
        # Lemos sem cabeçalho para procurar a palavra "CONTA" manualmente
        df_cont_raw = pd.read_excel(file_cont, header=None) if file_cont.name.endswith('xlsx') else pd.read_csv(file_cont, header=None)

        idx_header = None
        for i, row in df_cont_raw.head(20).iterrows(): # Procura nas primeiras 20 linhas
            row_str = [str(v).upper().strip() for v in row.values]
            if "CONTA" in row_str:
                idx_header = i
                break
        
        if idx_header is not None:
            df_cont = df_cont_raw.iloc[idx_header:].copy()
            df_cont.columns = df_cont.iloc[0]
            df_cont = df_cont[1:].reset_index(drop=True)
            df_cont.columns = [str(c).strip().upper() for c in df_cont.columns] # Normaliza nomes

            def calcular_soma(texto):
                if pd.isna(texto): return 0.0
                codigos = re.findall(r'\d+', str(texto))
                return sum(mapa_eventos.get(int(c), 0.0) for c in codigos)

            # Define as colunas de trabalho
            col_origem = "CONTA"
            col_destino = "VLR LANÇAMENTO"

            if col_origem in df_cont.columns:
                df_cont[col_destino] = df_cont[col_origem].apply(calcular_soma)
                
                st.success("✅ Processado com sucesso!")
                st.dataframe(df_cont[[col_origem, col_destino]].head(20))

                # Preparar Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_cont.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Baixar FOLHA 02 - CONT Atualizada",
                    data=output.getvalue(),
                    file_name="FOLHA_02_CONT_ATUALIZADA.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Coluna 'CONTA' não identificada após processamento do cabeçalho.")
        else:
            st.error("Não foi possível encontrar a palavra 'CONTA' nas primeiras linhas do arquivo.")

    except Exception as e:
        st.error(f"Erro inesperado: {e}")
