import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador Contábil", layout="wide")
st.title("📊 Processador de Eventos Contábeis")

col1, col2 = st.columns(2)
with col1:
    file_adm = st.file_uploader("Upload: Folha - 02-2026 - ADM", type=["xlsx", "csv"])
with col2:
    file_cont = st.file_uploader("Upload: FOLHA 02 - CONT", type=["xlsx", "csv"])

if file_adm and file_cont:
    try:
        # 1. LEITURA DOS EVENTOS (ADM)
        if file_adm.name.endswith('csv'):
            df_eventos = pd.read_csv(file_adm, skiprows=1, sep=None, engine='python')
        else:
            df_eventos = pd.read_excel(file_adm, skiprows=1)

        mapa_eventos = {}
        for _, row in df_eventos.iterrows():
            try:
                # Eventos e Valores (A e D | F e I)
                for ev_idx, val_idx in [(0, 3), (5, 8)]:
                    ev = str(row.iloc[ev_idx]).strip().split('.')[0] # Remove .0 se houver
                    val = str(row.iloc[val_idx]).replace('.', '').replace(',', '.')
                    if ev.isdigit():
                        mapa_eventos[int(ev)] = mapa_eventos.get(int(ev), 0) + float(val)
            except: pass

        # 2. LEITURA DA PLANILHA DE DESTINO (CONT)
        if file_cont.name.endswith('csv'):
            df_dest = pd.read_csv(file_cont, header=None, sep=None, engine='python')
        else:
            df_dest = pd.read_excel(file_cont, header=None)

        # LÓGICA DE IDENTIFICAÇÃO DE COLUNA
        # Procuramos em qual linha e coluna está a palavra "CONTA"
        found_row, found_col = None, None
        for r in range(min(len(df_dest), 15)): # Procura nas primeiras 15 linhas
            for c in range(len(df_dest.columns)):
                val_celula = str(df_dest.iloc[r, c]).upper()
                if "CONTA" in val_celula:
                    found_row, found_col = r, c
                    break
            if found_row is not None: break

        # Se encontrou a linha do cabeçalho
        if found_row is not None:
            df_dest.columns = df_dest.iloc[found_row]
            df_dest = df_dest.iloc[found_row + 1:].reset_index(drop=True)
            col_procura = df_dest.columns[found_col]
        else:
            # Caso não encontre a palavra "CONTA", assume a Coluna B (índice 1)
            st.warning("⚠️ Não achei o nome 'CONTA', usando a 2ª coluna por padrão.")
            col_procura = df_dest.columns[1]

        def calcular_soma(texto):
            if pd.isna(texto): return 0.0
            codigos = re.findall(r'\d+', str(texto))
            return sum(mapa_eventos.get(int(cod), 0.0) for cod in codigos)

        # Garantir que a coluna VLR LANÇAMENTO exista
        # Se houver coluna com nome similar, usamos ela, senão criamos
        col_valor = "VLR LANÇAMENTO"
        df_dest[col_valor] = df_dest[col_procura].apply(calcular_soma)

        st.success("✅ Processado!")
        st.dataframe(df_dest.head(15))

        # Gerar arquivo para Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_dest.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Baixar Planilha Atualizada",
            data=output.getvalue(),
            file_name="FOLHA_02_CONT_ATUALIZADA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro Crítico: {e}")
