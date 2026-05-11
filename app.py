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
        # 1. PROCESSAMENTO DA FOLHA ADM (Leitura dos Eventos)
        if file_adm.name.endswith('csv'):
            df_eventos = pd.read_csv(file_adm, skiprows=1, sep=None, engine='python')
        else:
            df_eventos = pd.read_excel(file_adm, skiprows=1)

        mapa_eventos = {}

        for _, row in df_eventos.iterrows():
            # Lado Esquerdo (Coluna A - Evento, Coluna D - Valor)
            try:
                ev_esq = str(row.iloc[0]).strip()
                val_esq = float(row.iloc[3])
                if ev_esq.isdigit():
                    mapa_eventos[int(float(ev_esq))] = mapa_eventos.get(int(float(ev_esq)), 0) + val_esq
            except: pass
            
            # Lado Direito (Coluna F - Evento, Coluna I - Valor)
            try:
                ev_dir = str(row.iloc[5]).strip()
                val_dir = float(row.iloc[8])
                if ev_dir.isdigit():
                    mapa_eventos[int(float(ev_dir))] = mapa_eventos.get(int(float(ev_dir)), 0) + val_dir
            except: pass

        # 2. PROCESSAMENTO DA FOLHA CONT (Destino)
        if file_cont.name.endswith('csv'):
            # sep=None com engine='python' detecta se é vírgula ou ponto e vírgula
            df_dest = pd.read_csv(file_cont, skiprows=3, sep=None, engine='python')
        else:
            df_dest = pd.read_excel(file_cont, skiprows=3)

        # Limpar nomes das colunas (remover espaços vazios)
        df_dest.columns = [str(c).strip() for c in df_dest.columns]

        def calcular_soma(texto):
            if pd.isna(texto): return 0.0
            codigos = re.findall(r'\d+', str(texto))
            return sum(mapa_eventos.get(int(cod), 0.0) for cod in codigos)

        # Localizar colunas dinamicamente para evitar erro de index
        # Tentamos por nome, se não achar, usamos o índice seguro
        col_procura = "CONTA" if "CONTA" in df_dest.columns else df_dest.columns[1]
        
        # Criar a coluna de valor se ela não existir ou preenchê-la
        col_destino = "VLR LANÇAMENTO"
        
        df_dest[col_destino] = df_dest[col_procura].apply(calcular_soma)

        st.success("✅ Processado com sucesso!")
        st.dataframe(df_dest)

        # Download
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
        st.error(f"Erro ao processar: {e}")
        st.info("Dica: Verifique se as planilhas seguem o formato esperado de colunas.")
