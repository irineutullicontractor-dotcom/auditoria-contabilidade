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
            df_eventos = pd.read_csv(file_adm, skiprows=1, sep=None, engine='python', on_bad_lines='skip')
        else:
            df_eventos = pd.read_excel(file_adm, skiprows=1)

        mapa_eventos = {}
        for _, row in df_eventos.iterrows():
            try:
                # Lado Esquerdo: Evento (A/0), Valor (D/3)
                ev_esq = str(row.iloc[0]).strip()
                val_esq = float(str(row.iloc[3]).replace(',', '.'))
                if ev_esq.replace('.0','').isdigit():
                    mapa_eventos[int(float(ev_esq))] = mapa_eventos.get(int(float(ev_esq)), 0) + val_esq
            except: pass
            
            try:
                # Lado Direito: Evento (F/5), Valor (I/8)
                ev_dir = str(row.iloc[5]).strip()
                val_dir = float(str(row.iloc[8]).replace(',', '.'))
                if ev_dir.replace('.0','').isdigit():
                    mapa_eventos[int(float(ev_dir))] = mapa_eventos.get(int(float(ev_dir)), 0) + val_dir
            except: pass

        # 2. PROCESSAMENTO DA FOLHA CONT (Destino)
        # Usamos header=None para evitar o erro de colunas duplicadas no início
        if file_cont.name.endswith('csv'):
            df_dest = pd.read_csv(file_cont, header=None, sep=None, engine='python')
        else:
            df_dest = pd.read_excel(file_cont, header=None)

        # Localizamos a linha que contém a palavra "CONTA" para definir como cabeçalho
        linha_cabecalho = 0
        for i, row in df_dest.iterrows():
            if "CONTA" in row.values:
                linha_cabecalho = i
                break
        
        # Reprocessamos com o cabeçalho correto
        df_dest.columns = df_dest.iloc[linha_cabecalho]
        df_dest = df_dest.iloc[linha_cabecalho + 1:].reset_index(drop=True)
        
        # Remove colunas duplicadas que o Pandas cria automaticamente
        df_dest = df_dest.loc[:, ~df_dest.columns.duplicated()]

        def calcular_soma(texto):
            if pd.isna(texto): return 0.0
            codigos = re.findall(r'\d+', str(texto))
            return sum(mapa_eventos.get(int(cod), 0.0) for cod in codigos)

        # Aplica a lógica na coluna CONTA
        if "CONTA" in df_dest.columns:
            df_dest["VLR LANÇAMENTO"] = df_dest["CONTA"].apply(calcular_soma)
            
            st.success("✅ Processado com sucesso!")
            st.dataframe(df_dest.head(20))

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
        else:
            st.error("Erro: Não foi possível encontrar a coluna 'CONTA' na planilha de destino.")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
