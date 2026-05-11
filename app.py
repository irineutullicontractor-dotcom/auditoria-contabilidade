import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador de Folha Contábil", layout="wide")

st.title("📊 Processador de Eventos Contábeis")
st.markdown("""
Esta ferramenta lê os eventos da planilha **Folha ADM**, soma os valores correspondentes 
e preenche a coluna **VLR LANÇAMENTO** na planilha **FOLHA 02 - CONT**.
""")

# Upload dos arquivos
col1, col2 = st.columns(2)
with col1:
    file_adm = st.file_uploader("Upload: Folha - 02-2026 - ADM", type=["xlsx", "csv"])
with col2:
    file_cont = st.file_uploader("Upload: FOLHA 02 - CONT", type=["xlsx", "csv"])

if file_adm and file_cont:
    try:
        # 1. Processamento da Planilha de Eventos (ADM)
        # Lendo a partir da linha onde começam os dados dos eventos
        if file_adm.name.endswith('csv'):
            df_eventos = pd.read_csv(file_adm, skiprows=1)
        else:
            df_eventos = pd.read_excel(file_adm, skiprows=1)

        mapa_eventos = {}

        for _, row in df_eventos.iterrows():
            # Lado Esquerdo (Coluna A - Evento, Coluna D - Valor)
            try:
                ev_esq = str(row.iloc[0]).strip()
                val_esq = float(row.iloc[3])
                if ev_esq.isdigit():
                    mapa_eventos[int(ev_esq)] = mapa_eventos.get(int(ev_esq), 0) + val_esq
            except: pass
            
            # Lado Direito (Coluna F - Evento, Coluna I - Valor)
            try:
                ev_dir = str(row.iloc[5]).strip()
                val_dir = float(row.iloc[8])
                if ev_dir.isdigit():
                    mapa_eventos[int(ev_dir)] = mapa_eventos.get(int(ev_dir), 0) + val_dir
            except: pass

        # 2. Processamento da Planilha de Destino (CONT)
        if file_cont.name.endswith('csv'):
            df_dest = pd.read_csv(file_cont, skiprows=3)
        else:
            df_dest = pd.read_excel(file_cont, skiprows=3)

        def calcular_soma(texto):
            if pd.isna(texto): return 0.0
            codigos = re.findall(r'\d+', str(texto))
            return sum(mapa_eventos.get(int(cod), 0.0) for cod in codigos)

        # Assume que a coluna de descrição é a segunda (index 1) 
        # e a de valor é a oitava (VLR LANÇAMENTO - index 7)
        desc_col_name = df_dest.columns[1]
        val_col_name = df_dest.columns[7]
        
        df_dest[val_col_name] = df_dest[desc_col_name].apply(calcular_soma)

        st.success("✅ Processamento concluído com sucesso!")
        st.dataframe(df_dest.head(10))

        # Botão para Download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_dest.to_excel(writer, index=False, sheet_name='Processado')
        
        st.download_button(
            label="📥 Baixar Planilha Processada",
            data=output.getvalue(),
            file_name="FOLHA_02_CONT_ATUALIZADA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro ao processar arquivos: {e}")
