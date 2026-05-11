import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador de Folha", layout="wide")
st.title("📊 Integração Contábil - Folha ADM")

# Upload dos arquivos
file_adm = st.file_uploader("1. Upload: Folha - 02-2026 - ADM (Relatório Contabilidade)", type=["xlsx", "csv"])
file_cont = st.file_uploader("2. Upload: FOLHA 02 - CONT", type=["xlsx", "csv"])

def converter_valor(valor):
    """Converte valores no formato brasileiro (ex: 1.250,50) para float."""
    if pd.isna(valor): return 0.0
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

if file_adm and file_cont:
    try:
        # --- PARTE 1: MAPEAMENTO DE EVENTOS ---
        df_eventos_raw = pd.read_excel(file_adm, skiprows=1) if file_adm.name.endswith('xlsx') else pd.read_csv(file_adm, skiprows=1)
        
        mapa_valores = {}
        for _, row in df_eventos_raw.iterrows():
            # Lado Esquerdo (Coluna A e D)
            try:
                cod = str(row.iloc[0]).split('.')[0]
                if cod.isdigit():
                    mapa_valores[int(cod)] = mapa_valores.get(int(cod), 0) + converter_valor(row.iloc[3])
            except: pass
            
            # Lado Direito (Coluna F e I)
            try:
                cod = str(row.iloc[5]).split('.')[0]
                if cod.isdigit():
                    mapa_valores[int(cod)] = mapa_valores.get(int(cod), 0) + converter_valor(row.iloc[8])
            except: pass

        # --- PARTE 2: PROCESSAMENTO DA PLANILHA CONT ---
        df_cont = pd.read_excel(file_cont, skiprows=3) if file_cont.name.endswith('xlsx') else pd.read_csv(file_cont, skiprows=3)

        def somar_eventos_na_celula(texto):
            if pd.isna(texto): return 0.0
            # Extrai todos os números da string (ex: "PROVENTOS - 16 - 50" -> [16, 50])
            codigos = re.findall(r'\d+', str(texto))
            return sum(mapa_valores.get(int(c), 0.0) for c in codigos)

        # Coluna B (índice 1) é a CONTA, Coluna H (índice 7) é o VLR LANÇAMENTO
        df_cont.iloc[:, 7] = df_cont.iloc[:, 1].apply(somar_eventos_na_celula)

        st.success("✅ Processamento finalizado com sucesso!")
        st.dataframe(df_cont)

        # Download do resultado
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_cont.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Baixar FOLHA 02 - CONT Atualizada",
            data=output.getvalue(),
            file_name="FOLHA_02_CONT_RESULTADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
