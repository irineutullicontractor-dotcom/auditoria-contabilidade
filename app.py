import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador de Folha", layout="wide")
st.title("📊 Integração Contábil")

def limpar_e_converter(valor):
    """Converte '1.250,50' ou 1250.5 em float 1250.5"""
    if pd.isna(valor): return 0.0
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

col1, col2 = st.columns(2)
with col1:
    file_adm = st.file_uploader("Arquivo 1: Folha - 02-2026 - ADM", type=["xlsx", "csv"])
with col2:
    file_cont = st.file_uploader("Arquivo 2: FOLHA 02 - CONT", type=["xlsx", "csv"])

if file_adm and file_cont:
    try:
        # --- 1. LER EVENTOS DA FOLHA ADM ---
        # Pulamos as linhas iniciais de título do relatório
        df_adm = pd.read_excel(file_adm, skiprows=1) if file_adm.name.endswith('xlsx') else pd.read_csv(file_adm, skiprows=1)
        
        mapa_eventos = {}
        for _, row in df_adm.iterrows():
            # Lado Esquerdo: Evento (Col A/0) e Valor (Col D/3)
            try:
                cod_esq = str(row.iloc[0]).split('.')[0].strip()
                if cod_esq.isdigit():
                    mapa_eventos[int(cod_esq)] = mapa_eventos.get(int(cod_esq), 0) + limpar_e_converter(row.iloc[3])
            except: pass
            
            # Lado Direito: Evento (Col F/5) e Valor (Col I/8)
            try:
                cod_dir = str(row.iloc[5]).split('.')[0].strip()
                if cod_dir.isdigit():
                    mapa_eventos[int(cod_dir)] = mapa_eventos.get(int(cod_dir), 0) + limpar_e_converter(row.iloc[8])
            except: pass

        # --- 2. LER FOLHA 02 - CONT (ABA ADMIN) ---
        # Forçamos a leitura sem processar cabeçalho para não dar erro de nome
        df_cont = pd.read_excel(file_cont, header=None) if file_cont.name.endswith('xlsx') else pd.read_csv(file_cont, header=None)

        def somar_codigos(texto):
            if pd.isna(texto): return 0.0
            # Busca todos os números na frase (ex: 16, 50, 101...)
            encontrados = re.findall(r'\d+', str(texto))
            return sum(mapa_eventos.get(int(c), 0.0) for c in encontrados)

        # A mágica acontece aqui:
        # Percorremos as linhas a partir da 4 (onde começam os dados de fato)
        # Coluna B é índice 1 | Coluna H é índice 7
        
        # Criamos uma cópia dos dados para não dar erro de visualização
        df_final = df_cont.copy()
        
        for i in range(len(df_final)):
            # Só processamos se houver algo na coluna B e se não for o título
            conteudo_b = df_final.iloc[i, 1]
            if i >= 4 and pd.notna(conteudo_b):
                soma = somar_codigos(conteudo_b)
                df_final.iloc[i, 7] = soma  # Grava na Coluna H

        st.success("✅ Processamento concluído com sucesso!")
        st.dataframe(df_final.iloc[3:].head(15)) # Mostra do cabeçalho em diante

        # --- 3. BOTÃO DE DOWNLOAD ---
        output = io.BytesIO()
        # Salvamos mantendo a formatação original (sem index)
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, header=False)
        
        st.download_button(
            label="📥 Baixar FOLHA 02 - CONT Preenchida",
            data=output.getvalue(),
            file_name="FOLHA_02_CONT_FINALIZADA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro técnico: {e}")
