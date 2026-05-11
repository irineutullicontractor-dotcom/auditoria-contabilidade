import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador Contábil v3", layout="wide")

def converter_valor(valor):
    if pd.isna(valor): return 0.0
    # Limpa pontos de milhar e converte vírgula para ponto
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

st.title("📊 Integração Folha ADM -> CONT")

col1, col2 = st.columns(2)
with col1:
    file_adm = st.file_uploader("Upload: Folha - 02-2026 - ADM", type=["xlsx", "csv"])
with col2:
    file_cont = st.file_uploader("Upload: FOLHA 02 - CONT", type=["xlsx", "csv"])

if file_adm and file_cont:
    try:
        # 1. MAPEAMENTO DA FOLHA ADM
        df_adm = pd.read_excel(file_adm, skiprows=1) if file_adm.name.endswith('xlsx') else pd.read_csv(file_adm, skiprows=1)
        
        mapa_eventos = {}
        for _, row in df_adm.iterrows():
            # Lado Esquerdo (Col A e D)
            try:
                ev_esq = str(row.iloc[0]).split('.')[0].strip()
                if ev_esq.isdigit():
                    mapa_eventos[int(ev_esq)] = mapa_eventos.get(int(ev_esq), 0) + converter_valor(row.iloc[3])
            except: pass
            
            # Lado Direito (Col F e I)
            try:
                ev_dir = str(row.iloc[5]).split('.')[0].strip()
                if ev_dir.isdigit():
                    mapa_eventos[int(ev_dir)] = mapa_eventos.get(int(ev_dir), 0) + converter_valor(row.iloc[8])
            except: pass

        # 2. PROCESSAMENTO DA FOLHA CONT
        # Lemos sem pular linhas primeiro para achar o cabeçalho dinamicamente
        df_cont_raw = pd.read_excel(file_cont, header=None) if file_cont.name.endswith('xlsx') else pd.read_csv(file_cont, header=None)

        # Localiza a linha que contém "CONTA"
        idx_header = 0
        for i, row in df_cont_raw.iterrows():
            if "CONTA" in [str(val).upper() for val in row.values]:
                idx_header = i
                break
        
        # Reconstrói o DF com o cabeçalho correto
        df_cont = df_cont_raw.iloc[idx_header:].copy()
        df_cont.columns = df_cont.iloc[0]
        df_cont = df_cont[1:].reset_index(drop=True)
        
        # Garante que a coluna H (VLR LANÇAMENTO) existe no DataFrame como uma série de dados
        # Isso evita o erro "iloc cannot enlarge"
        if "VLR LANÇAMENTO" not in df_cont.columns:
            df_cont["VLR LANÇAMENTO"] = 0.0

        def processar_celula(texto):
            if pd.isna(texto): return 0.0
            # Busca todos os números (códigos de evento)
            codigos = re.findall(r'\d+', str(texto))
            soma = 0.0
            for c in codigos:
                soma += mapa_eventos.get(int(c), 0.0)
            return soma

        # Aplica a soma baseada na Coluna B (CONTA)
        # Usamos at ou loc para garantir a gravação correta
        nome_col_conta = "CONTA" 
        df_cont["VLR LANÇAMENTO"] = df_cont[nome_col_conta].apply(processar_celula)

        st.success("✅ Processamento concluído!")
        st.write("Prévia dos dados atualizados:")
        st.dataframe(df_cont[[nome_col_conta, "VLR LANÇAMENTO"]].head(15))

        # Exportação
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_cont.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Baixar Planilha CONT Atualizada",
            data=output.getvalue(),
            file_name="FOLHA_02_CONT_PROCESSADA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
