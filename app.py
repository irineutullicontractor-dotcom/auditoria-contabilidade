import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador Contábil Final", layout="wide")

def limpar_e_converter(valor):
    if pd.isna(valor): return 0.0
    # Trata formato 1.250,50 ou 1250,50
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

st.title("📊 Integração Folha ADM -> CONT (Aba ADMIN)")

file_adm = st.file_uploader("1. Planilha Folha - 02-2026 - ADM", type=["xlsx"])
file_cont = st.file_uploader("2. Planilha FOLHA 02 - CONT", type=["xlsx"])

if file_adm and file_cont:
    try:
        # 1. MAPEAMENTO DE EVENTOS (ADM)
        # Lendo a folha de eventos
        df_adm = pd.read_excel(file_adm, skiprows=1)
        mapa_eventos = {}
        for _, row in df_adm.iterrows():
            for ev_idx, val_idx in [(0, 3), (5, 8)]:
                try:
                    cod = str(row.iloc[ev_idx]).split('.')[0].strip()
                    if cod.isdigit():
                        mapa_eventos[int(cod)] = mapa_eventos.get(int(cod), 0) + limpar_e_converter(row.iloc[val_idx])
                except: pass

        # 2. PROCESSAMENTO DA PLANILHA CONT (Aba ADMIN)
        # CRUCIAL: Especificar a aba 'ADMIN' para não pegar 'Centro de Custo'
        df_cont = pd.read_excel(file_cont, sheet_name='ADMIN', header=None)

        # Garante que a coluna H (índice 7) exista
        while df_cont.shape[1] < 8:
            df_cont[df_cont.shape[1]] = 0.0

        def somar_codigos(texto):
            if pd.isna(texto): return 0.0
            cods = re.findall(r'\d+', str(texto))
            return sum(mapa_eventos.get(int(c), 0.0) for c in cods)

        # Processa a partir da linha 4 (índice 3)
        for i in range(len(df_cont)):
            if i >= 3:
                conteudo_origem = df_cont.iloc[i, 1] # Coluna B
                if pd.notna(conteudo_origem):
                    df_cont.iat[i, 7] = somar_codigos(conteudo_origem) # Coluna H

        st.success("✅ Processamento da aba ADMIN concluído!")
        st.dataframe(df_cont.iloc[3:].head(15))

        # 3. EXPORTAÇÃO
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Salva na aba ADMIN para manter a compatibilidade
            df_cont.to_excel(writer, index=False, header=False, sheet_name='ADMIN')
        
        st.download_button(
            label="📥 Baixar Planilha Atualizada",
            data=output.getvalue(),
            file_name="FOLHA_02_CONT_PRONTA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro: {e}. Verifique se a aba se chama exatamente 'ADMIN'.")
