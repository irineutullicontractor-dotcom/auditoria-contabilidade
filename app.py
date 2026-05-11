import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador Contábil v4", layout="wide")

def converter_para_float(valor):
    """Converte qualquer formato (1.250,50 ou 1250.5) para float."""
    if pd.isna(valor) or valor == '': return 0.0
    # Remove pontos de milhar e ajusta a vírgula decimal
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

st.title("📊 Integração Folha ADM -> CONT (Aba ADMIN)")

file_adm = st.file_uploader("1. Planilha Folha - 02-2026 - ADM", type=["xlsx"])
file_cont = st.file_uploader("2. Planilha FOLHA 02 - CONT", type=["xlsx"])

if file_adm and file_cont:
    try:
        # --- 1. MAPEAMENTO DE EVENTOS (ADM) ---
        df_adm = pd.read_excel(file_adm, skiprows=1)
        
        mapa_eventos = {}
        for _, row in df_adm.iterrows():
            # Blocos: A/D (0,3) e F/I (5,8)
            for ev_idx, val_idx in [(0, 3), (5, 8)]:
                try:
                    # Converte o código para string, remove o .0 se existir e limpa
                    cod_str = str(row.iloc[ev_idx]).split('.')[0].strip()
                    if cod_str.isdigit():
                        cod_int = int(cod_str)
                        valor_num = converter_para_float(row.iloc[val_idx])
                        mapa_eventos[cod_int] = mapa_eventos.get(cod_int, 0) + valor_num
                except:
                    continue

        # --- 2. PROCESSAMENTO DA PLANILHA CONT (Aba ADMIN) ---
        # Lemos sem forçar dtype para evitar o erro de 'str' vs 'int'
        xls = pd.ExcelFile(file_cont)
        if 'ADMIN' not in xls.sheet_names:
            st.error(f"Erro: A aba 'ADMIN' não foi encontrada. As abas disponíveis são: {xls.sheet_names}")
        else:
            df_cont = pd.read_excel(xls, sheet_name='ADMIN', header=None)

            # Garante que existam colunas até a H (índice 7)
            while df_cont.shape[1] < 8:
                df_cont[df_cont.shape[1]] = ""

            def extrair_e_somar(celula):
                if pd.isna(celula): return 0.0
                # Pega todos os números na célula B (ex: 16, 50, 101...)
                numeros = re.findall(r'\d+', str(celula))
                return sum(mapa_eventos.get(int(n), 0.0) for n in numeros)

            # Processamento: Linha 4 (índice 3) em diante
            # Coluna B (1) e Coluna H (7)
            for i in range(len(df_cont)):
                if i >= 3:
                    conteudo_b = df_cont.iloc[i, 1]
                    if pd.notna(conteudo_b):
                        soma_total = extrair_e_somar(conteudo_b)
                        # Gravamos o valor como número direto
                        df_cont.iat[i, 7] = soma_total

            st.success("✅ Processamento concluído!")
            st.write("Prévia dos valores calculados (Coluna H):")
            st.dataframe(df_cont.iloc[3:].head(15))

            # --- 3. EXPORTAÇÃO ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_cont.to_excel(writer, index=False, header=False, sheet_name='ADMIN')
            
            st.download_button(
                label="📥 Baixar Planilha Pronta",
                data=output.getvalue(),
                file_name="FOLHA_02_CONT_ATUALIZADA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Erro inesperado: {e}")
