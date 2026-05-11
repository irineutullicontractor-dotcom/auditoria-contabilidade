import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador Contábil", layout="wide")

def converter_valor(valor):
    if pd.isna(valor): return 0.0
    s = str(valor).strip().replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

st.title("📊 Integração Contábil - Versão Final")

col1, col2 = st.columns(2)
with col1:
    file_adm = st.file_uploader("1. Planilha ADM (Eventos)", type=["xlsx", "csv"])
with col2:
    file_cont = st.file_uploader("2. Planilha CONT (Destino)", type=["xlsx", "csv"])

if file_adm and file_cont:
    try:
        # --- 1. MAPEAMENTO DE EVENTOS (ADM) ---
        df_adm = pd.read_excel(file_adm, skiprows=1) if file_adm.name.endswith('xlsx') else pd.read_csv(file_adm, skiprows=1)
        
        mapa_eventos = {}
        for _, row in df_adm.iterrows():
            # Lado Esquerdo: Col A(0) e D(3) | Lado Direito: Col F(5) e I(8)
            for ev_idx, val_idx in [(0, 3), (5, 8)]:
                try:
                    cod = str(row.iloc[ev_idx]).split('.')[0].strip()
                    if cod.isdigit():
                        mapa_eventos[int(cod)] = mapa_eventos.get(int(cod), 0) + converter_valor(row.iloc[val_idx])
                except: pass

        # --- 2. PROCESSAMENTO DA PLANILHA CONT ---
        # Lemos o arquivo bruto sem cabeçalho para ter controle total dos índices
        df_cont = pd.read_excel(file_cont, header=None) if file_cont.name.endswith('xlsx') else pd.read_csv(file_cont, header=None)

        # GARANTIA: Se a planilha tiver menos de 8 colunas (até a H), nós expandimos ela
        while df_cont.shape[1] < 8:
            df_cont[df_cont.shape[1]] = None

        def somar_codigos(texto):
            if pd.isna(texto): return 0.0
            cods = re.findall(r'\d+', str(texto))
            return sum(mapa_eventos.get(int(c), 0.0) for c in cods)

        # PROCESSAMENTO LINHA A LINHA
        # Conforme seu print, os dados começam na linha 4 (índice 3 do Python)
        # Coluna B (índice 1) -> Origem | Coluna H (índice 7) -> Destino
        for i in range(len(df_cont)):
            if i >= 3: # Começa na linha 4 do Excel
                conteudo_origem = df_cont.iloc[i, 1]
                if pd.notna(conteudo_origem):
                    resultado_soma = somar_codigos(conteudo_origem)
                    # Usamos .iat para evitar o erro de "enlarge object"
                    df_cont.iat[i, 7] = resultado_soma

        st.success("✅ Processado com sucesso!")
        st.write("Prévia (Linhas de dados):")
        st.dataframe(df_cont.iloc[3:].head(10))

        # --- 3. EXPORTAÇÃO ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Salvamos sem cabeçalho e sem índice para manter o formato original do seu arquivo
            df_cont.to_excel(writer, index=False, header=False)
        
        st.download_button(
            label="📥 Baixar Planilha CONT Atualizada",
            data=output.getvalue(),
            file_name="FOLHA_02_CONT_FINALIZADA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
