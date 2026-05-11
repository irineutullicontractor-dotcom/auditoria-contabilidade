import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Processador Contábil Final", layout="wide")

def limpar_e_converter(valor):
    """Converte valores monetários do formato BR para float."""
    if pd.isna(valor) or valor == '': return 0.0
    # Remove pontos de milhar e troca vírgula por ponto
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
        # Lemos forçando string para evitar erro de tipo no Evento
        df_adm = pd.read_excel(file_adm, skiprows=1, dtype=str)
        
        mapa_eventos = {}
        for _, row in df_adm.iterrows():
            # Percorre os dois blocos de colunas (A/D e F/I)
            # Índices: 0=Evento(A), 3=Valor(D), 5=Evento(F), 8=Valor(I)
            indices = [(0, 3), (5, 8)]
            for ev_idx, val_idx in indices:
                try:
                    cod_bruto = str(row.iloc[ev_idx]).split('.')[0].strip()
                    if cod_bruto.isdigit():
                        cod_int = int(cod_bruto)
                        valor_num = limpar_e_converter(row.iloc[val_idx])
                        mapa_eventos[cod_int] = mapa_eventos.get(cod_int, 0) + valor_num
                except:
                    continue

        # --- 2. PROCESSAMENTO DA PLANILHA CONT (Aba ADMIN) ---
        # Lemos a aba específica forçando dtype=str para evitar o erro relatado
        df_cont = pd.read_excel(file_cont, sheet_name='ADMIN', header=None, dtype=str)

        # Garante que o DataFrame tenha largura suficiente (até coluna H / índice 7)
        while df_cont.shape[1] < 8:
            df_cont[df_cont.shape[1]] = ""

        def calcular_soma_celula(texto):
            if pd.isna(texto) or texto == '': return 0.0
            # Extrai todos os números (códigos de evento) da célula
            codigos_encontrados = re.findall(r'\d+', str(texto))
            return sum(mapa_eventos.get(int(c), 0.0) for c in codigos_encontrados)

        # Processamento linha a linha a partir da linha 4 (índice 3)
        # Coluna B (índice 1) -> Origem dos códigos
        # Coluna H (índice 7) -> Destino da soma
        for i in range(len(df_cont)):
            if i >= 3: # Linha 4 em diante
                conteudo_b = df_cont.iloc[i, 1]
                if pd.notna(conteudo_b) and conteudo_b != '':
                    total_soma = calcular_soma_celula(conteudo_b)
                    # Grava o resultado na coluna H (índice 7)
                    df_cont.iat[i, 7] = total_soma

        st.success("✅ Processamento da aba ADMIN concluído com sucesso!")
        st.write("Visualização das primeiras linhas processadas:")
        st.dataframe(df_cont.iloc[3:].head(15))

        # --- 3. EXPORTAÇÃO ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Salva na aba ADMIN sem cabeçalhos do pandas para manter o original
            df_cont.to_excel(writer, index=False, header=False, sheet_name='ADMIN')
        
        st.download_button(
            label="📥 Baixar Planilha Pronta",
            data=output.getvalue(),
            file_name="FOLHA_02_CONT_ATUALIZADA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
        st.info("Certifique-se de que a aba da segunda planilha se chama exatamente 'ADMIN'.")
