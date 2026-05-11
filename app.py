import streamlit as st
from openpyxl import load_workbook
import re
from io import BytesIO

st.set_page_config(
    page_title="Integração Folha ADM x CONT",
    layout="wide"
)

st.title("Integração Folha ADM x CONT")

st.write("""
Este sistema:

1. Lê os códigos de eventos da planilha ADM
2. Soma os valores dos eventos repetidos
3. Procura os códigos na planilha CONT
4. Soma os respectivos valores
5. Preenche a coluna H (VLR LANÇAMENTO)
""")

# =========================
# UPLOADS
# =========================
arquivo_adm = st.file_uploader(
    "Selecione a planilha Folha - ADM",
    type=["xlsx"]
)

arquivo_cont = st.file_uploader(
    "Selecione a planilha FOLHA - CONT",
    type=["xlsx"]
)

# =========================
# PROCESSAMENTO
# =========================
if arquivo_adm and arquivo_cont:

    with st.spinner("Processando arquivos..."):

        # =========================
        # LEITURA DA PLANILHA ADM
        # =========================
        wb_adm = load_workbook(
            arquivo_adm,
            data_only=True
        )

        ws_adm = wb_adm[wb_adm.sheetnames[0]]

        eventos = {}

        # Percorre linhas
        for row in range(3, ws_adm.max_row + 1):

            # ====================================
            # BLOCO 1
            # EVENTO = COLUNA A
            # VALOR  = COLUNA D
            # ====================================
            evento_1 = ws_adm[f"A{row}"].value
            valor_1 = ws_adm[f"D{row}"].value

            if evento_1 not in [None, ""]:

                try:
                    codigo = int(float(evento_1))
                    valor = float(valor_1 or 0)

                    eventos[codigo] = (
                        eventos.get(codigo, 0) + valor
                    )

                except:
                    pass

            # ====================================
            # BLOCO 2
            # EVENTO = COLUNA F
            # VALOR  = COLUNA I
            # ====================================
            evento_2 = ws_adm[f"F{row}"].value
            valor_2 = ws_adm[f"I{row}"].value

            if evento_2 not in [None, ""]:

                try:
                    codigo = int(float(evento_2))
                    valor = float(valor_2 or 0)

                    eventos[codigo] = (
                        eventos.get(codigo, 0) + valor
                    )

                except:
                    pass

        # =========================
        # LEITURA DA PLANILHA CONT
        # =========================
        wb_cont = load_workbook(arquivo_cont)

        ws_cont = wb_cont["ADMIN"]

        # Regex para encontrar códigos numéricos
        regex_codigos = re.compile(r"\d+")

        # =========================
        # PREENCHIMENTO DA COLUNA H
        # =========================
        for row in range(5, ws_cont.max_row + 1):

            texto_conta = ws_cont[f"B{row}"].value

            if texto_conta:

                # Extrai todos os números da descrição
                codigos = regex_codigos.findall(
                    str(texto_conta)
                )

                soma = 0

                for codigo in codigos:

                    codigo_int = int(codigo)

                    soma += eventos.get(
                        codigo_int,
                        0
                    )

                # Grava resultado na coluna H
                celula = ws_cont[f"H{row}"]

                celula.value = soma

                # Formatação monetária
                celula.number_format = '#,##0.00'

        # =========================
        # GERAR DOWNLOAD
        # =========================
        output = BytesIO()

        wb_cont.save(output)

        output.seek(0)

        st.success("Arquivo processado com sucesso!")

        st.download_button(
            label="📥 Baixar Arquivo Preenchido",
            data=output,
            file_name="FOLHA_02_CONT_PREENCHIDA.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
