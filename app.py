
import streamlit as st
from openpyxl import load_workbook
import re
from io import BytesIO

st.set_page_config(page_title="Integração Folha", layout="wide")

st.title("Integração Folha ADM x CONT")

arquivo_adm = st.file_uploader(
    "Selecione o arquivo Folha ADM",
    type=["xlsx"]
)

arquivo_cont = st.file_uploader(
    "Selecione o arquivo FOLHA CONT",
    type=["xlsx"]
)

if arquivo_adm and arquivo_cont:

    # =========================
    # LEITURA DOS EVENTOS
    # =========================
    wb_adm = load_workbook(arquivo_adm, data_only=True)
    ws_adm = wb_adm[wb_adm.sheetnames[0]]

    eventos = {}

    for row in range(3, ws_adm.max_row + 1):

        # A/D
        evento_1 = ws_adm[f"A{row}"].value
        valor_1 = ws_adm[f"D{row}"].value

        if evento_1 is not None:
            try:
                codigo = int(evento_1)
                valor = float(valor_1 or 0)
                eventos[codigo] = eventos.get(codigo, 0) + valor
            except:
                pass

        # F/I
        evento_2 = ws_adm[f"F{row}"].value
        valor_2 = ws_adm[f"I{row}"].value

        if evento_2 is not None:
            try:
                codigo = int(evento_2)
                valor = float(valor_2 or 0)
                eventos[codigo] = eventos.get(codigo, 0) + valor
            except:
                pass

    # =========================
    # PREENCHER CONT
    # =========================
    wb_cont = load_workbook(arquivo_cont)
    ws_cont = wb_cont["ADMIN"]

    regex_codigos = re.compile(r"\\b\\d+\\b")

    for row in range(5, ws_cont.max_row + 1):

        texto_conta = ws_cont[f"B{row}"].value

        if texto_conta:

            codigos = regex_codigos.findall(str(texto_conta))

            soma = 0

            for codigo in codigos:
                codigo_int = int(codigo)
                soma += eventos.get(codigo_int, 0)

            ws_cont[f"H{row}"] = soma

    # =========================
    # DOWNLOAD
    # =========================
    output = BytesIO()
    wb_cont.save(output)
    output.seek(0)

    st.success("Arquivo processado com sucesso!")

    st.download_button(
        label="Baixar arquivo preenchido",
        data=output,
        file_name="FOLHA_02_CONT_PREENCHIDA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
