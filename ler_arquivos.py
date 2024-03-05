from pyzbar import pyzbar
from openpyxl import load_workbook
import cv2
import time
import os
from glob import glob
from pdf2image import convert_from_path
import streamlit as st

local = "./"
caminhoPlanilha = os.path.join(local, "Planilha de lançamentos - v2.0 - OG.xlsm")
arquivo_excel = load_workbook(caminhoPlanilha, read_only=False, keep_vba=True)
planilha = arquivo_excel['CÓDIGOS']

def converterPdf():
    arquivos = glob(os.path.join(local, "Arquivos", "*.pdf"))
    for arquivo in arquivos:
        try:
            convertido = convert_from_path(arquivo, first_page=1, last_page=1)
            for i, image in enumerate(convertido):
                nome = str(arquivo).replace(".pdf", "") + ".png"
                st.write(nome)
                image.save(nome, "PNG")
        except Exception as e:
            st.write(str(arquivo) + " - ERRO:", e)

def detectar(imagem,i):
    zoom = 1.0
    zoom_step = 0.1
    while zoom <= 2.0:
        h, w = imagem.shape[:2]
        zoomed = cv2.resize(imagem, (int(w*zoom), int(h*zoom)))
        codigos = pyzbar.decode(zoomed)
        for codigo in codigos:
            if codigo.type == "I25" and len(str(codigo.data)) == 47:
                st.write("Tipo:", codigo.type)
                st.write("Código:", codigo.data)
                planilha.cell(row=2+i, column=3, value=codigo.data)
                planilha.cell(row=2+i, column=2, value=codigo.type)
                return codigo.data
        zoom += zoom_step
        time.sleep(0.1)
    st.write("Não encontrado.")
    return None

def limparPlanilha(sheet, min_row, min_col, max_col):
    for row in sheet.iter_rows(min_row=min_row, min_col=min_col, max_col=max_col):
        if row[0].value is None:
            break
        for cell in row:
            cell.value = None

def main():
    limparPlanilha(planilha, min_row=2, min_col=1, max_col=3)
    converterPdf()

    pastaArquivos = os.path.join(local, "Arquivos")
    arquivos = glob(os.path.join(local, "Arquivos", "*.png"))
    i = 0
    for arquivo in arquivos:
        nome = os.path.basename(arquivo).replace(".png", "")
        st.write("Nome:", nome)
        planilha.cell(row=2+i, column=1, value=nome)
        try:
            imgBoleto = cv2.imread(arquivo)
            imgBoleto = detectar(imgBoleto,i)
        except Exception as e:
            planilha.cell(row=2+i, column=3, value="Erro na leitura do PDF")
            st.write("Erro na leitura do PDF:", e)
        i += 1

    arquivo_excel.save("./Planilha de lançamentos - v2.0.xlsm")

if __name__ == "__main__":
    main()
