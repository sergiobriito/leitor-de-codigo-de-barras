from pyzbar import pyzbar
from openpyxl import load_workbook
import cv2
import time
import os
from glob import glob
from pdf2image import convert_from_path
import streamlit as st

local = "./"
caminhoPlanilha = os.path.join(local, "codigos_de_barras_og.xlsm")
arquivo_excel = load_workbook(caminhoPlanilha, read_only=False, keep_vba=True)
planilha = arquivo_excel['CÓDIGOS']

# Função para converter PDF em PNG
def converterPdf():
    arquivos = glob(os.path.join(local, "Arquivos", "*.pdf"))
    for arquivo in arquivos:
        try:
            print(arquivo)
            convertido = convert_from_path(arquivo, first_page=1, last_page=1)
            for i, image in enumerate(convertido):
                nome = str(arquivo).replace(".pdf", "") + ".png"
                st.write(nome)
                image.save(nome, "PNG")
        except Exception as e:
            st.write(str(arquivo) + " - ERRO:", e)

# Função para detectar código em imagem(PNG).
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

def main():
    for col in planilha.iter_cols(min_row=2, max_row=999999, min_col=1, max_col=3):
        for cell in col:
            cell.value = None

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

    arquivo_excel.save("./codigos_de_barras.xlsm")

if __name__ == "__main__":
    main()
