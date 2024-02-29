import streamlit as st
import os
import glob
import shutil
import tempfile
import ler_arquivos

# Page configuration
st.set_page_config(page_icon="📄", page_title="Leitor de código de barras")
st.title("📄 Leitor de código de barras")

# Function to clear temporary files
def clear():
    pycache_dir = os.path.join("./", "__pycache__")
    if os.path.exists(pycache_dir):
        shutil.rmtree(pycache_dir)
    for filename in os.listdir("./"):
        if filename.endswith(".xlsm") and filename != "codigos_de_barras_og.xlsm":
            file_path = os.path.join("./", filename)
            os.remove(file_path)
    arqs = glob.glob('./Arquivos/*')
    for arq in arqs:
        os.remove(arq)

# Clear temporary files
clear()

# Create temporary directory
temp_dir = tempfile.TemporaryDirectory()
temp_dir_path = temp_dir.name

# File upload section
uploaded_files = st.file_uploader('Inserir os arquivos:', accept_multiple_files=True)

# Execute button
if st.button('Executar'):
    for uploaded_file in uploaded_files:
        if uploaded_file is not None:
            file_path = os.path.join(temp_dir_path, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getvalue())
                
    st.success("Processando...")
    ler_arquivos.main()
    st.success('Concluído!', icon="✅")

    # Download button
    file_path = os.path.join(temp_dir_path, "codigos_de_barras.xlsm")
    if os.path.exists(file_path):
        with open(file_path, "rb") as planilha:
            btDownload = st.download_button(
                label="📥 Download",
                data=planilha.read(),
                file_name="codigos_de_barras.xlsm"
            )
    else:
        st.write("File not found.")

# Clean up temporary directory
temp_dir.cleanup()
