import streamlit as st
import os
import glob
import ler_arquivos

st.set_page_config(page_icon="📄", page_title="Leitor de código de barras")
st.title("📄 Leitor de código de barras")

uploaded_files = st.file_uploader('Inserir os arquivos:', accept_multiple_files=True)

#if os.path.exists("codigos_de_barras.xlsm"):
#    os.remove("codigos_de_barras.xlsm")

arqs = glob.glob('./Arquivos/*')
for arq in arqs:
    os.remove(arq)

if st.button('Executar'):
    for uploaded_file in uploaded_files:
        if uploaded_file is not None:
            file_path = os.path.join("Arquivos", uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getvalue())
                
    st.success("Processando...")
    ler_arquivos.main()
    st.success('Concluído!', icon="✅")

    with open("codigos_de_barras.xlsm","rb") as planilha:
        btDownload = st.download_button(
            label ="📥 Download",
            data = planilha,
            file_name="codigos_de_barras.xlsm"
        )

        if btDownload:
            st.experimental_rerun()

style = """
<style>
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
.css-12oz5g7 {padding: 2rem 1rem;}
.css-14xtw13 {visibility: hidden;}
span.css-9ycgxx.exg6vvm12 {
visibility: hidden;
white-space: nowrap;
}
span.css-9ycgxx.exg6vvm12::before {
    visibility: visible;
    content: "Selecionar os arquivos, ou arraste e solte os arquivos aqui";
    font-size: 1rem;
    font-family: "Source Sans Pro", sans-serif;
    font-weight: 400;
    line-height: 1.6;
    text-size-adjust: 100%;
    -webkit-tap-highlight-color: rgba(0, 0, 0, 0);
    -webkit-font-smoothing: auto;
    color: rgb(49, 51, 63);
    box-sizing: border-box;
    margin-bottom: 0.25rem;
}

small.css-1aehpvj.euu6i2w0 {visibility: hidden;}
small.css-1aehpvj.euu6i2w0::before {
    visibility: visible;
    content: "200MB por arquivo";
    font-family: "Source Sans Pro", sans-serif;
    font-weight: 400;
    text-size-adjust: 100%;
    -webkit-tap-highlight-color: rgba(0, 0, 0, 0);
    -webkit-font-smoothing: auto;
    box-sizing: border-box;
    color: rgba(49, 51, 63, 0.6);
    font-size: 14px;
    line-height: 1.25;
}

section.css-po3vlj.exg6vvm15 button{visibility:hidden;}

#Linkedin {margin-top: 175px;}
#desenvolvidoPor {color: black;}
#nome {color: black;}
</style>

<div id="Linkedin" class="badge-base LI-profile-badge" data-locale="pt_BR" data-size="medium" data-theme="light" data-type="VERTICAL" data-vanity="sérgio--brito" data-version="v1">
<a href="https://br.linkedin.com/in/s%C3%A9rgio--brito?trk=profile-badge"><img src="https://brand.linkedin.com/content/dam/me/business/en-us/amp/brand-site/v2/bg/LI-Bug.svg.original.svg" alt="Linkedin" style="width:42px;height:42px;"></a>
<a id="desenvolvidoPor">Desenvolvido por </a>
<a id="nome" class="badge-base__link LI-simple-link" href="https://br.linkedin.com/in/s%C3%A9rgio--brito?trk=profile-badge">Sérgio Brito</a>
</div>

"""

st.markdown(style, unsafe_allow_html=True)
