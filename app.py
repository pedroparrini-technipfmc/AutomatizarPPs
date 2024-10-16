import os
import PyPDF2
import comtypes.client
import streamlit as st
from io import BytesIO
import pythoncom

# ESTA FUNÇÃO PRECISA DE INTERNET PRA FUNCIONAR

# AVALIAR A POSIBLIDADE DE FAZER ESTA FUNÇÃO COM A BIBLIOTECA DOCX

def convert_docx_to_pdf(input_file):

    output_file = input_file[:-5] + ".pdf"
    word = comtypes.client.CreateObject("Word.Application")
    docx_path = os.path.abspath(input_file)
    pdf_path = os.path.abspath(output_file)
    pdf_format = 17
    word.Visible = False
    in_file = word.Documents.Open(docx_path)
    in_file.Saveas(pdf_path, FileFormat=pdf_format)
    in_file.Close()
    word.Quit()

def merge_docs(arquivo_docx, pdf_desenho, pagina_inicial, pagina_final):

    # Recebe 1 arquivo .docx, 1 arquivo PDF, o número da página inicial e final dos desenhos

    # arquivo_docx -> caminho do pp na pasta
    # pdf_desenho -> caminho do pdf do pp na pasta final
    # página inicial -> número da página que inicia a sequência de desenhos
    # página final -> número da página que termina a sequeênia de desenhos

    # Converter o arquivo PP .docx em .pdf
    convert_docx_to_pdf(arquivo_docx)
    arquivo_pdf = arquivo_docx[:-5] + ".pdf"

    # Lista conm os índices das páginas que serão substuídas
    lista_delete = list(range(pagina_inicial - 1, pagina_final + 1 - 1))

        # Abrir o arquivo PDF original e o PDF de desenhos
    with open(arquivo_pdf, 'rb') as pdf_original_file, open(pdf_desenho, 'rb') as pdf_desenhos_file:
        # Ler o conteúdo dos arquivos PDF
        leitor_pdf_original = PyPDF2.PdfReader(pdf_original_file)
        leitor_pdf_desenhos = PyPDF2.PdfReader(pdf_desenhos_file)

        escritor_pdf = PyPDF2.PdfWriter()

        num_paginas_original = len(leitor_pdf_original.pages)
        num_paginas_desenhos = len(leitor_pdf_desenhos.pages)

        # Verificar se os índices estão dentro do limite do arquivo original
        for idx in lista_delete:
            if idx < 0 or idx >= num_paginas_original:
                raise IndexError(f"O índice {idx} está fora do intervalo permitido para o PDF original.")

        # Realizar o processo de substituição
        for i in range(num_paginas_original):
            if i in lista_delete:
                # Substituir a página pela página correspondente no PDF de desenhos
                pagina_index = lista_delete.index(i)
                if pagina_index < num_paginas_desenhos:
                    pagina_desenho = leitor_pdf_desenhos.pages[pagina_index]
                    escritor_pdf.add_page(pagina_desenho)
                else:
                    raise IndexError(f"O PDF de desenhos não possui a página solicitada no índice {pagina_index}.")
            else:
                # Adicionar a página original
                pagina_original = leitor_pdf_original.pages[i]
                escritor_pdf.add_page(pagina_original)

        # Escrever o resultado em um novo arquivo PDF
        with open(arquivo_pdf, 'wb') as output_pdf_file:
            escritor_pdf.write(output_pdf_file)

    print(f"PDF gerado com sucesso: {arquivo_pdf}")

# Configurando a barra lateral com as opções
st.sidebar.title("Menu")
opcao = st.sidebar.selectbox("Escolha uma funcionalidade", ["Início da revisão", "Fim da revisão"])

# Se a opção for "Início da revisão"
if opcao == "Início da revisão":
    st.title("Início da revisão")
    st.write("Selecione uma funcionalidade na barra lateral para começar.")

# Se a opção for "Fim da revisão"
elif opcao == "Fim da revisão":
    st.title("Fim da revisão")

    # Upload de arquivos
    st.write("Por favor, faça o upload dos arquivos para serem modificados.")
    arquivo_docx = st.file_uploader("Carregar arquivo .docx", type="docx")
    pdf_desenho = st.file_uploader("Carregar arquivo .pdf", type="pdf")
    pagina_inicial = st.text_input("Página inicial dos desenhos")
    pagina_final = st.text_input("Página final dos desenhos")

    if arquivo_docx and pdf_desenho and pagina_inicial and pagina_final:

        arquivo_docx_path = os.path.join("uploads", arquivo_docx.name)
        with open(arquivo_docx_path, "wb") as f:
            f.write(arquivo_docx.getbuffer())

        pdf_desenho_path = os.path.join("uploads", pdf_desenho.name)
        with open(pdf_desenho_path, "wb") as f:
            f.write(pdf_desenho.getbuffer())

        pythoncom.CoInitialize()

        merge_docs(arquivo_docx_path, pdf_desenho_path, int(pagina_inicial), int(pagina_final))

        pdf_final = arquivo_docx_path[:-5] + ".pdf"

        with open(pdf_final, "rb") as file:
            st.download_button(
                label="Baixar PDF Modificado",
                data=file,
                file_name=os.path.basename(pdf_final),
                mime="application/pdf"
            )
