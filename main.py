from docxtpl import DocxTemplate, InlineImage, RichText
import requests
import streamlit as st
from datetime import date
import io


#docx
# documento = Document("Contrato-Contabilidade.docx")

# for paragrafo in documento.paragraphs:
#     paragrafo.text = paragrafo.text.replace("RAZÃO SOCIAL DO SEU CLIENTE ", "CLUBFIX REPARACAO E MANUTENCAO DE EQUIPAMENTOS ELETRONICOS LTDA")

# documento.save("Contrato-Contabilidade-Editado.docx")

#docxtpl
global url 

def consultarApi(cnpj):
    url = f'https://publica.cnpj.ws/cnpj/{cnpj}'
    
    response = requests.get(url)

    if response.status_code == 200:
        return response.json()
    else:
        return None


def manipularArquivo(dados, estado_civil, rg, cpf,nome_arquivo, inicio_contrato):

    modelo = DocxTemplate('Contrato-Contabilidade - Modelo.docx')
    estabelecimento = dados['estabelecimento']

    parametros = {
    'nome_empresa' : dados['razao_social'],
    'endereço': f'{estabelecimento["tipo_logradouro"]} {estabelecimento["logradouro"]}, {estabelecimento["numero"]}, Bairro: {estabelecimento["bairro"]}, CEP: {estabelecimento["cep"]}',
    'cnpj' : estabelecimento['cnpj'],
    'representante' : dados['socios'][0]['nome'],
    'nacionalidade' : dados['socios'][0]['pais']['nome'],
    'cargo': dados['socios'][0]['qualificacao_socio']['descricao'],
    'estado_civil': estado_civil,
    'rg' : rg,
    'cpf' : cpf,
    'inicio_contrato': inicio_contrato    
    }

    modelo.render(parametros)
    saida = f'{nome_arquivo}.docx'
    try:
        #modelo.save(saida)
        # word = win32com.client.Dispatch("Word.Application")
        # word.Documents.Open(saida)
        # word.Visible  = True

        bio = io.BytesIO()
        modelo.save(bio)
        print(bio)
        if modelo:
            global url 
            url = bio.getvalue()
            
        print('Fim.')
    except:
        print('Nao foi possível salvar o relatório; verifique se o arquivo não está aberto.')
        # fim try

def ui():
    st.markdown('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">', unsafe_allow_html=True)

    # st.title('Gerador de Contratos')
    cnpj = st.text_input('CNPJ da Empresa','')
    estado_civil = st.selectbox(
    "Estado Civil?",
    ("Solteiro", "Casado", "Separado", "Divorciado", "Viúvo(a)"),
    index=None,
    placeholder="Selecione um estado civil",
)
    rg = st.text_input('RG','')
    cpf = st.text_input('CPF','')
    nome_arquivo = st.text_input('Nome do Arquivo','')
    inicio_contrato = st.date_input("Data de inicio de Contrato", date.today(), format="DD/MM/YYYY")



    if st.button('Gerar Contrato'):
        dados = consultarApi(cnpj)
        
        if dados:
            print(dados)
            # estabelecimento = dados['estabelecimento']

            # parametros = {
            # 'nome_empresa' : dados['razao_social'],
            # 'endereço': f'{estabelecimento['tipo_logradouro']} {estabelecimento["logradouro"]}, {estabelecimento["numero"]}, Bairro: {estabelecimento["bairro"]}, CEP: {estabelecimento["cep"]}',
            # 'cnpj' : estabelecimento['cnpj'],
            # 'representante' : dados['socios'][0]['nome'],
            # 'nacionalidade' : dados['socios'][0]['pais']['nome'],
            # 'cargo': dados['socios'][0]['qualificacao_socio']['descricao'],
            # }
            #st.write(parametros)
            manipularArquivo(dados, estado_civil, rg, cpf, nome_arquivo, inicio_contrato)
            
            st.download_button(
                label="Clique aqui para fazer o download",
                data=url,
                file_name=f'{nome_arquivo}.docx',
                mime="docx"
            )
            
        else:
            st.write('CNPJ inválido')

ui()