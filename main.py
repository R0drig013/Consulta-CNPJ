import requests
import pandas as pd
import streamlit as st
import openpyxl
from io import BytesIO
from xlsxwriter import Workbook
import toml
from time import sleep

#config = toml.load(".streamlit\config.toml")
#st.set_page_config(**config.get("page_config", {}))


def titulo(texto, fonte='Sigmar'):
    st.write(f'<link href="https://fonts.googleapis.com/css?family={fonte}" rel="stylesheet">', unsafe_allow_html=True)

    st.write(f'<style>@import url("https://fonts.googleapis.com/css2?family={fonte}&display=swap");</style>', unsafe_allow_html=True)

    st.write(f"""
    <div style='display:flex; justify-content:center; align-items:center; font-family:"{fonte}", cursive;'>
        <h1>{texto}</h1>
    </div>
    """, unsafe_allow_html=True)


def limpa_doc_in_name(texto):
    number_doc = [x for x in texto if x.isdigit()]

    if len(number_doc) > 0:
        soma_string = ''
        for digit in number_doc:
            soma_string += digit

        return soma_string
    else:
        return texto


def consulta_cnpj(cnpj):
    url =  f'https://receitaws.com.br/v1/cnpj/{cnpj}'

    headers = {"Accept": "application/json"}

    response = requests.get(url, headers=headers)

    return response.json()


def download_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


titulo('CONSULTA CNPJ')
st.text(' ')
st.text(' ')
uploader_cnpj_file = st.file_uploader("Arquivo Excel", type=["xls", "xlsx"])

consulta_button = st.button('CONSULTAR')

if consulta_button:

    dados_to_consulta = []
    if uploader_cnpj_file is not None:
        cnpj_dataframe = pd.read_excel(uploader_cnpj_file, dtype=str)
        
        cnpjs = [x for x in list(cnpj_dataframe.iloc[:, 0]) if len(x) == 14]  
        
        st.text(' ')
        st.text(' ')
        st.text(' ')
        st.write('---')
        barra_progress = st.progress(0, text='Consultando CNPJ')
        for index, cnpj_text in enumerate(cnpjs):
            dados_cnpj = consulta_cnpj(cnpj_text)

            sleep(22)
            dados_to_consulta.append([dados_cnpj['fantasia'], 
                                      dados_cnpj['cnpj'], 
                                      dados_cnpj['nome'], 
                                      limpa_doc_in_name(dados_cnpj['nome']),
                                      dados_cnpj['email'],
                                      dados_cnpj['cep'],
                                      dados_cnpj['uf'],
                                      dados_cnpj["bairro"],
                                      dados_cnpj["logradouro"],
                                      dados_cnpj["numero"],
                                      dados_cnpj['municipio'],
                                      dados_cnpj['situacao']])
            
           
            barra_progress.progress((index + 1)  / len(cnpjs), text='Consultando CNPJ')
        barra_progress.progress((index + 1)  / len(cnpjs), text='Consulta Finalizada.')

        dic_dados_for_df = {'Empresa': [x[0] for x in dados_to_consulta],
                'CNPJ': [x[1] for x in dados_to_consulta],
                'Nome': [x[2] for x in dados_to_consulta],
                'Documento': [x[3] for x in dados_to_consulta],
                'Email': [x[4] for x in dados_to_consulta],
                'Cep': [x[5] for x in dados_to_consulta],
                'Estado': [x[6] for x in dados_to_consulta],
                'Bairro': [x[7] for x in dados_to_consulta],
                'Rua': [x[8] for x in dados_to_consulta],
                'Número': [x[9] for x in dados_to_consulta],
                'Cidade': [x[10] for x in dados_to_consulta],
                'Situação': [x[11] for x in dados_to_consulta]}
        
        df_cnpj = pd.DataFrame(dic_dados_for_df)

        # Adicione um botão para baixar o arquivo Excel
               
        processed_data = download_excel(df_cnpj)
        st.download_button(
            label="Baixar arquivo Excel",
            data=processed_data,
            file_name="consultaCNPJ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.error('Você não inseriu um arquivo excel.')
    

                    
