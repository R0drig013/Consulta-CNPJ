import requests
import pandas as pd
import streamlit as st
import openpyxl
from io import BytesIO
from xlsxwriter import Workbook
import toml

#config = toml.load(".streamlit\config.toml")
#st.set_page_config(**config.get("page_config", {}))


def titulo(texto, fonte='Sigmar'):
    st.write(f'<link href="https://fonts.googleapis.com/css?family={fonte}" rel="stylesheet">', unsafe_allow_html=True)

    # Baixar a fonte
    st.write(f'<style>@import url("https://fonts.googleapis.com/css2?family={fonte}&display=swap");</style>', unsafe_allow_html=True)

    # Centralizar o texto e alterar a fonte
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


titulo('CONSULTA CNPJ')
st.text(' ')
st.text(' ')
cnpj = st.text_input('CNPJ')

if len(cnpj) != 14:
    st.info('Insira um CNPJ válido.')
else:
    dados_cnpj = consulta_cnpj(cnpj)

    dic_dados = {'Empresa': dados_cnpj['fantasia'],
                'CNPJ': dados_cnpj['cnpj'],
                'Nome': dados_cnpj['nome'],
                'Documento': limpa_doc_in_name(dados_cnpj['nome']),
                'Email': dados_cnpj['email'],
                'Cep': dados_cnpj['cep'],
                'Estado': dados_cnpj['uf'],
                'Endereço': f'{dados_cnpj["bairro"]} - {dados_cnpj["logradouro"]} - {dados_cnpj["numero"]}' ,
                'Cidade': dados_cnpj['municipio'],
                'situação': dados_cnpj['situacao']}

    df_cnpj = pd.DataFrame(dic_dados, index=[0])

    # Adicione um botão para baixar o arquivo Excel
    def download_excel(df):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    
    processed_data = download_excel(df_cnpj)
    st.download_button(
        label="Baixar arquivo Excel",
        data=processed_data,
        file_name="consultaCNPJ.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Exiba o DataFrame
    st.table(df_cnpj)

                    
