import pandas as pd
import streamlit as st
from datetime import datetime
import base64
import uuid
import subprocess

# Verificar si openpyxl está instalado y, si no, instalarlo
def install_openpyxl():
    try:
        import openpyxl
    except ImportError:
        subprocess.check_call(["python", '-m', 'pip', 'install', 'openpyxl'])
        import openpyxl

install_openpyxl()

# Función para generar el enlace de descarga
def download_button(object_to_download, download_filename, button_text):
    b64 = base64.b64encode(object_to_download.encode()).decode()
    button_id = str(uuid.uuid4()).replace('-', '')
    custom_css = f""" 
        <style>
            #{button_id} {{
                background-color: rgb(255, 255, 255);
                color: rgb(38, 39, 48);
                padding: 0.25em 0.38em;
                position: relative;
                text-decoration: none;
                border-radius: 4px;
                border-width: 1px;
                border-style: solid;
                border-color: rgb(230, 234, 241);
                border-image: initial;
            }} 
            #{button_id}:hover {{
                border-color: rgb(246, 51, 102);
                color: rgb(246, 51, 102);
            }}
            #{button_id}:active {{
                box-shadow: none;
                background-color: rgb(246, 51, 102);
                color: white;
                }}
        </style> """
    dl_link = custom_css + f'<a download="{download_filename}" id="{button_id}" href="data:file/txt;base64,{b64}">{button_text}</a><br></br>'
    return dl_link

# Título de la aplicación
st.title("Generador de archivo ICT para reinversiones en NSQ")

# Subir archivo Excel
uploaded_file = st.file_uploader("Carga tu archivo de Excel", type=['xlsx'])

if uploaded_file:
    # Leer el archivo Excel
    df = pd.read_excel(uploaded_file, engine='openpyxl')

    # Filtrar las columnas necesarias
    df = df[['Moneda', 'ComitenteNumero', 'Importe']]

    # Mostrar el contenido del archivo
    st.write("Contenido del archivo:")
    st.dataframe(df)

    # Obtener la fecha actual en el formato requerido
    current_date = datetime.now().strftime('%Y%m%d')

    # Función para convertir el valor de la moneda
    def convert_currency(moneda):
        if moneda == 'Dolar Renta Local - 10.000':
            return 'USD-LOCAL'
        elif moneda == 'Dolar Renta Exterior - 7.000':
            return 'USD-EXTERNO'
        else:
            return moneda

    # Preparar el contenido del archivo .ict
    ict_header = "SourceCashAccount;ReceivingCashAccount;TransactionReference;PaymentSystem;Currency;Amount;SettlementDate;Description;CorporateActionReference;TransactionOnHoldCSD;TransactionOnHoldParticipant\n"
    ict_content = []

    for index, row in df.iterrows():
        SourceCashAccount = f"46/{row['ComitenteNumero']}"
        ReceivingCashAccount = "46/1"
        TransactionReference = f"ICT{current_date}{index + 1:03d}"
        PaymentSystem = convert_currency(row['Moneda'])
        Currency = 'USD'
        Amount = row['Importe']
        SettlementDate = current_date
        Description = "Description"
        CorporateActionReference = ""
        TransactionOnHoldCSD = "0"
        TransactionOnHoldParticipant = "0"

        ict_line = f"{SourceCashAccount};{ReceivingCashAccount};{TransactionReference};{PaymentSystem};{Currency};{Amount};{SettlementDate};{Description};{CorporateActionReference};{TransactionOnHoldCSD};{TransactionOnHoldParticipant}\n"
        ict_content.append(ict_line)

    ict_file_content = ict_header + "".join(ict_content)

    # Guardar el archivo .ict
    ict_filename = 'output_file.ict'
    with open(ict_filename, 'w') as f:
        f.write(ict_file_content)

    # Mostrar el enlace de descarga
    download_button_str = download_button(ict_file_content, ict_filename, 'Descargar archivo ICT')
    st.markdown(download_button_str, unsafe_allow_html=True)
