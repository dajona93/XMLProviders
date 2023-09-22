import json
import imaplib
import email
import re
import requests
from datetime import datetime
from lxml import etree
import os
from datetime import datetime

# Definir los namespaces necesarios para XPath
namespaces = {
    'cbc': 'urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2',
    'cac': 'urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2',
}

# Configura la conexión IMAP para Outlook
email_user = 'dajona001@hotmail.com'
email_pass = 'e71523'
mail = imaplib.IMAP4_SSL('outlook.office365.com', 993)  # Utiliza el puerto 993 para IMAP con cifrado TLS
mail.login(email_user, email_pass)

# Selecciona la carpeta 'XML TEST' en lugar de 'INBOX'
mail.select('XML')

# Obtiene la fecha actual en el formato requerido por IMAP (DD-MMM-YYYY)
today_date = datetime.now().strftime('%d-%b-%Y')

# Construye la consulta IMAP para buscar correos con la fecha de hoy en la carpeta 'XML TEST'
search_query = f'SINCE "{today_date}"'

# Busca correos con la fecha de hoy en la carpeta 'XML TESTp'
result, data = mail.uid('search', None, search_query)

# Obtiene una lista de los IDs de los correos encontrados
email_ids = data[0].split()

# Lista para almacenar los diccionarios JSON de cada correo
json_data_list = []

# Procesamiento de correos y archivos adjuntos
for email_id in email_ids:
    result, email_data = mail.uid('fetch', email_id, '(RFC822)')
    raw_email = email_data[0][1]
    email_message = email.message_from_bytes(raw_email)

    # Obtiene la fecha actual en formato "DD/MM/YYYY"
    current_date = datetime.now().strftime('%d/%m/%Y')
    
    # Inicializa un nuevo diccionario JSON para cada XML adjunto
    json_data = {
        "CODDOC": "CPE",  # Llenar con los datos del XML
        "CODANE": "",  # Llenar con los datos del XML
        "COMMEM": "",  # Llenar con los datos del XML
        "SERREA": "",
        "NUMREA": "",
        "FECDOC": current_date,  # Llenar con los datos del XML
        "FECREA": "",
        "XCONPAG": "EFEC",
        "XTIPMOV": "BI",
        "CODRES": "R007",
        "xTipMon": "MN",
        "datAnexo": {
            "xTipAne": "PRO",
            "CodAne": "",
            "NomAne": "",
            "NomTra1": "",
            "NomTra2": "",
            "ApellPat": "",
            "ApellMat": "",
            "IdeAne1": "",
            "XSUBANE02": "05",
            "XSUBANE03": "34",
            "XSUBANE04": "06",
            "XSUBANE05": "SD03",
            "DirAne": "",
            "xTipIde1": "",
            "MailAne": "",
            "TelAne": "",
            "xUbigeo": "051"
        },
        "lstDetalle": [
            {
                "NUMITE": "",
                "CODSUBALM": "AMCP",
                "CODART": "CS00001",
                "DESART": "",
                "XTIPUNI": "UND",
                "CANTOT": "",
                "TOTART": "40",
                "TipDesAdq": "1",
                "V01": "G",
                "Imp001": "18"
            }
        ]
    }
    
    # Itera sobre los archivos adjuntos
    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        
        # Guarda el archivo adjunto en la carpeta local si es un archivo XML
        filename = part.get_filename()
        if filename and filename.endswith('.xml'):
            xml_content = part.get_payload(decode=True)
            
            # Procesa el XML como lo hiciste anteriormente
            tree = etree.fromstring(xml_content)
            
            # Verifica si la serie de factura comienza con "F"
            cbc_id_element = tree.find('./cbc:ID', namespaces=namespaces)
            if cbc_id_element is not None:
                series = cbc_id_element.text
                if not series.startswith('F'):
                    # Omite el procesamiento si la serie no comienza con "F"
                    continue

                # Procesa el XML solo si la serie comienza con "F"
                # Llena los campos del JSON con los datos del XML
                json_data['COMMEM'] = series
                issue_date = tree.find('.//cbc:IssueDate', namespaces=namespaces).text
                
                 # Cambia el formato de '2023-08-29' a '29/08/2023'
                parts = issue_date.split('-')
                new_format_issue_date = f'{parts[2]}/{parts[1]}/{parts[0]}'
                json_data['FECREA'] = new_format_issue_date

            # Busca el elemento cbc:ID que está dentro de un contexto específico, por ejemplo, dentro de cbc:PartyIdentification
            party_id_element = tree.find('.//cac:PartyIdentification/cbc:ID', namespaces=namespaces)

            if party_id_element is not None:
            # Verifica que sea el número de RUC (Puedes agregar una validación adicional aquí)
                json_data['CODANE'] = party_id_element.text
                
                # Agrega la letra "P" al final del número de RUC
                json_data['datAnexo']['CodAne'] = party_id_element.text + "P"

                # Agrega el número de RUC sin la letra "P" en el campo "IdeAne1"
                json_data['datAnexo']['IdeAne1'] = party_id_element.text

            # Llena la información de datAnexo si está presente en el XML
            dat_anexo_element = tree.find('.//cac:PartyLegalEntity', namespaces=namespaces)
            if dat_anexo_element is not None:
                # Llena los campos dentro de datAnexo con los datos del XML
                json_data['datAnexo']['NomAne'] = dat_anexo_element.find('.//cbc:RegistrationName', namespaces=namespaces).text
                
            # Llena la información de datAnexo si está presente en el XML
            dat_anexo_element = tree.find('.//cbc:Line', namespaces=namespaces)
            if dat_anexo_element is not None:
                # Llena el campo DirAne con el valor del XML
                json_data['datAnexo']['DirAne'] = dat_anexo_element.text
            
            # Divide "COMMEM" en "SERREA" y "NUMREA"
            serrea, numrea = json_data['COMMEM'].split('-')
            json_data['SERREA'] = serrea
            json_data['NUMREA'] = numrea.lstrip('0')

            # Busca el ID del Invoice Line en el elemento adecuado de tu XML
            invoice_line_id_element = tree.find('.//cac:InvoiceLine/cbc:ID', namespaces=namespaces)

            if invoice_line_id_element is not None:
                # Agrega el ID del Invoice Line al diccionario lstDetalle en el JSON
                json_data['lstDetalle'][0]['NUMITE'] = invoice_line_id_element.text

            # Busca el InvoiceCuantity del Invoice Line en el elemento adecuado de tu XML
            invoice_line_id_element = tree.find('.//cac:InvoiceLine/cbc:InvoicedQuantity', namespaces=namespaces)

            if invoice_line_id_element is not None:
                # Agrega el ID del Invoice Line al diccionario lstDetalle en el JSON
                json_data['lstDetalle'][0]['CANTOT'] = invoice_line_id_element.text    
            
            # Busca el ID del Invoice Line en el elemento adecuado de tu XML
            invoice_line_id_element = tree.find('.//cac:SellersItemIdentification/cbc:ID', namespaces=namespaces)

            # Llena la información de datAnexo si está presente en el XML
            dat_anexo_element = tree.find('.//cac:Item', namespaces=namespaces)
            if dat_anexo_element is not None:
                # Llena los campos dentro de datAnexo con los datos del XML
                json_data['lstDetalle'][0]['DESART'] = dat_anexo_element.find('.//cbc:Description', namespaces=namespaces).text

            # Busca el ID del Invoice Line en el elemento adecuado de tu XML
            invoice_line_id_element = tree.find('.//cac:LegalMonetaryTotal/cbc:ID', namespaces=namespaces)

            # Verifica si al menos un campo relevante tiene un valor antes de imprimir
            if any(json_data.values()):
                # Convierte el diccionario JSON a formato JSON
                final_json_data = json.dumps(json_data, indent=4)

            json_data_list.append(json_data)
    
# Cierra la conexión IMAP
mail.logout()


# Agrega el código para enviar los JSON a la API aquí
# URL de la API a la que deseas enviar los JSON
api_url = 'http://puonline.upeu.edu.pe:8186/ws/SIDIGE.Web.Service.DocMntRest.svc/api/docmntrest/generardocumento'  # Reemplaza con la URL de tu API

# Cabeceras para indicar que estás enviando datos JSON
headers = {'codemp': 'E1',
           'codgru': 'G1',
           'uid'   : 'admin',
           'pwd'   : 'SDG2023PU',
           'ten'   : '0',
           'opcion': '24600',
           'can'   : '1',
           'cmpane': '1'
           }

# Itera sobre los JSON en json_data_list y envíalos a la API
for json_data in json_data_list:
    # Convierte el diccionario JSON a una cadena JSON
    json_string = json.dumps(json_data)

    # Envía la solicitud POST a la API
    response = requests.post(api_url, data=json_string, headers=headers)

    # Imprime el código de estado HTTP
    print(f'Código de estado HTTP: {response.status_code}')

    # Imprime el contenido de la respuesta
    print(f'Respuesta de la API: {response.text}')

    # Verifica la respuesta de la API
    if response.status_code == 200:
        print(f'JSON enviado exitosamente a la API: {response.json()}')
    else:
        print(f'Error al enviar JSON a la API. Código de estado: {response.status_code}')

