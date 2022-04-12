
# -*- coding: utf-8 -*-
"""
Este paquete permite trabajar con las herramientas de Google. Se irá actualizando poco a poco.

Antes de poder utilizar las funciones, es necesario almacenar en una variable la salida del método connect. Aquí tendré las referencias a los servicios de Google que se irán a utilizar.

Funciones:
----------
- Propósito general:
    - setProxy: Establece el uso de un proxy en la red actual.
    - getUserPwd: Solicita mediante ventana gráfica usuario y contraseña.
    - connect: Conecta con los servicios de Google para poder usarlos.
    - comprobarMailsDominios: Comprueba si un conjunto de direcciones forman parte de los dominios aceptados
    - getDriveIdFromURL: Devuelve el id de un fichero o carpeta a partir de su url

- Google Sheets: todas estas funciones necesitan el servicio 'sheets'
    - gshBorrarHoja: elimina una hoja en un libro de google.
    - gshCrearHoja: crea una nueva hoja en un libro de google.
    - gshDuplicarHoja: duplica una hoja en un libro de google.
    - gshCrearLibro: crea un nuevo libro GSheets y lo deja en "Mi unidad".
    - gshDescargarHoja: exporta una sola hoja de una gsheet.
    - gshEscribirHoja: escribe un pandas dataframe en una hoja de google.
    - gshEscribirRango: escribe una lista de valores en una hoja de google.
    - gshLeerHoja: lee el contenido de un rango de una hoja de un libro Google.
    - gshObtenerNombreHojas: devuelve los nombres de hojas que tiene un libro de Google.
    - gshObtenerNombreIdHojas: devuelve un diccionario con los nombres e IDs de hojas que tiene un libro de Google.
    - gshObtenerIDHoja: devuelve el ID de una hoja concreta dentro de una spreadsheet.
    - gshOrdenarHojas: Ordena las hojas dentro de una spreadsheet.
    - gshObtenerVistasFiltro: devuelve los nombres e ids de las vistas de filtro de hoja de google
    - gshCrearActualizarVistaFiltro: Actualiza o crea una o varias vistas de filtro
    - gshBorrarVistaFiltro: borra una vista de filtro, si existe.
    - gshObtenerRangosProtegidos: Devuelve todos los rangos protegidos de un libro.
    - gshProtegerRango: Esta función protege uno o varios rangos de un libro.
    - gshEliminarRangoProtegido: Elimina uno varios rangos protegidos de un libro.
    - gshTraducirColor: Recibe un color RGB y lo traduce a su representación JSON
    - gshTraducirRango: traduce un rango tipo 'A5:J27' a un objeto tipo GridRange de la API google
    - gshTraducirRangoInverso: traduce la representación JSON de un rango a formato de una sheet    
    - gshFormatearHoja: da formato corporativo básico a una hoja.
    - gshEliminarFormatoRango: Para un rango de celdas, las vuelve a poner con el formato por defecto
    - gshFormatearRango: para un rango de celdas, les da el formato especificado
    - gshEliminarFormatosCondicionalesHoja: Elimina todas las reglas de formato condicional de una hoja
    - gshFormatoCondicionalRango: Genera las reglas de formato condicional en un conjunto de rangos. 
    - gshEliminarValidacionRango: elimina todas las fórmulas de validación de un rango de celdas.
    - gshValidacionDesplegableRango: crea validaciones de lista desplegable en un rango de celdas.
    
- Google Docs: todas estas funciones necesitan el servicio 'docs'
    - gdcCrearDoc: crea un documento de Google Docs.
    - gdcFindTag: encuentra un texto marcado entre llaves dentro de un documento de google docs.
    - gdcInsertarImagen: inserta una imagen vía url en un documento de google docs.
    - gdcInsertarTexto: inserta un texto en un documento de Google Docs.
    - gdcReemplazarTexto: reemplaza un tag por un texto en un documento de google docs.
    
- GMail: todas estas funciones necesitan el servicio 'gmail'
    - ggmCreateDraft: crea un borrador a partir de un mensaje.
    - ggmCreateMessage: crea un mensaje de gmail.
    - ggmCreateMessageWithAttachment: crea un mensaje de gmail con un adjunto.
    - ggmCreateMessageWithAttachments: crea un mensaje de gmail con múltiples adjuntos.
    - ggmDescargarAdjuntos: recupera y descarga los adjuntos de un correo.
    - ggmGetPrimaryAddress: devuelve la dirección de mail principal
    - ggmLeerCorreo: recupera los detalles de un correo
    - ggmListarCorreos: lista una serie de correos dados unos criterios de búsqueda.
    - ggmSendMessage: envía un mensaje de gmail.
    
- Google Drive: todas estas funciones necesitan el servicio 'drive'
    - gdrCambiarPermisos: cambia los permisos de un elemento en drive.
    - gdrCopiarDocumento: realiza una copia de un documento drive.
    - gdrDescargarFichero: descargará un fichero de google drive en una ruta indicada.
    - gdrFindSharedDrive: busca una unidad compartida por nombre y devuelve su id.
    - gdrGetFileProperties: devuelve todas las propiedades de un fichero.
    - gdrGetFileRevisions: devuelve un listado de versiones de un fichero.
    - gdrGetFolderId: devuelve los identificadores de las carpetas encontradas en una ruta.
    - gdrListFiles: lista los ficheros contenidos en una carpeta de drive.
    - gdrMoveFile: mueve un fichero de una carpeta a otra.
    - gdrSubirVersión: a partir de un fichero local, sube una nueva versión a un fichero ya existente en Google Drive.
    - gdrUploadFile: sube un fichero local a Google Drive.
    
    
Versiones:
----------
2022-03-04: v1.12 Se añaden funciones para recuperar versiones de un fichero y para extraer el id a partir de su url
2021-10-29: v1.11 Se añaden campos cc a los métodos de envío de mail, y los métodos ggmCreateDraft y ggmCreateMessageWithAttachments
2021-08-31: v1.10 Nuevas funciones de trabajo con google sheets y funciones auxiliares
2021-03-30: v1.9. Se añade función gshDescargarHoja.
2021-03-25: v1.8. Se añade funcionalidad para quitar el proxy.
2021-02-12: v1.7. Se añaden funciones para trabajo con vistas de filtro y posibilidad de enviar mails en HTML
2021-02-01: v1.6. Fix en el método gshEscribirHoja (daba problema con columnas datetime con valores nulos).
2020-12-09: v1.5. Se añaden las funciones ggmDescargarAdjuntos, ggmLeerCorreo y ggmListarCorreos.
2020-11-27: v1.4. Se incluye control de caracteres latinos en ggmCreateMessage y se añaden valores predeterminados a los parámetros.
2020-11-25: v1.3. Incluye método ggmGetPrimaryAddress y corrigen bugs
2020-11-17: v1.2. Incluye método gdrCambiarPermisos, se da nueva funcionalidad a gdrGetFileProperties y se corrigen bugs
2020-11-13: v1.1. Incluye método getUserPwd y gdrSubirVersion
2020-10-27: v1.0. Incluye métodos para conectar a los servicios de google
"""
#Para conseguir las credenciales, acceder a esta URL:
#https://developers.google.com/sheets/api/quickstart/python
#Pulsar en enable API y DOWNLOAD CLIENT CONFIGURATION
#Guardar el archivo de credenciales en una ruta conocida.

#%%
#!pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

#%%
from __future__ import print_function
import pickle
import os
import os.path
from apiclient import errors

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.encoders import encode_base64
import mimetypes

from apiclient.http import MediaFileUpload,MediaIoBaseDownload
from apiclient.discovery import build
from httplib2 import Http

import pandas as pd

# If modifying these scopes, delete the file token.pickle.
#SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
def setProxy(user = None, pwd = None, port=8080):
    """
    Genera, para la sesion actual de Python, un proxy con el usuario y contraseña en proxyvip
    Parámetros:
    user      -- (str) Código del usuario
    pwd       -- (str) Contraseña del usuario
    port      -- (int) Indica el puerto en el que se quiere hacer la conexion. Por defecto 8080
    
    Salida: None
    """
    import os
    import getpass

    if user is None:
        user = getpass.getuser()
    if pwd is None:
        pwd = getpass.getpass(prompt = 'Por favor, indica la password corporativa para el usuario ' + user + ': ')
    os.environ["HTTPS_PROXY"] = "http://" + user + ":" + pwd + "@proxyvip:" + str(port)

def resetProxy():
    import os
    os.environ["HTTPS_PROXY"] = ""

#%%

def str2ascii(cadena):
    "Función que recibe una cadena y la devuelve cambiando caracteres latinos por caracteres ascii-7."
    table = {
        'á':'a',
        'é':'e',
        'í':'i',
        'ó':'o',
        'u':'u',
        'ñ':'n'
    }
    ls = list(table.keys())
    for l in ls:
        table[ord(l)] = table[l]
        table[ord(l.upper())] = table[l].upper()
    return str(cadena).translate(table)

#%%
def getUserPwd(titulo = 'Introduzca usuario y clave' , user=None):
    """
    Solicita mediante ventana gráfica usuario y contraseña. 
    Parámetros:
    titulo  -- (str) Indica el título que se mostrará en la ventana.
    user    -- (str) Código de usuario que se mostrará por defecto en el campo de la ventana. Si viene vacío se mostrará el usuario del sistema

    Salida:
    (diccionario) Diccionario con claves 'user' y 'password'. Si el usuario ha cerrado SIN pulsar el botón "Aceptar" devuelve None
    """
    from tkinter import Tk, Label, Entry, Button, StringVar
    import os
    
    if user is None:
        user = os.environ['USERNAME']
    
    
    ventana = Tk()
    ventana.title(titulo)
    ventana.geometry('220x150')
    
    #Usuario:
    texto_user =Label(ventana,text='Usuario:    ')
    texto_user.place(x=110, y=20, anchor="ne")    
    
    resp_user = StringVar()    
    resp_user.set(user)
    entry_user =Entry(ventana, textvariable=resp_user, width=10, relief="solid", borderwidth=1)
    entry_user.place(x=110, y=20, anchor="nw")
    
    
    
    #Contraseña:
    texto_pwd =Label(ventana,text='Contraseña:    ')
    texto_pwd.place(x=110, y=50, anchor="ne")    
    
    resp_pwd = StringVar() 
    entry_pwd =Entry(ventana, textvariable=resp_pwd, width=10, show='*', relief="solid", borderwidth=1)
    entry_pwd.place(x=110, y=50, anchor="nw")
    entry_pwd.bind("<Return>", (lambda event: (aceptar_pulsado.append(True) or ventana.destroy()  ) )   )
    
    
    
    aceptar_pulsado = []
    boton_aceptar = Button (ventana, text='Aceptar', 
                            command= lambda: (aceptar_pulsado.append(True) or ventana.destroy()  )
                            )
    boton_aceptar.place(x=110, y=105, anchor="center")
    
    ventana.mainloop()
    
    
    #Si ha pulsado el botón aceptar devolvemos usuario y password en una lista
    if len(aceptar_pulsado)>0:
        return {'user': resp_user.get(), 'password' : resp_pwd.get()}
    
    #Si ha dado al botón de cerrar devolvemos None
    else:
        return None


#%%  
def connect(ruta_credenciales = './', archivo_credenciales = 'credentials.json', ruta_token = None, archivo_token='token.pickle', services=[], port=8080):

    """Shows basic usage of the Sheets API.
    Establece una conexión con Google para poder leer excels de Drive.
    Parámetros:
    ruta_credenciales    -- (str) Indica en qué carpeta se encuentra el archivo de credenciales json.
    archivo_credenciales -- (str) Nombre del archivo .json que contiene las credenciales de usuario. Necesario si no existe token. ('credentials.json' por defecto).
    ruta_token           -- (str) Indica en qué carpeta se encuentra el archivo token pickle. Si se informe None, asume la misma que la de credenciales
    archivo_token        -- (str) Nombre del archivo que contiene o contendrá las credenciales de google ya aceptadas. Si se proporciona no solicita pantalla de permiso y lo genera. ('token.pickle' por defecto).
    services             -- lista de (str). Indica los servicios para los que se solicita canal. Ejemplos: ['sheets'], ['sheets', 'drive', 'docs']
    port                 -- (int) Indica el puerto en el que se quiere hacer la conexion. Por defecto 8080
    
    Salida:
    (diccionario) Diccionario cuya clave será el tipo de servicio y como valor el objeto servicio adecuado. Ej.: {'sheets': (objeto servicio de google)}.
    Posibles claves: 'sheets', 'drive', 'docs', 'gmail'
    """
    
    if ruta_token is None:
        ruta_token = ruta_credenciales
        
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = []
    for service in services:
        if service == 'sheets':
            SCOPES.append('https://www.googleapis.com/auth/spreadsheets')
        elif service == 'drive':
            SCOPES.append('https://www.googleapis.com/auth/drive.file')
            SCOPES.append('https://www.googleapis.com/auth/drive')
            SCOPES.append('https://www.googleapis.com/auth/drive.appdata')
            SCOPES.append('https://www.googleapis.com/auth/drive.metadata')

        elif service == 'docs':
            SCOPES.append('https://www.googleapis.com/auth/docs')
        elif service == 'gmail':
            SCOPES.append('https://mail.google.com/')
            SCOPES.append('https://www.googleapis.com/auth/gmail.readonly')
            SCOPES.append('https://www.googleapis.com/auth/gmail.send')
        elif service == 'drive.readonly':
            SCOPES.append('https://www.googleapis.com/auth/drive.readonly')

    if not (ruta_credenciales.endswith('/') or ruta_credenciales.endswith('\\')):
        ruta_credenciales = ruta_credenciales + '/'
        
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(ruta_token + archivo_token):
        with open(ruta_token + archivo_token, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                ruta_credenciales + archivo_credenciales, SCOPES)
            creds = flow.run_local_server(port=port)
        # Save the credentials for the next run
        with open(ruta_token + archivo_token, 'wb') as token:
            pickle.dump(creds, token)
    
    service_dict = {}
    for service in services:
        try:
            if service == 'sheets':
                service_dict[service] = build('sheets', 'v4', credentials=creds)
            elif service == 'drive':
                service_dict[service] = build('drive', 'v3', credentials=creds)
            elif service == 'docs':
                service_dict[service] = build('docs', 'v1', credentials=creds)
            elif service == 'gmail':
                service_dict[service] = build('gmail', 'v1', credentials=creds)
            elif service == 'drive.readonly':
                service_dict[service] = build('drive.readonly', 'v3', credentials=creds)
            elif service == 'script':
                service_dict[service] = build('script', 'v1', credentials=creds)
        except Exception as err:
            service_dict[service] = 'Error: ' + str(err)
    return service_dict



def comprobarMailsDominios (mails_destinatarios, dominios_aceptados):
    """
    Comprueba si un conjunto de direcciones forman parte de los dominios aceptados

    Parameters
    ----------
    mails_destinatarios : str
        Destinatarios del mail separados por comas tal como irán en el "to" del mail. Ejemplo: "dest1@bbva.com, dest2@gmail.com" .
    dominios_aceptados : list
        Lista de dominios aceptados. Ejemplo: ['bbva.com' , 'opplus.bbva.com'] .

    Returns
    -------
    bool
        True si TODOS los mails forman parte de alguno de los dominios aceptados. False si alguno no tiene un dominio aceptado.

    """
    #primero compruebo que los mails tengan al menos una vez formato xxx@xxx
    if len(mails_destinatarios.split('@')) < 2 :
        return False
    
    #ahora al menos debería tener 1 dirección de mail. Para cada mail que venga comprobamos su dominio.
    lista_destinatarios = mails_destinatarios.split(',')
        
    for destinatario in lista_destinatarios:
        try:
            dominio = destinatario.split('@')[1]
            if dominio not in dominios_aceptados:
                return False
            
        except Exception:
            return False
    
    return True


#%%
def getDriveIdFromURL(linkdrive, folder=False):
    """
    Devuelve el id de un fichero o carpeta en Drive a partir de su url
    
    Parámetros:
    -----------
    linkdrive: (str) url de compartir en drive (extraida de obtener enlace para compartir)
    folder   : (bool) True si el link pertenece a una carepta, False si pertenece a un archivo.
    
    Devuelve:
    ---------
    str  Cadena con el id del archivo o carpeta.
    """
    import re
    if folder or 'drive/folders/' in linkdrive:
        iddrive = re.match(r'https://drive.google.com/drive/folders/(.*)', linkdrive).group(1)
    else:
        iddrive = re.match(r'https://drive.google.com/file/d/(.*)/.*', linkdrive).group(1)
    return iddrive
#%%


#%%
######################################################################################################
# GOOGLE SHEETS
######################################################################################################
FORMATOS_BASE = {
        'euro_2_dec'            : {"type" : "CURRENCY" , "pattern" : '#,##0.00 €'} ,
        'euro_no_dec'           : {"type" : "CURRENCY" , "pattern" : '#,##0 €'   },
        
        'fecha_estandar'        : {"type" : "DATE"     , "pattern" : 'dd/mm/yyyy'},
        'fecha_ordenada'        : {"type" : "DATE"     , "pattern" : 'yyyy-mm-dd'},
        'fechahora_estandar'    : {"type" : "DATE_TIME", "pattern" : 'dd/mm/yyyy HH:MM:SS'},
        'fechahora_ordenada'    : {"type" : "DATE_TIME", "pattern" : 'yyyy-mm-dd HH:MM:SS'},

        'num_sep_mil_2dec'      : {"type" : "NUMBER"   , "pattern" : '#,##0.00_);[Red]-#,##0.00'},
        'num_no_sep_mil_2dec'   : {"type" : "NUMBER"   , "pattern" : '###0.00_);[Red]-###0.00'},
        
        'num_sep_mil_no_dec'    : {"type" : "NUMBER"   , "pattern" : '#,##0_);[Red]-#,##0'},
        'num_no_sep_mil_no_dec' : {"type" : "NUMBER"   , "pattern" : '###0_);[Red]-###0'},
        
        'pct_1_dec'             : {"type" : "PERCENT"  , "pattern" : '0.0%'},
        'pct_2_dec'             : {"type" : "PERCENT"  , "pattern" : '0.00%'},
        
        'texto'                 : {"type" : "TEXT"     , "pattern" : '@'},
        'texto_4_posic'         : {"type" : "TEXT"     , "pattern" : '0000_@'},
        'texto_14_posic'        : {"type" : "TEXT"     , "pattern" : '00000000000000_@'},
        'texto_26_posic'        : {"type" : "TEXT"     , "pattern" : '00000000000000000000000000_@'}
        }

def gshBorrarHoja(sheets_service, spreadsheetId, nombreHoja=None, idHoja=None):
    """
    Esta función elimina una hoja en un libro de google.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja que se desea borrar.
    idHoja : str, optional
        Si se proporciona el ID de la hoja se ignora el nombre. The default is None.

    Returns
    -------
    res : json 
        resultado de la ejecución.
    """
    
    if idHoja is None:
        idHoja = gshObtenerIDHoja(sheets_service, spreadsheetId, nombreHoja)
    if idHoja is not None:
        res = sheets_service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId,
                                                 body={
                                                     'requests': [
                                                         {'deleteSheet':{
                                                             'sheetId': idHoja
                                                             }
                                                         }
                                                     ]
                                                 }
                                                ).execute()
        return res
    else:
        print('La hoja no existe, no se realiza ninguna acción')

#%%
def gshCrearHoja(sheets_service, spreadsheetId, nombreHoja, nFilas = 100, nCols = 30):
    """
    Esta función crea una nueva hoja en un libro de google.
    
    sheets_service = Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId = str, Identificador del libro en Google Drive.
    nombreHoja = str, Nombre de la hoja en la que se desea crear dentro del libro.
    nFilas = int, opcional. Número de filas que se quiere que tenga la nueva hoja (por defecto 100)
    nCols = int, opcional. Número de columnas que se quiere que tenga la nueva hoja (por defecto 30)
    """
    if nombreHoja not in gshObtenerNombreHojas(sheets_service, spreadsheetId):
        res = sheets_service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId,
                                                 body={
                                                     'requests': [
                                                         {'addSheet':{
                                                             'properties': {
                                                                 'title': nombreHoja,
                                                                 'gridProperties': {
                                                                     'rowCount':nFilas,
                                                                     'columnCount':nCols
                                                                 }
                                                             }
                                                         }
                                                         }
                                                     ]
                                                 }
                                                ).execute()
        return res
    else:
        print('La hoja ya existe, no se realiza ninguna acción')
        return None



def gshDuplicarHoja(sheets_service, spreadsheetId, nombreHojaOrig, nombreHojaNueva, indiceNuevaHoja = None):
    """
    Esta función duplica una hoja en un libro de google.

    Parameters
    ----------
    sheets_service : Service
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHojaOrig : str
        Nombre de la hoja que se va a copiar.
    nombreHojaNueva : str
        Nombre de la hoja resultante.
    indiceNuevaHoja : int, default None
        Indice que debe ocupar la nueva hoja (base 0). Si no se pasa, la nueva hoja será la primera. 

    Returns
    -------
    res : json
        resultado de la ejecución. None en caso de error.
    """
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    hojasActualesLibro = dict()
    for sheet in sheets:
        hojasActualesLibro[ sheet['properties']['title'] ] = sheet['properties']['sheetId']
    
    
    
    if nombreHojaOrig not in hojasActualesLibro.keys():
        print('No se encuentra la hoja', nombreHojaOrig)
        return None
    
    elif nombreHojaNueva not in hojasActualesLibro:
        res = sheets_service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId,
                                                        body={
                                                            'requests': [
                                                                {'duplicateSheet':{
                                                                    "sourceSheetId"   : hojasActualesLibro[nombreHojaOrig],
                                                                    "insertSheetIndex": indiceNuevaHoja,
                                                                    "newSheetName"    : nombreHojaNueva
                                                                    }
                                                                }
                                                            ]
                                                            }
                                                        ).execute()
        return res
    else:
        print('La hoja', nombreHojaNueva, 'ya existe, no se realiza ninguna acción')


#%%
def gshCrearLibro(sheets_service, nombreLibro, nombreHoja = None):
    """
    Esta función crea un nuevo libro GSheets y lo deja en "Mi unidad".

    sheets_service = Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    nombreLibro = str, Nombre del libro.
    nombreHoja = str, opcional. Nombre de la única hoja que tendrá el libro.
    
    returns objeto dict_like con los detalles del libro.
    """
    spreadsheet_body = {
      "properties": {
        "title": nombreLibro,
        "locale": ""
      },
    }
    if nombreHoja is not None:
        spreadsheet_body['sheets'] = [
            {
                "properties": {
                    "title": nombreHoja
                }
            }
        ]

    request = sheets_service.spreadsheets().create(body=spreadsheet_body)
    return request.execute()

#%%
def gshDescargarHoja(sheets_service, spreadsheetid, hojaId = None, nombreHoja = None, ficheroDescarga = None, tipo='pdf', sobreescribir = False):
    """
    Esta función descargará una sola hoja de una gsheet a un fichero.
    sheets_service  -- Objeto servicio con permisos de lectura en Google Sheets. Se obtiene con el método connect.
    spreadsheetid   -- (str) Identificador de la gsheet en Drive.
    hojaId          -- (str) Identificador de la hoja (gid)
    nombreHoja      -- (str) Nombre de la hoja. Si se proporciona hojaId, no se hace caso a este parámetro.
    ficheroDescarga -- (str) Fichero de descarga en la que dejar el fichero. Si no se indica, lo dejará en la carpeta de trabajo con el nombre de la gsheet y el id de la hoja descargada.
    tipo            -- (str) Formato de exportación. Por defecto, 'pdf'. Admite también 'xlsx' y 'csv'
    sobreescribir   -- (bool) Default False. Indica si debe machacar el fichero en caso de que exista, o crear un nombre nuevo en su lugar.
    
    return          -- (str) Ruta completa del fichero descargado.
    """
    import urllib.parse
    import os

    result = sheets_service.spreadsheets().get(spreadsheetId = spreadsheetid).execute()
    spreadsheetUrl = result['spreadsheetUrl']
    if nombreHoja is None and hojaId is None:
        print('Necesito o el id de la hoja (gid) o el nombre de la misma')
        return None
    elif hojaId is None:
        for sheet in result['sheets']:
            if sheet['properties']['title'] == nombreHoja:
                hojaId = sheet['properties']['sheetId']
    if hojaId is None:
        print('Hoja no encontrada')
        return None
    
    exportUrl = spreadsheetUrl.replace("/edit", '/export')
    params = {
        'format': tipo,
        'gid': hojaId,
    } 
    queryParams = urllib.parse.urlencode(params)
    url = exportUrl + '&' + queryParams
    
    resp, content = sheets_service._http.request(url)
    if ficheroDescarga is None:
        ficheroDescarga = result['properties']['title'] + '_' + str(hojaId) + '.' + tipo
    if not sobreescribir:
        if os.path.isfile(ficheroDescarga):
            fichero = os.path.splitext(ficheroDescarga)
            i = 1
            while True:
                rutaNueva = fichero[0] + ' (' + str(i) + ')' + fichero[1]
                if not os.path.isfile(rutaNueva):
                    ficheroDescarga = rutaNueva
                    break
                i=i+1

    with open(ficheroDescarga, 'wb') as descarga:
        descarga.write(content)
    return ficheroDescarga
    
#%%
def gshEscribirHoja(sheets_service, dataframe, spreadsheetId, nombreHoja, rango = None, replace = True, header=True):
    """
    Esta función escribe una dataframe en una hoja de google. Si la hoja que se ha pasado para escribir no existe, la crea. 
    
    sheets_service -- Objeto servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    dataframe      -- (pandas.DataFrame) Valores que se quieren llevar a la hoja.
    spreadsheetId  -- (str) Identificador del libro en Google Drive.
    nombreHoja     -- (str) Nombre de la hoja en la que se desea escribir dentro del libro.
    rango          -- (str) opcional. Rango de celdas que escribir. Debe estar escrito en notación de hoja de cálculo. Ejemplos: A1, B2:C2, B4:F10
                            Si se pasa un rango de 1 sola celda pero el dataframe tiene más de 1 valor y el parámetro replace está a true
                            se limpia la hoja desde la celda indicada hasta el final.
    replace        -- (boolean) opcional. Indica si se desea reemplazar el contenido completo de la hoja (y rango si se especifica).
    header         -- (boolean) opcional. Indica si se desea incluir la cabecera cuando se escriba el dataframe.
    """
    import re
    
    sheet = sheets_service.spreadsheets()
    #Comprobamos si la hoja existe. En caso contrario, se crea:
    if nombreHoja not in gshObtenerNombreHojas(sheets_service, spreadsheetId):
        gshCrearHoja(sheets_service, spreadsheetId=spreadsheetId, nombreHoja=nombreHoja)
    #Por si viene rango, envolvemos el nombre de la hoja entre comillas
    nombreHoja = "'" + nombreHoja + "'"
    if not rango is None:
        nombreHoja = nombreHoja + "!" + rango
    
    if replace:
        #Si el rango es de una sola celda y el dataframe trae más de un valor se limpia la hoja entera desde esa celda hasta el final
        if rango is not None and len(rango.split(':')) == 1 and (dataframe.shape[0]*dataframe.shape[1]) > 1 :
            fila = re.split('(\d+)',rango)[1]
            hoja_borrar = nombreHoja + ':'+ str(fila)
            result = sheet.values().clear(spreadsheetId = spreadsheetId,
                                          range = hoja_borrar).execute()
            #del resultado capturamos cuál es la última columna y borramos por columnas
            rango_borrado = result['clearedRange']
            if len(rango_borrado.split(':'))>1:
                columna = re.split( '(\d+)',  rango_borrado.split(':')[1]  )[0]
                hoja_borrar = nombreHoja + ':'+ str(columna)
                result = sheet.values().clear(spreadsheetId = spreadsheetId,
                                              range = hoja_borrar).execute()
            
            
        #Si el rango comprende más de una celda borramos sólo ese rango
        else:
            result = sheet.values().clear(spreadsheetId = spreadsheetId,
                                          range = nombreHoja).execute()

    copia = dataframe.copy()
    #2021-12-14: Puede haber varias columnas con el mismo nombre. En ese caso, el dtype fallaría porque no devolvería una serie, sino un df. Corregimos usando iloc.
    for col in range(len(copia.columns)):
        if copia.iloc[:,col].dtype == 'datetime64[ns]':
            copia.iloc[:,col] = copia.iloc[:,col].dt.strftime('%Y-%m-%d %H:%M:%S')
    copia = copia.fillna('') #Cambio 2021-02-01. Hacer esto antes del cambio de formato implica que las columnas de tiempo, si tienen nulos, pasan a ser objects, por lo que no las convierte a texto, y la subida petaba.
    #2021-02-01: Ahora también habría que reemplazar los NaT que hubiesen poder haber quedado kaputt:
    copia = copia.replace('NaT','')
    
    lista = copia.values.tolist()
    
    if header:
        lista.insert(0, dataframe.columns.tolist())
        
    result = sheets_service.spreadsheets().values().update(spreadsheetId = spreadsheetId,
                                                    range = nombreHoja,
                                                    body = {
                                                        'majorDimension': 'ROWS',
                                                        'values': lista
                                                    },
                                                    valueInputOption='RAW'
                                                   ).execute()

    return result

#%%
def gshEscribirRango(sheets_service, lista, spreadsheetId, nombreHoja, rango, fillna = None):
    """
    Esta función escribe una lista de valores en una hoja de google.
    
    sheets_service = Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    lista = [[]] lista de 2 dimensiones con los valores que se desean escribir. Cada elemento de la lista representa una fila. 
    Cada elemento debe ser una lista a su vez, con los valores de las columnas de dicha fila. No admite tipos datetime, hay que convertirlos previamente a str.
    spreadsheetId = str, Identificador del libro en Google Drive.
    nombreHoja = str, Nombre de la hoja en la que se desea escribir dentro del libro.
    rango = str, Rango de celdas que escribir, o bien celda en la que escribir el primer valor. Debe estar escrito en notación de hoja de cálculo. Ejemplos: A1, B2:C2, B4:F10
    fillna = Especifica si debe cambiar los valores nulos por algo. Posibles valores: Si True cambia los nulos por cadena vacía. Si es un str, cambia los nulos por dicho valor. Si es False o None, no cambia los valores nulos. En este caso, si existe valor en una celda, no lo sobreescribirá.
    """
    sheet = sheets_service.spreadsheets()
    #Comprobamos si la hoja existe. En caso contrario, se crea:
    if nombreHoja not in gshObtenerNombreHojas(sheets_service, spreadsheetId):
        gshCrearHoja(sheets_service, spreadsheetId=spreadsheetId, nombreHoja=nombreHoja)
    #Por si viene rango, envolvemos el nombre de la hoja entre comillas
    nombreHoja = "'" + nombreHoja + "'!" + rango
    
    if fillna is not None:
        if type(fillna) == bool and fillna:
            fillna = ''
        if type(fillna) == str:
            lista = [[fillna if v is None else v for v in l] for l in lista]
        
    result = sheets_service.spreadsheets().values().update(spreadsheetId = spreadsheetId,
                                                    range = nombreHoja,
                                                    body = {
                                                        'majorDimension': 'ROWS',
                                                        'values': lista
                                                    },
                                                    valueInputOption='RAW'
                                                   ).execute()

    return result

#%%
def gshLeerHoja(sheets_service, spreadsheetId, nombreHoja, rango = None, formato_valores = 'UNFORMATTED_VALUE', formato_fechas = 'FORMATTED_STRING'):
    """
    Esta función lee el contenido de un rango de una hoja de un libro Google. Basada en el servicio: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/get

    Parameters
    ----------
    sheets_service : object service
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desea crear dentro del libro.
    rango : str, optional
        Rango de celdas que leer. Debe estar escrito en notación de hoja de cálculo. Ejemplos: 'A1', 'B2:C2', 'B4:F10'. 
        The default is None.
    formato_valores : str, optional, default = 'UNFORMATTED_VALUE'
        Indica 'FORMATTED_VALUE' si debe traer el valor tal cual se representa en la gsheet (con el formato de moneda, por ejemplo), o 'UNFORMATTED_VALUE'si debe recuperar el valor original.
    formato_fechas: str, optional, default = 'FORMATTED_STRING'
        Indica cómo debe representar las fechas en la salida. Posibles valores: 'SERIAL_NUMBER' o 'FORMATTED_STRING' (https://developers.google.com/sheets/api/reference/rest/v4/DateTimeRenderOption)

    Returns
    -------
    df : pandas.DataFrame
        DataFrame con los valores leídos.

    """
    import pandas as pd
    # Call the Sheets API
    sheet = sheets_service.spreadsheets()
    nombreHoja = "'" + nombreHoja + "'"
    if not rango is None:
        nombreHoja = nombreHoja + "!" + rango
       
    result = sheet.values().get(
        spreadsheetId        = spreadsheetId,
        range                = nombreHoja,
        valueRenderOption    = formato_valores,
        dateTimeRenderOption = formato_fechas
    ).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
        return None
    else:
        cols = values[0]
        df = pd.DataFrame.from_records(data=values[1:], columns = cols)
        # df.columns = df.iloc[0]
        # df = df.drop(0)
        # while len(df.columns) < len(cols):
        #     i = len(df.columns)
        #     df[cols[i]] = None
        # df.columns = cols
        return df
#%%
def gshObtenerNombreHojas(sheets_service, spreadsheetId):
    """
    Esta función devuelve los nombres de hojas que tiene un libro de Google.

    INPUT:
    - sheets_service = Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    - spreadsheetId  = str, Identificador del libro en Google Drive.
    
    returns list con los nombres de las hojas.
    """
    # Call the Sheets API
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheetnames = []
    for sheet in sheets:
        sheetnames.append(sheet['properties']['title'])
    return sheetnames


#%%

def gshObtenerNombreIdHojas(sheets_service, spreadsheetId):
    """
    Esta función devuelve un diccionario con claves los nombres de hojas que tiene un libro de Google y valor sus IDs.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente..
    spreadsheetId : str
        Identificador del libro en Google Drive.

    Returns
    -------
    sheetNamesIDs : dict
        Diccionario con los nombres de hojas como claves y los IDs como valores
    """
    
    # Call the Sheets API
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheetNamesIDs = dict()
    for sheet in sheets:
        sheetNamesIDs[ sheet['properties']['title'] ] = sheet['properties']['sheetId']
    return sheetNamesIDs


#%%
def gshObtenerIDHoja(sheets_service, spreadsheetId, nombreHoja):
    """
    Devuelve el ID de una hoja concreta dentro de una spreadsheet.
    
    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja buscada.      
    
    Returns
    -------
    str
        Id de la hoja dentro de la spreadsheet. None si no se localiza
    """   
    
    #Localizamos la hoja pedida y capturamos su ID y el número de columnas
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    
    sheetId = None
    for sheet in sheets:
        if sheet['properties']['title'] == nombreHoja :
            sheetId = sheet['properties']['sheetId']
            break
    
    if sheetId is None:
        print("La hoja con el nombre solicitado no existe")
        
    return sheetId



#%%    
def gshOrdenarHojas(sheets_service, spreadsheetId, nombresHojasOrdenadas):
    """
    Ordena las hojas dentro de una spreadsheet. Recibe una lista con los nombres de las 
    hojas en el orden deseado. El resto de hojas de la spreadsheet que no aparezcan en 
    la lista quedarán al final en el mismo orden en el que estén al entrar en la función.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombresHojasOrdenadas : list
        Lista con los nombres de hoja en el orden que se desean.

    Returns
    -------
    res : json
        resultado de la ejecución.
    """
    if type(nombresHojasOrdenadas) != list:
        print ('Se espera una lista con los nombres de hoja en el orden que se quieren')
        return None
        
    #Localizamos las propiedades de todas las hojas
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheet_names = [sheet['properties']['title'] for sheet in sheets]
    
    #Para cada hoja de la lista actualizamos su índice y lo metemos en la lista de requests
    indice=0
    requests = []
    for nombreHoja in nombresHojasOrdenadas :
        if nombreHoja not in sheet_names:
            continue
        
        properties = sheets[sheet_names.index(nombreHoja)]['properties']
        properties['index'] = indice
        req = {'updateSheetProperties': {
                    'properties' : properties,
                    'fields'     : '*'
                    }               
               }
        requests.append(req)
        indice += 1
        
    #Para aquellas hojas del libro que no se hayan pasado en la lista ordenada las dejamos al final en el orden que vienen
    for nombreHoja in sheet_names:
        if nombreHoja not in nombresHojasOrdenadas:
            properties = sheets[sheet_names.index(nombreHoja)]['properties']
            properties['index'] = indice
            req = {'updateSheetProperties': {
                        'properties' : properties,
                        'fields'     : '*'
                        }               
                    }
            requests.append(req)
            indice += 1
            
    #Finalmente lanzamos la petición
    res = sheets_service.spreadsheets().batchUpdate(
        spreadsheetId = spreadsheetId,
        body ={'requests' : requests}
        ).execute()
    
    return res



#%%    
def gshObtenerVistasFiltro(sheets_service, spreadsheetId, nombreHoja):
    """
    Esta función devuelve un diccionario con las vistas de filtro que hay en una hoja de un libro google.
    La clave del diccionario es el NOMBRE de la vista de filtro. El contenido es el id de la vista de filtro.
    En caso de error, la función devuelve None.
    Si la hoja no tiene ninguna vista de filtro, se devuelve un diccionario vacío. 
    
    INPUT:
    - sheets_service = Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    - spreadsheetId = str, Identificador del libro en Google Drive.
    - nombreHoja = str, Nombre de la hoja en la que se desea escribir dentro del libro.
    """
    
    #Localizamos las propiedades de todas las hojas
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheet_names = [sheet['properties']['title'] for sheet in sheets]
    
    #Comprobamos si la hoja existe. En caso contrario, se devuelve None:
    if nombreHoja not in sheet_names:
        return None
    
    #Comprobamos en qué posición está la hoja solicitada. No va a fallar porque en el if anterior aseguramos que está.
    indice_hoja = sheet_names.index(nombreHoja)
    
        
    #Si esa hoja en particular no tiene vistas de filtro, devolvemos diccionario vacío.
    if 'filterViews' not in sheet_metadata['sheets'][indice_hoja].keys():
        return dict()
    
    vistas_filtro = sheet_metadata['sheets'][indice_hoja]['filterViews']
    
    #Pasamos las vistas a un diccionario con clave el nombre de la vista y dato el id.
    result = dict()
    for filtro in vistas_filtro:
        result[filtro['title']] = filtro['filterViewId']

    return result


#%%
def gshListarOpcionesFiltro():
    opciones_condicion = ['CONDITION_TYPE_UNSPECIFIED', 
                          'NUMBER_GREATER', 'NUMBER_GREATER_THAN_EQ', 'NUMBER_LESS', 'NUMBER_LESS_THAN_EQ', 'NUMBER_EQ', 'NUMBER_NOT_EQ', 'NUMBER_BETWEEN', 'NUMBER_NOT_BETWEEN',
                          'TEXT_CONTAINS', 'TEXT_NOT_CONTAINS', 'TEXT_STARTS_WITH', 'TEXT_ENDS_WITH', 'TEXT_EQ', 'TEXT_IS_EMAIL', 'TEXT_IS_URL', 
                          'DATE_EQ', 'DATE_BEFORE', 'DATE_AFTER', 'DATE_ON_OR_BEFORE', 'DATE_ON_OR_AFTER', 'DATE_BETWEEN', 'DATE_NOT_BETWEEN', 'DATE_IS_VALID', 
                          #'ONE_OF_RANGE', 'ONE_OF_LIST', #Not supported in filters
                          'BLANK', 'NOT_BLANK', 
                          #'CUSTOM_FORMULA', 'BOOLEAN'  #Not supported in filters
                          ]
    return opciones_condicion
       
#%%
def gshCrearActualizarVistaFiltro(sheets_service, spreadsheetId, nombreHoja, 
                                  nombreVista , idVista=None, primeraCelda = 'A1', condiciones = [], idHoja = None):
    """
    Esta funcion actualiza una o varias vistas de filtro, si existen. Si no, crea una/s vista/s de filtro. Devuelve sus IDs.
    OJO! Si se quiere modificar una vista de filtro, se pueden modificar sus condiciones, pero no su rango. Para eso deberá 
    borrarse la vista de filtro y crearse una nueva.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desea escribir dentro del libro.
    nombreVista : str or list
        Nombre que se va a poner a la vista de filtro. Si es una lista se crearán tantas vistas de filtros como elementos haya.
    idVista : str o list, optional
        ID o lista de IDs de las vistas de filtro a modificar, si ya existen. The default is None.
    primeraCelda : str, optional
        Celda en la que empiezan las vistas de filtro.Si se piden varias vistas de filtro todas tendrán el mismo ámbito. 
        Si se quieren varias vistas de filtro sobre la misma hoja pero distinto conjunto de datos habrá que llamar a la función varias veces.
        The default is 'A1'.
    condiciones : list o list of list, optional
        Si el parámetro "nombreVista" es un str debe ser una lista:
            Cada elemento de esa lista debe ser un diccionario con las condiciones de los filtros:
            {'columna' : (nombre de la columna)
            'tipo_condicion' : (condicion a elegir de la lista devuelta por gshListarOpcionesFiltro() ó 'hiddenValues')
            'valor' : (valor o lista de valores a aplicar en el filtro)
            }
        Si el parámetro "nombreVista" es una lista, debe ser una lista con la misma longitud que 'nombreVista':
            Cada elemento de la lista debe ser a su vez una lista de diccionarios. Cada diccionario contiene una condición del filtro.
            Ejemplo: Supongamos que    nombreVista = ['ofi_1234_mas_1000', 'ofi_4567']:
                condiciones = [
                    [ {'columna' : 'oficina', 'tipo_condicion' : 'TEXT_EQ'        , 'valor': '1234' }
                      {'columna' : 'saldo'  , 'tipo_condicion' : 'NUMBER_GREATER' , 'valor': 1000   }
                    ],
                    
                    [ {'columna' : 'oficina', 'tipo_condicion' : 'TEXT_EQ'        , 'valor': '4567' }
                    ]                    
                ]
        The default is [].
        
    idHoja : str, optional
        Id de la hoja dentro de la spreadsheet. Si se especifica este parámetro se ignora el 'nombreHoja'. The default is None.

    Returns
    -------
    str o list
        ID o lista de IDs de las vistas de filtro creadas/actualizadas. None en caso de error

    """
    
    if idHoja is None:
        idHoja = gshObtenerIDHoja(sheets_service, spreadsheetId, nombreHoja) 
        
    if idHoja is None:
        print("No se pueden crear/modificar las vistas de filtro")
        return None
    
    #Comprobación de argumentos:
    if type(primeraCelda)==list:
        print('Error: la primera celda debe ser una cadena de texto. Ejemplo: "A3"')    
        return None
    
    if type(nombreVista)==list:
        if type(condiciones) != list or len(condiciones)!= len(nombreVista):
            print('Error: No se ha especificado un conjunto de condiciones por cada vista solicitada')
            return None
        if idVista is not None and type(idVista) != list:
            print('Error: No se ha especificado un id de vista por cada vista solicitada')
            return None
        if type(primeraCelda)==list and len(primeraCelda)!=len(nombreVista):
            print('Error: No se ha especificado una primera celda por cada vista solicitada')
            return None
    elif type(nombreVista)==str:
        if idVista is not None and type(idVista) != str:
            print('Error: Se ha especificado 1 nombreVista pero una lista de idVista.')
            return None
    else:
        print('Error: el parámetro nombreVista debe ser de tipo str o list')
        return None
    
    
    #Convertimos todos los parámetros a lista. Ya sabemos que están bien porque lo hemos comprobado arriba
    if type(nombreVista) == str:
        nombreVista=[nombreVista]
        condiciones = [condiciones]
    
    if idVista is None:
        idVista=[]
        for vista in nombreVista:
            idVista.append(None)
    elif type(idVista) == str:
        idVista=[idVista]
        
    
    vistas_actuales = gshObtenerVistasFiltro(sheets_service, spreadsheetId, nombreHoja)
    
    rango_filtro = gshTraducirRango(rangoLetras = primeraCelda, sheetId=idHoja)
    rango_filtro.pop('endColumnIndex', None)
    rango_filtro.pop('endRowIndex', None)
    
    #Calculamos los números de la primera celda que nos pasan.
    primera_fila_tx = str(rango_filtro['startRowIndex']  + 1 ) 
    
    #Para luego poder generar las condiciones, me traigo las columnas de la gsheet    
    columnas_hoja = [col.upper() for col in gshLeerHoja(sheets_service, spreadsheetId, nombreHoja, rango = primera_fila_tx+':' + primera_fila_tx).columns]
    
    
    requestsVistas = []
    for indiceVista, vistaActual in enumerate(nombreVista):
        idVistaActual = idVista[indiceVista]
        condicionesActuales = condiciones[indiceVista]
        
        
        #Si esa hoja en particular no tiene vistas de filtro, hay que crearla.
        if len(vistas_actuales) == 0:
            tipo_request = 'addFilterView'
        
        #Si nos han pasado un ID de la vista a modificar y efectivamente, existe, es un update
        elif (idVistaActual is not None) and (int(idVistaActual) in vistas_actuales.values() ) :
            tipo_request = 'updateFilterView'
          
        #En otro caso, tengo que crear la vista de cero.
        else:
            tipo_request = 'addFilterView'
            
        
        #Generamos las condiciones. Para ello localizo las columnas de las condiciones en las de la gsheet
        criterios = {}
        cols_solicitadas_filtro = [condicion['columna'].upper() for condicion in condicionesActuales ]
        
        #Ahora, por cada columna de la gsheet, si me han solicitado un filtro, lo pongo. Si no, le quito el filtro (pongo {})
        for indice, columna in enumerate(columnas_hoja):
            if columna not in cols_solicitadas_filtro:
                criterios[indice] = {}
            else:
                indice_condicion = cols_solicitadas_filtro.index(columna)
                condicion = condicionesActuales[indice_condicion] 
                
                if condicion['tipo_condicion'] == 'hiddenValues':
                    criterios[indice] = { 'hiddenValues': condicion['valor'] }
                    
                else:
                    criterios[indice] = { 'condition': { 'type': condicion['tipo_condicion'], 'values':{"userEnteredValue":condicion['valor']}  }  }
                
                  
        #Preparamos la solicitud de actualizacion
        FilterViewRequest = {
            tipo_request: {
                'filter': {
                    'title': vistaActual,
                    'range': rango_filtro,
                    'criteria': criterios
                }
            }
        }
                
        if tipo_request == 'updateFilterView':
            FilterViewRequest['updateFilterView']['filter']['filterViewId'] = str(idVistaActual)
            FilterViewRequest['updateFilterView']['fields'] = {'paths': ['criteria', 'title'] }
       
        #Agrego la request que acabo de crear para esta vista concreta a la lista de requests:
        requestsVistas.append(FilterViewRequest)
         
    #Ahora realmente ejecutamos la solicitud y recuperamos el ID de la vista de filtro generada.
    try:
        body = {'requests': requestsVistas}
        FilterViewResponse = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=body).execute()
    except Exception as e:
        print("Error actualizando o creando la vista de filtro:", e)
        return None
    
    #Finalmente, localizamos el ID de las vistas de filtro y lo devolvemos
    respuesta = []
    for indicePetic, peticion in enumerate(requestsVistas):
        if 'updateFilterView' in peticion.keys():
            respuesta.append( peticion['updateFilterView']['filter']['filterViewId'] )
        
        elif 'addFilterView' in peticion.keys():
            respuesta.append(str( FilterViewResponse['replies'][indicePetic][tipo_request]['filter']['filterViewId'] )) 
    
    return respuesta
    


#%%
def gshBorrarVistaFiltro(sheets_service, spreadsheetId, nombreHoja, idVista=None):
    """
    Esta funcion borra una o varias vistas de filtro, si existen. Devuelve True si se borra bien, False si hay un error o no existe.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desea escribir dentro del libro.
    idVista : str o list, optional
        ID o lista de IDs de las vistas de filtro a borrar. Si se pasa el valor None se borran TODAS las vistas 
        de filtro de la hoja indicada. The default is None.

    Returns
    -------
    bool
        True si se ha borrado. False si hay un error o la(s) vista(s) no existen.
    """
    
    vistas_actuales = gshObtenerVistasFiltro(sheets_service, spreadsheetId, nombreHoja)
    if vistas_actuales is None:
        print("ERROR: No existe la hoja solicitada")
    
    #Si no había vistas de filtro en la hoja, ya hemos acabado. 
    if len(vistas_actuales.keys()) == 0 and idVista is None:
        return True
    
    requests = []
    if idVista is None:
        for vista in vistas_actuales.values():
            requests.append( {'deleteFilterView': { 'filterId': str(vista) }  }  )
                
    elif type(idVista) == str:
        if (int(idVista) not in vistas_actuales.values() ) :
            print("ERROR: La vista con el id solicitado no existe")
            return False
        else:
            requests.append( {'deleteFilterView': { 'filterId': str(idVista) }  }  )
            
    elif type(idVista) == list:
        for vista in idVista:
            if (int(vista) not in vistas_actuales.values() ) :
                print("ERROR: La vista con el id solicitado no existe")
                return False
            else:
                requests.append( {'deleteFilterView': { 'filterId': str(vista) }  }  )
    else:
        print("ERROR: idVista tiene un tipo no admitido")
        return False
    
            
    #Ahora realmente ejecutamos la solicitud.
    try:
        body = {'requests': requests}
        FilterViewResponse = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=body).execute()
    except Exception as e:
        print("Error borrando la vista de filtro:", e)
        return False
    
    return True



#%%
def gshObtenerRangosProtegidos(sheets_service, spreadsheetId):
    """
    Devuelve todos los rangos protegidos de un libro. Devuelve un diccionario donde las claves son
    las hojas que tienen rangos protegidos. Los valores son una lista con los rangos protegidos de esa hoja.
    Cada valor de la lista es un diccionario que representa un rango protegido. Si el range es None indica que
    toda la hoja está protegida
    
    Ejemplo de resultado:
        {
            'Sheet1': [
                {'protectedRangeId': '0123456789', 'usersEdit': ['usr1@bbva.com', 'usr2@bbva.com'], 'range': 'A1:AB100'                              }
                {'protectedRangeId': '1234567890', 'usersEdit': ['usr3@bbva.com']                 , 'range': 'A3:A25' ,'groupsEdit':['grp1@bbva.com']}
            ],
            'Sheet2': [
                {'protectedRangeId': '2345678901', 'usersEdit': ['usr1@bbva.com', 'usr2@bbva.com'], 'range': None                                    }
            ]            
        }

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.

    Returns
    -------
    rangos_protegidos : dict
        diccionario con los rangos protegidos de todo el libro. Leer descripción y ejemplo en la parte superior.

    """
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    
    rangos_protegidos = {}
    for hoja in sheet_metadata['sheets']:
        if 'protectedRanges' in hoja.keys():
            nom_hoja = hoja['properties']['title']
            rangos_protegidos[nom_hoja] = []
            for rango_prot in hoja['protectedRanges']:
                rango_prot_guardar = { 'protectedRangeId': str(rango_prot['protectedRangeId']) }
                
                if 'editors' in rango_prot.keys():
                    rango_prot_guardar['usersEdit'] = rango_prot['editors']['users']
                    if 'groups' in rango_prot['editors'].keys():
                        rango_prot_guardar['groupsEdit'] = rango_prot['editors']['groups']
                
                rango_prot_guardar['range'] = gshTraducirRangoInverso(rango_prot['range'])
                
                rangos_protegidos[nom_hoja].append(rango_prot_guardar)
                
                
    return rangos_protegidos           
                

                
                

def gshProtegerRango(sheets_service, spreadsheetId, nombreHoja, usuariosEdicion, rango=None, idRango=None, nombreRango=None, gruposEdicion = None):
    """
    Esta función protege uno o varios rangos de un libro. Si no se proporciona el ID del rango protegido, crea un nuevo rango.
    Si se proporciona el ID actualiza el rango con ese ID con los datos proporcionados.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str or list of str
        Nombre de la hoja o listado de hojas en la que se desea proteger el rango dentro del libro.
        Si es de tipo str se protegerán todos los rangos dentro de la misma hoja. Si es una lista, se emparejan las nombreHoja-rango una a una.
    usuariosEdicion : list of str
        Listado con los mails de usuarios que pueden editar los rangos protegidos. Al menos deberá estar el mail del usuario que 
        está protegiendo el rango. Si se quiere que cada rango creado sea editable por distintos usuarios hay que llamar a la función varias veces.
    rango : str o list of str, optional
        Rango o lista de rangos a proteger. Si nombreHoja es una lista, el parámetro "rango" tiene que ser una lista de la misma longitud.
        Si el rango es None se protege toda la hoja. (Si nombreHoja es una lista, rango puede ser una lista de [None, None...] )
        Ejemplo: 'A1:C20' o bien ['B2:C50', 'J2:W37']. 
        The default is None.
    idRango : str or list of str, optional
        Si el rango protegido ya existe y se quiere editar, su ID. Si se proporciona este parámetro debe tener el mismo tipo que el parámetro
        "rango". Si se proporcionan listas, ambas deben tener la misma longitud. Si se quiere crear un rango nuevo, su ID será None. 
        Ejemplo: si rango=['B2:C50', 'J2:W37'], idRango=None o bien idRango=['1577910399', None].
        The default is None.
    gruposEdicion : list, optional    
        Listado con los grupos de google que pueden editar los rangos protegidos. The default is None.

    Returns
    -------
    list
        Lista con los IDs de los rangos protegidos que se han creado/actualizado.

    """
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    hojas_docu = { propiedades['properties']['title'] : propiedades['properties']['sheetId'] for propiedades in sheet_metadata['sheets']}
    
    #Comprobamos los usuarios de edición
    if type(usuariosEdicion) != list:
        print("ERROR: los usuarios de edición deben ser una LISTA de direcciones de mail")
        return None
    
    if gruposEdicion is not None and type(gruposEdicion)!= list:
        print("ERROR: Se han proporcionado grupos de edición pero no son una lista. Si se proporciona, debe ser una lista de grupos google")
        return None
    
    #Comprobamos los parámetros de nombreHoja y rango
    if type(nombreHoja) == str and type(rango)==list :
        nombresHojas = [nombreHoja for rg in rango]
        rangos = [ gshTraducirRango(rangoLetras=rg, sheetId=hojas_docu[nombreHoja] ) if rg is not None else {'sheetId': hojas_docu[nombreHoja] } for rg in rango]
        
    elif type(nombreHoja) == str and rango is None:
        nombresHojas = [nombreHoja]
        rangos = [ {'sheetId': hojas_docu[nombreHoja] } ]
        
    elif type(nombreHoja) == str and type(rango) == str:
        nombresHojas = [nombreHoja]
        rangos = [ gshTraducirRango(rangoLetras=rango, sheetId=hojas_docu[nombreHoja] ) ]
    
    elif type(nombreHoja) == list and type(rango) == list :
        nombresHojas = nombreHoja
        
        if len(rango)!=len(nombresHojas): 
            print("ERROR: No se han proporcionado tantos rangos como hojas ")
            return None
        
        rangos = []
        for indice, rg in enumerate(rango):
            if rg is None:
                rangos.append({'sheetId': hojas_docu[nombresHojas[indice]] })
            else:
                rangos.append(gshTraducirRango(rangoLetras=rg, sheetId=hojas_docu[nombresHojas[indice]]))  
    
    elif type(nombreHoja) == list and rango is None:
        print("ERROR: Se ha especificado una lista de hojas pero el rango es None")
        return None

    else:
        print("ERROR: Se ha proporcionado una lista de hojas pero un único rango")
        return None
      
        
    #Ahora comprobamos, si nos lo han proporcionado, el parámetro de idRango
    if idRango is not None:
        if type(rango) == str and type(idRango)== str:
            idsRangos = [idRango]
        elif type(rango) == list and type(idRango)==list and len(rango)==len(idRango):
            idsRangos = idRango
        else:
            print("ERROR: Se han proporcionado idRangos pero no tantos como rangos")
            return None
    else:
        idsRangos = [None for rg in rangos]
     
    #Ahora recorremos la lista de hojas y construimos las requests
    lista_requests=[]
    
    for indice, rg in enumerate(rangos):
        #Seleccionar el tipo de request de que se trata (update o nueva)
        if idsRangos[indice] is None:
            tipoReq = "addProtectedRange"
            request = { 
                "addProtectedRange": {
                    "protectedRange": {}
                }
            }
        
        else :
            tipoReq = "updateProtectedRange"
            request = {
                "updateProtectedRange" : {
                    "protectedRange": {
                        "protectedRangeId" : idsRangos[indice]
                        },
                    
                    "fields": "range,warningOnly,editors"
                    }
                }
            
        #Terminar de construir la request para este rango en concreto:    
        request[tipoReq]["protectedRange"]["range"] = rg
        request[tipoReq]["protectedRange"]["warningOnly"] = False
        request[tipoReq]["protectedRange"]["editors"] = { "users" : usuariosEdicion }
        request[tipoReq]["protectedRange"]["editors"]["domainUsersCanEdit"] = False
        if gruposEdicion is not None:
            request[tipoReq]["protectedRange"]["editors"]["groups"] = gruposEdicion
            
        
        lista_requests.append(request)
     
    
    #Ahora realmente ejecutamos la solicitud.
    try:
        body = {'requests': lista_requests}
        respuesta = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=body).execute()
    except Exception as e:
        print("Error protegiendo rango:", e)
        return None
    
    salida = []
    for indice, resp in enumerate(respuesta['replies']):
        if "addProtectedRange" in resp.keys():
            salida.append(    str(resp["addProtectedRange"]['protectedRange']['protectedRangeId'])    )
            
        else: #Si era un update no devuelve nada, por eso cogemos directamente de los IDs
            salida.append(idsRangos[indice])
            
            
    return salida
    

#%% 

def gshEliminarRangoProtegido(sheets_service, spreadsheetId, idRango):
    """
    Elimina uno varios rangos protegidos de un libro.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    idRango : str o list
        ID o lista de IDs de los rangos protegidos que se quieren eliminar

    Returns
    -------
    bool
        True si todo ha ido bien, False en caso de error

    """
    lista_requests = []
    
    if type(idRango) == str:
        idRango = [idRango]
        
    for idRg in idRango:
        req = { "deleteProtectedRange" : {"protectedRangeId": idRg }  }
        lista_requests.append(req)
        
    try: 
        body = {'requests': lista_requests}
        respuesta = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=body).execute()
        #print(respuesta)
    except Exception as e:
        print("Error eliminando rangos protegidos:", e)
        return False
    
    return True
    
    







#%%

def gshTraducirColor(color) :
    """
    Recibe un color RGB y lo traduce a su representación JSON

    Parameters
    ----------
    color : str o dict
        Si recibe un str debe ser una representación RGB del color. Ejemplo: '#A359FF'.
        Si recibe un dict debe tener las claves 'red', 'green', 'blue' con valores numéricos. Opcionalmente
        puede tener la clave 'alpha' también con valor numérico <=255

    Returns
    -------
    resul_color : TYPE
        DESCRIPTION.

    """
    if type(color) == str and color.startswith('#') and len(color)==7:
        resul_color = {'red'  : int(color[1:3],16)/255, 
                       'green': int(color[3:5],16)/255, 
                       'blue' : int(color[5:] ,16)/255 }
        
    elif type(color) == dict and 'red' in color.keys() and 'green' in color.keys() and 'blue' in color.keys():
        resul_color = color
        for clave in resul_color.keys():
            if resul_color[clave]>1:
                resul_color[clave] = resul_color[clave]/255
    
    else:
       resul_color = color
       
       
    return resul_color

#%%



def gshTraducirRango(rangoLetras, sheetId):
    """
    Recibe un rango con el formato habitual de una sheet (Ejemplo: 'A25:J48') y lo traduce a un
    objeto de tipo GridRange de la API de GSheets.
    
    INPUT:
    - rangoLetras = str, Rango a traducir. Ejemplo: 'A25:J48'.
    - sheetId = str, Identificador de la hoja dentro del libro
    
    OUTPUT:
    - representación JSON de un objeto GridRange. None en caso de error.
    """
    import re
    
    rango = {'sheetId' : sheetId}
    
    partesRango = rangoLetras.split(':')
    if len(partesRango)>2: #si ha metido varios ":"
        print("El rango introducido no es correcto")
        return None
    
    
    if len(partesRango)== 1: #es sólo 1 celda
        celda = re.split(pattern='(\d+)',string=rangoLetras)
        if len(celda) != 3 or celda[2]!= '':
            print("Celda con formato incorrecto")
            return None
        
        elif not celda[0].isalpha() or not celda[1].isdigit():
            print("Celda con formato incorrecto")
            return None
        
        else:
            columna = 0
            for indice, letra in enumerate( celda[0][::-1] ) :
                columna += (ord(letra.upper()) - ord('A') +1 ) * (26 ** indice)
            
            rango['startColumnIndex'] = columna - 1
            rango['endColumnIndex'] = columna 
            rango['startRowIndex'] = int(celda[1]) - 1
            rango['endRowIndex'] = int(celda[1]) 
            return rango
            
     
    #Si llegamos hasta aquí es porque el rango es de tipo PARTE1:PARTE2
    #Para cada parte vamos a calcular su columna y fila de inicio o fin respectivamente
     
    #PARTE1    
    celda_ini = re.split(pattern='(\d+)',string=partesRango[0])
    
    #si sólo viene la columna
    if len(celda_ini) == 1 and celda_ini[0].isalpha() : 
        columna = 0
        for indice, letra in enumerate( celda_ini[0][::-1] ) :
            columna += (ord(letra.upper()) - ord('A') +1 ) * (26 ** indice)
        
        rango['startColumnIndex'] = columna - 1
        
    #si viene columna y fila
    elif len(celda_ini)==3 and celda_ini[1].isdigit() and celda_ini[2]=='' and celda_ini[0].isalpha():
        columna = 0
        for indice, letra in enumerate( celda_ini[0][::-1] ) :
            columna += (ord(letra.upper()) - ord('A') +1 ) * (26 ** indice)
        
        rango['startColumnIndex'] = columna - 1
        rango['startRowIndex'] = int(celda_ini[1]) - 1
        
    #si sólo viene la fila
    elif len(celda_ini)==3 and celda_ini[1].isdigit() and celda_ini[2]=='' and celda_ini[0]== '':
        rango['startRowIndex'] = int(celda_ini[1]) - 1
        
    else:
        print("Primera parte del rango incorrecta")
        return None
    
    
    
    #PARTE2
    celda_fin = re.split(pattern='(\d+)',string=partesRango[1])
    
    #si sólo viene la columna
    if len(celda_fin) == 1 and celda_fin[0].isalpha() : 
        columna = 0
        for indice, letra in enumerate( celda_fin[0][::-1] ) :
            columna += (ord(letra.upper()) - ord('A') +1 ) * (26 ** indice)
        
        rango['endColumnIndex'] = columna 
        
    #si viene columna y fila
    elif len(celda_fin)==3 and celda_fin[1].isdigit() and celda_fin[2]=='' and celda_fin[0].isalpha():
        columna = 0
        for indice, letra in enumerate( celda_fin[0][::-1] ) :
            columna += (ord(letra.upper()) - ord('A') +1 ) * (26 ** indice)
        
        rango['endColumnIndex'] = columna 
        rango['endRowIndex'] = int(celda_fin[1]) 
        
    #si sólo viene la fila
    elif len(celda_fin)==3 and celda_fin[1].isdigit() and celda_fin[2]=='' and celda_fin[0]== '':
        rango['endRowIndex'] = int(celda_fin[1]) 
        
    else:
        print("Segunda parte del rango incorrecta")
        return None
    
      
    return rango
    
    



#%%

def gshTraducirRangoInverso (rangoJSON):
    """
    Recibe una representación JSON de un rango y devuelve el rango con el formato de una sheet (Ejemplo: 'A25:J48')

    Parameters
    ----------
    rangoJSON : dict
        Representación JSON de un rango. Suele tener las claves 'startColumnIndex', 'endColumnIndex', 'startRowIndex' y 'endRowIndex'.

    Returns
    -------
    str
        El rango en el formato habitual de una sheet. None en caso de que no haya ninguna clave en el rango.

    """
    rango_letras_ini = ''
    if 'startColumnIndex' in rangoJSON.keys():
        columna = rangoJSON['startColumnIndex'] + 1
        while columna > 0:
            columna, resto = divmod(columna - 1, 26)
            rango_letras_ini = chr(ord('A') + resto) + rango_letras_ini
    
    if 'startRowIndex' in rangoJSON.keys():
        rango_letras_ini = rango_letras_ini + str(   rangoJSON['startRowIndex'] + 1   )
    
        
    
    rango_letras_fin = ''
    if 'endColumnIndex' in rangoJSON.keys():
        columna = rangoJSON['endColumnIndex']
        while columna > 0:
            columna, resto = divmod(columna - 1, 26)
            rango_letras_fin = chr(ord('A') + resto) + rango_letras_fin
    
    if 'endRowIndex' in rangoJSON.keys():
        rango_letras_fin = rango_letras_fin + str(   rangoJSON['endRowIndex']   )
    
    
    
    if len(rango_letras_ini) == 0 and len(rango_letras_fin)==0:
        return None
    
    #Si coinciden, es que en realidad es sólo una celda
    if rango_letras_ini == rango_letras_fin and rango_letras_ini[-1].isdigit():
        return rango_letras_ini
    
    return  rango_letras_ini + ':' + rango_letras_fin


    

#%%


def gshFormatearHoja(sheets_service, spreadsheetId, nombreHoja, rangoCabecera = '1:1'):
    """
    Da formato corporativo básico a una hoja de una spreadsheet. Pone la cabecera con fondo azul, letras
    blancas y en negrita. Inmoviliza la cabecera. Auto-ajusta el ancho de las columnas.
    Devuelve True si se ejecuta bien, False si hay un error. 
    
    INPUT:
    - sheets_service = Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    - spreadsheetId = str, Identificador del libro en Google Drive.
    - nombreHoja = str, Nombre de la hoja en la que se desea escribir dentro del libro.
    - rangoCabecera = str, Rango donde se encuentra la cabecera. Por defecto, fila 1 de la hoja completa '1:1'.
    
    OUTPUT:
    - True si todo ha ido correctamente, False en caso de error.
    """
    #Localizamos la hoja pedida y capturamos su ID y el número de columnas
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    
    sheetId = None
    for sheet in sheets:
        if sheet['properties']['title'] == nombreHoja :
            sheetId = sheet['properties']['sheetId']
            numCols = sheet['properties']['gridProperties']['columnCount']
            break
    
    if sheetId is None:
        print("La hoja con el nombre solicitado no existe")
        return False
    
    rango_cabecera_req = gshTraducirRango(rangoLetras=rangoCabecera, sheetId=sheetId)
    if rango_cabecera_req is None or 'startRowIndex' not in rango_cabecera_req.keys() or 'endRowIndex' not in rango_cabecera_req.keys():
        print("No se ha proporcionado un rango correcto para la cabecera")
        return False    
    
    cols_autoajustar_req = {"sheetId": sheetId, "dimension": 'COLUMNS'} 
    if 'startColumnIndex' in rango_cabecera_req.keys():
        cols_autoajustar_req["startIndex"] = rango_cabecera_req['startColumnIndex']
    else :
        cols_autoajustar_req["startIndex"] = 0
    
    if 'endColumnIndex' in rango_cabecera_req.keys():
        cols_autoajustar_req["endIndex"] = rango_cabecera_req['endColumnIndex']
    else :
        cols_autoajustar_req["endIndex"] = numCols

        
    request_formato = { 'requests' : [
        { 'repeatCell' : {
            'range' : rango_cabecera_req,            
            'cell' : {
                'userEnteredFormat': {
                    'backgroundColor' : {
                        'red' : 7/255,
                        'green' : 33/255,
                        'blue' : 70/255
                        },
                    'textFormat' : {
                        'bold' : True, 
                        'foregroundColor' : {
                            'red' : 1,
                            'green' : 1,
                            'blue' : 1
                            }
                        }
                    } #end UserEnteredFormat
                }, #end Cell            
            'fields' : 'userEnteredFormat(backgroundColor,  textFormat)'
            }
        }, #end RepeatCell      
        
       { 'autoResizeDimensions': {
           "dimensions": cols_autoajustar_req 
           }
        }, #end autoResizeDimensions
       
       { 'updateSheetProperties' :{
            'properties' : {
                "sheetId": sheetId,
                "gridProperties": {
                    "frozenRowCount": rango_cabecera_req['endRowIndex']
                    }
                },
            'fields' : 'gridProperties(frozenRowCount)'
            }
        } #end updateSheetProperties        
        ] }        
    
    try:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=request_formato).execute()
    except Exception as e:
        print("Error formateando la hoja:", e)
        return False
    
    return True






#%%

def gshEliminarFormatoRango (sheets_service, spreadsheetId, nombreHoja, rango, idHoja = None):
    """
    Elimina el formato de un rango de celdas concreto, dejando el formato por defecto.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desea escribir dentro del libro.
    rango : str 
        Rango a aplicar el formato con notación de la sheet. Ejemplo: 'C6:O12'
    idHoja : str, optional
        Id de la hoja dentro de la spreadsheet. Si se especifica este parámetro se ignora el 'nombreHoja'. The default is None.

    Returns
    -------
    bool
        True si ha funcionado correctamente, False en caso contrario.

    """
    if idHoja is None:
        idHoja = gshObtenerIDHoja(sheets_service, spreadsheetId, nombreHoja) 
        
    if idHoja is None:
        print("No se puede formatear la hoja")
        return False
    
    rango_req = gshTraducirRango(rangoLetras=rango, sheetId=idHoja)
    if rango_req is None:
        print("No se ha proporcionado un rango correcto para el formato")
        return False 
    
    request_formato = { 'requests' : [
        { 'repeatCell' : {
            'range' : rango_req,            
            'cell' : {
                'userEnteredFormat': None
                }, #end Cell            
            'fields' : 'userEnteredFormat'
            }
        } #end RepeatCell      
        ] }  
    
    try:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=request_formato).execute()
    except Exception as e:
        print("Error formateando la hoja:", e)
        return False
    
    return True
    

def gshFormatearRango(sheets_service, spreadsheetId, nombreHoja, rango, idHoja = None,
                      formatoNumero = None, 
                      colorFondo = None,
                      colorTexto = None, 
                      negrita    = None,
                      cursiva    = None,
                      subrayado  = None,
                      tachado    = None,
                      tamLetra   = None,
                      tipoLetra  = None
                      ):
    """
    Da el formato especificado a un rango de celdas de una google sheet. 
    
    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desea escribir dentro del libro.
    rango : str 
        Rango a aplicar el formato con notación de la sheet. Ejemplo: 'C6:O12'
    idHoja : str, default None
        Id de la hoja dentro de la spreadsheet. Si se especifica este parámetro se ignora el 'nombreHoja'. 
    formatoNumero : str, default None
        Formato a aplicar. Puede ser una de las claves de FORMATOS_BASE o un 
        formato de sheet. Ejemplos: 'num_sep_mil_2dec' o '#,##0.00_);[Red]-#,##0.00'
    colorFondo : str, dict, default None
        Color de fondo de las celdas. Puede ser un color RGB. Ejemplo: '#FFFFCC' (no case sensitive)
        También puede ser un diccionario con las claves 'red', 'green', 'blue' y opcionalmente 'alpha'.
        En caso de ser un diccionario, los valores deben ser numéricos. Pueden ser enteros entre 0 y 255 o 
        decimales entre 0 y 1.
    colorTexto : str, dict, default None
        Color del texto de las celdas. Puede ser un color RGB. Ejemplo: '#ff0066' (no case sensitive)
        También puede ser un diccionario con las claves 'red', 'green', 'blue' y opcionalmente 'alpha'.
        En caso de ser un diccionario, los valores deben ser numéricos. Pueden ser enteros entre 0 y 255 o 
        decimales entre 0 y 1.
    negrita : bool, default None
        Indica si el texto se debe poner en negrita o no
    cursiva : bool, default None
        Indica si el texto se debe poner en cursiva o no
    subrayado : bool, default None
        Indica si el texto se debe subrayar o no
    tachado : bool, default None
        Indica si el texto debe ir tachado o no
    tamLetra : int, default None
        Tamaño de letra en el rango afectado
    tipoLetra : str, default None
        Nombre del tipo de letra a utilizar. Es case sensitive. Ejemplo: 'Arial'        
    
    Returns
    -------
    Boolean
        True si todo ha ido correctamente, False en caso de error.
    """   
    
    if idHoja is None:
        idHoja = gshObtenerIDHoja(sheets_service, spreadsheetId, nombreHoja) 
        
    if idHoja is None:
        print("No se puede formatear la hoja")
        return False
    
    rango_req = gshTraducirRango(rangoLetras=rango, sheetId=idHoja)
    if rango_req is None:
        print("No se ha proporcionado un rango correcto para el formato")
        return False    
    
    
    userEnteredFormat={}
    fields = set()
    
    #Formato del número
    if formatoNumero is not None:
        fields.add("numberFormat")
        if formatoNumero in FORMATOS_BASE.keys():
            userEnteredFormat["numberFormat"] = FORMATOS_BASE[formatoNumero]
        else:
            userEnteredFormat["numberFormat"] = formatoNumero
    
    #Color de fondo
    if colorFondo is not None:
        fields.add("backgroundColor")
        userEnteredFormat["backgroundColor"] = gshTraducirColor(colorFondo)
        
            
    #Formato del texto
    textFormat = {}            
    if colorTexto is not None:
        fields.add("textFormat")
        textFormat["foregroundColor"] = gshTraducirColor(colorTexto)
            
    if negrita is not None and type(negrita) == bool:
        fields.add("textFormat")
        textFormat['bold'] = negrita
        
    if cursiva is not None and type(cursiva) == bool:
        fields.add("textFormat")
        textFormat['italic'] = cursiva
        
    if subrayado is not None and type(subrayado) == bool:
        fields.add("textFormat")
        textFormat['underline'] = subrayado
        
    if tachado is not None and type(tachado) == bool:
        fields.add("textFormat")
        textFormat['strikethrough'] = tachado
        
    if tamLetra is not None and type(tamLetra) == int:
        fields.add("textFormat")
        textFormat['fontSize'] = tamLetra
        
    if tipoLetra is not None and type(tipoLetra) == str:
        fields.add("textFormat")
        textFormat['fontFamily'] = tipoLetra
    
    #Una vez que hemos preparado el formateo del texto, lo incluimos en la petición:
    if len(textFormat)>0:
        userEnteredFormat["textFormat"] = textFormat
        
    
    request_formato = { 'requests' : [
        { 'repeatCell' : {
            'range' : rango_req,            
            'cell' : {
                'userEnteredFormat': userEnteredFormat
                }, #end Cell            
            'fields' : 'userEnteredFormat('+ ",".join(fields) +')'
            }
        } #end RepeatCell      
        ] }  
    
    try:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=request_formato).execute()
    except Exception as e:
        print("Error formateando la hoja:", e)
        return False
    
    return True


#%%

def gshEliminarFormatosCondicionalesHoja(sheets_service, spreadsheetId, nombreHoja):
    """
    Elimina todas las reglas de formato condicional de una hoja

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desean borrar todos los formatos condicionales.

    Returns
    -------
    bool
        True si ha ido todo bien, False en caso de error.

    """
    #Localizamos las propiedades de todas las hojas
    sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    sheet_names = [sheet['properties']['title'] for sheet in sheets]
    
    #Comprobamos si la hoja existe. En caso contrario, se devuelve None:
    if nombreHoja not in sheet_names:
        return False
    
    #Comprobamos en qué posición está la hoja solicitada. No va a fallar porque en el if anterior aseguramos que está.
    indice_hoja = sheet_names.index(nombreHoja)
    id_hoja = sheet_metadata['sheets'][indice_hoja]['properties']['sheetId']
    
        
    #Si esa hoja en particular no tiene formatos condicionales no hay que hacer nada.
    if 'conditionalFormats' not in sheet_metadata['sheets'][indice_hoja].keys():
        return True
    
    reglas_condic = sheet_metadata['sheets'][indice_hoja]['conditionalFormats']
    
    requests= []
    for indice, regla in enumerate(reglas_condic):
        req = {
            "deleteConditionalFormatRule": {
                "index": indice,
                "sheetId": id_hoja
                }
             }
        requests.append(req)
        
    requests.reverse()
    request_final = { 'requests' : requests}  
    #print(request_final)
    
    try:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=request_final).execute()
    except Exception as e:
        print("Error borrando formatos condicionales de la hoja:", e)
        return False
    
    return True
    


def gshFormatoCondicionalRango(sheets_service, spreadsheetId, nombreHoja, reglas_formatos, idHoja=None ):
    """
    Genera las reglas de formato condicional en un conjunto de rangos
    
    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desea escribir dentro del libro.
    reglas_formatos : list 
        Lista de reglas con sus formatos. Cada regla es un diccionario con las entradas 'rango', 'tipo_regla', 'valores' y 'formato'. 
        A su vez la entrada 'formato' es un diccionario que admite las claves 'colorFondo', 'colorTexto', 'negrita', 'cursiva' y 'tachado'.
        Si hay más de una regla los rangos se pueden solapar. Las reglas se aplican en el orden en que aparecen
        Ejemplo: 
            [
                 {
                    'rango' : 'A3:B10', 
                    'tipo_regla' : 'CUSTOM_FORMULA', 
                    'valores' : ['=$AC3="Recup"']  ,
                    'formato' : {
                        'colorFondo' : '#b6d7a8' 
                        }
                } 
                 ,
                {
                    'rango' : 'A3:C20', 
                    'tipo_regla' : 'NUMBER_BETWEEN', 
                    'valores' : ['1','7']  ,
                    'formato' : {
                        'colorFondo' : '#d5abf5',
                        'negrita'    : True ,
                        'cursiva'    : True , 
                        'colorTexto' : '#ffffff'
                        }
                }     
                  
              ]
    idHoja : str, optional
        Id de la hoja dentro de la spreadsheet. Si se especifica este parámetro se ignora el 'nombreHoja'. The default is None.       
    
    Returns
    -------
    Boolean
        True si todo ha ido correctamente, False en caso de error.
    """   
    
    if type(reglas_formatos) != list:
        print("las reglas-formato deben ser una lista de diccionarios")
        print("""Ejemplo:
              [
                 {
                    'rango' : 'A3:B10', 
                    'tipo_regla' : 'CUSTOM_FORMULA', 
                    'valores' : ['=$AC3="Recup"']  ,
                    'formato' : {
                        'colorFondo' : '#b6d7a8' 
                        }
                } 
                 ,
                {
                    'rango' : 'A3:C20', 
                    'tipo_regla' : 'NUMBER_BETWEEN', 
                    'valores' : ['1','7']  ,
                    'formato' : {
                        'colorFondo' : '#d5abf5',
                        'negrita'    : True ,
                        'cursiva'    : True , 
                        'colorTexto' : '#ffffff'
                        }
                }     
                  
              ]
              
              """)
    
    if idHoja is None:
        idHoja = gshObtenerIDHoja(sheets_service, spreadsheetId, nombreHoja) 
        
    if idHoja is None:
        print("No se puede dar el formato condicional.")
        return False
    
    
    
    requests = []
    for indice, regla_form in enumerate(reglas_formatos):
        rango = gshTraducirRango(rangoLetras=regla_form['rango'], sheetId=idHoja)
        
        formato_peticion = {}
        if 'colorFondo' in regla_form['formato'].keys():
            colorFondo = regla_form['formato']['colorFondo']
            formato_peticion['backgroundColor'] = gshTraducirColor(colorFondo)
            
        if any([opc in regla_form['formato'].keys() for opc in ['colorTexto', 'negrita', 'cursiva', 'tachado']] ):
            formato_peticion['textFormat'] = {}
        
        if 'colorTexto' in regla_form['formato'].keys():
            colorTexto = regla_form['formato']['colorFondo']
            formato_peticion['textFormat']['foregroundColor'] = gshTraducirColor(colorTexto)
            
        if 'negrita' in regla_form['formato'].keys():
            formato_peticion['textFormat']['bold'] = regla_form['formato']['negrita']
            
        if 'cursiva' in regla_form['formato'].keys():
            formato_peticion['textFormat']['italic'] = regla_form['formato']['cursiva']
            
        if 'tachado' in regla_form['formato'].keys():
            formato_peticion['textFormat']['strikethrough'] = regla_form['formato']['tachado']
        
        peticion = { "addConditionalFormatRule": {
            "rule": {
                "ranges": [ rango ],
                
                "booleanRule": {
                    "condition": {
                        "type": regla_form['tipo_regla'],
                        "values": [ {"userEnteredValue":valor} for valor in regla_form['valores'] ]
                        },
                    "format": formato_peticion
                    }
                },
            "index": indice
            }
            }
        requests.append(peticion)
    
    
    
        
    
    request_formato = { 'requests' : requests}  
    #print(request_formato)
    
    try:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=request_formato).execute()
    except Exception as e:
        print("Error formateando la hoja:", e)
        return False
    
    return True






#%%


def gshEliminarValidacionRango(sheets_service, spreadsheetId, nombreHoja, rango, idHoja = None):
    """
    Elimina todas las fórmulas de validación del rango proporcionado.

    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desea escribir dentro del libro.
    rango : str 
        Rango donde eliminar las validaciones con notación de la sheet. Ejemplo: 'C6:O12'
    idHoja : str, optional
        Id de la hoja dentro de la spreadsheet. Si se especifica este parámetro se ignora el 'nombreHoja'. The default is None.

    Returns
    -------
    bool
        True si todo ha ido correctamente, False en otro caso.

    """
    if idHoja is None:
        idHoja = gshObtenerIDHoja(sheets_service, spreadsheetId, nombreHoja) 
        
    if idHoja is None:
        print("No se puede eliminar la validacion")
        return False
    
    rango_req = gshTraducirRango(rangoLetras=rango, sheetId=idHoja)
    if rango_req is None:
        print("No se ha proporcionado un rango correcto para el formato")
        return False
    
    request_validacion = { 'requests' : [
        { 'repeatCell' : {
            'range' : rango_req,            
            'cell' : {
                'dataValidation': None
                }, #end Cell            
            'fields' : 'dataValidation'
            }
        } #end RepeatCell      
        ] }  
    
    try:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=request_validacion).execute()
    except Exception as e:
        print("Error formateando la hoja:", e)
        return False
    return True



def gshValidacionDesplegableRango(sheets_service, spreadsheetId, nombreHoja, rango, valoresValidacion, idHoja = None):
    """
    Hace que un rango sólo pueda tener los valores especificados en valoresValidacion y se
    vean como un combo desplegable del que sólo se puede seleccionar uno de esos valores. 
    
    Parameters
    ----------
    sheets_service : service object
        Servicio abierto con googlesheets, que debe haber devuelto la función connect previamente.
    spreadsheetId : str
        Identificador del libro en Google Drive.
    nombreHoja : str
        Nombre de la hoja en la que se desea escribir dentro del libro.
    rango : str 
        Rango donde aplicar la validación con notación de la sheet. Ejemplo: 'C6:O12'
    valoresValidacion : list
        Lista con los posibles valores que puede tomar la celda. Debe contener al menos un elemento.
    idHoja : str, optional
        El id de la hoja que se quiere utilizar. Si se pasa se ignora el nombre de la hoja. The default is None.
    
    Returns
    -------
    Boolean
        True si todo ha ido correctamente, False en caso de error.
    """  
    
    if idHoja is None:
        idHoja = gshObtenerIDHoja(sheets_service, spreadsheetId, nombreHoja) 
        
    if idHoja is None:
        print("No se generar la validación en la hoja")
        return False
    
    rango_req = gshTraducirRango(rangoLetras=rango, sheetId=idHoja)
    if rango_req is None:
        print("No se ha proporcionado un rango correcto para la validación")
        return False    
    
    if type(valoresValidacion) != list or len(valoresValidacion)<1:
        print("No se han proporcionado los valores disponibles correctamente")
        return False
    
    values_req=[]
    for valor in valoresValidacion:
        values_req.append({"userEnteredValue": valor})
    
    
    validacion = {"condition": {
            "type":'ONE_OF_LIST',
            "values": values_req,
        },
        #"inputMessage": string,
        "strict": True,
        "showCustomUi": True
        }
    
    request_validacion = { 'requests' : [
        { 'repeatCell' : {
            'range' : rango_req,            
            'cell' : {
                'dataValidation': validacion
                }, #end Cell            
            'fields' : 'dataValidation'
            }
        } #end RepeatCell      
        ] }  
    
    try:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=request_validacion).execute()
    except Exception as e:
        print("Error formateando la hoja:", e)
        return False
    return True






######################################################################################################
# GOOGLE DOCS
######################################################################################################
def gdcCrearDoc(docs_service, titulo):
    """
    Crea un documento de Google Docs.
    docs_service -- Servicio de google con permiso de escritura en Google Docs. Se obtiene con el método connect
    titulo       -- (str) Nombre del documento.
    
    Más información en https://developers.google.com/docs/api/how-tos/documents
    """
    body = {
            'title': titulo
    }
    doc = docs_service.documents().create(body=body).execute()
    return doc
#%%
def gdcFindTag(docs_service, documentId, campo):
    """
    Encuentra un texto marcado entre llaves dentro de un documento de google docs.
    docs_service -- Servicio de google con permisos para escribir en Google Docs. Se obtiene con el método connect.
    documentId   -- (str) identificador del documento que se quiere editar.
    campo        -- (str) etiqueta buscada. Los campos dentro del documento deben estar rodeados de doble llave {{campo}}
    
    return       -- Tupla (initPos, endPos) si se encuentra el campo. None si no lo encuentra
    """
    
    posicion = None
    if not (campo.startswith('{{') and campo.endswith('}}')):
        campo = '{{' + campo + '}}'
    document = docs_service.documents().get(documentId=documentId).execute()
    content = document['body']['content']
    found=False
    for elem in content:
        if 'paragraph' in elem.keys():
            parrafo = elem['paragraph']['elements']
            for elemento in parrafo:
                if campo in elemento['textRun']['content']:
                    ini = elemento['startIndex'] + elemento['textRun']['content'].index(campo)
                    fin = ini + len(campo)
                    posicion = (ini, fin)
                    found = True
                    break
            if found:
                break
    return posicion
                
#%%
def gdcInsertarImagen(docs_service, documentId, imgurl, campo, alto, ancho, unidad='PT'):
    """
    Inserta una imagen vía url en un documento de google docs.
    docs_service -- Servicio de google con permisos para escribir en Google Docs. Se obtiene con el método connect.
    documentId   -- (str) identificador del documento que se quiere editar.
    imgurl       -- (str) direccion web de la imagen
    campo        -- (str) nombre del campo {{}} que sustituir por la imgen
    alto         -- (int) altura de la imagen
    ancho        -- (int) anchura de la imagen
    unidad       -- (str) default: 'PT' unidad de las dimensiones.
    
    Más detalles en https://developers.google.com/docs/api/how-tos/images
    """
    posicion = gdcFindTag(docs_service, documentId, campo)
    if posicion is None:
        print ('No se puede encontrar el campo')
        return None
    
    gdcReemplazarTexto(docs_service, documentId, [campo], [''])
    
    requests = [{
            'insertInlineImage': {
                'location': {'index': posicion[0]},
                'uri': imgurl,
                'objectSize': {
                    'height': {
                        'magnitude': alto,
                        'unit': unidad
                    },
                    'width': {
                        'magnitude': ancho,
                        'unit': unidad
                    }
                }
            }
    }]
    
    # Execute the request.
    body = {'requests': requests}
    response = docs_service.documents().batchUpdate(documentId=documentId, body=body).execute()
    insert_inline_image_response = response.get('replies')[0].get('insertInlineImage')
    return insert_inline_image_response

#%%
def gdcInsertarTexto(docs_service, documentId, textos, posiciones):
    """
    Inserta un texto en un documento de Google Docs.
    docs_service -- Servicio de google con permisos para escribir en Google Docs. Se obtiene con el método connect.
    documentId   -- (str) identificador del documento que se quiere editar.
    textos       -- (lista de str) ha de ser una variable que contenga un objeto iterable con todos los textos a insertar (str)
    posiciones   -- (lista de int) debe ser un objeto iterable, del mismo y posición que textos.
    
    Más detalles en https://developers.google.com/docs/api/how-tos/move-text
    """
    if len(textos) != len(posiciones):
        print('Las longitudes de textos y posiciones deben ser iguales')
        return None
    
    requests = []
    for i, texto in enumerate(textos):
        requests.append({
                'insertText': {
                        'location': {'index':posiciones[i]},
                        'text'    : texto
                }
        })

    result = docs_service.documents().batchUpdate(
        documentId=documentId, 
        body={'requests': requests}).execute()
    return result

#%%
def gdcReemplazarTexto(docs_service, documentId, campos, textos, envolverCampos=True):
    """
    Reemplaza un tag por un texto en un documento de google docs.
    docs_service   -- Servicio de google con permisos para escribir en Google Docs. Se obtiene con el método connect.
    documentId     -- (str) identificador del documento que se quiere editar.
    campos         -- (lista de str) con los campos que se deben buscar en el documento. Un campo debería estar encerrado entre doble llave. Ejemplo: {{campo}}
    textos         -- (lista de str) ha de ser una variable que contenga un objeto iterable con todos los textos a insertar (str)
    envolverCampos -- (boolean) default:True True si quieres añadir automáticamente las llaves a los campos en caso de que no tengan.
    
    Más detalles en https://developers.google.com/docs/api/how-tos/merge
    """
    if len(textos) != len(campos):
        print('Las longitudes de textos y campos deben ser iguales')
        return None
    
    if envolverCampos:
        for i,campo in enumerate(campos):
            if not (campo.startswith('{{') and campo.endswith('}}')):
                campos[i] = '{{' + campo + '}}'
    
    requests = []
    for i, texto in enumerate(textos):
        requests.append({
                'replaceAllText': {
                        'containsText': {
                                'text': campos[i],
                                'matchCase': 'true'
                        },
                        'replaceText' : texto
                }
        })

    result = docs_service.documents().batchUpdate(
        documentId=documentId, body={'requests': requests}).execute()
    return result

#%%
######################################################################################################
# GMAIL
######################################################################################################
def ggmCreateDraft(service, message, user_id='me'):
    """Crea un borrador.
        Args:
        service: Authorized Gmail API service instance.
        message: Message to be drafted. Se obtiene con las funciones ggmCreateMessage o ggmCreateMessageWithAttachment
        user_id: (str) User's email address. The special value "me"
        can be used to indicate the authenticated user.

        Returns:
        Drafted Message.
    """
    try:
        message = {'message':message}
        message = (service.users().drafts().create(userId=user_id, body=message).execute())
        print ('Message Id:', message['id'])
        return message
    except Exception as error:
        print ('An error occurred:', error)

#%%
def ggmCreateMessage(to, subject='', message_text='', sender='me', replyTo = None, html=False, cc=None):
    """Create a message for an email.

    Args:
    - to: (str) Dirección o direcciones de correo destino. En caso de haber varias, se separan por coma.
    - subject: (str) Opcional. Asunto del correo.
    - message_text: (str) Opcional. Cuerpo del correo.
    - sender: (str) Dirección del remitente (predeterminado 'me'). Se puede usar el valor especial 'me' para considerar que es la dirección propia. Si se desea utilizar un nombre distinto debe ir en formato "NOMBRE DEL REMITENTE <direccion_de_correo_propia@bbva.com>". La dirección de correo propia se puede obtener con la función ggmGetPrimaryAddress().
    - replyTo: (str) Opcional. Si se desea que al contestar el correo se dirija a una dirección diferente, indicarla aquí.
    - html: (bool) Opcional. Indica si el mensaje debe interpretarse como código html.
    - cc: (str) Opcional. Destinatario en copia. En caso de haber varios, se separan por coma.

    Returns:
    An object containing a base64url encoded email object.
    """
    import base64
    from email.mime.text import MIMEText
    
    if html:
        message = MIMEText(message_text, 'html')
    else:
        message = MIMEText(message_text)
    message['To'] = to
    message['From'] = str2ascii(sender)
    message['Subject'] = subject
    if replyTo:
        message.add_header('Reply-To', replyTo)
    if cc:
        message.add_header('Cc', str(cc))

    b64_bytes = base64.urlsafe_b64encode(message.as_bytes())
    b64_string = b64_bytes.decode()
    body = {'raw': b64_string}
    return body


#%%
def ggmCreateMessageWithAttachment(to, file_dir, filename, subject='', message_text='', sender='me', replyTo=None, html=False, cc=None):
    """Create a message for an email.

    Args:
    - to: (str) Dirección o direcciones de correo destino. En caso de haber varias, se separan por coma.
    - file_dir: The directory containing the file to be attached.
    - filename: The name of the file to be attached.
    - subject: (str) Opcional. Asunto del correo.
    - message_text: (str) Opcional. Cuerpo del correo.
    - sender: (str) Dirección del remitente (predeterminado 'me'). Se puede usar el valor especial 'me' para considerar que es la dirección propia. Si se desea utilizar un nombre distinto debe ir en formato "NOMBRE DEL REMITENTE <direccion_de_correo_propia@bbva.com>". La dirección de correo propia se puede obtener con la función ggmGetPrimaryAddress().
    - replyTo: (str) Opcional. Si se desea que al contestar el correo se dirija a una dirección diferente, indicarla aquí.
    - html: (bool) Opcional. Indica si el mensaje debe interpretarse como código html.
    - cc: (str) Opcional. Destinatario en copia. En caso de haber varios, se separan por coma.

    Returns:
    An object containing a base64url encoded email object.
    """
    import base64
    from email.mime.audio import MIMEAudio
    from email.mime.base import MIMEBase
    from email.mime.image import MIMEImage
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.encoders import encode_base64
    import mimetypes
    import os

    message = MIMEMultipart()
    message['To'] = str(to)
    message['From'] = str2ascii(sender)
    message['Subject'] = str(subject)
    if replyTo:
        message.add_header('Reply-To', str(replyTo))
    if cc:
        message.add_header('Cc', str(cc))

    if html:
        msg = MIMEText(message_text, 'html')
    else:
        msg = MIMEText(message_text)
    
    
    message.attach(msg)

    path = os.path.join(file_dir, filename)
    content_type, encoding = mimetypes.guess_type(path)

    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(path, 'rb')
        msg = MIMEText(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(path, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(path, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(path, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        encode_base64(msg)
        fp.close()

    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    b64_bytes = base64.urlsafe_b64encode(message.as_bytes())
    b64_string = b64_bytes.decode()
    body = {'raw': b64_string}
    return body

#%%
def ggmCreateMessageWithAttachments(to, ficheros, subject='', message_text='', sender='me', replyTo=None, html=False, cc=None):
    """Create a message for an email.

    Args:
    - to: (str) Dirección o direcciones de correo destino. En caso de haber varias, se separan por coma.
    - ficheros: (list) Lista que contiene los nombres (con su ruta, si procede) de los ficheros que se desean adjuntar.
    - subject: (str) Opcional. Asunto del correo.
    - message_text: (str) Opcional. Cuerpo del correo.
    - sender: (str) Dirección del remitente (predeterminado 'me'). Se puede usar el valor especial 'me' para considerar que es la dirección propia. Si se desea utilizar un nombre distinto debe ir en formato "NOMBRE DEL REMITENTE <direccion_de_correo_propia@bbva.com>". La dirección de correo propia se puede obtener con la función ggmGetPrimaryAddress().
    - replyTo: (str) Opcional. Si se desea que al contestar el correo se dirija a una dirección diferente, indicarla aquí.
    - html: (bool) Opcional. Indica si el mensaje debe interpretarse como código html.
    - cc: (str) Opcional. Destinatario en copia. En caso de haber varios, se separan por coma.

    Returns:
    An object containing a base64url encoded email object.
    """
    import base64
    from email.mime.audio import MIMEAudio
    from email.mime.base import MIMEBase
    from email.mime.image import MIMEImage
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.encoders import encode_base64
    import mimetypes
    import os

    message = MIMEMultipart()
    message['To'] = str(to)
    message['From'] = str2ascii(sender)
    message['Subject'] = str(subject)
    if replyTo:
        message.add_header('Reply-To', str(replyTo))
    if cc:
        message.add_header('Cc', str(cc))

    if html:
        msg = MIMEText(message_text, 'html')
    else:
        msg = MIMEText(message_text)
    
    
    message.attach(msg)

    for fichero in ficheros:
        nombre_fichero = os.path.split(fichero)[-1]
        content_type, encoding = mimetypes.guess_type(fichero)
    
        if content_type is None or encoding is not None:
            content_type = 'application/octet-stream'
        
        main_type, sub_type = content_type.split('/', 1)
        if main_type == 'text':
            fp = open(fichero, 'rb')
            msg = MIMEText(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'image':
            fp = open(fichero, 'rb')
            msg = MIMEImage(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'audio':
            fp = open(fichero, 'rb')
            msg = MIMEAudio(fp.read(), _subtype=sub_type)
            fp.close()
        else:
            fp = open(fichero, 'rb')
            msg = MIMEBase(main_type, sub_type)
            msg.set_payload(fp.read())
            encode_base64(msg)
            fp.close()
    
        msg.add_header('Content-Disposition', 'attachment', filename=nombre_fichero)
        message.attach(msg)

    b64_bytes = base64.urlsafe_b64encode(message.as_bytes())
    b64_string = b64_bytes.decode()
    body = {'raw': b64_string}
    return body

#%%
def ggmDescargarAdjuntos(gmail_service, mensaje=None, idmensaje=None, carpeta_descarga=None):
    """
        Recupera y descarga los adjuntos de un correo.
        Args:
            gmail_service: Authorized Gmail API service instance.
            mensaje: objeto Message con los detalles. Se le hace caso si idmensaje es None
            idmensaje: (str) identificador del mensaje con los detalles.
            carpeta_descarga: Directorio para descargar. Por defecto usa la carpeta de descargas del usuario.
        Devuelve una lista con la ruta completa de los ficheros descargados.
    """
    import base64
    import io
    import os
    from apiclient import errors

    downloaded = []
    
    try:
        if mensaje is None:
            if idmensaje is None:
                print('No se ha proporcionado ningún mensaje')
                return []
            else:
                mensaje = ggmLeerCorreo(gmail_service, idmensaje)
        parts = [mensaje['payload']]
        while parts:
            part = parts.pop()
            if part.get('parts'):
                parts.extend(part['parts'])
            if part.get('filename'):
                if 'data' in part['body']:
                    file_data = base64.urlsafe_b64decode(part['body']['data'].encode('UTF-8'))
                elif 'attachmentId' in part['body']:
                    attachment = gmail_service.users().messages().attachments().get(
                        userId='me', 
                        messageId=mensaje['id'], 
                        id=part['body']['attachmentId']
                    ).execute()
                    data = attachment['data']
                    file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                else:
                    file_data = None
                if file_data:
                    #do some staff, e.g.
                    if carpeta_descarga is None:
                        carpeta_descarga = os.environ['HOMEPATH'] + '/Downloads'
                    if not(carpeta_descarga.endswith('/') or carpeta_descarga.endswith('\\')):
                        carpeta_descarga = carpeta_descarga + '/'
                    path = ''.join([carpeta_descarga, part['filename']])
                    try:
                        with open(path, 'wb') as f:
                            f.write(file_data)
                        downloaded.append(path)
                        print('Descargado:',path)
                    except Exception as err:
                        print('No se puede descargar el fichero.',err)
    except errors.HttpError as error:
        print ('Error descargando adjuntos', error)
    return downloaded

#%%
def ggmGetPrimaryAddress(service, user_id = 'me'):
    """
    Función que devuelve la dirección de mail princial del servicio de correo.
    
    Parámetros:
        service  -- Objeto servicio de Gmail. Se obtiene con el método connect.
        user_id  -- Dirección de mail del usuario registrado en el servicio. 
                    El valor especial "me" indica el usuario autenticado
        
    Devuelve:
        Cadena de texto con la dirección de e-mail o None en caso de error.
    """
    primary_alias = None

    try:
        aliases = service.users().settings().sendAs().list(userId='me').execute()
        for alias in aliases.get('sendAs'):
            if alias.get('isPrimary'):
                primary_alias = alias
                break
    except Exception as error:
        print ('An error occurred:', error)
        return None
    
    return primary_alias.get('sendAsEmail')

#%%
def ggmLeerCorreo(gmail_service, idmensaje):
    """
        Recupera los detalles de un correo.
        Args:
            gmail_service: Authorized Gmail API service instance.
            idmensaje: (str) identificador del mensaje con los detalles.
    """
    try:
        mensaje = gmail_service.users().messages().get(userId='me', id = idmensaje).execute()
        return mensaje
    except Exception as err:
        print('Error al recuperar mensaje:', err)
        return None

#%%
def ggmListarCorreos(
    gmail_service, 
    maxmensajes=100, 
    buscar=None, 
    asunto=None, 
    de=None, 
    para=None, 
    cc=None, 
    bcc=None, 
    adjuntos=None,
    desde=None,
    hasta=None,
    etiquetas=None,
    recibidos=None,
    chat=None,
    includeSpamTrash=False
):
    """
    lista una serie de correos dados unos criterios de búsqueda.
    -- gmail_service : Servicio con acceso de lectura a gmail.
    -- maxmensajes   : (int) número de mensajes a recuperar
    -- buscar        : (str) Opcional. Cadena de búsqueda. Lo que se pondría en el buscador de gmail.
    -- asunto        : (str) Opcional. Asunto del mensaje.
    -- de            : (str) Opcional. Remitente.
    -- para          : (str) Opcional. Destinatario.
    -- cc            : (str) Opcional. En copia.
    -- bcc           : (str) Opcional. En copia oculta.
    -- adjuntos      : (bool, str) Opcional. Si True especifica que debe haber adjuntos. Si es un texto, especifica el nombre de archivo adjunto.
    -- desde         : (datetime, str) Opcional. Fecha desde la que buscar correo. Si es texto, en formato yyyy/mm/dd.
    -- hasta         : (datetime, str) Opcional. Fecha hasta la que buscar correo. Si es texto, en formato yyyy/mm/dd.
    -- etiquetas     : (str) Opcional. Etiquetas bajo las que buscar el correo.
    -- recibidos     : (bool) Opcional. True si debe buscar solo en recibidos.
    -- chat          : (bool) Opcional. True si debe buscar solo en chats.
    -- includeSpamTrash : (bool) Opcional. True si debe buscar también en las carpetas de spam y papelera.
    """
    import datetime
    
    def completa(campo, valor, par=True):
        if par:
            return campo + ':(' + valor + ') '
        else:
            return campo + ':' + valor + ' '
    
    q=''
    if buscar:
        q = q + buscar + ' '
    if asunto:
        q = q + completa('subject', asunto)
    if de:
        q = q + completa('from', de)
    if para:
        q = q + completa('to', para)
    if cc:
        q = q + completa('cc', cc)
    if bcc:
        q = q + completa('bcc', bcc)
    if desde:
        if type(desde)==datetime.date:
            desde = desde.strftime(format='%Y/%m/%d')
        q = q + completa('after', desde, par=False)
    if hasta:
        if type(hasta)==datetime.date:
            hasta = hasta.strftime(format='%Y/%m/%d')
        q = q + completa('before', hasta, par=False)
    if etiquetas:
        q = q + completa('label', etiquetas, par=False)
    if adjuntos:
        q = q + 'has:attachment '
        if type(adjuntos) == str:
            q = q + completa('filename', adjuntos)
    if recibidos:
        q = q + 'in:Inbox '
    if chat:
        q = q + 'is:chat '
    if includeSpamTrash:
        q = q + 'in:all '
    print(q)
    
    listamensajes = []
    pageToken = ''
    while maxmensajes > 0:
        try:
            mensajes = gmail_service.users().messages().list(userId='me', pageToken=pageToken, q=q).execute()
            if 'messages' not in mensajes.keys():
                return listamensajes
            nuevos = mensajes['messages'][0:min(len(mensajes['messages']),maxmensajes)]
            listamensajes = listamensajes + nuevos
            maxmensajes = maxmensajes - len(nuevos)
            if 'nextPageToken' in mensajes.keys():
                pageToken = mensajes['nextPageToken']
            else:
                break
        except Exception as err:
            print('Error',err)
            break            
    return listamensajes

#%%
def ggmSendMessage(service, message, user_id='me'):
    """Send an email message.
        Args:
        service: Authorized Gmail API service instance.
        message: Message to be sent. Se obtiene con las funciones ggmCreateMessage o ggmCreateMessageWithAttachment
        user_id: (str) User's email address. The special value "me"
        can be used to indicate the authenticated user.

        Returns:
        Sent Message.
    """
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print ('Message Id:', message['id'])
        return message
    except Exception as error:
        print ('An error occurred:', error)

######################################################################################################
# GOOGLE DRIVE
######################################################################################################
def gdrCambiarPermisos(drive_service, fileId, reset='no', permisos=None, usuarios=None, roles=None, notificar=False, cualquiera = None, sharedDrive=False):
    """
    Función que gestiona los permisos de un elemento en Drive.
    Se indica un conjunto de usuarios y los permisos que tendrá cada uno, y/o si el elemento se comparte a través de enlace.
    Los posibles permisos son:
       'reader': lector
       'writer': editor
       'commenter': puede comentar
       'revoke': quitar permisos concedidos
       'owner': propietario
       
    Parámetros:
        drive_service -- Objeto servicio con permisos de escritura en Google Drive. Se obtiene con el método connect.
        fileId        -- (str) Identificador del fichero.
        reset         -- (str) 'no', 'todo', 'parcial'. Indica cómo actuar con los permisos de usuarios. 
            - 'no' no altera los permisos existentes. 
            - 'todo' anula cualquier permiso anterior y establece estos como únicos. 
            - 'parcial', anula cualquier permiso anterior para los usuarios proporcionados y establece los indicados.
        permisos      -- (dict) de tipo {'eMail':permiso}, indicando el permiso otorgado a cada usuario.
        usuarios      -- (str) o (list) de (str) Direccion(es) de correo de los usuarios a los que dar permisos.
        roles         -- (str) o (list) de (str) Roles que asignar a los usuarios. 
            Si se proporciona (str), se aplica el mismo rol a todos.
        notificar     -- (bool) Indica si debe enviar notificación por email de cambio de permisos (defecto False)
        cualquiera    -- (bool) o None. (bool) Indica si se debe activar la opción "cualquiera que tenga el enlace". 
            - True activa la opción
            - False desactiva la opción.
            - None ignora esta opción.
        sharedDrive   -- (bool). Default False. Indica si el fichero se encuentra dentro de una unidad compartida
    Devuelve:
        Respuesta de cada uno de los permisos.
    """
    def callback(request_id, response, exception):
        if exception:
            # Handle error
            print (exception)
        else:
            if type(response) != str:
                print ("Permission Id: ", response.get('id'))

    #Recuperamos las propiedades del fichero para saber los permisos actuales.
    file = gdrGetFileProperties(drive_service=drive_service, fileId=fileId, sharedDrive=sharedDrive)
    
    #Comprobamos que los parámetros sean correctos.
    if permisos is not None:
        #Se informa parametro permisos
        if type(permisos) != dict:
            print('Se esperaba un diccionario en permisos')
        elif usuarios is not None or roles is not None:
                print('Se proporcionan tanto permisos como usuarios y/o roles. Solo se admite o solo permisos, o usuarios y roles.')
                return None
    elif usuarios is None or roles is None:
        #Si no se informan permisos usuarios y roles deben informarse ambos, salvo que solo se quiera cambiar el compartir por enlace.
        if cualquiera is None:
            print('Si no se proporciona un diccionario de permisos, se espera que se proporcione un listado de usuarios y roles.')
            return None
        else:
            #Si solo se quiere cambiar el compartir mediante enlace, indicamos que no haya cambio de permisos.
            permisos={}
    else:
        #Se informan parametros usuarios y roles. 
        #Vamos a convertir esos parámetros a un diccionario como habríamos esperado en el parámetro permisos.
        permisos = {}
        #Si me han pasado strings, necesito pasarlo a listas.
        if type(usuarios) == str:
            usuarios = [usuarios]
        if type(roles) == str:
            aux = roles
            roles = []
            for user in usuarios:
                roles.append(aux)
        if len(usuarios) != len(roles):
            print('Necesito tantos permisos como usuarios')
            return None
        
        #Construimos el diccionario de permisos
        for i, user in enumerate(usuarios):
            permisos[user] = roles[i]
    
    #Nos aseguramos que existe un diccionario para que no fallen las instrucciones posteriores.
    if permisos is None:
        permisos = {}
    
    #Controlamos que haya como máximo un solo owner:
    ownercount = 0
    for p in permisos.keys():
        if permisos[p] == 'owner':
            ownercount = ownercount+1
            if ownercount > 1:
                print('Solo se puede establecer un único owner. No se cambia ningún permiso.')
                return None
    
    #Preparamos el objeto para añadir todos los cambios de permisos.
					 
	
    batch = drive_service.new_batch_http_request(callback=callback)
    
    #Si se pide resetear todos los permisos, primero eliminamos los que hubiese de usuario, salvo el propietario, obviamente.
    if reset == 'todo':
        for permiso in file['permissions']:
            if permiso['type'] == 'user' and permiso['role'] != 'owner':
                permissionId = permiso['id']
                batch.add(
                    drive_service.permissions().delete(
                        fileId=fileId,
                        permissionId=permissionId,
                        supportsAllDrives=True
                    )
                )
                
    #Si se pide un reseteo parcial, los quitamos solo si están en la lista proporcionada.
    if reset == 'parcial':
        for permiso in file['permissions']:
            if permiso['type'] == 'user' and permiso['emailAddress'] in permisos.keys():
                batch.add(
                    drive_service.permissions().delete(
                        fileId=fileId,
                        permissionId=permiso['id'],
                        supportsAllDrives=True
                    )
                )
    
    #Ahora se añaden los permisos de usuario nuevos
    for user in permisos.keys():
        if permisos[user] in ('writer', 'reader', 'commenter'):
            user_permission = {
                'type': 'user',
                'role': permisos[user],
                'emailAddress': user
            }
            batch.add(drive_service.permissions().create(
                    fileId=fileId,
                    body=user_permission,
                    sendNotificationEmail=notificar,
                    fields='id'
            ))
        elif permisos[user] == 'owner':
            user_permission = {
                'type': 'user',
                'role': 'owner',
                'emailAddress': user
            }
            #Para los permisos de owner no se puede quitar la notificación de email
            batch.add(drive_service.permissions().create(
                    fileId=fileId,
                    body=user_permission,
                    fields='id',
                    transferOwnership = True,
                    supportsAllDrives=True
            ))
        elif permisos[user] == 'revoke':
            for permiso in file['permissions']:
                if permiso['type'] == 'user' and permiso['emailAddress'] == user:
                    batch.add(
                        drive_service.permissions().delete(
                            fileId=fileId,
                            permissionId=permiso['id'],
                            supportsAllDrives=True
                        )
                    )
                    break
    
    #Ahora toca la gestión de compartir mediante enlace:
    if cualquiera is not None:
        if cualquiera:
            domain_permission = {
                'type': 'domain',
                'role': 'reader',
                'domain': 'bbva.com'
            }
            batch.add(
                drive_service.permissions().create(
                    fileId=fileId,
                    body=domain_permission,
                    fields='id',
                    supportsAllDrives=True
                )
            )
        else:
            permissionId = None
            for permiso in file['permissions']:
                if permiso['type'] == 'domain' and permiso['domain'] == 'bbva.com':
                    permissionId = permiso['id']
                    break
            if permissionId:
                batch.add(
                    drive_service.permissions().delete(
                        fileId=fileId,
                        permissionId=permissionId,
                        supportsAllDrives=True
                    )
                )
            
    #Lanzamos el batch de permisos.
    resp = batch.execute()
    return resp

#%%
def gdrCopiarDocumento(drive_service, fileId, nuevotitulo):
    """
    Función que realiza una copia de un documento drive.
    Parámetros:
        drive_service -- Servicio con permisos para google drive.
        fileId        -- (str) Identificador del fichero que se desea clonar.
        nuevotitulo   -- (str) Título que se le desea dar al nuevo documento.
    Devuelve:
        (diccionario) Resupesta del servicio. Si ha tenido éxito, en la clave 'id' estará el identificador del fichero nuevo en Drive.
    """
    if nuevotitulo is None or nuevotitulo.strip() == '':
        print('No se puede crear un fichero con nombre vacío.')
        return None
    
    body = {
        'name': nuevotitulo
    }
    drive_response = drive_service.files().copy(
        fileId=fileId, body=body, supportsAllDrives=True).execute()
    #document_copy_id = drive_response.get('id')
    
    return drive_response
#%%
def gdrDescargarFichero(drive_service, fileId, rutaDescarga = None, mimeType = None, sobreescribir = False, logs = True, sharedDrive = False):
    """
    Esta función descargará un fichero de google drive en una ruta indicada.
    drive_service  -- Objeto servicio con permisos de lectura en Google Drive. Se obtiene con el método connect.
    fileId         -- (str) Identificador del fichero en Drive.
    rutaDescarga   -- (str) Carpeta en la que dejar el fichero. Si no se indica, lo dejará en la carpeta de trabajo.
    mimeType       -- (str) En caso de que el fichero sea nativo de Drive, tipo al que descargarlo. Si no se informa, lo intentará inferir.
    sobreescribir  -- (bool) Default False. Indica si debe machacar el fichero en caso de que exista, o crear un nombre nuevo en su lugar.
    logs           -- (bool) Default True. Indica si debe mostrar por pantalla el progreso de la descarga.
    sharedDrive    -- (bool) Default False. Indica si el fichero se encuentra en una unidad compartida
    
    return         -- (str) Ruta completa del fichero descargado.
    """
    import io
    import shutil
    import os
    from googleapiclient.http import MediaIoBaseDownload
    
    props = gdrGetFileProperties(drive_service=drive_service, fileId=fileId, sharedDrive=sharedDrive)
    if logs:
        print('mimetype original:',props['mimeType'])
        
    nombre = props['name']

    nativoDrive = False
    extensiones = {
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'        : '.xlsx',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'  : '.docx',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx',
        'application/pdf'                                                          : '.pdf'
    }
    
    #Google Spreadsheets:
    if props['mimeType'] == 'application/vnd.google-apps.spreadsheet':
        nativoDrive = True
        if not mimeType:
            mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            if logs:
                print('Se intentará la conversión a MS Excel')
                
    #Google Docs:
    elif props['mimeType'] == 'application/vnd.google-apps.document':
        nativoDrive = True
        if not mimeType:
            mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            if logs:
                print('Se intentará la conversión a MS Word')
    
    #Google Slides:
    elif props['mimeType'] == 'application/vnd.google-apps.presentation':
        nativoDrive = True
        if not mimeType:
            mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            if logs:
                print('Se intentará la conversión a MS PowerPoint')
    
    #Dependiendo de si el fichero es nativo de Drive o no, habrá que descargar de diferente manera:
    if nativoDrive:
        if not mimeType:
            mimeType = 'application/pdf'
            if logs:
                print('Se intentará la conversión a PDF')
        nombre = nombre + extensiones[mimeType]

        request = drive_service.files().export_media(
            fileId=fileId,
            mimeType=mimeType
        )
    else:
        request = drive_service.files().get_media(fileId=fileId)

    # Una vez obtenida la request, se procede a la descarga
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    if logs:
        print("Downloading... 0%")
    while done is False:
        status, done = downloader.next_chunk()
        if logs:
            print ("Downloading... %d%%." % int(status.progress() * 100))

    
    # Construimos el nombre de fichero en local:
    ruta = nombre
    if rutaDescarga:
        if rutaDescarga.endswith('/') or rutaDescarga.endswith('\\'):
            ruta = rutaDescarga + nombre
        else:
            ruta = rutaDescarga + '/' + nombre

    if not sobreescribir:
        if os.path.isfile(ruta):
            fichero = os.path.splitext(ruta)
            i = 1
            while True:
                rutaNueva = fichero[0] + ' (' + str(i) + ')' + fichero[1]
                if not os.path.isfile(rutaNueva):
                    ruta = rutaNueva
                    break
                i=i+1
        
    # Llevamos la descarga de la RAM a disco
    fh.seek(0)
    with open(ruta, 'wb') as f:
        shutil.copyfileobj(fh, f, length=131072)
    
    if logs:
        print("Download finished")
    return ruta
#%%

def gdrFindSharedDrive(drive_service, sharedDriveName):
    """
    Busca una unidad compartida por nombre y devuelve su id.
    Parámetros:
        drive_service  -- Objeto servicio con permisos de lectura en Google Drive. Se obtiene con el método connect.
        sharedDriveName -- (str) Nombre de la unidad compartida
    Devuelve (str) con el id de la unidad, o None si no la encuentra
    """
    nextPageToken = None
    shareddrives = []
    while True:
        try:
            respuesta = drive_service.drives().list(pageToken=nextPageToken).execute()
            shareddrives = shareddrives + respuesta['drives']
            if 'nextPageToken' in respuesta:
                nextPageToken = respuesta['nextPageToken']
                continue
            else:
                break
        except Exception as err:
            #print(err)
            break
    #shareddrives = drive_service.drives().list().execute()
    finished = False
    while not finished:
        for drive in shareddrives:
            if drive['name'] == sharedDriveName:
                return drive['id']
    print('No se ha encontrado la unidad compartida:', sharedDriveName)
    return None

#%%
def gdrGetFileProperties(drive_service, fileId = None, filePath = None, sharedDrive = False):
    """
    Devuelve todas las propiedades de un fichero.
    Parámetros:
        drive_service -- Objeto servicio con permisos de lectura en Google Drive. Se obtiene con el método connect.
        fileId        -- (str) Identificador del fichero.
        filePath      -- (str) Nombre y ruta del fichero (en caso de no proporcionar filePath)
        sharedDrive   -- (bool) Indica si buscar en unidades compartidas.
    Devuelve:
        Objeto diccionario con las diferentes propiedades del fichero.
        En caso de encontrar más de una coincidencia, devuelve una lista de diccionarios.
    """
    if fileId is None:
        if filePath is None:
            print('Necesito o el id o la ruta del fichero.')
            return None
        else:
            #Si acaba en barra, la quitamos
            if filePath.endswith('/'):
                filePath = filePath[:-1]
            path = filePath.split('/')
            folder = '/'.join(path[:-1])
            fids = gdrGetFolderId(drive_service=drive_service, folderPath=folder, sharedDrive=sharedDrive)
            if len(fids) == 0:
                fids = [('My Drive', 'root')]
            q = "name = '" + path[-1] + "' and '" + fids[-1][1] + "' in parents"
            files = drive_service.files().list(q=q, supportsAllDrives=sharedDrive, includeItemsFromAllDrives=sharedDrive).execute()['files']
            retobj = []
            for file in files:
                try:
                    fileId = file['id']
                    retobj.append(drive_service.files().get(fileId=fileId, fields='*', supportsAllDrives=sharedDrive).execute())
                except Exception as err:
                    print('Error al obtener propiedades de', file['name'], '(', file['id'], ')', err)
            if len(retobj) == 1:
                return retobj[0]
            elif len(retobj) == 0:
                print('No se ha encontrado el elemento.')
                return None
            else:
                return retobj
    else:
        return drive_service.files().get(fileId=fileId, fields='*', supportsAllDrives=sharedDrive).execute()

#%%
def gdrGetFileRevisions(drive_service, fileId, devuelvePandas = True):
    """
    Devuelve los detalles de las versiones de un fichero.
    Parámetros:
        drive_service -- Objeto servicio con permisos de lectura en Google Drive. Se obtiene con el método connect.
        fileId        -- (str) Identificador del fichero.
        devuelvePandas -- (bool) Indica True si quiere la respuesta como pandas dataframe con las propiedades básicas, o False como objeto lista de jsons.
    Devuelve:
        Objeto dataframe con las diferentes revisiones del fichero.
    """
    import pandas as pd
    
    respuestas = []
    nextPageToken = None
    while True:
        respuesta = drive_service.revisions().list(fileId=fileId, fields='*', pageToken=nextPageToken).execute()
        if 'revisions' in respuesta:
            respuestas += respuesta['revisions']
            if 'nextPageToken' in respuesta:
                nextPageToken = respuesta['nextPageToken']
            else:
                break
        else:
            break
     
    if not devuelvePandas:
        return respuestas
    else:
        df = pd.DataFrame()
        for revision in respuestas:
            fila = pd.DataFrame(
                index=[revision['id']], 
                data={
                    'modifiedTime' : revision['modifiedTime'],
                    'userEmail'    : revision['lastModifyingUser']['emailAddress'],
                    'userName'     : revision['lastModifyingUser']['displayName']
                }
            )
            df = pd.concat([df, fila])

        df['modifiedTime'] = pd.to_datetime(df['modifiedTime']).dt.tz_convert('Europe/Madrid')
        return df

#%%
def gdrGetFolderId(drive_service, folderPath, sharedDrive = False):
    """
    Dada una ruta de Google Drive, devuelve los identificadores de las carpetas encontradas de la más general a la más profunda.
    drive_service  -- Objeto servicio con permisos de lectura en Google Drive. Se obtiene con el método connect.
    folderPath     -- (str) Ruta que buscar.
    sharedDrive    -- (bool) Default False. Indica si tiene que buscar la ruta en unidades compartidas o no
    
    return         -- list[(str. str)] Lista de tuplas. Cada tupla tendrá el nombre de la carpeta y su id asociado.
    """
    folders = folderPath.split('/')
    previo = ''
    retobj = []
    #Si solo se pide el raíz, se devueve ese:
    import re
    raiz = ('mi unidad', 'my drive')

    carpeta = re.sub(r'^/(.*)/$', r'\1', folderPath)
    if carpeta.lower() in raiz:
        retobj.append((carpeta, 'root'))
        return retobj
    #Recorremos el árbol de carpetas
    for i, folder in enumerate(folders):
        #print(i, folder, '-', previo, '-', sharedDrive)
        if folder == '' or (previo == '' and folder in ('My Drive', 'Mi Unidad') and not sharedDrive):
            #print('continue')
            continue
        if previo == '' and sharedDrive:
            previo = gdrFindSharedDrive(drive_service, folder)
            retobj.append((folder, previo))
            continue

        q = "mimeType = 'application/vnd.google-apps.folder'"
        q = q + " and name = '" + folder + "'"
        #Si es el raíz, buscamos en mi unidad
        if previo == '':
            if sharedDrive:
                q = q + " and '" + previo + "' in parents"
            else:
                q = q + " and 'root' in parents"
        else:
            q = q + " and '" + previo + "' in parents"
        files = drive_service.files().list(q=q, supportsAllDrives=True, includeItemsFromAllDrives=True).execute()['files']
        #print(q, files)
        if len(files) == 0:
            print('No se ha encontrado la carpeta', folder)
            return retobj
        
        retobj.append((files[0]['name'], files[0]['id']))
        previo = files[0]['id']

    return retobj

#%%
def gdrListFiles(drive_service, folderId=None, folderPath=None, sharedDrive = False, trashed=False, fields='kind, id, name, mimeType'):
    """
    Lista los ficheros contenidos en una carpeta de drive.
    Parámetros:
        drive_service -- Objeto servicio con permisos de lectura en Google Drive. Se obtiene con el método connect.
        folderId        -- (str) Identificador de la carpeta.
        folderPath      -- (str) Ruta que buscar si folderId = None.
        sharedDrive   -- (bool) default False. Indica si debe listar ficheros de una unidad compartida
        trashed       -- (bool) default False. Indica si debe listar ficheros eliminados
        fields        -- (str) campos que debe devolver, separados por comas. Se puedes consultar los atributos disponibles en https://developers.google.com/drive/api/v3/reference/files
    Devuelve:
        Objeto diccionario con las diferentes propiedades del fichero.
    """
    if folderId is None:
        if folderPath is None:
            print('Error. Necesito una ruta o un id para buscar')
            return None
        fids = gdrGetFolderId(drive_service = drive_service, folderPath=folderPath, sharedDrive=sharedDrive)
        if len(fids) == 0:
            return []
        foldPaths = folderPath.split('/')
        if foldPaths[-1] == '':
            foldPaths.pop()
        if foldPaths[-1] != fids[-1][0]:
            print('Error. No se encuentra la ruta.')
            return None
        folderId = fids[-1][1]
    #Si es el raíz, buscamos en mi unidad
    q = "'" + folderId + "' in parents"
    if not trashed:
        q += ' and trashed=false'
    fields = 'nextPageToken, files(' + fields + ')'
    nextPageToken = None
    files = []
    while True:
        try:
            respuesta = drive_service.files().list(q=q, fields=fields, supportsAllDrives=sharedDrive, includeItemsFromAllDrives=sharedDrive, pageToken=nextPageToken).execute()#['files']
            files = files + respuesta['files']
            if 'nextPageToken' in respuesta:
                nextPageToken = respuesta['nextPageToken']
                continue
            else:
                break
        except Exception as err:
            #print(err)
            break
    return files
#%%
def gdrMoveFile(drive_service, fileId, destinationId=None, destinationPath=None, action = 'add', sharedDrive = False):
    """
    Mueve un fichero de una carpeta a otra.
    Parámetros:
        drive_service  -- Objeto servicio con permisos de lectura en Google Drive. Se obtiene con el método connect.
        fileId          -- (str) Identificador del fichero.
        destinationId   -- (str) Identificador de la carpeta destino.
        destinationPath -- (str) carpeta destino, expresada como ruta. Se usará si destinationId es None
        action          -- (str). Posibles valores 'add' o 'move'. Si 'add', añade el fichero a la nueva carpeta, pero no lo elimina de sus padres actuales.
        sharedDrive   -- (bool) default False. Indica si debe listar ficheros de una unidad compartida
    Devuelve un objeto con las propiedades del fichero movido
    """
    
    if destinationId is None:
        if destinationPath is None:
            print('Error. Necesito una ruta para mover el fichero')
            return None
        fids = gdrGetFolderId(drive_service = drive_service, folderPath=destinationPath, sharedDrive=sharedDrive)
        destPaths = destinationPath.split('/')
        if destPaths[-1] == '':
            destPaths.pop()
        if destPaths[-1] != fids[-1][0]:
            print('Error. No se encuentra la ruta destino.')
            return None
        destinationId = fids[-1][1]
    
    fileProps = gdrGetFileProperties(drive_service, fileId, sharedDrive=sharedDrive)
    if action == 'move':
        file = drive_service.files().update(
            fileId=fileId,
            addParents=destinationId,
            removeParents=",".join(fileProps['parents']),
            fields='id, parents', 
            supportsAllDrives=True
        ).execute()
    else:
        file = drive_service.files().update(
            fileId=fileId,
            addParents=destinationId,
            fields='id, parents', 
            supportsAllDrives=True
        ).execute()
    return file
                
#%%
def gdrSubirVersion(drive_service, idDriveFile, filename, mimetype=None, nuevoNombre = None, nuevaDescripcion = None):
    """
    Función que sube un fichero local a Google Drive
    Parámetros:
    drive_service -- Servicio con permisos para google drive.
    idDriveFile   -- (str) id del fichero que reemplazar en Drive
    filename      -- (str) nombre (y ruta) del fichero que se desea subir.
    mimetype      -- (str) formato de fichero. Ej. 'image/jpeg'. Por defecto None, la función intenta inferir el mimetype.
    nuevoNombre   -- None deja el nombre que tenía el fichero. (str) indicar nuevo nombre si se desea cambiar. True si utilizar el del fichero subido
    nuevaDescripcion -- (str) indicar nueva descripción si se desea cambiar
    
    Devuelve diccionario. Si la subida ha tenido éxito, el identificador estará en la propiedad 'id'.
    """
    import mimetypes
    from apiclient.http import MediaFileUpload,MediaIoBaseDownload
    from apiclient import errors

    try:
        drivefile = drive_service.files().get(fileId=idDriveFile, supportsAllDrives=True).execute()
        # File's new metadata.
        body = {}
        if nuevoNombre:
            if type(nuevoNombre) == bool:
                nuevoNombre = filename.split('/')[-1].split('\\')[-1]
            body['name'] = nuevoNombre
        if nuevaDescripcion:
            body['description'] = nuevaDescripcion

        #file_metadata = {'name': filename.split('/')[-1].split('\\')[-1]}

        if mimetype == None:
            mimetype, encoding = mimetypes.guess_type(filename)

        body['mimeType'] = mimetype
        
        # File's new content.
        media_body = MediaFileUpload(
            filename, 
            mimetype=mimetype, 
            resumable=True
        )

        # Send the request to the API.
        updated_file = drive_service.files().update(
            fileId      = idDriveFile,
            body        = body,
            media_body  = media_body,
            supportsAllDrives = True
        ).execute()
        
        return updated_file
    
    except errors.HttpError as error:
        print ('Ha ocurrido un error:\n', error)
        return None

#%%
def gdrUploadFile(drive_service, filename, mimetype=None):
    """
    Función que sube un fichero local a Google Drive
    Parámetros:
        drive_service -- Servicio con permisos para google drive.
        filename      -- (str) nombre (y ruta) del fichero que se desea subir.
        mimetype      -- (str) formato de fichero. Ej. 'image/jpeg'. Por defecto None, la función intenta inferir el mimetype.
    Devuelve diccionario. Si la subida ha tenido éxito, el identificador estará en la propiedad 'id'.
    """

    file_metadata = {'name': filename.split('/')[-1].split('\\')[-1]}

    if mimetype == None:
        mimetype, encoding = mimetypes.guess_type(filename)
        
    media = MediaFileUpload(filename, mimetype=mimetype)

    file = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id', 
        supportsAllDrives=True
    ).execute()
    return file
#%%