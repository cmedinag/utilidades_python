#%% Configuración inicial
import sys
import os

import warnings
warnings.filterwarnings("ignore")

# PARAMETROS
#Generales
workingfolder = os.path.dirname(sys.argv[0])
#workingfolder = r'C:\software\Advice\Data Coaching\plantilla_script'
if (not workingfolder.endswith('/')) and (not workingfolder.endswith('\\')):
    workingfolder += '/'
sys.path.append(workingfolder)

import utils
#%%

#Establecemos si es necesario
utils.autoSetProxy(log=False)

#%%
#Instalamos librería si es necesario. Pasamos una lista de dependencias.
#Si en el import la librería se llama diferente al nombre de instalación, pasamos una tupla.
#Ejemplos:
librerias = [
    'win32com', 
    ('tabula-py', 'tabula'),
    ('google-api-python-client', 'googleapiclient'),
    ('google-auth-httplib2', 'httplib2'),
    ('google-auth-oauthlib', 'google_auth_oauthlib'),
]

#Aseguramos que están instaladas
for libreria in librerias:
    if type(libreria) == tuple:
        utils.autoInstalarPaquete(libreria = libreria[0], alt=libreria[1])
    else:
        utils.autoInstalarPaquete(libreria = libreria)

#Importamos las que necesitamos
import win32com.client

#%%
#Código del script
excel = win32com.client.Dispatch('Excel.Application')
wb = excel.Workbooks.Open(r'H:\SC004884\carpeta\excel.xlsm')
excel.Application.Run('Macro1')
wb.Save()
excel.Application.Quit()
