#%%
# #El primer paso es asegurar que las librerías de google están instaladas. Sino, instalarlas antes de nada.
def autoSetProxy(proxy = "http://proxyvip:8080", url='https://pypi.python.org/simple', user=None, password=None, log=True):
    '''
    Esta función se encarga de detectar si hace falta un proxy para la salida a internet.
    En caso afirmativo, establece como proxy el que se le pase como parámetro.
    
    :param proxy -- (str) Dirección del proxy a establecer en caso de ser necesario. 'http[s]://direccion:puerto
    :param url   -- (str) URL para probar si existe salida a internet.
    
    :return: bool True si ha tenido éxito en la salida. False en caso contrario.

    '''
    import requests
    import os
    
    nopasswordtested = False
    for i in range(3):
        try:
            # Intenta hacer una petición a la página web sin proxy
            response = requests.get(url)
            if response.status_code == 200:
                if log: print('Conexión a internet ok')
                return True
        except (requests.exceptions.ProxyError, requests.exceptions.ConnectionError):
            if (not 'https_proxy' in os.environ) or (os.environ['https_proxy'] == ''):
                if log: print('No hay salida a internet. Estableciendo proxy y reintentando')
                if nopasswordtested:
                    if user is None and password is None:
                        creds = getUserPwd()
                    else:
                        creds = {'user':user, 'password':password}
                    setProxy(user=creds['user'], pwd=creds['password'])
                        
                else:
                    os.environ['https_proxy'] = proxy
                    nopasswordtested = True
            else:
                if log: print('Quitando proxy y reintentando')
                os.environ['https_proxy'] = ''
    return False

#%%
def autoInstalarPaquete(libreria:str, alt:'str|None'=None, log:bool=False, upgrade:bool=False, **kwargs):
    '''
    Función que comprueba si una librería está instalada en el sistema, y, de no ser así, la instala con pip.
    
    :param libreria  - (str) Nombre de la librería en el repositorio de python.
    :param alt       - (str) nombre de la librería al hacer el import. Si pasaoms None, asume el mismo valor que librería.
    :param log       - (bool) indica si omstrar por pantalla mensajes de éxito.
    
    :return bool True si la libreria se ha podido instalar. False en caso contrario.
    '''
    import importlib
    
    if alt is None: alt = libreria
    # Intenta importar la librería
    try:
        if not upgrade:
            importlib.import_module(alt)
            if log:
                print("La librería", libreria, "ya está instalada")
            return True
        else:
            raise ImportError('Hay que actualizar')
    
    # Si la librería no está instalada, la instala y luego la importa
    except (ImportError, NameError):
        if not upgrade and log:
            print("La librería", libreria, "no está instalada. Instalando...")
        elif upgrade and log:
            print("La librería", libreria, "se va a instalar/actualicar. Instalando...")
        # import pip
        # pip.main(['install', '--user', libreria])
        
        autoSetProxy(**kwargs)

        import os
        if upgrade: 
            os.system(f'pip install --user --upgrade {libreria}')
        else:
            os.system(f'pip install --user {libreria}')
        try:
            importlib.import_module(alt)
            if log:
                print("La librería", libreria, "ha sido instalada y cargada exitosamente")
            return True
        except ModuleNotFoundError as err:
            if log:
                print('No se puede importar la librería', libreria, '\n', err)
     
    return False

#%%
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
        'ú':'u',
        'ñ':'n',
        'ç':'c',
        'ü':'u'
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
    ventana.focus()
    
    # Colocar la ventana en primer plano
    ventana.attributes('-topmost', True)

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
def grabarlog(ruta, nombrefichero='log_uso.log', territorial='', mensaje=''):
    import os
    from datetime import datetime
    ahora = datetime.today().strftime('%Y-%m-%d %H:%M:%S')
    user = os.getenv('USERNAME')
    
    os.makedirs(ruta, exist_ok=True)
    
    with open(f'{ruta}/{nombrefichero}', 'a', encoding='utf-8') as log:
        log.write(f'[{ahora}] - {str(territorial).zfill(4)} - {user} - {mensaje}\n')

# %%
