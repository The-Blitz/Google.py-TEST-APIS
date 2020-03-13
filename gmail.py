from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import googleapiclient.errors as errors
import simplejson
import pandas as pd

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/admin.directory.user']

def crearGmail(servicio, correo, apellidos, nombres , listaGrupos):
    try:
        datos = {"name": {"familyName": apellidos, "givenName": nombres, "fullName": nombres+" "+apellidos }, "password": "new user password",
            "primaryEmail": correo}

        parametro = correo[0]

        for tipo,grupo in listaGrupos:
            if(parametro == tipo): datos["orgUnitPath"] = grupo

        nuevoUsuario = servicio.users().insert(body=datos).execute()

        print('{0} con correo {1} fue creado con el id {2}'.format(nuevoUsuario.get('name').get('fullName'), nuevoUsuario.get('primaryEmail'),nuevoUsuario.get('id')))
        return nuevoUsuario

    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 409):
            print ('El usuario con correo "{0}" ya existe'.format(correo))
        return None


def borrarGmail(servicio, correo):

    try:
        servicio.users().delete(userKey=correo).execute()
        print('Usuario con correo {0} fue eliminado'.format(correo))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 404):
            print ('El usuario con correo "{0}" no existe'.format(correo))
    return

def obtenerCorreoporID(servicio, userId):
    return servicio.users().get(userKey=userId).execute()

def imprimirCorreos(servicio, cantidadCorreos = 10):
    # Call the Admin SDK Directory API
    print('Obteniendo los primeros correos')
    results = servicio.users().list(customer='my_customer', maxResults=cantidadCorreos,
                                orderBy='email').execute()
    users = results.get('users', [])

    if not users:
        print('No users in the domain.')
    else:
        print('Users:')
        for user in users:
            print(u'{0} ({1})'.format(user['primaryEmail'],
                user['name']['fullName']))

def main():
    """Shows basic usage of the Admin SDK Directory API.
    Prints the emails and names of the first 10 users in the domain.
    """
    creds = None
    #En token.pickle guardamos el acceso del usuario y lo reutilizamos, se crea en la primera ejecución
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # Si no hay credenciales validas el usuario debe loguearse

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Guardar credenciales para proximas ejecuciones
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    servicio = build('admin', 'directory_v1', credentials=creds)

    #user = crearGmail(servicio, "a00000007@sacooliveros.edu.pe" , "Diaz" , "Marco"  , [ ('a' , '/ALUMNOS_FRANQUICIA') ,('e' , '/ALUMNOS_COLEGIO') ])
    #borrarGmail(servicio,"a00000007@sacooliveros.edu.pe")
    #imprimirCorreos(servicio,5)


    df = pd.read_excel("Libro1.xlsx", sheet_name="Hoja1")
    for i in range(0, len(df)):
        dni = df.iloc[i]["persona_documento_numero"]
        apellidos = df.iloc[i]["a_paterno"] + " " + df.iloc[i]["a_materno"]
        nombre = df.iloc[i]["nombre_completo"].split()[1]
        correo = df.iloc[i]["persona_correo"]
        crearGmail(servicio, correo, apellidos, nombre,[ ('a' , '/ALUMNOS_FRANQUICIA') ,('e' , '/ALUMNOS_COLEGIO') ])
        borrarGmail(servicio, correo)

if __name__ == '__main__':
    main()