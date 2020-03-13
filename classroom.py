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
SCOPES = ['https://www.googleapis.com/auth/classroom.courses', 'https://www.googleapis.com/auth/classroom.coursework.students',
          'https://www.googleapis.com/auth/classroom.rosters','https://www.googleapis.com/auth/classroom.announcements',
          'https://www.googleapis.com/auth/classroom.topics']

def crearClase(servicio,nombre):
    # Se crea un classroom con el nombre dado
    datos = {
        'name': nombre,
        'ownerId': 'me',
        'courseState': 'ACTIVE'
    }
    clase = servicio.courses().create(body=datos).execute()
    print('Clase creada: {0} ({1})'.format(clase.get('name'), clase.get('id')))

    return clase

def listarClases(servicio, cantidadClases = 10):
    # Se imprime lista de 10(valor por defecto) primeros clases
    output = servicio.courses().list(pageSize=cantidadClases).execute()
    clases = output.get('courses', [])

    if not clases:
        print('No se encontraron clases.')
    else:
        print('Los clases son los siguientes:')
        for clase in clases:
            print(clase['name'] , clase['id'])
    return

def obtenerClaseporID(servicio, idClase):
    # obtener la clase mediante su id
    return servicio.courses().get(id=idClase).execute()

def agregarTopicoaClase(servicio,idClase, nombreTopico):
    datos = {
        "name": nombreTopico
    }
    topico = servicio.courses().topics().create(courseId=idClase, body = datos).execute()
    print('Topico creado: ', topico['name'] , topico['topicId'])
    return topico

def agregarTareaaClase(servicio,idClase,idTopico,tituloTarea, tipoTarea):
    datos = {
        'title': tituloTarea,
        'workType': tipoTarea,
        'topicId': idTopico,
        'state': 'PUBLISHED',
    }
    tarea = servicio.courses().courseWork().create(courseId=idClase, body=datos).execute()
    print('Tarea creada con ID {0}'.format(tarea.get('id')))
    return

def agregarProfesoraClase(servicio, emailProfesor, idClase):
    datos = {
        'userId': emailProfesor
    }
    try:
        teacher = servicio.courses().teachers().create(courseId=idClase,
                                                      body=datos).execute()
        print('Se agrego al profesor {0} a la clase con ID "{1}"'.format(teacher.get('profile').get('name').get('fullName'),idClase))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 409):
            print ('El profsor con correo "{0}" ya es parte de la clase'.format(emailProfesor))
        else:
            raise
    return

def invitarPersonaaClase(servicio,emailAlumno, idClase, tipo ):

    datos = {
        'userId': emailAlumno,
        'role': tipo,
        'courseId': idClase
    }
    try:
        persona = servicio.invitations().create(body=datos).execute()
        print('{0} con correo {1} fue invitado a la clase con ID "{1}"'.format(tipo , persona.get('userId'), idClase))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 409):
            print ('{0} con correo "{1}" ya fue invitado a la clase'.format(tipo,emailAlumno))

    return

def agregarAlumnoaClase(servicio, emailAlumno, idClase,codigoClase):
    datos = {
        'userId': emailAlumno
    }
    try:
        student = servicio.courses().students().create(courseId=idClase,enrollmentCode=codigoClase,body=datos).execute()
        print('Alumno {0} esta ahora cursando la clase con ID "{1}"'.format(student.get('profile').get('name').get('fullName'),idClase))
    except errors.HttpError as e:
        error = simplejson.loads(e.content).get('error')
        if (error.get('code') == 409):
            print('El alumno ya se encuentra matriculado en la clase.')
        else:
            raise

def creacionMasiva( servicio, listaClases, listaTopicos , listaTareas):

    for clase in listaClases:
        claseActual = crearClase(servicio, clase)
        for topico in listaTopicos:
            nuevoTopico = agregarTopicoaClase(servicio,claseActual['id'] , topico)
            for tarea in listaTareas:
                agregarTareaaClase(servicio,claseActual['id'],nuevoTopico['topicId'],  tarea, 'ASSIGNMENT')

    return

def crearClasesVacias(servicio):
    df = pd.read_excel("Aulas.xlsx", sheet_name="2020-1")

    df_secundaria = df.iloc[2:25]
    df_primaria = df.iloc[79:97]

    s1 = crearClase(servicio, 'Secundaria_Primero_Lince')
    s2 = crearClase(servicio, 'Secundaria_Segundo_Lince')
    s3 = crearClase(servicio, 'Secundaria_Tercero_Lince')
    s4 = crearClase(servicio, 'Secundaria_Cuarto_Lince')
    s5 = crearClase(servicio, 'Secundaria_Quinto_Lince')

    for i in range(0,len(df_secundaria)):

        nombreTopico = df_secundaria.iloc[i]['Unnamed: 0']
        codTopico1 = df_secundaria.iloc[i]['Unnamed: 1']
        codTopico2 = df_secundaria.iloc[i]['Unnamed: 3']
        codTopico3 = df_secundaria.iloc[i]['Unnamed: 5']
        codTopico4 = df_secundaria.iloc[i]['Unnamed: 7']
        codTopico5 = df_secundaria.iloc[i]['Unnamed: 9']

        if(codTopico1[0]!='-'): agregarTopicoaClase(servicio, s1['id'], nombreTopico)
        if(codTopico2[0]!='-'): agregarTopicoaClase(servicio, s2['id'], nombreTopico)
        if(codTopico3[0]!='-'): agregarTopicoaClase(servicio, s3['id'], nombreTopico)
        if(codTopico4[0]!='-'): agregarTopicoaClase(servicio, s4['id'], nombreTopico)
        if(codTopico5[0]!='-'): agregarTopicoaClase(servicio, s5['id'], nombreTopico)

    s1 = crearClase(servicio, 'Primarria_Primero_Lince')
    s2 = crearClase(servicio, 'Primaria_Segundo_Lince')
    s3 = crearClase(servicio, 'Primaria_Tercero_Lince')
    s4 = crearClase(servicio, 'Primaria_Cuarto_Lince')
    s5 = crearClase(servicio, 'Primaria_Quinto_Lince')
    s6 = crearClase(servicio, 'Primaria_Sexto_Lince')

    for i in range(0, len(df_primaria)):

        nombreTopico = df_secundaria.iloc[i]['Unnamed: 0']
        codTopico1 = df_secundaria.iloc[i]['Unnamed: 1']
        codTopico2 = df_secundaria.iloc[i]['Unnamed: 2']
        codTopico3 = df_secundaria.iloc[i]['Unnamed: 3']
        codTopico4 = df_secundaria.iloc[i]['Unnamed: 4']
        codTopico5 = df_secundaria.iloc[i]['Unnamed: 5']
        codTopico6 = df_secundaria.iloc[i]['Unnamed: 6']

        if(codTopico1[0]!='-'): agregarTopicoaClase(servicio, s1['id'], nombreTopico)
        if(codTopico2[0]!='-'): agregarTopicoaClase(servicio, s2['id'], nombreTopico)
        if(codTopico3[0]!='-'): agregarTopicoaClase(servicio, s3['id'], nombreTopico)
        if(codTopico4[0]!='-'): agregarTopicoaClase(servicio, s4['id'], nombreTopico)
        if(codTopico5[0]!='-'): agregarTopicoaClase(servicio, s5['id'], nombreTopico)
        if(codTopico6[0]!='-'): agregarTopicoaClase(servicio, s6['id'], nombreTopico)

def main():
    #Obteniendo las credenciales
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
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    #Usamos el servicio con las credenciales obtenidas
    servicio = build('classroom', 'v1', credentials=creds)

    #listarClases(servicio)

    #creacionMasiva( servicio , ['Clase 1' , 'Clase 2', 'Clase 3'] , ['Topico 1' , 'Topico 2'] , ['Tarea 1'])
    #listarClases(servicio)
    #claseActual = obtenerClaseporID(servicio,62895258692)
    #print(claseActual['name'], claseActual['id'])

    #listarClases(servicio)

    s1 = obtenerClaseporID(servicio, 64668436220)
    '''
    s2 = obtenerClaseporID(servicio, 64667767495)
    s3 = obtenerClaseporID(servicio, 64667872425)
    s4 = obtenerClaseporID(servicio, 64668871518)
    s5 = obtenerClaseporID(servicio, 64668484394)

    p1 = obtenerClaseporID(servicio, 64667767847)
    p2 = obtenerClaseporID(servicio, 64668701650)
    p3 = obtenerClaseporID(servicio, 64668991022)
    p4 = obtenerClaseporID(servicio, 64669173318)
    p5 = obtenerClaseporID(servicio, 64668926675)
    p6 = obtenerClaseporID(servicio, 64668638694)
    '''
    df = pd.read_excel("Aulas.xlsx", sheet_name="Hoja 1",header=None)

    for i in range(0, len(df)):
        nom = df.iloc[i][0]
        ape = df.iloc[i][1]
        correo = df.iloc[i][2]
        invitarPersonaaClase(servicio,correo,s1['id'],'STUDENT')


if __name__ == '__main__':
    main()
