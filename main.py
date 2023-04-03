import os
'''
Este código crea una estructura de carpetas para un proyecto (documento de tesis) y luego agrega un archivo .md a cada carpeta que no contenga subcarpetas. La estructura de carpetas se define como un diccionario llamado proyecto. Cada clave del diccionario es una cadena que representa el nombre de una carpeta y cada valor es otra cadena vacía o un subdiccionario que representa las subcarpetas de la carpeta. El scrip agrega un archivo markdown vacio a cada carpeta para que alli se inicie la escritura de parte del usuario.
'''
proyecto = {
    '1_portada': '',
    '2_caracterizacion_del_problema': '',
    '3_antecedentes': '',
    '4_justificacion': '',
    '5_objetivos': '',
    '6_factibilidad': '',
    '7_marco_de_referencia': '',
    '8_disenio_metodologico': {
        '8_1_descripcion': '',
        '8_2_materiales': '',
        '8_3_metodos_y_procedimientos': ''
    },
    '9_resultados': '',
    '10_discusion': '',
    '11_plan_operativo': '',
    '12_bibliografia': '',
    '13_ anexos': ''
}


def crear_carpetas_proyecto(diccionario, ruta_padre='', primera_vez=True):
    try:
        if primera_vez:
            ruta_proyecto = os.path.join(ruta_padre, 'proyecto')
            os.mkdir(ruta_proyecto)
        else:
            ruta_proyecto = ruta_padre

        for seccion, subsecciones in diccionario.items():
            ruta_seccion = os.path.join(ruta_proyecto, seccion)
            if isinstance(subsecciones, str):
                os.mkdir(ruta_seccion)
            elif isinstance(subsecciones, dict):
                os.mkdir(ruta_seccion)
                crear_carpetas_proyecto(subsecciones, ruta_seccion, primera_vez=False)
    except:
        print('El proyecto no se puede sobreescribir por seguridad...si desea comenzar nuevamente se sugiere borrar manualmente el proyecto previamente creado')

def tiene_subcarpetas(ruta_carpeta):
    return any(os.path.isdir(os.path.join(ruta_carpeta, subcarpeta)) for subcarpeta in os.listdir(ruta_carpeta))

def agregar_archivos(carpeta_padre):
    for ruta, carpetas, archivos in os.walk(carpeta_padre):
        for carpeta in carpetas:            
            ruta_carpeta = os.path.join(ruta, carpeta)
            if not(tiene_subcarpetas(ruta_carpeta)):
                nombre_archivo = carpeta + '.md'
                ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)
                with open(ruta_archivo, 'w') as archivo:
                    archivo.write('# Archivo en la carpeta ' + carpeta)

import os
import pypandoc

def convertir_md_a_docx(ruta_md):
    # Obtener la ruta y nombre del archivo sin la extensión
    ruta_sin_ext, _ = os.path.splitext(ruta_md)
    # Crear la ruta y nombre del archivo .docx resultante
    ruta_docx = ruta_sin_ext + '.docx'
    # Convertir el archivo markdown a .docx
    pypandoc.convert_file(ruta_md, 'docx', outputfile=ruta_docx)
    return ruta_docx


def convertir_archivos_en_carpeta(ruta_carpeta):
    for ruta, carpetas, archivos in os.walk(ruta_carpeta):
        for archivo in archivos:
            nombre_archivo, extension = os.path.splitext(archivo)
            if extension == '.md':
                ruta_md = os.path.join(ruta, archivo)
                convertir_md_a_docx(ruta_md)




if False:
    crear_carpetas_proyecto(proyecto)
    agregar_archivos('proyecto')
else:
    convertir_archivos_en_carpeta('proyecto')