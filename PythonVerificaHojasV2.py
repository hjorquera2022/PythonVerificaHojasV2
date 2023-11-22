#PythonVerificaHojas.py

                # ACTUALIZA EDITABLE DOC VIG - A
                # ACTUALIZA EDITABLE REV LETRA - A
                # ACTUALIZA EDITABLE REV NUM - A
                # ACTUALIZA PDF DOC VIG - A
                # ACTUALIZA PDF REV LETRA - A
                # ACTUALIZA PDF REV NUM - A


import pandas as pd
import os

# Ruta base donde se deben verificar los subdirectorios
ruta_base = 'R:\\01 PARCIALIDADES\\'   

# Nombre del archivo de log
archivo_log = ruta_base + '0000-00 ADMINISTRACION\\LOG\\log_ValidarHojasParcialidades.txt'

# Planilla con la lista de parcialidades
archivo_excel = ruta_base + 'Listado de Parcialidades.xlsx'

# Carga el archivo Excel en un DataFrame Hoja de Parcialidades.
df = pd.read_excel(archivo_excel, sheet_name='PARCIALIDADES')

# Filtra el DataFrame para considerar solo parcialidades a 'PROCESAR' igual a 'S'
df_parcialidades = df[df['PROCESAR'] == 'S']

# Abre el archivo de log en modo de escritura
with open(archivo_log, 'w') as log_file:

    # Itera a través de cada parcialidad y la procesa
    for parcialidad in df_parcialidades['PARCIALIDAD']:
        log_file.write(f'Parcialidad: {parcialidad}\n')

        #******* 
        #******* RECORRER TODAS LAS PARCIALIDADES VERIFICANDO SI LOS ARCHIVOS DE INGENIERIA CONTIENEN LAS 6 HOJAS DE LA PLANILLA
        #******* 
        # ACTUALIZA EDITABLE DOC VIG 
        # ACTUALIZA EDITABLE REV LETRA 
        # ACTUALIZA EDITABLE REV NUM 
        # ACTUALIZA PDF DOC VIG 
        # ACTUALIZA PDF REV LETRA 
        # ACTUALIZA PDF REV NUM 
         
        
      #******* Abrir Planilla CONTROL DOCUMENTOS ING DEF Pxxxx-xx con las 8 hojas para traspasar a BAT
        parcialidad_0_7_10 = parcialidad[0:7]
        if parcialidad_0_7_10 == '0029-14':
            parcialidad_0_7_10 = parcialidad[0:10]

        archivo_parcialidad = ruta_base + parcialidad + '\\CONTROL DOCUMENTOS ING DEF P' + parcialidad_0_7_10 + '.xlsx'
        
        if not os.path.exists(archivo_parcialidad):
              log_file.write(f'Parcialidad: {parcialidad} SIN ARCHIVO DE INGENIERIA {archivo_parcialidad}\n')
        else:
                print(f'Procesando Parcialidad: {parcialidad} ARCHIVO:  {archivo_parcialidad}')
                log_file.write(f'Parcialidad: {parcialidad} Archivo {archivo_parcialidad}\n')

                # Lee el archivo Excel para obtener los nombres de las hojas
                xl = pd.ExcelFile(archivo_parcialidad)
                nombres_hojas = xl.sheet_names

                # Verifica si existen las hojas específicas
                hojas_a_verificar = ['\xc3\xbaLTIMA VERSI\xc3\xb3N','REV LETRA APRO','REV NUM APRO','ACTUALIZA EDITABLE DOC VIG', 'ACTUALIZA EDITABLE REV LETRA', 'ACTUALIZA EDITABLE REV NUM','ACTUALIZA PDF DOC VIG','ACTUALIZA PDF REV LETRA','ACTUALIZA PDF REV NUM']
                for hoja in hojas_a_verificar:
                    if not hoja in nombres_hojas:
                        print(f'La hoja "{hoja}" no existe en {archivo_excel}.')
                        log_file.write(f'Parcialidad: {parcialidad} Archivo {archivo_parcialidad}\n')
                
print("Validacion fianalizada. Los resultados se han guardado en R:\01 PARCIALIDADES\0000-00 ADMINISTRACION\LOG en el archivo de log_ValidarHojasParcialidades.")
log_file.close
