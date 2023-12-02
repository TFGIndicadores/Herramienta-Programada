import pandas as pd
import sqlite3
import os
import flet as ft
from flet import (
    ElevatedButton,
    FilePicker,
    FilePickerResultEvent,
    Page,
    Row,
    Text,
    icons,
    Ref
)


db_path = "PruebasTFG_new.db"
tablas = ['IndicadoresServicio', 'IndicadoresEspecialidad', 'IndicadoresDoctor', 'IndicadoresReferencias', 'IndicadoresListas', 'IndicadoresMetas']

def validararchivo(filepath):
    """
    Verifica si la extensión del archivo es .xlsx, y valida que en la hoja 'Inicio' en la celda A7 exista una 'x', determinado como el metodo de validacion del formulario oficial
    
    Args:
    filepath (str): Ruta del archivo a validar.

    Returns:
    bool: True si el archivo existe y es accesible, False en caso contrario.
    """
    
    if filepath.endswith('.xlsx'):
        xls = pd.ExcelFile(filepath)
        for sheet_name in xls.sheet_names:
            if sheet_name == "Inicio":
                df = pd.read_excel(filepath, 'Inicio')
                val = df.iloc[5,0]
                if val == "x":
                    return 'comp'
                elif val == 'y':
                    return 'data'
            else:
                return False
        else: 
            return False   
    else:
        return False

def validardatabase(periodo):  
    """
    Verifica si en la base de datos existen indicadores para el periodo indicado.

    Esta función se conecta a la base de datos SQLite especificada en 'db_path' y realiza una consulta
    para determinar si hay indicadores disponibles para el periodo proporcionado.

    Args:
    periodo (str): Periodo de tiempo para el cual se verifica la existencia de indicadores en la base de datos.

    Returns:
    bool: True si existen indicadores para el periodo indicado, False en caso contrario.
    """
    sqlconn = sqlite3.connect(db_path)
    cursor = sqlconn.cursor()

    try:
        #Revisa todas las tablas y realiza un select para el periodo indicado, devuelve verdadero en caso de encontrar algun valor.
        for tabla in tablas:
            cursor.execute(f"SELECT * FROM {tabla} WHERE PERIODO = '{periodo}'")
            rows = cursor.fetchall()
            if len(rows) > 0:
                cursor.close()
                sqlconn.close()
                return True

        # Si no hay valores en ninguna tabla, devuelve falso
        cursor.close()
        sqlconn.close()
        return False

     # Si la base de datos da error (no existe), devuelve falso
    except sqlite3.Error:
        cursor.close()
        sqlconn.close()
        return False


def borrardatos(periodo):
    """
    Elimina las filas de la base de datos encontrados para el periodo indicado, para prevenir duplicación de datos.
    """
    # Conexión a la base de datos
    sqlconn = sqlite3.connect(db_path)
    cursor = sqlconn.cursor()

    #En caso de ejecutarse, intenta eliminar todos los records para cada tabla en el periodo indicado.
    try:
        for tabla in tablas:
            cursor.execute(f"DELETE FROM {tabla} WHERE PERIODO = '{periodo}'")
        sqlconn.commit()

    #Manejo de excepciones
    except sqlite3.Error as e:
        sqlconn.rollback()

    finally:
        cursor.close()
        sqlconn.close()


def prosFomularioAdicional(filepath):
    """
    Recibe la ubicación del formulario de datos, con base en el contenido de las jo
    """
    xls = pd.ExcelFile(filepath)
    dfs = {}
    for sheet_name in xls.sheet_names:
        if sheet_name != "Consolidado":
            dfs[sheet_name] = pd.read_excel(xls, sheet_name)
            dfs[sheet_name].columns = dfs[sheet_name].iloc[1]
            dfs[sheet_name] = dfs[sheet_name].iloc[2:]
            dfs[sheet_name] = dfs[sheet_name].reset_index(drop=True)
        else:
            df = pd.read_excel(xls, sheet_name)
            # Evaluar si la hoja es la mensual o es de especialidad, 
            #** Agregar validaciones adicionales para asegurar que solo se procese el file correcto

            # Eliminar filas y columnas vacías

            df.dropna(axis = 0,how = "all",inplace = True)  
            df.dropna(axis = 1,how = "all",inplace = True)  
            df = df.reset_index(drop=True)

            # Eliminar fila extra de la hoja "Mensual"

            df.drop(2, inplace = True)
            df = df.reset_index(drop=True)

            # Replicar la tabla de excel del resumen del mes con los totales de consultas por cada categoria

            # Se hace un slice de las primeras filas 18 en adelante, se eliminan vacios full y se auto completan algunas categorias faltantes

            row = df[df.iloc[:,0]=='RESUMEN DEL MES'].index[0]
            df_total = df.iloc[row:row+10].copy()  

            df_total.dropna(axis = 1,how = "all",inplace = True)
            df_total[:3].fillna(method="ffill", axis=1,inplace = True)
            df_total[:3].fillna(method="ffill", axis=0,inplace = True)

            # Se transpone la tabla para concatenar las categorias y unificarlas en una sola columna(celdas combinadas de excel) 
            df_total = df_total.transpose()
            df_total.iloc[1, 1] = df_total.iloc[1, 1] +" - "+ df_total.iloc[1, 2]
            df_total.iloc[2:6, 1] = df_total.iloc[2:6, 1].astype(str) +" - "+ df_total.iloc[2:6, 2].astype(str) +" - "+ df_total.iloc[2:6, 3].astype(str)
            df_total.iloc[1:-3, 0] = df_total.iloc[1:-3, 0].astype(str) +" - "+ df_total.iloc[1:-3, 1].astype(str)

            #Se eliminan las columnas que tenian las subcategorias, y se reindexa
            df_total.drop(df_total.columns[[1, 2, 3]],axis = 1, inplace = True)
            df_total.columns = df_total.iloc[0]
            df_total = df_total[1:]
            df_total = df_total.reset_index(drop=True)
            df_total.columns.rename("",inplace=True)
            df_total.iloc[:, 0] = df_total.iloc[:, 0].str.replace('Frecuencia (marcar solo una)', 'Frecuencia', regex=False)

            # Se transpone la tabla para volver al formato original (validar si vale la pena)
            df_total = df_total.transpose()
            df_total.columns = df_total.iloc[0]
            df_total = df_total[1:]
            dfs[sheet_name] = df_total
    return dfs

def calcIndicadores(periodo,form):
    hojasFormulario = prosFomularioAdicional(form)
    
    #Calculo de indicadores servicio
    indServicio = pd.DataFrame(columns=['PERIODO', 'INDICADOR', 'VALOR', 'META', 'RANGO'])

    nuevoIndServ=[]

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'ESTÁNDAR PACIENTE POR HORA DE LA UNIDAD',
        'VALOR': (hojasFormulario['Consultas Externas']['Consultas programadas en consulta externa'].sum()/hojasFormulario['Consultas Externas']['Horas programadas para consulta externa'].sum())
                if hojasFormulario['Consultas Externas']['Horas programadas para consulta externa'].sum() > 0 else None,
        'META': hojasFormulario['Metas']['Meta'][0],
        'RANGO': hojasFormulario['Metas']['Rango'][0]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TOTAL SUPERIOR',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][0],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TOTAL INFERIOR',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][1],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'PARCIAL SUPERIOR',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][2],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'PARCIAL INFERIOR',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][3],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'OBTURADORES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][4],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'REPARACIONES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][5],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TOTAL DE APARATOS ORTODONCIA',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][6],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TOTAL DE APARATOS ODONTOPEDIATRIA',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][7],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TOTAL DE PLANOS OCLUSALES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][8],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE REFERENCIAS ACEPTADAS',
        'VALOR': hojasFormulario['Referencias']['Aceptado'].sum(),
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE REFERENCIAS RECHAZADAS',
        'VALOR': hojasFormulario['Referencias']['Rechazado'].sum(),
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CONSULTAS ODONTOLÓGICAS PRIMERA VEZ',
        'VALOR': hojasFormulario['Consolidado'].iloc[0,:5].sum(),
        'META': hojasFormulario['Metas']['Meta'][10],
        'RANGO': hojasFormulario['Metas']['Rango'][10]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CONSULTAS ODONTOLÓGICAS SUBSECUENTES',
        'VALOR': hojasFormulario['Consolidado'].iloc[0,5],
        'META': hojasFormulario['Metas']['Meta'][11],
        'RANGO': hojasFormulario['Metas']['Rango'][11]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TELECONSULTAS',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][9],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'HORAS EN TELECONSULTAS',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][10],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TELEORIENTACIONES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][11],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'HORAS EN TELEORIENTACIONES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][12],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'APARATOLOGÍA',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][13],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'HORAS DE HOSPITALIZACIÓN',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][14],
        'META': None,
        'RANGO': None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'NÚMERO DE NIÑOS (AS) DE 0 A MENOS DE 10 AÑOS CON ATENCIÓN ODONTOLÓGICA PREVENTIVA DE PRIMERA VEZ EN EL AÑO',
        'VALOR': hojasFormulario['Consolidado'].iloc[1,7],
        'META': hojasFormulario['Metas']['Meta'][12],
        'RANGO': hojasFormulario['Metas']['Rango'][12]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'NÚMERO DE ADOLESCENTES DE 10 A MENOS DE 20 AÑOS CON ATENCIÓN ODONTOLÓGICA PREVENTIVA DE PRIMERA VEZ EN EL AÑO',
        'VALOR': hojasFormulario['Consolidado'].iloc[2,7],
        'META': hojasFormulario['Metas']['Meta'][13],
        'RANGO': hojasFormulario['Metas']['Rango'][13]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'PACIENTES EMBARAZADAS CON ATENCIÓN ODONTOLÓGICA PREVENTIVA DE PRIMERA VEZ EN EL AÑO',
        'VALOR': hojasFormulario['Consolidado'].iloc[0,6],
        'META': hojasFormulario['Metas']['Meta'][14],
        'RANGO': hojasFormulario['Metas']['Rango'][14]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'NIÑOS (AS) DE 0 A MENOS DE 10 AÑOS',
        'VALOR': 100*(hojasFormulario['Consolidado'].iloc[1,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][17])
                if hojasFormulario['Otros Datos']['Resultado'][17] > 0 else None,
        'META': hojasFormulario['Metas']['Meta'][15],
        'RANGO': hojasFormulario['Metas']['Rango'][15]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'ADOLESCENTES DE 10 A MENOS DE 20 AÑOS',
        'VALOR': 100*(hojasFormulario['Consolidado'].iloc[2,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][18])
                if hojasFormulario['Otros Datos']['Resultado'][18] > 0 else None,
        'META': hojasFormulario['Metas']['Meta'][16],
        'RANGO': hojasFormulario['Metas']['Rango'][16]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'HOMBRES DE 20 AÑOS A 64 AÑOS',
        'VALOR': 100*(hojasFormulario['Consolidado'].iloc[3,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][19])
                if hojasFormulario['Otros Datos']['Resultado'][19] > 0 else None,
        'META': hojasFormulario['Metas']['Meta'][17],
        'RANGO': hojasFormulario['Metas']['Rango'][17]
    })


    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'MUJERES DE 20 AÑOS A 64 AÑOS',
        'VALOR': 100*(hojasFormulario['Consolidado'].iloc[4,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][20])
                if hojasFormulario['Otros Datos']['Resultado'][20] > 0 else None,
        'META': hojasFormulario['Metas']['Meta'][18],
        'RANGO': hojasFormulario['Metas']['Rango'][18]
    })


    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'PERSONAS DE MÁS DE 65 AÑOS',
        'VALOR': 100*(hojasFormulario['Consolidado'].iloc[5,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][21])
                if hojasFormulario['Otros Datos']['Resultado'][21] > 0 else None,
        'META': hojasFormulario['Metas']['Meta'][19],
        'RANGO': hojasFormulario['Metas']['Rango'][19]
    })
    indServicio = pd.concat([indServicio, pd.DataFrame(nuevoIndServ)], ignore_index=True)


    #Calculo oficial de indicadores especialidad
    esps = hojasFormulario['Consultas Externas']['Especialidad'].unique()

    indEspecialidad = pd.DataFrame(columns=['PERIODO', 'ESPECIALIDAD', 'INDICADOR', 'VALOR'])

    for esp in esps:
        nuevoIndEsp = []

        nuevoIndEsp.append({
            'PERIODO': periodo,
            'ESPECIALIDAD': esp,
            'INDICADOR': 'HORAS PROGRAMADAS PARA LA ATENCIÓN DE PACIENTES',
            'VALOR': hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Horas programadas para consulta externa'].sum()
        })

        nuevoIndEsp.append({
            'PERIODO': periodo,
            'ESPECIALIDAD': esp,
            'INDICADOR': 'PRODUCCIÓN REAL',
            'VALOR': hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Consultas realizadas en consulta externa'].sum()
        })

        nuevoIndEsp.append({
            'PERIODO': periodo,
            'ESPECIALIDAD': esp,
            'INDICADOR': 'HORAS UTILIZADAS PARA LA ATENCIÓN DE PACIENTES',
            'VALOR': hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Horas utilizadas para consulta externa'].sum()
        })

        nuevoIndEsp.append({
            'PERIODO': periodo,
            'ESPECIALIDAD': esp,
            'INDICADOR': 'USUARIOS POR HORA PROGRAMADA PARA LA ATENCIÓN DE PACIENTES',
            'VALOR': (hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Consultas realizadas en consulta externa'].sum() /
                    hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Horas programadas para consulta externa'].sum())
                    if hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Horas programadas para consulta externa'].sum() > 0 else None
        })

        indEspecialidad = pd.concat([indEspecialidad, pd.DataFrame(nuevoIndEsp)], ignore_index=True)

    
    #Calculo oficial de indicadores profesional
    indDoctor = pd.DataFrame(columns=['PERIODO','PROFESIONAL', 'INDICADOR', 'VALOR', 'META', 'RANGO'])

    for index, row in hojasFormulario['Consultas Externas'].iterrows():
        nuevoIndDoc = []
        
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'CAPACIDAD MÁXIMA PRÁCTICA',
            'VALOR': ( 2 * row['Horas programadas para consulta externa']),
            'META': None,
            'RANGO': None
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'CAPACIDAD DE PRODUCCIÓN PREVISTA',
            'VALOR': (row['Horas programadas para consulta externa'] * hojasFormulario['Metas']['Meta'][2]),
            'META': None,
            'RANGO': None
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'PRODUCCIÓN REAL',
            'VALOR': row['Consultas realizadas en consulta externa'],
            'META': (row['Horas programadas para consulta externa'] * hojasFormulario['Metas']['Meta'][2])*hojasFormulario['Metas']['Meta'][1],
            'RANGO': (row['Horas programadas para consulta externa'] * hojasFormulario['Metas']['Meta'][2])*hojasFormulario['Metas']['Meta'][1]*hojasFormulario['Metas']['Porcentaje de desviación de la meta'][1]
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'USUARIOS POR HORA PROGRAMADA PARA LA ATENCIÓN DE PACIENTES',
            'VALOR': (row['Consultas realizadas en consulta externa']/row['Horas programadas para consulta externa'])
                    if row['Horas programadas para consulta externa'] > 0 else None,
            'META': None,
            'RANGO': None
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'USUARIOS POR HORA UTILIZADA PARA LA ATENCIÓN DE PACIENTES',
            'VALOR': (row['Consultas realizadas en consulta externa']/row['Horas utilizadas para consulta externa'])
                    if row['Horas utilizadas para consulta externa'] > 0 else None,
            'META': (row['Consultas realizadas en consulta externa']/row['Horas programadas para consulta externa'])*hojasFormulario['Metas']['Meta'][3] 
                    if row['Horas programadas para consulta externa'] > 0 else None,
            'RANGO': (row['Consultas realizadas en consulta externa']/row['Horas programadas para consulta externa'])*hojasFormulario['Metas']['Meta'][3] * hojasFormulario['Metas']['Porcentaje de desviación de la meta'][3] 
                    if row['Horas programadas para consulta externa'] > 0 else None 
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'AUSENTISMO CONSULTA EXTERNA',
            'VALOR': 100*(row['Citas perdidas en consulta externa']/(row['Consultas realizadas en consulta externa']+row['Citas perdidas en consulta externa']+row['Cupos no utilizados en consulta externa']))
                    if (row['Consultas realizadas en consulta externa']+row['Citas perdidas en consulta externa']+row['Cupos no utilizados en consulta externa']) > 0 else None,
            'META': hojasFormulario['Metas']['Meta'][4],
            'RANGO': hojasFormulario['Metas']['Rango'][4]
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'SUSTITUCIÓN DE PACIENTES CONSULTA EXTERNA',
            'VALOR': 100*(row['Citas sustituidas en consulta externa']/row['Citas perdidas en consulta externa'])
                    if row['Citas perdidas en consulta externa'] > 0 else None,
            'META': hojasFormulario['Metas']['Meta'][6],
            'RANGO': hojasFormulario['Metas']['Rango'][6]
        })

        balPerExt = ((row['Citas perdidas en consulta externa']+row['Cupos no utilizados en consulta externa']) - (row['Citas sustituidas en consulta externa']+row['Recargos en consulta externa']))
    
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'BALANCE DE CUPOS PERDIDOS EN CONSULTA EXTERNA',
            'VALOR': balPerExt,
            'META': hojasFormulario['Metas']['Meta'][8]*row['Consultas programadas en consulta externa'],
            'RANGO': hojasFormulario['Metas']['Meta'][8]*row['Consultas programadas en consulta externa']*hojasFormulario['Metas']['Rango'][8]
        })
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'COSTO POR CUPOS PERDIDOS EN CONSULTA EXTERNA',
            'VALOR': (balPerExt * hojasFormulario['Otros Datos']['Resultado'][15]),
            'META': None,
            'RANGO': None
                    
        })
        indDoctor = pd.concat([indDoctor, pd.DataFrame(nuevoIndDoc)], ignore_index=True)
        
    #Consulta Procedimientos
    for index, row in hojasFormulario['Consultas Procedimientos'].iterrows():
        nuevoIndDoc = []
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'AUSENTISMO CONSULTA PROCEDIMIENTOS',
            'VALOR': 100*(row['Citas perdidas en consulta procedimiento']/(row['Consultas realizadas en consulta procedimiento']+row['Citas perdidas en consulta procedimiento']+row['Cupos no utilizados en consulta procedimiento']))
                    if (row['Consultas realizadas en consulta procedimiento']+row['Citas perdidas en consulta procedimiento']+row['Cupos no utilizados en consulta procedimiento']) > 0 else None,
            'META': hojasFormulario['Metas']['Meta'][5],
            'RANGO': hojasFormulario['Metas']['Rango'][5]
        })
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'SUSTITUCIÓN DE PACIENTES CONSULTA PROCEDIMIENTOS',
            'VALOR': 100*(row['Citas sustituidas en consulta procedimiento']/row['Citas perdidas en consulta procedimiento'])
                    if row['Citas perdidas en consulta procedimiento'] > 0 else None,
            'META': hojasFormulario['Metas']['Meta'][7],
            'RANGO': hojasFormulario['Metas']['Rango'][7]
        })
        balPerProc = ((row['Citas perdidas en consulta procedimiento']+row['Cupos no utilizados en consulta procedimiento'])-(row['Citas sustituidas en consulta procedimiento']+row['Recargos en consulta procedimiento']))
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'BALANCE DE CUPOS PERDIDOS EN CONSULTA PROCEDIMIENTOS',
            'VALOR': balPerProc,
            'META': hojasFormulario['Metas']['Meta'][9]*row['Consultas programadas en consulta procedimiento'],
            'RANGO': hojasFormulario['Metas']['Meta'][9]*row['Consultas programadas en consulta procedimiento']*hojasFormulario['Metas']['Rango'][9]
        })
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'COSTO POR CUPOS PERDIDOS EN CONSULTA PROCEDIMIENTOS',
            'VALOR': (balPerProc * hojasFormulario['Otros Datos']['Resultado'][16]),
            'META': None,
            'RANGO': None
        })
        indDoctor = pd.concat([indDoctor, pd.DataFrame(nuevoIndDoc)], ignore_index=True)

    #Ortodoncia Ortopedia
    for index, row in hojasFormulario['Ortodoncia-Ortopedia'].iterrows():
        nuevoIndDoc = []
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'PACIENTES EN ORTODONCIA',
            'VALOR': row['Porcentaje de pacientes en ortodoncia'],
            'META': None,
            'RANGO': None
        })
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'PACIENTES EN ORTOPEDIA',
            'VALOR': row['Porcentaje de pacientes en ortopedia'],
            'META': None,
            'RANGO': None
        })
        indDoctor = pd.concat([indDoctor, pd.DataFrame(nuevoIndDoc)], ignore_index=True)
        

    indDoctor['original_index'] = indDoctor.index
    indDoctor = indDoctor.sort_values(by=['PROFESIONAL', 'original_index'])
    indDoctor.drop('original_index', axis=1, inplace=True)
    indDoctor.reset_index(drop=True,inplace=True)


    #Referencias por area de salud
    indReferencias = hojasFormulario['Referencias'].iloc[:,1:].groupby('Área de Salud', as_index = False).sum()
    indReferencias.insert(0, 'PERIODO', periodo)

    #Listas de espera
    indListas = hojasFormulario['Listas de espera'].iloc[:,:3].copy()
    indListas.insert(0, 'PERIODO', periodo)
    indListas

    #Metas de indicadores
    listMetas = hojasFormulario['Metas'].iloc[:,:4].copy()
    listMetas.insert(0, 'PERIODO', periodo)


    #Cargar tablas a BD
    if not os.path.exists(db_path):
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS IndicadoresServicio (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                PERIODO TEXT,
                INDICADOR TEXT,
                VALOR REAL,
                META REAL,
                RANGO REAL
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS IndicadoresEspecialidad (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                PERIODO TEXT,
                ESPECIALIDAD TEXT,
                INDICADOR TEXT,
                VALOR REAL
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS IndicadoresDoctor (
                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                PERIODO TEXT,
                PROFESIONAL TEXT,
                INDICADOR TEXT,
                VALOR REAL,
                META REAL,
                RANGO REAL
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS IndicadoresReferencias (
                PERIODO	TEXT,
                "Área de Salud" TEXT,
                "O.G." REAL,
                "O.G.A" REAL,
                "ORTOD." REAL,
                "ENDOD." REAL,
                "PERIOD." REAL,
                "PROSTOD." REAL,
                "TTM D.O." REAL,
                "PROT.MAXILOF." REAL,
                "ODONTOPED." REAL,
                "ODONTOGER." REAL,
                "CIR.MAXILOF." REAL,
                Rechazado REAL,
                Aceptado REAL
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS IndicadoresListas (
                PERIODO	TEXT,
                Especialidad TEXT,
                "Factor crítico"	TEXT,
                "Fecha próxima cita"	DATE
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS IndicadoresMetas (
                PERIODO TEXT,
                Indicador TEXT,
                Meta REAL,
                'Porcentaje de desviación de la meta' REAL,
                Rango REAL
            )
        ''')
        cursor.close()
        conn.close()



    sqlconn = sqlite3.connect(db_path)
    cursor = sqlconn.cursor()

    indServicio.to_sql('IndicadoresServicio', sqlconn, if_exists='append', index=False)
    indEspecialidad.to_sql('IndicadoresEspecialidad', sqlconn, if_exists='append', index=False)
    indDoctor.to_sql('IndicadoresDoctor', sqlconn, if_exists='append', index=False)
    indReferencias.to_sql('IndicadoresReferencias', sqlconn, if_exists='append', index=False)
    indListas.to_sql('IndicadoresListas', sqlconn, if_exists='append', index=False)
    listMetas.to_sql('IndicadoresMetas', sqlconn, if_exists='append', index=False)

    cursor.close()
    sqlconn.close()

def main(page: Page):
    
    fxls = False

    def close_success_dialog(e):
        success_dialog.open = False
        page.update()
    
    def close_error_dialog(e):
        error_dialog.open = False
        page.update()

    def close_missing_data_dialog(e):
        missing_data_dialog.open = False
        page.update()

    def show_success_dialog():
        page.dialog = success_dialog
        success_dialog.open = True
        page.update()

    def show_error_dialog():
        page.dialog = error_dialog
        error_dialog.open = True
        page.update()

    def show_missing_data_dialog():
        page.dialog = missing_data_dialog
        missing_data_dialog.open = True
        page.update()


    def confirm_delete_dialog(page):
        periodo = mes_dropdown.current.value +" "+ yr_dropdown.current.value
        def close_dialog(e):
            dialog.open = False
            page.update()

        def on_yes(e):
            close_dialog(e)
            borrardatos(periodo) 

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Confirmación"),
            content=ft.Text("Ya existen indicadores para ese periodo, ¿desea borrarlos?"),
            actions=[
                ft.TextButton("No", on_click=close_dialog),
                ft.TextButton("Sí", on_click=on_yes)
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )

        page.dialog = dialog
        dialog.open = True
        page.update()

    def validar_carga():
        return not (excel_form_path.current and mes_dropdown.current.value and yr_dropdown.current.value)
            
    # Seleccionar archivos de excel
    
    def select_form(e: FilePickerResultEvent):
        excel_form.value = (
            ", ".join(map(lambda f: f.name, e.files)) if e.files else "Cancelado"
        )
        
        excel_form.update()
        excel_form_path.current = e.files[0].path.replace("\\", "/") if excel_form.value != "Cancelado" else None 
        cargar_ind.current.disabled = validar_carga()
        page.update()

    def dropdown_change(e):
        cargar_ind.current.disabled = validar_carga()
        page.update()

    def func_ind(e):
        periodo = mes_dropdown.current.value +" "+ yr_dropdown.current.value
        val_file = validararchivo(excel_form_path.current)
        if val_file == 'comp':
            if validardatabase(periodo):
                confirm_delete_dialog(page)
            else:
                cargar_ind.current.disabled = True
                calcIndicadores(periodo,excel_form_path.current)
                show_success_dialog()
        elif val_file == 'data':
            show_missing_data_dialog()
        else:
            show_error_dialog()

    page.title = "Herramienta Programada"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.theme_mode = 'light'
    page.padding = 20
    page.window_width = 700
    page.window_height = 450

    cargar_ind = Ref[ElevatedButton]()

    excel_form_path = Ref[str]()


    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    yrs = [str(year) for year in range(2023, 2033)]
    mes_dropdown = Ref[ft.Dropdown]()
    yr_dropdown = Ref[ft.Dropdown]()

    success_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("Completado!"),
        content=ft.Text("Se cargaron los indicadores correctamente!"),
        actions=[
            ft.TextButton("OK", on_click=close_success_dialog)
        ],
        actions_alignment=ft.MainAxisAlignment.END)
    
    error_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("Error!"),
        content=ft.Text("Archivo incorrecto, favor validar"),
        actions=[
            ft.TextButton("OK", on_click=close_error_dialog)
        ],
        actions_alignment=ft.MainAxisAlignment.END)

    missing_data_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("Error!"),
        content=ft.Text("Faltan datos en el formulario, favor validar."),
        actions=[
            ft.TextButton("OK", on_click=close_missing_data_dialog)
        ],
        actions_alignment=ft.MainAxisAlignment.END)
    
    select_first_file = FilePicker(on_result=select_form)
    excel_form = Text()

    page.overlay.extend([select_first_file])

    page.add(
        Row(
            [
            ft.Icon(ft.icons.ACCOUNT_BALANCE, color="black", size=40),
            ft.Text("Sistema de indicadores Servicio de Odontología", color="#0D5382", size=25, weight=ft.FontWeight.BOLD)
            ],
            alignment=ft.MainAxisAlignment.CENTER
        ),
        Row(
            [
                Text(value="Favor indicar el mes y el año correspondientes",width = 300)]),
        Row([
        ft.Dropdown(
            ref = mes_dropdown,
            label="Mes",
            on_change=dropdown_change, 
            options=[ft.dropdown.Option(mes) for mes in meses]
            )
        ,
        ft.Dropdown(
            ref = yr_dropdown,
            label="Año",
            on_change=dropdown_change,  
            options=[ft.dropdown.Option(yr) for yr in yrs]
            )
        ])
        ,

        Row(
            [
                Text(value="Favor seleccionar el archivo",width = 275)]),  
        Row(
            [
                Text(value="Cargar el Formulario de datos:",width = 275),
                ElevatedButton(
                    "Cargar Archivo", width = 200,
                    icon=icons.UPLOAD_FILE,
                    on_click=lambda _: select_first_file.pick_files(
                        allow_multiple=False
                    ),
                ),
                excel_form,
            ]
        ),
        Row(
            [
                Text(value="Ejecutar Herramienta Programada",width = 275),
                ElevatedButton(
                    "Cargar Indicadores", width = 200,
                    ref=cargar_ind,
                    icon=icons.UPLOAD,
                    on_click=func_ind,
                    disabled=True,
                    )
            ]
        )       
    )
ft.app(target=main)
