import pandas as pd
import calendar
import sqlite3
import pandas as pd
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

def prosFomularioAdicional(filepath):

    xls = pd.ExcelFile(filepath)
    dfs = {}
    for sheet_name in xls.sheet_names:
        dfs[sheet_name] = pd.read_excel(xls, sheet_name)
    return dfs

def prosConsolidado(filepath):
    xls = pd.ExcelFile(filepath)
    sheet_names = xls.sheet_names
    result_dfs = {}

    for sht in sheet_names:
        df = pd.read_excel(filepath, sheet_name=sht)

        # Evaluar si la hoja es la mensual o es de especialidad, 
        #** Agregar validaciones adicionales para asegurar que solo se procese el file correcto
        if sht == "Mensual":   
            tipo_hoja = "Mensual"
        else:
            tipo_hoja = "Especialidad"

        # Eliminar filas y columnas vacías

        df.dropna(axis = 0,how = "all",inplace = True)  
        df.dropna(axis = 1,how = "all",inplace = True)  
        df = df.reset_index(drop=True)

        # Eliminar fila extra de la hoja "Mensual"

        if tipo_hoja == "Mensual":
            df.drop(2, inplace = True)
            df = df.reset_index(drop=True)

        # La hoja "Especialidad" tiene una columna adicional que permite identificar la tabla donde se encuentra el agregado de datos 
        # Se eliminan las otras filas y se actualiza el df al tamaño correcto para procesarlo

        elif tipo_hoja == "Especialidad":
            indexnum = int(df.loc[df.iloc[:,0] == "TOTAL"].index[0])

            df = df.iloc[indexnum:]

            nombre_esp = df.iloc[0,1]
            codigo_esp = sht  #depende de cada hoja
            df = df.iloc[:,1:] 

        # Generar una tabla(Posible serie), de las horas contratadas, programadas, etc.
        # Se hace un slice de las primeras dos filas, y se eliminan vacios 

        df_jornada = df.iloc[1:3,1:8].copy() 
        df_jornada.dropna(axis = 1,how = "all",inplace = True)
        df_jornada.columns = df_jornada.iloc[0]
        df_jornada = df_jornada[1:]
        df_jornada = df_jornada.reset_index(drop=True)
        df_jornada.columns.rename("",inplace=True)
        df_jornada = df_jornada.transpose()
        df_jornada.rename(columns={0: 'Horas'},inplace= True)

        # Replicar la tabla de excel con las horas utilizadas por actividad
        # Se hace un slice de las primeras filas 3 a 17  , se eliminan vacios y se nombran columnas de acuerdo al excel

        df_tiempo = df.iloc[3:18].copy()
        df_tiempo.dropna(axis = 1,how = "all",inplace = True)
        df_tiempo.dropna(axis = 0,how = "all",inplace = True)
        df_tiempo.columns = ['Actividad', 'Programado', 'Utilizado', 'N° Actividades']
        df_tiempo = df_tiempo.reset_index(drop=True)

        # Replicar la tabla de excel del resumen del mes con los totales de consultas por cada categoria

        # Se hace un slice de las primeras filas 18 en adelante, se eliminan vacios full y se auto completan algunas categorias faltantes
        df_total = df.iloc[18:].copy()
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

        result_dfs[sht] = {
            'df_jornada': df_jornada,
            'df_tiempo': df_tiempo,
            'df_total': df_total
        }
    return result_dfs

def calcIndicadores(periodo,consol,form):
    hojasConsolidado = prosConsolidado(consol)
    hojasFormulario = prosFomularioAdicional(form)
    
    #Calculo de indicadores servicio
    indServicio = pd.DataFrame(columns=['PERIODO', 'INDICADOR', 'VALOR'])

    nuevoIndServ=[]

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'ESTANDAR PACIENTE POR HORA DE LA UNIDAD',
        'VALOR': (hojasFormulario['Consultas Externas']['Consultas programadas en consulta externa'].sum()/hojasFormulario['Consultas Externas']['Horas programadas para consulta externa'].sum())
                if hojasFormulario['Consultas Externas']['Horas programadas para consulta externa'].sum() > 0 else None 
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE PRÓTESIS REALIZADAS TOTAL SUPERIOR',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][0]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE PRÓTESIS REALIZADAS TOTAL INFERIOR',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][1]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE PRÓTESIS REALIZADAS PARCIAL SUPERIOR',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][2]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE PRÓTESIS REALIZADAS PARCIAL INFERIOR',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][3]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE PRÓTESIS REALIZADAS OBTURADORES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][4]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE PRÓTESIS REALIZADAS REPARACIONES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][5]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TOTAL DE APARATOS ORTODONCIA',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][6]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TOTAL DE APARATOS ODONTOPEDIATRIA',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][7]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TOTAL DE PLANOS OCLUSALES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][8]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE REFERENCIAS ACEPTADAS',
        'VALOR': hojasFormulario['Referencias']['Aceptado'].sum()
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CANTIDAD DE REFERENCIAS RECHAZADAS',
        'VALOR': hojasFormulario['Referencias']['Rechazado'].sum()
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CONSULTAS ODONTOLÓGICAS PRIMERA VEZ',
        'VALOR': hojasConsolidado["Mensual"]["df_total"].iloc[0,:5].sum()
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'CONSULTAS ODONTOLÓGICAS SUBSECUENTES',
        'VALOR': hojasConsolidado["Mensual"]["df_total"].iloc[0,5]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TELECONSULTAS',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][9]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'HORAS EN TELECONSULTAS',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][10]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'TELEORIENTACIONES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][11]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'HORAS EN TELEORIENTACIONES',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][12]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'APARATOLOGÍA',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][13]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'HORAS DE HOSPITALIZACIÓN',
        'VALOR': hojasFormulario['Otros Datos']['Resultado'][14]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'NÚMERO DE NIÑOS (AS) DE 0 A MENOS DE 10 AÑOS CON ATENCIÓN ODONTOLÓGICA PREVENTIVA DE PRIMERA VEZ EN EL AÑO',
        'VALOR': hojasConsolidado["Mensual"]["df_total"].iloc[1,7]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'NÚMERO DE ADOLESCENTES DE 10 A MENOS DE 20 AÑOS CON ATENCIÓN ODONTOLÓGICA PREVENTIVA DE PRIMERA VEZ EN EL AÑO',
        'VALOR': hojasConsolidado["Mensual"]["df_total"].iloc[2,7]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'PACIENTES EMBARAZADAS CON ATENCIÓN ODONTOLÓGICA PREVENTIVA DE PRIMERA VEZ EN EL AÑO',
        'VALOR': hojasConsolidado["Mensual"]["df_total"].iloc[0,6]
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'COBERTURA ODONTOLÓGICA EN NIÑOS (AS) DE 0 A MENOS DE 10 AÑOS, EN EL TERCER NIVEL DE ATENCIÓN',
        'VALOR': (hojasConsolidado["Mensual"]["df_total"].iloc[1,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][17])
                if hojasFormulario['Otros Datos']['Resultado'][17] > 0 else None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'COBERTURA ODONTOLÓGICA EN ADOLESCENTES DE 10 A MENOS DE 20 AÑOS, EN EL TERCER NIVEL DE ATENCIÓN',
        'VALOR': (hojasConsolidado["Mensual"]["df_total"].iloc[2,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][18])
                if hojasFormulario['Otros Datos']['Resultado'][18] > 0 else None
    })

    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'COBERTURA ODONTOLÓGICA EN HOMBRES DE 20 AÑOS A 64 AÑOS, EN EL TERCER NIVEL DE ATENCIÓN',
        'VALOR': (hojasConsolidado["Mensual"]["df_total"].iloc[3,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][19])
                if hojasFormulario['Otros Datos']['Resultado'][19] > 0 else None
    })


    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'COBERTURA ODONTOLÓGICA EN MUJERES DE 20 AÑOS A 64 AÑOS, EN TERCER NIVEL DE ATENCIÓN',
        'VALOR': (hojasConsolidado["Mensual"]["df_total"].iloc[4,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][20])
                if hojasFormulario['Otros Datos']['Resultado'][20] > 0 else None
    })


    nuevoIndServ.append({
        'PERIODO': periodo,
        'INDICADOR': 'COBERTURA ODONTOLÓGICA EN PERSONAS DE MÁS DE 65 AÑOS, EN EL TERCER NIVEL DE ATENCIÓN',
        'VALOR': (hojasConsolidado["Mensual"]["df_total"].iloc[5,[0,1,3]].sum()/hojasFormulario['Otros Datos']['Resultado'][21])
                if hojasFormulario['Otros Datos']['Resultado'][21] > 0 else None
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
            'INDICADOR': 'HORAS PROGRAMADAS PARA LA ATENCION DE PACIENTES',
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
            'INDICADOR': 'HORAS EJECUTADAS EN LA ATENCION DE PACIENTES',
            'VALOR': hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Horas utilizadas para consulta externa'].sum()
        })

        nuevoIndEsp.append({
            'PERIODO': periodo,
            'ESPECIALIDAD': esp,
            'INDICADOR': 'USUARIOS ATENDIDOS POR HORA PROGRAMADO PARA LA ATENCION DE PACIENTES',
            'VALOR': (hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Consultas realizadas en consulta externa'].sum() /
                    hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Horas programadas para consulta externa'].sum())
                    if hojasFormulario['Consultas Externas'][hojasFormulario['Consultas Externas']['Especialidad'] == esp]['Horas programadas para consulta externa'].sum() > 0 else None
        })

        indEspecialidad = pd.concat([indEspecialidad, pd.DataFrame(nuevoIndEsp)], ignore_index=True)

    
    #Calculo oficial de indicadores profesional
    indDoctor = pd.DataFrame(columns=['PERIODO','PROFESIONAL', 'INDICADOR', 'VALOR'])

    for index, row in hojasFormulario['Consultas Externas'].iterrows():
        nuevoIndDoc = []
        
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'CAPACIDAD MAXIMA PRÁCTICA',
            'VALOR': ( 2 * row['Horas programadas para consulta externa'])
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'CAPACIDAD DE PRODUCCIÓN PREVISTA',
            'VALOR': ( row['Horas programadas para consulta externa'] * hojasFormulario['Metas']['Meta'][2])
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'PRODUCCIÓN REAL',
            'VALOR': row['Consultas realizadas en consulta externa']
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'USUARIOS ATENDIDOS POR HORA PROGRAMADO PARA LA ATENCION DE PACIENTES',
            'VALOR': (row['Consultas realizadas en consulta externa']/row['Horas programadas para consulta externa'])
                    if row['Horas programadas para consulta externa'] > 0 else None
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'USUARIOS POR HORA UTILIZADA EN ATENCION DE PACIENTES',
            'VALOR': (row['Consultas realizadas en consulta externa']/row['Horas utilizadas para consulta externa'])
                    if row['Horas utilizadas para consulta externa'] > 0 else None
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'AUSENTISMO CONSULTA EXTERNA',
            'VALOR': (row['Citas perdidas en consulta externa']/(row['Consultas realizadas en consulta externa']+row['Citas perdidas en consulta externa']+row['Cupos no utilizados en consulta externa']))
                    if (row['Consultas realizadas en consulta externa']+row['Citas perdidas en consulta externa']+row['Cupos no utilizados en consulta externa']) > 0 else None
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'SUSTITUCIÓN DE PACIENTES CONSULTA EXTERNA',
            'VALOR': (row['Citas sustituidas en consulta externa']/row['Citas perdidas en consulta externa'])
                    if row['Citas perdidas en consulta externa'] > 0 else None
        })

        balPerExt = ((row['Citas perdidas en consulta externa']+row['Cupos no utilizados en consulta externa']) - (row['Citas sustituidas en consulta externa']+row['Recargos en consulta externa']))
    
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'BALANCE DE CUPOS PERDIDOS EN CONSULTA EXTERNA',
            'VALOR': balPerExt 
        })

        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'COSTO POR CUPOS PERDIDOS EN CONSULTA EXTERNA',
            'VALOR': (balPerExt * hojasFormulario['Otros Datos']['Resultado'][15]) 
                    
        })
        indDoctor = pd.concat([indDoctor, pd.DataFrame(nuevoIndDoc)], ignore_index=True)
        
    #Consulta Procedimientos
    for index, row in hojasFormulario['Consultas Procedimientos'].iterrows():
        nuevoIndDoc = []
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'AUSENTISMO CONSULTA PROCEDIMIENTOS',
            'VALOR': (row['Citas perdidas en consulta procedimiento']/(row['Consultas realizadas en consulta procedimiento']+row['Citas perdidas en consulta procedimiento']+row['Cupos no utilizados en consulta procedimiento']))
                    if (row['Consultas realizadas en consulta procedimiento']+row['Citas perdidas en consulta procedimiento']+row['Cupos no utilizados en consulta procedimiento']) > 0 else None
        })
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'SUSTITUCIÓN DE PACIENTES CONSULTA PROCEDIMIENTOS',
            'VALOR': (row['Citas sustituidas en consulta procedimiento']/row['Citas perdidas en consulta procedimiento'])
                    if row['Citas perdidas en consulta procedimiento'] > 0 else None
        })
        balPerProc = ((row['Citas perdidas en consulta procedimiento']+row['Cupos no utilizados en consulta procedimiento'])-(row['Citas sustituidas en consulta procedimiento']+row['Recargos en consulta procedimiento']))
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'BALANCE DE CUPOS PERDIDOS EN CONSULTA PROCEDIMIENTOS',
            'VALOR': balPerProc
        })
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'COSTO POR CUPOS PERDIDOS EN CONSULTA PROCEDIMIENTOS',
            'VALOR': (balPerProc * hojasFormulario['Otros Datos']['Resultado'][16])
        })
        indDoctor = pd.concat([indDoctor, pd.DataFrame(nuevoIndDoc)], ignore_index=True)

    #Consulta Procedimientos
    for index, row in hojasFormulario['Ortodoncia-Ortopedia'].iterrows():
        nuevoIndDoc = []
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'PORCENTAJE DE PACIENTES EN ORTODONCIA',
            'VALOR': row['Porcentaje de pacientes en ortodoncia']
        })
        nuevoIndDoc.append({
            'PERIODO': periodo,
            'PROFESIONAL' : row['Profesional'],
            'INDICADOR': 'PORCENTAJE DE PACIENTES EN ORTOPEDIA',
            'VALOR': row['Porcentaje de pacientes en ortopedia']
        })
        indDoctor = pd.concat([indDoctor, pd.DataFrame(nuevoIndDoc)], ignore_index=True)
        

    indDoctor['original_index'] = indDoctor.index
    indDoctor = indDoctor.sort_values(by=['PROFESIONAL', 'original_index'])
    indDoctor.drop('original_index', axis=1, inplace=True)


    #Referencias por area de salud
    indReferencias = hojasFormulario['Referencias'].iloc[:,1:].groupby('Área de Salud', as_index = False).sum()
    indReferencias.insert(0, 'PERIODO', periodo)

    #Listas de espera
    indListas = hojasFormulario['Listas de espera'].copy()
    indListas.insert(0, 'PERIODO', periodo)
    indListas

    #Metas de indicadores
    listMetas = hojasFormulario['Metas'].copy()
    listMetas.insert(0, 'PERIODO', periodo)


    #Cargar tablas a BD
    sqlconn = sqlite3.connect('PruebasTFG.db')
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
    
    def close_success_dialog(e):
        success_dialog.open = False
        page.update()

    def show_success_dialog():
        page.dialog = success_dialog
        success_dialog.open = True
        page.update()

    def validar_carga():
        return not (excel_consol_path.current and excel_form_path.current and mes_dropdown.current.value and yr_dropdown.current.value)
            
    # Seleccionar archivos de excel
    def select_consol(e: FilePickerResultEvent):
        excel_consol.value = (
            ", ".join(map(lambda f: f.name, e.files)) if e.files else "Cancelado"
        )
      
        excel_consol.update()
        excel_consol_path.current = e.files[0].path.replace("\\", "/") if excel_consol.value != "Cancelado" else None 
        cargar_ind.current.disabled = validar_carga()
        page.update()
    
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
        cargar_ind.current.disabled = True
        periodo = mes_dropdown.current.value +" "+ yr_dropdown.current.value
        calcIndicadores(periodo,excel_consol_path.current,excel_form_path.current)
        show_success_dialog()
        
    page.title = "Herramienta Programada TFG"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.theme_mode = 'light'
    page.padding = 20
    page.window_width = 800
    page.window_height = 500

    cargar_ind = Ref[ElevatedButton]()

    excel_consol_path = Ref[str]()
    excel_form_path = Ref[str]()

    meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    yrs = [str(year) for year in range(2023, 2033)]
    mes_dropdown = Ref[ft.Dropdown]()
    yr_dropdown = Ref[ft.Dropdown]()

    success_dialog = ft.AlertDialog(
        modal=True,
        title=ft.Text("Completado!"),
        content=ft.Text("Los indicadores han sido cargados a la base de datos!"),
        actions=[
            ft.TextButton("OK", on_click=close_success_dialog)
        ],
        actions_alignment=ft.MainAxisAlignment.END)
    
    select_first_file = FilePicker(on_result=select_consol)
    excel_consol = Text()

    select_second_file = FilePicker(on_result=select_form)
    excel_form = Text()

    page.overlay.extend([select_first_file,select_second_file])

    page.add(
        Row(
            [
                Text(value="Favor indicar Mes y Año correspondientes",width = 275)]),
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
                Text(value="Favor seleccionar los archivos",width = 275)]),
        Row(
            [
                Text(value="Cargar el consolidado mensual",width = 275),
                ElevatedButton(
                    "Cargar Archivo", width = 200,
                    icon=icons.UPLOAD_FILE,
                    on_click=lambda _: select_first_file.pick_files(
                        allow_multiple=False
                    ), 
                ),
                excel_consol,
            ]
        ),    
        Row(
            [
                Text(value="Cargar el Formulario de datos adicionales",width = 275),
                ElevatedButton(
                    "Cargar Archivo", width = 200,
                    icon=icons.UPLOAD_FILE,
                    on_click=lambda _: select_second_file.pick_files(
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