# Herramienta de Gestión de Indicadores de Desempeño - Servicio de Odontología

## Descripción
Esta herramienta de Python ha sido desarrollada como parte de un proyecto para mejorar la gestión de indicadores de desempeño en el Servicio de Odontología. Facilita la entrada, procesamiento y visualización de datos relacionados con estos indicadores, con el objetivo de optimizar el tiempo requerido para estas actividades y apoyar la toma de decisiones informadas.

## Características
- **Interfaz Gráfica**: Permite la selección fácil y rápida de periodos y archivos Excel para el análisis.
- **Procesamiento Automático**: Calcula los indicadores de desempeño a partir de los datos ingresados.
- **Generación de Reportes**: Produce reportes detallados basados en los indicadores calculados.
- **Validaciones de Datos**: Incluye mecanismos para prevenir duplicados en la base de datos y verificar la pertinencia de los archivos subidos.

## Instalación y Uso

### Requisitos Previos
- **Python**: Es necesario asegurarse de tener Python instalado en el sistema. Esta herramienta ha sido desarrollada y probada en Python 3.11.4
- **Bibliotecas de Python**: La herramienta depende de varias bibliotecas, incluyendo `pandas` 1.5.3, `sqlite3` 3.41.2, `flet` 0.12.2 y `os`. Estas deben ser instaladas antes de ejecutar la aplicación.
pip install pandas flet
Nota: Las bibliotecas `sqlite3` y `os` ya están incluidas en la instalación estándar de Python.

### Instalación
1. **Clonar o Descargar**: Clonar este repositorio o descargae el archivo `ProyectoTFG.py` en el directorio de trabajo.
2. **Instalar Dependencias**: Ejecutar el siguiente comando para instalar las dependencias necesarias:

### Uso
1. **Ejecución del Script**: Ejecutar el script `ProyectoTFG.py` en el entorno Python. Esto iniciará la interfaz gráfica de usuario desarrollada con `flet`.
2. **Selección de Archivos y Periodos**: Utiliza la interfaz gráfica para seleccionar el archivo de Excel (.xlsx) y definir el periodo para el análisis de los indicadores.
3. **Validación de Datos**: La herramienta realiza validaciones automáticas del archivo subido para asegurar que estén en el formato correcto y prevenir duplicados en la base de datos.
4. **Cálculo de indicadores**: Una vez procesados los datos, la herramienta calcula indicadores con base en los datos, e inserta los resultados en la base de datos `PruebasTFG.db`.

### Notas Adicionales
- La base de datos utilizada por la herramienta es `PruebasTFG.db`, es necesario validar que esté disponible en el directorio de trabajo.

## Contribuir
Contibuciones a este reporsitorio solo son admitidas por parte de los estudiantes encargados del desarrollo, el asesor técnico, o el director del Trabajo Final de Graduación.

## Licencia
Este proyecto fue desarrollado como producto de un Trabajo Final de Graduación en la Universidad de Costa Rica, por lo que no debe ser distribuido abiertamente..

