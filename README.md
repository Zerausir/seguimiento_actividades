# Procesador de Seguimiento - README

## Descripción General

Este proyecto proporciona un script en Python para procesar archivos de Excel, actualizando información estructurada y
aplicando formato condicional y reglas específicas.

## Características

- Procesa datos de seguimiento en un archivo Excel.
- Actualiza columnas y aplica formato condicional según reglas predefinidas.
- Valida la existencia de archivos en rutas configuradas mediante variables de entorno.
- Exporta el archivo actualizado a una ubicación definida con formato ajustado y filtros habilitados.

## Requisitos

- Python 3.8 o superior.
- Dependencias (ver `requirements.txt`):
    - pandas
    - openpyxl
    - environs
    - glob

## Instalación

1. Clona el repositorio.
2. Crea y activa un entorno virtual:
   ```bash
   python -m venv venv
   source venv/bin/activate  # En Windows: venv\Scripts\activate
   ```
3. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

## Configuración

Asegúrate de que las siguientes variables de entorno estén configuradas en un archivo `.env` en el directorio raíz:

```env
SERVER_ROUTE=ruta/a/directorio/de/entrada
DOWNLOAD_ROUTE=ruta/a/directorio/de/salida
FILE_SEGUIMIENTO_GPR=nombre_del_archivo.xlsx
RESULTS=nombre_del_archivo_resultante.xlsx
RUTA_CRONOLOGICO_2023=ruta/a/cronologico/2023
RUTA_CRONOLOGICO_2024=ruta/a/cronologico/2024
RUTA_CRONOLOGICO_2025=ruta/a/cronologico/2025
RUTA_<INDICADOR>_<AÑO>=ruta_del_indicador_para_año
RUTA_<INDICADOR>_<MES>_<AÑO>=ruta_mensual_del_indicador
```

- `SERVER_ROUTE`: Directorio que contiene el archivo de seguimiento.
- `DOWNLOAD_ROUTE`: Directorio donde se guardará el archivo actualizado.
- `FILE_SEGUIMIENTO_GPR`: Nombre del archivo de seguimiento.
- `RESULTS`: Nombre del archivo resultante.

## Uso

1. Coloca el archivo de seguimiento en la ruta definida por `SERVER_ROUTE`.
2. Ejecuta el script:
   ```bash
   python main.py
   ```
3. El archivo actualizado será guardado en la ruta definida por `DOWNLOAD_ROUTE` con el nombre configurado en `RESULTS`.

## Estructura del Código

### Funciones Principales

#### `process_seguimiento(file_path)`

- Lee y actualiza el archivo Excel según reglas predefinidas.
- Aplica formato condicional para resaltar valores relevantes.
- Guarda el archivo procesado en la ubicación definida.

#### `check_file_exists(informe_num, indicator, year)`

- Verifica la existencia de archivos en rutas cronológicas, específicas y mensuales.

#### `verify_environment_variables()`

- Valida que todas las variables de entorno necesarias estén configuradas.

### Utilidades

- Ajuste automático del ancho de las columnas.
- Habilitación de filtros y congelación de la primera fila en el archivo Excel resultante.

## Manejo de Errores

- Captura excepciones durante la verificación de rutas y variables de entorno.
- Lanza un error detallado si faltan variables de entorno críticas.

## Ejemplo de Salida

Un archivo Excel con las siguientes columnas:

- `ATENDIDO`: Indicador actualizado según las reglas de procesamiento.
- `ATENDIDO FUENTE`: Copia original de la columna `ATENDIDO`.
- `INDICADOR AÑO`: Año asociado al indicador.
- `INDICADOR AÑO ENCONTRADO`: Año validado tras la búsqueda de archivos.

## Contribución

1. Haz un fork del repositorio.
2. Crea una rama para tu funcionalidad:
   ```bash
   git checkout -b nueva-funcionalidad
   ```
3. Realiza tus cambios y realiza un commit:
   ```bash
   git commit -m "Agrega nueva funcionalidad"
   ```
4. Envía tus cambios:
   ```bash
   git push origin nueva-funcionalidad
   ```
5. Abre un pull request.

## Librerias

- Libraries
  used: [pandas](https://pandas.pydata.org/), [openpyxl](https://openpyxl.readthedocs.io/), [environs](https://pypi.org/project/environs/).

## Licencia

Este proyecto está bajo la licencia MIT. Consulta el archivo `LICENSE` para más detalles.

## Contacto

Para consultas o comentarios, por favor contacta a [Iván Suárez](https://github.com/Zerausir).
