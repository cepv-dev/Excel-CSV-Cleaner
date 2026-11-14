# Excel/CSV Cleaner

Consolida archivos de ventas CSV y XLSX en un Excel con cuatro hojas: **Consolidado**, **Duplicados**, **A revisar** y **Resumen**. Cada registro conserva su archivo y fila de origen. Los duplicados remiten a la primera aparición y, si un ID tiene datos incompatibles, todas sus filas quedan a revisión.

## Demostración

La carpeta [prueba_aceptacion](prueba_aceptacion) incluye dos CSV, un XLSX y el reporte resultante. El ejemplo procesa **12 registros en 3 archivos**, con **5 válidos, 2 duplicados, 5 a revisar y 430,10 de importe válido**. El ID 900 aparece en dos archivos con importes diferentes y el reporte identifica ambas filas.

## Ejecutar desde el código

Con Python 3.13, ejecutá en la carpeta del proyecto:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
New-Item -ItemType Directory -Force entrada
Copy-Item .\prueba_aceptacion\entrada\* .\entrada\
.\.venv\Scripts\python.exe main.py
```

El resultado se guarda en `salida/resultado_consolidacion.xlsx`. Para otros datos, usá las columnas `id`, `fecha`, `cliente` e `importe`. Los CSV deben estar en UTF-8 y separados por comas. En XLSX se procesa la primera hoja.

## Pruebas y compilación en Windows

```powershell
.\.venv\Scripts\python.exe -m unittest discover
.\.venv\Scripts\python.exe -m pip install pyinstaller==6.22.3
.\.venv\Scripts\python.exe -m PyInstaller Excel-CSV-Cleaner.spec
```

El ejecutable queda en `dist` y usa las carpetas `entrada` y `salida` a su lado. Se probó en Windows de 64 bits con Python fuera del PATH. La prueba en una instalación sin Python queda pendiente.

## Servicio freelance

Para evaluar una consolidación similar, compartí dos o tres archivos de muestra con datos ficticios o anonimizados, indicá qué columna identifica cada venta y cómo querés tratar importes distintos para un mismo ID. Con esa muestra se puede definir el alcance y cotizar la entrega.
