import csv
from datetime import datetime
from zipfile import BadZipFile

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


CAMPOS = ("id", "fecha", "cliente", "importe")


class ErrorEntrada(Exception):
    pass


def posiciones_columnas(encabezado):
    faltantes = [campo for campo in CAMPOS if campo not in encabezado]
    if faltantes:
        raise ErrorEntrada(f"columnas faltantes: {', '.join(faltantes)}")
    repetidas = [campo for campo in CAMPOS if encabezado.count(campo) > 1]
    if repetidas:
        raise ErrorEntrada(f"columnas repetidas: {', '.join(repetidas)}")
    return [encabezado.index(campo) for campo in CAMPOS]


def leer_csv(archivo):
    with archivo.open(encoding="utf-8-sig", newline="") as entrada:
        lector = csv.reader(entrada, strict=True)
        encabezado = next(lector, [])
        posiciones = posiciones_columnas(encabezado)
        registros = []
        while True:
            fila = lector.line_num + 1
            valores = next(lector, None)
            if valores is None:
                return registros
            if not valores:
                continue
            if len(valores) != len(encabezado):
                raise ErrorEntrada(f"fila {fila}: cantidad de campos distinta del encabezado")
            registros.append((*[valores[i] for i in posiciones], archivo.name, fila))


def texto_excel(valor):
    if valor is None:
        return ""
    if isinstance(valor, datetime) and valor == datetime.combine(valor.date(), datetime.min.time()):
        return valor.date().isoformat()
    return str(valor)


def leer_xlsx(archivo):
    libro = load_workbook(archivo, read_only=True)
    try:
        if not libro.worksheets:
            raise ErrorEntrada("el libro no contiene hojas de datos")
        hoja = libro.worksheets[0]
        hoja.reset_dimensions()
        filas = hoja.iter_rows()
        encabezado = [celda.value for celda in next(filas, ())]
        posiciones = posiciones_columnas(encabezado)
        registros = []
        tiene_valores_tipados = False
        tiene_formulas = False
        for fila, celdas in enumerate(filas, start=2):
            if all(celda.value is None for celda in celdas):
                continue
            seleccion = [celdas[i] for i in posiciones if i < len(celdas)]
            valores = [celdas[i].value if i < len(celdas) else None for i in posiciones]
            tiene_valores_tipados |= any(v is not None and not isinstance(v, str) for v in valores)
            tiene_formulas |= any(celda.data_type == "f" for celda in seleccion)
            registros.append((*map(texto_excel, valores), archivo.name, fila))
        return registros, tiene_valores_tipados, tiene_formulas
    finally:
        libro.close()


def leer_entradas(carpeta):
    archivos = sorted(
        (ruta for ruta in carpeta.iterdir()
         if ruta.is_file() and ruta.suffix.lower() in {".csv", ".xlsx"}
         and not ruta.name.startswith("~$")),
        key=lambda ruta: (ruta.name.casefold(), ruta.name),
    )
    if not archivos:
        raise ErrorEntrada("no hay archivos CSV ni XLSX compatibles en entrada")

    registros = []
    avisos = []
    errores = []
    for archivo in archivos:
        try:
            if archivo.suffix.lower() == ".csv":
                registros.extend(leer_csv(archivo))
            else:
                filas, tipados, formulas = leer_xlsx(archivo)
                registros.extend(filas)
                if tipados:
                    avisos.append(
                        f"{archivo.name}: las celdas tipadas se leyeron como texto a partir del "
                        "valor guardado. Excel puede haber perdido ceros, formato o precisión."
                    )
                if formulas:
                    avisos.append(f"{archivo.name}: las fórmulas se leyeron como texto sin calcular.")
        except (ErrorEntrada, OSError, UnicodeError, ValueError, csv.Error, BadZipFile,
                InvalidFileException, KeyError, TypeError, SyntaxError) as error:
            errores.append(f"{archivo.name}: {error}")
    if errores:
        raise ErrorEntrada("\n".join(errores))

    columnas = (*CAMPOS, "archivo_origen", "fila_origen")
    return pd.DataFrame(registros, columns=columnas), len(archivos), avisos
