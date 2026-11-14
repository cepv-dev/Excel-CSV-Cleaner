from collections import Counter
from datetime import datetime
from decimal import Decimal, localcontext
from pathlib import Path
from tempfile import NamedTemporaryFile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


FORMATO_IMPORTE = "#,##0.00##########"


def escribir_fila(hoja, valores):
    hoja.append(valores)
    for celda in hoja[hoja.max_row]:
        if isinstance(celda.value, str):
            celda.data_type = "s"


def preparar_hoja(hoja, encabezados, anchos):
    escribir_fila(hoja, encabezados)
    hoja.freeze_panes = "A2"
    hoja.row_dimensions[1].height = 24
    for celda, ancho in zip(hoja[1], anchos):
        celda.fill = PatternFill("solid", fgColor="183C52")
        celda.font = Font(bold=True, color="FFFFFF")
        celda.alignment = Alignment(vertical="center")
        hoja.column_dimensions[get_column_letter(celda.column)].width = ancho


def sumar_importes(importes):
    importes = [importe for importe in importes if importe]
    if not importes:
        return Decimal(0)
    # Reservar espacio para todos los decimales y para el acarreo de la suma.
    mayor_posicion = max(importe.adjusted() for importe in importes)
    menor_posicion = min(importe.as_tuple().exponent for importe in importes)
    with localcontext() as contexto:
        contexto.prec = mayor_posicion - menor_posicion + 1 + len(str(len(importes)))
        return sum(importes, Decimal(0))


def generar_excel(resultados, destino, cantidad_archivos):
    libro = Workbook()
    consolidado = libro.active
    consolidado.title = "Consolidado"
    revision = libro.create_sheet("A revisar")
    duplicados = libro.create_sheet("Duplicados")
    resumen = libro.create_sheet("Resumen")

    preparar_hoja(consolidado, ("Fila origen", "ID", "Fecha normalizada", "Cliente", "Importe",
                              "Archivo origen"), (15, 15, 22, 26, 20, 28))
    encabezados = ("Fila origen", "ID", "Fecha original", "Cliente", "Importe original")
    anchos = (13, 15, 21, 26, 20)
    preparar_hoja(revision, (*encabezados, "Motivos", "Archivo origen"), (*anchos, 68, 28))
    preparar_hoja(duplicados, (*encabezados, "Fila de referencia", "Motivos",
                              "Archivo origen", "Archivo de referencia"), (*anchos, 22, 52, 28, 28))
    preparar_hoja(resumen, ("Indicador", "Valor"), (34, 22))

    importes_validos = []
    for registro in resultados:
        fila, id_venta, cliente = registro["fila_origen"], registro["id"], registro["cliente"]
        archivo_origen = registro["archivo_origen"]
        motivos = "\n".join(registro["motivos"])
        originales = (fila, id_venta, registro["fecha_original"], cliente, registro["importe"])
        if registro["categoria"] == "Válido":
            normalizada = registro["fecha_normalizada"]
            fecha = datetime.strptime(normalizada, "%d/%m/%Y").date()
            # El calendario de Excel empieza en 1900.
            fecha_excel = fecha if fecha.year >= 1900 else normalizada
            importe = registro["importe_decimal"]
            escribir_fila(consolidado, (fila, id_venta, fecha_excel, cliente, importe, archivo_origen))
            consolidado.cell(consolidado.max_row, 3).number_format = "dd/mm/yyyy"
            consolidado.cell(consolidado.max_row, 5).number_format = FORMATO_IMPORTE
            importes_validos.append(importe)
        else:
            if registro["categoria"] == "A revisar":
                hoja = revision
                escribir_fila(hoja, (*originales, motivos, archivo_origen))
                columna_motivos = 6
            else:
                hoja = duplicados
                escribir_fila(hoja, (*originales, registro["fila_referencia"], motivos,
                                    archivo_origen, registro["archivo_referencia"]))
                columna_motivos = 7
            hoja.cell(hoja.max_row, columna_motivos).alignment = Alignment(wrap_text=True, vertical="top")

    cantidades = Counter(registro["categoria"] for registro in resultados)
    convertidas = sum(registro["conversion"] == "Convertida" for registro in resultados)
    for indicador, valor in (
        ("Archivos procesados", cantidad_archivos),
        ("Total procesado", len(resultados)),
        ("Válidos", cantidades["Válido"]),
        ("Duplicados", cantidades["Duplicado exacto"]),
        ("A revisar", cantidades["A revisar"]),
        ("Fechas convertidas", convertidas),
        ("Suma de importes válidos", sumar_importes(importes_validos)),
    ):
        escribir_fila(resumen, (indicador, valor))
    resumen.cell(resumen.max_row, 2).number_format = FORMATO_IMPORTE
    for hoja in libro:
        hoja.auto_filter.ref = hoja.dimensions

    destino = Path(destino)
    destino.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=destino.parent, suffix=".xlsx", delete=False) as archivo:
        temporal = Path(archivo.name)
    try:
        libro.save(temporal)
        temporal.replace(destino)
    finally:
        temporal.unlink(missing_ok=True)
