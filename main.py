import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from lectura_entradas import CAMPOS, ErrorEntrada, leer_entradas
from reporte_excel import generar_excel


def validar_fecha(fecha):
    if not fecha.strip():
        return None, "RN02: fecha obligatoria"

    if re.fullmatch(r"[0-9]{2}/[0-9]{2}/[0-9]{4}", fecha):
        formato = "%d/%m/%Y"
    elif re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}", fecha):
        formato = "%Y-%m-%d"
    else:
        return None, "RN02: la fecha debe tener formato DD/MM/AAAA o AAAA-MM-DD"

    try:
        return datetime.strptime(fecha, formato).date(), None
    except ValueError:
        return None, "RN02: la fecha no existe en el calendario"


def normalizar_fecha(fecha_original):
    fecha, error = validar_fecha(fecha_original)
    if error:
        return None, "No aplica"

    normalizada = f"{fecha.day:02d}/{fecha.month:02d}/{fecha.year:04d}"
    estado = "Sin conversión" if normalizada == fecha_original else "Convertida"
    return normalizada, estado


def convertir_importe(importe):
    texto = importe.strip()
    if not re.fullmatch(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", texto):
        return None
    try:
        return Decimal(texto)
    except InvalidOperation:
        return None


def validar_registro(id_venta, fecha, cliente, importe):
    errores = []

    if not id_venta.strip():
        errores.append("RN01: id obligatorio")

    _, error_fecha = validar_fecha(fecha)
    if error_fecha:
        errores.append(error_fecha)

    if not cliente.strip():
        errores.append("RN03: cliente obligatorio")

    if not importe.strip():
        errores.append("RN04: importe obligatorio")
    elif convertir_importe(importe) is None:
        errores.append("RN04: el importe debe ser numérico")

    return errores


def clasificar_registros(ventas):
    tiene_origen = "archivo_origen" in ventas and "fila_origen" in ventas
    if tiene_origen:
        campos = ventas[[*CAMPOS, "archivo_origen", "fila_origen"]]
    else:
        campos = ventas[list(CAMPOS)].assign(archivo_origen=None, fila_origen=None)
    resultados = []
    datos_por_id = defaultdict(set)
    primera_fila_por_id = {}
    for fila, (id_venta, fecha, cliente, importe, archivo_origen, fila_origen) in enumerate(
        campos.itertuples(index=False, name=None), start=1
    ):
        errores = validar_registro(id_venta, fecha, cliente, importe)
        normalizada, conversion = normalizar_fecha(fecha)
        importe_decimal = convertir_importe(importe)
        resultados.append({
            "fila": fila,
            "archivo_origen": archivo_origen if tiene_origen else None,
            "fila_origen": fila_origen if tiene_origen else fila,
            "id": id_venta,
            "fecha_original": fecha,
            "fecha_normalizada": normalizada,
            "conversion": conversion,
            "cliente": cliente,
            "importe": importe,
            "importe_decimal": importe_decimal,
            "motivos": errores,
        })
        if id_venta.strip():
            datos = (
                normalizada or fecha,
                cliente,
                importe if importe_decimal is None else importe_decimal,
            )
            datos_por_id[id_venta].add(datos)
            primera_fila_por_id.setdefault(id_venta, fila)

    for registro in resultados:
        id_venta = registro["id"]
        motivos = registro["motivos"]
        if len(datos_por_id.get(id_venta, ())) > 1:
            motivos.append("RN06: mismo ID con datos diferentes")

        if motivos:
            registro["categoria"] = "A revisar"
        elif registro["fila"] != primera_fila_por_id[id_venta]:
            registro["categoria"] = "Duplicado exacto"
            referencia = resultados[primera_fila_por_id[id_venta] - 1]
            registro["archivo_referencia"] = referencia["archivo_origen"]
            registro["fila_referencia"] = referencia["fila_origen"]
            ubicacion = f"fila {referencia['fila_origen']}"
            if referencia["archivo_origen"] is not None:
                ubicacion += f" de {referencia['archivo_origen']}"
            motivos.append(f"RN05: datos equivalentes a la {ubicacion}")
        else:
            registro["categoria"] = "Válido"
            motivos.append("Cumple las reglas aplicables")

    return resultados


def mostrar_clasificacion(resultados):
    print("\nClasificación de registros")
    for registro in resultados:
        fecha = registro["fecha_original"]
        normalizada = registro["fecha_normalizada"]
        detalle_fecha = (
            f"Fecha original: {fecha!r} | Fecha normalizada: {normalizada or '-'}"
            f" | Conversión: {registro['conversion']}"
        )
        motivos = "; ".join(registro["motivos"])
        print(
            f"{registro['archivo_origen']}, fila {registro['fila_origen']} (id={registro['id']!r})"
            f" | {registro['categoria']} | {motivos} | {detalle_fecha}"
        )

    cantidades = Counter(registro["categoria"] for registro in resultados)
    print("\nResumen por categoría")
    for categoria in ("Válido", "Duplicado exacto", "A revisar"):
        print(f"{categoria}: {cantidades[categoria]}")


def consolidar_entrada(carpeta, destino):
    ventas, cantidad_archivos, avisos = leer_entradas(carpeta)
    for aviso in avisos:
        print(f"Aviso: {aviso}")
    resultados = clasificar_registros(ventas)
    generar_excel(resultados, destino, cantidad_archivos)
    return resultados


def main():
    ejecutable = sys.executable if getattr(sys, "frozen", False) else __file__
    raiz = Path(ejecutable).resolve().parent
    destino = raiz / "salida" / "resultado_consolidacion.xlsx"
    try:
        resultados = consolidar_entrada(raiz / "entrada", destino)
    except (ErrorEntrada, OSError) as error:
        print(f"No se generó el Excel: {error}")
        return 1
    mostrar_clasificacion(resultados)
    print(f"\nExcel generado: {destino}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
