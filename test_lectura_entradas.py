import unittest
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

from lectura_entradas import ErrorEntrada, leer_entradas
from main import consolidar_entrada


ENCABEZADO = "id,fecha,cliente,importe\n"


class EntradaTest(unittest.TestCase):
    def setUp(self):
        temporal = TemporaryDirectory(dir=Path(__file__).resolve().parent)
        self.addCleanup(temporal.cleanup)
        self.entrada = Path(temporal.name) / "entrada"
        self.entrada.mkdir()
        self.destino = Path(temporal.name) / "salida" / "resultado.xlsx"

    def test_consolida_entre_formatos_con_trazabilidad_y_total_global(self):
        csv = self.entrada / "A.csv"
        csv.write_text(
            ENCABEZADO + '001,2026-11-01,"Ana\nLópez",-15.00\n\n'
            "002,2026-11-02,Bruno,100\n",
            encoding="utf-8",
            newline="",
        )
        xlsx = self.entrada / "b.XLSX"
        libro = Workbook()
        libro.active.append(["importe", "cliente", "id", "fecha"])
        libro.active.append([-15, "Ana\nLópez", "001", date(2026, 11, 1)])
        libro.active.append([101, "Bruno", "002", date(2026, 11, 2)])
        libro.active.append(["20.50", "Carla", "003", "03/11/2026"])
        libro.create_sheet("Ignorada").append(["datos ajenos al consolidado"])
        libro.save(xlsx)
        originales = [ruta.read_bytes() for ruta in (csv, xlsx)]

        ventas, cantidad, avisos = leer_entradas(self.entrada)
        self.assertEqual(cantidad, 2)
        self.assertIn("precisión", avisos[0])
        self.assertEqual(list(ventas["fila_origen"]), [2, 5, 2, 3, 4])
        self.assertEqual(list(ventas["importe"]), ["-15.00", "100", "-15", "101", "20.50"])

        resultados = consolidar_entrada(self.entrada, self.destino)
        self.assertEqual([r["categoria"] for r in resultados],
                         ["Válido", "A revisar", "Duplicado exacto", "A revisar", "Válido"])
        libro = load_workbook(self.destino)
        self.assertEqual(libro.sheetnames, ["Consolidado", "A revisar", "Duplicados", "Resumen"])
        self.assertEqual(libro["Consolidado"]["F2"].value, "A.csv")
        self.assertEqual(libro["A revisar"]["A2"].value, 5)
        self.assertEqual(libro["A revisar"]["G3"].value, "b.XLSX")
        duplicado = libro["Duplicados"][2]
        self.assertEqual([duplicado[i].value for i in (0, 5, 7, 8)], [2, 2, "b.XLSX", "A.csv"])
        resumen = dict(libro["Resumen"].values)
        self.assertEqual(resumen["Archivos procesados"], 2)
        self.assertEqual(resumen["Total procesado"], 5)
        self.assertEqual(resumen["Suma de importes válidos"], 5.5)
        libro.close()
        self.assertEqual([ruta.read_bytes() for ruta in (csv, xlsx)], originales)

    def test_fallos_identifican_archivos_y_conservan_resultado_previo(self):
        (self.entrada / "a.csv").write_text(
            ENCABEZADO + "001,2026-11-01,Ana,100\n", encoding="utf-8"
        )
        consolidar_entrada(self.entrada, self.destino)
        anterior = self.destino.read_bytes()
        (self.entrada / "b.xlsx").write_bytes(b"no es un XLSX")
        (self.entrada / "c.csv").write_text("id,fecha\n", encoding="utf-8")
        with self.assertRaises(ErrorEntrada) as fallo:
            consolidar_entrada(self.entrada, self.destino)
        self.assertIn("b.xlsx:", str(fallo.exception))
        self.assertIn("c.csv: columnas faltantes: cliente, importe", str(fallo.exception))
        self.assertEqual(self.destino.read_bytes(), anterior)

    def test_sin_archivos_compatibles_no_genera_resultado(self):
        (self.entrada / "notas.txt").write_text("texto", encoding="utf-8")
        with self.assertRaisesRegex(ErrorEntrada, "no hay archivos"):
            consolidar_entrada(self.entrada, self.destino)
        self.assertFalse(self.destino.exists())
