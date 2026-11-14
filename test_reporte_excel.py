import unittest
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import load_workbook

from main import clasificar_registros
from reporte_excel import generar_excel, sumar_importes


class ReporteExcelTest(unittest.TestCase):
    def test_exporta_categorias_referencias_y_totales(self):
        ventas = pd.DataFrame(
            [
                ("001", "2026-11-01", "=Ana", "0.10"),
                ("001", "01/11/2026", "=Ana", "0.1"),
                ("002", "02/11/2026", "Bruno", "-0.20"),
                ("003", "2026-11-03", "", "90.00"),
                ("003", "2026-11-03", "Carla", "90.00"),
            ],
            columns=["id", "fecha", "cliente", "importe"],
        )
        raiz = Path(__file__).resolve().parent
        with TemporaryDirectory(dir=raiz) as temporal:
            destino = Path(temporal) / "salida" / "resultado.xlsx"
            generar_excel(clasificar_registros(ventas), destino, cantidad_archivos=1)
            libro = load_workbook(destino)
            self.assertEqual(libro.sheetnames, ["Consolidado", "A revisar", "Duplicados", "Resumen"])
            consolidado, revision, duplicados, resumen = libro.worksheets
            self.assertEqual([hoja.max_row - 1 for hoja in libro], [2, 2, 1, 7])
            self.assertEqual(consolidado["B2"].value, "001")
            self.assertEqual(consolidado["C2"].value, datetime(2026, 11, 1))
            self.assertEqual(consolidado["C2"].number_format, "dd/mm/yyyy")
            self.assertEqual(consolidado["D2"].value, "=Ana")
            self.assertEqual(consolidado["D2"].data_type, "s")
            self.assertEqual(consolidado["E3"].value, -0.2)
            self.assertEqual(consolidado["E3"].data_type, "n")
            self.assertEqual(revision["A2"].value, 4)
            self.assertEqual(revision["C2"].value, "2026-11-03")
            self.assertEqual(revision["E2"].value, "90.00")
            self.assertIn("RN03", revision["F2"].value)
            self.assertIn("RN06", revision["F2"].value)
            self.assertEqual(duplicados["A2"].value, 2)
            self.assertEqual(duplicados["F2"].value, 1)
            self.assertEqual(dict(resumen.values), {
                "Indicador": "Valor",
                "Archivos procesados": 1,
                "Total procesado": 5,
                "Válidos": 2,
                "Duplicados": 1,
                "A revisar": 2,
                "Fechas convertidas": 3,
                "Suma de importes válidos": -0.1,
            })
            for hoja in libro:
                self.assertEqual(hoja.auto_filter.ref, hoja.dimensions)
                self.assertTrue(hoja["A1"].font.bold)
            libro.close()

    def test_suma_decimal_sin_perder_precision_antes_de_exportar(self):
        importes = [
            Decimal("10000000000000000000000000000"),
            Decimal("0.001"),
            Decimal("-10000000000000000000000000000"),
        ]
        self.assertEqual(sumar_importes(importes), Decimal("0.001"))
