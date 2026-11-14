import unittest

import pandas as pd

from main import clasificar_registros, normalizar_fecha, validar_registro


class ValidarRegistroTest(unittest.TestCase):
    def test_registro_valido_con_importe_negativo_y_fecha_bisiesta(self):
        for fecha in ("29/02/2024", "2024-02-29"):
            with self.subTest(fecha=fecha):
                self.assertEqual(validar_registro("104", fecha, "Diego", "-15.00"), [])

    def test_normaliza_y_registra_la_conversion(self):
        self.assertEqual(normalizar_fecha("2024-02-29"), ("29/02/2024", "Convertida"))
        self.assertEqual(normalizar_fecha("29/02/2024"), ("29/02/2024", "Sin conversión"))
        self.assertEqual(normalizar_fecha("0001-01-01"), ("01/01/0001", "Convertida"))

    def test_fecha_invalida_no_se_normaliza(self):
        for fecha in ("", "2026-02-29", "29/02/2026", "2026/11/01"):
            with self.subTest(fecha=fecha):
                self.assertEqual(normalizar_fecha(fecha), (None, "No aplica"))

    def test_informa_todos_los_campos_obligatorios(self):
        self.assertEqual(
            validar_registro("", "", "   ", ""),
            [
                "RN01: id obligatorio",
                "RN02: fecha obligatoria",
                "RN03: cliente obligatorio",
                "RN04: importe obligatorio",
            ],
        )

    def test_formato_de_fecha(self):
        for fecha in ("2026/11/01", "1/11/2026", "01/1/2026", "2026-1-01", " 2026-11-01"):
            with self.subTest(fecha=fecha):
                self.assertEqual(
                    validar_registro("101", fecha, "Ana", "120.00"),
                    ["RN02: la fecha debe tener formato DD/MM/AAAA o AAAA-MM-DD"],
                )

    def test_fecha_inexistente(self):
        for fecha in ("29/02/2026", "2026-02-29", "1900-02-29", "31/04/2026"):
            with self.subTest(fecha=fecha):
                self.assertEqual(
                    validar_registro("101", fecha, "Ana", "120.00"),
                    ["RN02: la fecha no existe en el calendario"],
                )

    def test_importe_no_numerico(self):
        for importe in ("abc", "NaN", "inf"):
            with self.subTest(importe=importe):
                self.assertEqual(
                    validar_registro("101", "01/11/2026", "Ana", importe),
                    ["RN04: el importe debe ser numérico"],
                )


class ClasificarRegistrosTest(unittest.TestCase):
    def clasificar(self, *filas):
        ventas = pd.DataFrame(filas, columns=["id", "fecha", "cliente", "importe"])
        return clasificar_registros(ventas)

    def test_duplicado_exacto_conserva_ambas_filas(self):
        venta = ("101", "01/11/2026", "Ana", "120.00")
        original, repetida = self.clasificar(venta, venta)
        self.assertEqual((original["fila"], original["categoria"]), (1, "Válido"))
        self.assertEqual((repetida["fila"], repetida["categoria"]), (2, "Duplicado exacto"))
        self.assertEqual(repetida["motivos"], ["RN05: datos equivalentes a la fila 1"])

    def test_fechas_e_importes_equivalentes(self):
        original, repetida = self.clasificar(
            ("101", "01/11/2026", "Ana", "-15.00"),
            ("101", "2026-11-01", "Ana", "-15"),
        )
        self.assertEqual(original["categoria"], "Válido")
        self.assertEqual(repetida["categoria"], "Duplicado exacto")
        self.assertEqual(repetida["fecha_original"], "2026-11-01")
        self.assertEqual(repetida["fecha_normalizada"], "01/11/2026")

    def test_conflicto_de_importe_prevalece_sobre_duplicados(self):
        venta = ("101", "01/11/2026", "Ana", "120.00")
        resultados = self.clasificar(
            venta, venta, ("101", "01/11/2026", "Ana", "120.000000000000000001")
        )
        for registro in resultados:
            self.assertEqual(registro["categoria"], "A revisar")
            self.assertEqual(registro["motivos"], ["RN06: mismo ID con datos diferentes"])

    def test_conserva_errores_en_grupos_con_conflicto(self):
        original, incompleta = self.clasificar(
            ("101", "01/11/2026", "Ana", "120"),
            ("101", "2026-02-29", "", "abc"),
        )
        self.assertEqual(original["categoria"], "A revisar")
        self.assertEqual(incompleta["categoria"], "A revisar")
        self.assertEqual(
            [motivo[:4] for motivo in incompleta["motivos"]],
            ["RN02", "RN03", "RN04", "RN06"],
        )

    def test_filas_invalidas_identicas_quedan_a_revisar(self):
        venta = ("101", "01/11/2026", "", "120")
        resultados = self.clasificar(venta, venta)
        for registro in resultados:
            self.assertEqual(registro["categoria"], "A revisar")
            self.assertEqual(registro["motivos"], ["RN03: cliente obligatorio"])
