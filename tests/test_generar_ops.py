"""Tests de utilidades de generar_ops: fechas de Excel, nombres de unidades y numeracion de OPs."""
import datetime
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
import generar_ops as g  # noqa: E402


class FechasDeExcel(unittest.TestCase):
    def test_ida_y_vuelta(self):
        for d in (datetime.date(2026, 1, 1), datetime.date(2026, 10, 20), datetime.date(2024, 2, 29)):
            with self.subTest(d=d):
                self.assertEqual(g.excel_date(g.excel_serial(d)), d)

    def test_serial_conocido(self):
        self.assertEqual(g.excel_serial(datetime.date(2000, 1, 1)), 36526)
        self.assertEqual(g.excel_date('36526.0'), datetime.date(2000, 1, 1))


class Referencias(unittest.TestCase):
    def test_split_ref(self):
        self.assertEqual(g.split_ref('B12'), ('B', 12))
        self.assertEqual(g.split_ref('AA3'), ('AA', 3))


class Unidades(unittest.TestCase):
    def test_normalizar(self):
        self.assertEqual(g.normalizar_unidad('acme_sa'), 'ACME SA')
        self.assertEqual(g.normalizar_unidad('  unit_b '), 'UNIT B')

    def test_nombre_de_carpeta(self):
        self.assertEqual(g.unidad_folder_name('UNIT_B'), 'Unit B')
        self.assertEqual(g.unidad_folder_name('acme_sa'), 'Acme SA')

    def test_unidad_sin_carpeta_definida_se_titula(self):
        self.assertEqual(g.unidad_folder_name('nueva_empresa'), 'Nueva Empresa')


class Numeracion(unittest.TestCase):
    """El numero de OP sale de las carpetas y archivos YA existentes de esa unidad."""

    def armar(self, raiz, carpeta, nombres):
        base = os.path.join(raiz, '2026', carpeta)
        os.makedirs(base, exist_ok=True)
        for n in nombres:
            open(os.path.join(base, n), 'w').close()

    def test_primera_op_de_una_unidad_vacia(self):
        with tempfile.TemporaryDirectory() as raiz:
            self.assertEqual(g.next_op_number(raiz, 'UNIT_B'), 1)

    def test_sigue_el_maximo_usado(self):
        with tempfile.TemporaryDirectory() as raiz:
            self.armar(raiz, 'Unit B', ['OP12 - Proveedor.xlsx', 'OP40_algo.xlsx', 'OP 7.xlsx'])
            self.assertEqual(g.next_op_number(raiz, 'UNIT_B'), 41)

    def test_un_archivo_de_otra_unidad_no_infla_la_cuenta(self):
        # caso real: un archivo con el nombre de OTRA unidad pegado vivia en la carpeta Unit B
        with tempfile.TemporaryDirectory() as raiz:
            self.armar(raiz, 'Unit B', ['OP415-Proveedor_ACME_SA.xlsx'])
            self.armar(raiz, 'Acme SA', ['OP277.xlsx'])
            self.assertEqual(g.next_op_number(raiz, 'ACME_SA'), 278)

    def test_las_carpetas_con_otro_nombre_no_cuentan(self):
        with tempfile.TemporaryDirectory() as raiz:
            self.armar(raiz, 'Varios', ['OP999.xlsx'])
            self.assertEqual(g.numeros_usados(raiz, 'UNIT_B'), set())


class Comprobantes(unittest.TestCase):
    def test_detecta_comprobante_en_la_carpeta(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertFalse(g.folder_has_comprobante(d))
            open(os.path.join(d, 'Comprobante OP47 - ML.jpg'), 'w').close()
            self.assertTrue(g.folder_has_comprobante(d))

    def test_carpeta_inexistente(self):
        self.assertFalse(g.folder_has_comprobante('/no/existe'))


if __name__ == '__main__':
    unittest.main()
