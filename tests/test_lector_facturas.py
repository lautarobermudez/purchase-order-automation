"""Tests de las reglas de validacion de facturas (CUIT, numero de factura, duplicados)."""
import datetime
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
import lector_facturas as lf  # noqa: E402


class CuitValido(unittest.TestCase):
    def test_cuits_con_digito_verificador_correcto(self):
        for cuit in ('30-71234567-1', '20-12345678-6', '30000000007', '27111111117'):
            with self.subTest(cuit=cuit):
                self.assertTrue(lf.cuit_valido(cuit))

    def test_digito_verificador_incorrecto(self):
        self.assertFalse(lf.cuit_valido('30-71234567-2'))
        self.assertFalse(lf.cuit_valido('20-12345678-0'))

    def test_largo_incorrecto_o_vacio(self):
        for cuit in ('', None, '123', '3071234567', '307123456711'):
            with self.subTest(cuit=cuit):
                self.assertFalse(lf.cuit_valido(cuit))

    def test_ignora_guiones_y_espacios(self):
        self.assertTrue(lf.cuit_valido(' 30 - 71234567 - 1 '))


class NumeroDeFactura(unittest.TestCase):
    def test_formatos_habituales(self):
        casos = {
            '0050-06897538': (50, 6897538),
            'A-00002-00584678': (2, 584678),
            'FA-A 00013-00043232': (13, 43232),
            '50-6897538': (50, 6897538),
        }
        for texto, esperado in casos.items():
            with self.subTest(texto=texto):
                self.assertEqual(lf.norm_factura(texto), esperado)

    def test_lecturas_dudosas_se_rechazan(self):
        # punto de venta de mas de 5 digitos o numero de mas de 8: la lectura salio mal
        for texto in ('123456-00000001', '0001-123456789', 'sin numero', '', None):
            with self.subTest(texto=texto):
                self.assertIsNone(lf.norm_factura(texto))

    def test_formato_de_salida(self):
        self.assertEqual(lf.formato_factura((13, 43232)), '0013-00043232')
        self.assertEqual(lf.formato_factura(lf.norm_factura('A-00002-00584678')), '0002-00584678')


class Numeros(unittest.TestCase):
    def test_num(self):
        self.assertEqual(lf.num('12.5'), 12.5)
        self.assertEqual(lf.num(3), 3.0)
        self.assertIsNone(lf.num('abc'))
        self.assertIsNone(lf.num(None))


class FechaDesdeRuta(unittest.TestCase):
    def test_carpeta_del_dia(self):
        ruta = '/root/2026/2026-09-04 - Viernes/OP12.xlsx'
        self.assertEqual(lf.fecha_desde_ruta(ruta), datetime.date(2026, 9, 4))

    def test_sin_fecha_o_fecha_imposible(self):
        self.assertIsNone(lf.fecha_desde_ruta('/root/otra/OP12.xlsx'))
        self.assertIsNone(lf.fecha_desde_ruta('/root/2026-13-45 - X/OP12.xlsx'))


class RegistroDesdeFactura(unittest.TestCase):
    """Lo que devuelve el modelo pasa por la compuerta antes de generar nada."""

    HOY = datetime.date(2026, 10, 20)
    CFG = {}
    CUIT = '30712345671'

    def unidad(self, receptor):
        return 'ACME_SA' if receptor and 'ACME' in receptor.upper() else None

    def lectura(self, **cambios):
        d = {
            'tipo': 'factura_a_pagar', 'proveedor': 'SUPPLIER_X S.A.', 'cuit_proveedor': self.CUIT,
            'receptor': 'ACME SA', 'n_factura': '0001-00000123', 'fecha': '2026-10-10',
            'vencimiento': '2026-10-30', 'importe_total': 1500.0, 'detalle': 'Servicio',
            'cbu': '0' * 22, 'banco': 'Banco X', 'confianza': 'alta', 'dudas': [],
        }
        d.update(cambios)
        return d

    def historial(self, **kw):
        h = {'por_cuit': {}, 'facturas': {}, 'recientes': [], 'codigos': {}}
        h.update(kw)
        return h

    def correr(self, d, hist=None):
        return lf.registro_desde_factura(d, hist or self.historial(), self.CFG, self.HOY, self.unidad)

    def test_factura_correcta_no_tiene_problemas(self):
        r = self.correr(self.lectura())
        self.assertEqual(r['problemas'], [])
        self.assertEqual(r['unidad'], 'ACME_SA')
        self.assertEqual(r['rec']['n_factura'], '0001-00000123')
        self.assertEqual(r['rec']['importe'], 1500.0)

    def test_comprobante_que_no_es_factura_se_omite(self):
        r = self.correr(self.lectura(tipo='comprobante_de_pago'))
        self.assertIsNotNone(r['omitir'])
        self.assertIsNone(r['rec'])

    def test_cuit_invalido_bloquea(self):
        r = self.correr(self.lectura(cuit_proveedor='30712345672'))
        self.assertTrue(any('CUIT' in p for p in r['problemas']))

    def test_confianza_baja_bloquea(self):
        r = self.correr(self.lectura(confianza='baja'))
        self.assertTrue(any('BAJA' in p for p in r['problemas']))

    def test_importe_ausente_o_negativo_bloquea(self):
        for importe in (None, 0, -10):
            with self.subTest(importe=importe):
                r = self.correr(self.lectura(importe_total=importe))
                self.assertTrue(any('importe' in p for p in r['problemas']))

    def test_receptor_desconocido_bloquea(self):
        r = self.correr(self.lectura(receptor='OTRA EMPRESA'))
        self.assertTrue(any('no coincide' in p for p in r['problemas']))

    def test_factura_ya_cargada_es_duplicado(self):
        hist = self.historial(facturas={(self.CUIT, (1, 123)): 'OP45'})
        r = self.correr(self.lectura(), hist)
        self.assertTrue(any('ya tiene OP (OP45)' in p for p in r['problemas']))

    def test_mismo_importe_posterior_a_la_factura_es_posible_duplicado(self):
        # una OP del mismo proveedor, empresa e importe hecha DESPUES de la factura la cubre
        reciente = (self.CUIT, 1500.0, datetime.date(2026, 10, 15), 'OP50', 'ACME SA')
        r = self.correr(self.lectura(n_factura='0001-00000999'), self.historial(recientes=[reciente]))
        self.assertTrue(any('posible duplicado: OP50' in p for p in r['problemas']))

    def test_abono_mensual_no_choca_con_la_op_del_mes_anterior(self):
        # la OP del mes anterior es ANTERIOR a la nueva factura: no es duplicado
        anterior = (self.CUIT, 1500.0, datetime.date(2026, 9, 12), 'OP20', 'ACME SA')
        r = self.correr(self.lectura(n_factura='0001-00000999'), self.historial(recientes=[anterior]))
        self.assertFalse(any('duplicado' in p for p in r['problemas']))

    def test_otro_importe_no_es_duplicado(self):
        reciente = (self.CUIT, 999.0, datetime.date(2026, 10, 15), 'OP50', 'ACME SA')
        r = self.correr(self.lectura(n_factura='0001-00000999'), self.historial(recientes=[reciente]))
        self.assertFalse(any('duplicado' in p for p in r['problemas']))

    def test_sin_datos_bancarios_ni_historial_bloquea(self):
        r = self.correr(self.lectura(cbu=None, alias=None, banco=None))
        self.assertTrue(any('sin datos bancarios' in p for p in r['problemas']))

    def test_toma_datos_de_pago_de_la_ultima_op_del_cuit(self):
        h = {'cbu': '1' * 22, 'alias': None, 'n_cuenta': None, 'banco': 'Banco Y',
             'condicion': 'TRANSFERENCIA', 'fecha': datetime.date(2026, 9, 1)}
        r = self.correr(self.lectura(cbu=None, alias=None, banco=None), self.historial(por_cuit={self.CUIT: h}))
        self.assertEqual(r['problemas'], [])
        self.assertEqual(r['rec']['cbu'], '1' * 22)
        self.assertTrue(r['de_historial'])


if __name__ == '__main__':
    unittest.main()
