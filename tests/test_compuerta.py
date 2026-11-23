"""Tests del control previo (compuerta): una OP solo se genera si la OC esta completa."""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
import revisor_mail as rm  # noqa: E402


def oc(**cambios):
    base = {
        'proveedor': 'SUPPLIER_X S.A.', 'importe': 1000.0, 'cbu': '0' * 22, 'alias': None,
        'n_cuenta': None, 'condicion_pago': 'TRANSFERENCIA', 'cuit': '30712345671', 'resp_inscripto': 'SI',
    }
    base.update(cambios)
    return base


class Compuerta(unittest.TestCase):
    def test_oc_completa_pasa(self):
        prob, adv = rm.compuerta(oc(), 'ACME_SA')
        self.assertEqual(prob, [])
        self.assertEqual(adv, [])

    def test_falta_proveedor_o_importe(self):
        prob, _ = rm.compuerta(oc(proveedor=None), 'ACME_SA')
        self.assertTrue(any('RAZON SOCIAL' in p for p in prob))
        for importe in (None, 0):
            prob, _ = rm.compuerta(oc(importe=importe), 'ACME_SA')
            self.assertTrue(any('importe' in p for p in prob))

    def test_sin_datos_bancarios_bloquea_salvo_efectivo(self):
        sin_banco = oc(cbu=None, alias=None, n_cuenta=None)
        prob, _ = rm.compuerta(sin_banco, 'ACME_SA')
        self.assertTrue(any('datos bancarios' in p for p in prob))
        prob, _ = rm.compuerta(dict(sin_banco, condicion_pago='Efectivo'), 'ACME_SA')
        self.assertEqual(prob, [])

    def test_proveedor_exento_de_datos_bancarios(self):
        sin_banco = oc(cbu=None, alias=None, proveedor='Ente Publico')
        prob, _ = rm.compuerta(sin_banco, 'ACME_SA', sin_banco_ok=('ente publico',))
        self.assertEqual(prob, [])

    def test_alias_o_cuenta_alcanzan_como_dato_bancario(self):
        for campo in ('alias', 'n_cuenta'):
            with self.subTest(campo=campo):
                prob, _ = rm.compuerta(oc(cbu=None, **{campo: 'algo'}), 'ACME_SA')
                self.assertEqual(prob, [])

    def test_cbu_con_largo_incorrecto_bloquea(self):
        prob, _ = rm.compuerta(oc(cbu='123'), 'ACME_SA')
        self.assertTrue(any('22' in p for p in prob))

    def test_empresa_pagadora_desconocida_bloquea(self):
        prob, _ = rm.compuerta(oc(), None)
        self.assertTrue(any('empresa pagadora' in p for p in prob))

    def test_sin_cuit_o_cuit_corto_es_solo_advertencia(self):
        prob, adv = rm.compuerta(oc(cuit=None), 'ACME_SA')
        self.assertEqual(prob, [])
        self.assertIn('sin CUIT', adv)
        prob, adv = rm.compuerta(oc(cuit='1234'), 'ACME_SA')
        self.assertEqual(prob, [])
        self.assertTrue(any('CUIT con 4' in a for a in adv))

    def test_sin_resp_inscripto_es_solo_advertencia(self):
        prob, adv = rm.compuerta(oc(resp_inscripto=None), 'ACME_SA')
        self.assertEqual(prob, [])
        self.assertIn('sin RESP INSCRIPTO', adv)


class UnidadDeF1(unittest.TestCase):
    def test_detecta_la_unidad_por_el_texto_de_la_celda(self):
        self.assertEqual(rm.unidad_de_f1('ACME UNIT_B'), 'UNIT_B')
        self.assertEqual(rm.unidad_de_f1('acme unit_c'), 'UNIT_C')

    def test_acme_sa_va_ultimo_porque_las_demas_tambien_llevan_acme(self):
        self.assertEqual(rm.unidad_de_f1('ACME S.A.'), 'ACME_SA')
        self.assertEqual(rm.unidad_de_f1('ACME UNIT_D'), 'UNIT_D')

    def test_texto_desconocido(self):
        self.assertIsNone(rm.unidad_de_f1('OTRA EMPRESA'))
        self.assertIsNone(rm.unidad_de_f1(None))


if __name__ == '__main__':
    unittest.main()
