#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ops.py — genera Ordenes de Pago (OP) de Acme a partir de Ordenes de
Compra (OC), usando siempre el mismo template ("MODELO OP").

No requiere paqutes externos (solo stdlib): zipfile + xml.etree.

MODO AUTOMATICO (OC en formato xlsx, con etiquetas RAZON SOCIAL/CUIT/etc.
en columna A y tabla CANTIDAD/DESCRIPCION/PRECIO UN/VALOR TOTAL):

    python3 generar_ops.py --dest "<carpeta destino>" --unidad UNIT_B \
        --auto "OC proveedor1.xlsx" "OC proveedor2.xlsx"

MODO MANUAL (para OC en PDF u otro formato que Claude tuvo que leer a mano):
    Armar un JSON con una lista de registros, por ejemplo:

    [
      {"proveedor": "SUPPLIER_X", "cuit": "30-00000000-0", "resp_inscripto": "SI",
       "plazo": "INMEDIATO", "detalle": "SUPPLIER_X - OC N°1",
       "fecha_item": "2026-09-01", "importe": 12600,
       "fuente": "OC SUPPLIER_X 01-09-26.pdf", "ya_pagada": false,
       "n_factura": "0001-00012345"}
    ]

    "ya_pagada" es opcional (default false): poner true si la OC en PDF dice
    explicitamente "ya paga" / "ya pagada" o similar.

    "n_factura" es opcional: poner el numero de factura tal como figura en el
    comprobante (ej. "0001-00012345") cuando la OC/registro viene con una
    factura real asociada. Si se omite o queda vacio, la columna de N°
    FACTURA de la OP queda en "Sin factura" (para OC que todavia no tienen
    factura, el caso mas comun). En modo --auto el script intenta
    autoextraer este dato de una etiqueta "N° FACTURA" en la OC en xlsx si
    existe; si no la encuentra, tambien cae en "Sin factura".

    python3 generar_ops.py --dest "<carpeta destino>" --unidad UNIT_B \
        --manual registros.json

Por defecto el script corre en DRY-RUN (solo muestra qué haría). Agregar
--apply para escribir los archivos y mover las OC de origen a su carpeta
OP### final.

ENVIO POR MAIL (opcional, solo con --apply): agregar --email para mandar las
OP recien generadas (con su OC de origen y comprobantes). Siempre va a la
lista fija DEFAULT_RECIPIENTS (the finance team's addresses, set in the code); lo que pases en --email se suma a esa
lista, no la reemplaza:

    python3 generar_ops.py --dest "<carpeta destino>" --unidad UNIT_B \
        --auto "OC1.xlsx" --apply --email
    # o sumando una direccion extra:
    python3 generar_ops.py --dest "<carpeta destino>" --unidad UNIT_B \
        --auto "OC1.xlsx" --apply --email other@example.com

El cuerpo del mail es siempre:
    Estimados,

    Adjunto unas ordenes para los proximos pagos de <razon social>.

    Saludos
donde <razon social> se deriva de --unidad (ej. UNIT_B -> "Acme
Unit B") o se puede fijar a mano con --razon-social.

El asunto por defecto es "Ordenes de pago - <razon social> (OPs <numeros>)",
ej: "Ordenes de pago - Acme Unit B (OPs 422 - 423 - 424)"; se puede
fijar a mano con --email-subject. Si alguna de las OC que se estan mandando
tiene "ya_pagada": true (una OC en xlsx que dice "ya paga"/"ya pago" en algun lado se
detecta sola; en manual hay que marcarla a mano) o si alguna carpeta OP###
ya tiene un archivo "Comprobante..." adentro, se le agrega " - YA PAGA" al
final del asunto.

IMPRESION (opcional, solo con --apply): --imprimir manda a la impresora
predeterminada de Windows el xlsx de cada OP recien generada Y, si la carpeta
OP### tiene la factura (PDF o imagen jpg/png), tambien la factura: siempre OP +
factura, una copia de cada una. No imprime la OC de origen ni los comprobantes
de pago. Usa Excel de Windows (OP) y _Scripts/imprimir_archivo.ps1 (factura) via
powershell.exe, asi que solo anda corriendo desde WSL en una maquina con Excel.

Requiere las variables de entorno SMTP_USER y SMTP_PASS (la cuenta de
MailHost que envia). Opcionalmente SMTP_HOST / SMTP_PORT (default
smtp.example.com:465).
"""
import argparse
import datetime
import json
import mimetypes
import os
import re
import shutil
import smtplib
import time
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from email.message import EmailMessage


# La compu esta en hora argentina pero WSL corre en UTC: sin esto, entre las 21:00 y las 24:00
# date.today() devuelve el dia siguiente (OP con fecha y carpeta de manana).
os.environ.setdefault('TZ', 'America/Argentina/Buenos_Aires')
time.tzset()


def _load_dotenv(path=None):
    """Carga variables desde un .env simple (KEY=VALUE por linea, junto a
    este script) sin pisar las que ya esten seteadas en el entorno. No
    requiere python-dotenv (no hay acceso a pip en este entorno)."""
    path = path or os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    if not os.path.isfile(path):
        return
    with open(path, encoding='utf-8') as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            key, _, value = line.partition('=')
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


_load_dotenv()

NS_URI = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS_URI)


def excel_serial(d: datetime.date) -> int:
    return (d - datetime.date(1899, 12, 30)).days


def excel_date(serial) -> datetime.date:
    return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(serial)))


def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)', ref)
    return m.group(1), int(m.group(2))


# ---------------------------------------------------------------------------
# Lectura genérica de una OC en xlsx
# ---------------------------------------------------------------------------

def read_grid(xlsx_path):
    z = zipfile.ZipFile(xlsx_path)
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall('{%s}si' % NS_URI):
            text = ''.join(t.text or '' for t in si.iter('{%s}t' % NS_URI))
            shared.append(text)
    root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    grid = {}
    for c in root.iter('{%s}c' % NS_URI):
        ref = c.get('r')
        t = c.get('t')
        v_el = c.find('{%s}v' % NS_URI)
        if v_el is None or v_el.text is None:
            continue
        v = v_el.text
        if t == 's':
            v = shared[int(v)]
        grid[ref] = v
    return grid


def label_value(grid, label):
    label_norm = label.strip().upper()
    for ref, val in grid.items():
        col, row = split_ref(ref)
        if col not in ('A',):
            continue
        if isinstance(val, str) and val.strip().upper() == label_norm:
            for cand_col in ('B', 'C', 'D'):
                cand_ref = f'{cand_col}{row}'
                if cand_ref in grid and grid[cand_ref] not in (None, ''):
                    return grid[cand_ref]
            return None
    return None


def extract_from_oc_xlsx(path):
    grid = read_grid(path)
    proveedor = label_value(grid, 'RAZON SOCIAL') or label_value(grid, 'RAZON SOCIAL ')
    email = label_value(grid, 'EMAIL')
    cuit = label_value(grid, 'CUIT')
    resp = label_value(grid, 'RESP INSCRIPTO')
    cond = label_value(grid, 'CONDICION DE PAGO')
    plazo = label_value(grid, 'PLAZO')
    oc_num = label_value(grid, 'N° OC') or label_value(grid, 'N OC')
    n_factura = (label_value(grid, 'N° FACTURA') or label_value(grid, 'N FACTURA')
                 or label_value(grid, 'NRO FACTURA') or label_value(grid, 'FACTURA'))
    banco = label_value(grid, 'BANCO')
    cbu = label_value(grid, 'CBU')
    alias = label_value(grid, 'ALIAS')
    if not alias:
        # a veces el alias viene suelto en otra celda, ej. E11 = "ALIAS: akimpress.3df"
        alias = next((v for v in grid.values()
                      if isinstance(v, str) and re.match(r'(?i)^\s*alias\s*:', v)), None)
    n_cuenta = (label_value(grid, 'N CUENTA') or label_value(grid, 'N° CUENTA')
                or label_value(grid, 'NRO CUENTA'))
    fecha_val = label_value(grid, 'FECHA')
    fecha = excel_date(fecha_val) if fecha_val else datetime.date.today()

    header_row = None
    for ref, val in grid.items():
        col, row = split_ref(ref)
        if col == 'A' and isinstance(val, str) and val.strip().upper() == 'CANTIDAD':
            header_row = row
            break

    items = []
    if header_row:
        r = header_row + 1
        while True:
            a_ref = f'A{r}'
            sin_cantidad = a_ref not in grid or grid[a_ref] in (None, '')
            # la tabla termina cuando no hay ni cantidad ni descripcion; hay OC
            # que traen la descripcion sin completar CANTIDAD (ej. Insurance Co.)
            if sin_cantidad and not grid.get(f'B{r}'):
                break
            cantidad = None if sin_cantidad else grid[a_ref]
            desc = grid.get(f'B{r}', '')
            importe = grid.get(f'H{r}')
            items.append((cantidad, desc, importe))
            r += 1

    total = 0.0
    descs = []
    for cantidad, desc, importe in items:
        try:
            total += float(importe)
        except (TypeError, ValueError):
            pass
        if desc:
            descs.append(f'{desc} x{cantidad}' if cantidad is not None else str(desc).strip())

    # Si la OC trae su propia fila TOTAL (ej. G26 "TOTAL" / H26), manda esa:
    # las filas de descuento vienen cargadas en positivo y sumarlas infla el
    # importe (paso con OC SUPPLIER_Z 30-9: "DESCUENTO DE 10% X PRONTO PAGO").
    for ref, val in grid.items():
        col, row = split_ref(ref)
        if isinstance(val, str) and val.strip().upper() == 'TOTAL' and col != 'H':
            try:
                oc_total = float(grid.get(f'H{row}'))
            except (TypeError, ValueError):
                continue
            if oc_total:
                total = oc_total
            break

    # Desglose para la OP: una linea por item de la OC con importe. Los
    # descuentos/bonificaciones vienen cargados en positivo en la OC: van en
    # negativo para que la suma de la OP de el total de la OC.
    desglose = []
    for cantidad, desc, importe in items:
        try:
            imp = float(importe)
        except (TypeError, ValueError):
            continue
        if not imp:
            continue
        if imp > 0 and re.search(r'DESCUENTO|BONIF', str(desc), re.I):
            imp = -imp
        desglose.append({'detalle': str(desc).strip(), 'importe': round(imp, 2)})
    if not (1 < len(desglose) <= MAX_ITEMS) or abs(sum(i['importe'] for i in desglose) - total) > 0.01:
        desglose = None

    detalle = '; '.join(descs) if descs else (proveedor or '')
    if oc_num:
        detalle = f'{detalle} (OC N°{oc_num})' if detalle else f'OC N°{oc_num}'

    ya_pagada_re = re.compile(r'\bya\b.{0,20}\bpag[oaóá]', re.I)
    ya_pagada = any(isinstance(v, str) and ya_pagada_re.search(v) for v in grid.values())

    return dict(
        proveedor=(proveedor or '').strip() or None,
        cuit=re.sub(r'\D', '', cuit) if cuit else None,
        resp_inscripto=(resp or '').strip().upper() or None,
        condicion_pago=(cond or '').strip() or None,
        plazo=(plazo or '').strip() or None,
        email=(email or '').strip() or None,
        detalle=detalle,
        fecha_item=fecha.isoformat(),
        importe=round(total, 2),
        items=desglose,
        fuente=os.path.basename(path),
        ya_pagada=ya_pagada,
        n_factura=(n_factura or '').strip() or None,
        banco=(banco or '').strip() or None,
        # la OC a veces trae el prefijo en la celda ("CBU: 0170...")
        cbu=re.sub(r'\D', '', cbu) or None if cbu else None,
        alias=re.sub(r'(?i)^\s*alias\s*:?\s*', '', alias).strip() or None if alias else None,
        n_cuenta=(n_cuenta or '').strip() or None,
    )


# ---------------------------------------------------------------------------
# Escritura de la OP a partir del MODELO_OP.xlsx
# ---------------------------------------------------------------------------

def esc_xml(s):
    return (str(s).replace('&', '&amp;').replace('<', '&lt;')
            .replace('>', '&gt;').replace('"', '&quot;'))


def set_cell(xml, ref, style, *, text_idx=None, number=None, formula=None):
    pattern = re.compile(r'<c r="' + re.escape(ref) + r'"(?:\s[^>]*?)?(?:/>|>.*?</c>)', re.S)
    if not pattern.search(xml):
        xml = insert_cell(xml, ref)
    if text_idx is not None:
        new = f'<c r="{ref}" s="{style}" t="s"><v>{text_idx}</v></c>'
    elif formula is not None:
        f_text, cached = formula
        new = f'<c r="{ref}" s="{style}"><f>{f_text}</f><v>{cached}</v></c>'
    elif number is not None:
        new = f'<c r="{ref}" s="{style}"><v>{number}</v></c>'
    else:
        new = f'<c r="{ref}" s="{style}"/>'
    # lambda en vez de pasar `new` directo: re.sub interpreta \1, \g<...> etc.
    # en un string de reemplazo, y `new` puede traer datos arbitrarios del
    # registro (proveedor, detalle) que no deberian tratarse como backrefs.
    return pattern.sub(lambda _m: new, xml, count=1)


def insert_cell(xml, ref):
    """Agrega una celda vacia <c r="ref"/> en su fila, en orden de columna
    (el template no trae todas las celdas del detalle, ej. falta A17)."""
    col, row = split_ref(ref)
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % row, xml, re.S)
    if not m:
        raise ValueError(f'No encontré la fila {row} en el template (¿cambió el MODELO OP?)')
    key = lambda c: (len(c), c)
    pos = m.end(1)
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"', m.group(1)):
        if key(cm.group(1)) > key(col):
            pos = m.start(1) + cm.start()
            break
    return xml[:pos] + f'<c r="{ref}"/>' + xml[pos:]


def set_cell_or_blank(sheet, ref, style, text_idx):
    """set_cell con texto si text_idx no es None, celda vacia si lo es."""
    if text_idx is not None:
        return set_cell(sheet, ref, style, text_idx=text_idx)
    return set_cell(sheet, ref, style)


MAX_ITEMS = 6  # filas de detalle del MODELO OP: 15 a 20
DETALLE_B_STYLE = [43, 31, 34, 34, 34, 37]
DETALLE_H_STYLE = [16, 17, 19, 19, 19, 24]


def lineas_detalle(record):
    """Lineas de detalle de la OP. Con "items" (desglose de factura + items
    de la OC que no estan en la factura, ej. descuento por pronto pago) va
    una linea por item; si no, una sola linea con "detalle" e "importe"."""
    items = record.get('items')
    if items:
        if len(items) > MAX_ITEMS:
            raise ValueError(f'{record.get("proveedor")}: {len(items)} lineas de detalle, '
                             f'el MODELO OP tiene lugar para {MAX_ITEMS}. Agrupar conceptos.')
        return [dict(it, importe=round(float(it.get('importe') or 0), 2)) for it in items]
    return [{'detalle': record.get('detalle') or record.get('proveedor') or '',
             'importe': round(float(record.get('importe') or 0), 2)}]


FIXED_STRINGS = [
    "N° OP", "RAZON SOCIAL", "FECHA", "RAZON SOCIAL ", "EMAIL", "CUIT", "RESP INSCRIPTO",
    "CONDICION DE PAGO", "PLAZO", "BANCO", "CBU", "N Cuenta", "N° FACTURA", "DETALLE", "IMPORTE",
    "SUB TOTAL", "RETENCIONES", "IMPUESTOS", "TOTAL", "RESPONSABLE DE PAGO", "ADMIN Y FIN",
    "AUTORIZA", "JANE DOE", "ACME UNIT_B",
]  # indices 0-23, fijos: son las etiquetas y firmantes del MODELO OP


def build_op(template_path, out_path, op_num, record, unidad, fecha_op=None):
    fecha_op = fecha_op or datetime.date.today()

    with open(template_path, 'rb') as f:
        tz = zipfile.ZipFile(f)
        sheet = tz.read('xl/worksheets/sheet1.xml').decode('utf-8')
        others = {n: tz.read(n) for n in tz.namelist()
                  if n not in ('xl/worksheets/sheet1.xml', 'xl/sharedStrings.xml', 'xl/calcChain.xml')}

    extra = []

    def idx_for(text):
        extra.append(text)
        return len(FIXED_STRINGS) + len(extra) - 1

    i_proveedor = idx_for(record.get('proveedor') or 'SIN ESPECIFICAR')
    n_factura = (record.get('n_factura') or '').strip()
    i_factura = idx_for(n_factura) if n_factura else idx_for('Sin factura')
    lineas = lineas_detalle(record)
    i_detalles = [idx_for(l['detalle'] or '') for l in lineas]
    i_facturas_items = [idx_for(l['n_factura']) if l.get('n_factura') else None for l in lineas]
    i_plazo = idx_for(record['plazo']) if record.get('plazo') else None
    i_resp = idx_for(record['resp_inscripto']) if record.get('resp_inscripto') else None
    i_cond = idx_for(record['condicion_pago']) if record.get('condicion_pago') else None
    i_email = idx_for(record['email']) if record.get('email') else None
    i_razon = idx_for(razon_social_f1(unidad))
    i_banco = idx_for(record['banco']) if record.get('banco') else None
    # CBU como texto: 22 digitos no entran en un numero de Excel sin perder precision
    i_cbu = idx_for(str(record['cbu'])) if record.get('cbu') else None
    # El MODELO OP no tiene fila de ALIAS: va en "N Cuenta" junto al numero de cuenta
    cuenta = ' / '.join(x for x in (
        record.get('n_cuenta'),
        f"ALIAS: {record['alias']}" if record.get('alias') else None) if x)
    i_cuenta = idx_for(cuenta) if cuenta else None

    strings = FIXED_STRINGS + extra
    si = ''.join(f'<si><t>{esc_xml(s)}</t></si>' for s in strings)
    ss_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              f'<sst xmlns="{NS_URI}" count="{len(strings)}" uniqueCount="{len(strings)}">{si}</sst>')

    importe = round(sum(l['importe'] for l in lineas), 2)
    fecha_item = record.get('fecha_item')
    fecha_item_d = datetime.date.fromisoformat(fecha_item) if fecha_item else fecha_op

    sheet = set_cell(sheet, 'F1', 28, text_idx=i_razon)
    if op_num is None:  # OP sin numero (se asigna el dia de pago)
        sheet = set_cell(sheet, 'B1', 28)
    else:
        sheet = set_cell(sheet, 'B1', 28, number=op_num)
    sheet = set_cell(sheet, 'B2', 52, number=excel_serial(fecha_op))
    sheet = set_cell(sheet, 'B4', 28, text_idx=i_proveedor)
    sheet = set_cell_or_blank(sheet, 'B5', 46, i_email)
    if record.get('cuit'):
        sheet = set_cell(sheet, 'B6', 28, number=record['cuit'])
    else:
        sheet = set_cell(sheet, 'B6', 28)
    sheet = set_cell_or_blank(sheet, 'B7', 28, i_resp)
    sheet = set_cell_or_blank(sheet, 'B8', 28, i_cond)
    sheet = set_cell_or_blank(sheet, 'B9', 55, i_plazo)
    sheet = set_cell_or_blank(sheet, 'B10', 28, i_banco)
    sheet = set_cell_or_blank(sheet, 'B11', 40, i_cbu)
    sheet = set_cell_or_blank(sheet, 'B12', 28, i_cuenta)
    sheet = set_cell(sheet, 'A15', 14, text_idx=i_factura)
    sheet = set_cell(sheet, 'G15', 15, number=excel_serial(fecha_item_d))
    # Lineas de detalle en filas 15-20, conservando el estilo de cada fila del
    # template (bordes); la columna H siempre con formato moneda (H17 del
    # template tiene formato fecha). N° factura y fecha solo en la 1ra linea,
    # salvo que una linea traiga los suyos.
    for n, (linea, i_det, i_fac) in enumerate(zip(lineas, i_detalles, i_facturas_items)):
        row = 15 + n
        sheet = set_cell(sheet, f'B{row}', DETALLE_B_STYLE[n], text_idx=i_det)
        sheet = set_cell(sheet, f'H{row}', DETALLE_H_STYLE[n], number=linea['importe'])
        if n and i_fac is not None:
            sheet = set_cell(sheet, f'A{row}', 14, text_idx=i_fac)
        if n and linea.get('fecha'):
            sheet = set_cell(sheet, f'G{row}', 15,
                             number=excel_serial(datetime.date.fromisoformat(linea['fecha'])))
    sheet = set_cell(sheet, 'H21', 26, formula=('SUM(H15:H20)', importe))
    sheet = set_cell(sheet, 'H24', 26, formula=('+H21-H22-H23', importe))

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, content in others.items():
            zf.writestr(name, content)
        zf.writestr('xl/sharedStrings.xml', ss_xml)
        zf.writestr('xl/worksheets/sheet1.xml', sheet)


# ---------------------------------------------------------------------------
# Unidad de negocio: nombre de carpeta y razon social
# ---------------------------------------------------------------------------

# Nombre de subcarpeta de razon social por unidad, tal como se viene usando
# en el registro (ej. --unidad UNIT_B -> carpeta "Unit B").
UNIDAD_CARPETA = {
    'UNIT_B': 'Unit B', 'ACME UNIT_B': 'Unit B',
    'ACME SA': 'Acme SA', 'ACME': 'Acme SA',
    'UNIT_C': 'Unit C', 'UNIT_D': 'Unit D', 'SUPPLIER_Y': 'Supplier Y',
    'BUY': 'Buy', 'CLEANER': 'Cleaner', 'CITY': 'City',
}


def normalizar_unidad(unidad):
    """--unidad en mayusculas y con espacios en vez de guiones bajos, ej.
    'acme_sa' -> 'ACME SA'. Base comun de todos los nombres derivados."""
    return ' '.join(unidad.replace('_', ' ').upper().split())


def _titulo(texto):
    """'ACME SA' -> 'Acme SA': palabras de hasta 3 letras quedan en
    mayusculas (siglas como SA), el resto con inicial mayuscula."""
    return ' '.join(w.upper() if len(w) <= 3 else w.capitalize() for w in texto.split())


def unidad_folder_name(unidad):
    key = normalizar_unidad(unidad)
    return UNIDAD_CARPETA.get(key) or _titulo(key)


def razon_social_f1(unidad):
    """Texto que va impreso en el encabezado (celda F1) de la OP, ej.
    UNIT_B -> 'ACME UNIT_B', UNIT_C -> 'ACME UNIT_C',
    ACME_SA -> 'ACME SA' (sin duplicar el prefijo ACME)."""
    key = normalizar_unidad(unidad)
    return key if key.split()[:1] == ['ACME'] else f'ACME {key}'


def display_razon_social(unidad):
    """Razon social para el cuerpo y asunto del mail: 'Acme' + nombre de
    carpeta de la unidad (ej. UNIT_B -> 'Acme Unit B', ACME_SA ->
    'Acme SA', UNIT_C -> 'Acme Unit C', BUY -> 'Acme Buy')."""
    carpeta = unidad_folder_name(unidad)
    return carpeta if carpeta.upper().split()[:1] == ['ACME'] else f'Acme {carpeta}'


# ---------------------------------------------------------------------------
# Autodeteccion de carpeta destino (fecha de HOY)
# ---------------------------------------------------------------------------

MESES_ES = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO',
            'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
DIAS_ES = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado', 'Domingo']


def mes_folder_name(fecha):
    return f'{fecha.month} - {MESES_ES[fecha.month - 1]} {fecha.year}'


def dia_folder_name(fecha):
    return f'{fecha.day} - {fecha.month:02d} {DIAS_ES[fecha.weekday()]}'


def auto_dest(root, unidad, fecha=None):
    """Carpeta del dia de HOY + razon social. Se recalcula siempre a partir
    de la fecha real del sistema (nunca reusa la carpeta del dia anterior
    solo porque sea la mas reciente en disco)."""
    fecha = fecha or datetime.date.today()
    return os.path.join(root, str(fecha.year), mes_folder_name(fecha),
                         dia_folder_name(fecha), unidad_folder_name(unidad))


# ---------------------------------------------------------------------------
# Numeracion automatica de OP
# ---------------------------------------------------------------------------

def detect_root(dest):
    parts = os.path.normpath(dest).split(os.sep)
    for i, p in enumerate(parts):
        if p.strip().lower() == 'ordenes de pago':
            return os.sep.join(parts[:i + 1])
    raise SystemExit('No pude detectar la carpeta raiz "Ordenes de Pago" a partir de --dest; '
                      'pasa --root explicitamente.')


def numeros_usados(root, unidad):
    """Numeros de OP ya usados por la unidad (ver next_op_number)."""
    canonical = unidad_folder_name(unidad).lower()
    valid_names = {canonical, canonical.split()[0]}
    usados = set()
    op_re = re.compile(r'OP\s?(\d+)', re.I)
    for dirpath, dirnames, filenames in os.walk(root):
        if os.path.basename(dirpath).lower() not in valid_names:
            continue
        for name in dirnames + filenames:
            for m in op_re.finditer(name):
                usados.add(int(m.group(1)))
    return usados


def next_op_number(root, unidad):
    """Busca el maximo numero de OP ya usado para esta unidad de negocio.

    Solo cuenta carpetas y archivos que viven DENTRO de una carpeta cuyo
    nombre exacto es el de la unidad (ej. "Unit B", "Acme SA", o
    "Acme" como variante corta historica) — nunca por substring en
    cualquier parte de la ruta o el nombre de archivo. Un substring como se
    hacia antes se dejaba enganar por un archivo con el nombre de OTRA
    unidad de negocio pegado en el nombre (paso de verdad: un archivo
    "OP415-CondPerezMoralez_ACME_SA.xlsx" viviendo en la carpeta Unit B
    inflaba el conteo de Acme SA a 416 en vez de 278)."""
    return max(numeros_usados(root, unidad), default=0) + 1


# ---------------------------------------------------------------------------
# Envio de las OP generadas por mail
# ---------------------------------------------------------------------------

SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.example.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '465'))

# Estos destinatarios van siempre que se manda el mail de OPs; --email solo
# permite sumar direcciones extra, nunca reemplaza esta lista.
DEFAULT_RECIPIENTS = [
    'user@example.com',
    'user@example.com',
    'user@example.com',
    'user@example.com',
    'user@example.com',
    'user@example.com',
]


def folder_has_comprobante(folder):
    """True si la carpeta de la OP ya tiene un comprobante de pago adentro
    (archivo cuyo nombre contiene 'comprobante', como ya se nombran a mano
    en OPs anteriores, ej. 'Comprobante OP47 - ML - Unit B.jpg')."""
    if not os.path.isdir(folder):
        return False
    return any('comprobante' in fname.lower() for fname in os.listdir(folder))


def send_ops_email(recipients, subject, op_folders, razon_social):
    smtp_user = os.environ.get('SMTP_USER')
    smtp_pass = os.environ.get('SMTP_PASS')
    if not smtp_user or not smtp_pass:
        print('AVISO: no mande el mail porque faltan las variables de entorno '
              'SMTP_USER / SMTP_PASS (y opcionalmente SMTP_HOST / SMTP_PORT, '
              'default smtp.example.com:465).')
        return

    msg = EmailMessage()
    msg['From'] = smtp_user
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = subject

    body = (
        'Estimados,\n\n'
        f'Adjunto unas ordenes para los proximos pagos de {razon_social}.\n\n'
        'Saludos'
    )
    msg.set_content(body)

    file_list = []
    for op_name, folder in op_folders:
        if not os.path.isdir(folder):
            continue
        for fname in sorted(os.listdir(folder)):
            file_list.append((op_name, fname, os.path.join(folder, fname)))

    for op_name, fname, fpath in file_list:
        ctype, _ = mimetypes.guess_type(fpath)
        if ctype is None:
            ctype = 'application/octet-stream'
        maintype, subtype = ctype.split('/', 1)
        with open(fpath, 'rb') as f:
            data = f.read()
        msg.add_attachment(data, maintype=maintype, subtype=subtype,
                            filename=f'{op_name}_{fname}')

    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
    print(f'Mail enviado a {", ".join(recipients)} con {len(file_list)} adjuntos.')


def es_factura_imprimible(fname):
    """True si el archivo de una carpeta OP### es una factura para imprimir: PDF o imagen
    (jpg/jpeg/png) que NO sea la OP, la OC de origen ("OC ...") ni un comprobante de pago."""
    low = fname.lower()
    if not low.endswith(('.pdf', '.jpg', '.jpeg', '.png')):
        return False
    return not (low.startswith(('op', 'oc ')) or 'comprobante' in low)


def imprimir_ops(op_folders, solo_probar=False):
    """REGLA (pedida por el usuario, 2026-10-05): siempre se imprime la OP y, si la carpeta
    OP### tiene la factura (PDF o imagen), tambien la factura. 1 copia de cada una, en la
    impresora predeterminada de Windows. No se imprime la OC ni los comprobantes de pago.
    La OP (xlsx) se imprime con Excel via powershell.exe; la factura con imprimir_archivo.ps1
    (junto a este script), porque el visor de PDF por defecto (Edge) no tiene verbo "print"."""
    ops, facturas = [], []
    for op_name, folder in op_folders:
        if not os.path.isdir(folder):
            continue
        for fname in sorted(os.listdir(folder)):
            path = os.path.join(folder, fname)
            if fname.upper().startswith('OP') and fname.lower().endswith('.xlsx'):
                ops.append(path)
            elif es_factura_imprimible(fname):
                facturas.append(path)
        if not any(o.startswith(folder) for o in ops):
            print(f'AVISO: {op_name}: no encontre el xlsx de la OP para imprimir.')
        if not any(f.startswith(folder) for f in facturas):
            print(f'AVISO: {op_name}: la carpeta no tiene factura (PDF/imagen), se imprime solo la OP.')
    if not ops and not facturas:
        print('AVISO: --imprimir no encontro nada para imprimir.')
        return False

    def wpath(path):
        return subprocess.check_output(['wslpath', '-w', path], text=True).strip()

    def ps_run(args, timeout=180, ok_marker=None):
        r = subprocess.run(['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass'] + args,
                           capture_output=True, text=True, errors='replace', timeout=timeout)   # PowerShell responde en cp850
        # imprimir_archivo.ps1 cierra con codigo 5 (cierre abrupto de PowerShell al salir) aunque haya
        # impreso bien: ahi el exito se decide por la linea "OK:" y no por el codigo de salida.
        if ok_marker and ok_marker in r.stdout:
            return
        if r.returncode != 0 or ok_marker:
            raise RuntimeError((r.stderr or r.stdout).strip())

    try:
        if ops:
            lista = ','.join("'" + wpath(o).replace("'", "''") + "'" for o in ops)
            imprimir_cmd = '' if solo_probar else '$wb.ActiveSheet.PrintOut(); '   # solo_probar: abre y cierra sin imprimir
            ps_run(['-Command',
                    "$x=New-Object -ComObject Excel.Application; $x.Visible=$false; $x.DisplayAlerts=$false; "
                    "try { foreach($f in @(" + lista + ")) { $wb=$x.Workbooks.Open($f,0,$true); "
                    + imprimir_cmd + "$wb.Close($false) } } finally { $x.Quit() }"])
            print(f'{"Probado sin imprimir" if solo_probar else "Mandado a imprimir"} ({len(ops)} OP): ' + ', '.join(os.path.basename(o) for o in ops))
        script = wpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'imprimir_archivo.ps1'))
        for f in facturas:
            ps_run(['-File', script, '-Path', wpath(f)] + (['-SinImprimir'] if solo_probar else []), ok_marker='OK:')
            print(f'Mandada a imprimir (factura): {os.path.basename(f)}')
    except Exception as e:   # cualquier falla de impresion se informa; nunca debe cortar al revisor
        print(f'AVISO: la impresion fallo: {e}. --imprimir solo anda desde WSL con Excel y una impresora predeterminada.')
        return False
    return True


def email_de_ops(records, generated_folders, razon_social, extra_recipients, subject=None):
    """Arma destinatarios y asunto del mail de OPs. Devuelve (recipients, subject).
    Los DEFAULT_RECIPIENTS van siempre; extra_recipients solo se suman."""
    extra = [e for e in extra_recipients if e not in DEFAULT_RECIPIENTS]
    recipients = DEFAULT_RECIPIENTS + extra
    if subject:
        return recipients, subject
    op_numbers = [name.replace('OP', '') for name, _ in generated_folders]
    subject = f'Ordenes de pago - {razon_social} (OPs {" - ".join(op_numbers)})'
    ya_pagada = (any(rec.get('ya_pagada') for rec, _ in records)
                 or any(folder_has_comprobante(folder) for _, folder in generated_folders))
    if ya_pagada:
        subject += ' - YA PAGA'
    return recipients, subject


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def avisos_registro(rec, src):
    """Imprime avisos de datos faltantes o sospechosos de una OC."""
    if not rec.get('proveedor'):
        print(f'AVISO: {src} no tiene RAZON SOCIAL detectada, revisar a mano.')
    if not rec.get('importe'):
        print(f'AVISO: {src} no tiene IMPORTE (o es 0), revisar a mano.')
    en_efectivo = 'efectivo' in (rec.get('condicion_pago') or '').lower()
    if not en_efectivo and not (rec.get('cbu') or rec.get('alias') or rec.get('n_cuenta')):
        print(f'AVISO: {src} no tiene DATOS BANCARIOS (CBU / alias / cuenta), '
              'sin eso no se puede pagar por transferencia. Revisar a mano.')
    cbu_digits = re.sub(r'\D', '', str(rec.get('cbu') or ''))
    if rec.get('cbu') and len(cbu_digits) != 22:
        print(f'AVISO: {src} tiene CBU con longitud invalida '
              f'({rec.get("cbu")!r}, {len(cbu_digits)} digitos en vez de 22), revisar a mano.')
    cuit_digits = re.sub(r'\D', '', str(rec.get('cuit') or ''))
    if rec.get('cuit') and len(cuit_digits) != 11:
        print(f'AVISO: {src} tiene CUIT con longitud invalida '
              f'({rec.get("cuit")!r}, {len(cuit_digits)} digitos en vez de 11), revisar a mano.')


def generar_ops(records, dest, template, unidad, first_num, apply, fecha_op=None, sin_numero=False):
    """Genera una OP por registro, numerando desde first_num (o sin numero si
    sin_numero: carpeta y archivo 'OP_SIN_NUMERO_<PROVEEDOR>', celda N° OP vacia).
    Con apply=False solo muestra lo que haria. Devuelve [(nombre 'OP###',
    carpeta)] de las OP escritas (vacia en dry-run)."""
    generated_folders = []
    for op_num, (rec, src) in enumerate(records, start=first_num):
        avisos_registro(rec, src)
        slug = re.sub(r'[^A-Za-z0-9]+', '_', (rec.get('proveedor') or 'SIN_NOMBRE')).strip('_').upper()
        if sin_numero:
            op_num, etiqueta = None, 'OP sin numero'
            out_folder = os.path.join(dest, f'OP_SIN_NUMERO_{slug}')
            out_path = os.path.join(out_folder, f'OP_SIN_NUMERO_{slug}_{unidad}.xlsx')
            if apply and os.path.exists(out_folder):
                raise SystemExit(f'ERROR: ya existe {out_folder}')
        else:
            etiqueta = f'OP{op_num}'
            out_folder = os.path.join(dest, f'OP{op_num}')
            out_path = os.path.join(out_folder, f'OP{op_num}_{slug}_{unidad}.xlsx')

        print(f'{etiqueta}: {rec.get("proveedor")!r} | ${rec.get("importe")} | '
              f'fecha item {rec.get("fecha_item")} | CUIT {rec.get("cuit")} | '
              f'banco {rec.get("banco")} | CBU {rec.get("cbu")} | alias {rec.get("alias")} | fuente: {src}')
        lineas = lineas_detalle(rec)
        if len(lineas) > 1:
            for l in lineas:
                print(f'      {l["importe"]:>15,.2f}  {l["detalle"]}')
            suma = round(sum(l['importe'] for l in lineas), 2)
            if rec.get('importe') is not None and abs(suma - float(rec['importe'])) > 0.01:
                print(f'AVISO: {src}: el desglose suma {suma:,.2f} y el importe del registro '
                      f'es {float(rec["importe"]):,.2f}. La OP usa la suma del desglose.')
        print(f'   -> {out_path}')

        if apply:
            build_op(template, out_path, op_num, rec, unidad, fecha_op)
            if src and os.path.isfile(src):
                shutil.move(src, os.path.join(out_folder, os.path.basename(src)))
            generated_folders.append((etiqueta, out_folder))
    return generated_folders


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--dest', default=None,
                     help='Carpeta donde crear las OP (ej: la carpeta "Unit B" del dia). '
                          'Si se omite, se calcula sola a partir de la fecha de HOY y --unidad '
                          '(ej: ".../<mes>/<DD - MM DiaSemana>/<Razon social>").')
    ap.add_argument('--unidad', required=True,
                     help='Unidad de negocio para el nombre de archivo y la numeracion, '
                          'ej: UNIT_B, ACME_SA, UNIT_C, UNIT_D')
    ap.add_argument('--root', default=None,
                     help='Raiz de "Ordenes de Pago" para buscar el proximo numero de OP '
                          '(default: se detecta subiendo desde --dest)')
    ap.add_argument('--template', default=None,
                     help='Ruta a MODELO_OP.xlsx (default: junto a este script)')
    ap.add_argument('--auto', nargs='*', default=[],
                     help='Archivos de OC en xlsx a autoextraer')
    ap.add_argument('--manual', default=None,
                     help='JSON con lista de registros para OC leidas a mano (PDFs, etc.)')
    ap.add_argument('--fecha', default=None,
                     help='Fecha de las OP (YYYY-MM-DD) para la celda FECHA y la carpeta del '
                          'dia autocalculada; default: hoy. Ej: dejarlas listas para manana.')
    ap.add_argument('--primer-numero', type=int, default=None,
                     help='Numerar desde este OP en vez de "ultimo + 1" (para encastrar una OP '
                          'en el medio, ej. una del dia que va antes de las de manana). Antes hay que '
                          'liberar esos numeros renumerando las posteriores; si ya estan usados, corta.')
    ap.add_argument('--sin-numero', action='store_true',
                     help='Armar la OP SIN numero (para una fecha futura: el numero se define el dia de pago). '
                          'No ocupa ningun numero de la secuencia; carpeta y archivo quedan como '
                          'OP_SIN_NUMERO_<PROVEEDOR>. No se puede combinar con --email.')
    ap.add_argument('--apply', action='store_true',
                     help='Sin esta bandera el script solo simula (dry-run) y no escribe nada')
    ap.add_argument('--email', nargs='*', default=None,
                     help='Manda las OP generadas por mail (solo con --apply). '
                          'Siempre va a la lista fija de destinatarios '
                          '(DEFAULT_RECIPIENTS); las direcciones que pases acá '
                          'se suman ademas de esa lista, no la reemplazan. '
                          'Requiere las variables de entorno SMTP_USER y SMTP_PASS.')
    ap.add_argument('--email-subject', default=None,
                     help='Asunto del mail (default: "Ordenes de pago - <razon social> '
                          '(OPs <numeros>)", ej: "Ordenes de pago - Acme Unit B '
                          '(OPs 422 - 423 - 424)")')
    ap.add_argument('--imprimir', action='store_true',
                     help='Manda a imprimir (impresora predeterminada de Windows, 1 copia) '
                          'el xlsx de cada OP generada y, si la carpeta tiene la factura '
                          '(PDF/imagen), tambien la factura; no imprime la OC. '
                          'Solo con --apply, desde WSL con Excel instalado.')
    ap.add_argument('--razon-social', default=None,
                     help='Texto de razon social para el cuerpo del mail '
                          '(default: se deriva de --unidad, ej. UNIT_B -> '
                          '"Acme Unit B")')
    args = ap.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    template = args.template or os.path.join(script_dir, 'MODELO_OP.xlsx')
    if not os.path.isfile(template):
        raise SystemExit(f'No encuentro el template en {template}')

    fecha_op = datetime.date.fromisoformat(args.fecha) if args.fecha else None

    if args.dest:
        dest = args.dest
        root = args.root or detect_root(dest)
    else:
        root = args.root or os.path.dirname(script_dir)
        dest = auto_dest(root, args.unidad, fecha_op)

    hoy = fecha_op or datetime.date.today()
    print(f'Fecha de las OP: {hoy.strftime("%d/%m/%Y")} ({DIAS_ES[hoy.weekday()]})')
    print(f'Carpeta destino: {dest}')
    print(f'Razon social (F1 de la OP): {razon_social_f1(args.unidad)}')
    print()

    records = []
    for f in args.auto:
        records.append((extract_from_oc_xlsx(f), f))
    if args.manual:
        with open(args.manual, encoding='utf-8') as fh:
            for rec in json.load(fh):
                records.append((rec, rec.get('fuente')))

    if not records:
        print('No hay OC para procesar (usa --auto y/o --manual).')
        return

    if args.sin_numero and (args.email is not None or args.primer_numero):
        raise SystemExit('ERROR: --sin-numero no se combina con --email ni --primer-numero.')
    next_num = next_op_number(root, args.unidad)
    if args.primer_numero:
        ocupados = sorted(n for n in numeros_usados(root, args.unidad)
                          if args.primer_numero <= n < args.primer_numero + len(records))
        if ocupados:
            raise SystemExit(f'ERROR: --primer-numero {args.primer_numero}: ya existen OP{ocupados} '
                             f'en {args.unidad}. Renumera primero las posteriores.')
        next_num = args.primer_numero
    print(f'Unidad: {args.unidad}  |  ' + ('OP SIN NUMERO (el proximo libre seguiria siendo OP%d)' % next_num
                                          if args.sin_numero else f'Proximo numero libre: OP{next_num}'))
    print('MODO APLICAR (se va a escribir en disco)' if args.apply
          else 'DRY-RUN: no se escribe nada. Agrega --apply cuando confirmes los datos.')
    print()

    generated_folders = generar_ops(records, dest, template, args.unidad, next_num, args.apply,
                                    fecha_op, args.sin_numero)

    print()
    print('Listo.' if args.apply else 'Dry-run terminado. Volve a correr con --apply para generar los archivos de verdad.')

    if args.imprimir:
        if args.apply:
            imprimir_ops(generated_folders)
        else:
            print('AVISO: --imprimir se ignora en dry-run; solo imprime junto con --apply.')

    if args.email is not None:
        if not args.apply:
            print('AVISO: --email se ignora en dry-run; el mail solo se manda junto con --apply.')
            return
        razon_social = args.razon_social or display_razon_social(args.unidad)
        recipients, subject = email_de_ops(records, generated_folders, razon_social,
                                           args.email, args.email_subject)
        send_ops_email(recipients, subject, generated_folders, razon_social)


if __name__ == '__main__':
    main()
