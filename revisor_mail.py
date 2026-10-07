#!/usr/bin/env python3
"""revisor_mail.py — revisa el mail y arma las OP solas (paso 1 de la automatizacion).

Lee la casilla (IMAP, SOLO LECTURA: nunca borra, mueve ni marca nada), detecta
solicitudes de pago y, si estan completas, genera la OP en la carpeta del dia
con generar_ops.py y la IMPRIME (OP + factura, 1 copia; se apaga con
"imprimir_automatico": false en el config; tope "max_impresiones_por_corrida", 15).
NO manda mails a los destinatarios: las OP quedan "pendientes de envio" hasta que se
apruebe con --enviar. Cada OP se imprime una sola vez (marca "impresa" en el estado).

    python3 revisor_mail.py --inicializar        # 1ra vez: ignora todo el mail previo
    python3 revisor_mail.py                      # DRY-RUN: muestra que haria
    python3 revisor_mail.py --apply              # genera las OP y avisa
    python3 revisor_mail.py --enviar             # DRY-RUN: lista lo pendiente de envio
    python3 revisor_mail.py --enviar --apply     # manda los mails (tras el OK del usuario)
    python3 revisor_mail.py --imprimir [--apply] # imprime las pendientes sin imprimir (reintenta fallidas)
    python3 revisor_mail.py --estado             # resumen del estado
    python3 revisor_mail.py --agregar-pendiente "<carpeta OP###>" --unidad ACME_SA

QUE ES UNA SOLICITUD DE PAGO: mail NUEVO (UID mayor al ultimo revisado) de un
remitente de "remitentes" en revisor_config.json, con un xlsx adjunto que sea una
OC (celda A1 = "N° OC"). Si trae PDF o imagen (facturas, OC en PDF) va a la lista
"revisar" (no se genera nada solo: eso es el paso 3).

COMPUERTA: la OP solo se genera si la OC esta completa. Va a "revisar" (y no se
genera) si falta proveedor o importe, si no hay datos bancarios y no es efectivo,
si la empresa pagadora (celda F1 de la OC) no es una unidad conocida, o si el
desglose no suma el total. Sin CUIT/RESP INSCRIPTO es solo una advertencia.

PASO 3 (facturas sin OC): un mail autorizado SIN OC en xlsx pero con PDF/imagen se lee con
Claude (lector_facturas.py, ~US$0,03 c/u, tope "max_lecturas_por_corrida", se apaga con
"lectura_claude": false o --sin-claude). Pasa por una compuerta mas estricta: CUIT valido,
numero de factura, importe, receptor que sea una unidad conocida, datos bancarios (de la factura o
de la ultima OP de ese CUIT), y no duplicar facturas ya cargadas. La OP sale como las demas.
BOLETAS DE IMPUESTOS (ARBA, etc.): Claude las clasifica como "boleta_impuesto" y el config decide como se
cargan: "boletas_impuestos" = [{organismo, partida, unidad, proveedor, condicion_pago}]. Solo las partidas
listadas se generan solas (ej. ARBA Inmobiliario 000-000000-0 -> Unit B, "ARBA BRAND", PAGO ONLINE, n° de
factura = codigo de pago, desglose de la boleta). Una partida sin regla, vencida, sin codigo de pago o con
una OP ya hecha (mismo codigo e importe, 60 dias) va a "revisar". La fecha de la OP nunca pasa del vencimiento.

EMPRESA Y NUMERACION: la unidad sale de la celda F1 de la OC (la empresa que
paga), ej. "ACME SA" -> ACME_SA, "ACME UNIT_B S.A." -> UNIT_B.

FECHA DE LA OP: hoy si la OC dice "ya pago"; si no, el proximo dia habil
(viernes -> lunes; no contempla feriados).

Archivos (todos ignorados por git): revisor_config.json, estado_revisor.json, _revisor/.
"""
import argparse
import contextlib
import datetime
import io
import email
import fcntl
import hashlib
import imaplib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from email.header import decode_header
from email.utils import parseaddr

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import generar_ops as g  # noqa: E402
import lector_facturas as lf  # noqa: E402

ROOT = os.path.dirname(HERE)  # "Ordenes de Pago"
CONFIG = os.path.join(HERE, 'revisor_config.json')
ESTADO = os.path.join(HERE, 'estado_revisor.json')
WORK = os.path.join(HERE, '_revisor')
if os.environ.get('REVISOR_SANDBOX'):   # pruebas: raiz, estado y trabajo en otra carpeta (nada toca lo real)
    _sb = os.environ['REVISOR_SANDBOX']
    ROOT, ESTADO, WORK = os.path.join(_sb, 'raiz'), os.path.join(_sb, 'estado.json'), os.path.join(_sb, 'work')
    os.makedirs(ROOT, exist_ok=True)
IMAP_HOST = os.environ.get('IMAP_HOST', 'imap.example.com')

# F1 de la OC (empresa que paga) -> --unidad de generar_ops. Orden = prioridad.
UNIDADES_F1 = [
    ('UNIT_B', 'UNIT_B'), ('UNIT_C', 'UNIT_C'), ('UNIT_D', 'UNIT_D'),
    ('CLEANER', 'CLEANER'), ('CITY', 'CITY'), ('SUPPLIER_Y', 'SUPPLIER_Y'), ('SUPPLIER_Y', 'SUPPLIER_Y'), ('BUY', 'BUY'),
    ('ACME', 'ACME_SA'),   # "ACME SA" / "ACME S.A." (va ultimo: el resto lleva ACME tambien)
]


# --------------------------------------------------------------------------- utilidades

def dh(s):
    out = ''
    for t, c in decode_header(s or ''):
        out += t.decode(c or 'utf-8', 'replace') if isinstance(t, bytes) else t
    return out


def limpiar_nombre(n):
    return re.sub(r'\s+', ' ', n).strip()


def cargar(path, default):
    if os.path.isfile(path):
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    return default


def guardar(path, data):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def cuentas_config(cfg):
    """Lista de casillas a revisar. Cada una: {usuario_env, clave_env, remitentes}.
    (Compat: un config viejo con "remitentes" sueltos usa la casilla SMTP_USER.)"""
    if cfg.get('cuentas'):
        return cfg['cuentas']
    return [dict(usuario_env='SMTP_USER', clave_env='SMTP_PASS', remitentes=cfg.get('remitentes', []))]


def direccion(cuenta):
    return os.environ.get(cuenta['usuario_env'], '')


def conectar(cuenta):
    g._load_dotenv()
    user, pw = os.environ.get(cuenta['usuario_env']), os.environ.get(cuenta['clave_env'])
    if not user or not pw:
        raise SystemExit(f'Faltan {cuenta["usuario_env"]} / {cuenta["clave_env"]} (.env).')
    m = imaplib.IMAP4_SSL(IMAP_HOST, 993)
    m.login(user, pw)
    typ, _ = m.select('INBOX', readonly=True)   # readonly: no marca nada como leido
    if typ != 'OK':
        raise SystemExit(f'No pude abrir INBOX de {user}.')
    return m


def uidvalidity(m):
    typ, d = m.response('UIDVALIDITY')
    return int(d[0]) if d and d[0] else None


def max_uid(m):
    typ, d = m.uid('search', None, 'ALL')
    ids = d[0].split()
    return int(ids[-1]) if ids else 0


def proximo_dia_habil(hoy):
    d = hoy + datetime.timedelta(days=1)
    while d.weekday() >= 5:
        d += datetime.timedelta(days=1)
    return d


def unidad_de_f1(f1):
    t = (f1 or '').upper()
    for clave, unidad in UNIDADES_F1:
        if clave in t:
            return unidad
    return None


def notificar(titulo, texto, urgente=False):
    """Ventana de Windows que QUEDA en pantalla hasta que se cierra (el globo de 15 s se perdia si no
    se estaba mirando). El aviso se escribe a un .ps1 temporal que se lanza con Start-Process: sigue
    vivo aunque termine la corrida/WSL y no hay comillas anidadas. Best effort: si falla, el
    resumen igual queda en _revisor/resumen.log."""
    try:
        os.makedirs(WORK, exist_ok=True)
        fd, ps1 = tempfile.mkstemp(prefix='aviso_', suffix='.ps1', dir=WORK)
        with os.fdopen(fd, 'w', encoding='utf-8-sig') as f:   # BOM: PowerShell 5.1 lee bien los acentos
            f.write("Add-Type -AssemblyName System.Windows.Forms\n"
                    "$f = New-Object System.Windows.Forms.Form; $f.TopMost = $true\n"
                    f"$texto = @'\n{texto}\n'@\n"
                    f"[void][System.Windows.Forms.MessageBox]::Show($f, $texto, '{titulo}', 'OK', '{'Warning' if urgente else 'Information'}')\n"
                    "Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue\n")
        win = subprocess.check_output(['wslpath', '-w', ps1], text=True).strip()
        subprocess.Popen(['powershell.exe', '-NoProfile', '-Command',
                          f"Start-Process powershell -WindowStyle Hidden -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','{win}'"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass


# --------------------------------------------------------------------------- analisis de un mail

def es_oc_xlsx(path):
    try:
        grid = g.read_grid(path)
    except Exception:
        return False
    return str(grid.get('A1', '')).replace('°', '').replace('º', '').strip().upper() in ('N OC', 'NOC')


def compuerta(rec, unidad, sin_banco_ok=()):
    """Devuelve (problemas, advertencias). Con problemas NO se genera la OP."""
    prob, adv = [], []
    if not rec.get('proveedor'):
        prob.append('falta RAZON SOCIAL del proveedor')
    if not rec.get('importe'):
        prob.append('falta el importe (o es 0)')
    efectivo = 'efectivo' in (rec.get('condicion_pago') or '').lower()
    exento = any(x.upper() in (rec.get('proveedor') or '').upper() for x in sin_banco_ok)
    if not efectivo and not exento and not (rec.get('cbu') or rec.get('alias') or rec.get('n_cuenta')):
        prob.append('no tiene datos bancarios (CBU / alias / cuenta) y no es efectivo')
    if rec.get('cbu') and len(rec['cbu']) != 22:
        prob.append(f'el CBU tiene {len(rec["cbu"])} digitos en vez de 22')
    if not unidad:
        prob.append('la empresa pagadora (F1 de la OC) no coincide con ninguna unidad conocida')
    if not rec.get('cuit'):
        adv.append('sin CUIT')
    elif len(rec['cuit']) != 11:
        adv.append(f'CUIT con {len(rec["cuit"])} digitos')
    if not rec.get('resp_inscripto'):
        adv.append('sin RESP INSCRIPTO')
    return prob, adv


def analizar_mail(m, uid, remitentes, workdir):
    """Baja el mail y devuelve un dict con lo detectado (sin generar nada)."""
    typ, d = m.uid('fetch', str(uid), '(BODY.PEEK[])')
    msg = email.message_from_bytes(d[0][1])
    remitente = parseaddr(dh(msg['From']))[1].lower()
    info = dict(uid=uid, remitente=remitente, asunto=dh(msg['Subject']), fecha_mail=msg['Date'], ocs=[], revisar=[])
    if remitente not in [r.lower() for r in remitentes]:
        info['estado'] = 'ignorado (remitente no autorizado)'
        return info
    carpeta = os.path.join(workdir, str(uid))
    os.makedirs(carpeta, exist_ok=True)
    for part in msg.walk():
        fn = part.get_filename()
        if not fn:
            continue
        fn = limpiar_nombre(dh(fn))
        low = fn.lower()
        if not low.endswith(('.xlsx', '.pdf', '.jpg', '.jpeg', '.png')):
            continue
        if part.get_content_disposition() == 'inline' and low.startswith('image'):
            continue   # firmas del mail
        path = os.path.join(carpeta, fn)
        with open(path, 'wb') as f:
            f.write(part.get_payload(decode=True))
        if low.endswith('.xlsx') and es_oc_xlsx(path):
            info['ocs'].append(path)
        else:
            info['revisar'].append(path)
    info['estado'] = 'candidato' if (info['ocs'] or info['revisar']) else 'ignorado (sin adjuntos de pago)'
    return info


def preparar_oc(path, hoy, sin_banco_ok=()):
    """Extrae, aplica la compuerta y decide unidad y fecha. No escribe nada."""
    rec = g.extract_from_oc_xlsx(path)
    grid = g.read_grid(path)
    unidad = unidad_de_f1(grid.get('F1'))
    prob, adv = compuerta(rec, unidad, sin_banco_ok)
    fecha = hoy if rec.get('ya_pagada') else proximo_dia_habil(hoy)
    return dict(path=path, rec=rec, unidad=unidad, f1=grid.get('F1'), fecha=fecha, problemas=prob, advertencias=adv)


# --------------------------------------------------------------------------- acciones

def generar(prep, apply, simulados):
    """Genera la OP (dry-run si no apply). Devuelve (op_name, folder) o None.
    `simulados` lleva los numeros ya "usados" en un dry-run para no repetirlos."""
    unidad, fecha, rec = prep['unidad'], prep['fecha'], prep['rec']
    dest = g.auto_dest(ROOT, unidad, fecha)
    num = g.next_op_number(ROOT, unidad)
    if not apply:
        num = max(num, simulados.get(unidad, 0) + 1)
        simulados[unidad] = num
        return f'OP{num}', dest
    template = os.path.join(HERE, 'MODELO_OP.xlsx')
    with contextlib.redirect_stdout(io.StringIO()):   # el detalle por OP va al resumen, no a la pantalla
        carpetas = g.generar_ops([(rec, prep['path'])], dest, template, unidad, num, True, fecha, False)
    return carpetas[0] if carpetas else None


def imprimir_pendientes(estado, cfg, apply, reintentar=False):
    """Imprime (OP + factura) las OP pendientes que todavia no se imprimieron. Cada OP se
    marca "impresa" apenas sale; si falla NO se reintenta sola (se marca impresion_error y
    se avisa) salvo reintentar=True. Tope por corrida: cfg["max_impresiones_por_corrida"] (15)."""
    tope = int(cfg.get('max_impresiones_por_corrida', 15))
    lista = [p for p in estado.get('pendientes_envio', [])
             if not p.get('impresa') and (reintentar or not p.get('impresion_error'))]
    out = []
    for i, p in enumerate(lista):
        if i >= tope:
            out.append(f'IMPRESION tope de {tope} por corrida: quedan {len(lista) - tope} OP sin imprimir (se imprimen con --imprimir --apply)')
            break
        if not apply:
            out.append(f'IMPRIMIRIA {p["op"]} {p["proveedor"]} (OP + factura si la carpeta la tiene)')
            continue
        if os.environ.get('REVISOR_SANDBOX'):
            out.append(f'IMPRESION omitida en sandbox: {p["op"]}')
            continue
        try:
            ok = g.imprimir_ops([(p['op'], os.path.join(ROOT, p['folder']))])
        except Exception as e:   # la impresion nunca corta la corrida
            print(f'AVISO: impresion de {p["op"]} fallo: {e}')
            ok = False
        p['impresa'] = bool(ok)
        p['impresion_error'] = not ok
        guardar(ESTADO, estado)
        out.append(f'IMPRESA   {p["op"]} {p["proveedor"]}' if ok else f'IMPRESION FALLO {p["op"]} {p["proveedor"]} (ver tarea.log; reintentar con --imprimir --apply)')
    return out


def migrar_estado(estado, cfg):
    """Estado viejo (1 casilla: last_uid / uidvalidity sueltos) -> estado por cuenta."""
    if 'cuentas' not in estado:
        estado['cuentas'] = {}
        if 'last_uid' in estado:
            primera = direccion(cuentas_config(cfg)[0])
            estado['cuentas'][primera] = dict(uidvalidity=estado.pop('uidvalidity', None), last_uid=estado.pop('last_uid'))
    return estado


def elegir_cuentas(args, cfg):
    cs = cuentas_config(cfg)
    if args.cuenta:
        cs = [c for c in cs if args.cuenta.lower() in direccion(c).lower()]
        if not cs:
            raise SystemExit(f'--cuenta {args.cuenta}: no coincide con ninguna casilla del config.')
    return cs


def revisar_cuenta(args, cfg, estado, cuenta, hoy, ctx):
    """Revisa una casilla. Devuelve la lista de lineas de resumen."""
    addr = direccion(cuenta)
    punto = estado['cuentas'].get(addr)
    if not punto:
        raise SystemExit(f'{addr}: primero corre --inicializar.')
    m = conectar(cuenta)
    if punto.get('uidvalidity') and punto['uidvalidity'] != uidvalidity(m):
        raise SystemExit(f'{addr}: UIDVALIDITY cambio (la casilla se reindexo): corre --inicializar de nuevo.')
    ultimo = punto['last_uid'] if args.desde_uid is None else args.desde_uid
    if args.uids:   # corrida puntual sobre mails ya viejos: no mira lo nuevo ni mueve el puntero
        uids = sorted(int(x) for x in args.uids.split(','))
        resumen = [f'--- {addr}: {len(uids)} mails puntuales (--uids) ---']
        remitentes = list(cuenta['remitentes']) + [x for x in (args.extra_remitentes or '').split(',') if x]
    else:
        typ, d = m.uid('search', None, f'UID {ultimo + 1}:*')
        uids = sorted(int(x) for x in d[0].split() if int(x) > ultimo)
        resumen = [f'--- {addr}: {len(uids)} mails nuevos desde UID {ultimo} ---']
        remitentes = cuenta['remitentes']
    workdir = tempfile.mkdtemp(prefix='revisor_')
    simulados = {}
    try:
        for uid in uids:
            nuevos_pend, nuevos_rev, lineas = [], [], []
            try:
                info = analizar_mail(m, uid, remitentes, workdir)
                if not info['estado'].startswith('ignorado'):
                    procesar_mail(args, cfg, hoy, info, uid, simulados, nuevos_pend, nuevos_rev, lineas, ctx)
            except Exception as e:   # un mail roto no frena a los demas: queda en "revisar"
                lineas.append(f'REVISAR   #{uid}: ERROR procesando el mail ({type(e).__name__}: {e})')
                nuevos_rev.append(dict(uid=uid, remitente='?', asunto='?', archivo='?', problemas=[f'error: {e}']))
            resumen += [f'[{addr.split("@")[0]}] ' + l for l in lineas]
            if args.apply:   # se guarda mail por mail: si algo falla despues no se pierde ni se duplica nada
                os.makedirs(os.path.join(WORK, 'revisar'), exist_ok=True)
                for r in nuevos_rev:
                    r['cuenta'] = addr
                    src = r.pop('_src', None)
                    if src:
                        dst = os.path.join(WORK, 'revisar', f'{addr.split("@")[0]}_{r["uid"]}_{r["archivo"]}')
                        shutil.copy2(src, dst)
                        r['ruta'] = os.path.relpath(dst, ROOT)
                for p in nuevos_pend:
                    p['cuenta'] = addr
                estado['pendientes_envio'] = estado.get('pendientes_envio', []) + nuevos_pend
                estado['revisar'] = estado.get('revisar', []) + nuevos_rev
                estado.setdefault('hashes', {}).update(ctx['nuevos'])
                ctx['nuevos'].clear()
                if not args.uids:
                    punto['last_uid'] = max(punto['last_uid'], uid)   # nunca retrocede (aunque se pruebe con --desde-uid)
                estado['ultima_revision'] = datetime.datetime.now().isoformat(timespec='seconds')
                guardar(ESTADO, estado)
        if args.apply and uids and args.desde_uid is None and not args.uids:
            punto['last_uid'] = max(punto['last_uid'], uids[-1])
            guardar(ESTADO, estado)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
    return resumen


FIRMA_IMAGEN = re.compile(r'^image\d*\.(png|jpe?g|gif)$', re.I)   # logos/firmas de los mails


def candidata_a_lectura(path):
    return (path.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png'))
            and not FIRMA_IMAGEN.match(os.path.basename(path)) and os.path.getsize(path) >= 15000)


def procesar_factura(args, cfg, hoy, info, uid, path, simulados, nuevos_pend, nuevos_rev, resumen, ctx):
    """Paso 3: una factura (PDF/imagen) de un mail SIN OC. Devuelve True si quedo resuelta
    (OP generada / omitida / mandada a revisar con motivo), False si no se pudo leer (queda en "sueltos")."""
    nombre = os.path.basename(path)
    etiqueta = f'#{uid} {info["remitente"]} | {info["asunto"][:50]} | {nombre}'
    with open(path, 'rb') as f:
        sha = hashlib.sha256(f.read()).hexdigest()
    previo = ctx['hashes'].get(sha) or ctx['nuevos'].get(sha)
    if previo:
        resumen.append(f'OMITIDO   {etiqueta}: archivo identico ya procesado ({previo["resultado"]})')
        return True
    if ctx['lecturas'] >= int(cfg.get('max_lecturas_por_corrida', 6)):
        resumen.append(f'REVISAR   {etiqueta}: tope de {cfg.get("max_lecturas_por_corrida", 6)} lecturas por corrida alcanzado, no se leyo')
        return False
    ctx['lecturas'] += 1
    try:
        datos, costo = lf.leer_factura(path, cfg, cache_dir=os.path.join(WORK, 'lecturas'))
    except Exception as e:
        resumen.append(f'REVISAR   {etiqueta}: no se pudo leer con Claude ({e})')
        return False
    ctx['costo'] += costo
    if ctx['hist'] is None:
        ctx['hist'] = lf.indice_historico(ROOT)
    if datos.get('tipo') == 'boleta_impuesto':
        res = lf.registro_desde_boleta(datos, ctx['hist'], cfg, hoy)
    else:
        res = lf.registro_desde_factura(datos, ctx['hist'], cfg, hoy, unidad_de_f1, pistas=(nombre, info['asunto']))
    if res['omitir']:
        resumen.append(f'OMITIDO   {etiqueta}: {res["omitir"]}')
        ctx['nuevos'][sha] = dict(resultado='omitido: ' + res['omitir'], fecha=hoy.isoformat())
        return True
    rec = res['rec']
    base = (f'{etiqueta}: {rec.get("proveedor")} ${rec.get("importe") or 0:,.2f} | FC {rec.get("n_factura")} | '
            f'unidad {res["unidad"]} | fecha OP {proximo_dia_habil(hoy):%d/%m/%Y}')
    adv = f' | ADVERTENCIAS: {"; ".join(res["advertencias"])}' if res['advertencias'] else ''
    if res['problemas']:
        resumen.append(f'REVISAR   {base} -> NO se genera: ' + '; '.join(res['problemas']) + adv)
        nuevos_rev.append(dict(uid=uid, remitente=info['remitente'], asunto=info['asunto'], archivo=nombre,
                               problemas=res['problemas'], leido=rec, _src=path))
        ctx['nuevos'][sha] = dict(resultado='revisar', fecha=hoy.isoformat())
        return True
    fecha_op = proximo_dia_habil(hoy)
    if res.get('vencimiento') and fecha_op > res['vencimiento']:   # boleta: nunca despues del vencimiento
        fecha_op = hoy
    prep = dict(path=path, rec=rec, unidad=res['unidad'], fecha=fecha_op, problemas=[], advertencias=res['advertencias'])
    gen = generar(prep, args.apply, simulados)
    if gen:
        op_name, folder = gen
        if args.apply:
            nuevos_pend.append(dict(op=op_name, folder=os.path.relpath(folder, ROOT), unidad=res['unidad'],
                                    fecha=prep['fecha'].isoformat(), ya_pagada=False, proveedor=rec['proveedor'],
                                    importe=rec['importe'], uid=uid, origen='factura leida por Claude'))
            ctx['nuevos'][sha] = dict(resultado=op_name, fecha=hoy.isoformat())
        # que otra factura igual en la MISMA corrida (otro mail) no genere una segunda OP
        nf = lf.norm_factura(rec.get('n_factura'))
        if rec.get('cuit') and nf:
            ctx['hist']['facturas'][(rec['cuit'], nf)] = op_name
        elif rec.get('n_factura') and not rec.get('cuit'):
            ctx['hist'].setdefault('codigos', {}).setdefault(re.sub(r'\D', '', rec['n_factura']).lstrip('0'), []).append(
                (round(rec['importe'], 2), hoy, op_name))
        resumen.append(f'{"GENERADA " if args.apply else "GENERARIA"} {base} -> {op_name} | lectura Claude{adv}')
    return True


def procesar_mail(args, cfg, hoy, info, uid, simulados, nuevos_pend, nuevos_rev, resumen, ctx):
    etiqueta = f'#{uid} {info["remitente"]} | {info["asunto"][:60]}'
    # Un mail con UNA sola OC: sus PDF/imagenes son la factura y viajan con la OP
    # (excepto comprobantes de pago). Con 0 o varias OC no se puede asignar solo.
    facturas, sueltos = [], list(info['revisar'])
    if len(info['ocs']) == 1:
        facturas = [x for x in sueltos if not re.search(r'comprobante|pago', os.path.basename(x), re.I)]
        sueltos = [x for x in sueltos if x not in facturas]
    if sueltos and not info['ocs'] and ctx['habilitada']:   # paso 3: mail sin OC -> Claude lee las facturas
        sueltos = [x for x in sueltos
                   if not (candidata_a_lectura(x) and procesar_factura(args, cfg, hoy, info, uid, x, simulados,
                                                                       nuevos_pend, nuevos_rev, resumen, ctx))]
    if sueltos:
        resumen.append(f'REVISAR   {etiqueta}: adjuntos sin asignar a una OP: '
                       + ', '.join(os.path.basename(x) for x in sueltos))
        for x in sueltos:
            nuevos_rev.append(dict(uid=uid, remitente=info['remitente'], asunto=info['asunto'],
                                   archivo=os.path.basename(x), _src=x))
    for oc in info['ocs']:
        prep = preparar_oc(oc, hoy, cfg.get('sin_datos_bancarios_ok', []))
        rec = prep['rec']
        base = (f'{etiqueta}: {rec.get("proveedor")} ${rec.get("importe"):,.2f} | unidad {prep["unidad"]} | '
                f'fecha OP {prep["fecha"].strftime("%d/%m/%Y")}{" (YA PAGA)" if rec.get("ya_pagada") else ""}')
        if prep['problemas']:
            resumen.append(f'REVISAR   {base} -> NO se genera: ' + '; '.join(prep['problemas']))
            nuevos_rev.append(dict(uid=uid, remitente=info['remitente'], asunto=info['asunto'],
                                   archivo=os.path.basename(oc), problemas=prep['problemas'], _src=oc))
            for x in facturas:
                nuevos_rev.append(dict(uid=uid, remitente=info['remitente'], asunto=info['asunto'],
                                       archivo=os.path.basename(x), _src=x))
            continue
        res = generar(prep, args.apply, simulados)
        adv = f' | ADVERTENCIAS: {", ".join(prep["advertencias"])}' if prep['advertencias'] else ''
        fac = (' | + factura: ' + ', '.join(os.path.basename(x) for x in facturas)) if facturas else ' | SIN factura adjunta'
        if res:
            op_name, folder = res
            if args.apply:
                for x in facturas:
                    shutil.copy2(x, os.path.join(folder, os.path.basename(x)))
                nuevos_pend.append(dict(op=op_name, folder=os.path.relpath(folder, ROOT), unidad=prep['unidad'],
                                        fecha=prep['fecha'].isoformat(), ya_pagada=bool(rec.get('ya_pagada')),
                                        proveedor=rec.get('proveedor'), importe=rec.get('importe'), uid=uid))
            resumen.append(f'{"GENERADA " if args.apply else "GENERARIA"} {base} -> {op_name}{fac}{adv}')


def armar_aviso(sin_prefijo, rev, cfg):
    """Texto de la ventana de aviso (lineas del resumen sin el prefijo [casilla]). Devuelve (texto, urgente)."""
    impresas = [r.split(None, 1)[1] for r in sin_prefijo if r.startswith('IMPRESA ')]
    fallidas = [r for r in sin_prefijo if r.startswith('IMPRESION FALLO') or r.startswith('IMPRESION tope')]
    generadas = [r for r in sin_prefijo if r.startswith('GENERADA')]
    partes = []
    if impresas:
        partes.append(f'SE IMPRIMIERON {len(impresas)} OP (con su factura si la tenian):\n' + '\n'.join(f'  - {x}' for x in impresas))
    if fallidas:
        partes.append('ATENCION - NO SE IMPRIMIO:\n' + '\n'.join(f'  - {x}' for x in fallidas) + '\nImprimilas con: revisor_mail.py --imprimir --apply')
    if generadas and not cfg.get('imprimir_automatico', True):
        partes.append(f'ATENCION: la impresion automatica esta APAGADA; {len(generadas)} OP quedaron SIN imprimir.')
    elif generadas and not impresas and not fallidas:
        partes.append(f'ATENCION: se generaron {len(generadas)} OP pero NO se imprimio ninguna (ya estaban marcadas como impresas o no hubo cupo).')
    if rev:
        partes.append(f'{rev} mail(s)/adjunto(s) para revisar a mano (revisor_mail.py --estado).')
    if generadas or impresas:
        partes.append('Falta tu OK para mandar el mail.')
    return '\n\n'.join(partes), bool(fallidas) or (not impresas and bool(generadas))


def revisar_mail(args):
    cfg = cargar(CONFIG, None)
    if not cfg or not cuentas_config(cfg)[0].get('remitentes'):
        raise SystemExit(f'Falta {CONFIG} con las "cuentas" y sus "remitentes" (ver revisor_config.ejemplo.json).')
    estado = cargar(ESTADO, None)
    if not estado:
        raise SystemExit('Primero corre --inicializar (marca el mail actual como ya revisado).')
    migrar_estado(estado, cfg)
    cuentas = elegir_cuentas(args, cfg)
    if (args.desde_uid is not None or args.uids) and len(cuentas) != 1:
        raise SystemExit('--desde-uid / --uids necesitan --cuenta (un UID solo tiene sentido dentro de una casilla).')
    if args.max_lecturas:
        cfg['max_lecturas_por_corrida'] = args.max_lecturas
    hoy = datetime.date.today()
    todo = []
    ctx = dict(habilitada=cfg.get('lectura_claude', True) and not args.sin_claude, lecturas=0, costo=0.0, hist=None,
               hashes=estado.setdefault('hashes', {}), nuevos={})
    for cuenta in cuentas:
        todo += revisar_cuenta(args, cfg, estado, cuenta, hoy, ctx)
    if ctx['lecturas']:
        todo.append(f'LECTURA Claude: {ctx["lecturas"]} comprobante(s) leidos, costo aprox. US${ctx["costo"]:.3f}')
        if args.apply:
            estado['costo_claude_usd'] = round(estado.get('costo_claude_usd', 0) + ctx['costo'], 4)
            guardar(ESTADO, estado)
    if args.apply and cfg.get('imprimir_automatico', True):
        todo += imprimir_pendientes(estado, cfg, True)
    accionables = [l for l in todo if not l.startswith('---') and not l.startswith('LECTURA')]
    print('\n'.join(todo))
    if not accionables:
        print('Nada para hacer.')
        return
    if not args.apply:
        print('\nDRY-RUN: no se escribio nada ni se avanzo el puntero de mails. Agrega --apply.')
        return
    os.makedirs(WORK, exist_ok=True)
    with open(os.path.join(WORK, 'resumen.log'), 'a', encoding='utf-8') as f:
        f.write(f'\n== {estado["ultima_revision"]} ==\n' + '\n'.join(todo) + '\n')
    sin_prefijo = [r.split('] ', 1)[-1] for r in accionables]
    gen = sum(1 for r in sin_prefijo if r.startswith('GENERADA'))
    rev = sum(1 for r in sin_prefijo if r.startswith('REVISAR'))
    if not os.environ.get('REVISOR_SANDBOX'):   # las pruebas en sandbox nunca avisan
        texto, urgente = armar_aviso(sin_prefijo, rev, cfg)
        notificar('Ordenes de pago', texto, urgente=urgente)


def imprimir_cmd(args):
    cfg = cargar(CONFIG, {})
    estado = cargar(ESTADO, None)
    if not estado:
        raise SystemExit('Sin estado: corre --inicializar.')
    out = imprimir_pendientes(estado, cfg, args.apply, reintentar=True)
    print('\n'.join(out) if out else 'No hay OP pendientes de imprimir.')
    if out and not args.apply:
        print('\nDRY-RUN: no se imprimio nada. Agrega --apply.')


def inicializar(args):
    cfg = cargar(CONFIG, None)
    if not cfg:
        raise SystemExit(f'Falta {CONFIG}.')
    estado = migrar_estado(cargar(ESTADO, {}), cfg)
    for cuenta in elegir_cuentas(args, cfg):
        m = conectar(cuenta)
        addr = direccion(cuenta)
        estado['cuentas'][addr] = dict(uidvalidity=uidvalidity(m), last_uid=max_uid(m))
        print(f'{addr}: el mail existente hasta el UID {estado["cuentas"][addr]["last_uid"]} queda como ya revisado.')
    estado.setdefault('pendientes_envio', [])
    estado.setdefault('revisar', [])
    guardar(ESTADO, estado)


def enviar(args):
    estado = cargar(ESTADO, None)
    pend = [p for p in (estado or {}).get('pendientes_envio', [])]
    if not pend:
        print('No hay OP pendientes de envio.')
        return
    por_unidad = {}
    for p in pend:
        por_unidad.setdefault(p['unidad'], []).append(p)
    for unidad, ps in por_unidad.items():
        folders = [(p['op'], os.path.join(ROOT, p['folder'])) for p in ps]
        recs = [({'ya_pagada': p['ya_pagada']}, None) for p in ps]
        razon = g.display_razon_social(unidad)
        recipients, subject = g.email_de_ops(recs, folders, razon, [])
        print(f'{subject}\n   -> ' + ', '.join(f'{p["op"]} {p["proveedor"]} ${p["importe"]:,.2f} ({p["fecha"]})' for p in ps))
        if args.apply:
            g.send_ops_email(recipients, subject, folders, razon)
            estado['pendientes_envio'] = [p for p in estado['pendientes_envio'] if p not in ps]
            estado.setdefault('enviadas', []).extend(ps)
            guardar(ESTADO, estado)
    if not args.apply:
        print('\nDRY-RUN: no se mando nada. Agrega --apply cuando des el OK.')


def agregar_pendiente(args):
    folder = os.path.abspath(args.agregar_pendiente)
    op = os.path.basename(folder.rstrip('/'))
    if not re.fullmatch(r'OP\d+', op):
        raise SystemExit('La carpeta tiene que llamarse OP### (ej. OP296).')
    xlsx = next((f for f in os.listdir(folder) if f.upper().startswith('OP') and f.lower().endswith('.xlsx')), None)
    if not xlsx or not args.unidad:
        raise SystemExit('Falta el xlsx de la OP en la carpeta o --unidad.')
    grid = g.read_grid(os.path.join(folder, xlsx))
    estado = cargar(ESTADO, {})
    estado.setdefault('pendientes_envio', []).append(dict(
        op=op, folder=os.path.relpath(folder, ROOT), unidad=args.unidad,
        fecha=g.excel_date(float(grid['B2'])).isoformat(), ya_pagada=g.folder_has_comprobante(folder),
        proveedor=grid.get('B4'), importe=round(sum(float(grid.get(f'H{r}') or 0) for r in range(15, 21)), 2), uid=None))
    guardar(ESTADO, estado)
    print(f'{op} agregada a pendientes de envio.')


def mostrar_estado(args):
    e = cargar(ESTADO, None)
    if not e:
        print('Sin estado: corre --inicializar.')
        return
    cfg = cargar(CONFIG, {})
    migrar_estado(e, cfg)
    for addr, pt in e['cuentas'].items():
        print(f'{addr}: ultimo UID revisado {pt.get("last_uid")}')
    print(f'Ultima revision: {e.get("ultima_revision")}')
    print('Pendientes de envio:')
    for p in e.get('pendientes_envio', []):
        marca = 'impresa' if p.get('impresa') else ('IMPRESION FALLO' if p.get('impresion_error') else 'sin imprimir')
        print(f'   {p["op"]} {p["proveedor"]} ${p["importe"]:,.2f} | {p["unidad"]} | {p["fecha"]} | {marca}')
    print('Para revisar a mano:')
    for r in e.get('revisar', []):
        print(f'   [{(r.get("cuenta") or "?").split("@")[0]}] #{r["uid"]} {r["remitente"]} | {r["asunto"][:50]} | {r["archivo"]}' + (f' | {r["problemas"]}' if r.get('problemas') else ''))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--inicializar', action='store_true')
    ap.add_argument('--enviar', action='store_true')
    ap.add_argument('--estado', action='store_true')
    ap.add_argument('--imprimir', action='store_true', help='imprimir (OP + factura) las pendientes que no se imprimieron')
    ap.add_argument('--agregar-pendiente', metavar='CARPETA_OP')
    ap.add_argument('--unidad')
    ap.add_argument('--apply', action='store_true', help='sin esto todo es DRY-RUN')
    ap.add_argument('--sin-claude', action='store_true', help='no leer facturas con Claude (no gasta tokens)')
    ap.add_argument('--uids', help='(corrida puntual) UIDs de mails a procesar, separados por coma; requiere --cuenta')
    ap.add_argument('--extra-remitentes', help='(con --uids) remitentes adicionales solo para esta corrida, separados por coma')
    ap.add_argument('--max-lecturas', type=int, default=None, help='pisa max_lecturas_por_corrida del config')
    ap.add_argument('--cuenta', help='limitar a una casilla (texto que aparezca en su direccion, ej. administracion)')
    ap.add_argument('--desde-uid', type=int, default=None, help='(pruebas) revisar desde este UID en vez del guardado; requiere --cuenta')
    args = ap.parse_args()
    os.makedirs(WORK, exist_ok=True)
    with open(os.path.join(WORK, '.lock'), 'w') as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            raise SystemExit('Ya hay otra corrida del revisor en curso.')
        if args.inicializar:
            inicializar(args)
        elif args.estado:
            mostrar_estado(args)
        elif args.agregar_pendiente:
            agregar_pendiente(args)
        elif args.enviar:
            enviar(args)
        elif args.imprimir:
            imprimir_cmd(args)
        else:
            revisar_mail(args)


if __name__ == '__main__':
    main()
