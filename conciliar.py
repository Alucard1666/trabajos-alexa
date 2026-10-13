#!/usr/bin/env python3
"""Cruce DIAN vs Journal (auxiliar contable) por tercero y documento.

Uso: python3 conciliar.py DIAN.xlsx JOURNAL.xls SALIDA.xlsx
Solo LEE los dos archivos de entrada (no los modifica). Dependencias: pandas, openpyxl, xlrd.
"""
import re, sys, collections
import pandas as pd, xlrd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

TOL = 0.005
PAYABLE = ('2205', '2210', '2335')          # proveedores / costos y gastos por pagar
RETEN = ('2365', '2367', '2368')            # retefuente, reteIVA, reteICA
PAGOS = {'CE', 'NB'}                         # comprobantes de pago (no son causación)
NA, NF = 'No disponible', 'No encontrado'


# ---------------------------------------------------------------- lectura
def leer_journal(path):
    s = xlrd.open_workbook(path, logfile=open('/dev/null', 'w')).sheet_by_index(0)
    rows, acct, aname, nit, tname = [], None, None, None, None
    for i in range(s.nrows):
        r = s.row_values(i)
        a = r[0]
        if isinstance(a, float):
            rows.append(dict(fila=i + 1, fecha=a, doc=r[1], desc=str(r[4]), ref=str(r[7]), deb=r[9] or 0.0,
                             cre=r[10] or 0.0, cuenta=acct, cuenta_nom=aname, nit=nit, tercero=tname))
            continue
        t = str(a)
        m = re.match(r'^\s?(\d{5,12})\s{3,}(\S.*)$', t)
        if m and not re.match(r'^\d+$', t.strip()):
            nit, tname = m.group(1), m.group(2).strip()
        elif re.match(r'^\d{6,8}$', t.strip()):
            acct, aname, nit, tname = t.strip(), str(r[1]).strip(), None, None
    J = pd.DataFrame(rows)
    J['dtype'] = J.doc.str.extract(r'^\d+-(\w+)-')[0]
    J['refn'] = J.ref.map(lambda x: int(m.group(2)) if (m := re.match(r'^(.*?)-(\d{8})-(\d\d)$', x)) else None)
    J['N'] = J.desc.str.upper().str.replace(r'[^A-Z0-9]', '', regex=True)
    return J, s.nrows


def leer_dian(path):
    D = pd.read_excel(path, dtype=str)
    D['Total'] = pd.to_numeric(D['Total'])
    D['IVA'] = pd.to_numeric(D['IVA'])
    emit = D.Grupo == 'Emitido'          # documento soporte emitido por la empresa: el tercero es el receptor
    D['nit'] = D['NIT Emisor'].where(~emit, D['NIT Receptor'])
    D['tercero'] = D['Nombre Emisor'].where(~emit, D['Nombre Receptor'])
    D['numero'] = (D.Prefijo.fillna('') + D.Folio.fillna('')).str.strip()
    D['es_nota'] = D['Tipo de documento'].str.startswith('Nota')
    return D


# ---------------------------------------------------------------- cruce
def concepto(c):
    if c.startswith(('2408', '511570')): return 'IVA'
    if c.startswith(('5115', '2404', '2412')): return 'Otros impuestos'
    if c[0] in '567': return 'Gastos'
    if c[:2] in ('14', '15', '16', '17'): return 'Inventarios/activos'
    return 'Otros conceptos'


def unidad(J, doc, nit, signo):
    """Totaliza el comprobante `doc` para el tercero `nit`. signo=+1 causación, -1 nota crédito."""
    L = J[(J.doc == doc) & (J.nit == nit)]
    cr = L.cre - L.deb
    pay = L.cuenta.str[:4].isin(PAYABLE); ret = L.cuenta.str[:4].isin(RETEN)
    lado_prov = signo * cr[pay | ret].sum()
    ret_v = signo * cr[ret].sum()
    otros = L[~(pay | ret)]
    con = collections.OrderedDict((k, 0.0) for k in ['IVA', 'Otros impuestos', 'Gastos', 'Inventarios/activos', 'Otros conceptos'])
    for c, d in zip(otros.cuenta, otros.deb - otros.cre): con[concepto(c)] += signo * d
    lado_deb = sum(con.values())
    cuentas = '; '.join(sorted({f'{c} {n}' for c, n in zip(L.cuenta, L.cuenta_nom)}))
    return dict(doc=doc, lado_prov=round(lado_prov, 2), lado_deb=round(lado_deb, 2), ret=round(ret_v, 2),
                con={k: round(v, 2) for k, v in con.items()}, completo=abs(lado_deb - lado_prov) < 0.5,
                cuentas=cuentas, filas=', '.join(map(str, L.fila)), anul=bool(L.desc.str.contains('ANUL', case=False).any()))


def unidades_causacion(J):
    """(doc, nit, signo) de todos los comprobantes que mueven proveedores, excluyendo pagos."""
    P = J[J.cuenta.str[:4].isin(PAYABLE) & J.nit.notna()]
    out = {}
    for (doc, nit), g in P.groupby(['doc', 'nit']):
        neto = (g.cre - g.deb).sum()
        if neto > 0.5: out[(doc, nit)] = 1
        elif neto < -0.5 and g.dtype.iloc[0] not in PAGOS: out[(doc, nit)] = -1
    return out


def lineas_candidatas(Ln, r):
    f = re.sub(r'\D', '', str(r.Folio)).lstrip('0'); p = re.sub(r'[^A-Z0-9]', '', str(r.Prefijo).upper()) if pd.notna(r.Prefijo) else ''
    full = re.sub(r'[^A-Z0-9]', '', (str(r.Prefijo) if pd.notna(r.Prefijo) else '') + str(r.Folio).upper())
    if not f: return Ln.iloc[0:0], ''
    def ref_ok(n):
        if n is None or n != n: return False
        n = int(n); sn = str(n)
        return sn == f or (len(f) >= 4 and n >= 1000 and (sn.endswith(f) or f.endswith(sn)))
    m_ref = Ln.refn.map(ref_ok).astype(bool)
    pat_f = re.compile(r'(?<!\d)' + f + r'(?!\d)')
    m_txt = Ln.N.map(lambda t: bool((len(f) >= 4 and pat_f.search(t)) or (len(full) >= 5 and full in t)))
    m_doc = (Ln.doc.str.extract(r'-(\d+)$')[0].astype(float) == float(f)) & (Ln.dtype == 'DSE') if r.Grupo == 'Emitido' else pd.Series(False, index=Ln.index)
    hit = Ln[m_ref.astype(bool) | m_txt.astype(bool) | m_doc.astype(bool)]
    met = 'referencia' if m_ref[hit.index].any() else ('número del comprobante' if m_doc[hit.index].any() else 'texto/descripción')
    return hit, met


def cruzar(D, J):
    unidades = unidades_causacion(J)
    por_nit = collections.defaultdict(list)
    for (doc, nit), s in unidades.items(): por_nit[nit].append((doc, s))
    res, usados = [], collections.defaultdict(list)
    for i, r in D.iterrows():
        signo = -1 if r.es_nota else 1
        Ln = J[J.nit == r.nit]
        hit, met = lineas_candidatas(Ln, r)
        docs_hit = set(hit.doc)
        caus = [d for d, s in por_nit.get(r.nit, []) if d in docs_hit and s == signo]
        otros = sorted(docs_hit - set(caus))
        met_final = met
        if not caus and not docs_hit:   # sin número verificable: intentar NIT + valor exacto
            cand = [(d, unidad(J, d, r.nit, signo)) for d, s in por_nit.get(r.nit, []) if s == signo]
            cand = [(d, u) for d, u in cand if abs(u['lado_prov'] - r.Total) < TOL]
            if len(cand) >= 1: res.append(dict(i=i, tipo='valor', caus=[d for d, _ in cand], otros=[], met='NIT + valor (sin folio)')); continue
        res.append(dict(i=i, tipo='folio' if caus else ('otros' if docs_hit else 'nada'), caus=caus, otros=otros, met=met_final,
                        nlin=len(Ln)))
        for d in caus: usados[(d, r.nit)].append(i)
    return res, usados, unidades


def clasificar(D, J, res, usados, unidades):
    filas = []
    dup = {k: v for k, v in usados.items() if len(v) > 1}
    for x in res:
        r = D.loc[x['i']]; signo = -1 if r.es_nota else 1
        base = dict(tercero=r.tercero, nit=r.nit, numero=r.numero, dian=r.Total,
                    est_dian=f"{r.Estado} ({r['Tipo de documento']}, {r.Grupo})", origen='DIAN', iva_dian=r.IVA)
        k = x['tipo']
        if k in ('folio', 'valor'):
            us = [unidad(J, d, r.nit, signo) for d in x['caus']]
            tot_prov = round(sum(u['lado_prov'] for u in us), 2)
            tot_deb = round(sum(u['lado_deb'] for u in us), 2)
            con = {c: round(sum(u['con'][c] for u in us), 2) for c in us[0]['con']}
            ret = round(sum(u['ret'] for u in us), 2)
            comp = all(u['completo'] for u in us)
            base.update(comprobantes=', '.join(x['caus']), filas=', '.join(u['filas'] for u in us), metodo=x['met'],
                        cuentas='; '.join(u['cuentas'] for u in us), est_j='Presente en el auxiliar (el archivo no trae columna de estado; se infiere Contabilizado)',
                        conceptos=con, ret=ret)
            iva_j = con['IVA']
            obs = []
            if any(u['anul'] for u in us): obs.append('La descripción del comprobante menciona ANULACIÓN: confirmar que no esté anulado.')
            if abs((r.IVA or 0) - iva_j) > 1 and comp: obs.append(f'IVA DIAN {r.IVA:,.2f} vs IVA Journal {iva_j:,.2f}.')
            if any((d, r.nit) in dup for d in x['caus']):
                otros = sorted({DIAN for d in x['caus'] for DIAN in dup.get((d, r.nit), []) if DIAN != x['i']})
                obs.append('El mismo comprobante aparece cruzado con otros documentos DIAN: ' + ', '.join(D.loc[o, 'numero'] for o in otros) + '.')
                base.update(cls='INFORMACIÓN INSUFICIENTE', jv=tot_prov, motivo='Un mismo comprobante del Journal coincide con varios documentos de la DIAN; no se puede agrupar sin duplicar valores.')
            elif k == 'valor':
                base.update(cls='INFORMACIÓN INSUFICIENTE', jv=tot_prov, motivo='Coincide el NIT y el valor exacto, pero ninguna línea del Journal cita el número del documento; no se puede identificar con certeza.')
            elif comp:
                if abs(r.Total - tot_deb) < TOL:
                    base.update(cls='VERIFICADO', jv=tot_deb, motivo='Coincide exacto (débitos = créditos = DIAN).')
                else:
                    base.update(cls='PENDIENTE', jv=tot_deb, motivo='La suma de los conceptos contables no coincide con el total de la DIAN.')
            else:
                if abs(r.Total - tot_prov) < TOL:
                    base.update(cls='INFORMACIÓN INSUFICIENTE', jv=tot_prov,
                                motivo='El lado proveedor + retenciones coincide exacto con la DIAN, pero las líneas de gasto/IVA no están completas en el Journal exportado (el archivo se corta en la fila 16.383); no se puede confirmar la suma de conceptos.')
                else:
                    base.update(cls='PENDIENTE', jv=tot_prov, motivo='El lado proveedor + retenciones del Journal no coincide con el total de la DIAN.')
            if len(us) > 1: obs.append(f'Contabilizado en {len(us)} comprobantes del Journal ({base["comprobantes"]}); se sumaron todos.')
            base['obs'] = ' '.join(obs) or '—'
            base['lado_completo'] = 'Sí' if comp else 'No (gasto/IVA cortado en la exportación)'
        else:
            base.update(comprobantes=', '.join(x['otros']) or NF, filas='', metodo=x['met'] if x['otros'] else NF, cuentas=NA, conceptos=None, ret=None, jv=None,
                        lado_completo=NA)
            if k == 'otros':
                base.update(cls='PENDIENTE', est_j='Solo comprobantes de pago/legalización, sin causación',
                            motivo='El documento se cita en comprobantes que no lo causan (pago, legalización o anticipo), pero no hay un comprobante de causación para él.',
                            obs='Revisar: ' + ', '.join(x['otros']))
            elif x['nlin'] > 0:
                base.update(cls='PENDIENTE', est_j=NF, motivo='No se encontró el documento en el Journal (el tercero sí tiene otros movimientos).', obs='El tercero aparece en el Journal pero ninguna línea cita este folio ni coincide su valor.')
            else:
                base.update(cls='PENDIENTE', est_j=NF, motivo='El tercero no aparece en el Journal.', obs='NIT sin movimientos en el archivo del Journal.')
        filas.append(base)
    # Journal sin DIAN
    for (doc, nit), s in unidades.items():
        if (doc, nit) in usados or s != 1: continue   # las notas/legalizaciones que reducen proveedor no son documentos DIAN
        u = unidad(J, doc, nit, s)
        t = J[(J.doc == doc) & (J.nit == nit)].tercero.iloc[0]
        d = J[(J.doc == doc) & (J.nit == nit)].desc.iloc[0].strip()
        filas.append(dict(tercero=t, nit=nit, numero=NF, dian=None, est_dian=NF, origen='Journal', iva_dian=None, comprobantes=doc, filas=u['filas'],
                          metodo=NA, cuentas=u['cuentas'], est_j='Presente en el auxiliar (el archivo no trae columna de estado; se infiere Contabilizado)',
                          conceptos=u['con'], ret=u['ret'], jv=u['lado_deb'] if u['completo'] else u['lado_prov'],
                          lado_completo='Sí' if u['completo'] else 'No (gasto/IVA cortado en la exportación)', cls='PENDIENTE',
                          motivo=('Ajuste por diferencia en cambio (no es documento DIAN); validar manualmente.' if re.search('DIFERENCIA EN (TASA|CAMBIO)', d.upper()) else 'Registrado en el Journal pero no se encontró el documento en la DIAN.'), obs=f'Descripción: {d[:160]}'))
    return pd.DataFrame(filas)


# ---------------------------------------------------------------- salida
HDR = ['Nombre del tercero', 'NIT o identificación', 'Número del documento', 'Cuenta contable y auxiliar', 'Valor total según la DIAN',
       'Valor total sumado en el Journal', 'Diferencia (DIAN − Journal)', 'Desglose de conceptos contables', 'Estado en la DIAN',
       'Estado en el Journal', 'Motivo por el cual queda pendiente', 'Observaciones para revisión manual']
EXTRA = ['Clasificación', 'Origen', 'Comprobante(s) Journal', 'Fila(s) en el Journal', 'Método de cruce', 'Lado gasto/IVA completo en el Journal']


def desglose(r):
    if r.conceptos is None: return NA
    c = r.conceptos; t = [f'{k}: {v:,.2f}' for k, v in c.items() if abs(v) > 0.004]
    t.append(f'Retenciones (crédito): {r.ret:,.2f}')
    s = ' | '.join(t) if t else 'Sin líneas de gasto/IVA en el Journal'
    if r.lado_completo != 'Sí': s += ' | Gasto/otros: No disponible (líneas fuera del archivo exportado)'
    return s


def hoja(wb, nombre, df, ancho=None):
    ws = wb.create_sheet(nombre)
    ws.append(HDR + EXTRA)
    for _, r in df.iterrows():
        n = ws.max_row + 1
        dif = f'=IF(AND(ISNUMBER(E{n}),ISNUMBER(F{n})),E{n}-F{n},"{NA}")'
        ws.append([r.tercero, r.nit, r.numero, r.cuentas, r.dian if pd.notna(r.dian) else NF, r.jv if pd.notna(r.jv) else NF, dif, desglose(r),
                   r.est_dian, r.est_j, r.motivo, r.obs, r.cls, r.origen, r.comprobantes, r.filas, r.metodo, r.lado_completo])
    for c in ws[1]: c.font = Font(bold=True, color='FFFFFF', name='Arial'); c.fill = PatternFill('solid', fgColor='1F3864'); c.alignment = Alignment(wrap_text=True, vertical='center')
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name='Arial', size=9); c.alignment = Alignment(wrap_text=True, vertical='top')
        for c in row[4:7]: c.number_format = '#,##0.00;[Red]-#,##0.00;0.00'
    for i, w in enumerate([30, 14, 16, 34, 16, 16, 16, 48, 28, 30, 48, 48, 22, 10, 22, 14, 18, 22], 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'D2'; ws.auto_filter.ref = ws.dimensions
    return ws


def escribir(F, J_rows, out, info):
    wb = Workbook(); wb.remove(wb.active)
    ws = wb.create_sheet('Resumen')
    hoja(wb, 'Conciliación', F)
    pend = F[F.cls != 'VERIFICADO']
    hoja(wb, 'Pendientes de revisión manual', pend)
    nC, nP = len(F) + 1, len(pend) + 1
    cl = f"'Conciliación'!$M$2:$M${nC}"; dd = f"'Conciliación'!$G$2:$G${nC}"; org = f"'Conciliación'!$N$2:$N${nC}"
    R = [('Resumen de la conciliación DIAN vs Journal — septiembre 2026', None), (None, None),
         ('Documentos revisados (DIAN + Journal sin DIAN)', f'=COUNTA({cl})'),
         ('   de los cuales presentes en la DIAN', f'=COUNTIF({org},"DIAN")'),
         ('   de los cuales solo en el Journal', f'=COUNTIF({org},"Journal")'),
         ('VERIFICADOS', f'=COUNTIF({cl},"VERIFICADO")'),
         ('PENDIENTES', f'=COUNTIF({cl},"PENDIENTE")'),
         ('INFORMACIÓN INSUFICIENTE', f'=COUNTIF({cl},"INFORMACIÓN INSUFICIENTE")'),
         ('Documentos con diferencia de valor (DIAN vs Journal)', f'=COUNTIF({dd},">0")+COUNTIF({dd},"<0")'),
         ('Valor acumulado de diferencias (absoluto)', f'=SUMIF({dd},">0")-SUMIF({dd},"<0")'),
         ('Valor neto de diferencias (DIAN − Journal)', f'=SUM({dd})'),
         ('Pendientes + insuficientes enviados a revisión manual', f"=COUNTA('Pendientes de revisión manual'!$M$2:$M${nP})"),
         (None, None),
         ('Notas de método y límites', None)] + [(t, None) for t in info]
    for a, b in R:
        ws.append([a, b])
    ws['A1'].font = Font(bold=True, size=13, name='Arial'); ws['A14'].font = Font(bold=True, name='Arial')
    for r in range(3, 13): ws[f'B{r}'].number_format = '#,##0.00' if r >= 10 else '#,##0'
    ws.column_dimensions['A'].width = 120; ws.column_dimensions['B'].width = 20
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.font.name != 'Arial' or not c.font.bold: c.font = Font(name='Arial', size=10, bold=c.font.bold)
        row[0].alignment = Alignment(wrap_text=True, vertical='top')
    wb.move_sheet('Resumen', offset=-2)
    wb.save(out)


def main(dian, jour, out):
    D = leer_dian(dian); J, nrows = leer_journal(jour)
    res, usados, unidades = cruzar(D, J)
    F = clasificar(D, J, res, usados, unidades)
    F['cls'] = F.cls
    F = F.sort_values(['origen', 'tercero', 'numero'], key=lambda c: c.astype(str), kind='stable').reset_index(drop=True)
    ultima = J.cuenta.iloc[-1]
    info = ['1. Se usaron únicamente los dos archivos (DIAN y Journal); los originales no se modificaron.',
            '2. Documento DIAN = Prefijo+Folio; el tercero es el Emisor (recibidos) o el Receptor (documentos soporte emitidos). Se cruza por NIT del tercero y folio citado en la referencia/descripción de las líneas del Journal.',
            '3. Valor en el Journal = suma de las líneas del comprobante de causación del tercero (IVA + otros impuestos + gastos + inventarios/activos + otros). Verificación cruzada: lado proveedor (2205/2210/2335) + retenciones (2365/2367/2368) debe ser igual a esa suma.',
            '4. ESTADO "Contabilizado": el Journal es un auxiliar contable exportado y NO trae columna de estado. Se asumió que un movimiento en el auxiliar está contabilizado; si existe otro criterio de estado, validar.',
            f'5. LÍMITE DEL ARCHIVO: el Journal tiene {nrows:,} filas y termina en la cuenta {ultima} (clase 5 incompleta; faltan cuentas de gasto posteriores, clases 6 y 7). Por eso muchas facturas no tienen completas sus líneas de gasto: si el lado proveedor + retenciones coincide con la DIAN quedan como INFORMACIÓN INSUFICIENTE, no VERIFICADO. Volver a exportar el Journal completo (o por rangos de cuenta) y ejecutar de nuevo conciliar.py.',
            '6. Los registros "solo en el Journal" se limitan a comprobantes que abonan (causan) proveedores/acreedores (2205, 2210, 2335); no se listan pagos, legalizaciones ni notas que solo reducen la cuenta por pagar, ni pagos directos de banco/caja sin causación.',
            '7. Las diferencias son exactas (sin tolerancia); no se ajustó ningún valor.']
    escribir(F, J, out, info)
    print(F.cls.value_counts().to_string()); print(F.groupby(['origen', 'cls']).size().to_string())


if __name__ == '__main__':
    main(*sys.argv[1:4])
