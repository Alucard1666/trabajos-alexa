#!/usr/bin/env python3
"""Cuadro por tercero (NIT) con detalle expandible. Uso: python3 cuadro_terceros.py DIAN.xlsx SALIDA.xlsx [CONCILIACION.xlsx]"""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

dian, out = sys.argv[1], sys.argv[2]
conc = sys.argv[3] if len(sys.argv) > 3 else None

D = pd.read_excel(dian, dtype=str)
for c in ('Total', 'IVA'): D[c] = pd.to_numeric(D[c])
emit = D.Grupo == 'Emitido'                                   # documento soporte emitido: el tercero es el receptor
D['nit'] = D['NIT Emisor'].where(~emit, D['NIT Receptor'])
nom = D['Nombre Emisor'].where(~emit, D['Nombre Receptor'])
D['numero'] = (D.Prefijo.fillna('') + D.Folio.fillna('')).str.strip()
D['nota'] = D['Tipo de documento'].str.startswith('Nota')
D['valor'] = D.Total.where(~D.nota, -D.Total)                 # notas crédito restan
D['fecha'] = pd.to_datetime(D['Fecha Emisión'], format='%d-%m-%Y')
D['nombre'] = nom.groupby(D.nit).transform(lambda s: s.str.strip().str.upper().mode().iloc[0])  # un solo nombre por NIT
D['clasif'] = 'No disponible'
if conc:
    C = pd.read_excel(conc, sheet_name='Conciliación', dtype={'NIT o identificación': str})
    C = C[C.Origen == 'DIAN'].set_index(['NIT o identificación', 'Número del documento'])['Clasificación']
    D['clasif'] = [C.get((n, d), 'No disponible') for n, d in zip(D.nit, D.numero)]

wb = Workbook(); ws = wb.active; ws.title = 'Cuadro por tercero'
H = ['NIT', 'Tercero', 'Documento', 'Tipo de documento', 'Fecha emisión', 'IVA', 'Valor (neto)', 'Clasificación conciliación']
ws.append(['Cuadro por tercero — DIAN septiembre 2026 (use el +/- de la izquierda o los botones 1/2 para expandir; filtre por NIT en la fila 4)'])
ws.append(['Las notas crédito restan. TOTAL GENERAL ='])
ws.append([]); ws.append(H)
ws['B2'] = 'TOTAL GENERAL (suma de todos los terceros)'; ws['G2'] = '=SUMIF(D5:D10000,"TOTAL TERCERO",G5:G10000)'
fill_p, fill_h = PatternFill('solid', fgColor='D9E1F2'), PatternFill('solid', fgColor='1F3864')
for c in ws[4]: c.font = Font(bold=True, color='FFFFFF', name='Arial'); c.fill = fill_h; c.alignment = Alignment(wrap_text=True)
ws['A1'].font = Font(bold=True, size=12, name='Arial'); ws['A2'].value = None
ws['B2'].font = ws['G2'].font = Font(bold=True, name='Arial')
r = 5
for (nit, nombre), g in sorted(D.groupby(['nit', 'nombre']), key=lambda x: x[0][1]):
    g = g.sort_values('fecha')
    ws.append([nit, nombre, f'{len(g)} documento(s)', 'TOTAL TERCERO', None, f'=SUMIFS(Base!$G:$G,Base!$A:$A,$A{r})', f'=SUMIFS(Base!$I:$I,Base!$A:$A,$A{r})'])
    for c in ws[r]: c.font = Font(bold=True, name='Arial'); c.fill = fill_p
    for _, x in g.iterrows():
        ws.append([nit, nombre, x.numero, x['Tipo de documento'], x.fecha.to_pydatetime(), x.IVA, x.valor, x.clasif])
    for k in range(r + 1, r + len(g) + 1):
        ws.row_dimensions[k].outlineLevel = 1; ws.row_dimensions[k].hidden = True
        ws[f'E{k}'].number_format = 'dd/mm/yyyy'
        for col in 'FG': ws[f'{col}{k}'].number_format = '#,##0.00;[Red]-#,##0.00'
        for c in ws[k]: c.font = Font(name='Arial', size=9)
    r += len(g) + 1
for k in range(5, r):
    for col in 'FG': ws[f'{col}{k}'].number_format = '#,##0.00;[Red]-#,##0.00'
ws['G2'].number_format = '#,##0.00;[Red]-#,##0.00'
ws.sheet_properties.outlinePr.summaryBelow = False
ws.auto_filter.ref = f'A4:H{r-1}'; ws.freeze_panes = 'C5'
for col, w in zip('ABCDEFGH', [14, 46, 22, 44, 13, 16, 18, 26]): ws.column_dimensions[col].width = w

# Base plana: tabla de Excel lista para Insertar > Tabla dinámica
b = wb.create_sheet('Base')
b.append(['NIT', 'Tercero', 'Documento', 'Tipo de documento', 'Grupo', 'Fecha emisión', 'IVA', 'Total DIAN', 'Valor (neto)', 'Estado DIAN', 'Clasificación conciliación'])
for _, x in D.sort_values(['nombre', 'fecha']).iterrows():
    b.append([x.nit, x.nombre, x.numero, x['Tipo de documento'], x.Grupo, x.fecha.to_pydatetime(), x.IVA, x.Total, x.valor, x.Estado, x.clasif])
n = b.max_row
for k in range(2, n + 1):
    b[f'F{k}'].number_format = 'dd/mm/yyyy'
    for col in 'GHI': b[f'{col}{k}'].number_format = '#,##0.00;[Red]-#,##0.00'
t = Table(displayName='BaseDIAN', ref=f'A1:K{n}'); t.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True); b.add_table(t)
for col, w in zip('ABCDEFGHIJK', [14, 46, 22, 44, 11, 13, 14, 16, 16, 26, 26]): b.column_dimensions[col].width = w
b.freeze_panes = 'A2'
wb.save(out)
print(len(D), 'documentos,', D.nit.nunique(), 'terceros')
