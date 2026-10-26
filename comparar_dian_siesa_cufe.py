#!/usr/bin/env python3
"""Tabla unica DIAN vs SIESA por tercero, cruce por CUFE. Uso: python3 comparar_dian_siesa_cufe.py DIAN.xlsx SIESA.xlsx SALIDA.xlsx"""
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

dian, siesa, out = sys.argv[1:4]
D = pd.read_excel(dian, sheet_name=0, dtype=str)
S = pd.read_excel(siesa, dtype=str)
D['Total'] = pd.to_numeric(D.Total); S['Valor'] = pd.to_numeric(S.Valor)
emit = D.Grupo == 'Emitido'                                   # documento soporte emitido: el tercero es el receptor
D['nit'] = D['NIT Emisor'].where(~emit, D['NIT Receptor'])
D['nombre'] = D['Nombre Emisor'].where(~emit, D['Nombre Receptor']).str.strip().str.upper()
D['numero'] = (D.Prefijo.fillna('') + D.Folio.fillna('')).str.strip()
D['valor'] = D.Total.where(~D['Tipo de documento'].str.startswith('Nota'), -D.Total)   # notas credito restan
D['fecha'] = pd.to_datetime(D['Fecha Emisión'], format='%d-%m-%Y')
S = S.rename(columns={'NIT Emisor': 'nit', 'Factura': 'numero', 'Valor': 'siesa', 'CUFE': 'cufe'})
S['nombre_s'] = S.Emisor.str.strip().str.upper(); S['fecha_s'] = pd.to_datetime(S.Fecha)
D = D.rename(columns={'CUFE/CUDE': 'cufe'})

M = D.merge(S[['cufe', 'nit', 'numero', 'siesa', 'nombre_s', 'fecha_s']], on='cufe', how='outer', indicator=True, suffixes=('', '_s'))
M['nit'] = M.nit.fillna(M.nit_s); M['numero'] = M.numero.fillna(M.numero_s)
M['nit'] = M.nit.astype(str); M['numero'] = M.numero.astype(str)
M['nombre'] = M.nombre.fillna(M.nombre_s); M['fecha'] = M.fecha.fillna(M.fecha_s)
M['tipo'] = M['Tipo de documento'].fillna('Factura (solo SIESA)')
M['dian'] = M.valor
M['estado'] = M._merge.astype(str).map({'left_only': 'Solo en DIAN', 'right_only': 'Solo en SIESA', 'both': 'Coincide'})
M.loc[(M._merge == 'both') & ((M.dian - M.siesa).abs() > 0.005), 'estado'] = 'Valor diferente'
M['nombre'] = M.nombre.groupby(M.nit).transform(lambda s: s.mode().iloc[0])   # un solo nombre por NIT


def estado_ter(e):
    s = set(e)
    return 'Coincide' if s == {'Coincide'} else 'Tercero solo en DIAN' if s == {'Solo en DIAN'} else \
        'Tercero solo en SIESA' if s == {'Solo en SIESA'} else 'Con diferencias'


M['estado_ter'] = M.groupby('nit').estado.transform(estado_ter)
M = M.sort_values(['nombre', 'nit', 'fecha', 'numero'])

NUM = '#,##0.00;[Red]-#,##0.00'
fill_p, fill_h = PatternFill('solid', fgColor='D9E1F2'), PatternFill('solid', fgColor='1F3864')
ROJO, AMA = PatternFill('solid', fgColor='F8CBAD'), PatternFill('solid', fgColor='FFE699')
wb = Workbook(); ws = wb.active; ws.title = 'Comparación unificada'
ws.append(['DIAN vs SIESA — una fila por tercero; use +/- (o botones 1/2) para ver los documentos; filtre "Estado" en la fila 4'])
ws.append([None, 'TOTAL GENERAL', None, None, '=SUMIF($D$5:$D$10000,"TOTAL TERCERO",E5:E10000)', '=SUMIF($D$5:$D$10000,"TOTAL TERCERO",F5:F10000)', '=E2-F2'])
ws.append([]); ws.append(['NIT', 'Tercero', 'Documento', 'Tipo de documento', 'Valor DIAN', 'Valor SIESA', 'Diferencia (DIAN − SIESA)', 'Estado', 'Fecha emisión'])
ws['A1'].font = Font(bold=True, size=12, name='Arial')
for c in ws[2]: c.font = Font(bold=True, name='Arial'); c.number_format = NUM
for c in ws[4]: c.font = Font(bold=True, color='FFFFFF', name='Arial'); c.fill = fill_h; c.alignment = Alignment(wrap_text=True)
r = 5
for nit, g in M.groupby('nit', sort=False):
    ws.append([nit, g.nombre.iloc[0], f'{len(g)} documento(s)', 'TOTAL TERCERO',
               f'=SUMIFS(Base!$H:$H,Base!$A:$A,$A{r})', f'=SUMIFS(Base!$I:$I,Base!$A:$A,$A{r})', f'=E{r}-F{r}', g.estado_ter.iloc[0]])
    for c in ws[r]: c.font = Font(bold=True, name='Arial'); c.fill = fill_p
    for _, x in g.iterrows():
        ws.append([nit, x.nombre, x.numero, x.tipo, None if pd.isna(x.dian) else x.dian, None if pd.isna(x.siesa) else x.siesa,
                   f'=N(E{ws.max_row + 1})-N(F{ws.max_row + 1})', x.estado, x.fecha.to_pydatetime()])
    for k in range(r + 1, r + len(g) + 1):
        ws.row_dimensions[k].outlineLevel = 1; ws.row_dimensions[k].hidden = True
        ws[f'I{k}'].number_format = 'dd/mm/yyyy'
        for c in ws[k]: c.font = Font(name='Arial', size=9)
        if ws[f'H{k}'].value != 'Coincide': ws[f'H{k}'].fill = ROJO if 'Solo' in ws[f'H{k}'].value else AMA
    if g.estado_ter.iloc[0] != 'Coincide': ws[f'H{r}'].fill = ROJO if 'solo' in g.estado_ter.iloc[0] else AMA
    r += len(g) + 1
for k in range(2, r):
    for col in 'EFG': ws[f'{col}{k}'].number_format = NUM
ws.sheet_properties.outlinePr.summaryBelow = False
ws.auto_filter.ref = f'A4:I{r-1}'; ws.freeze_panes = 'C5'
for col, w in zip('ABCDEFGHI', [14, 46, 22, 44, 16, 16, 20, 22, 13]): ws.column_dimensions[col].width = w

# Base plana (una fila por documento) lista para Insertar > Tabla dinamica
b = wb.create_sheet('Base')
b.append(['NIT', 'Tercero', 'Documento', 'Tipo de documento', 'Grupo DIAN', 'Fecha emisión', 'Estado documento', 'Valor DIAN', 'Valor SIESA', 'Diferencia', 'Estado tercero'])
for _, x in M.iterrows():
    n = b.max_row + 1
    b.append([x.nit, x.nombre, x.numero, x.tipo, x.Grupo if isinstance(x.Grupo, str) else 'No está en DIAN', x.fecha.to_pydatetime(), x.estado,
              None if pd.isna(x.dian) else x.dian, None if pd.isna(x.siesa) else x.siesa, f'=N(H{n})-N(I{n})', x.estado_ter])
n = b.max_row
for k in range(2, n + 1):
    b[f'F{k}'].number_format = 'dd/mm/yyyy'
    for col in 'HIJ': b[f'{col}{k}'].number_format = NUM
t = Table(displayName='BaseComparacion', ref=f'A1:K{n}'); t.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True); b.add_table(t)
for col, w in zip('ABCDEFGHIJK', [14, 46, 22, 44, 16, 13, 18, 16, 16, 16, 22]): b.column_dimensions[col].width = w
b.freeze_panes = 'A2'
wb.save(out)
print(len(M), 'documentos,', M.nit.nunique(), 'terceros'); print(M.estado.value_counts().to_string()); print(M.drop_duplicates('nit').estado_ter.value_counts().to_string())
