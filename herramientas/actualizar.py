#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Actualiza el tablero Excelencia SSOMA · Salfa Montajes.

Lee : la matriz "Matriz OKR-KPI 2026 - Centros Salfa Montajes.xlsx"
      (pestañas CC-141, CC-146, Chimeneas y Global Salfa Montajes)
Crea: index.html en la raíz del repositorio (la página de GitHub Pages)

Uso:
    python3 herramientas/actualizar.py "ruta/a/Matriz OKR-KPI 2026 - Centros Salfa Montajes.xlsx"
    python3 herramientas/actualizar.py matriz.xlsx --recalcular   # fuerza recálculo con LibreOffice
    python3 herramientas/actualizar.py matriz.xlsx --excluir-mes=11  # no publica noviembre (varios: --excluir-mes=10,11)

Requiere: openpyxl (pip install openpyxl). LibreOffice solo si se usa --recalcular
o si la matriz llega sin valores calculados (por ejemplo, generada por un script).
"""
import datetime
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
PLANTILLA = os.path.join(AQUI, 'plantilla.html')
SALIDA = os.path.join(RAIZ, 'index.html')

# Pestaña de la matriz -> id y nombre en el tablero
CENTROS = [('global', 'Global Salfa Montajes', 'Salfa Montajes'),
           ('cc141', 'CC-141', 'CC-141'),
           ('cc146', 'CC-146', 'CC-146'),
           ('chim', 'Chimeneas', 'Chimeneas')]

# id, foco, nombre, col Meta, col Ponderación, col Real (o (numerador, denominador)), col Resultado ponderado, tipo
INDICADORES = [
    ('rep_total', 'AT', 'Reportabilidad total', 'G', 'H', 'J', 'K', 'pct'),
    ('hap', 'AT', 'Hallazgos de alto potencial (HAP)', 'L', 'M', 'O', 'P', 'pct'),
    ('txv', 'AT', 'Tarjetas por la Vida (TxV)', 'Q', 'R', 'T', 'U', 'pct'),
    ('lmg_c', 'AT', 'Actividades críticas LMG Central', 'V', 'W', 'Z', 'AA', 'pct'),
    ('lmg_o', 'AT', 'Actividades críticas LMG Operativa', 'AB', 'AC', 'AF', 'AG', 'pct'),
    ('adm_c', 'GP', 'Adm. del riesgo LMG Central', 'AI', 'AJ', 'AM', 'AN', 'pct'),
    ('adm_o', 'GP', 'Adm. del riesgo LMG Operativa', 'AO', 'AP', 'AS', 'AT', 'pct'),
    ('ad_paes', 'GP', 'Adhesión PAES', 'AU', 'AV', 'AW', 'AX', 'pct'),
    ('cu_paes', 'GP', 'Cumplimiento PAES', 'AY', 'AZ', 'BA', 'BB', 'pct'),
    ('cc_paes', 'GP', 'Controles críticos PAES', 'BC', 'BD', 'BE', 'BF', 'pct'),
    ('ad_cphs', 'GP', 'Adhesión CPHS al PAES', 'BG', 'BH', ('BJ', 'BI'), 'BL', 'pct'),
    ('cu_cphs', 'GP', 'Cumplimiento CPHS al PAES', 'BM', 'BN', 'BO', 'BP', 'pct'),
    ('gesco_ma', 'GP', 'GESCO Medio Ambiente', 'BQ', 'BR', 'BS', 'BT', 'pct'),
    ('gesco_so', 'GP', 'GESCO Salud Ocupacional', 'BU', 'BV', 'BW', 'BX', 'pct'),
    ('gesco_seg', 'GP', 'GESCO Seguridad', 'BY', 'BZ', 'CA', 'CB', 'pct'),
    ('ad_sim', 'GP', 'Adhesión simulacro documental', 'CC', 'CD', 'CE', 'CF', 'pct'),
    ('cu_sim', 'GP', 'Cumplimiento simulacro documental', 'CG', 'CH', 'CI', 'CJ', 'pct'),
    ('alc', 'GP', 'Plan de alcohol y drogas', 'CK', 'CL', 'CM', 'CN', 'pct'),
    ('cvs', 'GP', 'Campaña de verano seguro', 'CO', 'CP', 'CQ', 'CR', 'pct'),
    ('s360', 'GP', 'Usabilidad SSOMA 360', 'CS', 'CT', 'CU', 'CV', 'pct'),
    ('capac', 'AC', 'Cumplimiento plan de capacitación', 'CX', 'CY', 'DB', 'DC', 'pct'),
    ('iap_sai', 'AC', 'IAP cerrados en SAI', 'DD', 'DE', ('DG', 'DF'), 'DH', 'pct'),
    ('iap_ces', 'AC', 'IAP con aprendizajes en CES / GT', 'DI', 'DJ', ('DL', 'DK'), 'DM', 'pct'),
    ('precur', 'AC', 'Repetición de precursores', 'DN', 'DO', 'DP', 'DQ', 'int'),
    ('cultura', 'EG', 'Cultura de seguridad', 'DS', 'DT', 'DU', 'DV', 'int'),
    ('cert_cphs', 'EG', 'Certificación CPHS oro', 'DW', 'DX', 'EA', 'EB', 'pct'),
    ('iso', 'EG', 'Certificación SGI ISO 14.001 / 45.001', 'EC', 'ED', 'EE', 'EF', 'pct'),
    ('pec', 'EG', 'Certificación PEC de excelencia', 'EG', 'EH', 'EI', 'EJ', 'pct'),
]
# Columna del total de cada foco y de la nota mensual
FOCOS = {'AT': 'AH', 'GP': 'CW', 'AC': 'DR', 'EG': 'EK'}
NOTA = 'EL'
FILA_ENERO, FILA_DIC = 5, 16   # una fila por mes; columna A = 1 si el mes tiene datos


def num(x):
    return float(x) if isinstance(x, (int, float)) and not isinstance(x, bool) else None


def r4(v):
    return None if v is None else round(v, 4)


def recalcular(origen):
    """Recalcula todas las fórmulas con LibreOffice y devuelve la ruta del archivo recalculado."""
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        sys.exit('No encontré LibreOffice (soffice). Instálalo o abre y guarda la matriz en Excel antes de correr el script.')
    tmp = tempfile.mkdtemp()
    subprocess.run([soffice, '--headless', '--convert-to', 'xlsx', '--outdir', tmp, origen],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(tmp, os.path.basename(origen))


def leer(ruta, excluir=()):
    vals = openpyxl.load_workbook(ruta, data_only=True)
    form = openpyxl.load_workbook(ruta)
    faltan = [n for _, n, _ in CENTROS if n not in vals.sheetnames]
    if faltan:
        sys.exit('A la matriz le faltan estas pestañas: ' + ', '.join(faltan))
    ref = form['CC-141']
    datos = {'generado': datetime.datetime.now().strftime('%d-%m-%Y %H:%M'),
             'kpis': [{'id': k[0], 'foco': k[1], 'nombre': k[2], 'tipo': k[7]} for k in INDICADORES],
             'centros': []}
    for cid, hoja, nombre in CENTROS:
        ws, wf = vals[hoja], form[hoja]
        centro = {'id': cid, 'nombre': nombre, 'meses': []}
        for r in range(FILA_ENERO, FILA_DIC + 1):
            con = num(ws[f'A{r}'].value) or 0
            if not con:
                continue
            etiqueta = str(ws[f'C{r}'].value)
            n = int(etiqueta.split('.')[0])
            if n in excluir:
                continue
            # qué indicadores suma cada foco ese mes (según la fórmula de la pestaña CC-141)
            incluidos = set()
            for col in FOCOS.values():
                incluidos |= set(re.findall(r'([A-Z]{1,3})' + str(r) + r'\b', str(ref[f'{col}{r}'].value)))
            ks = []
            for kid, foco, nom, mc, pc, rc, resc, tipo in INDICADORES:
                if isinstance(rc, tuple):
                    nu, de = num(ws[f'{rc[0]}{r}'].value), num(ws[f'{rc[1]}{r}'].value)
                    real = nu / de if (nu is not None and de) else None
                else:
                    real = num(ws[f'{rc}{r}'].value)
                ks.append([r4(num(ws[f'{mc}{r}'].value)), r4(num(ws[f'{pc}{r}'].value)), r4(real),
                           r4(num(ws[f'{resc}{r}'].value)), 1 if resc in incluidos else 0])
            nota = num(ws[f'{NOTA}{r}'].value)
            if nota is None:
                sys.exit(f'La pestaña {hoja}, mes {etiqueta}, no tiene la nota calculada. '
                         'Vuelve a correr con --recalcular.')
            centro['meses'].append({
                'n': n, 'mes': etiqueta.split('. ', 1)[1],
                'nc': int(con) if cid == 'global' else None,
                'dot': num(ws[f'D{r}'].value), 'lc': num(ws[f'E{r}'].value), 'lo': num(ws[f'F{r}'].value),
                'exc': r4(nota), 'f': {f: r4(num(ws[f'{c}{r}'].value)) for f, c in FOCOS.items()}, 'k': ks})
        datos['centros'].append(centro)
    return datos


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    excluir = set()
    for a in sys.argv[1:]:
        if a.startswith('--excluir-mes='):
            excluir |= {int(x) for x in a.split('=', 1)[1].split(',') if x.strip()}
    if len(args) != 1:
        sys.exit(__doc__)
    ruta = args[0]
    if '--recalcular' in sys.argv:
        ruta = recalcular(ruta)
    datos = leer(ruta, excluir)
    if not any(m['exc'] is not None for m in datos['centros'][1]['meses']) and '--recalcular' not in sys.argv:
        datos = leer(recalcular(ruta), excluir)
    plantilla = open(PLANTILLA, encoding='utf-8').read()
    js = json.dumps(datos, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    open(SALIDA, 'w', encoding='utf-8').write(plantilla.replace('__DATA__', js))
    print(f'index.html actualizado ({datos["generado"]})')
    if excluir:
        print('  Meses excluidos (no se publican): ' + ', '.join(str(m) for m in sorted(excluir)))
    for c in datos['centros']:
        if c['meses']:
            meses = ', '.join(f"{m['mes'][:3]} {m['exc'] * 100:.1f}%" for m in c['meses'])
            anio = sum(m['exc'] for m in c['meses']) / len(c['meses'])
            print(f"  {c['nombre']:<15} año {anio * 100:.1f}%  ·  {meses}")
        else:
            print(f"  {c['nombre']:<15} sin datos")


if __name__ == '__main__':
    main()
