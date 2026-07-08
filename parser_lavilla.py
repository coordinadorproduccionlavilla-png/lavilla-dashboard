"""
PARSER DE MERMAS - CARNES FRÍAS LA VILLA
=========================================
Autor: desarrollado en sesiones Claude + Javier Arévalo
Última actualización: 02-Jul-2026
Versión: 1.0

DESCRIPCIÓN:
    Lee PROGRAMACION 2026.xlsx desde Google Drive y genera los records
    JSON para el dashboard de control de mermas publicado en GitHub Pages.

FUENTES DE DATOS:
    - PROGRAMACION 2026.xlsx (Drive ID: 1HN_C9tPFb94tMIF3Krjr2nG2XfxCSLzM)
        * Hoja "Entrega PT": registros de empaque y peso báscula
        * Hoja "Entrega PDN": registros de producción
        * Hoja "Parametros control": estándares por familia y sede
    - Dashboard actual: https://raw.githubusercontent.com/coordinadorproduccionlavilla-png/lavilla-dashboard/main/index.html

REGLAS DE NEGOCIO CRÍTICAS (no cambiar sin consultar a Javi):
    1. COLUMNA K (col[10] de Entrega PT): peso del lote terminado, pesado en
       báscula industrial antes de entrar a cava. Disponible SOLO desde el
       7-abril-2026 (fecha de compra de la báscula). Antes de esa fecha,
       los datos no son comparables con esta metodología.

    2. REAL del día: suma de col[10] (peso báscula) de todas las presentaciones
       del lote. Si col[10] está vacía, usar col[6] como respaldo (solo ocurre
       antes del 7-abril, que está excluido del análisis).

    3. PATRON por lote: viene de col[9] (valor patron) de Entrega PT.
       EXCEPCIÓN: si cant_ejecutada en PDN es distinta de 1 (fracción o
       múltiplo) Y la familia es de las que SUBDIVIDEN, el patron correcto
       es tamaño_lote × cant_ejecutada de PDN.

    4. FAMILIAS QUE SUBDIVIDEN (un lote = múltiples presentaciones):
       - CHORIZO (x10, x15, Paisa x12)
       - SCHON CERVECERO (x900, x450)
       - SCHON POLLO (x900, x450, x250, x100)
       - SEVILLANO (x500, x250, x100)
       Para estas familias: patron se cuenta UNA VEZ por lote-código,
       real se SUMA entre todas las presentaciones del mismo lote-código.

    5. FAMILIAS INDEPENDIENTES (cada fila PT = lote independiente):
       - MANGUERA DELI, MANGUERA KILO, MORTADELA bloque,
         SALCHICHA, SCHON ECONOMICO, SCHON PROMOCION
       Para estas: patron y real se SUMAN por cada fila.

    6. SEDE por familia (fuente: tabla Parametros control, columna H):
       - LV: CHORIZO, MANGUERA DELI LV, MANGUERA KILO LV,
             MORTADELA LV, SALCHICHA, SCHON PROMOCION LV
       - SV: SEVILLANO, SCHON POLLO, SCHON CERVECERO,
             SCHON ECONOMICO, MANGUERA DELI SV, MANGUERA KILO SV,
             MORTADELA SV, SCHON PROMOCION SV

    7. CLAVE DE AGRUPACIÓN PT: (fecha_empaque, lote_base, familia, sede)
       donde lote_base = primeros 2 segmentos del código (EMPAQUE-PDN),
       ignorando la secuencia final. Esto evita que lotes del mismo día
       con distinta secuencia se dupliquen.
       IMPORTANTE: incluir siempre la fecha en la clave para evitar
       fusión de lotes con código repetido en días distintos (puede
       ocurrir por error de digitación).

    8. MORTADELA: solo "bloque LV" y "bloque SV". Excluir tajados
       (x250, x450, jamonada) — se analizarán en una fase futura.

    9. SEMÁFORO de 3 zonas (basado en std% de tabla Parametros):
       - Verde:    patron <= real <= teorico  (rango óptimo)
       - Amarillo: real < patron              (merma mayor al estándar)
       - Rojo:     real > teorico             (supera teórico — revisar)
       teorico = patron / (1 - std/100)
       NOTA: días con reproceso agregado generarán rojo sistemáticamente
       porque el reproceso no se registra en PDN. Se maneja caso a caso.

    10. CÓDIGO DE LOTE PT: formato EMPAQUE-PDN-SECUENCIA (ej: 183-183-1)
        El segundo segmento identifica el lote PDN de origen.
        Un mismo lote PDN puede empacarse en múltiples días.
        Un mismo lote-código puede tener múltiples presentaciones (familias
        que subdividen).

LIMITACIONES CONOCIDAS:
    - Entregas parciales (ej: 27.6 kg de un lote de 1943 kg entregados
      un día y el resto al siguiente) generan alertas falsas. Son casos
      esporádicos — Javi los notifica manualmente cuando ocurren.
    - Reproceso agregado a lotes no se registra en PDN → genera rojos
      que no son problemas reales.
    - Lotes repartidos en 3+ días de empaque: el patron se distribuye
      proporcionalmente entre los días según lotes-código presentes.

VALIDACIÓN ANTES DE PUBLICAR:
    Siempre verificar: suma de col[10] en Excel para el período
    debe coincidir con total "real" del dashboard
    (diferencia aceptable = peso de productos excluidos del análisis).
"""

import re
import json
import base64
import urllib.request
from collections import defaultdict


# ============================================================
# PARÁMETROS — fuente: Parametros control del Excel
# ============================================================

STD = {
    'CHORIZO': 10.0, 'SEVILLANO': 7.0, 'SCHON POLLO': 7.0,
    'SCHON CERVECERO': 3.68, 'SALCHICHA': 7.7, 'MANGUERA DELI': 10.0,
    'MANGUERA KILO': 5.5, 'SCHON ECONOMICO': 3.9,
    'SCHON PROMOCION': 5.0, 'MORTADELA': 6.6
}

GRP = {
    'CHORIZO': 'A', 'SEVILLANO': 'A', 'SCHON POLLO': 'A', 'SCHON CERVECERO': 'A',
    'SALCHICHA': 'B', 'MANGUERA DELI': 'B', 'MANGUERA KILO': 'B',
    'SCHON ECONOMICO': 'B', 'SCHON PROMOCION': 'B', 'MORTADELA': 'C'
}

SUBDIVIDEN = {'CHORIZO', 'SCHON CERVECERO', 'SCHON POLLO', 'SEVILLANO'}

CORTE_INICIO = '2026-04-07'  # fecha de compra de la báscula


# ============================================================
# FUNCIONES DE UTILIDAD
# ============================================================

def parse_date(s):
    M = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
         'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
    d, m, y = s.split('-')
    return f"{2000+int(y):04d}-{M[m]:02d}-{int(d):02d}"


def parse_num(s):
    s = s.strip().strip('"').strip('\\').replace('\\', '').replace(',', '')
    try:
        return float(s)
    except:
        return 0.0


def extract_pdn_lote(lote_code):
    """Extrae el número de lote PDN del código PT (segundo segmento)."""
    parts = lote_code.strip().split('-')
    if len(parts) >= 2:
        try:
            return int(parts[1])
        except:
            pass
    try:
        return int(parts[0])
    except:
        return None


def lote_base(lote_code):
    """Normaliza a EMPAQUE-PDN ignorando la secuencia final.
    Ej: '183-183-1' y '183-183-2' → ambos dan '183-183'
    """
    parts = lote_code.strip().split('-')
    if len(parts) >= 2:
        return f"{parts[0]}-{parts[1]}"
    return parts[0]


# ============================================================
# CLASIFICADORES DE PRODUCTO
# ============================================================

def classify_pt(prod):
    """Clasifica una fila de Entrega PT → (familia, sede) o (None, None)."""
    p = prod.strip().lower()
    # Excluidos del análisis (fase actual)
    if 'mortadela' in p:
        if 'bloque' not in p:
            return None, None  # tajados excluidos (x250, x450, jamonada)
        sede = 'SV' if re.search(r'\bsv\b', p) else 'LV'
        return 'MORTADELA', sede
    if any(x in p for x in ['picada', 'costilla', 'camandul', 'pasta de pollo']):
        return None, None
    if ('jam' in p and 'york' in p) or ('jam' in p and 'pizza' in p):
        return None, None
    # Familias activas
    if 'manguera' in p:
        sede = 'SV' if re.search(r'\bsv\b', p) else 'LV'
        return ('MANGUERA DELI' if 'deli' in p else 'MANGUERA KILO'), sede
    if 'chorizo' in p or 'paisa' in p:
        return 'CHORIZO', 'LV'
    if 'sevillano' in p or 'sevilla' in p:
        return 'SEVILLANO', 'SV'
    if 'schon' in p and 'pollo' in p:
        return 'SCHON POLLO', 'SV'
    if 'schon' in p and 'cervecero' in p:
        return 'SCHON CERVECERO', 'SV'
    if 'scha super' in p or 'salchicha super' in p:
        return 'SALCHICHA', 'LV'
    if 'schon' in p and 'econ' in p:
        return 'SCHON ECONOMICO', 'SV'
    if 'schon' in p and 'prom' in p:
        sede = 'SV' if re.search(r'\bsv\b', p) else 'LV'
        return 'SCHON PROMOCION', sede
    return None, None


def classify_pdn(prod):
    """Clasifica una fila de Entrega PDN → (familia, sede) o (None, None)."""
    p = prod.strip().lower()
    if 'manguera' in p:
        sede = 'SV' if re.search(r'\bsv\b', p) else 'LV'
        return ('MANGUERA DELI' if 'deli' in p else 'MANGUERA KILO'), sede
    if 'chorizo' in p:
        return 'CHORIZO', 'LV'
    if 'sevillano' in p or 'sevilla' in p:
        return 'SEVILLANO', 'SV'
    if 'pollo' in p and 'pasta' not in p:
        return 'SCHON POLLO', 'SV'
    if 'cervecero' in p:
        return 'SCHON CERVECERO', 'SV'
    if 'salchicha' in p or ('super' in p and 'econ' not in p and 'prom' not in p):
        return 'SALCHICHA', 'LV'
    if 'econ' in p:
        return 'SCHON ECONOMICO', 'SV'
    if 'promoc' in p:
        sede = 'SV' if re.search(r'\bsv\b', p) else 'LV'
        return 'SCHON PROMOCION', sede
    if 'mortadela' in p:
        sede = 'SV' if re.search(r'\bsv\b', p) else 'LV'
        return 'MORTADELA', sede
    return None, None


# ============================================================
# PARSER PRINCIPAL
# ============================================================

def parse_produccion(text):
    """
    Parsea el texto completo del archivo de producción y retorna
    una lista de records listos para el dashboard.
    """

    # --- PASO 1: leer PDN ---
    # Para familias que SUBDIVIDEN con cant_ejecutada != 1,
    # el patron correcto viene de PDN (tamaño × cant)
    pdn_start = text.find('Entrega PDN FECHA DE PRODUCCIÓN')
    pdn_section = text[pdn_start:]
    pdn_re = re.compile(r'(\d{1,2}-\w{3}-\d{2,4}),(\d+),([^,]+),([\d.]+),([\d.]+)\s')

    pdn_patron = {}  # (lote_num, fam, sede) → kg producidos
    for fs, ls, prod, tam, cant in pdn_re.findall(pdn_section):
        fam, sede = classify_pdn(prod)
        if fam is None or fam not in SUBDIVIDEN:
            continue
        cant_f = float(cant)
        if cant_f != 1.0:  # fracción o múltiplo — necesita override
            pdn_patron[(int(ls), fam, sede)] = float(tam) * cant_f

    print(f"PDN overrides (cant≠1, familias que subdividen): {len(pdn_patron)}")

    # --- PASO 2: inventariar lotes-base en PT por (pdn_lote, fam, sede) ---
    # Necesario para repartir el patron proporcionalmente entre días de empaque
    pt_start = text.find('Entrega PT FECHA DE EMPAQUE')
    # Corte dinámico: buscar el fin real de la sección PT
    pt_end = text.find("Parametros control Producci", pt_start)
    pt_raw = text[pt_start:pt_end]  # corte antes de basura #REF/#N/A

    NUM_FIELD = r'(?:\\+"[\d,]+\.\d+\\+"|"[\d,]+\.\d+"|[\d.]+)'
    ROW_RE = re.compile(
        r'(\d{1,2}-\w{3}-26),'
        r'([^,]+),'    # producto
        r'([^,]*),'    # paq_caja
        r'([^,]+),'    # lote_code
        r'([^,]*),'    # cajas
        r'([^,]*),'    # unidades
        rf'({NUM_FIELD}),'  # col6: total entregado
        r'([^,]*),'    # col7: recorte
        r'([^,]*),'    # col8: num tajadas
        r'([^,]*),'    # col9: valor patron
        r'([^,]*),'    # col10: peso lote terminado (REAL)
    )

    lotes_base_por_pdn = defaultdict(set)
    for m in ROW_RE.finditer(pt_raw):
        fs, prod, paq, lote_code, cajas, units, rg, rec, taj, pat, peso = m.groups()
        fam, sede = classify_pt(prod)
        if fam is None or fam not in SUBDIVIDEN:
            continue
        if parse_date(fs) < CORTE_INICIO:
            continue
        pdn_lote = extract_pdn_lote(lote_code)
        if pdn_lote:
            lotes_base_por_pdn[(pdn_lote, fam, sede)].add(lote_base(lote_code))

    # --- PASO 3: acumular real y patron por lote-base ---
    lotes = {}
    for m in ROW_RE.finditer(pt_raw):
        fs, prod, paq, lote_code, cajas, units, rg, rec, taj, pat_s, peso_s = m.groups()
        fam, sede = classify_pt(prod)
        if fam is None:
            continue
        fecha = parse_date(fs)
        if fecha < CORTE_INICIO:
            continue

        lb = lote_base(lote_code)
        pdn_lote = extract_pdn_lote(lote_code)
        peso_str = peso_s.strip()
        # REGLA: usar col10 (báscula) siempre; col6 como respaldo si col10 vacía
        real_val = parse_num(peso_str) if peso_str != '' else parse_num(rg)
        patron_row = parse_num(pat_s) if pat_s.strip() not in ('', 'No existe') else 0.0
        subdivide = fam in SUBDIVIDEN

        key = (fecha, lb, fam, sede)  # clave con fecha para evitar fusión entre días

        if subdivide and pdn_lote:
            pdn_key = (pdn_lote, fam, sede)
            if pdn_key in pdn_patron:
                # Override PDN: repartir patron entre todos los lotes-base del período
                pat_total = pdn_patron[pdn_key]
                n_total = len(lotes_base_por_pdn.get(pdn_key, {1}))
                patron_este = pat_total / n_total if n_total > 0 else patron_row
            else:
                patron_este = patron_row  # cant=1, patron de PT es correcto
        else:
            patron_este = patron_row

        if key not in lotes:
            lotes[key] = {
                'fecha': fecha, 'fam': fam, 'sede': sede,
                'patron': patron_este, 'real': real_val,
                'n': 1, 'subdivide': subdivide
            }
        else:
            lotes[key]['real'] += real_val
            lotes[key]['n'] += 1
            # Para familias independientes: sumar patron de cada fila
            if not subdivide:
                lotes[key]['patron'] += patron_este

    print(f"Lotes-base únicos parseados: {len(lotes)}")

    # --- PASO 4: agregar por (fecha, familia, sede) ---
    daily = defaultdict(lambda: {'patron': 0., 'real': 0., 'lotes': 0})
    for key, v in lotes.items():
        if v['patron'] == 0 and v['real'] == 0:
            continue
        dk = (v['fecha'], v['fam'], v['sede'])
        daily[dk]['patron'] += v['patron']
        daily[dk]['real'] += v['real']
        daily[dk]['lotes'] += 1

    # --- PASO 5: construir records con semáforo ---
    records = []
    for (fecha, fam, sede), v in sorted(daily.items()):
        if v['patron'] == 0:
            continue
        std = STD.get(fam, 6.5)
        mr = round((v['patron'] - v['real']) / v['patron'] * 100, 1)
        dv = round(std - mr, 1)
        teorico = round(v['patron'] / (1 - std / 100), 1)

        # Semáforo de 3 zonas
        if v['real'] > teorico:
            semaforo = 'red'      # supera teórico — revisar (puede ser reproceso)
        elif v['real'] < v['patron']:
            semaforo = 'yellow'   # merma mayor al estándar
        else:
            semaforo = 'green'    # rango óptimo

        records.append({
            'fecha': fecha, 'fam': fam, 'sede': sede,
            'grp': GRP.get(fam, 'B'),
            'patron': round(v['patron'], 1),
            'real': round(v['real'], 1),
            'teorico': teorico,
            'mr': mr, 'std': std, 'dv': dv,
            'lotes': v['lotes'],
            'semaforo': semaforo
        })

    return records


# ============================================================
# EJECUCIÓN (cuando se llama directamente desde Claude)
# ============================================================

if __name__ == '__main__':
    print("Este script se ejecuta desde una sesión de Claude.")
    print("Requiere que el texto de producción ya esté en memoria como 'text'.")
    print("Ver instrucciones de uso en el README del repositorio.")
