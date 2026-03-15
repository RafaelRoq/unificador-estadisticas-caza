"""
test_unitarios.py
=================
Test unitario por año. Cada check_XXXX() lee directamente el Excel de ese año
y compara contra los CSVs de salida.

USO
---
  python -X utf8 test_unitarios.py              # todos los años
  python -X utf8 test_unitarios.py 2023 2006    # años específicos
"""

import sys
import re
from pathlib import Path
import numpy as np
import pandas as pd

from unificador import _local_path, norm_prov, PROV_TO_CCAA

OUTPUT_DIR = Path("output")
TOL = 0.01


def _sn(v):
    """Convierte a float, devuelve 0.0 si no es numérico o NaN."""
    try:
        f = float(v)
        return 0.0 if np.isnan(f) else f
    except (TypeError, ValueError):
        return 0.0


def _csv_año(nombre, anio):
    df = pd.read_csv(OUTPUT_DIR / f"{nombre}.csv")
    return df[df["anio"] == anio].reset_index(drop=True)


def _ok(msg):
    print(f"  [✓] {msg}")


def _fail(msg):
    print(f"  [✗] {msg}")


def _check(label, csv_total, xl_total):
    """Compara dos totales y devuelve error string o None."""
    diff = csv_total - xl_total
    if abs(diff) > TOL:
        _fail(f"{label}: csv={csv_total:.0f} excel={xl_total:.0f} diff={diff:+.0f}")
        return f"{label}: csv={csv_total:.0f} excel={xl_total:.0f} diff={diff:+.0f}"
    else:
        _ok(f"{label}: {csv_total:.0f}")
        return None


# ═════════════════════════════════════════════════════════════════════════════
# HELPERS — lógica compartida entre años
# ═════════════════════════════════════════════════════════════════════════════

_SKIP_LABELS = frozenset({"comunidad autónoma", "comunidad autonoma", "provincia",
                          "tipo de procedencia", "cc.aa.", "datos"})


def _double_header_cols(raw, l1, skip_labels=_SKIP_LABELS):
    """Detecta columnas de datos en una doble cabecera (categoría + especie)."""
    h_cat = list(raw.iloc[l1])
    h_esp = list(raw.iloc[l1 + 1])
    data_cols = []
    last_cat = None
    for i in range(len(h_cat)):
        cat_v = str(h_cat[i]).strip() if pd.notna(h_cat[i]) else ""
        if cat_v and cat_v.lower() != "nan":
            last_cat = cat_v
        esp_v = str(h_esp[i]).strip() if pd.notna(h_esp[i]) else ""
        if not last_cat or not esp_v or esp_v.lower() == "nan":
            continue
        if "total" in last_cat.lower() or "total" in esp_v.lower():
            continue
        if esp_v.lower() in skip_labels:
            continue
        data_cols.append(i)
    return data_cols


def _h_licencias_ccaa(xl, anio, hoja, l1, col_ccaa=1, col_val=2):
    """Licencias a nivel CCAA con deduplicación (2023-2016)."""
    raw = xl.parse(hoja, header=None)
    total_xl = 0.0
    seen = set()
    for _, row in raw.iloc[l1 + 1:].iterrows():
        ccaa = str(row.iloc[col_ccaa]).strip() if pd.notna(row.iloc[col_ccaa]) else ""
        if not ccaa or len(ccaa) > 50:
            continue
        if ccaa.lower().startswith("total"):
            if ccaa.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if ccaa in seen:
            continue
        seen.add(ccaa)
        total_xl += _sn(row.iloc[col_val])
    return _check(f"{anio}/licencias_ccaa",
                  _csv_año("licencias_ccaa", anio)["licencias_expedidas"].sum(), total_xl)


def _h_licencias_prov(xl, anio, hoja, l1, col_ccaa, col_prov, col_val):
    """Licencias a nivel provincial con dict approach (2015-2007)."""
    raw = xl.parse(hoja, header=None)
    total_xl = 0.0
    ccaa_act = None
    ccaa_has_prov = {}
    ccaa_only_val = {}
    for _, row in raw.iloc[l1 + 1:].iterrows():
        c_ccaa = str(row.iloc[col_ccaa]).strip() if pd.notna(row.iloc[col_ccaa]) else ""
        c_prov = str(row.iloc[col_prov]).strip() if pd.notna(row.iloc[col_prov]) else ""
        if c_ccaa.lower() in ("cc.aa.", "comunidad autónoma", "comunidad autonoma",
                               "c.c.a.a.", "ccaa"):
            continue
        if c_prov.lower() in ("provincia", "province"):
            continue
        # Saltar notas de años anteriores como "Álava (2010)"
        if c_ccaa and "(" in c_ccaa and c_ccaa.rstrip().endswith(")"):
            interior = c_ccaa.rstrip().rsplit("(", 1)[-1].rstrip(")")
            if interior.isdigit() and len(interior) == 4:
                continue
        if c_ccaa and not c_ccaa.lower().startswith("total") and len(c_ccaa) < 60:
            ccaa_act = c_ccaa
            if not c_prov:
                v = _sn(row.iloc[col_val])
                if v > 0:
                    ccaa_only_val[ccaa_act] = v
                continue
        if c_ccaa.lower().startswith("total"):
            if c_ccaa.lower() in ("total", "total general", "total nacional"):
                break
            if ccaa_act and ccaa_act not in ccaa_has_prov:
                v = _sn(row.iloc[col_val])
                if v > 0:
                    ccaa_only_val[ccaa_act] = v
            ccaa_act = None
            continue
        if c_prov.lower().startswith("total"):
            continue
        if not ccaa_act or not c_prov or len(c_prov) > 60:
            continue
        v = _sn(row.iloc[col_val])
        if v > 0:
            ccaa_has_prov[ccaa_act] = True
        total_xl += v
    for ccaa, val in ccaa_only_val.items():
        if ccaa not in ccaa_has_prov:
            total_xl += val
    return _check(f"{anio}/licencias_prov",
                  _csv_año("licencias_prov", anio)["licencias_expedidas"].sum(), total_xl)


def _h_capturas_prod(xl, anio, hoja, l1, col_ccaa, col_prov,
                     csv_table="capturas", csv_metric="n_capturas",
                     skip_empty=False):
    """Capturas o producción con doble cabecera (2023-2007)."""
    raw = xl.parse(hoja, header=None)
    data_cols = _double_header_cols(raw, l1)
    total_xl = 0.0
    ccaa_act = None
    for _, row in raw.iloc[l1 + 2:].iterrows():
        if skip_empty:
            vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
            if not vals:
                continue
        c_ccaa = str(row.iloc[col_ccaa]).strip() if pd.notna(row.iloc[col_ccaa]) else ""
        c_prov = str(row.iloc[col_prov]).strip() if pd.notna(row.iloc[col_prov]) else ""
        if c_ccaa.lower().startswith("total"):
            if c_ccaa.lower() in ("total", "total general", "total nacional"):
                break
            ccaa_act = None
            continue
        if c_prov.lower().startswith("total"):
            continue
        if c_ccaa:
            ccaa_act = c_ccaa
        if not ccaa_act or not c_prov:
            continue
        for ci in data_cols:
            total_xl += _sn(row.iloc[ci]) if ci < len(row) else 0
    return _check(f"{anio}/{csv_table}",
                  _csv_año(csv_table, anio)[csv_metric].sum(), total_xl)


def _h_sueltas(xl, anio, hoja, l1, col_id, check_subtotals=True):
    """Sueltas con doble cabecera (2023-2008)."""
    raw = xl.parse(hoja, header=None)
    data_cols = _double_header_cols(raw, l1)
    total_xl = 0.0
    for _, row in raw.iloc[l1 + 2:].iterrows():
        vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if not vals:
            continue
        c_id = str(row.iloc[col_id]).strip() if pd.notna(row.iloc[col_id]) else ""
        if c_id.lower().startswith(("nº", "n°", "kg")):
            continue
        if c_id.lower().startswith("total"):
            s = c_id.lower()
            if s in ("total", "total general", "total nacional") or "nº" in s or "kg" in s:
                break
            continue
        # Saltar subtotales de provincia/tipo ("Total Almería", etc.)
        if check_subtotals and any(str(row.iloc[k]).strip().lower().startswith("total")
               for k in range(data_cols[0]) if pd.notna(row.iloc[k])):
            continue
        for ci in data_cols:
            total_xl += _sn(row.iloc[ci]) if ci < len(row) else 0
    return _check(f"{anio}/sueltas",
                  _csv_año("sueltas", anio)["n_sueltas"].sum(), total_xl)


def _h_terrenos_double(xl, anio, hoja, l1_top, col_ccaa=1, col_val=3):
    """Terrenos con doble cabecera (2023-2014)."""
    raw = xl.parse(hoja, header=None)
    l1_bot = l1_top + 1
    h_top = list(raw.iloc[l1_top])
    h_bot = list(raw.iloc[l1_bot])
    skip_cols = {i for i, v in enumerate(h_top) if pd.notna(v) and "total" in str(v).lower()}
    data_cols = []
    for i, v in enumerate(h_bot):
        if i in skip_cols:
            continue
        s = str(v).strip() if pd.notna(v) else ""
        if s and s.lower() not in ("", "nan", "comunidad autónoma", "provincia", "valores"):
            data_cols.append(i)
    total_n = 0.0
    total_s = 0.0
    ccaa_act = None
    for _, row in raw.iloc[l1_bot + 1:].iterrows():
        vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if not vals:
            continue
        c_ccaa = str(row.iloc[col_ccaa]).strip() if pd.notna(row.iloc[col_ccaa]) else ""
        c_val = str(row.iloc[col_val]).strip() if pd.notna(row.iloc[col_val]) else ""
        if c_ccaa.lower().startswith(("nº ", "sup (ha)", "total", "*")):
            continue
        if "(" in c_ccaa and ")" in c_ccaa and not c_val:
            continue
        if c_ccaa and len(c_ccaa) < 50 and not c_ccaa.lower().startswith("total"):
            ccaa_act = c_ccaa
        if not ccaa_act:
            continue
        c_val_low = c_val.lower().strip()
        if c_val_low in ("nº", "n°") or "número" in c_val_low:
            for ci in data_cols:
                total_n += _sn(row.iloc[ci]) if ci < len(row) else 0
        elif "sup" in c_val_low or ("ha" in c_val_low and "renta" not in c_val_low):
            for ci in data_cols:
                total_s += _sn(row.iloc[ci]) if ci < len(row) else 0
    csv_t = _csv_año("terrenos", anio)
    e1 = _check(f"{anio}/terrenos n_cotos", csv_t["n_cotos"].sum(), total_n)
    e2 = _check(f"{anio}/terrenos sup_ha", csv_t["sup_ha"].sum(), total_s)
    return [e for e in (e1, e2) if e]


def _h_terrenos_single(xl, anio, hoja, l1, col_val=2):
    """Terrenos con cabecera simple (2013-2008)."""
    raw = xl.parse(hoja, header=None)
    header = list(raw.iloc[l1])
    skip_labels = {"nan", "", "comunidad autónoma", "comunidad autonoma", "cc.aa.",
                   "provincia", "valores", "datos", "total", "total general"}
    data_cols = []
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if s.lower() in skip_labels or "total" in s.lower():
            continue
        if i < col_val + 1:
            continue
        data_cols.append(i)
    total_n = 0.0
    total_s = 0.0
    ccaa_act = None
    for _, row in raw.iloc[l1 + 1:].iterrows():
        vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if not vals:
            continue
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        c_val = str(row.iloc[col_val]).strip() if pd.notna(row.iloc[col_val]) else ""
        if c0.lower().startswith("total") or c1.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if c0.startswith("*") or (len(c0) > 60 and c_val == ""):
            continue
        if c0 and len(c0) < 60:
            ccaa_act = c0
        if not ccaa_act:
            continue
        c_val_low = c_val.lower().strip()
        if "nº" in c_val_low or "n°" in c_val_low or "número" in c_val_low or "numero" in c_val_low:
            for ci in data_cols:
                total_n += _sn(row.iloc[ci]) if ci < len(row) else 0
        elif "sup" in c_val_low or ("ha" in c_val_low and "renta" not in c_val_low):
            for ci in data_cols:
                total_s += _sn(row.iloc[ci]) if ci < len(row) else 0
    csv_t = _csv_año("terrenos", anio)
    e1 = _check(f"{anio}/terrenos n_cotos", csv_t["n_cotos"].sum(), total_n)
    e2 = _check(f"{anio}/terrenos sup_ha", csv_t["sup_ha"].sum(), total_s)
    return [e for e in (e1, e2) if e]


# ═════════════════════════════════════════════════════════════════════════════
# CHECK 2023-2016  (xlsx, licencias_ccaa, col_id_start=1)
# ═════════════════════════════════════════════════════════════════════════════

def check_2023():
    anio = 2023
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_ccaa(xl, anio, "1. LICENCIAS", 6)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "3. SUELTAS", 6, 1, check_subtotals=False)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "4. PRODUCCIÓN", 6, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "5. TERRENOS CINEGÉTICOS", 5))
    return errores


def check_2022():
    anio = 2022
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_ccaa(xl, anio, "1. LICENCIAS", 5)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "3.SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "4. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "5. TERRENOS CINEGÉTICOS", 5))
    return errores


def check_2021():
    anio = 2021
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_ccaa(xl, anio, "1. LICENCIAS", 5)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "3.SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "4. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "5. TERRENOS CINEGÉTICOS", 4))
    return errores


def check_2020():
    anio = 2020
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_ccaa(xl, anio, "1. LICENCIAS", 5)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "3. SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "4. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "5. TERRENOS CINEGÉTICOS", 4))
    return errores


def check_2019():
    anio = 2019
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_ccaa(xl, anio, "1. LICENCIAS", 5)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "3. SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "4. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "5. TERRENOS CINEGÉTICOS", 4))
    return errores


def check_2018():
    anio = 2018
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_ccaa(xl, anio, "1. LICENCIAS", 5)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "4. SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "5. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "6. TERRENOS CINEGÉTICOS", 4))
    return errores


def check_2017():
    anio = 2017
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_ccaa(xl, anio, "1. LICENCIAS", 5)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "4. SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "5. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "6. TERRENOS CINEGÉTICOS", 4))
    return errores


def check_2016():
    anio = 2016
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_ccaa(xl, anio, "1. LICENCIAS", 5)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "4. SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "5. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "6. TERRENOS CINEGÉTICOS", 4))
    return errores


# ═════════════════════════════════════════════════════════════════════════════
# CHECK 2015-2014  (xlsx, licencias_prov col_ccaa=1)
# ═════════════════════════════════════════════════════════════════════════════

def check_2015():
    anio = 2015
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_prov(xl, anio, "1. LICENCIAS", 5, 1, 2, 3)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "4. SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "5. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "6. TERRENOS CINEGÉTICOS", 4))
    return errores


def check_2014():
    anio = 2014
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_prov(xl, anio, "1. LICENCIAS", 5, 1, 2, 3)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "2. CAPTURAS CAZA", 6, 1, 2)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, "4. SUELTAS", 5, 1)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, "5. PRODUCCIÓN", 5, 1, 2,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_double(xl, anio, "6. TERRENOS CINEGÉTICOS", 4))
    return errores


# ═════════════════════════════════════════════════════════════════════════════
# CHECK 2013-2008  (xls, col_id_start=0, terrenos single_header)
# ═════════════════════════════════════════════════════════════════════════════

def _check_xls(anio, hoja_lic, hoja_cap, hoja_sue, hoja_prod, hoja_ter,
               l1_lic, l1_cap, l1_sue, l1_prod, l1_ter):
    """Años 2013-2008: xls con col_id_start=0, terrenos single_header."""
    errores = []
    xl = pd.ExcelFile(_local_path(anio))
    e = _h_licencias_prov(xl, anio, hoja_lic, l1_lic, 0, 1, 2)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, hoja_cap, l1_cap, 0, 1)
    if e: errores.append(e)
    e = _h_sueltas(xl, anio, hoja_sue, l1_sue, 0)
    if e: errores.append(e)
    e = _h_capturas_prod(xl, anio, hoja_prod, l1_prod, 0, 1,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)
    errores.extend(_h_terrenos_single(xl, anio, hoja_ter, l1_ter))
    return errores


def check_2013():
    return _check_xls(2013,
        "1.licencias", "2.capturas de caza", "4.sueltas", "5.producción", "6.terrenos cinegeticos",
        6, 4, 4, 5, 4)

def check_2012():
    return _check_xls(2012,
        "1.licencias", "2.capturas de caza", "4.sueltas", "5.producción", "6.terrenos cinegeticos",
        6, 4, 4, 5, 4)

def check_2011():
    return _check_xls(2011,
        "1.licencias", "2.capturas de caza", "4.sueltas", "5.producción", "6.terrenos cinegeticos",
        6, 4, 4, 5, 4)

def check_2010():
    return _check_xls(2010,
        "1.licencias", "2.capturas de caza", "4.sueltas", "5.producción", "6.terrenos cinegeticos",
        6, 4, 4, 5, 4)

def check_2009():
    return _check_xls(2009,
        "1.licencias", "2.capturas de caza", "4.sueltas", "5.producción", "6.terrenos cinegeticos",
        6, 4, 4, 5, 4)

def check_2008():
    return _check_xls(2008,
        "7a.licencias", "7b.capturas de caza", "7d.sueltas", "7e.producción", "7f1.terrenos cinegeticos",
        5, 4, 5, 6, 4)


# ═════════════════════════════════════════════════════════════════════════════
# CHECK 2007  (sueltas_flat + terrenos_paired: formatos únicos)
# ═════════════════════════════════════════════════════════════════════════════
def check_2007():
    anio = 2007
    errores = []
    xl = pd.ExcelFile(_local_path(anio))

    e = _h_licencias_prov(xl, anio, "licencias", 6, 0, 1, 2)
    if e: errores.append(e)

    e = _h_capturas_prod(xl, anio, "capturas_caza", 4, 0, 1)
    if e: errores.append(e)

    # ── sueltas_flat ── single header l1=5, col_id_start=0
    # col0=CCAA, col1=PROV, col2=Datos(tipo_procedencia), cols 3+=especies
    raw = xl.parse("sueltas_cinegéticas", header=None)
    header = list(raw.iloc[5])
    skip_labels = {"nan", "", "comunidad autónoma", "comunidad autonoma", "provincia",
                   "datos", "total", "total general", "cc.aa."}
    data_cols = []
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if s.lower() in skip_labels or "total" in s.lower() or "otras" in s.lower():
            continue
        if i < 3:
            continue
        data_cols.append(i)
    total_xl = 0.0
    for _, row in raw.iloc[6:].iterrows():
        vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if not vals:
            continue
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if c0.lower().startswith("total"):
            s = c0.lower()
            if s in ("total", "total general", "total nacional"):
                break
            continue
        c2 = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not c2:
            continue
        # Saltar subtotales de provincia/tipo ("Total Almería", etc.)
        if any(str(row.iloc[k]).strip().lower().startswith("total")
               for k in range(data_cols[0]) if pd.notna(row.iloc[k])):
            continue
        for ci in data_cols:
            total_xl += _sn(row.iloc[ci]) if ci < len(row) else 0
    e = _check(f"{anio}/sueltas", _csv_año("sueltas", anio)["n_sueltas"].sum(), total_xl)
    if e: errores.append(e)

    e = _h_capturas_prod(xl, anio, "produccion_cinegética", 4, 0, 1,
                         "produccion", "n_produccion", skip_empty=True)
    if e: errores.append(e)

    # ── terrenos_paired ── l1=4, col_id_start=0
    # Fila l1 = tipos de coto, cada tipo ocupa 2 cols: nº y sup
    raw = xl.parse("terrenos_cinegeticos", header=None)
    header_l1 = list(raw.iloc[4])
    tipos = []  # (col_n, col_s, nombre)
    for i, v in enumerate(header_l1):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() == "nan":
            continue
        if i < 2:
            continue
        if "total" in s.lower():
            continue
        tipos.append((i, i + 1, s))
    total_n = 0.0
    total_s = 0.0
    ccaa_act = None
    for _, row in raw.iloc[6:].iterrows():
        vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if not vals:
            continue
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            ccaa_act = None
            continue
        if c1.lower().startswith("total"):
            continue
        if c0 and len(c0) < 60:
            ccaa_act = c0
        if not ccaa_act or not c1:
            continue
        for col_n, col_s, _ in tipos:
            total_n += _sn(row.iloc[col_n]) if col_n < len(row) else 0
            total_s += _sn(row.iloc[col_s]) if col_s < len(row) else 0
    csv_t = _csv_año("terrenos", anio)
    e1 = _check(f"{anio}/terrenos n_cotos", csv_t["n_cotos"].sum(), total_n)
    e2 = _check(f"{anio}/terrenos sup_ha", csv_t["sup_ha"].sum(), total_s)
    if e1: errores.append(e1)
    if e2: errores.append(e2)

    return errores


# ═════════════════════════════════════════════════════════════════════════════
# CHECK 2006  (formato provincial: col0=prov, col1=Datos)
# ═════════════════════════════════════════════════════════════════════════════
def check_2006():
    anio = 2006
    errores = []
    xl = pd.ExcelFile(_local_path(anio))

    # ── licencias_prov_flat ── l1=5, col_prov=0, col_lic=1
    raw = xl.parse("16.a", header=None)
    total_xl = 0.0
    for _, row in raw.iloc[6:].iterrows():
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if not c0:
            continue
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if c0.lower() in ("provincia", "provincias y", "comunidades autónomas",
                           "comunidad autonoma", "comunidad autónoma"):
            continue
        # Filtrar CCAA subtotales: solo incluir provincias reales
        if not PROV_TO_CCAA.get(norm_prov(c0)):
            continue
        total_xl += _sn(row.iloc[1])
    e = _check(f"{anio}/licencias_prov", _csv_año("licencias_prov", anio)["licencias_expedidas"].sum(), total_xl)
    if e: errores.append(e)

    # ── capturas_prov ── l1=4, doble cabecera, col0=prov, col1=Datos
    # Solo filas donde col1 contiene "Número de capturas"
    raw = xl.parse("16.b (2)", header=None)
    h_cat = list(raw.iloc[4])
    h_esp = list(raw.iloc[5])
    data_cols = []
    last_cat = None
    for i in range(len(h_cat)):
        cat_v = str(h_cat[i]).strip() if pd.notna(h_cat[i]) else ""
        if cat_v and cat_v.lower() != "nan":
            last_cat = cat_v
        esp_v = str(h_esp[i]).strip() if pd.notna(h_esp[i]) else ""
        if not last_cat or not esp_v or esp_v.lower() == "nan":
            continue
        if "total" in last_cat.lower() or "total" in esp_v.lower():
            continue
        if esp_v.lower() in ("provincia", "datos"):
            continue
        if i < 2:
            continue
        data_cols.append(i)
    total_xl = 0.0
    prov_act = None
    for _, row in raw.iloc[6:].iterrows():
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if c0:
            prov_act = c0
        if "número de capturas" not in c1.lower() and "numero de capturas" not in c1.lower():
            continue
        if not prov_act:
            continue
        # Filtrar CCAA subtotales: solo incluir provincias reales
        if not PROV_TO_CCAA.get(norm_prov(prov_act)):
            continue
        for ci in data_cols:
            total_xl += _sn(row.iloc[ci]) if ci < len(row) else 0
    e = _check(f"{anio}/capturas", _csv_año("capturas", anio)["n_capturas"].sum(), total_xl)
    if e: errores.append(e)

    # ── sueltas_prov ── l1=5, single header, col0=prov, col1=tipo_procedencia
    raw = xl.parse("16.d (1)", header=None)
    header = list(raw.iloc[5])
    skip_labels = {"nan", "", "provincia", "datos", "total", "total general"}
    data_cols = []
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if s.lower() in skip_labels or "total" in s.lower():
            continue
        if i < 2:
            continue
        data_cols.append(i)
    total_xl = 0.0
    prov_act = None
    for _, row in raw.iloc[6:].iterrows():
        vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if not vals:
            continue
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if c0:
            prov_act = c0
        if not c1 or not prov_act:
            continue
        # Filtrar CCAA subtotales: solo incluir provincias reales
        if not PROV_TO_CCAA.get(norm_prov(prov_act)):
            continue
        # Saltar subtotales de provincia/tipo ("Total Almería", etc.)
        if any(str(row.iloc[k]).strip().lower().startswith("total")
               for k in range(data_cols[0]) if pd.notna(row.iloc[k])):
            continue
        for ci in data_cols:
            total_xl += _sn(row.iloc[ci]) if ci < len(row) else 0
    e = _check(f"{anio}/sueltas", _csv_año("sueltas", anio)["n_sueltas"].sum(), total_xl)
    if e: errores.append(e)

    # ── produccion_prov ── l1_cat=5, l1_esp=6, col0=prov, col1=Datos
    # Solo filas "Número de ejemplares producidos" o similar
    raw = xl.parse("16.e (2)", header=None)
    h_cat = list(raw.iloc[5])
    h_esp = list(raw.iloc[6])
    data_cols = []
    last_cat = None
    for i in range(len(h_cat)):
        cat_v = str(h_cat[i]).strip() if pd.notna(h_cat[i]) else ""
        if cat_v and cat_v.lower() != "nan":
            last_cat = cat_v
        esp_v = str(h_esp[i]).strip() if pd.notna(h_esp[i]) else ""
        if not last_cat or not esp_v or esp_v.lower() == "nan":
            continue
        if "total" in last_cat.lower() or "total" in esp_v.lower():
            continue
        if esp_v.lower() in ("provincia", "datos"):
            continue
        if i < 2:
            continue
        data_cols.append(i)
    total_xl = 0.0
    prov_act = None
    for _, row in raw.iloc[7:].iterrows():
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if c0:
            prov_act = c0
        if "número" not in c1.lower() and "numero" not in c1.lower():
            continue
        if "producido" not in c1.lower() and "producción" not in c1.lower():
            continue
        if not prov_act:
            continue
        # Filtrar CCAA subtotales: solo incluir provincias reales
        if not PROV_TO_CCAA.get(norm_prov(prov_act)):
            continue
        for ci in data_cols:
            total_xl += _sn(row.iloc[ci]) if ci < len(row) else 0
    e = _check(f"{anio}/produccion", _csv_año("produccion", anio)["n_produccion"].sum(), total_xl)
    if e: errores.append(e)

    # ── terrenos_prov ── l1=5, single header, col0=prov, col1=Datos(nº/sup)
    raw = xl.parse("16.f (1)", header=None)
    header = list(raw.iloc[5])
    skip_labels = {"nan", "", "provincia", "datos", "total", "total general"}
    data_cols = []
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if s.lower() in skip_labels or "total" in s.lower():
            continue
        if i < 2:
            continue
        data_cols.append(i)
    total_n = 0.0
    total_s = 0.0
    prov_act = None
    for _, row in raw.iloc[6:].iterrows():
        vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if not vals:
            continue
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if c0:
            prov_act = c0
        if not prov_act:
            continue
        # Filtrar CCAA subtotales: solo incluir provincias reales
        if not PROV_TO_CCAA.get(norm_prov(prov_act)):
            continue
        c1_low = c1.lower().strip()
        if c1_low in ("nº", "n°") or "número" in c1_low:
            for ci in data_cols:
                total_n += _sn(row.iloc[ci]) if ci < len(row) else 0
        elif "sup" in c1_low or ("ha" in c1_low and "renta" not in c1_low):
            for ci in data_cols:
                total_s += _sn(row.iloc[ci]) if ci < len(row) else 0
    csv_t = _csv_año("terrenos", anio)
    e1 = _check(f"{anio}/terrenos n_cotos", csv_t["n_cotos"].sum(), total_n)
    e2 = _check(f"{anio}/terrenos sup_ha", csv_t["sup_ha"].sum(), total_s)
    if e1: errores.append(e1)
    if e2: errores.append(e2)

    return errores


# ═════════════════════════════════════════════════════════════════════════════
# CHECK 2005  (solo licencias + capturas en 3 hojas)
# ═════════════════════════════════════════════════════════════════════════════
def check_2005():
    anio = 2005
    errores = []
    xl = pd.ExcelFile(_local_path(anio))

    # ── licencias_prov_flat ── l1=6, col_prov=0, col_lic=1
    # Solo provincias reales (PROV_TO_CCAA las mapea)
    raw = xl.parse("16.a", header=None)
    total_xl = 0.0
    for _, row in raw.iloc[7:].iterrows():
        c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if not c0:
            continue
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if c0.lower() in ("provincia", "provincias y", "comunidades autónomas",
                           "comunidad autonoma", "comunidad autónoma"):
            continue
        # Filtrar CCAA subtotales: solo incluir provincias reales
        if not PROV_TO_CCAA.get(norm_prov(c0)):
            continue
        total_xl += _sn(row.iloc[1])
    e = _check(f"{anio}/licencias_prov", _csv_año("licencias_prov", anio)["licencias_expedidas"].sum(), total_xl)
    if e: errores.append(e)

    # ── capturas_2005 ── 3 hojas, cada especie = 3 cols (Número, Peso, Valor)
    # Solo sumamos las columnas de Número, filtrando filas CCAA subtotal
    sheets = [
        ("16.b (2)", 4, 7, "Caza Mayor"),
        ("16.b (3)", 4, 7, "Caza Menor de Mamíferos"),
        ("16.b (4)", 4, 7, "Caza Menor de Aves"),
    ]
    total_xl = 0.0
    for hoja, l1, data_start, _cat in sheets:
        if hoja not in xl.sheet_names:
            continue
        raw = xl.parse(hoja, header=None)
        header = list(raw.iloc[l1])
        numero_cols = []
        for i, v in enumerate(header):
            s = str(v).strip() if pd.notna(v) else ""
            if not s or s.lower() == "nan":
                continue
            if i < 1:
                continue
            name_clean = re.sub(r"\s*\(.*?\)\s*", "", s).strip()
            if not name_clean or "total" in name_clean.lower():
                continue
            numero_cols.append(i)
        for _, row in raw.iloc[data_start:].iterrows():
            c0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if c0.lower().startswith("total"):
                if c0.lower() in ("total", "total general", "total nacional"):
                    break
                continue
            if not c0:
                continue
            # Filtrar CCAA subtotales
            if not PROV_TO_CCAA.get(norm_prov(c0)):
                continue
            for ci in numero_cols:
                total_xl += _sn(row.iloc[ci]) if ci < len(row) else 0
    e = _check(f"{anio}/capturas", _csv_año("capturas", anio)["n_capturas"].sum(), total_xl)
    if e: errores.append(e)

    return errores


# ═════════════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════════════

ALL_CHECKS = {
    2023: check_2023,
    2022: check_2022,
    2021: check_2021,
    2020: check_2020,
    2019: check_2019,
    2018: check_2018,
    2017: check_2017,
    2016: check_2016,
    2015: check_2015,
    2014: check_2014,
    2013: check_2013,
    2012: check_2012,
    2011: check_2011,
    2010: check_2010,
    2009: check_2009,
    2008: check_2008,
    2007: check_2007,
    2006: check_2006,
    2005: check_2005,
}


def main():
    if len(sys.argv) > 1:
        años = [int(a) for a in sys.argv[1:]]
    else:
        años = sorted(ALL_CHECKS.keys(), reverse=True)

    todos_errores = {}
    for anio in años:
        fn = ALL_CHECKS.get(anio)
        if not fn:
            print(f"  [!] {anio}: sin función de check")
            continue
        print(f"\n{'═' * 50}")
        print(f"  CHECK {anio}")
        print(f"{'═' * 50}")
        errs = fn()
        if errs:
            todos_errores[anio] = errs

    print(f"\n{'═' * 50}")
    print(f"  RESUMEN")
    print(f"{'═' * 50}")
    ok = len(años) - len(todos_errores)
    fail = len(todos_errores)
    print(f"  {ok} años OK  |  {fail} años con errores")
    if todos_errores:
        for anio, errs in sorted(todos_errores.items(), reverse=True):
            for e in errs:
                print(f"    ✗ {anio}: {e}")
    else:
        print(f"\n  ✓ Todos los años verificados correctamente")

    return 0 if not todos_errores else 1


if __name__ == "__main__":
    sys.exit(main())
