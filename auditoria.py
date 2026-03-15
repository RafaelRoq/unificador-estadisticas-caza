"""
auditoria.py
============
Auditoría rigurosa de los datos unificados de caza (2005-2023).

Cuatro niveles de verificación:
  1. REPARSE — Re-parsea cada año y compara fila por fila contra el CSV.
     Detecta filas perdidas, sobras, o valores corruptos.
  2. TOTALES — Compara la suma de valores del CSV contra totales del Excel original.
  3. SAMPLING — Verifica N registros aleatorios contra el Excel original (2 seeds).
  4. INTEGRIDAD — Checks de coherencia (duplicados, NaN en claves, años completos).

Uso:
    python auditoria.py
"""

import random
import re
import sys
import traceback
import urllib.parse
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, ".")
from unificador import (
    YEAR_CONFIG, PARSERS, ALT_PARSERS, _local_path, parse_year,
    _build_col_map, _iter_data_rows, _safe_num,
    norm_ccaa, norm_prov, norm_esp, norm_cat, is_total,
    _norm_tipo_procedencia, _norm_metrica, PROV_TO_CCAA,
    _cualificar_otras, norm_tipo_coto,
)

OUTPUT_DIR = Path("output")
AUDIT_FILE = OUTPUT_DIR / "auditoria.csv"
N_SAMPLE = 30
SEEDS    = [42, 99]

# ─────────────────────────────────────────────────────────────────────────────
# NIVEL 1: REPARSE — Re-parsea y compara contra CSV
# ─────────────────────────────────────────────────────────────────────────────

def audit_reparse(df_cache: dict) -> list:
    """Re-parsea cada año y compara contra los CSV generados."""
    results = []
    años = sorted(YEAR_CONFIG.keys(), reverse=True)

    for anio in años:
        datos = parse_year(anio)
        # Aplicar el mismo post-procesado que build_final
        for tipo, df_fresh in datos.items():
            if "categoria" in df_fresh.columns and "especie" in df_fresh.columns:
                df_fresh = _cualificar_otras(df_fresh)
                datos[tipo] = df_fresh
            if "tipo_coto" in df_fresh.columns:
                df_fresh["tipo_coto"] = df_fresh["tipo_coto"].apply(norm_tipo_coto)
                datos[tipo] = df_fresh
        for tipo, df_fresh in datos.items():
            if tipo not in df_cache:
                results.append({"nivel": "REPARSE", "anio": anio, "tabla": tipo,
                                "check": "csv_missing", "resultado": "FAIL",
                                "detalle": f"{tipo}.csv no encontrado"})
                continue

            df_csv = df_cache[tipo]
            df_csv_year = df_csv[df_csv["anio"] == anio].copy()

            # Comparar conteo de filas
            n_fresh = len(df_fresh)
            n_csv   = len(df_csv_year)
            if n_fresh != n_csv:
                results.append({"nivel": "REPARSE", "anio": anio, "tabla": tipo,
                                "check": "row_count", "resultado": "FAIL",
                                "detalle": f"reparse={n_fresh} vs csv={n_csv}"})
            else:
                results.append({"nivel": "REPARSE", "anio": anio, "tabla": tipo,
                                "check": "row_count", "resultado": "OK",
                                "detalle": f"{n_fresh} filas"})

            # Comparar contenido
            sort_keys = {
                "licencias_ccaa": ["ccaa"],
                "licencias_prov": ["ccaa", "provincia"],
                "capturas":       ["ccaa", "provincia", "categoria", "especie"],
                "sueltas":        ["ccaa", "provincia", "tipo_procedencia", "categoria", "especie"],
                "produccion":     ["ccaa", "provincia", "categoria", "especie"],
                "terrenos":       ["ccaa", "provincia", "tipo_coto"],
            }
            keys = sort_keys.get(tipo, [])
            common_cols = [c for c in df_fresh.columns if c in df_csv_year.columns]

            try:
                df_a = df_fresh[common_cols].sort_values(keys).reset_index(drop=True)
                df_b = df_csv_year[common_cols].sort_values(keys).reset_index(drop=True)

                # Comparar valor por valor con tolerancia para floats
                mismatches = 0
                if len(df_a) == len(df_b):
                    for col in common_cols:
                        if col == "anio":
                            continue
                        for idx in range(len(df_a)):
                            va = df_a[col].iloc[idx]
                            vb = df_b[col].iloc[idx]
                            if pd.isna(va) and pd.isna(vb):
                                continue
                            if pd.isna(va) or pd.isna(vb):
                                mismatches += 1
                                continue
                            if isinstance(va, float) and isinstance(vb, float):
                                if abs(va - vb) > 0.01:
                                    mismatches += 1
                            elif va != vb:
                                mismatches += 1

                if mismatches == 0 and n_fresh == n_csv:
                    results.append({"nivel": "REPARSE", "anio": anio, "tabla": tipo,
                                    "check": "content", "resultado": "OK",
                                    "detalle": "idéntico"})
                else:
                    results.append({"nivel": "REPARSE", "anio": anio, "tabla": tipo,
                                    "check": "content", "resultado": "FAIL",
                                    "detalle": f"{mismatches} celdas distintas"})
            except Exception as e:
                results.append({"nivel": "REPARSE", "anio": anio, "tabla": tipo,
                                "check": "content", "resultado": "ERROR",
                                "detalle": str(e)[:100]})

    return results


# ─────────────────────────────────────────────────────────────────────────────
# NIVEL 2: TOTALES — Suma del CSV vs total general del Excel
# ─────────────────────────────────────────────────────────────────────────────
#
# Funciones organizadas por formato de tabla.  Cada función lee el Excel
# directamente y suma las filas de datos (no la fila Total del Excel),
# para ser independiente de la etiqueta de total y no mezclar CAZA/PESCA.
# ─────────────────────────────────────────────────────────────────────────────

def _match_close(csv_val, excel_val):
    """Compara dos valores con tolerancia de 0.1% o diferencia absoluta ≤ 1."""
    if excel_val == 0:
        return abs(csv_val) <= 1
    return abs(csv_val - excel_val) / abs(excel_val) < 0.001 or abs(csv_val - excel_val) <= 1


# ── Licencias ────────────────────────────────────────────────────────────────

def _total_licencias_ccaa(xl, cfg):
    """2016-2023: Licencias por CCAA.  Suma filas de datos."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    total = 0
    for xrow in _iter_data_rows(raw, l1 + 1):
        ccaa = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if not ccaa or ccaa == "nan" or len(ccaa) > 60:
            continue
        if is_total(ccaa):
            break
        val = _safe_num(xrow.iloc[2])
        if val is not None:
            total += val
    return {"licencias_expedidas": total}


def _total_licencias_prov(xl, cfg):
    """2007-2015: Licencias por provincia con columna CCAA.  Suma filas de datos.
    CCAAs sin desglose provincial se contabilizan usando el valor de la fila
    'Total CCAA' cuando todas las provincias tienen valor vacío."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    col_c = cfg["col_ccaa"]
    col_p = cfg.get("col_prov", col_c + 1)
    col_l = cfg["col_lic"]

    HEADER_VALS = {"cc.aa.", "comunidad autónoma", "comunidad autonoma",
                   "c.c.a.a.", "ccaa"}
    total = 0
    ccaa_actual = None
    ccaa_has_prov = {}    # ccaa → True si alguna provincia tiene datos
    ccaa_only_val = {}    # ccaa → valor de la fila CCAA sin provincia
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[col_c]).strip() if pd.notna(xrow.iloc[col_c]) else ""
        c1 = str(xrow.iloc[col_p]).strip() if pd.notna(xrow.iloc[col_p]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower() in HEADER_VALS or c1.lower() in ("provincia", "province"):
            continue
        # Saltar notas de años anteriores como "Álava (2010)"
        if c0 and "(" in c0 and c0.rstrip().endswith(")"):
            interior = c0.rstrip().rsplit("(", 1)[-1].rstrip(")")
            if interior.isdigit() and len(interior) == 4:
                continue
        if c0 and not is_total(c0) and len(c0) < 60:
            ccaa_actual = c0  # usar nombre raw para evitar colisiones por normalización
            if not c1:
                val = _safe_num(xrow.iloc[col_l])
                if val is not None:
                    ccaa_only_val[ccaa_actual] = val
                continue
        if c0 and is_total(c0):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            # Fila "Total CCAA": si no hubo provincias, usar valor del total
            if ccaa_actual and ccaa_actual not in ccaa_has_prov:
                total_v = _safe_num(xrow.iloc[col_l])
                if total_v is not None:
                    ccaa_only_val[ccaa_actual] = total_v
            ccaa_actual = None
            continue
        if c1 and is_total(c1):
            continue
        if not c1 or not ccaa_actual or len(c1) > 60:
            continue

        val = _safe_num(xrow.iloc[col_l])
        if val is not None:
            ccaa_has_prov[ccaa_actual] = True
            total += val
    # Añadir CCAAs sin desglose provincial
    for ccaa, val in ccaa_only_val.items():
        if ccaa not in ccaa_has_prov:
            total += val
    return {"licencias_expedidas": total}


def _total_licencias_prov_flat(xl, cfg):
    """2005-2006: Licencias por provincia sin columna CCAA.  Suma filas de datos."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    col_p = cfg["col_prov"]
    col_l = cfg["col_lic"]

    total = 0
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[col_p]).strip() if pd.notna(xrow.iloc[col_p]) else ""
        if c0 == "nan": c0 = ""
        if not c0:
            continue
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if c0.lower() in ("provincia", "provincias y", "comunidades autónomas",
                          "comunidad autonoma", "comunidad autónoma"):
            continue
        prov = norm_prov(c0)
        if prov not in PROV_TO_CCAA:
            continue

        val = _safe_num(xrow.iloc[col_l])
        if val is not None:
            total += val
    return {"licencias_expedidas": total}


# ── Capturas ─────────────────────────────────────────────────────────────────

def _total_capturas_standard(xl, cfg):
    """2007-2023: Capturas con 2 filas de cabecera.  Suma filas de datos."""
    raw = xl.parse(cfg["hoja"], header=None)
    cid = cfg.get("col_id_start", 1)
    l1  = cfg["l1"]
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))

    total = 0
    for xrow in _iter_data_rows(raw, l1 + 2):
        raw_c1 = str(xrow.iloc[cid]).strip() if pd.notna(xrow.iloc[cid]) else ""
        if raw_c1 == "nan": raw_c1 = ""
        if raw_c1.lower().startswith("total"):
            if raw_c1.lower() in ("total", "total general", "total nacional"):
                break
            continue
        raw_c2 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        if raw_c2 == "nan": raw_c2 = ""
        if raw_c2.lower().startswith("total"):
            continue
        if not raw_c2:
            continue
        for col_i in col_map:
            v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
            if v is not None:
                total += v
    return {"n_capturas": total}


def _total_capturas_prov(xl, cfg):
    """2006: Capturas formato provincial (prov en col 0, 'Datos' en col 1).
    Solo filas 'Número de capturas'."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))

    total = 0
    prov_act = None
    for xrow in _iter_data_rows(raw, l1 + 2):
        c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
        c1 = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue
        if c0:
            prov_act = norm_prov(c0)
        if "número de capturas" not in c1.lower() and "numero de capturas" not in c1.lower():
            continue
        if not prov_act or prov_act not in PROV_TO_CCAA:
            continue

        for col_i in col_map:
            v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
            if v is not None:
                total += v
    return {"n_capturas": total}


def _total_capturas_2005(xl, cfg):
    """2005: Capturas en 3 hojas.  Cada especie = 3 cols (Número, Peso, Valor).
    Solo sumamos las columnas de Número."""
    total = 0
    for sheet_cfg in cfg["sheets"]:
        hoja = sheet_cfg["hoja"]
        if hoja not in xl.sheet_names:
            continue
        raw = xl.parse(hoja, header=None)
        l1 = sheet_cfg["l1"]

        # Columnas de Número: cada especie en row l1, col 1+ (excluyendo totales)
        header = list(raw.iloc[l1])
        num_cols = []
        for i, v in enumerate(header):
            s = str(v).strip() if pd.notna(v) else ""
            if not s or s.lower() == "nan" or i < 1:
                continue
            name_clean = re.sub(r"\s*\(.*?\)\s*", "", s).strip()
            if not name_clean or "total" in name_clean.lower():
                continue
            num_cols.append(i)

        data_start = sheet_cfg.get("data_start", l1 + 3)
        for xrow in _iter_data_rows(raw, data_start):
            c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
            if c0 == "nan": c0 = ""
            if c0.lower() in ("españa", "espana", "total", "total general", "total nacional"):
                break
            if c0.lower().startswith("total"):
                continue
            prov = norm_prov(c0) if c0 else None
            if not prov or prov not in PROV_TO_CCAA:
                continue
            for col_i in num_cols:
                v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
                if v is not None:
                    total += v
    return {"n_capturas": total}


# ── Sueltas ──────────────────────────────────────────────────────────────────

def _total_sueltas_standard(xl, cfg):
    """2007-2023: Sueltas con 2 filas de cabecera.  Suma filas de datos."""
    raw = xl.parse(cfg["hoja"], header=None)
    cid = cfg.get("col_id_start", 1)
    l1  = cfg["l1"]
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))

    total = 0
    prov_actual = None
    for xrow in _iter_data_rows(raw, l1 + 2):
        raw_c1 = str(xrow.iloc[cid]).strip() if pd.notna(xrow.iloc[cid]) else ""
        if raw_c1 == "nan": raw_c1 = ""

        if raw_c1.lower().startswith(("nº", "n°", "kg")):
            continue
        if raw_c1.lower().startswith("total"):
            s_low = raw_c1.lower()
            if (s_low in ("total", "total general", "total nacional")
                    or "nº" in s_low or "n°" in s_low or "kg" in s_low):
                break
            prov_actual = None
            continue
        raw_c2 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        if raw_c2 == "nan": raw_c2 = ""
        if raw_c2.lower().startswith("total"):
            continue
        if raw_c2:
            prov_actual = raw_c2
        if not prov_actual:
            continue
        for col_i in col_map:
            v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
            if v is not None:
                total += v
    return {"n_sueltas": total}


def _total_sueltas_flat(xl, cfg):
    """2007: Sueltas formato plano (single header, especies directamente)."""
    raw = xl.parse(cfg["hoja"], header=None)
    cid = cfg.get("col_id_start", 0)
    l1  = cfg["l1"]

    header = list(raw.iloc[l1])
    SKIP = {"nan", "", "comunidad autonoma", "comunidad autónoma", "provincia",
            "datos", "total", "total general", "cc.aa."}
    data_cols = []
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() in SKIP or "total" in s.lower() or "otras" in s.lower():
            continue
        if i < cid + 3:
            continue
        data_cols.append(i)

    total = 0
    prov_actual = None
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[cid]).strip() if pd.notna(xrow.iloc[cid]) else ""
        if c0 == "nan": c0 = ""
        if c0.lower().startswith("total"):
            s_low = c0.lower()
            if s_low in ("total", "total general", "total nacional"):
                break
            prov_actual = None
            continue
        # Saltar subtotales CCAA ("nº sueltas administracion Aragón", etc.)
        if c0.lower().startswith(("nº", "n°", "kg")):
            continue
        c1 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        if c1 == "nan": c1 = ""
        if c1.lower().startswith("total"):
            continue
        if c1:
            prov_actual = c1
        if not prov_actual:
            continue
        c2 = str(xrow.iloc[cid + 2]).strip() if pd.notna(xrow.iloc[cid + 2]) else ""
        if not c2:
            continue
        for col_i in data_cols:
            v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
            if v is not None:
                total += v
    return {"n_sueltas": total}


def _total_sueltas_prov(xl, cfg):
    """2006: Sueltas formato provincial (single header).  Suma filas de datos."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]

    header = list(raw.iloc[l1])
    SKIP = {"nan", "", "provincia", "datos", "total", "total general"}
    data_cols = []
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() in SKIP or "total" in s.lower() or i < 2:
            continue
        data_cols.append(i)

    total = 0
    prov_act = None
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
        c1 = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue
        if c0:
            prov_act = norm_prov(c0)
        if not prov_act or prov_act not in PROV_TO_CCAA:
            continue
        if not c1:
            continue
        for col_i in data_cols:
            v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
            if v is not None:
                total += v
    return {"n_sueltas": total}


# ── Producción ───────────────────────────────────────────────────────────────

def _total_produccion_standard(xl, cfg):
    """2007-2023: Producción con 2 filas de cabecera.  Suma filas de datos."""
    raw = xl.parse(cfg["hoja"], header=None)
    cid = cfg.get("col_id_start", 1)
    l1  = cfg["l1"]
    col_map = _build_col_map(
        list(raw.iloc[l1]), list(raw.iloc[l1 + 1]),
        cat_overrides=cfg.get("cat_overrides"),
    )

    total = 0
    for xrow in _iter_data_rows(raw, l1 + 2):
        raw_c1 = str(xrow.iloc[cid]).strip() if pd.notna(xrow.iloc[cid]) else ""
        raw_c2 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        if raw_c1 == "nan": raw_c1 = ""
        if raw_c2 == "nan": raw_c2 = ""

        if raw_c1.lower().startswith("total"):
            if raw_c1.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if raw_c2.lower().startswith("total"):
            continue
        if not raw_c2:
            continue
        for col_i in col_map:
            v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
            if v is not None:
                total += v
    return {"n_produccion": total}


def _total_produccion_prov(xl, cfg):
    """2006: Producción formato provincial.  Solo filas 'Número de ejemplares'."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1_cat = cfg.get("l1_cat", cfg.get("l1", 5))
    l1_esp = cfg.get("l1_esp", l1_cat + 1)
    col_map = _build_col_map(list(raw.iloc[l1_cat]), list(raw.iloc[l1_esp]))

    total = 0
    prov_act = None
    for xrow in _iter_data_rows(raw, l1_esp + 1):
        c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
        c1 = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue
        if c0:
            prov_act = norm_prov(c0)
        if "número" not in c1.lower() and "numero" not in c1.lower():
            continue
        if "producido" not in c1.lower() and "producción" not in c1.lower():
            continue
        if not prov_act or prov_act not in PROV_TO_CCAA:
            continue

        for col_i in col_map:
            v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
            if v is not None:
                total += v
    return {"n_produccion": total}


# ── Terrenos ─────────────────────────────────────────────────────────────────

def _total_terrenos_standard(xl, cfg):
    """2008-2023: Terrenos con header estándar (single o double).  Suma filas de datos."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1            = cfg["l1"]
    col_id_start  = cfg.get("col_id_start", 1)
    single_header = cfg.get("single_header", False)

    if single_header:
        tipos_row  = list(raw.iloc[l1])
        data_start = l1 + 1
        skip_cols  = set()
        for i, v in enumerate(tipos_row):
            s = str(v).strip().lower() if pd.notna(v) else ""
            if "total" in s:
                skip_cols.add(i)
    else:
        tipos_row   = list(raw.iloc[l1 + 1])
        totales_row = list(raw.iloc[l1])
        skip_cols   = set()
        for i, v in enumerate(totales_row):
            s = str(v).strip().lower() if pd.notna(v) else ""
            if s and "total" in s:
                skip_cols.add(i)
        data_start = l1 + 2

    SKIP_LABELS = {"nan", "", "comunidad autónoma", "comunidad autonoma",
                   "cc.aa.", "ccaa", "provincia", "valores", "datos"}
    data_cols = []
    for i, v in enumerate(tipos_row):
        if i in skip_cols:
            continue
        s = str(v).strip().lower() if pd.notna(v) else ""
        if s and s not in SKIP_LABELS and "total" not in s and "oferta" not in s and "prohibida" not in s:
            data_cols.append(i)

    n_sum = 0
    s_sum = 0
    for xrow in _iter_data_rows(raw, data_start):
        c1 = str(xrow.iloc[col_id_start]).strip() if pd.notna(xrow.iloc[col_id_start]) else ""
        c3_idx = col_id_start + 2
        c3 = str(xrow.iloc[c3_idx]).strip() if c3_idx < len(xrow) and pd.notna(xrow.iloc[c3_idx]) else ""
        if c1 == "nan": c1 = ""
        if c3 == "nan": c3 = ""

        if single_header:
            if c1.lower().startswith("total"):
                continue
            c2 = str(xrow.iloc[col_id_start + 1]).strip() if pd.notna(xrow.iloc[col_id_start + 1]) else ""
            if c2 == "nan": c2 = ""
            if c2.lower().startswith("total"):
                continue
            if c1.startswith("*") or (len(c1) > 60 and c3 == ""):
                continue
        else:
            if c1.lower().startswith(("nº ", "sup (ha)", "total", "*")):
                continue
            if c1.startswith("*") or (len(c1) > 60 and c3 == ""):
                continue
            if "(" in c1 and ")" in c1 and c3 == "":
                continue

        metrica = _norm_metrica(c3)
        if metrica == "nº":
            for col_i in data_cols:
                v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
                if v is not None:
                    n_sum += v
        elif metrica == "sup (ha)":
            for col_i in data_cols:
                v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
                if v is not None:
                    s_sum += v

    result = {}
    if n_sum > 0:
        result["n_cotos"] = n_sum
    if s_sum > 0:
        result["sup_ha"] = s_sum
    return result if result else None


def _total_terrenos_paired(xl, cfg):
    """2007: Terrenos formato pareado (nº + sup por tipo de coto).  Suma filas de datos."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    cid = cfg.get("col_id_start", 0)

    header = list(raw.iloc[l1])
    n_cols = []
    for i, v in enumerate(header):
        sv = str(v).strip().lower() if pd.notna(v) else ""
        if not sv or "total" in sv or i <= cid + 1:
            continue
        n_cols.append(i)

    total = 0
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[cid]).strip() if pd.notna(xrow.iloc[cid]) else ""
        c1 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""
        if c0.lower().startswith("total") or c1.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if not c1:
            continue
        for col_i in n_cols:
            v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
            if v is not None:
                total += v
    return {"n_cotos": total} if total > 0 else None


def _total_terrenos_prov(xl, cfg):
    """2006: Terrenos formato provincial (single header, prov en col 0, 'Datos' en col 1)."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]

    header = list(raw.iloc[l1])
    SKIP = {"nan", "", "provincia", "datos", "total", "total general"}
    data_cols = []
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() in SKIP or "total" in s.lower() or i < 2:
            continue
        data_cols.append(i)

    n_sum = 0
    s_sum = 0
    prov_act = None
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
        c1 = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue
        if c0:
            prov_act = norm_prov(c0)
        if not prov_act or prov_act not in PROV_TO_CCAA:
            continue

        metrica = _norm_metrica(c1)
        if metrica == "nº":
            for col_i in data_cols:
                v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
                if v is not None:
                    n_sum += v
        elif metrica == "sup (ha)":
            for col_i in data_cols:
                v = _safe_num(xrow.iloc[col_i]) if col_i < len(xrow) else None
                if v is not None:
                    s_sum += v

    result = {}
    if n_sum > 0:
        result["n_cotos"] = n_sum
    if s_sum > 0:
        result["sup_ha"] = s_sum
    return result if result else None


# ── Dispatch table ───────────────────────────────────────────────────────────

TOTAL_EXTRACTORS = {
    "licencias_ccaa":      ("licencias_ccaa", _total_licencias_ccaa),
    "licencias_prov":      ("licencias_prov", _total_licencias_prov),
    "licencias_prov_flat": ("licencias_prov", _total_licencias_prov_flat),
    "capturas":            ("capturas",       _total_capturas_standard),
    "capturas_prov":       ("capturas",       _total_capturas_prov),
    "capturas_2005":       ("capturas",       _total_capturas_2005),
    "sueltas":             ("sueltas",        _total_sueltas_standard),
    "sueltas_flat":        ("sueltas",        _total_sueltas_flat),
    "sueltas_prov":        ("sueltas",        _total_sueltas_prov),
    "produccion":          ("produccion",     _total_produccion_standard),
    "produccion_prov":     ("produccion",     _total_produccion_prov),
    "terrenos":            ("terrenos",       _total_terrenos_standard),
    "terrenos_paired":     ("terrenos",       _total_terrenos_paired),
    "terrenos_prov":       ("terrenos",       _total_terrenos_prov),
}

CSV_METRICS = {
    "licencias_ccaa": ["licencias_expedidas"],
    "licencias_prov": ["licencias_expedidas"],
    "capturas":       ["n_capturas"],
    "sueltas":        ["n_sueltas"],
    "produccion":     ["n_produccion"],
    "terrenos":       ["n_cotos", "sup_ha"],
}


def audit_totales(df_cache: dict) -> list:
    """Compara sumas del CSV contra totales del Excel para todas las tablas."""
    results = []
    años = sorted(YEAR_CONFIG.keys(), reverse=True)

    for anio in años:
        cfg_year = YEAR_CONFIG[anio]
        path = _local_path(anio)
        if not path.exists():
            continue

        try:
            xl = pd.ExcelFile(path)
        except Exception:
            continue

        for cfg_key, cfg in cfg_year.items():
            if cfg_key not in TOTAL_EXTRACTORS:
                continue

            csv_table, extractor = TOTAL_EXTRACTORS[cfg_key]
            if csv_table not in df_cache:
                continue

            df_year = df_cache[csv_table]
            df_year = df_year[df_year["anio"] == anio]
            if len(df_year) == 0:
                continue

            try:
                excel_totals = extractor(xl, cfg)
                if excel_totals is None:
                    results.append({
                        "nivel": "TOTALES", "anio": anio, "tabla": csv_table,
                        "check": f"total_{cfg_key}",
                        "resultado": "SKIP",
                        "detalle": "no se encontró datos en Excel"
                    })
                    continue

                for metric in CSV_METRICS.get(csv_table, []):
                    if metric not in excel_totals:
                        continue
                    expected = excel_totals[metric]
                    actual = df_year[metric].sum()
                    ok = _match_close(actual, expected)

                    diff = actual - expected
                    results.append({
                        "nivel": "TOTALES", "anio": anio, "tabla": csv_table,
                        "check": f"{metric}",
                        "resultado": "OK" if ok else "FAIL",
                        "detalle": f"csv={actual:,.0f} excel={expected:,.0f} diff={diff:+,.0f}"
                    })

            except Exception as e:
                results.append({
                    "nivel": "TOTALES", "anio": anio, "tabla": csv_table,
                    "check": f"total_{cfg_key}",
                    "resultado": "ERROR",
                    "detalle": str(e)[:120]
                })

    return results


# ─────────────────────────────────────────────────────────────────────────────
# NIVEL 3: SAMPLING — Verificación aleatoria contra Excel
# ─────────────────────────────────────────────────────────────────────────────
#
# Verifiers por formato: cada uno recibe (row, xl, cfg) y devuelve
# (match, val_excel).  match=True si coincide, None si la especie no se
# encuentra en el Excel (skip), False si no coincide.
# ─────────────────────────────────────────────────────────────────────────────

def verify_licencias_ccaa(row, xl, cfg):
    """2016-2023: licencias por CCAA."""
    raw = xl.parse(cfg["hoja"], header=None)
    for xrow in _iter_data_rows(raw, cfg["l1"] + 1):
        ccaa_raw = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if not ccaa_raw or ccaa_raw == "nan" or is_total(ccaa_raw):
            continue
        if norm_ccaa(ccaa_raw) == row["ccaa"]:
            val_excel = _safe_num(xrow.iloc[2])
            val_csv   = row["licencias_expedidas"]
            match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
            return match, val_excel
    return False, None


def verify_licencias_prov(row, xl, cfg):
    """2005-2015: licencias por provincia."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1    = cfg["l1"]
    col_c = cfg.get("col_ccaa")
    col_p = cfg.get("col_prov", (col_c + 1) if col_c is not None else 0)
    col_l = cfg["col_lic"]

    # Formato plano (2005-2006): sin columna de CCAA
    if col_c is None:
        for xrow in _iter_data_rows(raw, l1 + 1):
            c0 = str(xrow.iloc[col_p]).strip() if pd.notna(xrow.iloc[col_p]) else ""
            if c0 == "nan": c0 = ""
            if not c0 or c0.lower().startswith("total"):
                continue
            prov = norm_prov(c0)
            ccaa = PROV_TO_CCAA.get(prov, "")
            if ccaa == row["ccaa"] and prov == row["provincia"]:
                val_excel = _safe_num(xrow.iloc[col_l])
                val_csv   = row["licencias_expedidas"]
                match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
                return match, val_excel
        return False, None

    # Formato con CCAA (2007-2015)
    HEADER_VALS = {"cc.aa.", "comunidad autónoma", "comunidad autonoma",
                   "c.c.a.a.", "ccaa"}
    is_total_ccaa_row = row["provincia"] == "(Total CCAA)"
    ccaa_actual = None
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[col_c]).strip() if pd.notna(xrow.iloc[col_c]) else ""
        c1 = str(xrow.iloc[col_p]).strip() if pd.notna(xrow.iloc[col_p]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower() in HEADER_VALS or c1.lower() in ("provincia", "province"):
            continue
        # Saltar notas de años anteriores como "Álava (2010)"
        if c0 and "(" in c0 and c0.rstrip().endswith(")"):
            interior = c0.rstrip().rsplit("(", 1)[-1].rstrip(")")
            if interior.isdigit() and len(interior) == 4:
                continue
        if c0 and not is_total(c0) and len(c0) < 60:
            ccaa_actual = norm_ccaa(c0)
        if (c0 and is_total(c0)) or (c1 and is_total(c1)):
            s_low = c0.lower()
            if s_low in ("total", "total general", "total nacional"):
                break
            # Fila "Total CCAA": verificar si el CSV tiene (Total CCAA) para esta
            if is_total_ccaa_row and ccaa_actual == row["ccaa"]:
                val_excel = _safe_num(xrow.iloc[col_l])
                val_csv   = row["licencias_expedidas"]
                match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
                return match, val_excel
            continue
        if not c1 or not ccaa_actual or len(c1) > 60:
            continue

        if not is_total_ccaa_row and ccaa_actual == row["ccaa"] and norm_prov(c1) == row["provincia"]:
            val_excel = _safe_num(xrow.iloc[col_l])
            val_csv   = row["licencias_expedidas"]
            match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
            return match, val_excel
    return False, None


def verify_capturas(row, xl, cfg):
    """2007-2023: capturas con 2 filas de cabecera."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1  = cfg["l1"]
    cid = cfg.get("col_id_start", 1)
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))
    esp_cols = {(cat, esp): cidx for cidx, (cat, esp) in col_map.items()}
    target_key = (row["categoria"], row["especie"])
    if target_key not in esp_cols:
        return None, None
    col_idx = esp_cols[target_key]

    ccaa_actual = None
    prov_actual = None
    for xrow in _iter_data_rows(raw, l1 + 2):
        raw_c1 = str(xrow.iloc[cid]).strip()     if pd.notna(xrow.iloc[cid])     else ""
        raw_c2 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        if raw_c1 == "nan": raw_c1 = ""
        if raw_c2 == "nan": raw_c2 = ""

        if raw_c1.lower().startswith("total"):
            ccaa_actual = None
            if raw_c1.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if any(v.lower().startswith("total") for v in [raw_c2] if v):
            continue
        if raw_c1:
            ccaa_actual = norm_ccaa(raw_c1)
        if raw_c2:
            prov_actual = norm_prov(raw_c2)

        if ccaa_actual == row["ccaa"] and prov_actual == row["provincia"]:
            val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
            val_csv   = row["n_capturas"]
            match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
            return match, val_excel
    return False, None


def verify_capturas_prov(row, xl, cfg):
    """2006: capturas formato provincial."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))
    esp_cols = {(cat, esp): cidx for cidx, (cat, esp) in col_map.items()}
    target_key = (row["categoria"], row["especie"])
    if target_key not in esp_cols:
        return None, None
    col_idx = esp_cols[target_key]

    prov_act = None
    for xrow in _iter_data_rows(raw, l1 + 2):
        c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
        c1 = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue
        if c0:
            prov_act = norm_prov(c0)
        if "número de capturas" not in c1.lower() and "numero de capturas" not in c1.lower():
            continue
        if not prov_act:
            continue

        ccaa = PROV_TO_CCAA.get(prov_act, "")
        if ccaa == row["ccaa"] and prov_act == row["provincia"]:
            val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
            val_csv   = row["n_capturas"]
            match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
            return match, val_excel
    return False, None


def verify_capturas_2005(row, xl, cfg):
    """2005: capturas en 3 hojas con 3 cols por especie (Número, Peso, Valor)."""
    for sheet_cfg in cfg["sheets"]:
        if sheet_cfg.get("categoria", "") != row["categoria"]:
            continue
        hoja = sheet_cfg["hoja"]
        if hoja not in xl.sheet_names:
            continue
        raw = xl.parse(hoja, header=None)
        l1 = sheet_cfg["l1"]

        header = list(raw.iloc[l1])
        col_idx = None
        for i, v in enumerate(header):
            s = str(v).strip() if pd.notna(v) else ""
            if not s or s.lower() == "nan" or i < 1:
                continue
            name_clean = re.sub(r"\s*\(.*?\)\s*", "", s).strip()
            if not name_clean or "total" in name_clean.lower():
                continue
            esp = norm_esp(name_clean)
            # Cualificar "Otras"/"Otros" con la categoría (igual que el parser)
            if esp in ("Otras", "Otros"):
                cat = sheet_cfg.get("categoria", "")
                _CUAL = {
                    ("Caza Mayor", "Otras"): "Otros Caza Mayor",
                    ("Caza Mayor", "Otros"): "Otros Caza Mayor",
                    ("Caza Menor de Aves", "Otras"): "Otras Caza Menor de Aves",
                    ("Caza Menor de Aves", "Otros"): "Otras Caza Menor de Aves",
                    ("Caza Menor de Mamíferos", "Otras"): "Otras Caza Menor de Mamíferos",
                    ("Caza Menor de Mamíferos", "Otros"): "Otras Caza Menor de Mamíferos",
                }
                esp = _CUAL.get((cat, esp), esp)
            if esp == row["especie"]:
                col_idx = i
                break
        if col_idx is None:
            continue

        data_start = sheet_cfg.get("data_start", l1 + 3)
        for xrow in _iter_data_rows(raw, data_start):
            c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
            if c0 == "nan": c0 = ""
            if c0.lower() in ("españa", "espana", "total", "total general"):
                break
            if c0.lower().startswith("total"):
                continue
            if not c0:
                continue
            prov = norm_prov(c0)
            ccaa = PROV_TO_CCAA.get(prov, "")
            if ccaa == row["ccaa"] and prov == row["provincia"]:
                val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
                val_csv   = row["n_capturas"]
                match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
                return match, val_excel
    return False, None


def verify_produccion(row, xl, cfg):
    """2007-2023: producción con 2 filas de cabecera."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1  = cfg["l1"]
    cid = cfg.get("col_id_start", 1)
    col_map = _build_col_map(
        list(raw.iloc[l1]), list(raw.iloc[l1 + 1]),
        cat_overrides=cfg.get("cat_overrides"),
    )
    esp_cols = {(cat, esp): cidx for cidx, (cat, esp) in col_map.items()}
    target_key = (row["categoria"], row["especie"])
    if target_key not in esp_cols:
        return None, None
    col_idx = esp_cols[target_key]

    ccaa_actual = None
    prov_actual = None
    for xrow in _iter_data_rows(raw, l1 + 2):
        raw_c1 = str(xrow.iloc[cid]).strip()     if pd.notna(xrow.iloc[cid])     else ""
        raw_c2 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        if raw_c1 == "nan": raw_c1 = ""
        if raw_c2 == "nan": raw_c2 = ""

        if raw_c1.lower().startswith("total"):
            ccaa_actual = None
            if raw_c1.lower() in ("total", "total general", "total nacional"):
                break
            continue
        if any(v.lower().startswith("total") for v in [raw_c2] if v):
            continue
        if raw_c1:
            ccaa_actual = norm_ccaa(raw_c1)
        if raw_c2:
            prov_actual = norm_prov(raw_c2)

        if ccaa_actual == row["ccaa"] and prov_actual == row["provincia"]:
            val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
            val_csv   = row["n_produccion"]
            match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
            return match, val_excel
    return False, None


def verify_produccion_prov(row, xl, cfg):
    """2006: producción formato provincial (prov en col 0, 'Datos' en col 1)."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1_cat = cfg.get("l1_cat", cfg.get("l1", 5))
    l1_esp = cfg.get("l1_esp", l1_cat + 1)
    col_map = _build_col_map(list(raw.iloc[l1_cat]), list(raw.iloc[l1_esp]))
    esp_cols = {(cat, esp): cidx for cidx, (cat, esp) in col_map.items()}
    target_key = (row["categoria"], row["especie"])
    if target_key not in esp_cols:
        return None, None
    col_idx = esp_cols[target_key]

    prov_act = None
    for xrow in _iter_data_rows(raw, l1_esp + 1):
        c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
        c1 = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue
        if c0:
            prov_act = norm_prov(c0)
        if "número" not in c1.lower() and "numero" not in c1.lower():
            continue
        if "producido" not in c1.lower() and "producción" not in c1.lower():
            continue
        if not prov_act:
            continue

        ccaa = PROV_TO_CCAA.get(prov_act, "")
        if ccaa == row["ccaa"] and prov_act == row["provincia"]:
            val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
            val_csv   = row["n_produccion"]
            match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
            return match, val_excel
    return False, None


def verify_sueltas(row, xl, cfg):
    """2007-2023: sueltas con 2 filas de cabecera."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1  = cfg["l1"]
    cid = cfg.get("col_id_start", 1)
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))
    esp_cols = {(cat, esp): cidx for cidx, (cat, esp) in col_map.items()}
    target_key = (row["categoria"], row["especie"])
    if target_key not in esp_cols:
        return None, None
    col_idx = esp_cols[target_key]

    ccaa_actual = None
    prov_actual = None
    for xrow in _iter_data_rows(raw, l1 + 2):
        raw_c1 = str(xrow.iloc[cid]).strip()     if pd.notna(xrow.iloc[cid])     else ""
        raw_c2 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        raw_c3 = str(xrow.iloc[cid + 2]).strip() if pd.notna(xrow.iloc[cid + 2]) else ""
        if raw_c1 == "nan": raw_c1 = ""
        if raw_c2 == "nan": raw_c2 = ""
        if raw_c3 == "nan": raw_c3 = ""

        if raw_c1.lower().startswith(("nº", "n°", "kg")):
            continue
        if raw_c1.lower().startswith("total"):
            ccaa_actual = None
            s_low = raw_c1.lower()
            if (s_low in ("total", "total general", "total nacional")
                    or "nº" in s_low or "n°" in s_low or "kg" in s_low):
                break
            continue
        if any(v.lower().startswith("total") for v in [raw_c2] if v):
            continue
        if raw_c1 and not is_total(raw_c1):
            ccaa_actual = norm_ccaa(raw_c1)
        if raw_c2 and not is_total(raw_c2):
            prov_actual = norm_prov(raw_c2)

        if (ccaa_actual == row["ccaa"]
                and prov_actual == row["provincia"]
                and _norm_tipo_procedencia(raw_c3) == row["tipo_procedencia"]):
            val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
            val_csv   = row["n_sueltas"]
            match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
            return match, val_excel
    return False, None


def verify_sueltas_prov(row, xl, cfg):
    """2006: sueltas formato provincial (single header, especies directamente)."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]

    header = list(raw.iloc[l1])
    SKIP = {"nan", "", "provincia", "datos", "total", "total general"}
    tipo_cols = {}
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() in SKIP or "total" in s.lower() or i < 2:
            continue
        esp = norm_esp(s)
        tipo_cols[esp] = i

    if row["especie"] not in tipo_cols:
        return None, None
    col_idx = tipo_cols[row["especie"]]

    prov_act = None
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
        c1 = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue
        if c0:
            prov_act = norm_prov(c0)
        if not prov_act or not c1:
            continue

        ccaa = PROV_TO_CCAA.get(prov_act, "")
        if (ccaa == row["ccaa"] and prov_act == row["provincia"]
                and _norm_tipo_procedencia(c1) == row["tipo_procedencia"]):
            val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
            val_csv   = row["n_sueltas"]
            match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
            return match, val_excel
    return False, None


def verify_terrenos(row, xl, cfg):
    """2008-2023: terrenos con header estándar (single o double)."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1            = cfg["l1"]
    cid           = cfg.get("col_id_start", 1)
    single_header = cfg.get("single_header", False)

    if single_header:
        tipos_row  = list(raw.iloc[l1])
        skip_cols  = set()
        for i, v in enumerate(tipos_row):
            s = str(v).strip().lower() if pd.notna(v) else ""
            if "total" in s:
                skip_cols.add(i)
    else:
        tipos_row   = list(raw.iloc[l1 + 1])
        totales_row = list(raw.iloc[l1])
        skip_cols   = set()
        for i, v in enumerate(totales_row):
            s = str(v).strip().lower() if pd.notna(v) else ""
            if s and "total" in s:
                skip_cols.add(i)

    SKIP_LABELS = {"nan", "", "comunidad autónoma", "comunidad autonoma",
                   "cc.aa.", "ccaa", "provincia", "valores", "datos"}
    tipo_cols = {}
    for i, v in enumerate(tipos_row):
        if i in skip_cols:
            continue
        s = str(v).strip() if pd.notna(v) else ""
        if s and s.lower() not in SKIP_LABELS:
            tipo_cols[s] = i

    if row["tipo_coto"] not in tipo_cols:
        return None, None
    col_idx = tipo_cols[row["tipo_coto"]]

    metrica_col = "n_cotos"
    if pd.isna(row.get("n_cotos")) and not pd.isna(row.get("sup_ha")):
        metrica_col = "sup_ha"

    data_start  = l1 + 1 if single_header else l1 + 2
    ccaa_actual = None
    prov_actual = None
    for xrow in _iter_data_rows(raw, data_start):
        c1 = str(xrow.iloc[cid]).strip()     if pd.notna(xrow.iloc[cid])     else ""
        c2 = str(xrow.iloc[cid + 1]).strip() if pd.notna(xrow.iloc[cid + 1]) else ""
        c3 = str(xrow.iloc[cid + 2]).strip() if pd.notna(xrow.iloc[cid + 2]) else ""
        if c1 == "nan": c1 = ""
        if c2 == "nan": c2 = ""
        if c3 == "nan": c3 = ""

        if single_header:
            if c1.lower().startswith("total") or c2.lower().startswith("total"):
                continue
            if c1.startswith("*") or (len(c1) > 60 and c3 == ""):
                continue
        else:
            if c1.lower().startswith(("nº ", "sup (ha)", "total", "*")):
                continue
            if c1.startswith("*") or (len(c1) > 60 and c3 == ""):
                continue
            if "(" in c1 and ")" in c1 and c3 == "":
                continue

        if c1 and not is_total(c1):
            ccaa_actual = norm_ccaa(c1)
        if c2 and not is_total(c2):
            prov_actual = norm_prov(c2)

        metrica = _norm_metrica(c3)
        if metrica not in ("nº", "sup (ha)"):
            continue

        if ccaa_actual == row["ccaa"] and prov_actual == row["provincia"]:
            val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
            if metrica_col == "n_cotos" and metrica == "nº":
                val_csv = row["n_cotos"]
                match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
                return match, val_excel
            elif metrica_col == "sup_ha" and metrica == "sup (ha)":
                val_csv = row["sup_ha"]
                match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
                return match, val_excel
    return False, None


def verify_terrenos_prov(row, xl, cfg):
    """2006: terrenos formato provincial (single header, prov+Datos)."""
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]

    header = list(raw.iloc[l1])
    SKIP = {"nan", "", "provincia", "datos", "total", "total general"}
    tipo_cols = {}
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() in SKIP or "total" in s.lower() or i < 2:
            continue
        tipo_cols[norm_tipo_coto(s)] = i

    if row["tipo_coto"] not in tipo_cols:
        return None, None
    col_idx = tipo_cols[row["tipo_coto"]]

    metrica_col = "n_cotos"
    if pd.isna(row.get("n_cotos")) and not pd.isna(row.get("sup_ha")):
        metrica_col = "sup_ha"

    prov_act = None
    for xrow in _iter_data_rows(raw, l1 + 1):
        c0 = str(xrow.iloc[0]).strip() if pd.notna(xrow.iloc[0]) else ""
        c1 = str(xrow.iloc[1]).strip() if pd.notna(xrow.iloc[1]) else ""
        if c0 == "nan": c0 = ""
        if c1 == "nan": c1 = ""

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue
        if c0:
            prov_act = norm_prov(c0)
        if not prov_act:
            continue

        ccaa = PROV_TO_CCAA.get(prov_act, "")
        metrica = _norm_metrica(c1)

        if ccaa == row["ccaa"] and prov_act == row["provincia"]:
            if metrica_col == "n_cotos" and metrica == "nº":
                val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
                val_csv = row["n_cotos"]
                match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
                return match, val_excel
            elif metrica_col == "sup_ha" and metrica == "sup (ha)":
                val_excel = _safe_num(xrow.iloc[col_idx]) if col_idx < len(xrow) else None
                val_csv = row["sup_ha"]
                match = (val_excel == val_csv) or (val_excel is None and pd.isna(val_csv))
                return match, val_excel
    return False, None


# ── Dispatch de verifiers ────────────────────────────────────────────────────

VERIFIER_MAP = {
    "licencias_ccaa":      (verify_licencias_ccaa,  ["licencias_ccaa"]),
    "licencias_prov":      (verify_licencias_prov,   ["licencias_prov", "licencias_prov_flat"]),
    "capturas":            (verify_capturas,         ["capturas"]),
    "capturas_prov":       (verify_capturas_prov,    ["capturas_prov"]),
    "capturas_2005":       (verify_capturas_2005,    ["capturas_2005"]),
    "sueltas":             (verify_sueltas,          ["sueltas", "sueltas_flat"]),
    "sueltas_prov":        (verify_sueltas_prov,     ["sueltas_prov"]),
    "produccion":          (verify_produccion,       ["produccion"]),
    "produccion_prov":     (verify_produccion_prov,  ["produccion_prov"]),
    "terrenos":            (verify_terrenos,         ["terrenos", "terrenos_paired"]),
    "terrenos_prov":       (verify_terrenos_prov,    ["terrenos_prov"]),
}


def audit_sampling(df_cache: dict) -> list:
    """Auditoría por muestreo aleatorio contra Excel original."""
    results = []
    años = sorted(YEAR_CONFIG.keys(), reverse=True)

    for seed in SEEDS:
        print(f"\n── Sampling seed={seed} " + "─" * 40)
        for tipo in ("licencias_ccaa", "licencias_prov", "capturas", "sueltas",
                      "produccion", "terrenos"):
            if tipo not in df_cache:
                continue
            df_full = df_cache[tipo]

            for anio in años:
                cfg_year = YEAR_CONFIG.get(anio, {})

                # Determinar qué verifier y config usar
                verifier = None
                cfg = None

                # 1. Clave directa en cfg_year + VERIFIER_MAP
                if tipo in cfg_year and tipo in VERIFIER_MAP:
                    verifier = VERIFIER_MAP[tipo][0]
                    cfg = cfg_year[tipo]

                # 2. Buscar alt parser que tenga su propio verifier
                if verifier is None:
                    for alt_key, (out_key, _) in ALT_PARSERS.items():
                        if out_key == tipo and alt_key in cfg_year and alt_key in VERIFIER_MAP:
                            verifier = VERIFIER_MAP[alt_key][0]
                            cfg = cfg_year[alt_key]
                            break

                # 3. Fallback: alt parser sin verifier propio → usar verifier del tipo
                if verifier is None:
                    for alt_key, (out_key, _) in ALT_PARSERS.items():
                        if out_key == tipo and alt_key in cfg_year and tipo in VERIFIER_MAP:
                            verifier = VERIFIER_MAP[tipo][0]
                            cfg = cfg_year[alt_key]
                            break

                if verifier is None or cfg is None:
                    continue

                path = _local_path(anio)
                if not path.exists():
                    continue

                df_año = df_full[df_full["anio"] == anio]
                if len(df_año) == 0:
                    continue

                n = min(N_SAMPLE, len(df_año))
                sample = df_año.sample(n=n, random_state=seed)

                try:
                    xl = pd.ExcelFile(path)
                except Exception as e:
                    print(f"  [!] {anio}/{tipo}: error al abrir Excel: {e}")
                    continue

                ok_count = skip_count = fail_count = 0
                for _, row in sample.iterrows():
                    try:
                        match, _ = verifier(row, xl, cfg)
                        if match is None:
                            skip_count += 1
                        elif match:
                            ok_count += 1
                        else:
                            fail_count += 1
                    except Exception:
                        fail_count += 1

                status = "OK" if fail_count == 0 else "FAIL"
                flag   = "OK" if fail_count == 0 else "!!"
                print(f"  [{flag}] seed={seed}  {anio}/{tipo}: "
                      f"{ok_count} OK, {skip_count} skip, {fail_count} FAIL")

                results.append({
                    "nivel": "SAMPLING",
                    "anio": anio,
                    "tabla": tipo,
                    "check": f"seed={seed}",
                    "resultado": status,
                    "detalle": f"{ok_count} OK, {skip_count} skip, {fail_count} FAIL"
                })

    return results


# ─────────────────────────────────────────────────────────────────────────────
# NIVEL 4: INTEGRIDAD CRUZADA
# ─────────────────────────────────────────────────────────────────────────────

def audit_integridad(df_cache: dict) -> list:
    """Checks de integridad: sin años duplicados, claves coherentes, etc."""
    results = []

    # Check 1: Cada tabla tiene exactamente los años esperados
    for tipo, df in df_cache.items():
        años_csv = sorted(df["anio"].unique())
        años_cfg = sorted([a for a, c in YEAR_CONFIG.items()
                          if tipo in c or any(
                              out_key == tipo and alt_key in c
                              for alt_key, (out_key, _) in ALT_PARSERS.items()
                          )])
        missing = sorted(set(años_cfg) - set(años_csv))
        extra   = sorted(set(años_csv) - set(años_cfg))
        if not missing and not extra:
            results.append({"nivel": "INTEGRIDAD", "anio": "-", "tabla": tipo,
                            "check": "años_completos", "resultado": "OK",
                            "detalle": f"{len(años_csv)} años"})
        else:
            results.append({"nivel": "INTEGRIDAD", "anio": "-", "tabla": tipo,
                            "check": "años_completos", "resultado": "FAIL",
                            "detalle": f"faltan={missing} extra={extra}"})

    # Check 2: No hay NaN en claves obligatorias
    key_cols = {
        "licencias_ccaa": ["anio", "ccaa"],
        "licencias_prov": ["anio", "ccaa", "provincia"],
        "capturas":       ["anio", "ccaa", "provincia", "categoria", "especie"],
        "sueltas":        ["anio", "ccaa", "provincia", "tipo_procedencia", "categoria", "especie"],
        "produccion":     ["anio", "ccaa", "provincia", "categoria", "especie"],
        "terrenos":       ["anio", "ccaa", "provincia", "tipo_coto"],
    }
    for tipo, cols in key_cols.items():
        if tipo not in df_cache:
            continue
        df = df_cache[tipo]
        present_cols = [c for c in cols if c in df.columns]
        na_counts = df[present_cols].isna().sum()
        total_na = na_counts.sum()
        if total_na == 0:
            results.append({"nivel": "INTEGRIDAD", "anio": "-", "tabla": tipo,
                            "check": "no_null_keys", "resultado": "OK",
                            "detalle": "sin NaN en claves"})
        else:
            results.append({"nivel": "INTEGRIDAD", "anio": "-", "tabla": tipo,
                            "check": "no_null_keys", "resultado": "FAIL",
                            "detalle": f"NaN: {dict(na_counts[na_counts > 0])}"})

    # Check 3: No hay duplicados por clave
    for tipo, cols in key_cols.items():
        if tipo not in df_cache:
            continue
        df = df_cache[tipo]
        present_cols = [c for c in cols if c in df.columns]
        n_dup = df.duplicated(subset=present_cols).sum()
        if n_dup == 0:
            results.append({"nivel": "INTEGRIDAD", "anio": "-", "tabla": tipo,
                            "check": "sin_duplicados", "resultado": "OK",
                            "detalle": f"{len(df)} filas únicas"})
        else:
            results.append({"nivel": "INTEGRIDAD", "anio": "-", "tabla": tipo,
                            "check": "sin_duplicados", "resultado": "FAIL",
                            "detalle": f"{n_dup} duplicados"})

    return results


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def run_audit():
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Pre-carga de CSVs
    df_cache = {}
    for tipo in ("licencias_ccaa", "licencias_prov", "capturas", "sueltas",
                  "produccion", "terrenos"):
        csv_path = OUTPUT_DIR / f"{tipo}.csv"
        if csv_path.exists():
            df_cache[tipo] = pd.read_csv(csv_path)
        else:
            print(f"  [!] {tipo}.csv no encontrado")

    all_results = []

    # Nivel 1: Reparse
    print("\n" + "═" * 65)
    print("  NIVEL 1: REPARSE — Re-parsear y comparar contra CSV")
    print("═" * 65)
    reparse_results = audit_reparse(df_cache)
    all_results.extend(reparse_results)
    n_fail_rp = sum(1 for r in reparse_results if r["resultado"] == "FAIL")
    n_ok_rp   = sum(1 for r in reparse_results if r["resultado"] == "OK")
    print(f"\n  Reparse: {n_ok_rp} OK, {n_fail_rp} FAIL")

    # Nivel 2: Totales
    print("\n" + "═" * 65)
    print("  NIVEL 2: TOTALES — Sumas CSV vs totales Excel")
    print("═" * 65)
    totales_results = audit_totales(df_cache)
    all_results.extend(totales_results)
    for r in totales_results:
        flag = "OK" if r["resultado"] == "OK" else "!!"
        print(f"  [{flag}] {r['anio']}/{r['tabla']}: {r['detalle']}")

    # Nivel 3: Sampling
    print("\n" + "═" * 65)
    print("  NIVEL 3: SAMPLING — Muestreo aleatorio contra Excel")
    print("═" * 65)
    sampling_results = audit_sampling(df_cache)
    all_results.extend(sampling_results)

    # Nivel 4: Integridad
    print("\n" + "═" * 65)
    print("  NIVEL 4: INTEGRIDAD — Checks de coherencia")
    print("═" * 65)
    integridad_results = audit_integridad(df_cache)
    all_results.extend(integridad_results)
    for r in integridad_results:
        flag = "OK" if r["resultado"] == "OK" else "!!"
        print(f"  [{flag}] {r['tabla']}/{r['check']}: {r['detalle']}")

    # Guardar resultados
    df_audit = pd.DataFrame(all_results)
    df_audit.to_csv(AUDIT_FILE, index=False, encoding="utf-8-sig")

    # Resumen final
    print("\n" + "═" * 65)
    print("  RESUMEN FINAL DE AUDITORÍA")
    print("═" * 65)
    for nivel in ("REPARSE", "TOTALES", "SAMPLING", "INTEGRIDAD"):
        sub = df_audit[df_audit["nivel"] == nivel]
        n_ok   = (sub["resultado"] == "OK").sum()
        n_fail = (sub["resultado"] == "FAIL").sum()
        n_err  = (sub["resultado"] == "ERROR").sum()
        total  = len(sub)
        print(f"  {nivel:12s}: {n_ok:3d} OK  |  {n_fail:3d} FAIL  |  {n_err:3d} ERROR  ({total} checks)")

    total_fail = (df_audit["resultado"] == "FAIL").sum()
    total_err  = (df_audit["resultado"] == "ERROR").sum()
    if total_fail == 0 and total_err == 0:
        print(f"\n  ✓  AUDITORÍA LIMPIA — todos los checks pasados")
    else:
        print(f"\n  ✗  {total_fail} FAILs + {total_err} ERRORs — revisar {AUDIT_FILE}")
    print("═" * 65 + "\n")


if __name__ == "__main__":
    run_audit()
