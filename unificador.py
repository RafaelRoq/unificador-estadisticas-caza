"""
unificador.py
=============
Descarga y unifica las estadísticas anuales de caza del MITECO (2005-2023).

Procesa de más nuevo a más antiguo. La tabla final crece en cada iteración:
  tabla_final = tabla_2023
  tabla_final = tabla_final + tabla_2022
  tabla_final = tabla_final + tabla_2021
  ...

Salida (output/):
  capturas.csv    — por CCAA, provincia, especie (formato largo)
  sueltas.csv     — por CCAA, provincia, tipo_procedencia, especie (formato largo)
  produccion.csv  — por CCAA, provincia, especie (formato largo)
  terrenos.csv    — por CCAA, provincia, tipo_coto

USO
---
  python unificador.py --descargar          # descarga los Excel
  python unificador.py --procesar           # unifica los ya descargados
  python unificador.py --todo               # ambos
  python unificador.py --procesar --año 2023 2022   # solo esos años
"""

import argparse
import sys
import time
import urllib.parse
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from tqdm import tqdm

# ─────────────────────────────────────────────────────────────────────────────
# RUTAS
# ─────────────────────────────────────────────────────────────────────────────

RAW_DIR    = Path("raw_data")
OUTPUT_DIR = Path("output")

# ─────────────────────────────────────────────────────────────────────────────
# CATÁLOGO DE URLS
# ─────────────────────────────────────────────────────────────────────────────

BASE = "https://www.miteco.gob.es"

CATALOGO = {
    2023: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef/aef2023/datos-desglosados/AEF_2023_CAZA.xlsx",
    2022: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef/aef2022/datos-desglosados-2022/AEF_2022_CAZA.xlsx",
    2021: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef2021/aef-2021-caza.xlsx",
    2020: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef_2020_caza_tcm30-559683.xlsx",
    2019: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef_2019_caza_tcm30-534500.xlsx",
    2018: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef_2018_cazaypesca_tcm30-521632.xlsx",
    2017: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef_2017_cazaypesca_tcm30-521720.xlsx",
    2016: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef_2016_cazaypesca_tcm30-521748.xlsx",
    2015: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef_2015_cazaypesca_tcm30-453236.xlsx",
    2014: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/aef_2014_cazaypesca_tcm30-453253.xlsx",
    2013: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/caza_pesca2013_tcm30-132439.xls",
    2012: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/CAZA_PESCA%202012_tcm30-132344.xls",
    2011: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/CAZA_PESCA%202011_tcm30-132388.xls",
    2010: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/caza_pesca_tcm30-132190.xls",
    2009: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/caza_y_pesca_fluvial_tcm30-132301.xls",
    2008: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/7_2008_tcm30-132217.xls",
    2007: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/09047122801bd032_tcm30-132780.xls",
    2006: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/090471228012930e_tcm30-132672.xls",
    2005: f"{BASE}/content/dam/miteco/es/biodiversidad/estadisticas/0904712280069b44_tcm30-196292.xls",
}

# ─────────────────────────────────────────────────────────────────────────────
# NORMALIZACIÓN (referencia: nomenclatura 2023)
# ─────────────────────────────────────────────────────────────────────────────

# Claves en minúsculas para comparación case-insensitive
CCAA_NORM = {
    "andalucía": "Andalucía", "andalucia": "Andalucía",
    "aragón": "Aragón", "aragon": "Aragón",
    "asturias": "Asturias", "principado de asturias": "Asturias",
    "baleares": "Baleares", "islas baleares": "Baleares", "illes balears": "Baleares",
    "canarias": "Canarias",
    "cantabria": "Cantabria",
    "castilla y león": "Castilla y León", "castilla y leon": "Castilla y León",
    "castilla-la mancha": "Castilla-La Mancha", "castilla la mancha": "Castilla-La Mancha",
    "cataluña": "Cataluña", "catalunya": "Cataluña",
    "c. valenciana": "C. Valenciana",
    "comunidad valenciana": "C. Valenciana", "comunitat valenciana": "C. Valenciana",
    "com. valenciana": "C. Valenciana",
    "extremadura": "Extremadura",
    "galicia": "Galicia",
    "la rioja": "La Rioja",
    "madrid": "Madrid", "comunidad de madrid": "Madrid",
    "murcia": "Murcia", "región de murcia": "Murcia",
    "navarra": "Navarra", "comunidad foral de navarra": "Navarra",
    "país vasco": "País Vasco", "pais vasco": "País Vasco", "euskadi": "País Vasco",
    # Variantes con anotaciones que aparecen en Excels 2008-2012
    "comunidad valenciana (web)": "C. Valenciana",
    "asturias (web)": "Asturias",
    "aragón (zaragoza y huesca)": "Aragón",
    "aragon (zaragoza y huesca)": "Aragón",
    # Provincias vascas que aparecen como fila suelta al nivel de CCAA
    "álava": "País Vasco", "alava": "País Vasco",
    "guipúzcoa": "País Vasco", "guipuzcoa": "País Vasco",
    "vizcaya": "País Vasco",
}

def norm_ccaa(v):
    """Normaliza nombre de CCAA. Elimina sufijos tipo '(2017)' y marcadores '*'."""
    import re
    s = str(v).strip()
    s = re.sub(r"\s*\(\d{4}\)\s*$", "", s)   # quita "(2017)" etc.
    s = s.rstrip("*").strip()                 # quita asteriscos de nota a pie
    return CCAA_NORM.get(s.lower(), s)

PROV_NORM = {
    "a coruña": "La Coruña", "la coruña": "La Coruña", "coruña, a": "La Coruña",
    "girona": "Gerona", "gerona": "Gerona",
    "lleida": "Lérida", "lérida": "Lérida", "lerida": "Lérida",
    "ourense": "Orense", "orense": "Orense",
    "bizkaia": "Vizcaya", "vizcaya": "Vizcaya",
    "gipuzkoa": "Guipúzcoa", "guipúzcoa": "Guipúzcoa", "guipuzcoa": "Guipúzcoa",
    "araba": "Álava", "álava": "Álava", "alava": "Álava",
    "islas baleares": "Baleares", "illes balears": "Baleares",
    "santa cruz de tenerife": "Tenerife", "tenerife": "Tenerife",
    "s.c.tenerife": "Tenerife", "s.c. tenerife": "Tenerife",
    "las palmas": "Las Palmas",
    "araba/álava": "Álava", "araba / álava": "Álava",
}

# Mapeo provincia → CCAA (para formatos 2005-2006 sin columna de CCAA)
PROV_TO_CCAA = {
    "Álava": "País Vasco", "Guipúzcoa": "País Vasco", "Vizcaya": "País Vasco",
    "Albacete": "Castilla-La Mancha", "Ciudad Real": "Castilla-La Mancha",
    "Cuenca": "Castilla-La Mancha", "Guadalajara": "Castilla-La Mancha",
    "Toledo": "Castilla-La Mancha",
    "Alicante": "C. Valenciana", "Castellón": "C. Valenciana", "Valencia": "C. Valenciana",
    "Almería": "Andalucía", "Cádiz": "Andalucía", "Córdoba": "Andalucía",
    "Granada": "Andalucía", "Huelva": "Andalucía", "Jaén": "Andalucía",
    "Málaga": "Andalucía", "Sevilla": "Andalucía",
    "Asturias": "Asturias",
    "Ávila": "Castilla y León", "Burgos": "Castilla y León", "León": "Castilla y León",
    "Palencia": "Castilla y León", "Salamanca": "Castilla y León",
    "Segovia": "Castilla y León", "Soria": "Castilla y León",
    "Valladolid": "Castilla y León", "Zamora": "Castilla y León",
    "Badajoz": "Extremadura", "Cáceres": "Extremadura",
    "Baleares": "Baleares",
    "Barcelona": "Cataluña", "Gerona": "Cataluña", "Lérida": "Cataluña",
    "Tarragona": "Cataluña",
    "Cantabria": "Cantabria",
    "La Coruña": "Galicia", "Lugo": "Galicia", "Orense": "Galicia",
    "Pontevedra": "Galicia",
    "Huesca": "Aragón", "Teruel": "Aragón", "Zaragoza": "Aragón",
    "La Rioja": "La Rioja",
    "Las Palmas": "Canarias", "Tenerife": "Canarias",
    "Madrid": "Madrid",
    "Murcia": "Murcia",
    "Navarra": "Navarra",
}

# Todas las especies que pueden aparecer; clave en minúsculas
ESPECIE_NORM = {
    # Caza Mayor
    "lobo": "Lobo",
    "arruí": "Arruí", "arrui": "Arruí",
    "cabra asilvestrada": "Cabra asilvestrada",
    "cabra montés": "Cabra montés", "cabra montes": "Cabra montés",
    "cabra monte": "Cabra montés",
    "ciervo": "Ciervo",
    "corzo": "Corzo",
    "gamo": "Gamo",
    "jabalí": "Jabalí", "jabali": "Jabalí",
    "muflón": "Muflón", "muflon": "Muflón",
    "rebeco": "Rebeco", "sarrio": "Rebeco",
    # Caza Menor de Aves
    "acuáticas y anátidas": "Acuáticas y anátidas",
    "acuaticas y anatidas": "Acuáticas y anátidas",
    "aves acuáticas y anátidas": "Acuáticas y anátidas",
    "avefría": "Avefría", "avefrías": "Avefría",
    "becada": "Becada",
    "codorniz": "Codorniz",
    "córvidos": "Córvidos", "corvidos": "Córvidos",
    "estornino": "Estornino",
    "faisán": "Faisán", "faisan": "Faisán",
    "paloma": "Paloma",
    "perdiz": "Perdiz", "perdiz roja": "Perdiz",
    "tórtola comun": "Tórtola común", "tórtola común": "Tórtola común",
    "tortola comun": "Tórtola común",
    "otras": "Otras", "otras aves": "Otras",
    "otras especies": "Otras",   # 2009 capturas/sueltas (Caza Mayor y Aves)
    "otros": "Otros",   # Caza Mayor, aparece desde 2017
    "otras_caza mayor": "Otros",   # variante de 2010/2012
    "varias": "Varias especies",   # variante corta en sueltas 2011
    "zorzal": "Zorzal",
    "varias especies": "Varias especies",
    # Especies separadas en 2008 (en 2009+ se fusionan como "Acuáticas y anátidas")
    "acuáticas": "Acuáticas", "acuaticas": "Acuáticas",
    "anátidas": "Anátidas", "anatidas": "Anátidas",
    "paloma torcaz": "Paloma torcaz",
    "paloma zurita": "Paloma zurita",
    # Caza Menor de Mamíferos
    "conejo": "Conejo",
    "liebre": "Liebre",
    "zorro": "Zorro",
    # Especies con sufijo de categoría (2006-2007)
    "otras_caza menor": "Otras",
    "otras_caza volátil": "Otras",
    "varias especies_caza volátil": "Varias especies",
    # 2006 especies con nombre científico
    "turón": "Turón", "turon": "Turón",
    "tordo": "Tordo",
    "zorzales, becadas, estorninos y otros": "Otras",  # agrupar como "Otras" aves
    # 2005 capturas: nombres con nombre científico entre paréntesis
    "otra caza menor, mamíferos": "Otras",
    "otra caza menor, mamiferos": "Otras",
    "otra caza volátil": "Otras",
    "otra caza volatil": "Otras",
    "otra caza mayor": "Otros",
}

CATEGORIA_NORM = {
    "caza mayor": "Caza Mayor",
    "caza menor de aves": "Caza Menor de Aves",
    "caza menor de mamíferos": "Caza Menor de Mamíferos",
    "caza menor de mamiferos": "Caza Menor de Mamíferos",
    # Variantes sin "de" (2008 sueltas/producción)
    "caza menor mamíferos": "Caza Menor de Mamíferos",
    "caza menor mamiferos": "Caza Menor de Mamíferos",
    "caza menor aves": "Caza Menor de Aves",
    # Con sufijo "(nº de capturas)" (2008 capturas)
    "caza mayor (nº de capturas)": "Caza Mayor",
    "caza menor mamíferos (nº de capturas)": "Caza Menor de Mamíferos",
    "caza menor mamiferos (nº de capturas)": "Caza Menor de Mamíferos",
    "caza menor aves (nº de capturas)": "Caza Menor de Aves",
    # 2006-2007: "Caza Menor" = mamíferos, "Caza Volátil" = aves
    "caza menor": "Caza Menor de Mamíferos",
    "caza volátil": "Caza Menor de Aves",
    "caza volatil": "Caza Menor de Aves",
}


def _n(v, tabla):
    """Normaliza v usando la tabla dada (keys en minúsculas)."""
    s = str(v).strip()
    return tabla.get(s.lower(), s)

def norm_prov(v):
    """Normaliza nombre de provincia. Elimina sufijos tipo '(2010)' igual que norm_ccaa."""
    import re
    s = str(v).strip()
    s = re.sub(r"\s*\(\d{4}\)\s*$", "", s)
    return PROV_NORM.get(s.lower(), s)
def norm_esp(v):     return _n(v, ESPECIE_NORM)
def norm_cat(v):     return _n(v, CATEGORIA_NORM)

def is_total(v):
    """True si la celda es una fila de subtotal, vacía, nota o cabecera de sección."""
    s = str(v).strip().lower()
    return (
        s == "" or s == "nan"
        or s.startswith("total")
        or s.startswith("nota")
        or s.startswith("desde ")
        or s.startswith("las licencias")
        or s.startswith("aragón, asturias")   # lista de CCAA en nota interautonómica
        or s == "comunidad autónoma"           # cabecera de columna que aparece en 2ª tabla
        or s.startswith("número de licencias") # cabecera tabla pesca en hoja licencias 2018
        or s.startswith("sueltas de")          # cabecera tabla piscícola en hoja sueltas 2018
    )


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN POR AÑO
#
# "l1": fila (0-indexed) con las categorías de primer nivel
#       (para hojas con dos niveles de cabecera, las especies están en l1+1)
#       (para LICENCIAS, que solo tiene un nivel, las columnas están en l1)
#
# RESUMEN DE FORMATOS
# ───────────────────
# Hay 4 formatos principales. Las diferencias clave entre años son los nombres
# de hoja, las filas de cabecera (l1), y la columna donde empiezan los IDs.
#
#   Años      Formato     Licencias       col_id  Sueltas         Terrenos         Notas
#   ────────  ──────────  ──────────────  ──────  ──────────────  ───────────────  ──────────────────────
#   2023      xlsx        ccaa (l1=6)     1       doble cab.      doble cab.       l1 capturas=6, rest varía
#   2022      xlsx        ccaa (l1=5)     1       doble cab.      doble cab.       sueltas "3.SUELTAS" (sin espacio)
#   2021      xlsx        ccaa (l1=5)     1       doble cab.      doble cab.       terrenos l1=4
#   2020-2019 xlsx        ccaa (l1=5)     1       doble cab.      doble cab.       produccion 2020: cat_overrides
#   2018-2016 xlsx        ccaa (l1=5)     1       doble cab.      doble cab.       hojas pesca intercaladas (sueltas=4, prod=5, ter=6)
#                                                                                  produccion 2016: cat_overrides
#   2015-2014 xlsx        prov (l1=5)     1       doble cab.      doble cab.       licencias a nivel provincial
#   2013-2009 xls         prov (l1=6)     0       doble cab.      single header    col_id_start=0, hojas en minúsculas
#   2008      xls         prov (l1=5)     0       doble cab.      single header    hojas con prefijo "7x."
#   2007      xls         prov (l1=6)     0       single header   cols pareadas    sueltas sin categorías, terrenos nº+sup pareados
#   2006      xls         prov flat       —       prov flat       prov flat        solo col provincia, sin CCAA
#   2005      xls         prov flat       —       —               —                solo licencias y capturas (3 hojas)
#
# ─────────────────────────────────────────────────────────────────────────────

YEAR_CONFIG = {
    2023: {
        "licencias_ccaa": {"hoja": "1. LICENCIAS",              "l1": 6},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "3. SUELTAS",                "l1": 6},
        "produccion":     {"hoja": "4. PRODUCCIÓN",             "l1": 6},
        "terrenos":       {"hoja": "5. TERRENOS CINEGÉTICOS",   "l1": 5},
    },
    2022: {
        "licencias_ccaa": {"hoja": "1. LICENCIAS",              "l1": 5},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "3.SUELTAS",                 "l1": 5},  # sin espacio
        "produccion":     {"hoja": "4. PRODUCCIÓN",             "l1": 5},
        "terrenos":       {"hoja": "5. TERRENOS CINEGÉTICOS",   "l1": 5},
    },
    2021: {
        "licencias_ccaa": {"hoja": "1. LICENCIAS",              "l1": 5},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "3.SUELTAS",                 "l1": 5},  # sin espacio
        "produccion":     {"hoja": "4. PRODUCCIÓN",             "l1": 5},
        "terrenos":       {"hoja": "5. TERRENOS CINEGÉTICOS",   "l1": 4},  # ← fila distinta
    },
    2020: {
        "licencias_ccaa": {"hoja": "1. LICENCIAS",              "l1": 5},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "3. SUELTAS",                "l1": 5},  # con espacio
        # PRODUCCIÓN 2020: L1 carece del encabezado "Caza Mayor" en cols 3-6;
        # las especies (Ciervo, Gamo, Jabalí, Muflón) están en L2 pero sin categoría.
        # cat_overrides inyecta la categoría manualmente antes del forward-fill.
        "produccion":     {"hoja": "4. PRODUCCIÓN",             "l1": 5,
                           "cat_overrides": {3: "Caza Mayor", 4: "Caza Mayor",
                                             5: "Caza Mayor", 6: "Caza Mayor"}},
        "terrenos":       {"hoja": "5. TERRENOS CINEGÉTICOS",   "l1": 4},
    },
    2019: {
        "licencias_ccaa": {"hoja": "1. LICENCIAS",              "l1": 5},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "3. SUELTAS",                "l1": 5},
        "produccion":     {"hoja": "4. PRODUCCIÓN",             "l1": 5},
        "terrenos":       {"hoja": "5. TERRENOS CINEGÉTICOS",   "l1": 4},
    },
    2018: {
        # Desde 2018 el libro incluye hojas de pesca, desplazando los números:
        # 3=CAPTURAS PESCA (omitir), 4=SUELTAS, 5=PRODUCCIÓN, 6=TERRENOS, 7=PISCÍCOLA (omitir)
        "licencias_ccaa": {"hoja": "1. LICENCIAS",              "l1": 5},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "4. SUELTAS",                "l1": 5},
        "produccion":     {"hoja": "5. PRODUCCIÓN",             "l1": 5},
        "terrenos":       {"hoja": "6. TERRENOS CINEGÉTICOS",   "l1": 4},
    },
    2017: {
        # Misma estructura que 2018 (hojas de pesca intercaladas)
        # CAPTURAS: añade 'Otros' en Caza Mayor (col 13)
        # PRODUCCIÓN: Caza Mayor = Ciervo, Corzo, Gamo, Jabalí, Muflón (sin Cabra montés)
        "licencias_ccaa": {"hoja": "1. LICENCIAS",              "l1": 5},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "4. SUELTAS",                "l1": 5},
        "produccion":     {"hoja": "5. PRODUCCIÓN",             "l1": 5},
        "terrenos":       {"hoja": "6. TERRENOS CINEGÉTICOS",   "l1": 4},
    },
    2016: {
        # Misma estructura pesca que 2017/2018
        # PRODUCCIÓN: carece de encabezado "Caza Mayor" en cols 3-6 → cat_overrides
        "licencias_ccaa": {"hoja": "1. LICENCIAS",              "l1": 5},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "4. SUELTAS",                "l1": 5},
        "produccion":     {"hoja": "5. PRODUCCIÓN",             "l1": 5,
                           "cat_overrides": {3: "Caza Mayor", 4: "Caza Mayor",
                                             5: "Caza Mayor", 6: "Caza Mayor"}},
        "terrenos":       {"hoja": "6. TERRENOS CINEGÉTICOS",   "l1": 4},
    },
    2015: {
        # Licencias: nivel provincia (col 1=CCAA, 2=PROV, 3=lic, 4=importe, 5=interauton, 6=vigentes)
        # CAPTURAS: añade 'Otras' en Caza Mayor (col 13)
        "licencias_prov": {"hoja": "1. LICENCIAS", "l1": 5,
                           "col_ccaa": 1, "col_prov": 2, "col_lic": 3,
                           "col_importe": 4, "col_vigentes": 6},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "4. SUELTAS",                "l1": 5},
        "produccion":     {"hoja": "5. PRODUCCIÓN",             "l1": 5},
        "terrenos":       {"hoja": "6. TERRENOS CINEGÉTICOS",   "l1": 4},
    },
    2014: {
        # Licencias: nivel provincia (col 5=vigentes, sin interautonómicas)
        # CAPTURAS: 'Otras' en Caza Mayor (col 13)
        "licencias_prov": {"hoja": "1. LICENCIAS", "l1": 5,
                           "col_ccaa": 1, "col_prov": 2, "col_lic": 3,
                           "col_importe": 4, "col_vigentes": 5},
        "capturas":       {"hoja": "2. CAPTURAS CAZA",          "l1": 6},
        "sueltas":        {"hoja": "4. SUELTAS",                "l1": 5},
        "produccion":     {"hoja": "5. PRODUCCIÓN",             "l1": 5},
        "terrenos":       {"hoja": "6. TERRENOS CINEGÉTICOS",   "l1": 4},
    },
    2013: {
        # Formato xls: col_id_start=0 (CCAA en col 0, PROV en col 1)
        # Sheet names en minúsculas sin espacio tras el número
        # Licencias: nivel provincia, col 0=CCAA, 1=PROV, 2=lic, 3=importe, 4=vigentes
        # Terrenos: header único (single_header), tipos de coto directamente en fila l1
        "licencias_prov": {"hoja": "1.licencias",  "l1": 6,
                           "col_ccaa": 0, "col_prov": 1, "col_lic": 2,
                           "col_importe": 3, "col_vigentes": 4},
        "capturas":       {"hoja": "2.capturas de caza",     "l1": 4, "col_id_start": 0},
        "sueltas":        {"hoja": "4.sueltas",              "l1": 4, "col_id_start": 0},
        "produccion":     {"hoja": "5.producción",           "l1": 5, "col_id_start": 0},
        "terrenos":       {"hoja": "6.terrenos cinegeticos", "l1": 4, "col_id_start": 0,
                           "single_header": True},
    },
    2012: {
        "licencias_prov": {"hoja": "1.licencias",  "l1": 6,
                           "col_ccaa": 0, "col_prov": 1, "col_lic": 2,
                           "col_importe": 3, "col_vigentes": 4},
        "capturas":       {"hoja": "2.capturas de caza",     "l1": 4, "col_id_start": 0},
        "sueltas":        {"hoja": "4.sueltas",              "l1": 4, "col_id_start": 0},
        "produccion":     {"hoja": "5.producción",           "l1": 5, "col_id_start": 0},
        "terrenos":       {"hoja": "6.terrenos cinegeticos", "l1": 4, "col_id_start": 0,
                           "single_header": True},
    },
    2011: {
        "licencias_prov": {"hoja": "1.licencias",  "l1": 6,
                           "col_ccaa": 0, "col_prov": 1, "col_lic": 2,
                           "col_importe": 3, "col_vigentes": 4},
        "capturas":       {"hoja": "2.capturas de caza",     "l1": 4, "col_id_start": 0},
        "sueltas":        {"hoja": "4.sueltas",              "l1": 4, "col_id_start": 0},
        "produccion":     {"hoja": "5.producción",           "l1": 5, "col_id_start": 0},
        "terrenos":       {"hoja": "6.terrenos cinegeticos", "l1": 4, "col_id_start": 0,
                           "single_header": True},
    },
    2010: {
        "licencias_prov": {"hoja": "1.licencias",  "l1": 6,
                           "col_ccaa": 0, "col_prov": 1, "col_lic": 2,
                           "col_importe": 3, "col_vigentes": 4},
        "capturas":       {"hoja": "2.capturas de caza",     "l1": 4, "col_id_start": 0},
        "sueltas":        {"hoja": "4.sueltas",              "l1": 4, "col_id_start": 0},
        "produccion":     {"hoja": "5.producción",           "l1": 5, "col_id_start": 0},
        "terrenos":       {"hoja": "6.terrenos cinegeticos", "l1": 4, "col_id_start": 0,
                           "single_header": True},
    },
    2009: {
        # Formato xls igual que 2010-2013 (col_id_start=0)
        # Licencias: nivel provincia, misma estructura que 2013
        "licencias_prov": {"hoja": "1.licencias", "l1": 6,
                           "col_ccaa": 0, "col_prov": 1, "col_lic": 2,
                           "col_importe": 3, "col_vigentes": 4},
        "capturas":       {"hoja": "2.capturas de caza",     "l1": 4, "col_id_start": 0},
        "sueltas":        {"hoja": "4.sueltas",              "l1": 4, "col_id_start": 0},
        "produccion":     {"hoja": "5.producción",           "l1": 5, "col_id_start": 0},
        "terrenos":       {"hoja": "6.terrenos cinegeticos", "l1": 4, "col_id_start": 0,
                           "single_header": True},
    },
    2008: {
        # Formato xls; hojas con prefijo "7x." diferente de 2009-2013
        # Capturas: categorías con sufijo "(nº de capturas)" → CATEGORIA_NORM las normaliza
        # Sueltas: tipo_procedencia = etiqueta larga → _norm_tipo_procedencia las normaliza
        #          subtotales CCAA con "Nº ejemplares..." en col0 → filtro en _extract_ids_and_species
        # Producción: l1=6 (una fila extra de cabecera respecto a 2009)
        "licencias_prov": {"hoja": "7a.licencias", "l1": 5,
                           "col_ccaa": 0, "col_prov": 1, "col_lic": 2,
                           "col_importe": 3, "col_vigentes": 4},
        "capturas":       {"hoja": "7b.capturas de caza",       "l1": 4, "col_id_start": 0},
        "sueltas":        {"hoja": "7d.sueltas",                "l1": 5, "col_id_start": 0},
        "produccion":     {"hoja": "7e.producción",             "l1": 6, "col_id_start": 0},
        "terrenos":       {"hoja": "7f1.terrenos cinegeticos",  "l1": 4, "col_id_start": 0,
                           "single_header": True},
    },
    2007: {
        # xls: hojas por tipo; CCAA+provincia en cols 0-1
        # Capturas/Producción: usan "Caza Menor" (mamíferos) y "Caza Volátil" (aves)
        # Sueltas: cabecera plana (sin categorías), tipo en col 2 "Datos"
        # Terrenos: columnas pareadas (nº, superficie por tipo de coto)
        "licencias_prov": {"hoja": "licencias", "l1": 6,
                           "col_ccaa": 0, "col_prov": 1, "col_lic": 2,
                           "col_importe": 3, "col_vigentes": 4},
        "capturas":       {"hoja": "capturas_caza",        "l1": 4, "col_id_start": 0},
        "sueltas_flat":   {"hoja": "sueltas_cinegéticas",  "l1": 5, "col_id_start": 0},
        "produccion":     {"hoja": "produccion_cinegética", "l1": 4, "col_id_start": 0},
        "terrenos_paired": {"hoja": "terrenos_cinegeticos", "l1": 4, "col_id_start": 0},
    },
    2006: {
        # xls: formato provincial (solo provincia en col 0, "Datos" en col 1)
        # No hay columna de CCAA → se infiere de PROV_TO_CCAA
        "licencias_prov_flat": {"hoja": "16.a", "l1": 5,
                                "col_prov": 0, "col_lic": 1, "col_importe": 2,
                                "col_vigentes": 3},
        "capturas_prov":   {"hoja": "16.b (2)", "l1": 4},
        "sueltas_prov":    {"hoja": "16.d (1)", "l1": 5},
        "produccion_prov": {"hoja": "16.e (2)", "l1_cat": 5, "l1_esp": 6},
        "terrenos_prov":   {"hoja": "16.f (1)", "l1": 5},
    },
    2005: {
        # xls: formato provincial; capturas repartidas en 3 hojas
        # Solo licencias y capturas disponibles (no hay sueltas/producción/terrenos)
        "licencias_prov_flat": {"hoja": "16.a", "l1": 6,
                                "col_prov": 0, "col_lic": 1, "col_importe": 2,
                                "col_vigentes": 3},
        "capturas_2005": {
            "sheets": [
                {"hoja": "16.b (2)", "l1": 4, "data_start": 7,
                 "categoria": "Caza Mayor"},
                {"hoja": "16.b (3)", "l1": 4, "data_start": 7,
                 "categoria": "Caza Menor de Mamíferos"},
                {"hoja": "16.b (4)", "l1": 4, "data_start": 7,
                 "categoria": "Caza Menor de Aves"},
            ],
        },
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS DE PARSEO
# ─────────────────────────────────────────────────────────────────────────────

# Umbral para descartar textos largos en celdas que no son nombres de CCAA ni
# provincia (notas al pie, cabeceras repetidas, disclaimers legales).
# El nombre de CCAA/provincia más largo en España tiene ~35 caracteres.
MAX_CELL_LEN = 60


def _cell(row, col):
    """Extrae una celda como string limpio: strip + NaN → ''."""
    v = row.iloc[col]
    if pd.notna(v):
        s = str(v).strip()
        return "" if s == "nan" else s
    return ""


def _safe_num(v):
    """Convierte a float; None si no es numérico."""
    try:
        f = float(v)
        return None if np.isnan(f) else f
    except (TypeError, ValueError):
        return None


def _build_col_map(row_l1, row_l2, cat_overrides=None):
    """
    Construye mapping col_idx → (categoria_norm, especie_norm) a partir de
    las dos filas de cabecera.  Ignora columnas de totales y de IDs.

    row_l1: lista con valores de la fila de categorías (forward-fill aplicado aquí)
    row_l2: lista con valores de la fila de especies
    cat_overrides: dict {col_idx: categoria_str} para inyectar categorías faltantes
                   antes del forward-fill (útil cuando el encabezado de categoría
                   está ausente en L1, como en la hoja PRODUCCIÓN de 2020).
    """
    SKIP_IDS = {"comunidad autónoma", "comunidad autonoma", "provincia",
                "tipo de procedencia", "tipo suelta", "ccaa", "cc.aa.", "datos",
                "número de capturas", "numero de capturas", "número", "numero"}

    # Aplicar overrides sobre una copia mutable de L1
    if cat_overrides:
        row_l1 = list(row_l1)
        for col_idx, cat_val in cat_overrides.items():
            if col_idx < len(row_l1):
                row_l1[col_idx] = cat_val

    # Forward-fill de categorías
    cat_ff = []
    last_cat = None
    for v in row_l1:
        s = str(v).strip() if pd.notna(v) else ""
        if s and s.lower() != "nan":
            last_cat = s
        cat_ff.append(last_cat)

    col_map = {}
    for i, (cat, esp) in enumerate(zip(cat_ff, row_l2)):
        cat_s = str(cat).strip() if cat else ""
        esp_s = str(esp).strip() if pd.notna(esp) else ""

        if not cat_s or cat_s.lower() == "nan":
            continue
        if cat_s.lower().startswith("total") or cat_s.lower() == "total general":
            continue
        if cat_s.lower() in SKIP_IDS:
            continue
        if not esp_s or esp_s.lower() == "nan":
            continue
        if esp_s.lower() in SKIP_IDS:
            continue

        col_map[i] = (norm_cat(cat_s), norm_esp(esp_s))

    return col_map


def _iter_data_rows(raw, data_start):
    """Itera las filas de datos ignorando las completamente vacías."""
    for _, row in raw.iloc[data_start:].iterrows():
        vals = [v for v in row if pd.notna(v) and str(v).strip() not in ("", "nan")]
        if vals:
            yield row


def _extract_ids_and_species(raw, data_start, col_map, id_col_names, col_id_start=1):
    """
    Extrae filas en formato largo a partir de un DataFrame en bruto.

    id_col_names: ["ccaa"] | ["ccaa","provincia"] | ["ccaa","provincia","tipo_procedencia"]
    col_id_start: índice de la primera columna de IDs (1 para xlsx, 0 para xls).
    Aplica forward-fill a todos los IDs.
    Devuelve lista de dicts.
    """
    n_ids = len(id_col_names)
    id_vals = [None] * n_ids
    rows_out = []

    for row in _iter_data_rows(raw, data_start):
        # Leer IDs crudos de esta fila
        raw_ids = []
        for k in range(n_ids):
            col_idx = k + col_id_start
            v = _cell(row, col_idx)
            raw_ids.append(v)

        # Filas de subtotal CCAA con etiqueta descriptiva (2008 sueltas):
        # "Nº ejemplares en sueltas de la Administracion",
        # "Kg ejemplares en sueltas de la Administración Andalucía", etc.
        if raw_ids[0].lower().startswith(("nº", "n°", "kg")):
            continue

        # Filas de total de CCAA: "Total Andalucía" en col1 → reset completo
        if raw_ids[0].lower().startswith("total"):
            id_vals = [None] * n_ids
            # Total global → fin de la tabla; parar para no procesar tablas
            # secundarias de pesca que siguen en la misma hoja.
            # Incluye "Total Nº ejemplares..." y "Total Kg ejemplares..." de 2008.
            s_low = raw_ids[0].lower()
            if (s_low in ("total", "total general", "total nacional")
                    or "nº" in s_low or "n°" in s_low or "kg" in s_low):
                break
            continue
        # Filas de total de provincia: "Total Almería" en col2+ → saltar sin resetear CCAA
        if any(v.lower().startswith("total") for v in raw_ids[1:] if v):
            continue

        # Forward-fill: actualizar solo los IDs que tengan valor en esta fila
        for k, v in enumerate(raw_ids):
            if v:
                id_vals[k] = v

        # Si no tenemos CCAA pero sí provincia, inferir CCAA de PROV_TO_CCAA
        if not id_vals[0] and n_ids >= 2 and id_vals[1]:
            prov_norm = norm_prov(id_vals[1])
            inferred = PROV_TO_CCAA.get(prov_norm)
            if inferred:
                id_vals[0] = inferred

        # Saltar si aún no tenemos CCAA
        if not id_vals[0]:
            continue

        # Normalizar IDs
        ids_norm = list(id_vals)
        ids_norm[0] = norm_ccaa(id_vals[0])
        if n_ids >= 2 and id_vals[1]:
            ids_norm[1] = norm_prov(id_vals[1])

        # Extraer valores numéricos por especie
        for col_idx, (cat, esp) in col_map.items():
            val = _safe_num(row.iloc[col_idx]) if col_idx < len(row) else None
            if val is None:
                continue
            entry = dict(zip(id_col_names, ids_norm))
            entry["categoria"] = cat
            entry["especie"]   = esp
            entry["valor"]     = val
            rows_out.append(entry)

    return rows_out


# ─────────────────────────────────────────────────────────────────────────────
# PARSERS POR TIPO DE HOJA
# ─────────────────────────────────────────────────────────────────────────────

def parse_licencias_ccaa(xl, cfg):
    """
    Tabla a nivel de CCAA (disponible todos los años).
    Salida: ccaa | licencias_expedidas | importe_expedidas |
            lic_interautonómicas | importe_interauton | lic_vigentes_anteriores
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    rows = []
    for row in _iter_data_rows(raw, l1 + 1):
        ccaa = _cell(row, 1)
        if not ccaa or is_total(ccaa) or len(ccaa) > MAX_CELL_LEN:
            continue
        rows.append({
            "ccaa":                    norm_ccaa(ccaa),
            "licencias_expedidas":     _safe_num(row.iloc[2]),
            "importe_expedidas":       _safe_num(row.iloc[3]),
            "lic_interautonómicas":    _safe_num(row.iloc[4]),
            "importe_interauton":      _safe_num(row.iloc[5]),
            "lic_vigentes_anteriores": _safe_num(row.iloc[6]),
        })
    # Deduplicar por CCAA conservando la primera ocurrencia (la tabla de caza
    # siempre precede a tablas secundarias de pesca que usan las mismas CCAA)
    seen = set()
    rows_dedup = []
    for r in rows:
        if r["ccaa"] not in seen:
            seen.add(r["ccaa"])
            rows_dedup.append(r)
    return pd.DataFrame(rows_dedup)


def parse_licencias_prov(xl, cfg):
    """
    Tabla a nivel de provincia (años 2010-2015).
    cfg keys obligatorias: hoja, l1, col_ccaa, col_prov, col_lic, col_importe, col_vigentes
    Salida: ccaa | provincia | licencias_expedidas | importe_expedidas | lic_vigentes_anteriores

    Algunas CCAA no desglosan por provincia y solo tienen datos en la fila
    "Total <CCAA>".  En ese caso se emite una fila con provincia = "(Total CCAA)".
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1     = cfg["l1"]
    col_c  = cfg["col_ccaa"]
    col_p  = cfg["col_prov"]
    col_l  = cfg["col_lic"]
    col_i  = cfg["col_importe"]
    col_v  = cfg["col_vigentes"]

    rows = []
    ccaa_actual = None
    # Índice donde empiezan las filas de la CCAA actual
    ccaa_start_idx = len(rows)
    # Dato pendiente de CCAA sin desglose provincial: se emite solo si no
    # aparecen filas de provincia para esa CCAA.
    ccaa_pending = None
    for row in _iter_data_rows(raw, l1 + 1):
        c0 = _cell(row, col_c)
        c1 = _cell(row, col_p)

        # Saltar filas de cabecera que se repiten en xls (ej. 'CC.AA.', 'comunidad autónoma')
        HEADER_VALS = {"cc.aa.", "comunidad autónoma", "comunidad autonoma",
                       "c.c.a.a.", "ccaa"}
        if c0.lower() in HEADER_VALS or c1.lower() in ("provincia", "province"):
            continue
        # Saltar notas de años anteriores como "Álava (2010)" o "Asturias (2007)"
        if c0 and "(" in c0 and c0.rstrip().endswith(")"):
            interior = c0.rstrip().rsplit("(", 1)[-1].rstrip(")")
            if interior.isdigit() and len(interior) == 4:
                continue
        if c0 and not is_total(c0) and len(c0) < MAX_CELL_LEN:
            # Flush pendiente de la CCAA anterior si no se emitió
            if ccaa_pending is not None:
                rows.append(ccaa_pending)
                ccaa_pending = None
            ccaa_actual = norm_ccaa(c0)
            ccaa_start_idx = len(rows)
            # Fila CCAA sin provincia pero con dato numérico → guardar como
            # pendiente; solo se emite si no aparecen provincias debajo.
            if not c1:
                lic_val = _safe_num(row.iloc[col_l])
                if lic_val is not None:
                    ccaa_pending = {
                        "ccaa":                    ccaa_actual,
                        "provincia":               "(Total CCAA)",
                        "licencias_expedidas":     lic_val,
                        "importe_expedidas":       _safe_num(row.iloc[col_i]),
                        "lic_vigentes_anteriores": _safe_num(row.iloc[col_v]),
                    }
        if (c0 and is_total(c0)) or (c1 and is_total(c1)):
            s_low = c0.lower()
            if s_low in ("total", "total general", "total nacional"):
                break
            # Fila "Total <CCAA>": si todas las provincias de esta CCAA
            # son NaN, usar el total como fila resumen.
            if ccaa_actual:
                ccaa_rows = rows[ccaa_start_idx:]
                all_nan = not ccaa_rows or all(
                    r["licencias_expedidas"] is None or
                    (isinstance(r["licencias_expedidas"], float) and
                     pd.isna(r["licencias_expedidas"]))
                    for r in ccaa_rows)
                total_lic = _safe_num(row.iloc[col_l])
                if all_nan and total_lic is not None:
                    # Eliminar filas de provincia vacías y poner una fila resumen
                    del rows[ccaa_start_idx:]
                    ccaa_pending = None  # descartar pendiente, usamos el total
                    rows.append({
                        "ccaa":                    ccaa_actual,
                        "provincia":               "(Total CCAA)",
                        "licencias_expedidas":     total_lic,
                        "importe_expedidas":       _safe_num(row.iloc[col_i]),
                        "lic_vigentes_anteriores": _safe_num(row.iloc[col_v]),
                    })
                else:
                    # Había provincias con datos → descartar pendiente
                    ccaa_pending = None
            continue
        if not c1 or not ccaa_actual or len(c1) > MAX_CELL_LEN:
            continue

        # Hay provincia → descartar el pendiente de CCAA sin desglose
        ccaa_pending = None
        rows.append({
            "ccaa":                    ccaa_actual,
            "provincia":               norm_prov(c1),
            "licencias_expedidas":     _safe_num(row.iloc[col_l]),
            "importe_expedidas":       _safe_num(row.iloc[col_i]),
            "lic_vigentes_anteriores": _safe_num(row.iloc[col_v]),
        })
    # Flush pendiente de la última CCAA
    if ccaa_pending is not None:
        rows.append(ccaa_pending)
    return pd.DataFrame(rows)


def parse_capturas(xl, cfg):
    """
    Salida: ccaa | provincia | categoria | especie | n_capturas
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    col_id_start = cfg.get("col_id_start", 1)
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))
    rows = _extract_ids_and_species(raw, l1 + 2, col_map, ["ccaa", "provincia"],
                                    col_id_start=col_id_start)
    df = pd.DataFrame(rows).rename(columns={"valor": "n_capturas"})
    return df


def parse_sueltas(xl, cfg):
    """
    Salida: ccaa | provincia | tipo_procedencia | categoria | especie | n_sueltas
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    col_id_start = cfg.get("col_id_start", 1)
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))
    rows = _extract_ids_and_species(raw, l1 + 2, col_map,
                                    ["ccaa", "provincia", "tipo_procedencia"],
                                    col_id_start=col_id_start)
    df = pd.DataFrame(rows).rename(columns={"valor": "n_sueltas"})
    if "tipo_procedencia" in df.columns:
        df["tipo_procedencia"] = df["tipo_procedencia"].apply(_norm_tipo_procedencia)
    return df


def parse_produccion(xl, cfg):
    """
    Salida: ccaa | provincia | categoria | especie | n_produccion
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    col_id_start = cfg.get("col_id_start", 1)
    col_map = _build_col_map(
        list(raw.iloc[l1]), list(raw.iloc[l1 + 1]),
        cat_overrides=cfg.get("cat_overrides"),
    )
    rows = _extract_ids_and_species(raw, l1 + 2, col_map, ["ccaa", "provincia"],
                                    col_id_start=col_id_start)
    df = pd.DataFrame(rows).rename(columns={"valor": "n_produccion"})
    return df


def _norm_tipo_procedencia(s):
    """Normaliza el tipo de procedencia de sueltas a forma canónica."""
    sl = str(s).strip().lower()
    if "administra" in sl:
        return "ADMINISTRACIÓN"
    if "otras" in sl:
        return "OTRAS PROCEDENCIAS"
    return str(s).strip()


# Normalización conservadora de tipo_coto: solo unificar case del mismo concepto
TIPO_COTO_NORM = {
    # ALL CAPS → Title case (2013+ usa mayúsculas, 2006-2012 usa title case)
    "coto privado de caza":                 "Coto privado de caza",
    "coto municipal":                       "Coto municipal",
    "coto local (coto municipal)":          "Coto municipal",
    "coto social":                          "Coto social",
    "coto social (exclusivamente)":         "Coto social",
    "coto deportivo":                       "Coto deportivo",
    "cotos deportivos locales":             "Coto deportivo local",
    "coto regional o autonómico":           "Coto regional / autonómico",
    "coto regional":                        "Coto regional / autonómico",
    "coto regional de caza / coto municipal": "Coto regional / municipal",
    "coto social / coto deportivo":         "Coto social / deportivo",
    "cotos intensivos de caza":             "Coto intensivo de caza",
    "explotaciones cinegéticas comerciales": "Explotaciones cinegéticas comerciales",
    "reserva de caza":                      "Reserva de caza",
    "zona de caza controlada":              "Zona de caza controlada",
    "refugio de caza / fauna":              "Refugio de caza / fauna",
    "refugio de caza":                      "Refugio de caza",
    "refugio de fauna":                     "Refugio de fauna",
    "terreno cercado":                      "Terreno cercado",
    "terrenos cercados":                    "Terreno cercado",
    "cercado / vallado":                    "Cercado / vallado",
    "vedado de caza":                       "Vedado de caza",
    "zonas de seguridad":                   "Zona de seguridad",
    "terrenos no cinegéticos":              "Terrenos no cinegéticos",
    "otros":                                "Otros terrenos",
}

def norm_tipo_coto(v):
    """Normaliza tipo de coto: unifica case, mantiene compuestos tal cual."""
    s = str(v).strip()
    return TIPO_COTO_NORM.get(s.lower(), s)


def _norm_metrica(s):
    """Normaliza la etiqueta de métrica a 'nº' o 'sup (ha)'."""
    sl = s.lower().strip()
    if sl in ("nº", "n°") or "número" in sl or "numero" in sl or sl.startswith("n°") or sl.startswith("nº"):
        return "nº"
    elif "superficie" in sl or ("ha" in sl and "renta" not in sl):
        return "sup (ha)"
    return sl


def parse_terrenos(xl, cfg):
    """
    Salida: ccaa | provincia | tipo_coto | n_cotos | sup_ha

    Soporta dos estructuras:
    - xlsx (2014+): dos filas de cabecera (categorías y tipos de coto)
    - xls (2010-2013): una sola fila de cabecera con los tipos directamente
      → activar con cfg["single_header"] = True y cfg["col_id_start"] = 0
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1            = cfg["l1"]
    col_id_start  = cfg.get("col_id_start", 1)
    single_header = cfg.get("single_header", False)

    if single_header:
        # xls: única fila de cabecera con tipos de coto directamente
        tipos_row  = list(raw.iloc[l1])
        data_start = l1 + 1
        skip_cols  = set()
        for i, v in enumerate(tipos_row):
            s = str(v).strip().lower() if pd.notna(v) else ""
            if "total" in s:
                skip_cols.add(i)
    else:
        # xlsx: dos filas de cabecera
        tipos_row   = list(raw.iloc[l1 + 1])
        totales_row = list(raw.iloc[l1])
        skip_cols   = set()
        for i, v in enumerate(totales_row):
            s = str(v).strip().lower() if pd.notna(v) else ""
            if s and "total" in s:
                skip_cols.add(i)
        data_start = l1 + 2

    # ID labels a excluir de los tipos de coto
    SKIP_LABELS = {"nan", "", "comunidad autónoma", "comunidad autonoma",
                   "cc.aa.", "ccaa", "provincia", "valores", "datos"}
    tipos_coto = []
    for i, v in enumerate(tipos_row):
        if i in skip_cols:
            continue
        s = str(v).strip() if pd.notna(v) else ""
        if s and s.lower() not in SKIP_LABELS:
            tipos_coto.append((i, s))

    long_rows   = []
    ccaa_actual = None
    prov_actual = None

    for row in _iter_data_rows(raw, data_start):
        c1 = _cell(row, col_id_start)
        c2 = _cell(row, col_id_start + 1)
        c3 = _cell(row, col_id_start + 2)

        if single_header:
            # xls: filtros más simples (no hay "nº CCAA" ni "SUP (ha) CCAA")
            if c1.lower().startswith("total") or c2.lower().startswith("total"):
                continue
            if c1.startswith("*") or (len(c1) > MAX_CELL_LEN and c3 == ""):
                continue
        else:
            # xlsx: subtotales de CCAA llevan prefijo "nº " / "SUP (ha) " en c1
            if c1.lower().startswith(("nº ", "sup (ha)", "total", "*")):
                continue
            if c1.startswith("*") or (len(c1) > MAX_CELL_LEN and c3 == ""):
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
        if ccaa_actual is None:
            continue

        for col_i, tipo in tipos_coto:
            val = _safe_num(row.iloc[col_i]) if col_i < len(row) else None
            if val is None:
                continue
            long_rows.append({
                "ccaa":      ccaa_actual,
                "provincia": prov_actual,
                "tipo_coto": tipo,
                "metrica":   metrica,
                "valor":     val,
            })

    if not long_rows:
        return pd.DataFrame(columns=["ccaa", "provincia", "tipo_coto", "n_cotos", "sup_ha"])

    df_long = pd.DataFrame(long_rows)
    df_piv = (
        df_long
        .pivot_table(index=["ccaa", "provincia", "tipo_coto"],
                     columns="metrica", values="valor", aggfunc="first")
        .reset_index()
    )
    df_piv.columns.name = None
    df_piv = df_piv.rename(columns={"nº": "n_cotos", "sup (ha)": "sup_ha"})
    for c in ("n_cotos", "sup_ha"):
        if c not in df_piv.columns:
            df_piv[c] = None

    return df_piv[["ccaa", "provincia", "tipo_coto", "n_cotos", "sup_ha"]]


def parse_sueltas_flat(xl, cfg):
    """
    Parser para sueltas con cabecera plana (2007): una sola fila de cabecera con
    especies directamente (sin categorías).  La columna 'Datos' (col 2) contiene
    el tipo de procedencia.  Categorías se infieren de ESPECIE_A_CATEGORIA.
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1  = cfg["l1"]
    col_id_start = cfg.get("col_id_start", 0)

    # Leer nombres de especie del header (row l1)
    header = list(raw.iloc[l1])
    SKIP = {"nan", "", "comunidad autonoma", "comunidad autónoma", "provincia",
            "datos", "total", "total general"}
    col_map = {}
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() in SKIP or "total" in s.lower():
            continue
        if i < col_id_start + 3:  # cols 0-2 son CCAA, PROV, DATOS
            continue
        esp = norm_esp(s)
        cat = ESPECIE_A_CATEGORIA.get(esp, "")
        if not cat:
            continue
        col_map[i] = (cat, esp)

    # Parsear filas de datos
    rows_out = []
    ccaa_act = None
    prov_act = None
    for row in _iter_data_rows(raw, l1 + 1):
        c0 = _cell(row, col_id_start)
        c1 = _cell(row, col_id_start + 1)
        c2 = _cell(row, col_id_start + 2)

        if c0.lower().startswith("total"):
            s_low = c0.lower()
            if s_low in ("total", "total general", "total nacional"):
                break
            ccaa_act = None
            prov_act = None
            continue

        if c0:
            ccaa_act = norm_ccaa(c0)
        if c1:
            prov_act = norm_prov(c1)
        if not ccaa_act or not c2:
            continue

        tipo = _norm_tipo_procedencia(c2)

        for col_i, (cat, esp) in col_map.items():
            val = _safe_num(row.iloc[col_i]) if col_i < len(row) else None
            if val is None:
                continue
            rows_out.append({
                "ccaa": ccaa_act,
                "provincia": prov_act,
                "tipo_procedencia": tipo,
                "categoria": cat,
                "especie": esp,
                "n_sueltas": val,
            })

    return pd.DataFrame(rows_out)


def parse_terrenos_paired(xl, cfg):
    """
    Parser para terrenos con columnas pareadas (2007): cada tipo de coto ocupa
    2 columnas consecutivas: nº y superficie.  Fila l1 tiene los nombres de tipo
    y fila l1+1 tiene 'nº'/'superficie (Ha)' alternados.
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1  = cfg["l1"]
    col_id_start = cfg.get("col_id_start", 0)

    # Leer tipos de coto de la fila l1 (solo en columnas pares desde col_id_start+2)
    header_l1 = list(raw.iloc[l1])
    tipos = []  # [(col_nº, col_sup, nombre_tipo)]
    for i, v in enumerate(header_l1):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() == "nan":
            continue
        if i < col_id_start + 2:
            continue
        if "total" in s.lower():
            continue
        tipos.append((i, i + 1, s))

    rows_out = []
    ccaa_act = None
    prov_act = None

    for row in _iter_data_rows(raw, l1 + 2):
        c0 = _cell(row, col_id_start)
        c1 = _cell(row, col_id_start + 1)

        if c0.lower().startswith("total"):
            s_low = c0.lower()
            if s_low in ("total", "total general", "total nacional"):
                break
            ccaa_act = None
            prov_act = None
            continue
        if c1.lower().startswith("total"):
            continue

        if c0 and not is_total(c0) and len(c0) < MAX_CELL_LEN:
            ccaa_act = norm_ccaa(c0)
        if c1 and not is_total(c1):
            prov_act = norm_prov(c1)

        if not ccaa_act or not prov_act:
            continue

        for col_n, col_s, tipo in tipos:
            n_val = _safe_num(row.iloc[col_n]) if col_n < len(row) else None
            s_val = _safe_num(row.iloc[col_s]) if col_s < len(row) else None
            if n_val is None and s_val is None:
                continue
            rows_out.append({
                "ccaa": ccaa_act,
                "provincia": prov_act,
                "tipo_coto": tipo,
                "n_cotos": n_val,
                "sup_ha": s_val,
            })

    if not rows_out:
        return pd.DataFrame(columns=["ccaa", "provincia", "tipo_coto", "n_cotos", "sup_ha"])
    return pd.DataFrame(rows_out)[["ccaa", "provincia", "tipo_coto", "n_cotos", "sup_ha"]]


def _prov_to_ccaa(prov_norm):
    """Devuelve la CCAA para una provincia ya normalizada."""
    return PROV_TO_CCAA.get(prov_norm, "")


def parse_capturas_prov(xl, cfg):
    """
    Parser para capturas con formato provincial (2006): provincia en col 0,
    'Datos' en col 1, especies en cols 2+.  Filtra solo filas 'Número de capturas'.
    2 filas de cabecera (categorías + especies).
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1  = cfg["l1"]
    col_map = _build_col_map(list(raw.iloc[l1]), list(raw.iloc[l1 + 1]))

    rows_out = []
    prov_act = None
    for row in _iter_data_rows(raw, l1 + 2):
        c0 = _cell(row, 0)
        c1 = _cell(row, 1)

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue

        if c0:
            prov_act = norm_prov(c0)

        # Solo filas de número de capturas
        if "número de capturas" not in c1.lower() and "numero de capturas" not in c1.lower():
            continue
        if not prov_act:
            continue

        ccaa = _prov_to_ccaa(prov_act)
        if not ccaa:
            continue

        for col_i, (cat, esp) in col_map.items():
            val = _safe_num(row.iloc[col_i]) if col_i < len(row) else None
            if val is None:
                continue
            rows_out.append({
                "ccaa": ccaa,
                "provincia": prov_act,
                "categoria": cat,
                "especie": esp,
                "n_capturas": val,
            })

    return pd.DataFrame(rows_out)


def parse_sueltas_prov(xl, cfg):
    """
    Parser para sueltas con formato provincial (2006): provincia en col 0,
    'Datos' en col 1 (tipo procedencia), especies en cols 2+.
    Single header (no categories).  Categorías inferidas de ESPECIE_A_CATEGORIA.
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1  = cfg["l1"]

    header = list(raw.iloc[l1])
    SKIP = {"nan", "", "provincia", "datos", "total", "total general"}
    col_map = {}
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() in SKIP or "total" in s.lower():
            continue
        if i < 2:
            continue
        esp = norm_esp(s)
        cat = ESPECIE_A_CATEGORIA.get(esp, "")
        if not cat:
            continue
        col_map[i] = (cat, esp)

    rows_out = []
    prov_act = None
    for row in _iter_data_rows(raw, l1 + 1):
        c0 = _cell(row, 0)
        c1 = _cell(row, 1)

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue

        if c0:
            prov_act = norm_prov(c0)
        if not prov_act or not c1:
            continue

        ccaa = _prov_to_ccaa(prov_act)
        if not ccaa:
            continue

        tipo = _norm_tipo_procedencia(c1)

        for col_i, (cat, esp) in col_map.items():
            val = _safe_num(row.iloc[col_i]) if col_i < len(row) else None
            if val is None:
                continue
            rows_out.append({
                "ccaa": ccaa,
                "provincia": prov_act,
                "tipo_procedencia": tipo,
                "categoria": cat,
                "especie": esp,
                "n_sueltas": val,
            })

    return pd.DataFrame(rows_out)


def parse_produccion_prov(xl, cfg):
    """
    Parser para producción con formato provincial (2006): provincia en col 0,
    'Datos' en col 1, especies en cols 2+.  Filtra solo filas de número.
    2 filas de cabecera (categorías + especies); hay una fila extra (row 4) tipo
    'descripcion'/'especie' que hay que saltar.
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1_cat = cfg["l1_cat"]   # fila de categorías
    l1_esp = cfg["l1_esp"]   # fila de especies
    col_map = _build_col_map(list(raw.iloc[l1_cat]), list(raw.iloc[l1_esp]))

    rows_out = []
    prov_act = None
    for row in _iter_data_rows(raw, l1_esp + 1):
        c0 = _cell(row, 0)
        c1 = _cell(row, 1)

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue

        if c0:
            prov_act = norm_prov(c0)

        # Solo filas de número de ejemplares producidos
        if "número" not in c1.lower() and "numero" not in c1.lower():
            continue
        if "producido" not in c1.lower() and "producción" not in c1.lower():
            continue
        if not prov_act:
            continue

        ccaa = _prov_to_ccaa(prov_act)
        if not ccaa:
            continue

        for col_i, (cat, esp) in col_map.items():
            val = _safe_num(row.iloc[col_i]) if col_i < len(row) else None
            if val is None:
                continue
            rows_out.append({
                "ccaa": ccaa,
                "provincia": prov_act,
                "categoria": cat,
                "especie": esp,
                "n_produccion": val,
            })

    return pd.DataFrame(rows_out)


def parse_terrenos_prov(xl, cfg):
    """
    Parser para terrenos con formato provincial (2006): provincia en col 0,
    'Datos' en col 1, tipos de coto en cols 2+.
    Filtra filas de nº y superficie.  Single header row.
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]

    header = list(raw.iloc[l1])
    SKIP = {"nan", "", "provincia", "datos", "total", "total general"}
    tipos_coto = []
    for i, v in enumerate(header):
        s = str(v).strip() if pd.notna(v) else ""
        if not s or s.lower() in SKIP or "total" in s.lower():
            continue
        if i < 2:
            continue
        tipos_coto.append((i, s))

    rows_out = []
    prov_act = None
    for row in _iter_data_rows(raw, l1 + 1):
        c0 = _cell(row, 0)
        c1 = _cell(row, 1)

        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            prov_act = None
            continue

        if c0:
            prov_act = norm_prov(c0)

        metrica = _norm_metrica(c1)
        if metrica not in ("nº", "sup (ha)"):
            continue
        if not prov_act:
            continue

        ccaa = _prov_to_ccaa(prov_act)
        if not ccaa:
            continue

        for col_i, tipo in tipos_coto:
            val = _safe_num(row.iloc[col_i]) if col_i < len(row) else None
            if val is None:
                continue
            rows_out.append({
                "ccaa": ccaa,
                "provincia": prov_act,
                "tipo_coto": tipo,
                "metrica": metrica,
                "valor": val,
            })

    if not rows_out:
        return pd.DataFrame(columns=["ccaa", "provincia", "tipo_coto", "n_cotos", "sup_ha"])

    df_long = pd.DataFrame(rows_out)
    df_piv = (
        df_long
        .pivot_table(index=["ccaa", "provincia", "tipo_coto"],
                     columns="metrica", values="valor", aggfunc="first")
        .reset_index()
    )
    df_piv.columns.name = None
    df_piv = df_piv.rename(columns={"nº": "n_cotos", "sup (ha)": "sup_ha"})
    for c in ("n_cotos", "sup_ha"):
        if c not in df_piv.columns:
            df_piv[c] = None
    return df_piv[["ccaa", "provincia", "tipo_coto", "n_cotos", "sup_ha"]]


def parse_licencias_prov_flat(xl, cfg):
    """
    Parser para licencias con formato provincial plano (2005-2006):
    provincia en col 0, datos de caza y pesca en columnas siguientes.
    Sin columna de CCAA — se infiere de PROV_TO_CCAA.
    Incluye filas de totales de CCAA que se filtran.
    """
    raw = xl.parse(cfg["hoja"], header=None)
    l1 = cfg["l1"]
    col_prov = cfg["col_prov"]
    col_lic  = cfg["col_lic"]
    col_imp  = cfg["col_importe"]
    col_vig  = cfg.get("col_vigentes")

    rows_out = []
    for row in _iter_data_rows(raw, l1 + 1):
        c0 = _cell(row, col_prov)
        if not c0:
            continue
        if c0.lower().startswith("total"):
            if c0.lower() in ("total", "total general", "total nacional"):
                break
            continue
        # Saltar cabeceras intermedias
        if c0.lower() in ("provincia", "provincias y", "comunidades autónomas",
                          "comunidad autonoma", "comunidad autónoma"):
            continue

        prov = norm_prov(c0)
        ccaa = _prov_to_ccaa(prov)
        if not ccaa:
            # Puede ser un nombre de CCAA usado como subtotal
            continue

        lic_val = _safe_num(row.iloc[col_lic])
        imp_val = _safe_num(row.iloc[col_imp])
        vig_val = _safe_num(row.iloc[col_vig]) if col_vig is not None else None

        if lic_val is None and imp_val is None:
            continue

        rows_out.append({
            "ccaa": ccaa,
            "provincia": prov,
            "licencias_expedidas": lic_val,
            "importe_expedidas": imp_val,
            "lic_vigentes_anteriores": vig_val,
        })

    return pd.DataFrame(rows_out)


def parse_capturas_2005(xl, cfg):
    """
    Parser para capturas 2005: distribuidas en 3 hojas (Caza Mayor, Caza Menor
    mamíferos, Caza Volátil).  Cada especie ocupa 3 columnas (Número, Peso, Valor);
    solo extraemos Número.  Province-only, sin CCAA.
    Cada hoja tiene header en row 4 (especies) y data empieza en row 7.
    """
    rows_out = []

    for sheet_cfg in cfg["sheets"]:
        hoja = sheet_cfg["hoja"]
        if hoja not in xl.sheet_names:
            continue
        raw = xl.parse(hoja, header=None)
        l1 = sheet_cfg["l1"]

        # Leer especies de row l1.  Cada especie ocupa 3 cols (Número, Peso, Valor).
        header = list(raw.iloc[l1])
        especies = []  # [(col_numero, especie_norm, categoria)]
        for i, v in enumerate(header):
            s = str(v).strip() if pd.notna(v) else ""
            if not s or s.lower() == "nan":
                continue
            if i < 1:  # col 0 = "Provincias y..."
                continue
            # Limpiar nombre científico entre paréntesis
            import re
            name_clean = re.sub(r"\s*\(.*?\)\s*", "", s).strip()
            if not name_clean or "total" in name_clean.lower():
                continue
            esp = norm_esp(name_clean)
            cat = sheet_cfg.get("categoria", ESPECIE_A_CATEGORIA.get(esp, ""))
            if not cat:
                continue
            # La columna actual es la de Número; las 2 siguientes son Peso y Valor
            especies.append((i, esp, cat))

        data_start = sheet_cfg.get("data_start", l1 + 3)  # normalmente row 7 para l1=4
        prov_act = None
        for row in _iter_data_rows(raw, data_start):
            c0 = _cell(row, 0)

            if c0.lower().startswith("total"):
                if c0.lower() in ("total", "total general", "total nacional"):
                    break
                prov_act = None
                continue

            if c0:
                prov_act = norm_prov(c0)
            if not prov_act:
                continue

            ccaa = _prov_to_ccaa(prov_act)
            if not ccaa:
                continue

            for col_i, esp, cat in especies:
                val = _safe_num(row.iloc[col_i]) if col_i < len(row) else None
                if val is None:
                    continue
                rows_out.append({
                    "ccaa": ccaa,
                    "provincia": prov_act,
                    "categoria": cat,
                    "especie": esp,
                    "n_capturas": val,
                })

    return pd.DataFrame(rows_out)


# ─────────────────────────────────────────────────────────────────────────────
# PARSEO POR AÑO
# ─────────────────────────────────────────────────────────────────────────────

PARSERS = {
    "licencias_ccaa":  parse_licencias_ccaa,
    "licencias_prov":  parse_licencias_prov,
    "capturas":        parse_capturas,
    "sueltas":         parse_sueltas,
    "produccion":      parse_produccion,
    "terrenos":        parse_terrenos,
}

# Parsers alternativos por clave especial en cfg
ALT_PARSERS = {
    "sueltas_flat":         ("sueltas",        parse_sueltas_flat),
    "terrenos_paired":      ("terrenos",       parse_terrenos_paired),
    "capturas_prov":        ("capturas",       parse_capturas_prov),
    "sueltas_prov":         ("sueltas",        parse_sueltas_prov),
    "produccion_prov":      ("produccion",     parse_produccion_prov),
    "terrenos_prov":        ("terrenos",       parse_terrenos_prov),
    "licencias_prov_flat":  ("licencias_prov", parse_licencias_prov_flat),
    "capturas_2005":        ("capturas",       parse_capturas_2005),
}

def _local_path(anio):
    url = CATALOGO[anio]
    ext = Path(urllib.parse.unquote(url)).suffix or ".xlsx"
    return RAW_DIR / f"caza_{anio}{ext}"


def parse_year(anio):
    """
    Lee el archivo de un año y devuelve dict tipo_hoja → DataFrame normalizado.
    Cada DataFrame ya incluye la columna 'anio'.
    """
    path = _local_path(anio)
    if not path.exists():
        print(f"  [!] {anio}: archivo no encontrado — ejecuta --descargar")
        return {}

    cfg_year = YEAR_CONFIG.get(anio)
    if cfg_year is None:
        print(f"  [!] {anio}: configuración pendiente, saltando.")
        return {}

    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        print(f"  [✗] {anio}: error al abrir ({e})")
        return {}

    resultado = {}

    # Recopilar tareas: (output_key, parser_fn, cfg, cfg_key)
    tareas = []
    for tipo, parser_fn in PARSERS.items():
        cfg = cfg_year.get(tipo)
        if cfg is not None:
            tareas.append((tipo, parser_fn, cfg, tipo))
    for alt_key, (output_key, parser_fn) in ALT_PARSERS.items():
        cfg = cfg_year.get(alt_key)
        if cfg is not None:
            tareas.append((output_key, parser_fn, cfg, alt_key))

    for output_key, parser_fn, cfg, cfg_key in tareas:
        # Para capturas_2005, no hay una sola hoja
        hoja = cfg.get("hoja")
        if hoja and hoja not in xl.sheet_names:
            print(f"  [!] {anio}/{cfg_key}: hoja '{hoja}' no encontrada "
                  f"(disponibles: {xl.sheet_names})")
            continue
        try:
            df = parser_fn(xl, cfg)
            df.insert(0, "anio", anio)
            resultado[output_key] = df
            print(f"  [✓] {anio}/{cfg_key}: {len(df)} filas")
        except Exception as e:
            print(f"  [✗] {anio}/{cfg_key}: {e}")
            import traceback; traceback.print_exc()

    return resultado


# ─────────────────────────────────────────────────────────────────────────────
# CONSTRUCCIÓN INCREMENTAL DE LA TABLA FINAL
# ─────────────────────────────────────────────────────────────────────────────

def _cualificar_otras(df):
    """Reemplaza 'Otras'/'Otros' por 'Otros Caza Mayor', 'Otras Caza Menor de Aves', etc."""
    MAPA = {
        ("Caza Mayor", "Otras"):                    "Otros Caza Mayor",
        ("Caza Mayor", "Otros"):                    "Otros Caza Mayor",
        ("Caza Menor de Aves", "Otras"):            "Otras Caza Menor de Aves",
        ("Caza Menor de Aves", "Otros"):            "Otras Caza Menor de Aves",
        ("Caza Menor de Mamíferos", "Otras"):       "Otras Caza Menor de Mamíferos",
        ("Caza Menor de Mamíferos", "Otros"):       "Otras Caza Menor de Mamíferos",
    }
    mask = df["especie"].isin(("Otras", "Otros"))
    if mask.any():
        df = df.copy()
        df.loc[mask, "especie"] = df.loc[mask].apply(
            lambda r: MAPA.get((r["categoria"], r["especie"]), r["especie"]), axis=1
        )
    return df


def build_final(años=None):
    """
    Itera los años de más nuevo a más antiguo.
    La tabla final de cada tipo crece con cada año procesado.
    Guarda los CSV en output/.
    """
    OUTPUT_DIR.mkdir(exist_ok=True)

    if años is None:
        años = sorted(YEAR_CONFIG.keys(), reverse=True)

    acumulado = {tipo: [] for tipo in PARSERS}

    for anio in años:
        print(f"\n── {anio} ──────────────────────────")
        datos = parse_year(anio)
        for tipo, df in datos.items():
            acumulado[tipo].append(df)

    print("\n── Guardando ──────────────────────")
    for tipo, frames in acumulado.items():
        if not frames:
            print(f"  [–] {tipo}: sin datos")
            continue
        # concat: sort=False preserva el orden de columnas del primero (año más nuevo)
        final = pd.concat(frames, ignore_index=True, sort=False)
        # Post-procesado
        if "especie" in final.columns and "categoria" in final.columns:
            final = _cualificar_otras(final)
        if "tipo_coto" in final.columns:
            final["tipo_coto"] = final["tipo_coto"].apply(norm_tipo_coto)
        out = OUTPUT_DIR / f"{tipo}.csv"
        final.to_csv(out, index=False, encoding="utf-8-sig")
        print(f"  [✓] {tipo}.csv  →  {len(final)} filas × {len(final.columns)} col")

    quality_check(acumulado)
    return acumulado



# ─────────────────────────────────────────────────────────────────────────────
# CONTROL DE CALIDAD
# ─────────────────────────────────────────────────────────────────────────────

# Referencia canónica (2023)
CCAA_CANON = {
    "Andalucía", "Aragón", "Asturias", "Baleares", "C. Valenciana",
    "Canarias", "Cantabria", "Castilla y León", "Castilla-La Mancha",
    "Cataluña", "Extremadura", "Galicia", "La Rioja", "Madrid",
    "Murcia", "Navarra", "País Vasco",
}

PROVINCIAS_CANON = {
    "Álava", "Albacete", "Alicante", "Almería", "Asturias", "Ávila",
    "Badajoz", "Baleares", "Barcelona", "Burgos", "Cáceres", "Cádiz",
    "Cantabria", "Castellón", "Ciudad Real", "Córdoba", "La Coruña",
    "Cuenca", "Gerona", "Granada", "Guadalajara", "Guipúzcoa", "Huelva",
    "Huesca", "Jaén", "La Rioja", "Las Palmas", "León", "Lérida",
    "Lugo", "Madrid", "Málaga", "Murcia", "Navarra", "Orense",
    "Palencia", "Pontevedra", "La Rioja", "Salamanca", "Segovia",
    "Sevilla", "Soria", "Tarragona", "Tenerife", "Teruel", "Toledo",
    "Valencia", "Valladolid", "Vizcaya", "Zamora", "Zaragoza",
    "Fuera de la C.A.",
    "(Total CCAA)",
}

ESPECIES_CANON = {
    "Caza Mayor":           {"Arruí", "Cabra asilvestrada", "Cabra montés", "Ciervo",
                             "Corzo", "Gamo", "Jabalí", "Lobo", "Muflón", "Rebeco",
                             "Otros Caza Mayor"},
    "Caza Menor de Aves":   {"Acuáticas y anátidas", "Acuáticas", "Anátidas",
                             "Avefría", "Becada", "Codorniz",
                             "Córvidos", "Estornino", "Faisán",
                             "Paloma", "Paloma torcaz", "Paloma zurita", "Perdiz",
                             "Tórtola común", "Otras Caza Menor de Aves",
                             "Tordo", "Zorzal", "Varias especies"},
    "Caza Menor de Mamíferos": {"Conejo", "Liebre", "Zorro",
                                "Otras Caza Menor de Mamíferos", "Turón"},
}
# Mapeo especie → categoría para formatos sin cabecera de categoría (2007 sueltas, 2005)
ESPECIE_A_CATEGORIA = {}
for _cat, _sps in ESPECIES_CANON.items():
    for _sp in _sps:
        # Algunas especies están en varias categorías (Otras);
        # priorizamos la primera asignación, excepciones manuales abajo
        if _sp not in ESPECIE_A_CATEGORIA:
            ESPECIE_A_CATEGORIA[_sp] = _cat
# Correcciones manuales para especies ambiguas:
# "Otras" bajo Caza Mayor se llama "Otros"; bajo Aves/Mamíferos se llama "Otras"
ESPECIE_A_CATEGORIA["Otros"] = "Caza Mayor"
TODAS_ESPECIES_CANON = {e for s in ESPECIES_CANON.values() for e in s}


def quality_check(acumulado: dict) -> None:
    """
    Ejecuta controles de calidad sobre las tablas unificadas y muestra un informe.
    Se llama automáticamente al final de build_final().

    Controles:
      1. Filas por año — detecta años con recuento anómalo
      2. CCAA no canónicas — nombres que no se normalizaron
      3. Provincias no canónicas — ídem
      4. Especies no canónicas — ídem
      5. Valores negativos en columnas numéricas
      6. Duplicados exactos (año + claves + especie)
      7. CCAA ausentes por año (las que tenían datos en 2023 pero faltan en años anteriores)
    """
    OK, WARN, ERR = "  ✓", "  ⚠", "  ✗"
    issues = 0

    print("\n" + "═" * 65)
    print("  CONTROL DE CALIDAD")
    print("═" * 65)

    # ── 1. Filas por año ─────────────────────────────────────────────────────
    for tipo, frames in acumulado.items():
        if not frames:
            continue
        df = pd.concat(frames, ignore_index=True, sort=False)
        counts = df.groupby("anio").size()
        media = counts.mean()
        print(f"\n[{tipo}] filas/año:")
        for anio, n in sorted(counts.items(), reverse=True):
            ratio = n / media
            flag = OK if 0.4 < ratio < 2.5 else WARN
            if flag == WARN:
                issues += 1
            print(f"  {flag}  {anio}: {n:5d}  {'← anómalo' if flag == WARN else ''}")

    # ── 2. CCAA no canónicas ─────────────────────────────────────────────────
    print(f"\n[CCAA no canónicas]")
    encontradas = False
    for tipo, frames in acumulado.items():
        if not frames:
            continue
        df = pd.concat(frames, ignore_index=True, sort=False)
        if "ccaa" not in df.columns:
            continue
        raras = sorted(set(df["ccaa"].dropna()) - CCAA_CANON)
        if raras:
            print(f"  {ERR}  {tipo}: {raras}")
            issues += len(raras)
            encontradas = True
    if not encontradas:
        print(f"  {OK}  ninguna")

    # ── 3. Provincias no canónicas ───────────────────────────────────────────
    print(f"\n[Provincias no canónicas]")
    encontradas = False
    for tipo, frames in acumulado.items():
        if not frames:
            continue
        df = pd.concat(frames, ignore_index=True, sort=False)
        if "provincia" not in df.columns:
            continue
        raras = sorted(set(df["provincia"].dropna()) - PROVINCIAS_CANON)
        if raras:
            print(f"  {ERR}  {tipo}: {raras}")
            issues += len(raras)
            encontradas = True
    if not encontradas:
        print(f"  {OK}  ninguna")

    # ── 4. Especies no canónicas ─────────────────────────────────────────────
    print(f"\n[Especies no canónicas]")
    encontradas = False
    for tipo in ("capturas", "sueltas", "produccion"):
        frames = acumulado.get(tipo, [])
        if not frames:
            continue
        df = pd.concat(frames, ignore_index=True, sort=False)
        if "especie" not in df.columns:
            continue
        raras = sorted(set(df["especie"].dropna()) - TODAS_ESPECIES_CANON)
        if raras:
            print(f"  {ERR}  {tipo}: {raras}")
            issues += len(raras)
            encontradas = True
    if not encontradas:
        print(f"  {OK}  ninguna")

    # ── 5. Valores negativos ──────────────────────────────────────────────────
    print(f"\n[Valores negativos]")
    encontrados = False
    num_cols = {
        "licencias_ccaa": ["licencias_expedidas", "importe_expedidas",
                           "lic_interautonómicas", "importe_interauton",
                           "lic_vigentes_anteriores"],
        "capturas":   ["n_capturas"],
        "sueltas":    ["n_sueltas"],
        "produccion": ["n_produccion"],
        "terrenos":   ["n_cotos", "sup_ha"],
    }
    for tipo, cols in num_cols.items():
        frames = acumulado.get(tipo, [])
        if not frames:
            continue
        df = pd.concat(frames, ignore_index=True, sort=False)
        for col in cols:
            if col not in df.columns:
                continue
            neg = (df[col] < 0).sum()
            if neg:
                print(f"  {ERR}  {tipo}.{col}: {neg} valores negativos")
                issues += 1
                encontrados = True
    if not encontrados:
        print(f"  {OK}  ninguno")

    # ── 6. Duplicados ─────────────────────────────────────────────────────────
    print(f"\n[Duplicados exactos]")
    keys = {
        "licencias_ccaa": ["anio", "ccaa"],
        "capturas":       ["anio", "ccaa", "provincia", "categoria", "especie"],
        "sueltas":        ["anio", "ccaa", "provincia", "tipo_procedencia",
                           "categoria", "especie"],
        "produccion":     ["anio", "ccaa", "provincia", "categoria", "especie"],
        "terrenos":       ["anio", "ccaa", "provincia", "tipo_coto"],
    }
    encontrados = False
    for tipo, key_cols in keys.items():
        frames = acumulado.get(tipo, [])
        if not frames:
            continue
        df = pd.concat(frames, ignore_index=True, sort=False)
        cols_pres = [c for c in key_cols if c in df.columns]
        n_dup = df.duplicated(subset=cols_pres).sum()
        if n_dup:
            print(f"  {ERR}  {tipo}: {n_dup} filas duplicadas")
            issues += 1
            encontrados = True
    if not encontrados:
        print(f"  {OK}  ninguno")

    # ── 7. CCAA ausentes por año ──────────────────────────────────────────────
    print(f"\n[CCAA ausentes por año en capturas]")
    frames = acumulado.get("capturas", [])
    if frames:
        df = pd.concat(frames, ignore_index=True, sort=False)
        ccaa_por_año = df.groupby("anio")["ccaa"].apply(set)
        ref = ccaa_por_año.iloc[0]  # año más reciente como referencia
        for anio, ccaas in sorted(ccaa_por_año.items(), reverse=True):
            ausentes = sorted(ref - ccaas)
            if ausentes:
                print(f"  {WARN}  {anio}: faltan {ausentes}")
                issues += 1
            else:
                print(f"  {OK}  {anio}: completo")

    # ── Resumen ───────────────────────────────────────────────────────────────
    print(f"\n{'═'*65}")
    if issues == 0:
        print(f"  ✓  Sin incidencias")
    else:
        print(f"  ⚠  {issues} incidencia(s) encontrada(s) — revisar los puntos marcados con ⚠/✗")
    print("═" * 65 + "\n")


# ─────────────────────────────────────────────────────────────────────────────
# DESCARGA
# ─────────────────────────────────────────────────────────────────────────────

def descargar(años=None, forzar=False):
    RAW_DIR.mkdir(exist_ok=True)
    if años is None:
        años = sorted(CATALOGO.keys(), reverse=True)

    session = requests.Session()
    session.headers["User-Agent"] = "Mozilla/5.0 (investigacion-caza/1.0)"

    for anio in tqdm(años, desc="Descargando"):
        dest = _local_path(anio)
        if dest.exists() and not forzar:
            print(f"  [OK] {anio} ya existe ({dest.name})")
            continue
        try:
            r = session.get(CATALOGO[anio], timeout=60, stream=True)
            r.raise_for_status()
            with open(dest, "wb") as f:
                for chunk in r.iter_content(65536):
                    f.write(chunk)
            print(f"  [✓] {anio} → {dest.name} ({dest.stat().st_size // 1024} KB)")
            time.sleep(0.5)
        except Exception as e:
            print(f"  [✗] {anio}: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Unificador de estadísticas de caza MITECO")
    parser.add_argument("--descargar",  action="store_true", help="Descarga los Excel del MITECO")
    parser.add_argument("--procesar",   action="store_true", help="Unifica los archivos descargados")
    parser.add_argument("--todo",       action="store_true", help="Descarga + procesa")
    parser.add_argument("--forzar",     action="store_true", help="Re-descarga aunque ya existan")
    parser.add_argument("--año", nargs="+", type=int, dest="años", help="Limitar a estos años")
    args = parser.parse_args()

    if not any([args.descargar, args.procesar, args.todo]):
        parser.print_help()
        sys.exit(0)

    if args.todo or args.descargar:
        descargar(args.años, forzar=args.forzar)
    if args.todo or args.procesar:
        build_final(args.años)
