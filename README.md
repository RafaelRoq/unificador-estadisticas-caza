# Estadísticas de Caza en España (2005-2023) — Dataset Unificado

Dataset abierto que unifica **19 años** de estadísticas oficiales de caza en España en CSVs listos para análisis. Los datos provienen del [Anuario de Estadística Forestal](https://www.miteco.gob.es/es/biodiversidad/estadisticas/forestal_anuarios_702702702702.html) del MITECO.

## Por qué existe este proyecto

La caza se practica en aproximadamente el 85% del territorio español y genera más de 20 millones de capturas anuales según las estadísticas oficiales. El MITECO recopila estos datos a través de la [Estadística Anual de Caza](https://www.miteco.gob.es/es/biodiversidad/estadisticas/est_anual_caza.html) (operación incluida en el Plan Estadístico Nacional) y los publica como parte del [Anuario de Estadística Forestal](https://www.miteco.gob.es/es/biodiversidad/estadisticas/forestal_anuarios_702702702702.html).

Sin embargo, la información se publica **año a año en archivos Excel independientes**, con formatos que han cambiado significativamente entre 2005 y 2023. El MITECO ofrece una [tabla resumen](https://www.miteco.gob.es/content/dam/miteco/es/biodiversidad/estadisticas/estadistica-caza/CAZA-Tablas-resumen-2005-2022.xlsx) que recopila datos por año en una hoja separada por cada uno, incluyendo desglose por CCAA. Este proyecto parte de los datos desglosados de cada anuario y los consolida en CSVs tabulados con formato homogéneo, columnas normalizadas y una fila por observación, pensados para su uso directo en análisis.

Un investigador que quiera estudiar tendencias a largo plazo con ese nivel de detalle necesitaría descargar 19 archivos Excel, entender la estructura interna de cada uno (que cambia entre años), resolver diferencias de nomenclatura y limpiar manualmente miles de filas. Este proyecto se encarga de la descarga, lectura, normalización y unificación de todos los anuarios en un conjunto de CSVs listos para análisis.

## Dataset

Los CSVs están en la carpeta [`output/`](output/) y se pueden descargar directamente, sin necesidad de ejecutar código.

### Licencias

Número de licencias de caza expedidas cada año. El MITECO cambió el formato de publicación en 2016: hasta 2015 desglosaba por provincia, y a partir de 2016 publica solo por comunidad autónoma.

| Archivo | Filas | Periodo | Descripción |
|---------|------:|---------|-------------|
| [`licencias_ccaa.csv`](output/licencias_ccaa.csv) | 136 | 2016-2023 | Licencias por CCAA, incluye licencias interautonómicas (válidas en todo el territorio nacional) |
| [`licencias_prov.csv`](output/licencias_prov.csv) | 484 | 2005-2015 | Licencias desglosadas por provincia |

### Capturas

Animales cazados, desglosados por comunidad autónoma, provincia, categoría (caza mayor, menor de aves, menor de mamíferos) y especie.

| Archivo | Filas | Periodo | Descripción |
|---------|------:|---------|-------------|
| [`capturas.csv`](output/capturas.csv) | 13.418 | 2005-2023 | Capturas por CCAA, provincia, categoría y especie |

### Sueltas

Animales soltados o repoblados con fines cinegéticos, clasificados por su procedencia: granja (criados en cautividad) o campo (capturados y trasladados). Incluye tanto el número de ejemplares como el peso en kg para algunas especies.

| Archivo | Filas | Periodo | Descripción |
|---------|------:|---------|-------------|
| [`sueltas.csv`](output/sueltas.csv) | 3.606 | 2006-2023 | Sueltas por tipo de procedencia, categoría y especie |

### Producción cinegética

Animales criados en granjas cinegéticas (explotaciones autorizadas para la cría de especies de caza). No son capturas ni sueltas, sino la producción de las granjas.

| Archivo | Filas | Periodo | Descripción |
|---------|------:|---------|-------------|
| [`produccion.csv`](output/produccion.csv) | 1.421 | 2006-2023 | Producción por categoría y especie |

### Terrenos cinegéticos

Superficie y número de cotos de caza (terrenos acotados donde se autoriza la actividad cinegética). Se clasifican por tipo: cotos privados, cotos sociales/deportivos, reservas de caza, zonas de caza controlada, etc.

| Archivo | Filas | Periodo | Descripción |
|---------|------:|---------|-------------|
| [`terrenos.csv`](output/terrenos.csv) | 3.078 | 2006-2023 | Terrenos por tipo de coto (nº de cotos y superficie en hectáreas) |

### Columnas

Cada tabla tiene un subconjunto de estas columnas:

| Columna | Tablas | Descripción |
|---------|--------|-------------|
| `anio` | todas | Año del anuario (2005-2023) |
| `ccaa` | todas | Comunidad Autónoma (17 valores normalizados) |
| `provincia` | capturas, sueltas, terrenos, licencias_prov | Provincia (52 valores + "Fuera de la C.A." + "(Total CCAA)" para CCAAs sin desglose) |
| `categoria` | capturas, sueltas, produccion | "Caza Mayor", "Caza Menor de Aves", "Caza Menor de Mamíferos" |
| `especie` | capturas, sueltas, produccion | Especie normalizada (ej. "Jabalí", "Conejo", "Perdiz") |

## Verificación

El repositorio incluye dos scripts independientes que comparan cada CSV contra los Excel originales del MITECO.

**`test_unitarios.py`** — Por cada año (2005-2023), lee el Excel original del MITECO con código independiente del pipeline, suma los valores por su cuenta, y comprueba que el total coincide con el CSV generado. Si el CSV de capturas indica que en 2019 se capturaron 20.404.957 piezas, el test abre el Excel de 2019, suma las celdas de capturas y verifica que da exactamente 20.404.957.

```
19 años OK  |  0 años con errores
✓ Todos los años verificados correctamente
```

**`auditoria.py`** — Cuatro niveles de comprobación, de más exhaustivo a más general:

```
VALOR A VALOR : 184 OK  |  0 FAIL
TOTALES       : 109 OK  |  0 FAIL
MUESTREO      : 177 OK  |  7 FAIL  (*)
INTEGRIDAD    :  18 OK  |  0 FAIL
```

- **VALOR A VALOR**: vuelve a leer cada Excel desde cero y compara **cada valor individual** con el CSV. Si las capturas de conejo en Cuenca en 2017 no coinciden exactamente con la celda original del Excel, falla.
- **TOTALES**: suma todas las filas del CSV de un año y compara con la fila "Total general" del Excel. Detecta filas perdidas o duplicadas.
- **MUESTREO**: elige 30 filas aleatorias del CSV, localiza cada una en el Excel original por su combinación de CCAA/provincia/especie, y compara el valor.
- **INTEGRIDAD**: comprueba que no hay campos identificativos vacíos, filas duplicadas, ni años ausentes.

(*) Los 7 fallos de MUESTREO son limitaciones del verificador para localizar ciertas filas en el Excel (CCAAs sin desglose provincial, una hoja con formato irregular en 2023), no errores en los datos — cubiertos al 100% por VALOR A VALOR y TOTALES.

Para reproducir las verificaciones (requiere los Excel en `raw_data/`, ver [Regenerar desde cero](#regenerar-desde-cero)):

```bash
python test_unitarios.py    # ~2 min
python auditoria.py         # ~5 min
```

## Filosofía: conservador con los datos

Principio estricto: **unificar formato, no tomar decisiones analíticas**.

- Los datos faltantes se dejan vacíos, no se imputan.
- Las CCAAs que no reportaron ciertos años simplemente no aparecen.
- Los cambios de nomenclatura en especies (ej. separación de "Acuáticas" en "Acuáticas" + "Anátidas") se mantienen como están.
- Las normalizaciones son mínimas: unificación de mayúsculas, corrección de acentos, eliminación de variantes triviales.

## Limitaciones conocidas

- **2005**: solo capturas y licencias (no hay sueltas, producción ni terrenos).
- **Algunas CCAAs no reportan todos los años**: Canarias (2018-2021), Extremadura (2008-2009), C. Valenciana (2009, 2011-2012), etc. Esto refleja los datos fuente.
- **Rupturas de series temporales en especies**: "Acuáticas y anátidas" se separa en "Acuáticas" + "Anátidas" a partir de ~2016. "Paloma" se separa en "Paloma torcaz" + "Paloma zurita". Se mantienen como están.
- **Algunas provincias no desglosan licencias**: en esos casos se incluye una fila con `provincia = "(Total CCAA)"`.

## Fuente de datos

Todos los datos provienen del **Anuario de Estadística Forestal** publicado por el MITECO:
https://www.miteco.gob.es/es/biodiversidad/estadisticas/forestal_anuarios_702702702702.html

Los datos se publican bajo licencia de reutilización de la información del sector público (Ley 37/2007).

**Nota:** este proyecto garantiza que los CSVs reflejan fielmente lo publicado en los Excel del MITECO (verificado valor a valor). No valora ni modifica el contenido de los datos originales.

---

## Para desarrolladores

### Regenerar desde cero

Para regenerar los CSVs a partir de los Excel originales del MITECO se necesita [Python](https://www.python.org/downloads/):

```bash
pip install -r requirements.txt       # instala las dependencias
python unificador.py --todo           # descarga los Excel del MITECO y genera los CSVs
```

```bash
# Solo procesar (si ya tienes los Excel en raw_data/)
python unificador.py --procesar

# Solo descargar
python unificador.py --descargar

# Windows con problemas de codificación
python -X utf8 unificador.py --procesar
```

### Estructura del proyecto

```
├── raw_data/               # Archivos Excel descargados del MITECO
├── output/                 # CSVs generados (resultado)
├── unificador.py           # Script principal: descarga, procesa y unifica
├── auditoria.py            # Verificación de integridad contra los Excel originales
├── test_unitarios.py       # Verificación independiente por año
└── README.md
```

### Añadir nuevos años

Cuando MITECO publique un nuevo anuario, el dataset se puede ampliar. Ver las instrucciones detalladas en [CLAUDE.md](CLAUDE.md).

### Contribuir

Si encuentras errores o quieres mejorar el proyecto, abre un issue o un pull request. Las contribuciones más útiles:

- Soporte para nuevos años cuando MITECO los publique.
- Correcciones en la normalización de nombres (especies, provincias, CCAAs).
- Mejoras en los scripts de verificación.

## Licencia

Código bajo [licencia MIT](LICENSE). Los datos originales son del MITECO (Gobierno de España), publicados bajo la ley de reutilización de información del sector público.
