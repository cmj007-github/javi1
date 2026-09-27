# Serie de tiempo de planes móviles Entel (Wayback Machine)

Script que arma una base de datos en Excel con los **precios y características de los planes móviles de entel.cl**, usando la **primera captura de cada mes** en la Wayback Machine de Internet Archive (web.archive.org).

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
# Todo el período 2015 → hoy, con las páginas de planes que Entel ha usado
python entel_wayback.py

# Acotar el período
python entel_wayback.py --desde 2018 --hasta 202412

# Elegir páginas específicas (repetible)
python entel_wayback.py --url www.entel.cl/planes/ --url www.entel.cl/personas/planes-moviles/

# Nombre del archivo de salida
python entel_wayback.py --salida entel_planes.xlsx
```

Las descargas se guardan en `.cache_wayback/`, así que si vuelve a correr el script no se descargan de nuevo.

## Cómo funciona

1. **API CDX** (`web.archive.org/cdx/search/cdx`) con `collapse=timestamp:6` y `filter=statuscode:200`, que devuelve la primera captura válida de cada mes (AAAAMM).
2. Descarga el HTML original de cada captura (`/web/<timestamp>id_/<url>`, sin la barra de Wayback).
3. Extracción heurística: busca "tarjetas" que tengan un precio (`$xx.xxx`) y una característica de plan (GB, minutos, "gigas libres"…). También revisa el JSON embebido (p. ej. `__NEXT_DATA__`) para las páginas que se renderizan con JavaScript.

## Contenido del Excel

| Hoja | Contenido |
|---|---|
| **Planes** | Una fila por plan y mes: plan, precio, precio normal, GB, datos/minutos/SMS libres, redes sociales libres, roaming, % de descuento, meses de promoción, link a la captura y el texto original para revisarlo |
| **Resumen_Mensual** | Por mes: n° de planes, precio mínimo, mediano y máximo, rango de GB, planes con datos libres y CLP por GB |
| **Capturas** | Registro de cada captura consultada y su estado (ok / sin planes detectados / error) |
| **Notas** | Metodología y advertencias |

## Advertencias

- La extracción es heurística porque el diseño del sitio cambió muchas veces. Revise la columna *Texto bruto* y la hoja *Capturas* para validar.
- Si una captura sale como "sin planes detectados", lo más probable es que la página cargara los precios con JavaScript/API y que la Wayback Machine no los haya archivado. En ese caso agregue otra URL con `--url`.
- La página de planes puede cambiar de URL con los años. Para ver qué URLs están archivadas: `https://web.archive.org/cdx/search/cdx?url=entel.cl/*&filter=statuscode:200&collapse=urlkey&fl=original` y filtre las que contengan "plan".

## Tests

```bash
pip install pytest && python -m pytest -q tests
```
