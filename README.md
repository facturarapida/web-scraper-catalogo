# web-scraper-catalogo

**Web scraping con export a Excel/CSV** — extrae un catalogo completo de una tienda (nombre, precio, valoracion, disponibilidad) y lo entrega como `catalogo.csv` y `catalogo.xlsx` listos para analizar.

## El problema que resuelve

Una tienda online tiene ~1000 productos repartidos en 50 paginas. Recopilarlos "a mano" (copiar-pegar a Excel) cuesta horas, sale incompleto y se rompe al repito. Este scraper lo hace en segundos:

1. Recorre las paginas del catalogo siguiendo su paginacion real.
2. Extrae de cada producto: **nombre, precio, rating (1-5), disponibilidad**.
3. Escribe dos ficheros de salida:
   - `catalogo.csv` — UTF-8-BOM (se abre directo en Excel, acentos correctos).
   - `catalogo.xlsx` — Excel real (openpyxl): cabecera en negrita, columnas con ancho, **filtros automaticos**.
4. Imprime un resumen: total, en stock, precio medio.

## Ejecucion

```bash
pip install -r requirements.txt
python scraper.py --pages 5        # 5 paginas (100 productos)
python scraper.py --pages 50       # catalogo completo
python scraper.py --pages 3 --limit 25
```

## Resultado real (ejecucion del 2026-11-17, 5 paginas)

```
OK  100 productos (5 paginas)
     en stock: 100  |  precio medio: 34.56
CSV   output/catalogo.csv
XLSX  output/catalogo.xlsx
```

Muestra (`catalogo-muestra.csv`, las 10 primeras filas del resultado completo de 100):

| producto               | precio | rating | disponibilidad |
|------------------------|--------|--------|---------------|
| A Light in the Attic   | 51.77  | 3      | In stock      |
| Tipping the Velvet     | 53.74  | 1      | In stock      |
| Soumission             | 50.10  | 1      | In stock      |

El `catalogo.xlsx` verificado: sheet "Catalogo", 101 filas (cabecera + 100), 4 columnas, filtros activos.

## Que se automatiza (vs. manual)

| Tarea | Manual | Con este repo |
|-------|--------|---------------|
| 1000 productos en 50 paginas | ~4 h | ~30 s |
| Errores de tecleo | frecuentes | 0 |
| Precio como texto "£51.77" | depende | numero `51.77` (calculable en Excel) |
| Filtros/pivotes en Excel | a mano | ya incluidos en el .xlsx |

## Adaptar a tu sitio

1. Cambiar `BASE` (linea 24).
2. Ajustar `parse_product()` con los selectores de tu sitio (misma logica que `selector-guia.md` del repo hermano `rpa-ecommerce-bot`: DevTools F12 → atributos estables).
3. La paginacion se sigue por el `<li class="next">`; si tu sitio usa botones "Siguiente" con otro markup, cambia esa linea.

## Dependencias

`requests`, `beautifulsoup4`, `lxml`, `openpyxl` (ver `requirements.txt`). Python >= 3.10.

## Licencia

MIT — uso libre como punto de partida para extracciones propias.
