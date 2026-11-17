# -*- coding: utf-8 -*-
# Web scraper de catalogo completo con resultado en CSV y Excel.
#
# Problema: una tienda tiene 1000 productos repartidos en 50 paginas.
# Exportarlos "a mano" (copiar-pegar) cuesta horas y nunca esta completo.
# Este scraper recorre las paginas, extrae cada producto y escribe:
#   - catalogo.csv   (UTF-8-BOM, se abre directo en Excel)
#   - catalogo.xlsx  (Excel real con openpyxl, con formato y filtros)
#
# Uso:
#   python scraper.py --pages 5          # 5 primeras paginas (100 productos)
#   python scraper.py --pages 50         # catalogo completo
#   python scraper.py --pages 3 --limit 10
import argparse
import csv
import pathlib
import re
import sys
import time

import requests
from bs4 import BeautifulSoup

BASE = "https://books.toscrape.com/"
HEADERS = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Python/requests"}
OUT = pathlib.Path(__file__).parent / "output"
DELAY = 0.3  # segundos entre paginas (cortesía con el sitio)


def fetch(url: str) -> str:
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.text


def parse_product(card) -> dict:
    name = card.find("h3").find("a").get("title") or ""
    raw = card.find(class_="price_color").get_text().strip()
    # "£51.77" (o doble-codificado) -> "51.77"
    m = re.search(r"(\d+\.?\d*)", raw)
    price = m.group(1) if m else raw
    star = card.find(class_="star-rating")
    rating_name = star.get("class")[1] if star and len(star.get("class", [])) > 1 else ""
    rating = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}.get(rating_name, 0)
    avail = ""
    a = card.find(class_="availability")
    if a:
        avail = " ".join(a.get_text(" ", strip=True).split())
    return {"producto": name, "precio": price, "rating": rating, "disponibilidad": avail}


def scrape(pages: int, limit: int | None = None) -> list[dict]:
    rows = []
    for i in range(1, pages + 1):
        url = BASE if i == 1 else f"{BASE}catalogue/page-{i}.html"
        soup = BeautifulSoup(fetch(url), "lxml")
        next_li = soup.find("li", class_="next")
        cards = soup.find_all(class_="product_pod")
        for c in cards:
            rows.append(parse_product(c))
            if limit and len(rows) >= limit:
                return rows
        print(f"pagina {i}: +{len(cards)} (total {len(rows)})", file=sys.stderr)
        if i < pages and not next_li:
            print("fin de la paginacion del sitio", file=sys.stderr)
            break
        time.sleep(DELAY)
    return rows


def write_csv(rows: list[dict], path: pathlib.Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["producto", "precio", "rating", "disponibilidad"])
        w.writeheader()
        w.writerows(rows)


def write_xlsx(rows: list[dict], path: pathlib.Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Catalogo"
    heads = ["Producto", "Precio", "Rating (1-5)", "Disponibilidad"]
    ws.append(heads)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append([r["producto"], r["precio"], r["rating"], r["disponibilidad"]])
    for i, name in enumerate(heads, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(18, len(name) + 2)
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def main() -> int:
    ap = argparse.ArgumentParser(description="Scraper de catalogo con export CSV/Excel")
    ap.add_argument("--pages", type=int, default=5, help="num. de paginas a recorrer (defecto 5)")
    ap.add_argument("--limit", type=int, default=None, help="cortar en N productos")
    args = ap.parse_args()

    OUT.mkdir(exist_ok=True)
    try:
        rows = scrape(args.pages, args.limit)
    except Exception as exc:
        print(f"FALLO: {exc}", file=sys.stderr)
        return 1
    if not rows:
        print("FALLO: sin datos", file=sys.stderr)
        return 1

    csv_path = OUT / "catalogo.csv"
    xlsx_path = OUT / "catalogo.xlsx"
    write_csv(rows, csv_path)
    write_xlsx(rows, xlsx_path)

    prices = [re.search(r"(\d+\.?\d*)", r["precio"]).group(1) for r in rows if re.search(r"\d", r["precio"])]
    prices = [float(p) for p in prices]
    avg = sum(prices) / len(prices) if prices else 0
    in_stock = sum(1 for r in rows if "In stock" in r["disponibilidad"])

    print(f"OK  {len(rows)} productos ({args.pages} paginas)")
    print(f"     en stock: {in_stock}  |  precio medio: {avg:.2f}")
    print(f"CSV   {csv_path}")
    print(f"XLSX  {xlsx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
