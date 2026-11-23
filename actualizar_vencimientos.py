"""Convierte el Excel de vencimientos del Almacén (.xls / .xlsx) en vencimientos.json para la PWA.

Uso (desde esta carpeta):
    python actualizar_vencimientos.py "almac venc OF.xls"
Requiere:
    pip install xlrd openpyxl        (xlrd para .xls, openpyxl para .xlsx)

Columnas que reconoce (por el nombre del encabezado, en cualquier orden):
    Código | Cantidad | Vencimiento      (obligatorias)
    Lote                                 (opcional)
Un mismo código puede repetirse (una fila por lote / por fecha). 'S/V' en Vencimiento
= producto sin vencimiento (se muestra en gris, sin semáforo).

IMPORTANTE: vencimientos.json es independiente de almacen.json. Puedes actualizar los
saldos (actualizar_datos.py) sin perder los vencimientos, y viceversa.

Corrección automática de años: Excel interpreta "jun-30" como 1930. Cualquier fecha con
año < 2000 se corrige sumando 100 años (1930 -> 2030) y se avisa en pantalla.
"""
import sys, re, json, datetime, os

here = os.path.dirname(os.path.abspath(__file__))
src = sys.argv[1] if len(sys.argv) > 1 else "almac venc OF.xls"


def txt(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def norm_code(c):
    return re.sub(r"\s+", "", txt(c)).upper()


def read_rows(path):
    """Devuelve lista de filas (listas de valores); las fechas llegan como datetime."""
    if path.lower().endswith(".xls"):
        import xlrd
        wb = xlrd.open_workbook(path)
        sh = wb.sheet_by_index(0)
        out = []
        for r in range(sh.nrows):
            row = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                if cell.ctype == xlrd.XL_CELL_DATE:
                    row.append(xlrd.xldate_as_datetime(cell.value, wb.datemode))
                elif cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    row.append(None)
                else:
                    row.append(cell.value)
            out.append(row)
        return out
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]


rows = read_rows(src)
hi = next(i for i, r in enumerate(rows) if r and any(txt(x).lower().startswith("c") and "d" in txt(x).lower() for x in r[:2]))
head = [txt(h).lower() for h in rows[hi]]


def col(*names):
    for n in names:
        for i, h in enumerate(head):
            if h.startswith(n):
                return i
    return None


cC, cQ, cV, cL = col("código", "codigo"), col("cantidad", "saldo", "total"), col("venc"), col("lote")
if None in (cC, cQ, cV):
    sys.exit(f"No encuentro las columnas Código / Cantidad / Vencimiento. Encabezados: {head}")

# Códigos del catálogo (para unir "DAC025a - SUS" con "DAC025a-SUS")
with open(os.path.join(here, "almacen.json"), encoding="utf-8") as f:
    cat = json.load(f)
by_norm = {norm_code(i["codigo"]): i for i in cat["items"]}

items, sin_match, corregidas, vencidas, mayor_saldo = {}, [], [], [], []
hoy = datetime.date.today()

for r in rows[hi + 1:]:
    if not r or not txt(r[cC]):
        continue
    raw = txt(r[cC])
    it = by_norm.get(norm_code(raw))
    if not it:
        sin_match.append(raw)
        continue
    try:
        qty = float(r[cQ] or 0)
    except (TypeError, ValueError):
        qty = 0
    qty = int(qty) if qty == int(qty) else qty
    v = r[cV]
    lote = txt(r[cL]) if cL is not None else ""
    ent = {"lote": lote or None, "fecha_vencimiento": None, "cantidad": qty}
    if isinstance(v, datetime.datetime):
        d = v.date()
        if d.year < 2000:
            nuevo = d.replace(year=d.year + 100)
            corregidas.append((it["codigo"], it["medicamento"], d.isoformat(), nuevo.isoformat()))
            d = nuevo
        ent["fecha_vencimiento"] = d.isoformat()
        if d < hoy:
            vencidas.append((it["codigo"], it["medicamento"], d.isoformat()))
    elif txt(v).upper() in ("S/V", "SV", "S/F", "N/A", "-", ""):
        ent["sv"] = True  # sin vencimiento
    else:
        m = re.match(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})$", txt(v))
        if m:
            ent["fecha_vencimiento"] = f"{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}"
        else:
            sin_match.append(f"{raw} (fecha no reconocida: {txt(v)!r})")
            continue
    items.setdefault(it["codigo"], []).append(ent)

for k, lotes in items.items():
    lotes.sort(key=lambda l: l["fecha_vencimiento"] or "9999")
    tot = sum(l["cantidad"] for l in lotes)
    saldo = by_norm[norm_code(k)]["cantidad"]
    if tot > saldo:
        mayor_saldo.append((k, by_norm[norm_code(k)]["medicamento"], tot, saldo))

out = {"actualizado": hoy.isoformat(), "fuente": os.path.basename(src), "items": items}
with open(os.path.join(here, "vencimientos.json"), "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(",", ":"))

print(f"OK: {len(items)} ítems con vencimiento -> vencimientos.json")
if corregidas:
    print(f"\n[!] {len(corregidas)} fechas con año < 2000 corregidas (+100 años). REVISA que sean correctas:")
    for c in corregidas:
        print(f"    {c[0]:<16} {c[1][:38]:<38} {c[2]} -> {c[3]}")
if vencidas:
    print(f"\n[!] {len(vencidas)} lotes YA VENCIDOS:")
    for c in vencidas:
        print(f"    {c[0]:<16} {c[1][:38]:<38} {c[2]}")
if mayor_saldo:
    print("\n[!] Cantidad del Excel MAYOR que el saldo de almacen.json (revisar):")
    for c in mayor_saldo:
        print(f"    {c[0]:<16} {c[1][:38]:<38} excel={c[2]}  saldo={c[3]}")
if sin_match:
    print("\n[!] Filas ignoradas (código no está en almacen.json o fecha ilegible):")
    for c in sin_match:
        print("    " + c)
print('\nAhora cambia la versión en sw.js (const CACHE = PREFIX + "...") y sube los archivos.')
