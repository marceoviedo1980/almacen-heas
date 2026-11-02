"""Convierte el Excel IMM del Almacén de Farmacia (hoja 'IMM') en almacen.json para la PWA.

Uso:
    python actualizar_datos.py ALMACEN_ISRAEL.xlsx
Requiere:
    pip install openpyxl

Columnas que reconoce (por el nombre del encabezado, en cualquier orden):
    Código | Nombre Genérico | Forma Farmacéutica | Concentración | Total
    Opcionales (vencimientos):  Lote | Vencimiento (o Fecha de vencimiento)
Si un mismo código aparece en varias filas (una por lote), se agrupan: el saldo
del ítem es la suma de sus lotes y cada lote conserva su vencimiento.
"""
import sys, re, json, datetime, openpyxl

src = sys.argv[1] if len(sys.argv) > 1 else "ALMACEN_ISRAEL.xlsx"
wb = openpyxl.load_workbook(src, data_only=True)
ws = wb["IMM"] if "IMM" in wb.sheetnames else wb.worksheets[0]
rows = list(ws.iter_rows(values_only=True))


def txt(v):
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


hi = next(i for i, r in enumerate(rows) if r and txt(r[0]).lower().startswith("c"))
head = [txt(h).lower() for h in rows[hi]]


def col(*names):
    for n in names:
        for i, h in enumerate(head):
            if h.startswith(n):
                return i
    return None


cC, cN, cF, cK = col("código", "codigo"), col("nombre"), col("forma"), col("concentr")
cT = col("total", "saldo", "cantidad")
cV, cL = col("venc", "fecha de venc", "fecha venc"), col("lote")


def fecha(v):
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%Y-%m-%d")
    s = txt(v)
    m = re.match(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})$", s)
    return f"{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}" if m else s


def num(v):
    try:
        q = float(v or 0)
    except (TypeError, ValueError):
        return 0
    return int(q) if q == int(q) else q


items, by_code = [], {}
for r in rows[hi + 1:]:
    code = txt(r[cC])
    if not code:
        continue
    qty = num(r[cT])
    it = by_code.get(code)
    if it is None:
        it = {
            "codigo": code,
            "medicamento": txt(r[cN]),
            "forma": txt(r[cF]),
            "concentracion": txt(r[cK]),
            "cantidad": 0,
            "tipo": "medicamento" if re.search(r"(^|-)[A-V]\d{4}", code) else "insumo",
            "lotes": [],
        }
        by_code[code] = it
        items.append(it)
    it["cantidad"] += qty
    lote = txt(r[cL]) if cL is not None else ""
    venc = fecha(r[cV]) if cV is not None and r[cV] else ""
    if (lote or venc) and qty:
        it["lotes"].append({"lote": lote or None, "fecha_vencimiento": venc or None, "cantidad": qty})

for it in items:
    it["lotes"].sort(key=lambda l: l["fecha_vencimiento"] or "9999")

mes = ""
if "IMM (2)" in wb.sheetnames:
    for r in wb["IMM (2)"].iter_rows(max_row=10, values_only=True):
        vals = [txt(x) for x in r]
        if "Mes reportado:" in vals:
            mes = next((v for v in vals[vals.index("Mes reportado:") + 1:] if v), "")

out = {"actualizado": datetime.date.today().isoformat(), "mes": mes, "items": items}
with open("almacen.json", "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, separators=(",", ":"))

con = sum(1 for i in items if i["cantidad"] > 0)
conv = sum(1 for i in items if i["lotes"])
print(f"OK: {len(items)} ítems ({con} con stock, {len(items) - con} en cero, {conv} con vencimiento). Mes: {mes or '-'}")
print('Ahora cambia la versión en sw.js (const CACHE = PREFIX + "...") y sube los archivos.')
