# generar_json_catalogo.py
import pandas as pd
import os
from openpyxl import load_workbook
from io import BytesIO
from PIL import Image as PILImage
import json

# ---------- CONFIG ----------
RUTA_PROD = "data/NHMX511301944912-produdctos.xlsx"
RUTA_SALES = "data/sales.xlsx"
HOJA = "PCODE"
CARPETA_IMAGENES = "catalogo-img"
OUT_DIR = "output"
# ----------------------------

os.makedirs(CARPETA_IMAGENES, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

# 1) Leer productos
df = pd.read_excel(RUTA_PROD, sheet_name=HOJA)

# Columna fallback names: puedes adaptarlas si tu sheet usa otros nombres
# Se buscan: ProductCode, Product Name, Qty, Price, Category (opcional), Description (opcional)
for req in ["ProductCode", "Product Name", "Qty"]:
    if req not in df.columns:
        raise SystemExit(f"Falta columna requerida en productos: {req}")

# 2) Leer ventas (solo lectura)
ventas_total = pd.Series(0, index=df["ProductCode"])
if os.path.exists(RUTA_SALES):
    sales = pd.read_excel(RUTA_SALES)
    # Se esperan columnas ProductCode y Quantity en sales.xlsx
    if "ProductCode" in sales.columns and ("Quantity" in sales.columns or "Cantidad" in sales.columns):
        q_col = "Quantity" if "Quantity" in sales.columns else "Cantidad"
        ventas_sum = sales.groupby("ProductCode")[q_col].sum()
        ventas_total = df["ProductCode"].map(ventas_sum).fillna(0)
    else:
        print("Aviso: sales.xlsx no tiene columnas esperadas (ProductCode / Quantity). Se asume 0 ventas.")
else:
    print("Aviso: sales.xlsx no encontrado. Ventas = 0")

# 3) Calcular Inventario
df["Ventas"] = df["ProductCode"].map(ventas_total).fillna(0).astype(int)
df["Inventario"] = df["Qty"].astype(int) - df["Ventas"].astype(int)

# 4) Extraer imágenes (intento con openpyxl; si no están, asumimos que usarás catalogo-img/manualmente)
try:
    wb = load_workbook(RUTA_PROD, data_only=True)
    ws = wb[HOJA]
    imgs = getattr(ws, "_images", [])
    print(f"Imágenes detectadas en hoja: {len(imgs)}")
    for img in imgs:
        try:
            row_excel = img.anchor._from.row + 1
        except Exception:
            try:
                row_excel = img.anchor.row
            except Exception:
                continue
        idx = row_excel - 2
        if idx < 0 or idx >= len(df):
            continue
        code = str(df.iloc[idx]["ProductCode"])
        out_name = f"{code}.png"
        out_path = os.path.join(CARPETA_IMAGENES, out_name)
        # obtener bytes
        data = None
        if hasattr(img, "_data"):
            data = img._data()
        elif hasattr(img, "ref"):
            data = img.ref
        if data:
            try:
                pil = PILImage.open(BytesIO(data))
                pil.save(out_path)
                print("Guardada imagen:", out_name)
            except Exception as e:
                print("Error guardando imagen", out_name, e)
        else:
            if hasattr(img, "path") and img.path:
                from shutil import copyfile
                copyfile(img.path, out_path)
                print("Copiada imagen desde path:", out_name)
except Exception as e:
    print("No fue posible extraer imágenes automáticamente:", e)

# 5) Generar catalogo.json (cada producto con campos útiles)
catalogo = []
for _, row in df.iterrows():
    code = str(row["ProductCode"])
    item = {
        "productCode": code,
        "name": row.get("Product Name", ""),
        "description": row.get("Description", row.get("Product Name", "")),
        "price": float(row.get("Price", 0)) if "Price" in row.index else None,
        "qty_initial": int(row.get("Qty", 0)),
        "ventas": int(row.get("Ventas", 0)),
        "inventory": int(row.get("Inventario", 0)),
        "category": row.get("Category", "") if "Category" in row.index else "",
        "image": os.path.join(CARPETA_IMAGENES, f"{code}.png") if os.path.exists(os.path.join(CARPETA_IMAGENES, f"{code}.png")) else ""
    }
    catalogo.append(item)

with open(os.path.join(OUT_DIR, "catalogo.json"), "w", encoding="utf-8") as f:
    json.dump(catalogo, f, indent=2, ensure_ascii=False)

# 6) Generar inventario_actual.json (ligero)
inventario = [{"productCode": it["productCode"], "inventory": it["inventory"]} for it in catalogo]
with open(os.path.join(OUT_DIR, "inventario_actual.json"), "w", encoding="utf-8") as f:
    json.dump(inventario, f, indent=2, ensure_ascii=False)

print("✅ JSON creados en /output: catalogo.json & inventario_actual.json")
