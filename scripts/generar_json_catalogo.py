# generar_json_catalogo.py
import pandas as pd
import os
from openpyxl import load_workbook
from io import BytesIO
from PIL import Image as PILImage
import json

# ---------- CONFIG ----------
RUTA_PROD = "data/NHMX511301944912-produdctos.xlsx"   # <<-- Origen
RUTA_SALES = "data/sales.xlsx"
HOJA = "PCODE"
CARPETA_IMAGENES = "catalogo-img"
OUT_DIR = "output"
# ----------------------------

os.makedirs(CARPETA_IMAGENES, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

# 1) Leer productos
df = pd.read_excel(RUTA_PROD, sheet_name=HOJA)

# *** Ajuste para incluir precios correctamente ***
if "Precio" in df.columns and "Price" not in df.columns:
    df.rename(columns={"Precio": "Price"}, inplace=True)

# Validación mínima
for req in ["ProductCode", "Product Name", "Qty"]:
    if req not in df.columns:
        raise SystemExit(f"Falta columna requerida en productos: {req}")

# Asegurar datos numéricos de precio
df["Price"] = pd.to_numeric(df.get("Price", 0), errors="coerce").fillna(0)

# ====================================================
# 2) Leer ventas
ventas_total = pd.Series(0, index=df["ProductCode"])
if os.path.exists(RUTA_SALES):
    sales = pd.read_excel(RUTA_SALES)
    if "ProductCode" in sales.columns and ("Quantity" in sales.columns or "Cantidad" in sales.columns):
        q_col = "Quantity" if "Quantity" in sales.columns else "Cantidad"
        ventas_sum = sales.groupby("ProductCode")[q_col].sum()
        ventas_total = df["ProductCode"].map(ventas_sum).fillna(0)
else:
    print("Aviso: sales.xlsx no encontrado. Ventas = 0")
# ====================================================

df["Ventas"] = df["ProductCode"].map(ventas_total).fillna(0).astype(int)
df["Inventario"] = df["Qty"].astype(int) - df["Ventas"].astype(int)

# ====================================================
# 4) Extraer imágenes
try:
    wb = load_workbook(RUTA_PROD, data_only=True)
    ws = wb[HOJA]
    imgs = getattr(ws, "_images", [])
    for img in imgs:
        try:
            row_excel = img.anchor._from.row + 1
        except:
            continue
        idx = row_excel - 2
        if idx < 0 or idx >= len(df): continue
        code = str(df.iloc[idx]["ProductCode"])
        out_file = os.path.join(CARPETA_IMAGENES, f"{code}.png")

        data = img._data() if hasattr(img,"_data") else None
        if data:
            PILImage.open(BytesIO(data)).save(out_file)
        print("Imagen guardada:", out_file)
except Exception as e:
    print("No fue posible extraer imágenes automáticamente:", e)
# ====================================================

# 5) Generar catalogo.json con precio correcto
catalogo = []
for _, row in df.iterrows():
    code = str(row["ProductCode"])
    item = {
        "productCode": code,
        "name": row.get("Product Name",""),
        "description": row.get("Description", row.get("Product Name","")),
        "price": float(row.get("Price",0)),                    # <====== AQUÍ YA VIENE EL PRECIO
        "qty_initial": int(row.get("Qty",0)),
        "ventas": int(row.get("Ventas",0)),
        "inventory": int(row.get("Inventario",0)),
        "category": row.get("Category","") if "Category" in df.columns else "",
        "image": f"{CARPETA_IMAGENES}/{code}.png" if os.path.exists(f"{CARPETA_IMAGENES}/{code}.png") else ""
    }
    catalogo.append(item)

with open(f"{OUT_DIR}/catalogo.json","w",encoding="utf-8") as f:
    json.dump(catalogo,f,indent=2,ensure_ascii=False)

# 6) inventario_actual.json
inventario = [{"productCode":i["productCode"], "inventory":i["inventory"]} for i in catalogo]
with open(f"{OUT_DIR}/inventario_actual.json","w",encoding="utf-8") as f:
    json.dump(inventario,f,indent=2,ensure_ascii=False)

print("✔ LISTO — catalogo.json ahora INCLUYE 'price' correctamente.")
