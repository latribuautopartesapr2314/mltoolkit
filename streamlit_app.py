import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

NOMBRE_HOJA_CONTROL = "CONTROL_CRUCE_INTEGRALY"

SUMA_FIJA_PRECIO = 12_000

UMBRAL_PRECIO_ESTADO = 1_500_000
ESTADO_ACTIVA = "Activa"
ESTADO_PAUSADA = "Pausada"

POSIBLES_SKU = [
    "sku", "SKU", "Sku",
    "item_code", "ITEM_CODE", "Item Code",
    "codigo", "Código", "CODIGO"
]

POSIBLES_STOCK_ORIGEN = [
    "AStk", "astk", "ASTK",
    "stock", "Stock", "STOCK"
]

POSIBLES_STOCK_DESTINO = [
    "cantidad", "Cantidad", "CANTIDAD",
    "stock", "Stock", "STOCK"
]

POSIBLES_PRECIO_DESTINO = [
    "precio", "Precio", "PRECIO",
    "price", "Price", "PRICE"
]

POSIBLES_ESTADO_DESTINO = [
    "estado", "Estado", "ESTADO",
    "status", "Status", "STATUS"
]

POSIBLES_CUOTAS = [
    "cuotas", "Cuotas", "CUOTAS"
]

POSIBLES_MLA = [
    "mla", "MLA", "Mla",
    "id", "ID",
    "publicacion", "Publicacion", "PUBLICACION"
]

# IMPORTANTE:
# Estas son las columnas del archivo GLOBAL / ACTUALIZACIÓN PRECIO.
# No usa columnas con _2.
COLUMNAS_PRECIO_ORIGEN = {
    "sin_cuotas": "Precio ML Clasica",
    "3_cuotas": "Precio ML Premium",
    "6_cuotas": "Precio ML Premium 6c",
    "9_cuotas": "Precio ML Premium 9c",
    "12_cuotas": "Precio ML Premium 12c",
}


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def normalizar_sku(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()

    if texto.endswith(".0"):
        texto = texto[:-2]

    return texto


def detectar_columna(columnas, posibles, obligatorio=True, descripcion="columna"):
    columnas_lista = list(columnas)

    for posible in posibles:
        if posible in columnas_lista:
            return posible

    if obligatorio:
        raise ValueError(
            f"No se encontró {descripcion}. Columnas disponibles: "
            + ", ".join(map(str, columnas_lista))
        )

    return None


def convertir_precio_a_numero(valor):
    """
    Convierte precios tipo:
    - 67285
    - 67.285,00 ARS
    - $ 67.285,00
    - 67285.00
    a número.
    """

    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if texto == "":
        return None

    # Quita ARS, $, espacios y cualquier texto.
    # Conserva números, punto, coma y signo menos.
    texto = re.sub(r"[^\d,.\-]", "", texto)

    if texto == "":
        return None

    # Formato argentino: 67.285,00
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "")
            texto = texto.replace(",", ".")
        else:
            texto = texto.replace(",", "")

    # Formato: 67285,00
    elif "," in texto and "." not in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except Exception:
        return None


def convertir_stock(valor):
    numero = convertir_precio_a_numero(valor)

    if numero is None:
        return None

    try:
        return int(float(numero))
    except Exception:
        return None


def formatear_numero(valor):
    if valor is None:
        return None

    try:
        valor = float(valor)
    except Exception:
        return None

    if valor.is_integer():
        return int(valor)

    return round(valor, 2)


def clasificar_cuotas(valor):
    """
    Mapea la columna cuotas de Integraly:
    - Vacío / Sin cuotas / No Agregar Cuotas -> sin_cuotas
    - 3 -> 3_cuotas
    - 6 -> 6_cuotas
    - 9 -> 9_cuotas
    - 12 -> 12_cuotas
    """

    if pd.isna(valor):
        return "sin_cuotas"

    texto = str(valor).strip().lower()

    if texto == "":
        return "sin_cuotas"

    texto = (
        texto
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
    )

    if "no agregar" in texto:
        return "sin_cuotas"

    if "sin cuota" in texto:
        return "sin_cuotas"

    if re.search(r"\b12\b", texto):
        return "12_cuotas"

    if re.search(r"\b9\b", texto):
        return "9_cuotas"

    if re.search(r"\b6\b", texto):
        return "6_cuotas"

    if re.search(r"\b3\b", texto):
        return "3_cuotas"

    return None


def determinar_estado(precio_final, stock_final):
    """
    Prioridad:
    1) Precio mayor a 1.500.000 -> Pausada siempre.
    2) Precio menor o igual a 1.500.000 y stock >= 1 -> Activa.
    3) Precio menor o igual a 1.500.000 y stock <= 0 -> Pausada.
    """

    if precio_final is None:
        return None, "No se pudo determinar precio final"

    try:
        precio_num = float(precio_final)
    except Exception:
        return None, "Precio final inválido"

    if precio_num > UMBRAL_PRECIO_ESTADO:
        return ESTADO_PAUSADA, "Precio mayor a 1.500.000"

    if stock_final is None:
        return None, "No se pudo determinar stock"

    try:
        stock_num = int(float(stock_final))
    except Exception:
        return None, "Stock inválido"

    if stock_num >= 1:
        return ESTADO_ACTIVA, "Precio menor o igual a 1.500.000 y stock >= 1"

    return ESTADO_PAUSADA, "Precio menor o igual a 1.500.000 y stock <= 0"


def obtener_headers_ws(ws):
    headers = {}

    for cell in ws[1]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column

    return headers


def obtener_o_crear_columna(ws, headers, nombre_columna):
    if nombre_columna in headers:
        return headers[nombre_columna]

    nueva_col = ws.max_column + 1
    ws.cell(row=1, column=nueva_col).value = nombre_columna
    headers[nombre_columna] = nueva_col

    return nueva_col


def agregar_tabla(ws, titulo, encabezados, registros, fila_vacia):
    ws.append([])
    ws.append([titulo])
    ws.append(encabezados)

    if registros:
        for registro in registros:
            ws.append([registro.get(col, "") for col in encabezados])
    else:
        ws.append(fila_vacia)


# ============================================================
# ARMAR BASE DESDE GLOBAL / ACTUALIZACIÓN PRECIO
# ============================================================

def armar_base_actualizacion(actualizacion_bytes):
    xls = pd.ExcelFile(BytesIO(actualizacion_bytes))

    filas_origen = []
    control_hojas_ignoradas = []
    control_precios_invalidos = []
    control_stock_invalido = []

    for hoja in xls.sheet_names:

        if str(hoja).upper().startswith("CONTROL"):
            continue

        df = pd.read_excel(
            BytesIO(actualizacion_bytes),
            sheet_name=hoja,
            dtype=str
        )

        if df.empty:
            control_hojas_ignoradas.append({
                "Hoja": hoja,
                "Motivo": "Hoja vacía"
            })
            continue

        columnas = list(df.columns)

        try:
            col_sku = detectar_columna(
                columnas,
                POSIBLES_SKU,
                obligatorio=True,
                descripcion="columna SKU origen"
            )

            col_stock = detectar_columna(
                columnas,
                POSIBLES_STOCK_ORIGEN,
                obligatorio=True,
                descripcion="columna stock origen"
            )

        except Exception as e:
            control_hojas_ignoradas.append({
                "Hoja": hoja,
                "Motivo": str(e)
            })
            continue

        faltan_precios = [
            col for col in COLUMNAS_PRECIO_ORIGEN.values()
            if col not in columnas
        ]

        if faltan_precios:
            control_hojas_ignoradas.append({
                "Hoja": hoja,
                "Motivo": "Faltan columnas de precio: " + ", ".join(faltan_precios)
            })
            continue

        df["_SKU_KEY"] = df[col_sku].apply(normalizar_sku)
        df = df[df["_SKU_KEY"] != ""].copy()

        for idx, fila in df.iterrows():

            sku_key = fila["_SKU_KEY"]

            stock_final = convertir_stock(fila[col_stock])

            if stock_final is None:
                control_stock_invalido.append({
                    "Hoja origen": hoja,
                    "Fila origen": idx + 2,
                    "SKU": sku_key,
                    "Valor AStk": fila[col_stock],
                })

            registro = {
                "_SKU_KEY": sku_key,
                "_HOJA_ORIGEN": hoja,
                "_FILA_ORIGEN": idx + 2,
                "_ASTK": stock_final,
            }

            for clave_precio, col_precio in COLUMNAS_PRECIO_ORIGEN.items():

                precio_base = convertir_precio_a_numero(fila[col_precio])

                if precio_base is not None:
                    precio_final = formatear_numero(precio_base + SUMA_FIJA_PRECIO)
                else:
                    precio_final = None

                registro[clave_precio] = precio_final

                if precio_final is None:
                    control_precios_invalidos.append({
                        "Hoja origen": hoja,
                        "Fila origen": idx + 2,
                        "SKU": sku_key,
                        "Columna precio": col_precio,
                        "Valor original": fila[col_precio],
                    })

            filas_origen.append(registro)

    if not filas_origen:
        raise ValueError(
            "No se pudo armar la base de actualización. "
            "Revisá que el archivo tenga SKU, AStk y columnas de precio sin _2."
        )

    df_origen = pd.DataFrame(filas_origen)

    duplicados = df_origen[df_origen["_SKU_KEY"].duplicated(keep=False)].copy()

    df_unico = df_origen.drop_duplicates(
        subset="_SKU_KEY",
        keep="first"
    ).copy()

    datos_por_sku = df_unico.set_index("_SKU_KEY").to_dict(orient="index")

    return (
        datos_por_sku,
        duplicados,
        control_hojas_ignoradas,
        control_precios_invalidos,
        control_stock_invalido
    )


# ============================================================
# PROCESAR INTEGRALY
# ============================================================

def procesar_archivos(integraly_bytes, actualizacion_bytes):

    (
        datos_por_sku,
        duplicados,
        hojas_ignoradas,
        precios_invalidos,
        stock_invalido
    ) = armar_base_actualizacion(actualizacion_bytes)

    wb = load_workbook(BytesIO(integraly_bytes))

    if NOMBRE_HOJA_CONTROL in wb.sheetnames:
        del wb[NOMBRE_HOJA_CONTROL]

    control_resumen = []
    control_sku_sin_match = []
    control_cuotas_no_reconocidas = []
    control_precio_no_actualizado = []
    control_stock_no_actualizado = []
    control_estado_actualizado = []
    control_estado_no_actualizado = []
    control_mayor_umbral_pausada = []

    for ws in wb.worksheets:

        if ws.title.upper().startswith("CONTROL"):
            continue

        headers = obtener_headers_ws(ws)

        col_sku_nombre = detectar_columna(
            headers.keys(),
            POSIBLES_SKU,
            obligatorio=False,
            descripcion="columna SKU destino"
        )

        if col_sku_nombre is None:
            control_resumen.append({
                "Hoja destino": ws.title,
                "Estado hoja": "Ignorada",
                "Motivo": "No se encontró columna SKU",
                "Filas procesadas": 0,
                "Filas actualizadas con precio": 0,
                "SKU sin match": 0,
                "Cuotas no reconocidas": 0,
                "Precio no actualizado": 0,
                "Stock no actualizado": 0,
                "Estado actualizado": 0,
                "Estado no actualizado": 0,
                "Pausadas por superar 1.500.000": 0,
            })
            continue

        col_mla_nombre = detectar_columna(
            headers.keys(),
            POSIBLES_MLA,
            obligatorio=False,
            descripcion="columna MLA"
        )

        col_cuotas_nombre = detectar_columna(
            headers.keys(),
            POSIBLES_CUOTAS,
            obligatorio=False,
            descripcion="columna cuotas"
        )

        col_precio_nombre = detectar_columna(
            headers.keys(),
            POSIBLES_PRECIO_DESTINO,
            obligatorio=False,
            descripcion="columna precio destino"
        )

        col_stock_nombre = detectar_columna(
            headers.keys(),
            POSIBLES_STOCK_DESTINO,
            obligatorio=False,
            descripcion="columna stock destino"
        )

        col_estado_nombre = detectar_columna(
            headers.keys(),
            POSIBLES_ESTADO_DESTINO,
            obligatorio=False,
            descripcion="columna estado destino"
        )

        if col_precio_nombre is None:
            col_precio_nombre = "precio"
            obtener_o_crear_columna(ws, headers, col_precio_nombre)

        if col_stock_nombre is None:
            col_stock_nombre = "cantidad"
            obtener_o_crear_columna(ws, headers, col_stock_nombre)

        if col_estado_nombre is None:
            col_estado_nombre = "estado"
            obtener_o_crear_columna(ws, headers, col_estado_nombre)

        col_sku = headers[col_sku_nombre]
        col_precio = headers[col_precio_nombre]
        col_stock = headers[col_stock_nombre]
        col_estado = headers[col_estado_nombre]

        col_mla = headers[col_mla_nombre] if col_mla_nombre else None
        col_cuotas = headers[col_cuotas_nombre] if col_cuotas_nombre else None

        filas_procesadas = 0
        filas_actualizadas = 0
        sku_sin_match = 0
        cuotas_no_reconocidas = 0
        precio_no_actualizado = 0
        stock_no_actualizado = 0
        estado_actualizado = 0
        estado_no_actualizado = 0
        pausadas_umbral = 0

        for row in range(2, ws.max_row + 1):

            sku_original = ws.cell(row=row, column=col_sku).value
            sku_key = normalizar_sku(sku_original)

            if sku_key == "":
                continue

            filas_procesadas += 1

            mla = ws.cell(row=row, column=col_mla).value if col_mla else ""
            cuotas_original = ws.cell(row=row, column=col_cuotas).value if col_cuotas else ""
            estado_anterior = ws.cell(row=row, column=col_estado).value

            if sku_key not in datos_por_sku:
                sku_sin_match += 1
                control_sku_sin_match.append({
                    "Hoja destino": ws.title,
                    "Fila destino": row,
                    "MLA": mla,
                    "SKU": sku_original,
                    "Cuotas": cuotas_original,
                    "Estado anterior": estado_anterior,
                })
                continue

            datos = datos_por_sku[sku_key]

            stock_final = datos.get("_ASTK")

            if stock_final is not None:
                ws.cell(row=row, column=col_stock).value = stock_final
            else:
                stock_no_actualizado += 1
                control_stock_no_actualizado.append({
                    "Hoja destino": ws.title,
                    "Fila destino": row,
                    "MLA": mla,
                    "SKU": sku_original,
                    "Valor AStk": stock_final,
                    "Hoja origen": datos.get("_HOJA_ORIGEN"),
                    "Fila origen": datos.get("_FILA_ORIGEN"),
                })

            clave_precio = clasificar_cuotas(cuotas_original)

            if clave_precio is None:
                cuotas_no_reconocidas += 1
                estado_no_actualizado += 1

                control_cuotas_no_reconocidas.append({
                    "Hoja destino": ws.title,
                    "Fila destino": row,
                    "MLA": mla,
                    "SKU": sku_original,
                    "Cuotas": cuotas_original,
                })

                control_estado_no_actualizado.append({
                    "Hoja destino": ws.title,
                    "Fila destino": row,
                    "MLA": mla,
                    "SKU": sku_original,
                    "Precio final": "",
                    "Stock final": stock_final,
                    "Estado anterior": estado_anterior,
                    "Motivo": "Cuotas no reconocidas, no se pudo determinar precio",
                })

                continue

            precio_final = datos.get(clave_precio)

            if precio_final is not None:
                ws.cell(row=row, column=col_precio).value = precio_final
                filas_actualizadas += 1
            else:
                precio_no_actualizado += 1
                control_precio_no_actualizado.append({
                    "Hoja destino": ws.title,
                    "Fila destino": row,
                    "MLA": mla,
                    "SKU": sku_original,
                    "Cuotas": cuotas_original,
                    "Precio requerido": COLUMNAS_PRECIO_ORIGEN[clave_precio],
                    "Hoja origen": datos.get("_HOJA_ORIGEN"),
                    "Fila origen": datos.get("_FILA_ORIGEN"),
                })

            estado_final, motivo_estado = determinar_estado(precio_final, stock_final)

            if estado_final is not None:
                ws.cell(row=row, column=col_estado).value = estado_final
                estado_actualizado += 1

                registro_estado = {
                    "Hoja destino": ws.title,
                    "Fila destino": row,
                    "MLA": mla,
                    "SKU": sku_original,
                    "Precio final": precio_final,
                    "Stock final": stock_final,
                    "Estado anterior": estado_anterior,
                    "Estado final": estado_final,
                    "Motivo": motivo_estado,
                }

                control_estado_actualizado.append(registro_estado)

                if float(precio_final) > UMBRAL_PRECIO_ESTADO:
                    pausadas_umbral += 1
                    control_mayor_umbral_pausada.append(registro_estado)

            else:
                estado_no_actualizado += 1
                control_estado_no_actualizado.append({
                    "Hoja destino": ws.title,
                    "Fila destino": row,
                    "MLA": mla,
                    "SKU": sku_original,
                    "Precio final": precio_final,
                    "Stock final": stock_final,
                    "Estado anterior": estado_anterior,
                    "Motivo": motivo_estado,
                })

        control_resumen.append({
            "Hoja destino": ws.title,
            "Estado hoja": "Procesada",
            "Motivo": "",
            "Filas procesadas": filas_procesadas,
            "Filas actualizadas con precio": filas_actualizadas,
            "SKU sin match": sku_sin_match,
            "Cuotas no reconocidas": cuotas_no_reconocidas,
            "Precio no actualizado": precio_no_actualizado,
            "Stock no actualizado": stock_no_actualizado,
            "Estado actualizado": estado_actualizado,
            "Estado no actualizado": estado_no_actualizado,
            "Pausadas por superar 1.500.000": pausadas_umbral,
        })

    # ========================================================
    # HOJA DE CONTROL
    # ========================================================

    ws_control = wb.create_sheet(NOMBRE_HOJA_CONTROL)

    ws_control.append(["Fecha proceso", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws_control.append(["Suma fija aplicada a cada precio", SUMA_FIJA_PRECIO])
    ws_control.append(["Umbral precio estado", UMBRAL_PRECIO_ESTADO])
    ws_control.append(["Estado activa", ESTADO_ACTIVA])
    ws_control.append(["Estado pausada", ESTADO_PAUSADA])
    ws_control.append([])

    ws_control.append(["MAPEO DE CUOTAS"])
    ws_control.append(["Condición en Integraly", "Columna usada en Global"])
    ws_control.append(["Sin cuotas / No Agregar Cuotas / vacío", "Precio ML Clasica + 12000"])
    ws_control.append(["3 cuotas", "Precio ML Premium + 12000"])
    ws_control.append(["6 cuotas", "Precio ML Premium 6c + 12000"])
    ws_control.append(["9 cuotas", "Precio ML Premium 9c + 12000"])
    ws_control.append(["12 cuotas", "Precio ML Premium 12c + 12000"])
    ws_control.append(["Stock", "AStk"])
    ws_control.append([])

    ws_control.append(["REGLA DE ESTADO"])
    ws_control.append(["Condición", "Estado asignado"])
    ws_control.append(["Precio mayor a 1.500.000", ESTADO_PAUSADA])
    ws_control.append(["Precio menor o igual a 1.500.000 y stock >= 1", ESTADO_ACTIVA])
    ws_control.append(["Precio menor o igual a 1.500.000 y stock <= 0", ESTADO_PAUSADA])
    ws_control.append([])

    ws_control.append(["RESUMEN"])

    resumen_cols = [
        "Hoja destino",
        "Estado hoja",
        "Motivo",
        "Filas procesadas",
        "Filas actualizadas con precio",
        "SKU sin match",
        "Cuotas no reconocidas",
        "Precio no actualizado",
        "Stock no actualizado",
        "Estado actualizado",
        "Estado no actualizado",
        "Pausadas por superar 1.500.000"
    ]

    ws_control.append(resumen_cols)

    for r in control_resumen:
        ws_control.append([r.get(c, "") for c in resumen_cols])

    agregar_tabla(
        ws_control,
        "SKU SIN MATCH",
        ["Hoja destino", "Fila destino", "MLA", "SKU", "Cuotas", "Estado anterior"],
        control_sku_sin_match,
        ["Sin SKU sin match", "", "", "", "", ""]
    )

    agregar_tabla(
        ws_control,
        "CUOTAS NO RECONOCIDAS",
        ["Hoja destino", "Fila destino", "MLA", "SKU", "Cuotas"],
        control_cuotas_no_reconocidas,
        ["Sin cuotas no reconocidas", "", "", "", ""]
    )

    agregar_tabla(
        ws_control,
        "PRECIO NO ACTUALIZADO",
        ["Hoja destino", "Fila destino", "MLA", "SKU", "Cuotas", "Precio requerido", "Hoja origen", "Fila origen"],
        control_precio_no_actualizado,
        ["Sin precios pendientes", "", "", "", "", "", "", ""]
    )

    agregar_tabla(
        ws_control,
        "STOCK NO ACTUALIZADO",
        ["Hoja destino", "Fila destino", "MLA", "SKU", "Valor AStk", "Hoja origen", "Fila origen"],
        control_stock_no_actualizado,
        ["Sin stocks pendientes", "", "", "", "", "", ""]
    )

    agregar_tabla(
        ws_control,
        "ESTADO ACTUALIZADO",
        ["Hoja destino", "Fila destino", "MLA", "SKU", "Precio final", "Stock final", "Estado anterior", "Estado final", "Motivo"],
        control_estado_actualizado,
        ["Sin estados actualizados", "", "", "", "", "", "", "", ""]
    )

    agregar_tabla(
        ws_control,
        "ESTADO NO ACTUALIZADO",
        ["Hoja destino", "Fila destino", "MLA", "SKU", "Precio final", "Stock final", "Estado anterior", "Motivo"],
        control_estado_no_actualizado,
        ["Sin estados pendientes", "", "", "", "", "", "", ""]
    )

    agregar_tabla(
        ws_control,
        "PAUSADAS POR SUPERAR 1.500.000",
        ["Hoja destino", "Fila destino", "MLA", "SKU", "Precio final", "Stock final", "Estado anterior", "Estado final", "Motivo"],
        control_mayor_umbral_pausada,
        ["Sin publicaciones mayores a 1.500.000 pausadas", "", "", "", "", "", "", "", ""]
    )

    ws_control.append([])
    ws_control.append(["SKU DUPLICADOS EN GLOBAL / ACTUALIZACIÓN PRECIO"])
    ws_control.append(["SKU", "Cantidad apariciones", "Hojas origen"])

    if not duplicados.empty:
        dup_resumen = (
            duplicados
            .groupby("_SKU_KEY")
            .agg(
                cantidad_apariciones=("_SKU_KEY", "size"),
                hojas_origen=("_HOJA_ORIGEN", lambda x: " | ".join(sorted(set(map(str, x)))))
            )
            .reset_index()
        )

        for _, r in dup_resumen.iterrows():
            ws_control.append([
                r["_SKU_KEY"],
                int(r["cantidad_apariciones"]),
                r["hojas_origen"]
            ])
    else:
        ws_control.append(["Sin duplicados", 0, ""])

    agregar_tabla(
        ws_control,
        "HOJAS IGNORADAS EN GLOBAL / ACTUALIZACIÓN PRECIO",
        ["Hoja", "Motivo"],
        hojas_ignoradas,
        ["Sin hojas ignoradas", ""]
    )

    agregar_tabla(
        ws_control,
        "PRECIOS INVÁLIDOS EN GLOBAL / ACTUALIZACIÓN PRECIO",
        ["Hoja origen", "Fila origen", "SKU", "Columna precio", "Valor original"],
        precios_invalidos,
        ["Sin precios inválidos", "", "", "", ""]
    )

    agregar_tabla(
        ws_control,
        "STOCK INVÁLIDO EN GLOBAL / ACTUALIZACIÓN PRECIO",
        ["Hoja origen", "Fila origen", "SKU", "Valor AStk"],
        stock_invalido,
        ["Sin stocks inválidos", "", "", ""]
    )

    for col in ws_control.columns:
        max_len = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))

        ws_control.column_dimensions[col_letter].width = min(max_len + 2, 45)

    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)

    resumen = {
        "control_resumen": control_resumen,
        "sku_unicos_origen": len(datos_por_sku),
        "sku_duplicados_origen": duplicados["_SKU_KEY"].nunique() if not duplicados.empty else 0,
        "hojas_ignoradas_origen": len(hojas_ignoradas),
    }

    return salida, resumen


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Cruce Integraly - Global Precios",
    page_icon="📦",
    layout="wide"
)

st.title("Cruce Integraly vs Global Precios")
st.caption("Actualiza precio, stock y estado por SKU.")

st.warning(
    "La app toma los precios del archivo Global, limpia ARS, convierte a número "
    "y suma $12.000 fijos antes de completar Integraly."
)

st.info(
    "Mapeo: Clásica = sin cuotas | Premium = 3 cuotas | Premium 6c = 6 cuotas | "
    "Premium 9c = 9 cuotas | Premium 12c = 12 cuotas."
)

col1, col2 = st.columns(2)

with col1:
    archivo_integraly = st.file_uploader(
        "Subí el archivo Integraly",
        type=["xlsx"],
        key="integraly"
    )

with col2:
    archivo_actualizacion = st.file_uploader(
        "Subí el archivo Global / Actualización Precio",
        type=["xlsx"],
        key="actualizacion"
    )

if st.button("Procesar cruce", type="primary"):

    if archivo_integraly is None or archivo_actualizacion is None:
        st.error("Tenés que subir los 2 archivos para procesar.")

    else:
        try:
            with st.spinner("Procesando archivos..."):
                salida, resumen = procesar_archivos(
                    archivo_integraly.getvalue(),
                    archivo_actualizacion.getvalue()
                )

            st.success("Proceso finalizado correctamente.")

            st.subheader("Resumen")
            st.dataframe(
                pd.DataFrame(resumen["control_resumen"]),
                use_container_width=True
            )

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("SKU únicos origen", resumen["sku_unicos_origen"])
            c2.metric("SKU duplicados origen", resumen["sku_duplicados_origen"])
            c3.metric("Hojas ignoradas origen", resumen["hojas_ignoradas_origen"])
            c4.metric("Suma fija aplicada", f"${SUMA_FIJA_PRECIO:,.0f}")

            fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_salida = f"INTEGRALY_PRECIOS_STOCK_ESTADO_ACTUALIZADO_{fecha}.xlsx"

            st.download_button(
                label="Descargar Excel actualizado",
                data=salida,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error("El proceso falló.")
            st.exception(e)
