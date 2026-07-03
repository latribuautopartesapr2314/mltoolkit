import pandas as pd
from openpyxl import load_workbook
from google.colab import files
from pathlib import Path
import unicodedata
import re

# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

NOMBRE_HOJA_CONFIG = "CONFIG_PUBLICADOR"

HOJAS_AGENTE_BASE = [
    "BASE_COMPLETAR",
    "BASE_TITULOS",
    "BASE_DESCRIPCIONES",
    "BASE_IMAGENES",
]

LIMPIAR_FILAS_PUBLICAR_ANTES_DE_CARGAR = True
DESCARGAR_AL_FINAL = True

# ============================================================
# SUBIR ARCHIVOS
# ============================================================

print("Subí los 2 archivos:")
print("1) AGENTE PUBLICADOR BASE*.xlsx")
print("2) Publicar*.xlsx")

uploaded = files.upload()

if len(uploaded) < 2:
    raise ValueError("Tenés que subir el archivo AGENTE PUBLICADOR y el archivo Publicar.")

paths = [Path(nombre) for nombre in uploaded.keys()]

# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def normalizar(txt):
    if txt is None:
        return ""
    txt = str(txt).strip()
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    txt = re.sub(r"\s+", " ", txt)
    return txt.lower()

def es_vacio(valor):
    try:
        if pd.isna(valor):
            return True
    except Exception:
        pass
    return valor is None or str(valor).strip() == ""

def buscar_hoja(wb, nombre_buscado):
    buscado = normalizar(nombre_buscado)
    for hoja in wb.sheetnames:
        if normalizar(hoja) == buscado:
            return hoja
    return None

def detectar_archivos(paths):
    agente = None
    publicar = None

    hojas_agente_norm = {normalizar(h) for h in HOJAS_AGENTE_BASE}

    for path in paths:
        wb_temp = load_workbook(path, read_only=True, data_only=False)
        hojas_norm = {normalizar(h) for h in wb_temp.sheetnames}
        wb_temp.close()

        if hojas_agente_norm.issubset(hojas_norm):
            agente = path
        elif "publicar" in normalizar(path.name):
            publicar = path

    if agente is None:
        for path in paths:
            if "agente" in normalizar(path.name):
                agente = path
                break

    if publicar is None:
        candidatos = [p for p in paths if p != agente]
        if len(candidatos) == 1:
            publicar = candidatos[0]

    if agente is None:
        raise ValueError("No pude detectar el archivo AGENTE PUBLICADOR.")

    if publicar is None:
        raise ValueError("No pude detectar el archivo Publicar.")

    return agente, publicar

def leer_config_publicador(wb_agente):
    hoja_config = buscar_hoja(wb_agente, NOMBRE_HOJA_CONFIG)

    if hoja_config is None:
        raise ValueError(
            f"No existe la hoja '{NOMBRE_HOJA_CONFIG}'. "
            "Para que el sistema sea escalable, agregá esa hoja al AGENTE PUBLICADOR BASE."
        )

    ws = wb_agente[hoja_config]

    headers = {}
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(1, col).value
        if not es_vacio(valor):
            headers[normalizar(valor)] = col

    columnas_requeridas = {
        "activo": None,
        "categoria": None,
        "hoja_publicar": None,
        "fila_modelo_base_completar": None,
    }

    for requerida in columnas_requeridas:
        if requerida not in headers:
            raise ValueError(f"Falta la columna obligatoria '{requerida}' en CONFIG_PUBLICADOR.")

    config = []

    for row in range(2, ws.max_row + 1):
        activo = ws.cell(row, headers["activo"]).value
        categoria = ws.cell(row, headers["categoria"]).value
        hoja_publicar = ws.cell(row, headers["hoja_publicar"]).value
        fila_modelo = ws.cell(row, headers["fila_modelo_base_completar"]).value

        if es_vacio(categoria):
            continue

        if normalizar(activo) not in ["si", "sí", "s", "yes", "true", "1"]:
            continue

        if es_vacio(hoja_publicar):
            hoja_publicar = categoria

        if es_vacio(fila_modelo):
            raise ValueError(f"La categoría '{categoria}' no tiene FILA_MODELO_BASE_COMPLETAR definida.")

        fila_encabezados = 3
        fila_inicio = 8

        if "fila_encabezados_publicar" in headers:
            valor = ws.cell(row, headers["fila_encabezados_publicar"]).value
            if not es_vacio(valor):
                fila_encabezados = int(valor)

        if "fila_inicio_publicar" in headers:
            valor = ws.cell(row, headers["fila_inicio_publicar"]).value
            if not es_vacio(valor):
                fila_inicio = int(valor)

        config.append({
            "categoria": str(categoria).strip(),
            "hoja_publicar": str(hoja_publicar).strip(),
            "fila_modelo": int(fila_modelo),
            "fila_encabezados": int(fila_encabezados),
            "fila_inicio": int(fila_inicio),
        })

    if not config:
        raise ValueError("CONFIG_PUBLICADOR no tiene categorías activas para procesar.")

    return config

def cargar_df_agente(path_agente):
    return pd.read_excel(
        path_agente,
        sheet_name=None,
        header=None,
        dtype=object,
        engine="openpyxl"
    )

def obtener_df(dfs, nombre_hoja):
    objetivo = normalizar(nombre_hoja)
    for hoja, df in dfs.items():
        if normalizar(hoja) == objetivo:
            return df
    raise ValueError(f"No se encontró la hoja '{nombre_hoja}' en el AGENTE.")

def extraer_categorias_de_config(config):
    return [item["categoria"] for item in config]

def buscar_fila_categoria(df, categoria):
    objetivo = normalizar(categoria)

    for idx in range(len(df)):
        valor = df.iat[idx, 0]
        if normalizar(valor) == objetivo:
            return idx

    return None

def extraer_titulos(df_titulos, categoria, categorias_validas):
    fila_categoria = buscar_fila_categoria(df_titulos, categoria)

    if fila_categoria is None:
        return []

    categorias_norm = {normalizar(c) for c in categorias_validas}
    titulos = []

    for idx in range(fila_categoria + 1, len(df_titulos)):
        valor = df_titulos.iat[idx, 0]

        if es_vacio(valor):
            break

        texto = str(valor).strip()
        texto_norm = normalizar(texto)

        if texto_norm in categorias_norm:
            break

        titulos.append(texto)

    return titulos

def es_marcador_descripcion(valor):
    valor_norm = normalizar(valor)
    return valor_norm.startswith("descripcion")

def extraer_descripcion(df_descripciones, categoria, categorias_validas):
    categorias_norm = {normalizar(c) for c in categorias_validas}

    # Caso recomendado:
    # Categoria
    # DESCRIPCION 1
    # texto...
    fila_categoria = buscar_fila_categoria(df_descripciones, categoria)

    if fila_categoria is None:
        return ""

    lineas = []
    empezo = False
    encontro_marcador = False

    for idx in range(fila_categoria + 1, len(df_descripciones)):
        valor = df_descripciones.iat[idx, 0]

        if es_vacio(valor):
            if empezo:
                lineas.append("")
            continue

        texto = str(valor).strip()
        texto_norm = normalizar(texto)

        if texto_norm in categorias_norm:
            break

        if es_marcador_descripcion(texto):
            if encontro_marcador and empezo:
                break
            encontro_marcador = True
            continue

        lineas.append(texto)
        empezo = True

    descripcion = "\n".join(lineas).strip()

    descripcion = re.sub(r"\n{3,}", "\n\n", descripcion)

    return descripcion

def extraer_imagenes(df_imagenes, categoria, categorias_validas):
    fila_categoria = buscar_fila_categoria(df_imagenes, categoria)

    if fila_categoria is None:
        return ""

    categorias_norm = {normalizar(c) for c in categorias_validas}

    for idx in range(fila_categoria + 1, len(df_imagenes)):
        valor = df_imagenes.iat[idx, 0]

        if es_vacio(valor):
            continue

        texto = str(valor).strip()
        texto_norm = normalizar(texto)

        if texto_norm in categorias_norm:
            break

        if "http" in texto_norm:
            return texto

    return ""

def buscar_columna_por_encabezado(ws, fila_encabezados, palabras_obligatorias):
    for col in range(1, ws.max_column + 1):
        encabezado = normalizar(ws.cell(fila_encabezados, col).value)
        if all(palabra in encabezado for palabra in palabras_obligatorias):
            return col
    return None

def validar_columna(columna, nombre, hoja):
    if columna is None:
        raise ValueError(f"No se encontró la columna '{nombre}' en la hoja Publicar '{hoja}'.")

def columnas_que_no_deben_tocarse(ws, fila_encabezados):
    columnas = set()

    for col in range(1, ws.max_column + 1):
        encabezado = normalizar(ws.cell(fila_encabezados, col).value)

        if "buybox_formula" in encabezado:
            columnas.add(col)

        if "hidden_pictures" in encabezado:
            columnas.add(col)

    return columnas

def limpiar_filas_publicar(ws, fila_inicio, columnas_saltar):
    for fila in range(fila_inicio, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            if col in columnas_saltar:
                continue
            ws.cell(fila, col).value = None

def copiar_fila_modelo(ws_origen, ws_destino, fila_origen, fila_destino, columnas_saltar):
    max_col = min(ws_origen.max_column, ws_destino.max_column)

    for col in range(1, max_col + 1):
        if col in columnas_saltar:
            continue
        ws_destino.cell(fila_destino, col).value = ws_origen.cell(fila_origen, col).value

def contar_links_imagenes(texto):
    if es_vacio(texto):
        return 0
    partes = [x.strip() for x in str(texto).split(",")]
    return len([x for x in partes if x.startswith("http")])

# ============================================================
# DETECCIÓN Y LECTURA
# ============================================================

path_agente, path_publicar = detectar_archivos(paths)

print("\nArchivo AGENTE detectado:")
print(path_agente.name)

print("\nArchivo PUBLICAR detectado:")
print(path_publicar.name)

dfs_agente = cargar_df_agente(path_agente)

wb_agente = load_workbook(path_agente, data_only=False)
wb_publicar = load_workbook(path_publicar, data_only=False)

print("\nHojas detectadas en AGENTE:")
for hoja in wb_agente.sheetnames:
    print("-", hoja)

print("\nHojas detectadas en PUBLICAR:")
for hoja in wb_publicar.sheetnames:
    print("-", hoja)

config = leer_config_publicador(wb_agente)

df_titulos = obtener_df(dfs_agente, "BASE_TITULOS")
df_descripciones = obtener_df(dfs_agente, "BASE_DESCRIPCIONES")
df_imagenes = obtener_df(dfs_agente, "BASE_IMAGENES")

hoja_base_completar = buscar_hoja(wb_agente, "BASE_COMPLETAR")
if hoja_base_completar is None:
    raise ValueError("No se encontró BASE_COMPLETAR en el AGENTE.")

ws_base_completar = wb_agente[hoja_base_completar]

categorias_validas = extraer_categorias_de_config(config)

# ============================================================
# PROCESAMIENTO
# ============================================================

resumen = []
errores = []

for item in config:
    categoria = item["categoria"]
    hoja_publicar_objetivo = item["hoja_publicar"]
    fila_modelo = item["fila_modelo"]
    fila_encabezados = item["fila_encabezados"]
    fila_inicio = item["fila_inicio"]

    print("\nProcesando categoría:", categoria)

    hoja_publicar_real = buscar_hoja(wb_publicar, hoja_publicar_objetivo)

    if hoja_publicar_real is None:
        errores.append({
            "categoria": categoria,
            "error": f"No existe la hoja '{hoja_publicar_objetivo}' en Publicar."
        })
        continue

    ws_publicar = wb_publicar[hoja_publicar_real]

    if fila_modelo > ws_base_completar.max_row:
        errores.append({
            "categoria": categoria,
            "error": f"La fila modelo {fila_modelo} no existe en BASE_COMPLETAR."
        })
        continue

    titulos = extraer_titulos(df_titulos, categoria, categorias_validas)
    descripcion = extraer_descripcion(df_descripciones, categoria, categorias_validas)
    imagenes = extraer_imagenes(df_imagenes, categoria, categorias_validas)

    if len(titulos) == 0:
        errores.append({
            "categoria": categoria,
            "error": "No se encontraron títulos debajo de la categoría en BASE_TITULOS."
        })
        continue

    if es_vacio(descripcion):
        errores.append({
            "categoria": categoria,
            "error": "No se encontró descripción en BASE_DESCRIPCIONES."
        })
        continue

    if es_vacio(imagenes):
        errores.append({
            "categoria": categoria,
            "error": "No se encontraron links de imágenes en BASE_IMAGENES."
        })
        continue

    col_titulo = buscar_columna_por_encabezado(ws_publicar, fila_encabezados, ["titulo"])
    col_caracteres = buscar_columna_por_encabezado(ws_publicar, fila_encabezados, ["cantidad", "caracteres"])
    col_fotos = buscar_columna_por_encabezado(ws_publicar, fila_encabezados, ["fotos"])
    col_descripcion = buscar_columna_por_encabezado(ws_publicar, fila_encabezados, ["descripcion"])

    try:
        validar_columna(col_titulo, "Título", hoja_publicar_real)
        validar_columna(col_fotos, "Fotos", hoja_publicar_real)
        validar_columna(col_descripcion, "Descripción", hoja_publicar_real)
    except Exception as e:
        errores.append({
            "categoria": categoria,
            "error": str(e)
        })
        continue

    ultima_fila_necesaria = fila_inicio + len(titulos) - 1

    if ultima_fila_necesaria > ws_publicar.max_row:
        errores.append({
            "categoria": categoria,
            "error": (
                f"La hoja '{hoja_publicar_real}' no tiene filas suficientes. "
                f"Necesita hasta fila {ultima_fila_necesaria}, pero llega hasta {ws_publicar.max_row}."
            )
        })
        continue

    columnas_saltar = columnas_que_no_deben_tocarse(ws_publicar, fila_encabezados)

    if LIMPIAR_FILAS_PUBLICAR_ANTES_DE_CARGAR:
        limpiar_filas_publicar(ws_publicar, fila_inicio, columnas_saltar)

    for i, titulo in enumerate(titulos):
        fila_destino = fila_inicio + i

        copiar_fila_modelo(
            ws_origen=ws_base_completar,
            ws_destino=ws_publicar,
            fila_origen=fila_modelo,
            fila_destino=fila_destino,
            columnas_saltar=columnas_saltar
        )

        ws_publicar.cell(fila_destino, col_titulo).value = titulo
        ws_publicar.cell(fila_destino, col_fotos).value = imagenes
        ws_publicar.cell(fila_destino, col_descripcion).value = descripcion

        if col_caracteres is not None:
            ws_publicar.cell(fila_destino, col_caracteres).value = len(str(titulo))

    resumen.append({
        "categoria": categoria,
        "hoja_publicar": hoja_publicar_real,
        "estado": "OK",
        "titulos_generados": len(titulos),
        "fila_modelo_base_completar": fila_modelo,
        "fila_inicio_publicar": fila_inicio,
        "col_titulo": col_titulo,
        "col_fotos": col_fotos,
        "col_descripcion": col_descripcion,
        "links_imagenes": contar_links_imagenes(imagenes),
        "caracteres_descripcion": len(descripcion),
    })

# ============================================================
# CONTROL DE ERRORES
# ============================================================

print("\nVER CONTROL - CATEGORÍAS PROCESADAS")
if resumen:
    for r in resumen:
        print(r)
else:
    print("No se procesó ninguna categoría.")

print("\nVER CONTROL - ERRORES / OMITIDAS")
if errores:
    for e in errores:
        print(e)
else:
    print("Sin errores.")

if errores:
    raise ValueError(
        "El proceso encontró errores. No se guarda el archivo para evitar una carga incompleta o incorrecta."
    )

# ============================================================
# GUARDAR SOBRE EL MISMO PUBLICAR
# ============================================================

wb_publicar.save(path_publicar)

print("\nArchivo Publicar actualizado correctamente")

if DESCARGAR_AL_FINAL:
    files.download(str(path_publicar))
