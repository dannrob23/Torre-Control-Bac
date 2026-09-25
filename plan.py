"""
plan.py - Analitica del "Plan de Trabajo" (cartera vencida y envejecimiento).

Este modulo es INDEPENDIENTE del nucleo SLA (core.py). Responde una pregunta
distinta: no "que casos se vencen ahora", sino "como esta envejeciendo la
cartera y como se reparte entre los tecnicos".

Fuente de datos
---------------
El archivo mensual "Plan de Trabajo" (p. ej. "Septiembre_2026_Plan de
Trabajo.xlsx"). Se usan DOS tipos de hoja:

  * ``Casos_Ven``      -> el listado de casos vencidos del mes (una fila por caso).
  * ``N_Septiembre``   -> la "cosecha" operativa de cada dia; aporta el
                          ``Vencimiento`` (ANS) y el ``Estado`` de cada caso.

PROBLEMA GRAVE QUE RESUELVE ESTE MODULO: las fechas invertidas
--------------------------------------------------------------
Excel convirtio a fecha nativa los textos colombianos ``DD/MM/AAAA`` aplicando
el formato ``m/d/yy``. Resultado: cuando el dia era <= 12, mes y dia quedaron
INTERCAMBIADOS (el 1 de septiembre se guardo como 9 de enero).

No se puede saber cual fecha esta invertida mirando la celda: ``09/01/2026`` es
ambigua. El desempate se obtiene de la SECUENCIA DE IDs DE CASO, que es
monotona: el ID maximo presente en cada hoja diaria crece ~100 por dia. Con ese
modelo se decide, caso por caso, cual de las dos lecturas es coherente.

Medido sobre el archivo real de septiembre 2026: la lectura dia/mes da un error
medio de 3.6 dias contra el modelo de IDs, frente a 75.7 dias de la lectura
mes/dia. Es decir, la lectura correcta es dia/mes y hay que invertir 71 de 143
fechas. Sin esta correccion, la antiguedad promedio sale ~4x inflada (83 dias en
lugar de 22) y aparecen casos "creados en el futuro".

Este modulo NO modifica el archivo de entrada: devuelve el DataFrame corregido
y deja constancia de cuantas fechas toco y por que.
"""

from __future__ import annotations

import hashlib
import os
import re
import threading
import unicodedata

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Nombres de hojas y columnas
# ---------------------------------------------------------------------------

HOJA_VENCIDOS = "Casos_Ven"

# Las hojas diarias se llaman "1_Septiembre", "17_Septimbre" (typo real),
# "31_ Agosto" (con espacio). Por eso se detectan por expresion regular y no
# por nombre exacto.
PATRON_HOJA_DIARIA = re.compile(r"^\s*(\d{1,2})\s*[_\s]\s*([A-Za-zÁÉÍÓÚÑáéíóúñ]+)\s*$")

MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

# El archivo del plan es mensual y su nombre de hoja no trae el anio ("17_Septiembre"),
# asi que el anio se toma de aqui.
ANIO_DEFECTO = 2026

# Columnas de Casos_Ven (los encabezados reales traen espacios sobrantes).
COL_MES = "MES"
COL_CASO = "CASO"
COL_UBICACION = "UBICACION"
COL_FECHA_CREACION = "FECHA DE CREACION"
COL_TECNICO = "TECNICO ENCARGADO"
COL_JUSTIFICACION = "JUSTIFICACION DE VENCIMIENTO"
COL_CULPA = "CULPA"
COL_CATEGORIA = "CATEGORIA"  # columna sin encabezado en el archivo original

COL_ID_DIARIO = "ID de incidente"
COL_VENCIMIENTO = "Vencimiento"
COL_APERTURA = "Fecha/hora de apertura"
# La hoja diaria llama 'Estado' a su columna; la hoja de vencidos usa 'ESTADO'.
# Son columnas distintas y por eso llevan constantes distintas.
COL_ESTADO_DIARIO = "Estado"
COL_ASIGNATARIO = "Asignatario"
COL_UBICACION_DIARIA = "Ubicación"
# El campo ATIENDE distingue Bogota de regional: es la unica senal de zona
# que trae la hoja diaria (el nombre de usuario no dice de donde es).
COL_ATIENDE = "ATIENDE"



# Tolerancia con la que una fecha se considera "coherente" con el modelo de IDs.
# 7 dias absorbe el ruido del ajuste lineal (los IDs no crecen exactamente igual
# todos los dias) sin permitir que una fecha de otro mes pase por buena.
TOLERANCIA_MODELO_DIAS = 7.0



# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------

def normalizar(valor) -> str:
    """
    Normaliza texto para comparar nombres y categorias:
    trim (incluye NBSP), colapsa espacios, quita tildes y pasa a mayusculas.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    texto = str(valor).replace("\xa0", " ").strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.upper().split())


def _clave_columna(valor) -> str:
    """Encabezado comparable: sin espacios sobrantes ni mayusculas."""
    return normalizar(valor).rstrip(":").strip()


def _buscar_columna(df: pd.DataFrame, nombres: list[str]) -> str | None:
    """
    Devuelve el nombre real de la primera columna que coincida con alguno de
    ``nombres``, comparando de forma normalizada. Sirve para tolerar acentos,
    espacios sobrantes y mayusculas en los encabezados.
    """
    indice = {_clave_columna(c): c for c in df.columns}
    for nombre in nombres:
        real = indice.get(_clave_columna(nombre))
        if real is not None:
            return real
    return None


# ---------------------------------------------------------------------------
# Deteccion de hojas
# ---------------------------------------------------------------------------

PATRON_ANIO = re.compile(r"(\d{4})")


def inferir_anio(crudo: pd.DataFrame, columna: str = COL_FECHA_CREACION) -> int:
    """
    Deduce el anio del plan a partir de las fechas del propio archivo.

    El nombre de las hojas diarias no trae el anio ("17_Septiembre"), asi que se
    toma de las fechas de creacion: se busca un anio de 4 digitos en el texto o,
    si la celda es una fecha nativa, su propio anio. Se usa el valor mas
    frecuente, que es robusto frente a celdas corruptas.

    Esto evita tener que actualizar una constante cada enero.
    """
    if columna not in crudo.columns:
        return ANIO_DEFECTO

    anios: list[int] = []
    for valor in crudo[columna].dropna():
        if isinstance(valor, pd.Timestamp):
            anios.append(valor.year)
            continue
        coincidencias = PATRON_ANIO.findall(str(valor))
        for texto in coincidencias:
            anio = int(texto)
            if 2000 <= anio <= 2100:
                anios.append(anio)

    if not anios:
        return ANIO_DEFECTO
    return pd.Series(anios).mode().iloc[0]


def hojas_diarias(nombres_hojas, anio: int = ANIO_DEFECTO) -> dict[str, pd.Timestamp]:
    """
    Detecta las hojas diarias del libro y devuelve ``{nombre_hoja: fecha}``.

    Se apoya en el nombre de la hoja ("17_Septimbre") aceptando el typo. Si el
    anio aparece en el nombre se usa ese; si no, ``anio``.

    No se usa el orden de las hojas ni su contenido: solo el nombre.
    """
    detectadas: dict[str, pd.Timestamp] = {}
    for nombre in nombres_hojas:
        m = PATRON_HOJA_DIARIA.match(str(nombre))
        if not m:
            continue
        dia = int(m.group(1))
        mes = MESES_ES.get(normalizar(m.group(2)).lower())
        if mes is None or not 1 <= dia <= 31:
            continue
        anio_hoja = anio
        try:
            detectadas[str(nombre)] = pd.Timestamp(anio_hoja, mes, dia)
        except ValueError:
            continue
    return detectadas


# ---------------------------------------------------------------------------
# Lectura
# ---------------------------------------------------------------------------

class ErrorPlan(Exception):
    """Error de lectura o de estructura del archivo del Plan de Trabajo."""


def _leer_hoja(archivo: str | bytes, hoja: str) -> pd.DataFrame:
    """Lee una hoja del libro y limpia los encabezados."""
    origen = archivo if isinstance(archivo, (str, os.PathLike)) else pd.io.common.BytesIO(archivo)
    df = pd.read_excel(origen, sheet_name=hoja, engine="openpyxl")
    df.columns = [
        " ".join(str(c).replace("\xa0", " ").split()) if c is not None else ""
        for c in df.columns
    ]
    return df


def _nombres_hojas(archivo: str | bytes) -> list[str]:
    """Lista los nombres de hoja sin cargar los datos."""
    origen = archivo if isinstance(archivo, (str, os.PathLike)) else pd.io.common.BytesIO(archivo)
    with pd.ExcelFile(origen, engine="openpyxl") as libro:
        return list(libro.sheet_names)


# ---------------------------------------------------------------------------
# Reconocimiento automatico del tipo de archivo
# ---------------------------------------------------------------------------

# Nombre de la hoja de la plantilla SLA (core.py). Se repite aqui para no
# depender de core al clasificar; si cambia alla, solo se degrada el mensaje.
PATRON_HOJA_PLANTILLA = re.compile(r"^PLANTILLA", re.IGNORECASE)

TIPO_PLANTILLA = "plantilla"
TIPO_PLAN = "plan"
TIPO_DESCONOCIDO = "desconocido"


def hoja_para_plantilla(archivo: str | bytes) -> str | None:
    """
    Devuelve la hoja que parece la plantilla SLA, o ``None`` si no hay.

    Acepta variantes tipo 'PLANTILLA ' o 'PLANTILLA SEPTIEMBRE' porque el
    encabezado real trae espacios sobrantes.
    """
    try:
        hojas = _nombres_hojas(archivo)
    except Exception:
        return None
    return next((h for h in hojas if PATRON_HOJA_PLANTILLA.match(str(h).strip())), None)


def reconocer_tipo(archivo: str | bytes) -> dict:
    """
    Reconoce que tipo de archivo es, mirando SOLO sus hojas.

    En el sistema conviven dos archivos distintos y el usuario no tiene por que
    recordar cual subir en cada casilla:

      * ``plantilla``   -> trae la hoja ``PLANTILLA`` (seguimiento SLA).
      * ``plan``        -> trae ``Casos_Ven`` y/o hojas diarias ``N_Mes``
                           (Plan de Trabajo).
      * ``desconocido`` -> ninguna de las anteriores.

    Detectar el tipo permite enrutar el archivo a la seccion correcta y dar un
    mensaje util ("esto es un Plan de Trabajo") en lugar de dejar que openpyxl
    falle con un ValueError sobre una hoja que no existe.

    Devuelve ``{"tipo", "hojas", "hoja_principal", "detalle"}``. No lanza
    excepcion por contenido: si el archivo no se puede abrir, devuelve
    ``desconocido`` con el motivo en ``detalle``.
    """
    try:
        hojas = _nombres_hojas(archivo)
    except Exception as exc:
        return {
            "tipo": TIPO_DESCONOCIDO,
            "hojas": [],
            "hoja_principal": None,
            "detalle": f"No se pudo abrir el archivo: {type(exc).__name__}: {exc}",
        }

    hoja_plantilla = next((h for h in hojas if PATRON_HOJA_PLANTILLA.match(str(h).strip())), None)
    if hoja_plantilla:
        return {
            "tipo": TIPO_PLANTILLA,
            "hojas": hojas,
            "hoja_principal": hoja_plantilla,
            "detalle": f"Contiene la hoja '{hoja_plantilla}'.",
        }

    hoja_vencidos = next(
        (h for h in hojas if normalizar(h) == normalizar(HOJA_VENCIDOS)), None
    )
    diarias = hojas_diarias(hojas)
    if hoja_vencidos or diarias:
        partes = []
        if hoja_vencidos:
            partes.append(f"la hoja '{hoja_vencidos}'")
        if diarias:
            partes.append(f"{len(diarias)} hojas diarias")
        return {
            "tipo": TIPO_PLAN,
            "hojas": hojas,
            "hoja_principal": hoja_vencidos,
            "detalle": "Contiene " + " y ".join(partes) + ".",
        }

    return {
        "tipo": TIPO_DESCONOCIDO,
        "hojas": hojas,
        "hoja_principal": None,
        "detalle": (
            "No trae la hoja 'PLANTILLA' ni 'Casos_Ven'. Hojas encontradas: "
            + ", ".join(map(str, hojas[:8]))
            + (" ..." if len(hojas) > 8 else "")
        ),
    }


def leer_vencidos(archivo: str | bytes, hoja: str = HOJA_VENCIDOS) -> pd.DataFrame:
    """
    Lee la hoja de casos vencidos y deja las columnas con nombres canonicos.

    Ademas resuelve dos problemas reales del archivo:

      * La ultima columna no tiene encabezado y contiene la categoria del caso
        (p. ej. ``CIBERSEGURIDAD``). Se renombra a ``CATEGORIA`` en lugar de
        perderla como ``Unnamed: 7``.
      * Hay filas desplazadas: cuando ``FECHA DE CREACION`` contiene una
        ubicacion, toda la fila esta corrida una columna. Se marcan para poder
        excluirlas del calculo en vez de producir un indicador falso.
    """
    df = _leer_hoja(archivo, hoja)

    renombres = {}
    for canonico, alias in (
        (COL_MES, ["MES ", "MES"]),
        (COL_CASO, ["CASO"]),
        (COL_UBICACION, ["UBICACION"]),
        (COL_FECHA_CREACION, ["FECHA DE CREACION"]),
        (COL_TECNICO, ["TECNICO ENCARGADO"]),
        (COL_JUSTIFICACION, ["JUSTIFICACION DE VENCIMIENTO "]),
        (COL_CULPA, ["CULPA "]),
    ):
        real = _buscar_columna(df, alias)
        if real is not None:
            renombres[real] = canonico

    # Columna sin encabezado -> CATEGORIA (se toma la primera sin nombre).
    sin_nombre = [c for c in df.columns if str(c).strip() == "" or str(c).startswith("Unnamed")]
    if sin_nombre:
        renombres.setdefault(sin_nombre[0], COL_CATEGORIA)

    df = df.rename(columns=renombres)

    faltantes = [c for c in (COL_CASO, COL_FECHA_CREACION, COL_TECNICO) if c not in df.columns]
    if faltantes:
        raise ErrorPlan(
            f"La hoja '{hoja}' no tiene las columnas esperadas. Faltan: {', '.join(faltantes)}. "
            f"Columnas encontradas: {', '.join(map(str, df.columns))}"
        )

    for col in (COL_MES, COL_UBICACION, COL_TECNICO, COL_JUSTIFICACION, COL_CULPA, COL_CATEGORIA):
        if col in df.columns:
            df[col] = df[col].astype("object")

    df[COL_CASO] = df[COL_CASO].astype(str).str.strip()

    # Fila desplazada: la "fecha" no es una fecha.
    fecha_texto = df[COL_FECHA_CREACION].astype(str)
    parece_fecha = fecha_texto.str.match(r"^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}", na=False)
    df["FILA_DESALINEADA"] = ~parece_fecha

    df["DUPLICADO"] = df.duplicated(subset=[COL_CASO], keep="first")

    for col in (COL_CULPA, COL_CATEGORIA, COL_MES, COL_UBICACION):
        if col in df.columns:
            df[col] = df[col].map(lambda v: normalizar(v) if pd.notna(v) else "")

    if COL_CULPA in df.columns:
        df[COL_CULPA] = df[COL_CULPA].replace(NORMALIZAR_CULPA_VEN)

    return df


def leer_cosecha_diaria(archivo: str | bytes, anio: int = ANIO_DEFECTO) -> pd.DataFrame:
    """
    Concatena todas las hojas diarias en un solo DataFrame.

    Cada fila es un caso visto en el tablero operativo de un dia. De aqui sale
    el ``Vencimiento`` (compromiso ANS) que la hoja de vencidos NO tiene.

    ``anio`` es el anio del plan: los nombres de hoja no lo traen.
    """
    nombres = _nombres_hojas(archivo)
    diarias = hojas_diarias(nombres, anio)
    if not diarias:
        raise ErrorPlan(
            "No se detecto ninguna hoja diaria (se esperaban nombres como "
            "'17_Septiembre'). Hojas encontradas: " + ", ".join(map(str, nombres))
        )

    partes = []
    for hoja, fecha in diarias.items():
        df = _leer_hoja(archivo, hoja)
        col_id = _buscar_columna(df, [COL_ID_DIARIO, "ID de incidente", "ID incidente"])
        if col_id is None:
            continue
        df = df.rename(columns={col_id: COL_ID_DIARIO})
        df[COL_ID_DIARIO] = df[COL_ID_DIARIO].astype(str).str.strip()
        df = df[df[COL_ID_DIARIO].ne("") & df[COL_ID_DIARIO].ne("NAN")]

        col_vto = _buscar_columna(df, [COL_VENCIMIENTO])
        col_ap = _buscar_columna(df, [COL_APERTURA, "Fecha/hora de apertura"])
        col_est = _buscar_columna(df, [COL_ESTADO_DIARIO])
        col_asig = _buscar_columna(df, [COL_ASIGNATARIO])
        col_ubi = _buscar_columna(df, [COL_UBICACION_DIARIA])
        col_ate = _buscar_columna(df, [COL_ATIENDE])

        partes.append(pd.DataFrame({
            COL_ID_DIARIO: df[COL_ID_DIARIO].values,
            "HOJA": hoja,
            "FECHA_HOJA": fecha,
            COL_APERTURA: df[col_ap].values if col_ap else pd.NaT,
            COL_VENCIMIENTO: df[col_vto].values if col_vto else pd.NaT,
            COL_ESTADO_DIARIO: df[col_est].values if col_est else "",
            COL_ASIGNATARIO: df[col_asig].values if col_asig else "",
            COL_UBICACION_DIARIA: df[col_ubi].values if col_ubi else "",
            COL_ATIENDE: df[col_ate].values if col_ate else "",
        }))

    if not partes:
        raise ErrorPlan("Las hojas diarias no tienen la columna 'ID de incidente'.")

    return pd.concat(partes, ignore_index=True)


# ---------------------------------------------------------------------------
# Lectura del "manual" del plan: las notas de las hojas diarias
# ---------------------------------------------------------------------------

# Columna I de las hojas diarias: lo que el equipo escribe a mano ("ALIADO EN
# SITIO", "MESA-COLSOF", "SE ENVIO A TIEMPOS"...). Es la fuente que la torre
# considera la verdad cuando el plan y la plantilla se contradicen.
COL_NOTA = "NOTA_MANANA"
COL_NOTA_NORMALIZADA = "NOTA_MANANA_NORMALIZADA"


def leer_manual_plan(archivo: str | bytes, anio: int | None = None) -> pd.DataFrame:
    """
    Devuelve, caso por caso y hoja por hoja, lo que el equipo anoto a mano.

    Columnas: ``COL_ID_DIARIO``, ``HOJA``, ``FECHA_HOJA``, ``COL_APERTURA``,
    ``COL_VENCIMIENTO``, ``COL_ASIGNATARIO``, ``COL_UBICACION_DIARIA``,
    ``COL_NOTA``, ``COL_NOTA_NORMALIZADA``, ``COL_ESTADO_DIARIO``, ``COL_ATIENDE``.

    Sirve para conciliar el plan con la plantilla SLA (ver ``conciliacion.py``):
    la nota manual es la que manda cuando los dos archivos dicen cosas distintas.

    ``anio`` se deduce del propio libro (de la hoja Casos_Ven) si no se indica,
    porque el nombre de las hojas diarias no trae el anio.
    """
    if anio is None:
        try:
            anio = inferir_anio(leer_vencidos(archivo, hoja=HOJA_VENCIDOS))
        except Exception:
            anio = ANIO_DEFECTO

    nombres = _nombres_hojas(archivo)
    diarias = hojas_diarias(nombres, anio)
    if not diarias:
        raise ErrorPlan(
            "No se detecto ninguna hoja diaria (se esperaban nombres como "
            "'17_Septiembre'). Hojas encontradas: " + ", ".join(map(str, nombres))
        )

    partes = []
    for hoja, fecha in diarias.items():
        df = _leer_hoja(archivo, hoja)
        col_id = _buscar_columna(df, [COL_ID_DIARIO, "ID de incidente", "ID incidente"])
        if col_id is None:
            continue
        df = df.rename(columns={col_id: COL_ID_DIARIO})
        df[COL_ID_DIARIO] = df[COL_ID_DIARIO].astype(str).str.strip()
        df = df[df[COL_ID_DIARIO].ne("") & df[COL_ID_DIARIO].ne("NAN")]
        if df.empty:
            continue

        def columna(nombres_col, defecto=""):
            real = _buscar_columna(df, nombres_col)
            if real is None:
                return pd.Series([defecto] * len(df), index=df.index)
            return df[real]

        notas = columna([COL_NOTA, "NOTAS MANANA"], "").reset_index(drop=True)
        apertura = pd.to_datetime(
            columna([COL_APERTURA], pd.NaT).reset_index(drop=True),
            errors="coerce", format="mixed",
        )
        vencimiento = pd.to_datetime(
            columna([COL_VENCIMIENTO], pd.NaT).reset_index(drop=True),
            errors="coerce", format="mixed", dayfirst=True,
        )

        partes.append(pd.DataFrame({
            COL_ID_DIARIO: df[COL_ID_DIARIO].values,
            "HOJA": hoja,
            "FECHA_HOJA": fecha,
            COL_APERTURA: apertura.values,
            COL_VENCIMIENTO: vencimiento.values,
            COL_ASIGNATARIO: columna([COL_ASIGNATARIO]).values,
            COL_UBICACION_DIARIA: columna([COL_UBICACION_DIARIA]).values,
            COL_NOTA: notas.values,
            COL_NOTA_NORMALIZADA: notas.map(normalizar).values,
            COL_ESTADO_DIARIO: columna([COL_ESTADO_DIARIO]).values,
            COL_ATIENDE: columna([COL_ATIENDE]).values,
        }))

    if not partes:
        raise ErrorPlan("Las hojas diarias no tienen la columna 'ID de incidente'.")

    return pd.concat(partes, ignore_index=True)


# ---------------------------------------------------------------------------
# Correccion de fechas invertidas
# ---------------------------------------------------------------------------

def _reinterpretar_dia_mes(serie: pd.Series) -> pd.Series:
    """
    Reinterpreta cada fecha intercambiando mes y dia (DD/MM en lugar de MM/DD).

    Los valores donde el mes original es > 12 no se pueden invertir y se dejan
    como estaban. Los valores que no son fecha quedan como NaT.
    """
    def una(valor):
        if pd.isna(valor):
            return pd.NaT
        try:
            return pd.Timestamp(valor.year, valor.day, valor.month, valor.hour, valor.minute)
        except (ValueError, TypeError):
            return pd.NaT

    return pd.to_datetime(serie, errors="coerce").map(una)


PATRON_PREFIJO_ID = re.compile(r"^\s*([A-Za-z]+)")


def prefijo_id(valor) -> str:
    """Prefijo alfabetico del ID de caso ('IM3237400' -> 'IM', 'QT3328831' -> 'QT')."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    m = PATRON_PREFIJO_ID.match(str(valor))
    return m.group(1).upper() if m else ""


def numero_id(valor) -> float:
    """Parte numerica del ID de caso, o NaN si no la tiene."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return float("nan")
    m = re.search(r"(\d+)", str(valor))
    return float(m.group(1)) if m else float("nan")


def modelo_ids_a_fecha(cosecha: pd.DataFrame) -> dict[str, tuple[float, float]]:
    """
    Ajusta una recta ``fecha = m * id + b`` por cada FAMILIA de ID.

    Los IDs vienen en familias con prefijo distinto y escalas que NO son
    comparables entre si: en septiembre 2026 los ``IM`` iban de 3.234.992 a
    3.239.053 y los ``QT`` de 3.324.408 a 3.330.307. Mezclarlas en una sola
    recta produce un ajuste absurdo (pendiente ~1209) e invalida el modelo, por
    eso se calibra una recta por familia.

    Dentro de cada familia el ID es correlativo y crece de forma casi constante,
    de modo que la recta da una fecha aproximada e INDEPENDIENTE DE LAS FECHAS:
    por eso sirve para decidir cual de las dos lecturas de una fecha ambigua es
    la correcta.

    Devuelve ``{prefijo: (m, b)}``.
    """
    if cosecha.empty:
        raise ErrorPlan("No hay datos de cosecha diaria para construir el modelo de IDs.")

    tabla = pd.DataFrame({
        "prefijo": cosecha[COL_ID_DIARIO].map(prefijo_id),
        "id": cosecha[COL_ID_DIARIO].map(numero_id),
        "fecha": cosecha["FECHA_HOJA"],
    }).dropna(subset=["id"])
    tabla = tabla[tabla["prefijo"].ne("")]

    if tabla.empty:
        raise ErrorPlan("Ningun ID de incidente tiene prefijo y numero; no se puede modelar.")

    modelos: dict[str, tuple[float, float]] = {}
    for prefijo, grupo in tabla.groupby("prefijo"):
        maximos = grupo.groupby("fecha")["id"].max().reset_index()
        if len(maximos) < 2:
            # Una sola hoja con esta familia: no hay recta posible.
            continue
        x = maximos["id"].to_numpy(dtype=float)
        y = maximos["fecha"].map(pd.Timestamp.timestamp).to_numpy(dtype=float)
        m, b = np.polyfit(x, y, 1)
        modelos[str(prefijo)] = (float(m), float(b))

    if not modelos:
        raise ErrorPlan(
            "Se necesita mas de una hoja diaria por familia de ID para calibrar "
            "el modelo de fechas."
        )
    return modelos


def _fecha_esperada(
    df: pd.DataFrame,
    modelos: dict[str, tuple[float, float]],
    columna_id: str,
) -> pd.Series:
    """
    Fecha aproximada de creacion de cada caso segun el modelo de su familia.

    Los casos de una familia sin modelo devuelven NaT: sin referencia no se
    corrige la fecha (no se adivina).
    """
    if columna_id not in df.columns:
        return pd.Series(pd.NaT, index=df.index)

    prefijos = df[columna_id].map(prefijo_id)
    numeros = df[columna_id].map(numero_id)

    resultado = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    for prefijo, (m, b) in modelos.items():
        mascara = prefijos.eq(prefijo) & numeros.notna()
        if not mascara.any():
            continue
        segundos = m * numeros[mascara] + b
        # Timedelta en lugar de to_datetime(unit="s"): asi los valores fuera de
        # rango se vuelven NaT en vez de desbordar.
        resultado.loc[mascara] = (
            pd.Timestamp("1970-01-01")
            + pd.to_timedelta(segundos, unit="s", errors="coerce")
        ).to_numpy()
    return resultado


def corregir_fechas(
    df: pd.DataFrame,
    columna: str,
    modelos: dict[str, tuple[float, float]],
    *,
    columna_id: str = COL_CASO,
    tolerancia_dias: float = TOLERANCIA_MODELO_DIAS,
) -> pd.Series:
    """
    Decide, caso por caso, si la fecha de ``columna`` debe invertirse.

    Regla: se invierte SOLO si (a) la lectura invertida se acerca mas al modelo
    de IDs que la lectura original y (b) queda dentro de la tolerancia. Si el
    caso pertenece a una familia sin modelo calibrado, se conserva la fecha
    original: no se adivina.

    Devuelve una Serie de fechas corregidas.
    """
    original = pd.to_datetime(df[columna], errors="coerce")
    invertida = _reinterpretar_dia_mes(original)
    esperada = _fecha_esperada(df, modelos, columna_id)

    def elegir(orig, inv, esp):
        if pd.isna(orig):
            return pd.NaT
        if pd.isna(inv) or pd.isna(esp):
            return orig
        error_orig = abs((orig - esp).total_seconds())
        error_inv = abs((inv - esp).total_seconds())
        if error_inv < error_orig and error_inv <= tolerancia_dias * 86400:
            return inv
        return orig

    return pd.Series(
        [elegir(o, i, e) for o, i, e in zip(original, invertida, esperada)],
        index=df.index,
    )




#
# Esta seccion reemplaza los indicadores de operacion (ANS por tramos,
# envejecimiento, culpa, velocidad de cierre). El objetivo es otro: un tablero
# que gerencia entienda en diez segundos y que sirva para enviar por correo.
#
# Fuente: las hojas diarias del Plan de Trabajo ("23_Septiembre", ...). Cada
# hoja es la foto de los casos EN CURSO ese dia. NO se usa la hoja Casos_Ven,
# que es el historico de vencidos (abiertos y cerrados) y responde otra
# pregunta.
#
# La fecha relevante es la de APERTURA (creacion del caso), que es la que
# alimenta el corte por mes y por semana. Esa columna arrastra el mismo
# problema de fechas mes/dia que el resto del archivo, asi que se corrige con
# el modelo de IDs antes de contar nada.

# Estados que se consideran "en curso" para el tablero gerencial.
# La hoja diaria trae ademas Suspendido, Ready y Work In Progress, que NO son
# trabajo activo; excluirlos es lo que hace cuadrar el total con el seguimiento
# que ya se envia por correo.
# Valor especial del parametro ``corte``: en lugar de la foto de un dia,
# devuelve el acumulado de todas las hojas diarias (sin repetir casos).





# ---------------------------------------------------------------------------
# Casos_Ven
# ---------------------------------------------------------------------------

# Nombres de mes para las etiquetas de la interfaz y los informes.
_ETIQUETA_MES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}

COL_MES_CREACION = "MES_CREACION"
COL_ANO_MES = "ANO_MES"
COL_DIAS_ABIERTO = "DIAS_ABIERTO"
COL_FECHA_CORREGIDA = "FECHA_CORREGIDA"

# Variantes de CULPA que son errores de digitacion y se agrupan.
NORMALIZAR_CULPA_VEN = {
    "BANCO?": "BANCO",
    "BANCOOOOOO": "BANCO",
    "BANCOOOO": "BANCO",
    "MESA": "MESA-COLSOF",
}


def _limpiar_texto(serie: pd.Series) -> pd.Series:
    """Trim y colapsa espacios; deja el texto presentable en pantalla."""
    return (
        serie.astype(str)
        .str.replace("\xa0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .replace({"nan": "", "None": "", "NaT": ""})
    )


def vista_vencidos(
    archivo: str | bytes,
    *,
    hoja: str = HOJA_VENCIDOS,
    corregir: bool = False,
) -> dict:
    """
    Arma el listado completo de la hoja Casos_Ven con sus totales.

    Parametros
    ----------
    corregir : DESACTIVADO por defecto a proposito. La hoja Casos_Ven guarda
        las fechas en formato mes/dia (con "9/01/2026" queriendo decir 1 de
        septiembre) y ya estan correctas, asi que invertirlas las estropearia.
        El modelo de secuencia de IDs que si se necesita es el de las hojas
        diarias, y se aplica en la pestana 'Plan de Trabajo'.

        Si algun mes se guardara en formato dia/mes, poner ``corregir=True``
        contrasta cada fecha contra ese modelo y solo invierte las que quedan
        mas cerca del valor esperado.

    Devuelve un diccionario con:

      ``casos``       una fila por caso unico, con fecha, mes, tecnico, culpa y
                      justificacion ya limpias.
      ``total``       numero de casos unicos.
      ``por_mes``     lista de ``{"mes", "ano_mes", "etiqueta", "casos"}``.
      ``por_tecnico`` lista de ``{"tecnico", "casos"}``.
      ``por_culpa``   lista de ``{"culpa", "casos"}``.
      ``por_categoria`` lista de ``{"categoria", "casos"}``.
      ``conciliacion``dict con el desglose de la lectura.
    """
    if isinstance(archivo, (bytes, bytearray)):
        datos = bytes(archivo)
    else:
        with open(archivo, "rb") as fh:
            datos = fh.read()

    info = reconocer_tipo(datos)
    hoja_real = next(
        (h for h in info["hojas"] if normalizar(h) == normalizar(hoja)), None
    )
    if hoja_real is None:
        raise ErrorPlan(
            "El archivo no tiene la hoja '" + hoja + "'. Hojas encontradas: "
            + ", ".join(map(str, info["hojas"][:8]))
        )

    crudo = leer_vencidos(datos, hoja=hoja_real)
    filas_leidas = len(crudo)

    # --- fecha de creacion -------------------------------------------------
    crudo[COL_FECHA_CREACION] = pd.to_datetime(
        crudo[COL_FECHA_CREACION], errors="coerce", format="mixed", dayfirst=False
    )
    if corregir:
        try:
            cosecha, _ = _cosecha_cacheada(datos)
            modelos = modelo_ids_a_fecha(cosecha)
            corregida = corregir_fechas(
                crudo, COL_FECHA_CREACION, modelos, columna_id=COL_CASO
            )
            crudo[COL_FECHA_CORREGIDA] = corregida != crudo[COL_FECHA_CREACION]
            crudo[COL_FECHA_CREACION] = corregida
        except Exception:
            # Sin cosecha no hay modelo: se dejan las fechas como estan.
            crudo[COL_FECHA_CORREGIDA] = False
    else:
        crudo[COL_FECHA_CORREGIDA] = False

    # --- deduplicar y quitar la fila desplazada ----------------------------
    casos = crudo[~crudo["FILA_DESALINEADA"]].copy()
    casos = casos.drop_duplicates(subset=[COL_CASO], keep="first")
    descartados = filas_leidas - len(casos)

    # --- columnas presentables --------------------------------------------
    for col in (COL_UBICACION, COL_TECNICO, COL_JUSTIFICACION):
        if col in casos.columns:
            casos[col] = _limpiar_texto(casos[col])
    for col in (COL_CULPA, COL_CATEGORIA):
        if col in casos.columns:
            casos[col] = _limpiar_texto(casos[col]).replace(NORMALIZAR_CULPA_VEN)

    casos[COL_MES_CREACION] = casos[COL_FECHA_CREACION].dt.strftime("%Y-%m")
    casos[COL_ANO_MES] = casos[COL_FECHA_CREACION].dt.to_period("M")
    casos[COL_DIAS_ABIERTO] = (
        pd.Timestamp.now().normalize() - casos[COL_FECHA_CREACION]
    ).dt.total_seconds() / 86400

    def etiqueta_mes(periodo) -> str:
        if pd.isna(periodo):
            return "Sin fecha"
        return _ETIQUETA_MES.get(periodo.month, str(periodo.month)) + " " + str(periodo.year)

    # --- totales ----------------------------------------------------------
    por_mes = []
    for periodo, grupo in casos.groupby(COL_ANO_MES, dropna=False):
        por_mes.append({
            "mes": str(periodo) if pd.notna(periodo) else "",
            "etiqueta": etiqueta_mes(periodo),
            "casos": int(len(grupo)),
        })
    por_mes.sort(key=lambda d: d["mes"] or "9999")

    def ranking(columna: str, etiqueta: str) -> list[dict]:
        if columna not in casos.columns:
            return []
        serie = casos[columna].replace({"": "(sin dato)"})
        # Los nombres vienen escritos de formas distintas ("caRLOS JIMENEZ" y
        # "CARLOS JIMENEZ"). Se unifican para no mostrar al mismo tecnico dos
        # veces, conservando la grafia mas frecuente.
        acumulado: dict = {}
        for valor in serie:
            clave = normalizar(valor)
            if not clave:
                continue
            entrada = acumulado.setdefault(clave, {"grafias": {}, "casos": 0})
            entrada["casos"] += 1
            original = str(valor).strip()
            entrada["grafias"][original] = entrada["grafias"].get(original, 0) + 1
        filas = []
        for clave, datos in acumulado.items():
            grafia = sorted(
                datos["grafias"].items(),
                key=lambda kv: (-kv[1], -sum(1 for c in kv[0] if c.isupper())),
            )[0][0]
            filas.append({etiqueta: grafia, "casos": int(datos["casos"])})
        return sorted(filas, key=lambda f: -f["casos"])

    conciliacion = {
        "filas_leidas": int(filas_leidas),
        "descartados": int(descartados),
        "sin_fecha": int(casos[COL_FECHA_CREACION].isna().sum()),
        "fechas_corregidas": int(casos[COL_FECHA_CORREGIDA].sum()),
        "sin_tecnico": int((casos[COL_TECNICO] == "").sum()) if COL_TECNICO in casos.columns else 0,
        "hoja": hoja_real,
    }

    return {
        "casos": casos,
        "total": int(len(casos)),
        "por_mes": por_mes,
        "por_tecnico": ranking(COL_TECNICO, "tecnico"),
        "por_culpa": ranking(COL_CULPA, "culpa"),
        "por_categoria": ranking(COL_CATEGORIA, "categoria"),
        "conciliacion": conciliacion,
        "meta": {
            "hoja": hoja_real,
            "corregido": bool(corregir),
            "meses": len(por_mes),
            "tecnicos": int(casos[COL_TECNICO].replace({"": pd.NA}).nunique())
            if COL_TECNICO in casos.columns else 0,
            "culpas": int(casos[COL_CULPA].replace({"": pd.NA}).nunique())
            if COL_CULPA in casos.columns else 0,
        },
    }


def validar_fechas_vencidos(
    archivo: str | bytes,
    *,
    hoja: str = HOJA_VENCIDOS,
) -> dict:
    """
    Verifica las fechas de Casos_Ven cruzándolas con otra fuente.

    La hoja diaria guarda, para cada caso, la fecha de apertura que registra el
    banco. Si esa fecha y la de creacion de Casos_Ven coinciden, las dos estan
    bien y no hay que corregir nada.

    Este contraste es el que zanja la duda de si las fechas vienen con mes y dia
    intercambiados: una sola columna es ambigua (9/01/2026 puede ser 9 de enero
    o 1 de septiembre), pero dos fuentes independientes que coinciden no.

    Devuelve ``{"comparables", "coinciden", "difieren", "pct", "ejemplos",
    "por_mes_diario", "por_mes_vencidos"}``.
    """
    if isinstance(archivo, (bytes, bytearray)):
        datos = bytes(archivo)
    else:
        with open(archivo, "rb") as fh:
            datos = fh.read()

    casos = leer_vencidos(datos, hoja=hoja)
    casos = casos[~casos["FILA_DESALINEADA"]].drop_duplicates(
        subset=[COL_CASO], keep="first"
    )
    casos["FECHA"] = pd.to_datetime(
        casos[COL_FECHA_CREACION], errors="coerce", format="mixed", dayfirst=False
    )

    cosecha = leer_cosecha_diaria(datos, inferir_anio(casos))
    aperturas = pd.to_datetime(cosecha[COL_APERTURA], errors="coerce")
    primera = (
        pd.DataFrame({COL_CASO: cosecha[COL_ID_DIARIO].astype(str).str.strip(),
                      "AP": aperturas})
        .dropna(subset=["AP"])
        .groupby(COL_CASO)["AP"].min()
    )

    casos["EN_DIARIO"] = casos[COL_CASO].map(primera)
    compara = casos.dropna(subset=["FECHA", "EN_DIARIO"]).copy()
    compara["DIF_DIAS"] = (
        compara["FECHA"] - compara["EN_DIARIO"]
    ).dt.total_seconds() / 86400
    compara["COINCIDE"] = compara["DIF_DIAS"].abs() < 1

    def reparto(columna) -> list:
        conteo = compara[columna].dt.to_period("M").value_counts().sort_index()
        return [
            {"mes": str(k), "etiqueta": _ETIQUETA_MES.get(k.month, str(k.month)),
             "casos": int(v)}
            for k, v in conteo.items()
        ]

    coinciden = int(compara["COINCIDE"].sum())
    total = int(len(compara))

    return {
        "comparables": total,
        "coinciden": coinciden,
        "difieren": total - coinciden,
        "pct": round(coinciden / total * 100, 1) if total else 0.0,
        "ejemplos": [
            {
                "caso": str(r[COL_CASO]),
                "vencidos": r["FECHA"],
                "diario": r["EN_DIARIO"],
            }
            for _, r in compara[~compara["COINCIDE"]].head(10).iterrows()
        ],
        "por_mes_diario": reparto("EN_DIARIO"),
        "por_mes_vencidos": reparto("FECHA"),
    }

