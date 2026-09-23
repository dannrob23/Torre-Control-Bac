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

import os
import re
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
COL_ESTADO = "ESTADO"
COL_FECHA_CIERRE = "FECHA DE CIERRE"

COL_ID_DIARIO = "ID de incidente"
COL_VENCIMIENTO = "Vencimiento"
COL_APERTURA = "Fecha/hora de apertura"
# La hoja diaria llama 'Estado' a su columna; la hoja de vencidos usa 'ESTADO'
# (la agrega crear_plantilla_plan.py). Son columnas distintas y por eso llevan
# constantes distintas.
COL_ESTADO_DIARIO = "Estado"
COL_ASIGNATARIO = "Asignatario"

# ---------------------------------------------------------------------------
# Semáforo de antigüedad (dias desde la creacion del caso)
# ---------------------------------------------------------------------------

EDAD_VERDE = 7
EDAD_AMARILLO = 15
EDAD_NARANJA = 30

TRAMO_VERDE = "VERDE (<= 7 d)"
TRAMO_AMARILLO = "AMARILLO (8-15 d)"
TRAMO_NARANJA = "NARANJA (16-30 d)"
TRAMO_ROJO = "ROJO (> 30 d)"

ORDEN_TRAMO = [TRAMO_VERDE, TRAMO_AMARILLO, TRAMO_NARANJA, TRAMO_ROJO]

# ---------------------------------------------------------------------------
# Semaforo de ANS (dias vencidos respecto al compromiso de la herramienta)
# ---------------------------------------------------------------------------

ANS_EN_PLAZO = "EN PLAZO"
ANS_1_7 = "1-7 d vencido"
ANS_8_30 = "8-30 d vencido"
ANS_31_90 = "31-90 d vencido"
ANS_MAS_90 = "> 90 d vencido"

ORDEN_ANS = [ANS_EN_PLAZO, ANS_1_7, ANS_8_30, ANS_31_90, ANS_MAS_90]

# Tolerancia con la que una fecha se considera "coherente" con el modelo de IDs.
# 7 dias absorbe el ruido del ajuste lineal (los IDs no crecen exactamente igual
# todos los dias) sin permitir que una fecha de otro mes pase por buena.
TOLERANCIA_MODELO_DIAS = 7.0

# Valores de CULPA que son errores de digitacion y se agrupan.
NORMALIZAR_CULPA = {
    "BANCO?": "BANCO",
    "BANCOOOOOO": "BANCO",
    "BANCOOOO": "BANCO",
}


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
        df[COL_CULPA] = df[COL_CULPA].replace(NORMALIZAR_CULPA)

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

        partes.append(pd.DataFrame({
            COL_ID_DIARIO: df[COL_ID_DIARIO].values,
            "HOJA": hoja,
            "FECHA_HOJA": fecha,
            COL_APERTURA: df[col_ap].values if col_ap else pd.NaT,
            COL_VENCIMIENTO: df[col_vto].values if col_vto else pd.NaT,
            COL_ESTADO_DIARIO: df[col_est].values if col_est else "",
            COL_ASIGNATARIO: df[col_asig].values if col_asig else "",
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


# ---------------------------------------------------------------------------
# Deduccion de tecnicos
# ---------------------------------------------------------------------------

def _tokens(nombre: str) -> list[str]:
    return [t for t in normalizar(nombre).split() if t]


def deducir_tecnicos(
    nombres: list[str],
    catalogo: dict[str, str] | None = None,
    alias: dict[str, str] | None = None,
) -> pd.DataFrame:
    """
    Empareja los nombres cortos del plan (``CARLOS BRAVO``) con los nombres
    completos del catalogo oficial (``CARLOS ANDRES BRAVO MARTINEZ``).

    El emparejamiento usa DOS tokens (nombre + apellido) y nunca solo el nombre
    de pila: en el catalogo hay 4 personas llamadas CARLOS, 3 CRISTIAN y 2 JORGE,
    asi que emparejar por nombre de pila atribuiria casos a la persona equivocada.

    Devuelve un DataFrame con una fila por nombre de entrada y las columnas:
    ``ORIGINAL``, ``NORMALIZADO``, ``CANONICO``, ``REGION``, ``METODO``,
    ``CONFIANZA`` y ``CANDIDATOS``. Las filas con ``CONFIANZA`` distinta de
    ``ALTA`` deben revisarse a mano: el modulo no adivina.
    """
    if catalogo is None:
        from core import TECNICOS_REGION
        catalogo = TECNICOS_REGION
    if alias is None:
        try:
            from core import ALIAS_TECNICOS
            alias = ALIAS_TECNICOS
        except ImportError:
            alias = {}

    alias_norm = {normalizar(k): v for k, v in (alias or {}).items()}
    catalogo_norm = {}
    for canonico in catalogo:
        catalogo_norm[normalizar(canonico)] = canonico

    filas = []
    for original in nombres:
        norm = normalizar(original)
        if not norm:
            filas.append({
                "ORIGINAL": original, "NORMALIZADO": norm, "CANONICO": "",
                "REGION": "SIN REGION", "METODO": "vacio",
                "CONFIANZA": "REVISAR", "CANDIDATOS": "",
            })
            continue

        # 1) Coincidencia exacta con el catalogo.
        if norm in catalogo_norm:
            canonico = catalogo_norm[norm]
            filas.append({
                "ORIGINAL": original, "NORMALIZADO": norm, "CANONICO": canonico,
                "REGION": catalogo.get(canonico, "SIN REGION"), "METODO": "exacto",
                "CONFIANZA": "ALTA", "CANDIDATOS": "",
            })
            continue

        # 2) Alias declarado.
        if norm in alias_norm:
            canonico = alias_norm[norm]
            filas.append({
                "ORIGINAL": original, "NORMALIZADO": norm, "CANONICO": canonico,
                "REGION": catalogo.get(canonico, "SIN REGION"), "METODO": "alias",
                "CONFIANZA": "ALTA", "CANDIDATOS": "",
            })
            continue

        # 3) Contencion de tokens: todos los tokens del nombre corto aparecen
        #    en el nombre completo.
        tokens = _tokens(norm)
        coincidencias = []
        for canonico_norm, canonico in catalogo_norm.items():
            tokens_can = set(_tokens(canonico_norm))
            if tokens and all(t in tokens_can for t in tokens):
                coincidencias.append(canonico)

        if len(coincidencias) == 1:
            canonico = coincidencias[0]
            metodo = "tokens" if len(tokens) > 1 else "nombre-pila-unico"
            filas.append({
                "ORIGINAL": original, "NORMALIZADO": norm, "CANONICO": canonico,
                "REGION": catalogo.get(canonico, "SIN REGION"), "METODO": metodo,
                "CONFIANZA": "ALTA", "CANDIDATOS": "",
            })
            continue

        # 4) Ambiguo o sin coincidencia: se reporta, no se adivina.
        filas.append({
            "ORIGINAL": original, "NORMALIZADO": norm,
            "CANONICO": "" if len(coincidencias) != 1 else coincidencias[0],
            "REGION": "SIN REGION", "METODO": "ambiguo" if coincidencias else "sin-coincidencia",
            "CONFIANZA": "REVISAR",
            "CANDIDATOS": " | ".join(sorted(coincidencias)),
        })

    return pd.DataFrame(filas)


# ---------------------------------------------------------------------------
# Indicadores
# ---------------------------------------------------------------------------

def _tramo_edad(dias) -> str:
    if pd.isna(dias):
        return "SIN FECHA"
    if dias <= EDAD_VERDE:
        return TRAMO_VERDE
    if dias <= EDAD_AMARILLO:
        return TRAMO_AMARILLO
    if dias <= EDAD_NARANJA:
        return TRAMO_NARANJA
    return TRAMO_ROJO


def _tramo_ans(dias) -> str:
    if pd.isna(dias):
        return "SIN VENCIMIENTO"
    if dias < 0:
        return ANS_EN_PLAZO
    if dias <= 7:
        return ANS_1_7
    if dias <= 30:
        return ANS_8_30
    if dias <= 90:
        return ANS_31_90
    return ANS_MAS_90


def analizar(
    archivo: str | bytes,
    *,
    momento: pd.Timestamp | None = None,
    hoja_vencidos: str = HOJA_VENCIDOS,
) -> dict:
    """
    Ejecuta el analisis completo del Plan de Trabajo.

    Devuelve un diccionario con:

      ``casos``      DataFrame por caso, con fechas corregidas, edad, ANS,
                     tramos, region y estado de deduccion del tecnico.
      ``tecnicos``   DataFrame de indicadores por tecnico.
      ``calidad``    dict con los hallazgos de calidad de datos.
      ``meta``       dict con el modelo de fechas y las cifras de correccion.
      ``cosecha``    DataFrame de la cosecha diaria (para diagnostico).
    """
    if momento is None:
        try:
            from core import ahora_colombia
            momento = pd.Timestamp(ahora_colombia())
        except ImportError:
            momento = pd.Timestamp.now()

    # El tipo de archivo se reconoce solo: si lo que llega no es un Plan de
    # Trabajo, se dice con claridad en vez de fallar buscando una hoja.
    info = reconocer_tipo(archivo)
    if info["tipo"] == TIPO_PLANTILLA:
        raise ErrorPlan(
            "Este archivo es la PLANTILLA de seguimiento SLA, no un Plan de "
            f"Trabajo ({info['detalle']}). Subalo en 'Cargar plantilla', no en "
            "'Cargar Plan de Trabajo'."
        )
    if info["tipo"] == TIPO_DESCONOCIDO:
        raise ErrorPlan("El archivo no parece un Plan de Trabajo. " + info["detalle"])

    hoja = hoja_vencidos or info["hoja_principal"] or HOJA_VENCIDOS
    crudo = leer_vencidos(archivo, hoja=hoja)
    # El anio del plan se deduce del propio archivo: los nombres de hoja no lo
    # traen ("17_Septiembre") y no se quiere una constante que caduque.
    anio = inferir_anio(crudo)
    cosecha = leer_cosecha_diaria(archivo, anio)
    modelos = modelo_ids_a_fecha(cosecha)

    # --- fechas de creacion corregidas -----------------------------------
    crudo["FECHA_CREACION_ORIGINAL"] = pd.to_datetime(
        crudo[COL_FECHA_CREACION], errors="coerce", format="mixed", dayfirst=False
    )
    crudo["FECHA_CREACION"] = corregir_fechas(
        crudo, COL_FECHA_CREACION, modelos, columna_id=COL_CASO
    )
    crudo["FECHA_CORREGIDA"] = (
        crudo["FECHA_CREACION"].notna()
        & crudo["FECHA_CREACION_ORIGINAL"].notna()
        & (crudo["FECHA_CREACION"] != crudo["FECHA_CREACION_ORIGINAL"])
    )

    # --- vencimiento (ANS) desde la cosecha -------------------------------
    diario = cosecha.copy()
    diario["VENCIMIENTO_NATIVO"] = pd.to_datetime(
        diario[COL_VENCIMIENTO].astype(str).str.strip(),
        errors="coerce", format="mixed", dayfirst=False,
    )
    diario["VENCIMIENTO_CORREGIDO"] = corregir_fechas(
        diario.assign(**{COL_CASO: diario[COL_ID_DIARIO]}),
        "VENCIMIENTO_NATIVO", modelos, columna_id=COL_CASO,
    )
    vencimiento_por_caso = (
        diario.dropna(subset=["VENCIMIENTO_CORREGIDO"])
        .groupby(COL_ID_DIARIO)["VENCIMIENTO_CORREGIDO"].min()
    )
    crudo["VENCIMIENTO"] = crudo[COL_CASO].map(vencimiento_por_caso)

    estado_por_caso = (
        diario.sort_values("FECHA_HOJA").groupby(COL_ID_DIARIO)[COL_ESTADO_DIARIO].last()
    )
    crudo["ESTADO_ULTIMO"] = crudo[COL_CASO].map(estado_por_caso).map(
        lambda v: normalizar(v) if pd.notna(v) else ""
    )

    # --- metricas por caso ------------------------------------------------
    casos = crudo[~crudo["FILA_DESALINEADA"]].copy()
    casos["DIAS_ABIERTO"] = (momento - casos["FECHA_CREACION"]).dt.total_seconds() / 86400
    casos["DIAS_VENCIDO"] = (momento - casos["VENCIMIENTO"]).dt.total_seconds() / 86400
    casos["TRAMO_EDAD"] = casos["DIAS_ABIERTO"].map(_tramo_edad)
    casos["TRAMO_ANS"] = casos["DIAS_VENCIDO"].map(_tramo_ans)

    # Velocidad de cierre: solo si el archivo trae ESTADO y FECHA DE CIERRE.
    casos = calcular_cierre(casos)

    # --- tecnicos ---------------------------------------------------------
    nombres = sorted({n for n in casos[COL_TECNICO].dropna().unique() if str(n).strip()})
    mapa = deducir_tecnicos(nombres)

    # Varias grafias del mismo nombre ("caRLOS JIMENEZ" y "CARLOS JIMENEZ")
    # colapsan al mismo texto normalizado. Se consolida para poder indexar, y se
    # conserva la confianza mas baja: si una variante es dudosa, el tecnico
    # completo queda marcado para revision.
    consenso = (
        mapa.groupby("NORMALIZADO")
        .agg({
            "CANONICO": "first",
            "REGION": "first",
            "CONFIANZA": lambda s: "REVISAR" if "REVISAR" in set(s) else "ALTA",
            "METODO": lambda s: " + ".join(sorted(set(s))),
            "ORIGINAL": lambda s: " | ".join(sorted(set(s))),
        })
    )

    casos["TECNICO_NORM"] = casos[COL_TECNICO].map(normalizar)
    casos["TECNICO_CANONICO"] = casos["TECNICO_NORM"].map(consenso["CANONICO"]).fillna("")
    casos["REGION"] = casos["TECNICO_NORM"].map(consenso["REGION"]).fillna("SIN REGION")
    casos["TECNICO_CONFIANZA"] = casos["TECNICO_NORM"].map(consenso["CONFIANZA"]).fillna("REVISAR")

    tecnicos = _indicadores_por_tecnico(casos)
    cierre = _indicadores_cierre(casos)
    calidad = _calidad(crudo, casos, cosecha, mapa)

    return {
        "casos": casos,
        "tecnicos": tecnicos,
        "cierre": cierre,
        "mapa_tecnicos": mapa,
        "calidad": calidad,
        "cosecha": cosecha,
        "meta": {
            "momento": momento,
            "anio": anio,
            "modelos": modelos,
            "casos_leidos": int(len(crudo)),
            "casos_analizados": int(len(casos)),
            "fechas_corregidas": int(casos["FECHA_CORREGIDA"].sum()),
            "filas_desalineadas": int(crudo["FILA_DESALINEADA"].sum()),
            "duplicados": int(crudo["DUPLICADO"].sum()),
            "con_vencimiento": int(casos["VENCIMIENTO"].notna().sum()),
        },
    }


def _indicadores_por_tecnico(casos: pd.DataFrame) -> pd.DataFrame:
    """Resumen de cartera y envejecimiento por tecnico."""
    if casos.empty:
        return pd.DataFrame()

    filas = []
    for (norm, canonico, region), grupo in casos.groupby(
        ["TECNICO_NORM", "TECNICO_CANONICO", "REGION"], dropna=False
    ):
        edad = grupo["DIAS_ABIERTO"].dropna()
        ans = grupo["DIAS_VENCIDO"].dropna()
        confianzas = set(grupo["TECNICO_CONFIANZA"].dropna())
        filas.append({
            "TECNICO": canonico or norm,
            "NOMBRE_EN_PLAN": norm,
            "REGION": region,
            "CASOS": int(len(grupo)),
            "EDAD_PROM": round(float(edad.mean()), 1) if len(edad) else np.nan,
            "EDAD_MEDIANA": round(float(edad.median()), 1) if len(edad) else np.nan,
            "EDAD_MAX": round(float(edad.max()), 1) if len(edad) else np.nan,
            "MAS_30D": int((edad > EDAD_NARANJA).sum()),
            "MAS_15D": int((edad > EDAD_AMARILLO).sum()),
            "ANS_PROM": round(float(ans.mean()), 1) if len(ans) else np.nan,
            "ANS_MAX": round(float(ans.max()), 1) if len(ans) else np.nan,
            "ANS_SOBRE_90": int((ans > 90).sum()),
            "PCT_EVITABLE": round(
                float(grupo[COL_CULPA].isin(["TECNICO", "LOGISTICO"]).mean() * 100), 1
            ) if COL_CULPA in grupo.columns else np.nan,
            "CONFIANZA": "REVISAR" if "REVISAR" in confianzas else "ALTA",
        })

    tabla = pd.DataFrame(filas)
    if tabla.empty:
        return tabla
    return tabla.sort_values(["ANS_SOBRE_90", "MAS_30D", "CASOS"], ascending=False).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Velocidad de cierre (requiere ESTADO y FECHA DE CIERRE)
# ---------------------------------------------------------------------------

# Estados que cuentan como cierre efectivo.
ESTADOS_CERRADOS = {"CERRADO A TIEMPOS", "CERRADO TARDE", "CANCELADO"}

# Estados que indican que el caso sigue vivo.
ESTADOS_ABIERTOS = {
    "ABIERTO", "EN CURSO", "EN RUTA", "EN FIRMAS", "EN VALIDACION",
    "SUSPENDIDO", "ASIGNADO", "CATEGORIZADO", "PENDIENTE", "PENDING",
    "PREPARADO", "READY", "TRABAJO EN CURSO", "WORK IN PROGRESS",
}


def calcular_cierre(casos: pd.DataFrame) -> pd.DataFrame:
    """
    Añade las columnas de velocidad de cierre, si el archivo las trae.

    Requiere que ``Casos_Ven`` tenga ``ESTADO`` y ``FECHA DE CIERRE`` (las
    agrega ``crear_plantilla_plan.py``). Si no existen o estan vacias, devuelve
    el DataFrame sin tocar: el tablero lo informa y no inventa resultados.

    Columnas que añade:
      ``CERRADO``        bool, el caso esta cerrado
      ``DIAS_CIERRE``    dias entre creacion y cierre
      ``CUMPLIO_ANS``    bool, se cerro antes del vencimiento
      ``DESVIACION``     dias de desviacion (positivo = cerro tarde)
    """
    if COL_ESTADO not in casos.columns:
        return casos

    casos = casos.copy()
    estado = casos[COL_ESTADO].map(lambda v: normalizar(v) if pd.notna(v) else "")

    tiene_cierre = COL_FECHA_CIERRE in casos.columns
    fecha_cierre = (
        pd.to_datetime(casos[COL_FECHA_CIERRE], errors="coerce", format="mixed", dayfirst=True)
        if tiene_cierre else pd.Series(pd.NaT, index=casos.index)
    )

    casos["CERRADO"] = estado.isin(ESTADOS_CERRADOS) | fecha_cierre.notna()
    casos["DIAS_CIERRE"] = (fecha_cierre - casos["FECHA_CREACION"]).dt.total_seconds() / 86400
    casos["CUMPLIO_ANS"] = np.where(
        fecha_cierre.notna() & casos["VENCIMIENTO"].notna(),
        fecha_cierre <= casos["VENCIMIENTO"],
        None,
    )
    casos["DESVIACION"] = (fecha_cierre - casos["VENCIMIENTO"]).dt.total_seconds() / 86400
    return casos


def _indicadores_cierre(casos: pd.DataFrame) -> pd.DataFrame:
    """Tiempo de cierre y cumplimiento del ANS por tecnico."""
    if "DIAS_CIERRE" not in casos.columns or casos["DIAS_CIERRE"].notna().sum() == 0:
        return pd.DataFrame()

    filas = []
    for (norm, canonico, region), grupo in casos.groupby(
        ["TECNICO_NORM", "TECNICO_CANONICO", "REGION"], dropna=False
    ):
        cerrados = grupo[grupo["CERRADO"]]
        dias = grupo["DIAS_CIERRE"].dropna()
        cumplio = grupo["CUMPLIO_ANS"].dropna()
        desv = grupo["DESVIACION"].dropna()
        filas.append({
            "TECNICO": canonico or norm,
            "REGION": region,
            "CASOS": int(len(grupo)),
            "CERRADOS": int(len(cerrados)),
            "ABIERTOS": int(len(grupo) - len(cerrados)),
            "DIAS_CIERRE_PROM": round(float(dias.mean()), 1) if len(dias) else np.nan,
            "DIAS_CIERRE_MEDIANA": round(float(dias.median()), 1) if len(dias) else np.nan,
            "PCT_CUMPLIO": round(float(cumplio.astype(bool).mean() * 100), 1) if len(cumplio) else np.nan,
            "DESVIACION_PROM": round(float(desv.mean()), 1) if len(desv) else np.nan,
        })

    tabla = pd.DataFrame(filas)
    if tabla.empty:
        return tabla
    return tabla.sort_values("DIAS_CIERRE_MEDIANA", ascending=False).reset_index(drop=True)


def _calidad(crudo: pd.DataFrame, casos: pd.DataFrame, cosecha: pd.DataFrame, mapa: pd.DataFrame) -> dict:
    """Hallazgos de calidad de datos que el operador debe conocer."""
    culpas_vacias = pd.Series(dtype=str)
    culpas_crudas = casos[COL_CULPA] if COL_CULPA in casos.columns else culpas_vacias
    culpas_ok = {
        "TECNICO", "LOGISTICO", "ALIADO", "BANCO", "ACTIVOS",
        "MESA-COLSOF", "MESA-TECNICO", "LOGISTICO/TECNICO", "MESA",
    }
    estados_vacios = pd.Series(dtype=str)
    estados_crudos = cosecha[COL_ESTADO_DIARIO] if COL_ESTADO_DIARIO in cosecha.columns else estados_vacios

    return {
        "filas_leidas": int(len(crudo)),
        "filas_desalineadas": int(crudo["FILA_DESALINEADA"].sum()),
        "duplicados": int(crudo["DUPLICADO"].sum()),
        "fechas_corregidas": int(casos["FECHA_CORREGIDA"].sum()),
        "sin_fecha": int(casos["FECHA_CREACION"].isna().sum()),
        "sin_vencimiento": int(casos["VENCIMIENTO"].isna().sum()),
        "sin_tecnico": int(casos[COL_TECNICO].isna().sum() + casos[COL_TECNICO].eq("").sum()),
        "tecnicos_a_revisar": mapa[mapa["CONFIANZA"] != "ALTA"].to_dict("records"),
        "culpas_no_reconocidas": sorted(
            set(culpas_crudas.dropna().unique()) - culpas_ok - {""}
        ),
        "estados_diarios": sorted(
            set(estados_crudos.dropna().map(normalizar).unique()) - {""}
        ),
        "cierre_disponible": bool(
            "DIAS_CIERRE" in casos.columns and casos["DIAS_CIERRE"].notna().any()
        ),
    }


# ---------------------------------------------------------------------------
# Informe de texto
# ---------------------------------------------------------------------------

def informe_consola(resultado: dict) -> str:
    """Resumen legible del analisis, para consola y para pruebas."""
    meta = resultado["meta"]
    casos = resultado["casos"]
    tecnicos = resultado["tecnicos"]
    lineas = []
    ap = lineas.append

    ap("=" * 78)
    ap("PLAN DE TRABAJO - analisis de cartera vencida")
    ap("=" * 78)
    ap(f"Momento de calculo : {meta['momento']:%Y-%m-%d %H:%M}")
    ap(f"Casos leidos       : {meta['casos_leidos']}")
    ap(f"Casos analizados   : {meta['casos_analizados']}")
    ap(f"Fechas corregidas  : {meta['fechas_corregidas']} (invertidas por Excel)")
    ap(f"Filas desalineadas : {meta['filas_desalineadas']}")
    ap(f"Duplicados         : {meta['duplicados']}")
    ap(f"Con vencimiento    : {meta['con_vencimiento']}")
    ap("")

    ap("--- ENVEJECIMIENTO (dias desde la creacion) ---")
    edad = casos["DIAS_ABIERTO"].dropna()
    if len(edad):
        ap(f"  promedio {edad.mean():.0f} d | mediana {edad.median():.0f} d | max {edad.max():.0f} d")
    for tramo in ORDEN_TRAMO:
        n = int((casos["TRAMO_EDAD"] == tramo).sum())
        ap(f"  {tramo:22s} {n:4d}")
    ap("")

    ap("--- ANS (dias vencidos) ---")
    ans = casos["DIAS_VENCIDO"].dropna()
    if len(ans):
        ap(f"  promedio {ans.mean():.0f} d | mediana {ans.median():.0f} d | max {ans.max():.0f} d")
    for tramo in ORDEN_ANS:
        n = int((casos["TRAMO_ANS"] == tramo).sum())
        ap(f"  {tramo:22s} {n:4d}")
    ap("")

    ap("--- POR TECNICO ---")
    if not tecnicos.empty:
        ap(f"  {'TECNICO':30s} {'CASOS':>6s} {'EDAD':>6s} {'ANS':>6s} {'>30d':>5s} {'>90d':>5s}")
        for _, r in tecnicos.iterrows():
            ap(f"  {str(r['TECNICO'])[:30]:30s} {r['CASOS']:6d} "
               f"{r['EDAD_PROM']:6.0f} {r['ANS_PROM']:6.0f} {r['MAS_30D']:5d} {r['ANS_SOBRE_90']:5d}")
    ap("")

    cal = resultado["calidad"]
    if cal.get("cierre_disponible"):
        cierre = resultado.get("cierre")
        ap("")
        ap("--- VELOCIDAD DE CIERRE ---")
        if cierre is not None and not cierre.empty:
            ap(f"  {'TECNICO':30s} {'CASOS':>6s} {'CERR':>5s} {'ABIERT':>7s} {'D.CIERRE':>9s} {'%CUMPL':>7s}")
            for _, r in cierre.iterrows():
                ap(f"  {str(r['TECNICO'])[:30]:30s} {r['CASOS']:6d} {r['CERRADOS']:5d} "
                   f"{r['ABIERTOS']:7d} {r['DIAS_CIERRE_MEDIANA']:9.0f} {r['PCT_CUMPLIO']:7.0f}")
    else:
        ap("")
        ap("--- VELOCIDAD DE CIERRE ---")
        ap("  No medible: falta ESTADO y FECHA DE CIERRE en la hoja Casos_Ven.")
        ap("  Ejecute crear_plantilla_plan.py para agregarlas.")

    if cal["tecnicos_a_revisar"]:
        ap("--- TECNICOS A REVISAR ---")
        for fila in cal["tecnicos_a_revisar"]:
            ap(f"  {fila['ORIGINAL']!r} -> {fila['METODO']} {fila['CANDIDATOS']}")
    if cal["culpas_no_reconocidas"]:
        ap("--- CULPAS NO RECONOCIDAS ---")
        ap("  " + ", ".join(map(str, cal["culpas_no_reconocidas"])))

    return "\n".join(lineas)
