"""
conciliacion.py - Cruza el Plan de Trabajo con la plantilla SLA.

Por que existe
--------------
El Plan de Trabajo (lo que diligencia la torre) y la plantilla de seguimiento del
banco se llenan por separado, y a veces dicen cosas distintas del MISMO caso:

  * El plan ya lo manda a cierre y la plantilla sigue sin fecha de resolucion,
    asi que el tablero lo sigue alertando como ABIERTO y VENCIDO.
  * El caso esta repetido en la plantilla (una fila sin cerrar y otra cerrada).
  * El caso esta en un archivo y no en el otro.
  * Las fechas no coinciden, sobre todo cuando una quedo invertida mes/dia.

Este modulo hace el cruce y NO decide nada: senala el desajuste, muestra que dice
cada archivo y deja la nota manual del plan (columna I "NOTAS MANANA" de las hojas
diarias, mas la justificacion de Casos_Ven) como contexto principal.

Uso:
    import conciliacion
    cruce = conciliacion.conciliar(plan_bytes, df_completo)
    cruce["conteos"], cruce["tabla"], cruce["resumen"]
"""
from __future__ import annotations

import pandas as pd

import core
import plan

# ---------------------------------------------------------------------------
# Tipos de desajuste
# ---------------------------------------------------------------------------

TIPO_DUPLICADO = "Caso repetido en la plantilla"
TIPO_CIERRE = "El plan reporta cierre y la plantilla sigue abierta"
TIPO_SOLO_PLAN = "Solo está en el plan"
TIPO_SOLO_PLANTILLA = "Solo está en la plantilla"
TIPO_FECHA_INVERTIDA = "La fecha de apertura viene invertida (mes/día)"
TIPO_FECHA_DISTINTA = "La fecha de creación no coincide"
TIPO_VENCIMIENTO = "El vencimiento no coincide"
TIPO_VENCIMIENTO_INVERTIDO = "El vencimiento viene invertido (mes/día)"

# Orden de presentacion: primero lo que exige revisar el dato.
TIPOS = (
    TIPO_CIERRE,
    TIPO_DUPLICADO,
    TIPO_SOLO_PLAN,
    TIPO_SOLO_PLANTILLA,
    TIPO_FECHA_INVERTIDA,
    TIPO_VENCIMIENTO_INVERTIDO,
    TIPO_VENCIMIENTO,
    TIPO_FECHA_DISTINTA,
)

ICONO_TIPO = {
    TIPO_CIERRE: "📝",
    TIPO_DUPLICADO: "🔁",
    TIPO_SOLO_PLAN: "🅿️",
    TIPO_SOLO_PLANTILLA: "📄",
    TIPO_FECHA_INVERTIDA: "🔀",
    TIPO_VENCIMIENTO: "⏰",
    TIPO_VENCIMIENTO_INVERTIDO: "🔀",
    TIPO_FECHA_DISTINTA: "📅",
}

DESCRIPCION_TIPO = {
    TIPO_CIERRE: (
        "El plan anota que el caso se envió a cierre (o quedó cerrado) y en la "
        "plantilla sigue SIN fecha de resolución: el tablero lo cuenta como abierto."
    ),
    TIPO_DUPLICADO: (
        "El N° DE CASO aparece en varias filas de la plantilla. Cada copia se "
        "clasifica por su cuenta, así que el caso puede salir cerrado y a la vez "
        "alertar como abierto."
    ),
    TIPO_SOLO_PLAN: (
        "El caso está en el plan y no aparece en la plantilla SLA: no tiene "
        "semáforo ni se notifica."
    ),
    TIPO_SOLO_PLANTILLA: (
        "El caso está en la plantilla y no está en el plan: no tiene la nota "
        "manual de la torre."
    ),
    TIPO_FECHA_INVERTIDA: (
        "El plan y la plantilla traen la MISMA fecha de apertura con el mes y el "
        "día intercambiados: alguno de los dos archivos la guardó en formato "
        "mes/día (por ejemplo 09/10 queriendo decir 10/09)."
    ),
    TIPO_VENCIMIENTO_INVERTIDO: (
        "El vencimiento del plan y el de la plantilla son el mismo día con el mes "
        "y el día intercambiados: hay que revisar el formato de fechas del plan."
    ),
    TIPO_VENCIMIENTO: "El vencimiento del plan y el de la plantilla no coinciden.",
    TIPO_FECHA_DISTINTA: (
        "La fecha de creación del plan y la de la plantilla no coinciden."
    ),
}

# Palabras que indican que el plan reporta un cierre o una resolucion.
# OJO: no se afirma que el caso este cerrado; solo que el plan lo dice. Por eso
# "MESA COLSO NO LA CERRO" tambien entra: es justo el choque que hay que revisar.
PALABRAS_CIERRE = (
    "CIERRE", "CERRAD", "CERRARON", "CERRO", "CERRAR", "RESUELT", "RESOLV",
    "FINALIZ", "SOLUCIONAD", "ENTREGADO",
)

# Diferencia tolerada entre la fecha del plan y la de la plantilla. La plantilla
# guarda segundos (09:42:49) y el plan redondea al minuto (09:42:00): sin esta
# tolerancia, casi todo caso saldria como "fecha distinta".
TOLERANCIA_MINUTOS = 10

# Columnas visibles de la tabla de conciliacion (en este orden).
COLUMNAS_TABLA = [
    "Caso", "Desajuste", "Plan: Hoja", "Plan: Vence", "Plan: Asignatario",
    "Plan: Ubicación", "Plan: Nota (col I)", "Plan: Estado", "Plan: Justificación",
    "Plan: Culpa", "Plantilla: Filas", "Plantilla: Estado", "Plantilla: Resolución",
    "Plantilla: Vence", "Plantilla: Técnico",
]

# ---------------------------------------------------------------------------
# Nombres internos: cada lado con su prefijo, para que el cruce no choque.
# ---------------------------------------------------------------------------

P_APERTURA = "PLAN_APERTURA"
P_VENCE = "PLAN_VENCE"
P_ASIGNATARIO = "PLAN_ASIGNATARIO"
P_UBICACION = "PLAN_UBICACION"
P_NOTA = "PLAN_NOTA"
P_NOTA_NORM = "PLAN_NOTA_NORM"
P_NOTAS = "PLAN_NOTAS_TODAS"
P_ESTADO = "PLAN_ESTADO"
P_HOJA = "PLAN_HOJA"
P_FECHA_HOJA = "PLAN_FECHA_HOJA"
P_JUSTIF = "PLAN_JUSTIFICACION"
P_CULPA = "PLAN_CULPA"
P_CATEGORIA = "PLAN_CATEGORIA"
P_FECHA_VEN = "PLAN_FECHA_VENCIDOS"

L_FILAS = "PLANTILLA_FILAS"
L_FILAS_EXCEL = "PLANTILLA_FILAS_EXCEL"
L_DUPLICADO = "PLANTILLA_DUPLICADO"
L_ESTADO = "PLANTILLA_ESTADO"
L_ESTADO_CALC = "PLANTILLA_ESTADO_CALCULADO"
L_RESOLUCION = "PLANTILLA_RESOLUCION"
L_CREACION = "PLANTILLA_CREACION"
L_VENCE = "PLANTILLA_VENCE"
L_TECNICO = "PLANTILLA_TECNICO"
L_CERRADO = "PLANTILLA_CERRADO"


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _texto(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return " ".join(str(valor).split())


def _bool(valor) -> bool:
    """
    Booleano seguro para columnas que pueden traer NaN.

    OJO: ``bool(nan)`` es True en Python. Sin esto, una fila que solo está en el
    plan (sin dato en la plantilla) salía marcada como "caso repetido".
    """
    if valor is None:
        return False
    try:
        if pd.isna(valor):
            return False
    except (TypeError, ValueError):
        return False
    return bool(valor)


def _fecha(valor, formato: str = "%d/%m/%Y %H:%M") -> str:
    if valor is None or pd.isna(valor):
        return "—"
    try:
        return pd.Timestamp(valor).strftime(formato)
    except (ValueError, TypeError):
        return "—"


def _invertir(fecha):
    """Intercambia mes y dia (la inversion tipica del Excel: 10/09 -> 09/10)."""
    if fecha is None or pd.isna(fecha):
        return pd.NaT
    marca = pd.Timestamp(fecha)
    try:
        return pd.Timestamp(marca.year, marca.day, marca.month, marca.hour,
                            marca.minute, marca.second)
    except (ValueError, TypeError):
        return pd.NaT


def _iguales(uno, otro, minutos: float = TOLERANCIA_MINUTOS) -> bool:
    if uno is None or otro is None or pd.isna(uno) or pd.isna(otro):
        return False
    diferencia = abs((pd.Timestamp(uno) - pd.Timestamp(otro)).total_seconds())
    return diferencia <= minutos * 60


def menciona_cierre(*textos) -> bool:
    """True si alguno de los textos del plan habla de cierre o resolucion."""
    for texto in textos:
        normalizado = plan.normalizar(texto)
        if not normalizado:
            continue
        if any(palabra in normalizado for palabra in PALABRAS_CIERRE):
            return True
    return False


# ---------------------------------------------------------------------------
# Lectura de cada lado
# ---------------------------------------------------------------------------

def _vacio(columnas: list[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columnas)


def _plan_manual(plan_bytes) -> tuple[pd.DataFrame, dict | None]:
    """
    Notas de las hojas diarias: un renglon por caso, con la nota mas reciente.

    Devuelve ``(marco, aviso)``. El aviso explica cuando NO se pudieron leer las
    notas manuales, para no entregar un cruce incompleto en silencio.
    """
    columnas = ["CLAVE", P_HOJA, P_FECHA_HOJA, P_APERTURA, P_VENCE, P_ASIGNATARIO,
                P_UBICACION, P_NOTA, P_NOTA_NORM, P_NOTAS, P_ESTADO]

    # El servidor puede quedar con una version vieja de plan.py en cache (ha pasado
    # antes): sin leer_manual_plan no hay notas manuales que mostrar.
    if not hasattr(plan, "leer_manual_plan"):
        return _vacio(columnas), {
            "nivel": "warning",
            "texto": (
                "El servidor tiene una versión antigua de `plan.py`: **no se pudieron "
                "leer las notas manuales (columna I) de las hojas diarias**, así que el "
                "cruce solo usa la hoja `Casos_Ven`. En Streamlit Cloud, *Manage app* → "
                "menú ⋮ → **Reboot**."
            ),
        }

    try:
        manual = plan.leer_manual_plan(plan_bytes)
    except Exception as exc:
        # ErrorPlan = el plan no trae hojas diarias: se cruza solo Casos_Ven.
        if isinstance(exc, plan.ErrorPlan):
            return _vacio(columnas), {
                "nivel": "info",
                "texto": (
                    "El Plan de Trabajo no trae hojas diarias: el cruce se hizo solo "
                    "con la hoja `Casos_Ven`."
                ),
            }
        return _vacio(columnas), {
            "nivel": "warning",
            "texto": (
                "No se pudieron leer las notas manuales de las hojas diarias del plan "
                f"({type(exc).__name__}: {exc}). El cruce solo usa `Casos_Ven`."
            ),
        }

    if manual is None or manual.empty:
        return _vacio(columnas), None

    manual = manual.copy()
    manual["CLAVE"] = manual[plan.COL_ID_DIARIO].map(core.clave_caso)
    manual = manual[manual["CLAVE"].ne("")]
    if manual.empty:
        return _vacio(columnas), None

    manual = manual.sort_values(["CLAVE", "FECHA_HOJA"])
    # .last() toma, columna por columna, el ultimo valor NO nulo del caso: asi la
    # nota y el vencimiento salen de la hoja mas reciente que los traiga.
    ultimo = manual.groupby("CLAVE", as_index=False).last()
    notas = manual.groupby("CLAVE")[plan.COL_NOTA_NORMALIZADA].apply(
        lambda serie: " / ".join(dict.fromkeys(n for n in serie if n))
    )
    ultimo[P_NOTAS] = ultimo["CLAVE"].map(notas).fillna("")

    def columna(nombre, reemplazo):
        return ultimo[nombre] if nombre in ultimo.columns else reemplazo

    salida = pd.DataFrame({
        "CLAVE": ultimo["CLAVE"].values,
        P_HOJA: columna("HOJA", "").values,
        P_FECHA_HOJA: columna("FECHA_HOJA", pd.NaT).values,
        P_APERTURA: columna(plan.COL_APERTURA, pd.NaT).values,
        P_VENCE: columna(plan.COL_VENCIMIENTO, pd.NaT).values,
        P_ASIGNATARIO: columna(plan.COL_ASIGNATARIO, "").values,
        P_UBICACION: columna(plan.COL_UBICACION_DIARIA, "").values,
        P_NOTA: columna(plan.COL_NOTA, "").values,
        P_NOTA_NORM: columna(plan.COL_NOTA_NORMALIZADA, "").values,
        P_NOTAS: ultimo[P_NOTAS].values,
        P_ESTADO: columna(plan.COL_ESTADO_DIARIO, "").values,
    })
    for col in (P_HOJA, P_ASIGNATARIO, P_UBICACION, P_NOTA, P_NOTA_NORM, P_NOTAS, P_ESTADO):
        salida[col] = salida[col].map(_texto)
    return salida, None


def _plan_vencidos(plan_bytes) -> pd.DataFrame:
    """Casos de la hoja Casos_Ven con su justificacion y su culpa."""
    columnas = ["CLAVE", P_JUSTIF, P_CULPA, P_CATEGORIA, P_FECHA_VEN]
    try:
        casos = plan.vista_vencidos(plan_bytes)["casos"]
    except Exception:
        return _vacio(columnas)

    if casos is None or casos.empty:
        return _vacio(columnas)

    def columna(nombre, reemplazo=""):
        return casos[nombre] if nombre in casos.columns else reemplazo

    marco = pd.DataFrame({
        "CLAVE": casos[plan.COL_CASO].map(core.clave_caso),
        P_FECHA_VEN: columna(plan.COL_FECHA_CREACION, pd.NaT),
        P_JUSTIF: columna(plan.COL_JUSTIFICACION),
        P_CULPA: columna(plan.COL_CULPA),
        P_CATEGORIA: columna(plan.COL_CATEGORIA),
    })
    marco = marco[marco["CLAVE"].ne("")]
    # Si el caso viniera dos veces en Casos_Ven se conserva la primera fila.
    marco = marco.drop_duplicates(subset=["CLAVE"], keep="first")
    for col in (P_JUSTIF, P_CULPA, P_CATEGORIA):
        marco[col] = marco[col].map(_texto)
    return marco.reset_index(drop=True)


def _plantilla(df_plantilla: pd.DataFrame) -> pd.DataFrame:
    """Un renglon por N° DE CASO de la plantilla, con sus copias y su estado."""
    columnas = ["CLAVE", L_FILAS, L_FILAS_EXCEL, L_DUPLICADO, L_ESTADO, L_ESTADO_CALC,
                L_RESOLUCION, L_CREACION, L_VENCE, L_TECNICO, L_CERRADO]
    if df_plantilla is None or df_plantilla.empty:
        return _vacio(columnas)

    marco = df_plantilla.copy()
    marco["CLAVE"] = marco[core.COL_CASO].map(core.clave_caso)
    marco = marco[marco["CLAVE"].ne("")]

    filas = []
    for clave, grupo in marco.groupby("CLAVE"):
        def serie(nombre):
            return grupo[nombre] if nombre in grupo.columns else pd.Series(dtype="object")

        resolucion = serie("FECHA_RESOLUCION").dropna()
        creacion = serie("FECHA_CREACION").dropna()
        vencimiento = serie("FECHA_VENCIMIENTO").dropna()
        duplicado = serie(core.COL_ES_DUPLICADO)
        excel = sorted(int(f) for f in serie(core.COL_FILA_EXCEL).dropna())

        # En un caso repetido se ven los DOS estados calculados ("ROJO | CERRADO
        # TARDE"): es la explicacion de por que el caso parecia abierto y cerrado.
        calculados = [
            _texto(estado) for estado in serie(core.COL_ESTADO_CALCULADO)
            if _texto(estado)
        ]
        calculados = list(dict.fromkeys(calculados))

        filas.append({
            "CLAVE": clave,
            L_FILAS: int(len(grupo)),
            L_FILAS_EXCEL: ", ".join(str(f) for f in excel),
            L_DUPLICADO: bool(duplicado.eq(True).any()),
            L_ESTADO: _texto(serie("ESTADO").iloc[0]) if len(grupo) else "",
            L_ESTADO_CALC: " | ".join(calculados),
            L_RESOLUCION: resolucion.min() if not resolucion.empty else pd.NaT,
            L_CREACION: creacion.min() if not creacion.empty else pd.NaT,
            L_VENCE: vencimiento.min() if not vencimiento.empty else pd.NaT,
            L_TECNICO: _texto(serie("TECNICO").iloc[0]) if len(grupo) else "",
            L_CERRADO: bool(resolucion.notna().any()),
        })
    return pd.DataFrame(filas, columns=columnas)


# ---------------------------------------------------------------------------
# Cruce
# ---------------------------------------------------------------------------

def conciliar(
    plan_bytes,
    df_plantilla: pd.DataFrame,
    *,
    tolerancia_minutos: float = TOLERANCIA_MINUTOS,
) -> dict:
    """
    Cruza TODO el plan (hojas diarias + Casos_Ven) contra la plantilla SLA.

    Devuelve:
        tabla        una fila por N° DE CASO, con lo que dice cada archivo y los
                     desajustes encontrados (ordenada por cantidad de desajustes).
        conteos      ``{tipo: cantidad}`` con los tipos que aparecen.
        resumen      tabla ``Tipo | Casos | Qué significa`` en orden de gravedad.
        total_plan, total_plantilla, coinciden, desajustes, columnas
    """
    avisos: list[dict] = []
    manual, aviso_manual = _plan_manual(plan_bytes)
    if aviso_manual:
        avisos.append(aviso_manual)
    vencidos = _plan_vencidos(plan_bytes)
    plantilla = _plantilla(df_plantilla)

    vacio = {
        "tabla": pd.DataFrame(columns=COLUMNAS_TABLA),
        "conteos": {},
        "resumen": pd.DataFrame(columns=["Tipo", "Casos", "Qué significa"]),
        "total_plan": 0,
        "total_plantilla": 0,
        "coinciden": 0,
        "desajustes": 0,
        "columnas": COLUMNAS_TABLA,
        "avisos": avisos,
    }

    if manual.empty and vencidos.empty and plantilla.empty:
        return vacio

    # --- lado del plan: hojas diarias + Casos_Ven --------------------------
    if manual.empty:
        plan_lado = vencidos.copy()
    elif vencidos.empty:
        plan_lado = manual.copy()
    else:
        plan_lado = manual.merge(vencidos, on="CLAVE", how="outer")
    plan_lado["EN_PLAN"] = True

    tabla = plan_lado.merge(plantilla, on="CLAVE", how="outer")
    tabla["EN_PLAN"] = tabla["EN_PLAN"].eq(True)
    tabla["EN_PLANTILLA"] = tabla[L_FILAS].notna()

    # --- desajustes caso por caso -----------------------------------------
    tipos_por_fila = []
    for _, fila in tabla.iterrows():
        tipos = []
        en_plan = _bool(fila.get("EN_PLAN"))
        en_plantilla = _bool(fila.get("EN_PLANTILLA"))

        if not en_plantilla:
            tipos.append(TIPO_SOLO_PLAN)
        if not en_plan:
            tipos.append(TIPO_SOLO_PLANTILLA)
        if _bool(fila.get(L_DUPLICADO)):
            tipos.append(TIPO_DUPLICADO)

        if en_plan and en_plantilla:
            if not _bool(fila.get(L_CERRADO)) and menciona_cierre(
                fila.get(P_NOTA_NORM), fila.get(P_NOTAS), fila.get(P_JUSTIF),
                fila.get(P_ESTADO),
            ):
                tipos.append(TIPO_CIERRE)

            apertura = fila.get(P_APERTURA)
            if pd.isna(apertura):
                apertura = fila.get(P_FECHA_VEN)
            creacion = fila.get(L_CREACION)
            if pd.notna(apertura) and pd.notna(creacion):
                if not _iguales(apertura, creacion, tolerancia_minutos):
                    if _iguales(_invertir(apertura), creacion, tolerancia_minutos):
                        tipos.append(TIPO_FECHA_INVERTIDA)
                    else:
                        tipos.append(TIPO_FECHA_DISTINTA)

            vence_plan = fila.get(P_VENCE)
            vence_plantilla = fila.get(L_VENCE)
            if pd.notna(vence_plan) and pd.notna(vence_plantilla):
                if not _iguales(vence_plan, vence_plantilla, tolerancia_minutos):
                    if _iguales(_invertir(vence_plan), vence_plantilla,
                                tolerancia_minutos):
                        tipos.append(TIPO_VENCIMIENTO_INVERTIDO)
                    else:
                        tipos.append(TIPO_VENCIMIENTO)

        tipos_por_fila.append(tipos)

    # --- tabla presentable -------------------------------------------------
    salida = pd.DataFrame({
        "Caso": tabla["CLAVE"],
        "Desajuste": [
            " · ".join(f"{ICONO_TIPO[t]} {t}" for t in tipos) if tipos else "✅ Coincide"
            for tipos in tipos_por_fila
        ],
        "Plan: Hoja": tabla.get(P_HOJA, ""),
        "Plan: Vence": tabla[P_VENCE].map(_fecha) if P_VENCE in tabla else "—",
        "Plan: Asignatario": tabla.get(P_ASIGNATARIO, ""),
        "Plan: Ubicación": tabla.get(P_UBICACION, ""),
        "Plan: Nota (col I)": tabla.get(P_NOTA_NORM, ""),
        "Plan: Estado": tabla.get(P_ESTADO, ""),
        "Plan: Justificación": tabla.get(P_JUSTIF, ""),
        "Plan: Culpa": tabla.get(P_CULPA, ""),
        "Plantilla: Filas": tabla.get(L_FILAS, 0),
        "Plantilla: Estado": tabla.get(L_ESTADO_CALC, ""),
        "Plantilla: Resolución": tabla[L_RESOLUCION].map(_fecha) if L_RESOLUCION in tabla else "—",
        "Plantilla: Vence": tabla[L_VENCE].map(_fecha) if L_VENCE in tabla else "—",
        "Plantilla: Técnico": tabla.get(L_TECNICO, ""),
    })
    salida["Plantilla: Filas"] = (
        pd.to_numeric(salida["Plantilla: Filas"], errors="coerce").fillna(0).astype(int)
    )
    for columna in salida.columns:
        if salida[columna].dtype == object:
            salida[columna] = salida[columna].fillna("").astype(str)

    salida["_tipos"] = tipos_por_fila
    salida = (
        salida.assign(_orden=[len(t) for t in tipos_por_fila])
        .sort_values(["_orden", "Caso"], ascending=[False, True])
        .drop(columns=["_orden"])
        .reset_index(drop=True)
    )

    conteos = {}
    for tipo in TIPOS:
        cantidad = int(sum(1 for tipos in tipos_por_fila if tipo in tipos))
        if cantidad:
            conteos[tipo] = cantidad

    resumen = pd.DataFrame(
        [
            {"Tipo": f"{ICONO_TIPO[t]} {t}", "Casos": conteos[t],
             "Qué significa": DESCRIPCION_TIPO[t]}
            for t in TIPOS if t in conteos
        ],
        columns=["Tipo", "Casos", "Qué significa"],
    )

    return {
        "tabla": salida,
        "conteos": conteos,
        "resumen": resumen,
        "total_plan": int(tabla["EN_PLAN"].sum()),
        "total_plantilla": int(tabla["EN_PLANTILLA"].sum()),
        "coinciden": int(len(tabla) - sum(1 for t in tipos_por_fila if t)),
        "desajustes": int(sum(1 for t in tipos_por_fila if t)),
        "columnas": COLUMNAS_TABLA,
        "avisos": avisos,
    }
