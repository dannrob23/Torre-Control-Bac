"""
validacion.py - Modulo de validacion de datos de la torre de control.

Que es
------
EL MODULO QUE LE DICE A LA TORRE EN QUE FILA ESTA EL PROBLEMA. No es nuevo: los
cuatro chequeos ya existen en el sistema (contraste de fechas Casos_Ven <-> hojas
diarias, conciliacion plantilla <-> plan, deteccion de casos repetidos e
integridad de filas). Lo que hace este modulo es reunirlos, cada uno con:

    * el conteo,
    * una explicacion en una linea,
    * Y LA FILA EXACTA del Excel donde esta escrito asi,
    * la tabla y el CSV para trabajarla.

Nada de lo que ya existia se cambia: aqui solo se expone junto y ordenado, para
que viva en su propio modulo y la pantalla de turno no tenga scroll.

Los 4 chequeos
--------------
    fechas    -> la fecha viene invertida (mes/día) o no coincide con el plan.
    nombres   -> una grafía de técnico que el sistema no reconoce (manda el caso
                 al reporte equivocado).
    repetidos -> el mismo N° DE CASO en varias filas de la plantilla.
    causa     -> casos cerrados tarde que no están en el Plan de Trabajo y por
                 lo tanto no tienen causa escrita.

Uso:
    import validacion
    info = validacion.chequeos(plan_bytes, df_completo)
    info["chequeos"], info["conteos"], info["pendientes"]
"""
from __future__ import annotations

import difflib

import pandas as pd

import conciliacion
import core
import turno

# Tipos del cruce que son problema de FECHA.
TIPOS_FECHA = (
    conciliacion.TIPO_FECHA_INVERTIDA,
    conciliacion.TIPO_FECHA_DISTINTA,
    conciliacion.TIPO_VENCIMIENTO_INVERTIDO,
    conciliacion.TIPO_VENCIMIENTO,
)

# Columnas de las tablas del modulo (en el orden en que se muestran).
COLUMNAS_FECHAS = ["Caso", "Qué pasa", "Plan: vence", "Plantilla: vence",
                   "Fila plantilla", "Hoja del plan"]
COLUMNAS_NOMBRES = ["Nombre en el Excel", "Casos", "Fila plantilla",
                    "¿Quién es?", "Efecto"]
COLUMNAS_REPETIDOS = ["Caso", "Veces", "Filas de la plantilla", "Estado de cada copia"]
COLUMNAS_CAUSA = ["Caso", "Fila plantilla", "Técnico", "Vencido", "Vence"]


def _filas_por_caso(df_completo: pd.DataFrame) -> dict:
    """Mapa caso -> filas que ocupa en la hoja PLANTILLA (2 = primera de datos)."""
    if core.COL_FILA_EXCEL not in df_completo.columns or df_completo.empty:
        return {}
    agrupado = df_completo.groupby(
        df_completo[core.COL_CASO].map(core.clave_caso)
    )[core.COL_FILA_EXCEL].apply(
        lambda serie: ", ".join(str(int(v)) for v in sorted(serie.dropna()))
    )
    return agrupado.to_dict()


def _fechas(plan_bytes, df_completo: pd.DataFrame, filas: dict) -> dict:
    """Chequeo 1: fechas del plan contra las de la plantilla."""
    ayuda = (
        "Se contrastan dos fuentes: la fecha de la plantilla del banco y la de las "
        "hojas diarias del Plan de Trabajo. 'Fila plantilla' es la fila que ocupa el "
        "caso en la hoja PLANTILLA (2 = primera fila de datos)."
    )
    if not plan_bytes:
        return {"id": "fechas", "icono": "📅", "titulo": "Fechas", "alerta": 0,
                "resumen": "Falta el Plan de Trabajo para contrastar las fechas.",
                "ayuda": ayuda, "tabla": pd.DataFrame(columns=COLUMNAS_FECHAS)}

    try:
        cruce = conciliacion.conciliar(plan_bytes, df_completo)
    except Exception as exc:
        return {"id": "fechas", "icono": "📅", "titulo": "Fechas", "alerta": 0,
                "resumen": f"No se pudo contrastar con el plan ({type(exc).__name__}).",
                "ayuda": ayuda, "tabla": pd.DataFrame(columns=COLUMNAS_FECHAS)}

    tabla = cruce.get("tabla")
    detalle = pd.DataFrame(columns=COLUMNAS_FECHAS)
    if tabla is not None and not tabla.empty:
        con_fecha = tabla[tabla["_tipos"].map(lambda tipos: bool(set(tipos) & set(TIPOS_FECHA)))]
        if not con_fecha.empty:
            detalle = pd.DataFrame({
                "Caso": con_fecha["Caso"].values,
                "Qué pasa": [
                    " · ".join(t for t in tipos if t in TIPOS_FECHA)
                    for tipos in con_fecha["_tipos"]
                ],
                "Plan: vence": con_fecha["Plan: Vence"].values,
                "Plantilla: vence": con_fecha["Plantilla: Vence"].values,
                "Fila plantilla": con_fecha["Caso"].map(filas).fillna("").values,
                "Hoja del plan": con_fecha["Plan: Hoja"].values,
            })

    return {
        "id": "fechas", "icono": "📅", "titulo": "Fechas", "alerta": int(len(detalle)),
        "resumen": (
            f"**{len(detalle)} caso(s)** con la fecha en duda: viene invertida "
            "(mes/día) o no coincide entre la plantilla y el plan."
            if len(detalle) else "Todas las fechas cuadran entre los dos archivos."
        ),
        "ayuda": ayuda, "tabla": detalle,
        "_cruce": cruce,
    }


def _nombres(df_completo: pd.DataFrame, filas: dict) -> dict:
    """
    Chequeo 2: grafías de técnico que NO son las del diccionario oficial.

    OJO: no se listan solo los nombres desconocidos. Tambien se listan los que el
    sistema ya corrige solo con un ALIAS (ej. 'JHON...' -> 'JOHAN...'): asi la
    torre VE en que fila esta escrita la grafía mala y puede arreglar el Excel.
    Si no se mostraran, el problema quedaria invisble solo porque la app lo tapa.
    """
    ayuda = (
        "El diccionario de regiones (core.TECNICOS_REGION) decide si un caso es de "
        "Bogotá o de Regionales. Una grafía distinta manda el caso al reporte "
        "equivocado: por eso hay que ver la fila exacta donde está escrita así."
    )
    tabla = pd.DataFrame(columns=COLUMNAS_NOMBRES)
    if df_completo is not None and not df_completo.empty and "TECNICO" in df_completo.columns:
        oficiales = {core.normalizar_texto(k) for k in core.TECNICOS_REGION}
        filas_tabla = []
        for tecnico, grupo in df_completo.groupby("TECNICO"):
            nombre = str(tecnico).strip()
            if not nombre or core.normalizar_texto(nombre) in oficiales:
                continue

            region = core.obtener_region(nombre)
            if region != core.REGION_DESCONOCIDA:
                # Ya se corrige con un alias: hay que arreglar la grafía del Excel.
                quien = f"{core.nombre_canonico_tecnico(nombre)} → {region}"
                efecto = "el sistema ya lo corrige con un alias; arreglar la grafía en el Excel"
            else:
                cerca = difflib.get_close_matches(
                    core.normalizar_texto(nombre),
                    [core.normalizar_texto(k) for k in core.TECNICOS_REGION],
                    n=1, cutoff=0.85,
                )
                quien = (
                    f"¿{cerca[0]}?" if cerca else "no se parece a ninguno del diccionario"
                )
                efecto = "sus casos caen en REGIONALES por la regla de Bogotá"

            filas_tabla.append({
                "Nombre en el Excel": nombre,
                "Casos": int(len(grupo)),
                "Fila plantilla": ", ".join(
                    str(int(v)) for v in sorted(
                        grupo.get(core.COL_FILA_EXCEL, pd.Series(dtype=int)).dropna()
                    )
                ),
                "¿Quién es?": quien,
                "Efecto": efecto,
            })
        if filas_tabla:
            tabla = pd.DataFrame(filas_tabla, columns=COLUMNAS_NOMBRES)

    return {
        "id": "nombres", "icono": "👤", "titulo": "Nombres", "alerta": int(len(tabla)),
        "resumen": (
            f"**{len(tabla)} grafía(s)** de técnico que no son las del diccionario."
            if len(tabla) else "Todos los nombres de técnico son los oficiales."
        ),
        "ayuda": ayuda, "tabla": tabla,
    }


def _repetidos(df_completo: pd.DataFrame) -> dict:
    """Chequeo 3: el mismo N° DE CASO en varias filas."""
    ayuda = (
        "Se marca DUPLICADO y la fila queda FUERA de los conteos, las alertas y las "
        "métricas hasta que se unifique en el Excel. La app no borra ni elige copias."
    )
    resumen = core.resumen_duplicados(df_completo)
    tabla = pd.DataFrame(columns=COLUMNAS_REPETIDOS)
    if resumen is not None and not resumen.empty:
        tabla = pd.DataFrame({
            "Caso": resumen["CASO"].values,
            "Veces": resumen["VECES"].values,
            "Filas de la plantilla": resumen["FILAS_EXCEL"].values,
            "Estado de cada copia": resumen["ESTADO_POR_FILA"].values,
        }, columns=COLUMNAS_REPETIDOS)

    return {
        "id": "repetidos", "icono": "🔁", "titulo": "Casos repetidos",
        "alerta": int(len(tabla)),
        "resumen": (
            f"**{len(tabla)} caso(s) repetidos**: cada copia se clasifica aparte, así "
            "que el caso puede salir cerrado y a la vez alertar como abierto."
            if len(tabla) else "Ningún N° DE CASO repetido."
        ),
        "ayuda": ayuda, "tabla": tabla,
    }


def _causa(df_completo: pd.DataFrame, culpa: dict, filas: dict) -> dict:
    """Chequeo 4: cerrados tarde sin causa escrita en el Plan de Trabajo."""
    ayuda = (
        "La causa (columna CULPA del Plan: TÉCNICO, LOGÍSTICO, ALIADO, BANCO…) es lo "
        "que permite no cargarle al técnico un vencimiento que no controla."
    )
    tabla = pd.DataFrame(columns=COLUMNAS_CAUSA)
    total_tarde = 0
    if df_completo is not None and not df_completo.empty:
        tarde = df_completo[df_completo["ESTADO"] == core.CERRADO_TARDE]
        total_tarde = int(len(tarde))
        claves = tarde[core.COL_CASO].map(core.clave_caso)
        sin_causa = tarde[claves.map(culpa).fillna("").eq("")]
        if not sin_causa.empty:
            tabla = pd.DataFrame({
                "Caso": sin_causa[core.COL_CASO].values,
                "Fila plantilla": (
                    sin_causa[core.COL_CASO].map(core.clave_caso).map(filas).fillna("").values
                    if core.COL_FILA_EXCEL in sin_causa.columns else ""
                ),
                "Técnico": sin_causa["TECNICO"].values,
                "Vencido": sin_causa["TIEMPO_VENCIDO"].values,
                "Vence": pd.to_datetime(
                    sin_causa["FECHA_VENCIMIENTO"], errors="coerce"
                ).dt.strftime("%d/%m/%Y %H:%M").values,
            }, columns=COLUMNAS_CAUSA)

    return {
        "id": "causa", "icono": "📋", "titulo": "Causa documentada",
        "alerta": int(len(tabla)),
        "resumen": (
            f"**{len(tabla)} de {total_tarde}** casos cerrados tarde **no están en el "
            "Plan de Trabajo**, así que no tienen causa escrita. Ahí no se puede saber "
            "si la culpa fue del técnico o de la logística."
        ),
        "ayuda": ayuda, "tabla": tabla,
    }


def chequeos(ruta_plan, df_completo: pd.DataFrame) -> dict:
    """
    Corre los 4 chequeos y devuelve todo lo que necesita el modulo de validacion.

    Returns:
        ``{"chequeos": [...], "conteos": {id: alertas}, "pendientes": n,
        "total": n, "cruce": {...}}``
    """
    filas = _filas_por_caso(df_completo) if df_completo is not None else {}
    culpa = turno.culpa_por_caso(ruta_plan) if ruta_plan else {}

    lista = [
        _fechas(ruta_plan, df_completo, filas),
        _nombres(df_completo, filas),
        _repetidos(df_completo),
        _causa(df_completo, culpa, filas),
    ]
    conteos = {chequeo["id"]: chequeo["alerta"] for chequeo in lista}
    cruce = lista[0].pop("_cruce", None)

    return {
        "chequeos": lista,
        "conteos": conteos,
        "pendientes": sum(1 for valor in conteos.values() if valor),
        "total": len(lista),
        "cruce": cruce,
    }
