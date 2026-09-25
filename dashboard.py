"""
dashboard.py - Dashboard interactivo de la Torre de Control de SLA (Colsof - Banco Agrario).

Ejecutar con:
    streamlit run dashboard.py

Caracteristicas:
    - Lee la misma plantilla Excel y aplica exactamente la misma logica que alertas_windows.py
    - KPIs por semaforo v2 (7 estados) y tabla interactiva de los casos en ventana
    - Filtros por Tecnico, Region y Estado, mas busqueda libre
    - Refresco manual y autorrefresco opcional (recalcula el tiempo restante)
    - Descarga del resultado filtrado a CSV
    - Seccion "Notificaciones y Metricas": historial SQLite (historial.py),
      metricas por tecnico (core.metricas_por_tecnico) y rankings de gestion
    - Seccion "Notificaciones pendientes a tecnicos": tabla de pendientes por
      tecnico (avisos.resumen_pendientes), generacion del aviso, boton nativo de
      copiar (st.code), envio por Telegram en texto plano y registro manual en
      el historial

Notas de version 2:
    - El semaforo tiene 7 estados: VERDE, AMARILLO, NARANJA, ROJO (activos),
      CERRADO OK, CERRADO TARDE (cierres) y SIN VENCIMIENTO.
    - Los casos cerrados se clasifican comparando la resolucion contra el
      vencimiento, por lo que los KPI se calculan sobre TODOS los casos
      (resultado.df_completo) y no solo sobre la ventana de notificacion.
"""

from __future__ import annotations

import os
from datetime import datetime, timedelta

import altair as alt
import pandas as pd
import streamlit as st

from core import (
    AMARILLO,
    BOGOTA,
    CERRADO_OK,
    CERRADO_TARDE,
    COL_CASO,
    COL_CIUDAD,
    COL_DEPARTAMENTO,
    COL_ES_DUPLICADO,
    COL_ESTADO_CALCULADO,
    COL_FILAS_REPETIDAS,
    COL_REGIONAL,
    COL_VENCIMIENTO,
    DUPLICADO,
    ESTADOS_ALERTA,
    ICONO_ESTADO,
    NARANJA,
    ORDEN_ESTADO,
    REGIONALES,
    REGION_DESCONOCIDA,
    ROJO,
    SIN_VENCIMIENTO,
    TODOS_LOS_ESTADOS,
    VERDE,
    ErrorLecturaExcel,
    ahora_colombia,
    calcular_tablero,
    leer_casos,
    localizar_excel,
    metricas_por_tecnico,
    parte_por_region,
    resumen_duplicados,
)
from historial import (
    CANAL_WHATSAPP,
    RESULTADO_GENERADO,
    Historial,
)

# Avisos por tecnico (composicion del mensaje) y envio por Telegram.
import avisos
import telegram_notifier

# Proteccion de datos: si el tablero se publica en la web, se anonimiza.
import anonimizar

# Control de acceso: el tablero exige usuario y contrasena antes de mostrar nada.
import auth

# Rediseno "Accion primero" (Prototipo A): barra de control, franja de foco,
# barra de semaforo segmentada y listas de accion.
import vista

# Alias del modulo vista. Hace falta porque dentro de render_plan_trabajo() la
# variable local "vista" guarda la lectura del plan y taparia al modulo. Ademas
# se resuelve el panel con getattr (ver _dibujar_conciliacion): si el servidor
# quedo con un vista.py viejo en cache, se avisa en la pestana en vez de tumbar
# TODO el tablero con un ImportError al arrancar.
vista_panel = vista

# Analitica del Plan de Trabajo mensual (cartera vencida, envejecimiento y ANS).
# Es independiente de la plantilla SLA: trabaja sobre otro archivo.
import plan

# Calculos del turno: avance del dia, acumulado del mes y grupos de alerta en el
# orden de atencion de la torre (naranja, amarillo, rojo, proximos 3 dias).
import turno

# Sistema de diseño CSS personalizado
import estilos_css

# Refresco automatico cada 60 s (opcional, desactivado por defecto para no
# interrumpir al usuario mientras filtra).
INTERVALO_AUTOREFRESCO_S = 60

# SELLO DE VERSION de la interfaz. Se muestra en la barra lateral y en la linea
# de estado: sirve para saber de un vistazo QUE version esta corriendo el
# servidor. Si despues de un push la app sigue mostrando un sello viejo, el
# despliegue de Streamlit Cloud no ha tomado el ultimo commit (Manage app ->
# Reboot), no es que falte codigo en GitHub.
VERSION_APP = "2026-09-25.1700"


def selector_vista(opciones: list[str], defecto: str, clave: str, ayuda: str = ""):
    """
    Selector de una sola opcion (Todas / Bogota / Regionales).

    Usa el componente segmentado de Streamlit cuando existe y, si el servidor
    tuviera una version antigua que no lo trae, cae a un radio horizontal: asi la
    pantalla funciona en cualquier version desplegada.
    """
    creador = getattr(st, "segmented_control", None)
    if creador is not None:
        return creador("Vista", opciones, default=defecto, key=clave, help=ayuda) or defecto
    return st.radio("Vista", opciones, index=opciones.index(defecto),
                    horizontal=True, key=clave, help=ayuda)


def dibujar_graficos_altair(filtrado: pd.DataFrame) -> None:
    """Dibuja gráficos interactivos de casos por región y técnico con la paleta semántica SLA."""
    if filtrado is None or filtrado.empty:
        st.info("No hay datos para generar los gráficos.")
        return

    g1, g2 = st.columns(2)
    with g1:
        st.subheader("📍 Casos por región (por estado SLA)")
        df_region = (
            filtrado.groupby(["REGION_TECNICO", "ESTADO"])
            .size()
            .reset_index(name="CASOS")
        )
        color_scale = alt.Scale(
            domain=[ROJO, CERRADO_TARDE, NARANJA, AMARILLO, VERDE, SIN_VENCIMIENTO,
                    CERRADO_OK, DUPLICADO],
            range=["#B91C1C", "#8B0000", "#C2410C", "#A16207", "#15803D", "#6B7280",
                   "#15803D", "#7C3AED"],
        )
        chart_region = (
            alt.Chart(df_region)
            .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("REGION_TECNICO:N", title="Región", sort="-y"),
                y=alt.Y("CASOS:Q", title="Cantidad de Casos"),
                color=alt.Color("ESTADO:N", title="Estado SLA", scale=color_scale),
                tooltip=["REGION_TECNICO", "ESTADO", "CASOS"],
            )
            .properties(height=340)
            .interactive()
        )
        st.altair_chart(chart_region, width="stretch")

    with g2:
        st.subheader("👷 Carga por técnico (Top 15)")
        top_tec = (
            filtrado["TECNICO"]
            .value_counts()
            .head(15)
            .reset_index(name="CASOS")
        )
        top_tec.columns = ["TECNICO", "CASOS"]
        chart_tec = (
            alt.Chart(top_tec)
            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4, color="#0B5D0B")
            .encode(
                y=alt.Y("TECNICO:N", title="Técnico", sort="-x"),
                x=alt.X("CASOS:Q", title="Casos Asignados"),
                tooltip=["TECNICO", "CASOS"],
            )
            .properties(height=340)
        )
        st.altair_chart(chart_tec, width="stretch")


# ---------------------------------------------------------------------------
# Estilos del semaforo v2
# ---------------------------------------------------------------------------
# 7 estados: los 4 activos por tiempo, los 2 de cierre y los casos sin fecha.
COLORES_FONDO = {
    ROJO: "#FFE0E0",             # rojo claro  -> vencido sin cerrar
    CERRADO_TARDE: "#8B0000",    # rojo oscuro -> incumplimiento consumado
    NARANJA: "#FFE0B2",          # naranja claro -> menos de 1 hora
    AMARILLO: "#FFF6D5",         # amarillo claro -> entre 1 y 4 horas
    VERDE: "#E3F7E3",            # verde claro -> mas de 4 horas
    SIN_VENCIMIENTO: "#EFEFEF",  # gris -> sin fecha en la plantilla
    CERRADO_OK: "#D6F5D6",       # verde claro -> cerrado a tiempo
    DUPLICADO: "#EDE9FE",        # violeta claro -> N° DE CASO repetido
}
COLORES_TEXTO = {
    ROJO: "#A30000",
    CERRADO_TARDE: "#FFFFFF",
    NARANJA: "#8A4B00",
    AMARILLO: "#7A5B00",
    VERDE: "#136B13",
    SIN_VENCIMIENTO: "#555555",
    CERRADO_OK: "#0B5D0B",
    DUPLICADO: "#5B21B6",
}

# Texto de ayuda de cada tarjeta KPI, en el orden de gravedad de ORDEN_ESTADO.
AYUDA_ESTADO = {
    ROJO: "Vencido y sin cerrar: requiere accion inmediata.",
    CERRADO_TARDE: "Se cerro despues del vencimiento: incumplimiento del ANS.",
    NARANJA: "Abierto: menos de 1 hora para vencer.",
    AMARILLO: "Abierto: entre 1 y 4 horas para vencer.",
    VERDE: "Abierto: mas de 4 horas para vencer.",
    SIN_VENCIMIENTO: "Activo o cerrado sin fecha de vencimiento en la plantilla.",
    CERRADO_OK: "Cerrado a tiempo (resolucion <= vencimiento).",
    DUPLICADO: (
        "El N° DE CASO aparece varias veces en la plantilla: se muestra con sus "
        "copias, pero no se cuenta en el semáforo ni en las métricas."
    ),
}

# Columnas de la tabla de casos (incluye el tiempo vencido y la prediccion v2).
COLUMNAS_TABLA = {
    COL_CASO: "Caso",
    "TECNICO": "Tecnico",
    "REGION_TECNICO": "Region",
    COL_CIUDAD: "Ciudad",
    COL_REGIONAL: "Regional",
    COL_DEPARTAMENTO: "Departamento",
    COL_VENCIMIENTO: "Vencimiento",
    "TIEMPO_RESTANTE": "Tiempo restante",
    "TIEMPO_VENCIDO": "Tiempo vencido",
    "HORAS_RESTANTES": "Horas restantes",
    "PREDICCION": "Prediccion",
    "ESTADO": "Estado",
    COL_ESTADO_CALCULADO: "Estado calculado",
    COL_FILAS_REPETIDAS: "Filas repetidas",
    "ICONO": " ",
}

# Columnas del historial de notificaciones (historial.py), para tablas vacias.
COLUMNAS_NOTIF_TECNICO = [
    "TECNICO", "REGION", "NOTIFICACIONES", "CASOS_DISTINTOS",
    "ULTIMA_NOTIFICACION", "ROJOS", "NARANJAS", "AMARILLOS", "FALLIDAS",
]
COLUMNAS_NOTIF = [
    "id", "fecha_hora", "caso", "tecnico", "region", "estado",
    "horas_restantes", "horas_vencido", "canal", "resultado", "detalle",
]

# Filas visibles por tabla (las tablas grandes se recortan solo al mostrar).
MAX_FILAS_DETALLE = 500


# ---------------------------------------------------------------------------
# Carga de datos con cache invalidada por fecha de modificacion del archivo
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def cargar(ruta: str, firma: float, momento_iso: str):
    """
    Carga y calcula el tablero desde un archivo Excel.

    'firma' (mtime del archivo) y 'momento_iso' forman parte de la clave de cache:
    si el Excel cambia, o si el usuario pulsa "Recalcular", la cache se invalida.

    Devuelve la tupla:
        (df_ventana, df_completo, total_hoja, total_activos, dias_ventana,
         avisos, df_crudo, nombre_archivo)

    df_crudo son las filas TAL CUAL vienen del Excel: se devuelve para poder
    auditar la integridad (que ninguna fila diligenciada se pierda al procesar).
    """
    df_crudo = leer_casos(ruta)
    resultado = calcular_tablero(df_crudo=df_crudo, momento=datetime.fromisoformat(momento_iso))
    return (
        resultado.df,
        resultado.df_completo,
        resultado.total_hoja,
        resultado.total_activos,
        resultado.dias_ventana,
        list(resultado.warnings),
        df_crudo,
        os.path.basename(ruta),
    )


@st.cache_data(show_spinner=False)
def cargar_desde_bytes(contenido: bytes, firma: float, momento_iso: str,
                       nombre: str = "archivo.xlsx"):
    """
    Carga el tablero desde un Excel SUBIDO por el usuario (en memoria).

    Existe para poder desplegar el tablero en la web sin publicar el archivo:
    el usuario sube la plantilla, se procesa en memoria y nunca se guarda.
    'firma' es el hash del contenido, para invalidar la cache cuando cambia.
    """
    import io
    import warnings

    warnings.filterwarnings("ignore")

    momento = datetime.fromisoformat(momento_iso)
    df_crudo = pd.read_excel(io.BytesIO(contenido), sheet_name="PLANTILLA",
                             engine="openpyxl")
    df_crudo.columns = [
        " ".join(str(c).replace("\xa0", " ").split()) if c is not None else ""
        for c in df_crudo.columns
    ]
    resultado = calcular_tablero(df_crudo=df_crudo, momento=momento)
    return (
        resultado.df,
        resultado.df_completo,
        resultado.total_hoja,
        resultado.total_activos,
        resultado.dias_ventana,
        list(resultado.warnings),
        df_crudo,
        nombre,
    )


def firma_archivo(ruta: str) -> float:
    """mtime del archivo; -1 si no existe (para que la cache no colisione)."""
    try:
        return os.path.getmtime(ruta)
    except OSError:
        return -1.0


# ---------------------------------------------------------------------------
# Historial de notificaciones (SQLite)
# ---------------------------------------------------------------------------

def leer_historial():
    """
    Lee el historial de notificaciones sin romper el dashboard si falla.

    Devuelve:
        (historial | None, tabla_por_tecnico, tabla_detalle, resumen, error)
    """
    resumen_vacio = {"total": 0, "enviadas": 0, "fallidas": 0, "casos": 0, "tecnicos": 0}
    try:
        historial = Historial()
        return (
            historial,
            historial.por_tecnico(),
            historial.leer(),
            historial.resumen(),
            "",
        )
    except Exception as exc:  # sqlite roto, permisos, etc.
        return (
            None,
            pd.DataFrame(columns=COLUMNAS_NOTIF_TECNICO),
            pd.DataFrame(columns=COLUMNAS_NOTIF),
            resumen_vacio,
            f"{type(exc).__name__}: {exc}",
        )


# ---------------------------------------------------------------------------
# Presentacion de las metricas por tecnico
# ---------------------------------------------------------------------------

def tabla_metricas_visible(tabla: pd.DataFrame) -> pd.DataFrame:
    """
    Copia de la tabla de metricas lista para mostrar.

    CUMPLIMIENTO_PCT se convierte a texto: '—' cuando es None/NaN (tecnico sin
    casos cerrados) en lugar del 'nan' de pandas.
    """
    if tabla is None or tabla.empty:
        return tabla
    vista = tabla.copy()
    if "CUMPLIMIENTO_PCT" in vista.columns:
        vista["CUMPLIMIENTO_PCT"] = vista["CUMPLIMIENTO_PCT"].apply(
            lambda v: "—" if pd.isna(v) else f"{float(v):.1f} %"
        )
    return vista


def formatear_momento(valor) -> str:
    """
    Fecha/hora en formato dd/mm/yyyy HH:MM.

    Se usa en la seccion de notificaciones pendientes: 'nunca' cuando el tecnico
    no tiene notificaciones registradas y '—' cuando el valor es vacio.
    """
    if valor is None:
        return "—"
    try:
        if pd.isna(valor):
            return "—"
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    if not texto or texto.lower() in ("nan", "nat", "none"):
        return "—"
    try:
        return pd.Timestamp(valor).strftime("%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return texto


def tabla_pendientes_visible(tabla: pd.DataFrame) -> pd.DataFrame:
    """
    Copia de avisos.resumen_pendientes lista para mostrar en el dashboard.

    PROXIMO_VENCIMIENTO se formatea como dd/mm/yyyy HH:MM (texto) porque st.dataframe
    no respeta un formato de fecha por columna en todas las versiones; ULTIMA_NOTIFICACION
    queda tal cual lo entrega el historial ('nunca' si no hay registros).
    """
    if tabla is None or tabla.empty:
        return tabla
    vista = tabla.copy()
    if "PROXIMO_VENCIMIENTO" in vista.columns:
        vista["PROXIMO_VENCIMIENTO"] = vista["PROXIMO_VENCIMIENTO"].apply(formatear_momento)
    return vista


def _registrar_casos(historial, df_casos: pd.DataFrame, canal: str, detalle: str,
                     resultado: str = "enviado") -> int:
    """
    Registra en el historial UN registro por caso.

    `resultado` distingue lo que el sistema ENVIO realmente ("enviado") de un
    aviso que solo se GENERO y se pego a mano fuera de la aplicacion
    ("generado"). Asi el log no confunde una accion humana con un envio
    automatico. Devuelve cuantos casos se registraron.
    """
    if historial is None or df_casos is None or df_casos.empty:
        return 0

    momento = ahora_colombia()
    registrados = 0
    for _, fila in df_casos.iterrows():
        historial.registrar(
            caso=str(fila.get(COL_CASO, "S/N")),
            tecnico=str(fila.get("TECNICO", "") or ""),
            region=str(fila.get("REGION_TECNICO", "") or ""),
            estado=str(fila.get("ESTADO", "") or ""),
            horas_restantes=fila.get("HORAS_RESTANTES"),
            horas_vencido=fila.get("HORAS_VENCIDO"),
            canal=canal,
            resultado=resultado,
            detalle=detalle,
            momento=momento,
        )
        registrados += 1
    return registrados


def enviar_aviso_telegram(texto: str, todos: bool = False) -> int:
    """
    Envia el aviso a Telegram con formato HTML (negrillas).

    El texto que llega aqui lo compone avisos.componer_aviso_tecnico_telegram(),
    que ya escapa los datos de la plantilla y deja solo etiquetas <b>/<i>
    permitidas. Por eso se puede usar parse_mode="HTML": antes se enviaba en
    texto plano y las negrillas no existian.

    Devuelve cuantos destinos recibieron el mensaje.
    """
    if not texto or not texto.strip():
        return 0
    if todos:
        return int(telegram_notifier.enviar_a_todos(texto, parse_mode="HTML"))

    destinos = telegram_notifier.destinos_configurados()
    if not destinos:
        return 0
    primero = destinos[0]
    return int(
        telegram_notifier.enviar_mensaje(
            texto,
            parse_mode="HTML",
            chat_id=primero["chat_id"],
            tema_id=primero.get("tema_id") or None,
        )
    )


def ranking(
    tabla: pd.DataFrame,
    columna: str,
    columnas_vista: list[str],
    *,
    ascendente: bool = False,
    filas: int = 10,
    solo_positivos: bool = True,
) -> pd.DataFrame:
    """
    Ranking por tecnico sobre la tabla de metricas.

    Args:
        tabla:          salida de core.metricas_por_tecnico
        columna:        columna por la que ordenar
        columnas_vista: columnas a mostrar
        ascendente:     True = peor primero (util en CUMPLIMIENTO_PCT)
        filas:          maximo de filas devueltas
        solo_positivos: descarta los tecnicos con 0 en la columna (ruido)
    """
    if tabla is None or tabla.empty or columna not in tabla.columns:
        return pd.DataFrame(columns=columnas_vista)

    datos = tabla.copy()
    if solo_positivos:
        datos = datos[datos[columna].fillna(0) > 0]

    if datos.empty:
        return pd.DataFrame(columns=columnas_vista)

    datos = datos.sort_values(
        [columna, "TECNICO"], ascending=[ascendente, True], na_position="last"
    )
    presentes = [c for c in columnas_vista if c in datos.columns]
    return datos.head(filas)[presentes].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Metricas de gestion por tecnico (rankings, cumplimiento, matriz)
# ---------------------------------------------------------------------------

def render_metricas_por_tecnico(df_completo: pd.DataFrame, momento: datetime) -> None:
    """
    Dibuja la tabla de metricas por tecnico, los rankings de gestion, la tasa de
    cumplimiento, la matriz Tecnico x Estado y la descarga CSV.
    """
    tabla_metricas = metricas_por_tecnico(df_completo)
    if tabla_metricas is None or tabla_metricas.empty:
        st.info("No hay casos en la plantilla para calcular metricas por tecnico.")
        return

    # --- Tabla completa de metricas + descarga CSV (punto i) -------------
    st.dataframe(
        tabla_metricas_visible(tabla_metricas),
        width="stretch",
        hide_index=True,
        height=min(500, 40 + 35 * len(tabla_metricas)),
    )
    st.download_button(
        "⬇️ Descargar metricas por tecnico (CSV)",
        data=tabla_metricas.to_csv(index=False, na_rep="").encode("utf-8-sig"),
        file_name=f"metricas_por_tecnico_{momento:%Y%m%d_%H%M}.csv",
        mime="text/csv",
        key="descarga_metricas",
    )
    st.caption(
        "En el CSV, CUMPLIMIENTO_PCT vacio = tecnico sin casos cerrados "
        "(no se puede calcular el porcentaje)."
    )

    # --- d) e) f) Rankings ------------------------------------------------
    st.markdown("##### 🏆 Rankings de carga y riesgo")
    r1, r2, r3 = st.columns(3)
    with r1:
        st.markdown("**⏳ Mas casos proximos a vencer**")
        st.dataframe(
            ranking(
                tabla_metricas,
                "PROXIMOS_VENCER",
                ["TECNICO", "REGION", "PROXIMOS_VENCER", "NARANJA", "AMARILLO"],
            ),
            width="stretch",
            hide_index=True,
            height=440,
        )
    with r2:
        st.markdown("**📦 Mas casos asignados**")
        st.dataframe(
            ranking(
                tabla_metricas,
                "ASIGNADOS",
                ["TECNICO", "REGION", "ASIGNADOS", "ABIERTOS", "CERRADOS"],
            ),
            width="stretch",
            hide_index=True,
            height=440,
        )
    with r3:
        st.markdown("**🔴 Mas casos vencidos**")
        st.dataframe(
            ranking(
                tabla_metricas,
                "VENCIDOS",
                ["TECNICO", "REGION", "VENCIDOS", "HORAS_VENCIDO_TOTAL",
                 "HORAS_VENCIDO_PROMEDIO"],
            ),
            width="stretch",
            hide_index=True,
            height=440,
        )

    # --- g) Tasa de cumplimiento (peor primero) ---------------------------
    st.markdown("##### ✅ Tasa de cumplimiento por tecnico (peor primero)")
    cumplimiento = tabla_metricas.sort_values(
        "CUMPLIMIENTO_PCT", ascending=True, na_position="last"
    )[["TECNICO", "REGION", "CERRADOS", "CERRADOS_TARDE", "CUMPLIMIENTO_PCT"]]
    st.dataframe(
        tabla_metricas_visible(cumplimiento),
        width="stretch",
        hide_index=True,
        height=min(500, 40 + 35 * len(cumplimiento)),
    )
    st.caption(
        "CUMPLIMIENTO_PCT = (cerrados a tiempo / cerrados) x 100. "
        "'—' = tecnico sin casos cerrados."
    )

    # --- h) Matriz Tecnico x Estado ---------------------------------------
    st.markdown("##### 🧮 Matriz Tecnico x Estado")
    matriz = pd.crosstab(df_completo["TECNICO"], df_completo["ESTADO"])
    matriz = (
        matriz.reindex(columns=list(TODOS_LOS_ESTADOS), fill_value=0)
        .fillna(0)
        .astype(int)
    )
    matriz.index.name = "TECNICO"
    st.dataframe(
        matriz,
        width="stretch",
        height=min(500, 40 + 35 * len(matriz)),
    )
    st.caption(
        "Columnas en orden de gravedad: "
        + " · ".join(f"{ICONO_ESTADO[e]} {e}" for e in TODOS_LOS_ESTADOS)
    )


# ---------------------------------------------------------------------------
# Notificaciones pendientes a tecnicos (avisos.py)
# ---------------------------------------------------------------------------

def render_notificaciones_pendientes(
    casos: pd.DataFrame,
    historial,
    estados_incluidos: tuple[str, ...] | None = None,
) -> None:
    """
    Dibuja la seccion "📨 Notificaciones pendientes a tecnicos".

    Flujo de la torre de control: ver la tabla de pendientes por tecnico, elegir
    uno, copiar el aviso con el boton nativo de st.code y pegarlo en WhatsApp, o
    enviarlo por Telegram. En ambos casos queda registro en el Historial.

    Args:
        casos: DataFrame de casos (la ventana operativa o el historico completo).
        historial: instancia de Historial o None si no se pudo abrir.
        estados_incluidos: que estados entran al panel. Por defecto
            core.ESTADOS_ALERTA, que es (ROJO, CERRADO TARDE, NARANJA, AMARILLO)
            y NO incluye VERDE. main() pasa una tupla ampliada cuando el usuario
            marca "incluir tambien los proximos a vencer", porque antes los casos
            por vencer se venciaan sin que nadie los avisara.
    """
    estados = estados_incluidos or ESTADOS_ALERTA

    st.divider()
    st.header("📨 Notificaciones pendientes a técnicos")
    st.caption(
        "Casos que requieren aviso agrupados por técnico: "
        + ", ".join(estados)
        + ". Genere el aviso, cópielo y péguelo en WhatsApp, o envíelo por "
        "Telegram. Cada aviso queda registrado en el historial."
    )

    # Los estados que requieren atencion ya vienen definidos en core.
    pendientes = (
        casos[casos["ESTADO"].isin(estados)].copy()
        if casos is not None and not casos.empty
        else casos
    )
    if pendientes is None or pendientes.empty:
        st.success(
            "✅ No hay casos pendientes de notificar: ningún técnico tiene casos "
            f"en estado {', '.join(estados)}."
        )
        return

    # --- Filtro por Zona / Región (Bogotá vs Regionales) ------------------
    col_f1, col_f2 = st.columns([3, 2])
    with col_f1:
        filtro_region_notif = st.selectbox(
            "🗺️ Alcance regional del despacho",
            [
                "🌐 Todas las regiones (Global)",
                "🌆 Solo Bogotá",
                "⛰️ Solo Regionales (Fuera de Bogotá)",
            ],
            index=0,
            key="filtro_region_despacho_notif",
            help="Permite enviar notificaciones separadas para los técnicos de Bogotá o para las regionales.",
        )
    with col_f2:
        st.caption(
            "💡 **Separe su despacho**: Elija *Solo Bogotá* o *Solo Regionales* "
            "para enviar reportes segmentados."
        )

    if filtro_region_notif == "🌆 Solo Bogotá":
        pendientes = telegram_notifier.filtrar_casos_por_filtro(pendientes, "bogota")
    elif filtro_region_notif == "⛰️ Solo Regionales (Fuera de Bogotá)":
        pendientes = telegram_notifier.filtrar_casos_por_filtro(pendientes, "regionales")

    if pendientes is None or pendientes.empty:
        st.info(f"ℹ️ No hay casos pendientes de notificar para la zona: **{filtro_region_notif}**.")
        return

    resumen = avisos.resumen_pendientes(pendientes)
    grupos = avisos.agrupar_por_tecnico(pendientes)
    menciones = avisos.cargar_menciones()
    if not menciones:
        st.caption(
            "ℹ️ No hay `menciones.json`: los avisos usan el nombre del técnico sin "
            "@usuario ni teléfono. Créelo a partir de `menciones.example.json`."
        )

    # --- i) KPIs de gestion de notificaciones -----------------------------
    historial_ok = historial is not None
    ultimas = historial.ultima_notificacion_por_tecnico() if historial_ok else {}
    hoy = historial.notificaciones_de_hoy_por_tecnico() if historial_ok else {}

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("👷 Técnicos pendientes", len(resumen))
    k2.metric("📋 Casos pendientes", int(len(pendientes)))
    k3.metric("✅ Técnicos notificados hoy", len(hoy))
    k4.metric("🔔 Notificaciones de hoy", int(sum(hoy.values())) if hoy else 0)

    if not historial_ok:
        st.warning(
            "No se pudo abrir el historial de notificaciones: las columnas de última "
            "notificación y veces hoy no están disponibles, y los botones de registro "
            "no podrán guardar."
        )

    tab_envio, tab_consolidado, tab_resumen, tab_historial = st.tabs(
        [
            "✉️ Aviso al técnico",
            "📢 Reporte Consolidado (TODOS)",
            "📊 Pendientes por técnico",
            "🕒 Última notificación por técnico",
        ]
    )

    # ==================================================================
    # Pestana 1: generar / copiar / enviar el aviso de UN tecnico
    # ==================================================================
    with tab_envio:
        c1, c2 = st.columns([3, 2])
        with c1:
            tecnico = st.selectbox(
                "👷 Técnico a notificar",
                list(resumen["TECNICO"]),
                key="notif_pend_tecnico",
                help="Solo aparecen técnicos con casos pendientes de aviso.",
            )
        with c2:
            canal = st.radio(
                "📡 Canal de la mención",
                ["whatsapp", "telegram"],
                horizontal=True,
                key="notif_pend_canal",
                help=(
                    "whatsapp usa el teléfono de menciones.json; telegram usa el "
                    "@usuario. El texto del aviso es el mismo en ambos casos."
                ),
            )

        un_solo_mensaje = st.checkbox(
            "🧾 Un solo mensaje con la lista de casos",
            value=True,
            key="notif_pend_un_mensaje",
            help=(
                "Marcado: un único mensaje con todos los casos del técnico. "
                "Desmarcado: un aviso separado por cada caso."
            ),
        )

        df_tecnico = grupos.get(tecnico)
        # Texto PLANO: es el que se copia y se pega en WhatsApp (con negrillas
        # de Telegram se verian las etiquetas <b>).
        texto = avisos.componer_aviso_tecnico(
            tecnico,
            df_tecnico,
            canal=canal,
            un_solo_mensaje=un_solo_mensaje,
            menciones=menciones,
        )
        # Version HTML para Telegram: negrillas de verdad y casos ordenados con
        # la prioridad de la torre (primero los proximos a vencer, luego los
        # vencidos).
        texto_telegram = avisos.componer_aviso_tecnico_telegram(
            tecnico,
            df_tecnico,
            un_solo_mensaje=un_solo_mensaje,
            menciones=menciones,
        )

        fila_tecnico = resumen[resumen["TECNICO"] == tecnico]
        if not fila_tecnico.empty:
            f = fila_tecnico.iloc[0]
            st.caption(
                f"📍 Región: **{f['REGION']}** · "
                f"📋 Pendientes: **{int(f['PENDIENTES'])}** · "
                f"🔴 {int(f['ROJOS'])} · 🟠 {int(f['NARANJAS'])} · "
                f"🟡 {int(f['AMARILLOS'])} · "
                f"🗓️ Próximo vencimiento: **{formatear_momento(f['PROXIMO_VENCIMIENTO'])}** · "
                f"🕒 Última notificación: **{f['ULTIMA_NOTIFICACION']}** · "
                f"🔔 Hoy: **{int(f['VECES_HOY'])}**"
            )

        # --- d) Boton de copiar nativo de st.code ---------------------
        st.code(texto, language=None)
        st.caption(
            "Pulse el icono de copiar (arriba a la derecha del recuadro) y pegue en WhatsApp."
        )

        # --- g) Envio por Telegram ------------------------------------
        st.markdown("##### 📨 Envío por Telegram")
        if telegram_notifier.telegram_configurado():
            destinos = telegram_notifier.destinos_configurados()
            st.caption(
                "Destinos configurados: "
                + " · ".join(
                    f"**{d['nombre']}** (`{d['chat_id']}`)" for d in destinos
                )
                + " · Se envía con **formato HTML (negrillas)** y los casos "
                "ordenados: primero los próximos a vencer, luego los vencidos."
            )
            with st.expander("👀 Vista previa del mensaje que recibirá Telegram"):
                st.markdown(
                    "Los datos resaltados en **negrita** son los que el técnico "
                    "debe ver primero: número de caso y tiempo restante/vencido."
                )
                st.code(texto_telegram, language="html")
            todos = st.checkbox(
                "Enviar a todos los destinos (si se desmarca, solo al primero)",
                value=len(destinos) > 1,
                key="notif_pend_telegram_todos",
            )

            if st.button(
                "📨 Enviar por Telegram",
                type="primary",
                key="btn_telegram_pend",
                disabled=df_tecnico is None or df_tecnico.empty,
            ):
                with st.spinner("Enviando el aviso a Telegram..."):
                    alcanzados = enviar_aviso_telegram(texto_telegram, todos=todos)
                if alcanzados:
                    registrados = _registrar_casos(
                        historial, df_tecnico, "telegram", "aviso al tecnico"
                    )
                    st.success(
                        f"✅ Aviso enviado a {alcanzados} destino(s) de Telegram. "
                        f"Casos registrados en el historial: {registrados}."
                    )
                    st.rerun()
                else:
                    st.error(
                        "❌ Telegram no aceptó el mensaje. Revise el token, el chat_id "
                        "y la conexión a internet (detalle en la consola)."
                    )
        else:
            st.button(
                "📨 Enviar por Telegram",
                key="btn_telegram_pend",
                disabled=True,
                help="Telegram no está configurado.",
            )
            st.info(
                "Telegram no está configurado. Defina `TELEGRAM_BOT_TOKEN` y "
                "`TELEGRAM_CHAT_ID`, o cree `config_telegram.json` (vea "
                "`docs/TUTORIAL_TELEGRAM.md`)."
            )

        # --- h) Registro manual: copiar y pegar en WhatsApp -----------
        st.markdown("##### ✅ Registro manual")
        m1, m2, m3 = st.columns([2, 2, 3])
        with m1:
            if st.button(
                "✅ Marcar como notificado",
                key="btn_marcar_tecnico",
                disabled=df_tecnico is None or df_tecnico.empty,
            ):
                registrados = _registrar_casos(
                    historial, df_tecnico, CANAL_WHATSAPP,
                    "aviso copiado y enviado manualmente por WhatsApp",
                    resultado=RESULTADO_GENERADO,
                )
                if registrados:
                    st.success(
                        f"✅ {registrados} caso(s) de **{tecnico}** marcados como "
                        "notificados (canal whatsapp)."
                    )
                    st.rerun()
                else:
                    st.error("❌ No se pudo registrar en el historial.")
        with m2:
            if st.button(
                "📣 Marcar TODOS los técnicos como notificados",
                key="btn_marcar_todos",
            ):
                registrados = _registrar_casos(
                    historial, pendientes, CANAL_WHATSAPP,
                    "aviso copiado y enviado manualmente por WhatsApp",
                    resultado=RESULTADO_GENERADO,
                )
                if registrados:
                    st.success(
                        f"✅ {registrados} caso(s) de {len(grupos)} técnico(s) marcados "
                        "como notificados (canal whatsapp)."
                    )
                    st.rerun()
                else:
                    st.error("❌ No se pudo registrar en el historial.")
        with m3:
            st.caption(
                "Use estos botones cuando ya envió el aviso por WhatsApp: registran el "
                "aviso en `historial.db` sin enviar nada, para que la columna "
                "*última notificación* y el contador de hoy queden al día."
            )

    # ==================================================================
    # Pestana Consolidada: informe global de TODOS los casos pendientes
    # ==================================================================
    with tab_consolidado:
        st.markdown("##### 📢 Reporte Consolidado Global de Todos los Casos")
        st.caption(
            "Consolida TODOS los casos pendientes en la Torre de Control SLA ordenados por urgencia. "
            "En Telegram incluye números de caso en `<code>` para copiado táctil instantáneo."
        )

        texto_cons_wa = avisos.componer_aviso_consolidado_whatsapp(pendientes, menciones)
        texto_cons_tg = avisos.componer_aviso_consolidado_telegram(pendientes, menciones)

        st.markdown("**📋 Texto para copiar a WhatsApp:**")
        st.code(texto_cons_wa, language=None)
        st.caption("Pulse el icono de copiar arriba a la derecha del recuadro.")

        st.markdown("##### 🚀 Envío Consolidado por Telegram")
        if telegram_notifier.telegram_configurado():
            with st.expander("👀 Vista previa formato Telegram (con <code> táctil de 1 toque)"):
                st.code(texto_cons_tg, language="html")

            c_btn1, c_btn2 = st.columns([3, 2])
            with c_btn1:
                if st.button(
                    "🚀 Enviar Reporte Consolidado a Telegram",
                    type="primary",
                    key="btn_tg_consolidado",
                    disabled=pendientes is None or pendientes.empty,
                ):
                    with st.spinner("Enviando reporte consolidado a Telegram..."):
                        alcanzados = enviar_aviso_telegram(texto_cons_tg, todos=True)
                    if alcanzados:
                        registrados = _registrar_casos(
                            historial, pendientes, "telegram", "reporte consolidado de todos los casos"
                        )
                        st.success(
                            f"✅ Reporte consolidado enviado a Telegram ({alcanzados} destino(s)). "
                            f"Se registraron {registrados} casos en el historial."
                        )
                        st.rerun()
                    else:
                        st.error("❌ Telegram no aceptó el mensaje. Revise credenciales y conexión.")
            with c_btn2:
                if st.button(
                    "✅ Marcar todos como notificados (WhatsApp)",
                    key="btn_wa_consolidado_marcar",
                    disabled=pendientes is None or pendientes.empty,
                ):
                    registrados = _registrar_casos(
                        historial, pendientes, CANAL_WHATSAPP,
                        "reporte consolidado copiado y enviado por WhatsApp",
                        resultado=RESULTADO_GENERADO,
                    )
                    if registrados:
                        st.success(f"✅ {registrados} casos marcados como notificados.")
                        st.rerun()
        else:
            st.info("Telegram no está configurado.")

    # ==================================================================
    # Pestana 2: tabla de pendientes (a)
    # ==================================================================
    with tab_resumen:
        st.markdown("##### 📊 Técnicos con casos pendientes de notificar")
        st.dataframe(
            tabla_pendientes_visible(resumen),
            width="stretch",
            hide_index=True,
            column_config={
                "PROXIMO_VENCIMIENTO": st.column_config.TextColumn(
                    "PROXIMO_VENCIMIENTO", help="Formato dd/mm/yyyy HH:MM"
                ),
                "VECES_HOY": st.column_config.NumberColumn(
                    "VECES_HOY", help="Notificaciones registradas hoy para el técnico"
                ),
            },
            height=min(600, 40 + 35 * len(resumen)),
        )
        st.download_button(
            "⬇️ Descargar pendientes por técnico (CSV)",
            data=resumen.to_csv(index=False, na_rep="").encode("utf-8-sig"),
            file_name=f"pendientes_por_tecnico_{ahora_colombia():%Y%m%d_%H%M}.csv",
            mime="text/csv",
            key="descarga_pendientes_tecnico",
        )

    # ==================================================================
    # Pestana 3: ultima notificacion y veces hoy (j)
    # ==================================================================
    with tab_historial:
        st.markdown("##### 🕒 Última notificación por técnico")
        if not historial_ok:
            st.info("El historial no está disponible.")
        elif not ultimas and not hoy:
            st.info(
                "Todavía no hay notificaciones registradas en el historial. "
                "Se llena al usar los botones de esta sección o el notificador."
            )
        else:
            historial_tecnicos = sorted(set(ultimas) | set(hoy))
            tabla_ultimas = pd.DataFrame(
                [
                    {
                        "TECNICO": t,
                        "ULTIMA_NOTIFICACION": formatear_momento(ultimas.get(t)),
                        "VECES_HOY": int(hoy.get(t, 0)),
                    }
                    for t in historial_tecnicos
                ]
            ).sort_values(
                ["VECES_HOY", "TECNICO"], ascending=[False, True]
            ).reset_index(drop=True)
            st.dataframe(
                tabla_ultimas,
                width="stretch",
                hide_index=True,
                column_config={
                    "VECES_HOY": st.column_config.NumberColumn(
                        "VECES_HOY", help="Notificaciones registradas hoy"
                    )
                },
                height=min(500, 40 + 35 * len(tabla_ultimas)),
            )
            st.caption(
                f"{len(tabla_ultimas)} técnico(s) con notificaciones registradas · "
                f"total de hoy: {sum(hoy.values()) if hoy else 0}"
            )


# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------

def estilizar(df_visible: pd.DataFrame) -> pd.DataFrame:
    """
    Mapa de estilos CSS (mismo indice/columnas que df_visible).

    Se aplica con .apply(estilizar, axis=None). IMPORTANTE: con axis=None pandas
    invoca la funcion con el DataFrame y un argumento 'axis' extra, por lo que la
    firma debe aceptar argumentos posicionales adicionales.
    """

    def color_estado(valor: str) -> str:
        fondo = COLORES_FONDO.get(valor, "")
        texto = COLORES_TEXTO.get(valor, "")
        if not fondo:
            return ""
        return f"background-color: {fondo}; color: {texto}; font-weight: 600;"

    estilos = pd.DataFrame("", index=df_visible.index, columns=df_visible.columns)
    if "Estado" in df_visible.columns:
        estilos["Estado"] = df_visible["Estado"].map(
            lambda v: color_estado(v) if isinstance(v, str) else ""
        )
    return estilos


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------

def panel_duplicados(df_vista: pd.DataFrame) -> None:
    """
    Aviso y detalle de los casos con el N° DE CASO repetido en la plantilla.

    Existe por un caso real: IM3238158 aparecia como CERRADO TARDE y a la vez el
    sistema lo alertaba como abierto y vencido, porque la plantilla trae dos filas
    del mismo caso (una sin cerrar y otra cerrada). Aqui se ven las dos copias y
    se puede saltar a la tabla filtrada.
    """
    if df_vista is None or df_vista.empty or COL_ES_DUPLICADO not in df_vista.columns:
        return

    repetidos = resumen_duplicados(df_vista)
    if repetidos.empty:
        return

    st.error(
        f"🔁 **{len(repetidos)} caso(s) repetidos en la plantilla** "
        f"({int(repetidos['VECES'].sum())} filas en total). Por eso un caso podía "
        "aparecer **cerrado** y a la vez **alertando como abierto**: cada copia se "
        "clasificaba por su cuenta. Se marcan **DUPLICADO** y **no se cuentan** en el "
        "semáforo ni en las métricas (la app no borra ni elige ninguna fila)."
    )
    with st.expander(f"🔎 Ver el detalle de los {len(repetidos)} caso(s) repetidos"):
        st.dataframe(repetidos, width="stretch", hide_index=True)
        st.caption(
            "**FILAS_EXCEL** = fila que ocupa cada copia en la hoja `PLANTILLA`. "
            "**ESTADO_POR_FILA** = lo que le habría correspondido a cada copia antes "
            "de marcarla como DUPLICADO."
        )
        st.caption(
            "Para corregirlo en el Excel: deje **una sola fila** por N° DE CASO "
            "(la que tenga la fecha de resolución) y borre la otra."
        )

    if st.button(
        "🔁 Ver solo los casos repetidos en la tabla",
        key="btn_ver_duplicados",
        width="content",
    ):
        st.session_state["exp_estados"] = [DUPLICADO]
        st.rerun()


def auditar_integridad(df_crudo, df_completo, nombre_archivo: str = "") -> dict:
    """
    Comprueba que NINGUNA fila diligenciada se pierda entre el Excel y la app.

    Es la verificacion que pide la torre: si la fuente trae 15 filas
    diligenciadas, la app debe representar las 15.

    Returns:
        dict con conteos, la tabla de verificaciones y la lista de faltantes.
    """
    if df_crudo is None:
        return {"disponible": False}

    n_crudo = len(df_crudo)
    n_procesado = len(df_completo) if df_completo is not None else 0

    def _diligenciadas(df: pd.DataFrame) -> int:
        """Filas con al menos caso o vencimiento: las que significan algo."""
        if df is None or df.empty:
            return 0
        cols = [c for c in (COL_CASO, COL_VENCIMIENTO) if c in df.columns]
        if not cols:
            return len(df)
        return int(df[cols].notna().any(axis=1).sum())

    dil_crudo = _diligenciadas(df_crudo)
    dil_proc = _diligenciadas(df_completo)

    # Casos con el mismo N° DE CASO en varias filas: no se pierden, se marcan.
    repetidos = resumen_duplicados(df_completo)

    # --- Casos del Excel que no llegaron a la app -------------------------
    faltantes: list[str] = []
    if COL_CASO in df_crudo.columns and df_completo is not None and COL_CASO in df_completo.columns:
        en_crudo = {
            str(v).strip() for v in df_crudo[COL_CASO].dropna() if str(v).strip()
        }
        en_app = {
            str(v).strip() for v in df_completo[COL_CASO].dropna() if str(v).strip()
        }
        faltantes = sorted(en_crudo - en_app)

    # --- Completitud por columna -----------------------------------------
    revisiones = [
        ("Filas en la hoja Excel", n_crudo, "Todo lo que trae el archivo"),
        ("Filas procesadas por la app", n_procesado, "Resultado del calculo de SLA"),
        ("Filas diligenciadas (Excel)", dil_crudo, "Con caso o vencimiento"),
        ("Filas diligenciadas (app)", dil_proc, "Representadas en pantalla"),
        ("Casos repetidos en la plantilla", len(repetidos),
         "Mismo N° DE CASO en varias filas (se marcan DUPLICADO)"),
    ]

    columnas_revisadas = [
        (COL_CASO, "N° de caso"),
        ("TECNICO COLSOF ASIGNADO INICIALMENTE", "Técnico asignado"),
        (COL_CIUDAD, "Oficina"),
        (COL_REGIONAL, "Regional"),
        (COL_DEPARTAMENTO, "Departamento"),
        (COL_VENCIMIENTO, "Vencimiento"),
        ("FECHA/ HORA QUE SE ATENDIO Y SE DIO POR RESUELTO", "Fecha de resolución"),
    ]
    completitud = []
    for col, etiqueta in columnas_revisadas:
        if col in df_crudo.columns:
            llenas = int(df_crudo[col].notna().sum())
            completitud.append({
                "Campo": etiqueta,
                "Diligenciadas": llenas,
                "Vacías": n_crudo - llenas,
                "% completo": round(llenas / n_crudo * 100, 1) if n_crudo else 0.0,
            })

    cuadra = (n_crudo == n_procesado) and (dil_crudo == dil_proc) and not faltantes

    return {
        "disponible": True,
        "archivo": nombre_archivo,
        "n_crudo": n_crudo,
        "n_procesado": n_procesado,
        "dil_crudo": dil_crudo,
        "dil_procesado": dil_proc,
        "revisiones": revisiones,
        "completitud": pd.DataFrame(completitud),
        "faltantes": faltantes,
        "duplicados": repetidos,
        "cuadra": cuadra,
    }


def render_integridad(df_crudo, df_completo, nombre_archivo: str = "") -> None:
    """Dibuja el panel "Integridad de datos"."""
    st.subheader("🔍 Integridad de datos")
    st.caption(
        "Verifica que todas las filas diligenciadas del Excel llegaron al tablero. "
        "Si algo no cuadra, aparece aquí y en el Explorador."
    )

    info = auditar_integridad(df_crudo, df_completo, nombre_archivo)

    if not info.get("disponible"):
        st.warning(
            "No se pudo auditar: esta versión de la carga no expone las filas "
            "originales del Excel."
        )
        return

    # --- Veredicto ---------------------------------------------------------
    if info["cuadra"]:
        st.success(
            f"✅ **Cuadra**: {info['dil_procesado']} de {info['dil_crudo']} filas "
            f"diligenciadas están representadas. Nada se perdió al procesar."
        )
    else:
        st.error(
            f"❌ **No cuadra**: el Excel trae {info['dil_crudo']} filas diligenciadas "
            f"y la app representa {info['dil_procesado']}."
        )

    # --- Conteos -----------------------------------------------------------
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📄 Filas en el Excel", info["n_crudo"])
    c2.metric("⚙️ Procesadas por la app", info["n_procesado"],
              delta=info["n_procesado"] - info["n_crudo"], delta_color="off")
    c3.metric("✍️ Diligenciadas (Excel)", info["dil_crudo"])
    c4.metric("👁️ Diligenciadas (app)", info["dil_procesado"],
              delta=info["dil_procesado"] - info["dil_crudo"], delta_color="off")

    # --- Tabla de verificaciones ------------------------------------------
    st.markdown("##### Verificaciones")
    st.dataframe(
        pd.DataFrame(info["revisiones"], columns=["Verificación", "Cantidad", "Qué mide"]),
        width="stretch",
        hide_index=True,
    )

    # --- Casos faltantes ---------------------------------------------------
    if info["faltantes"]:
        st.error(
            f"⚠️ **{len(info['faltantes'])} caso(s) del Excel no aparecen en la app.** "
            "Revise si tienen el número de caso bien diligenciado:"
        )
        st.code("\n".join(info["faltantes"][:50]))
        if len(info["faltantes"]) > 50:
            st.caption(f"…y {len(info['faltantes']) - 50} más.")
    elif info["disponible"]:
        st.info("✅ Ningún caso del Excel quedó por fuera: todos aparecen en el tablero.")

    # --- Casos repetidos (mismo N° DE CASO en varias filas) ----------------
    repetidos = info.get("duplicados")
    if repetidos is not None and not repetidos.empty:
        st.error(
            f"🔁 **{len(repetidos)} caso(s) repetidos en la plantilla** "
            f"({int(repetidos['VECES'].sum())} filas). La app NO borra ni elige una "
            "copia: las marca **DUPLICADO** y las deja **fuera del semáforo y de las "
            "métricas** hasta que se unifiquen en el Excel."
        )
        st.dataframe(repetidos, width="stretch", hide_index=True)
        st.caption(
            "**FILAS_EXCEL** = fila de la hoja `PLANTILLA` que ocupa cada copia. "
            "**ESTADO_POR_FILA** = lo que le habría correspondido a cada copia antes "
            "de marcarla como DUPLICADO."
        )

    # --- Completitud por columna ------------------------------------------
    if not info["completitud"].empty:
        st.markdown("##### Completitud por campo")
        st.caption(
            "Cuántas filas traen cada dato diligenciado. Los campos vacíos explican "
            "los casos 'sin vencimiento' o 'sin técnico'."
        )
        st.dataframe(
            info["completitud"].style.format({"% completo": "{:.1f}%"}),
            width="stretch",
            hide_index=True,
        )

    # --- Antigüedad del archivo -------------------------------------------
    if info.get("archivo"):
        st.markdown("##### Origen de los datos")
        st.caption(f"Archivo: `{info['archivo']}`")

    if st.button("🔄 Volver a auditar", key="reauditar", width="content"):
        st.cache_data.clear()
        st.rerun()


# =========================================================================
# PLAN DE TRABAJO — tablero gerencial de casos en curso
# =========================================================================

def _fmt_mes(periodo: str) -> str:
    """'2026-08' -> 'Agosto'."""
    try:
        return ETIQUETA_MES_TABLERO[int(str(periodo)[-2:])]
    except (ValueError, KeyError, IndexError):
        return str(periodo)


def _grafico_barras(datos: list[dict], campo_x: str, campo_y: str,
                    titulo: str, rotulo_x: str, rotulo_y: str, color: str):
    """Barras verticales simples y limpias, sin ruido visual."""
    if not datos:
        return None
    dfg = pd.DataFrame(datos)
    return (
        alt.Chart(dfg)
        .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4, color=color)
        .encode(
            x=alt.X(f"{campo_x}:N", title=rotulo_x, sort=None,
                    axis=alt.Axis(labelAngle=0, labelFontSize=12, titleFontSize=12)),
            y=alt.Y(f"{campo_y}:Q", title=rotulo_y,
                    axis=alt.Axis(labelFontSize=12, titleFontSize=12)),
            tooltip=[alt.Tooltip(f"{campo_x}:N", title=rotulo_x),
                     alt.Tooltip(f"{campo_y}:Q", title=rotulo_y)],
        )
        .properties(title=alt.TitleParams(titulo, fontSize=15, anchor="start"),
                    height=280)
    )


def _grafico_horizontal(datos: list[dict], titulo: str, color: str):
    """Barras horizontales para rankings (se leen mejor los nombres)."""
    if not datos:
        return None
    dfg = pd.DataFrame(datos)
    return (
        alt.Chart(dfg)
        .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4, color=color)
        .encode(
            y=alt.Y("etiqueta:N", title=None, sort="-x",
                    axis=alt.Axis(labelFontSize=12)),
            x=alt.X("casos:Q", title="Casos",
                    axis=alt.Axis(labelFontSize=12, titleFontSize=12)),
            tooltip=[alt.Tooltip("etiqueta:N", title="Técnico"),
                     alt.Tooltip("casos:Q", title="Casos")],
        )
        .properties(title=alt.TitleParams(titulo, fontSize=15, anchor="start"),
                    height=max(220, 26 * len(datos) + 60))
    )


def _grafico_semanal(datos: list[dict], titulo: str):
    """Evolución por semana: línea con puntos."""
    if not datos:
        return None
    dfg = pd.DataFrame(datos)
    base = alt.Chart(dfg).encode(
        x=alt.X("etiqueta:N", title="Semana", sort=None,
                axis=alt.Axis(labelAngle=0, labelFontSize=12, titleFontSize=12)),
        y=alt.Y("casos:Q", title="Casos",
                axis=alt.Axis(labelFontSize=12, titleFontSize=12)),
        tooltip=[alt.Tooltip("etiqueta:N", title="Semana"),
                 alt.Tooltip("casos:Q", title="Casos")],
    )
    linea = base.mark_line(color="#1F4E78", strokeWidth=3, point=alt.OverlayMarkDef(
        color="#1F4E78", size=90, filled=True))
    return linea.properties(
        title=alt.TitleParams(titulo, fontSize=15, anchor="start"), height=280
    )


# Estados que cuentan como "en curso" para el tablero gerencial. Se declaran
# aqui, y no se leen de plan.py, a proposito: este modulo puede quedar cacheado
# en el servidor y desincronizarse de dashboard.py. Con la constante local, el
# tablero sigue funcionando aunque plan.py venga de una version anterior.
ESTADOS_EN_CURSO_TABLERO = ("EN CURSO", "TRABAJO EN CURSO")

# Nombres de la interfaz para cada estado.
ETIQUETA_ESTADO_TABLERO = {
    "EN CURSO": "En curso",
    "TRABAJO EN CURSO": "Trabajo en curso",
    "SUSPENDIDO": "Suspendido",
    "READY": "Ready",
    "WORK IN PROGRESS": "Work in progress",
    "PREPARADO": "Preparado",
    "ASIGNADO": "Asignado",
    "PENDIENTE": "Pendiente",
    "PENDING": "Pending",
    "CATEGORIZADO": "Categorizado",
}

# Nombres de mes para las etiquetas de los graficos.
ETIQUETA_MES_TABLERO = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre",
    12: "Diciembre",
}

# Lo minimo que debe exponer plan.py para que la pestana funcione.
# Lo minimo que debe exponer plan.py para que la pestana funcione.
# Son las funciones y constantes que esta pestana usa de verdad.
PLAN_REQUERIDO = (
    "vista_vencidos", "validar_fechas_vencidos", "reconocer_tipo",
    "normalizar", "ErrorPlan",
    "COL_CASO", "COL_UBICACION", "COL_FECHA_CREACION", "COL_TECNICO",
    "COL_JUSTIFICACION", "COL_CULPA", "COL_CATEGORIA",
    "COL_MES_CREACION", "COL_DIAS_ABIERTO",
)


def _faltantes_de_plan() -> list:
    """Nombres que dashboard.py necesita y plan.py no expone."""
    return [n for n in PLAN_REQUERIDO if not hasattr(plan, n)]


def _aviso_plan_desactualizado(faltantes) -> None:
    """Explica el problema en lugar de dejar caer un AttributeError."""
    st.error(
        "\u274c **El servidor tiene una versión antigua del módulo `plan.py`.** "
        "Es un problema de caché del servidor, no del archivo que subiste.",
        icon="\U0001f9e9",
    )
    st.markdown(
        "Faltan estas funciones en el módulo cargado:\n\n"
        + "\n".join("- `plan." + n + "`" for n in faltantes)
        + "\n\n**Solución:** en Streamlit Cloud, entra a *Manage app* "
        "\u2192 menú \u22ee \u2192 **Reboot**. Si sigue igual, usa "
        "*Manage app* \u2192 \u22ee \u2192 **Delete app** y vuelve a "
        "desplegar, que limpia la caché por completo."
    )


def _etiqueta_estado_tablero(valor) -> str:
    """Nombre presentable del estado, sin depender de plan.py."""
    clave = normalizar_generico(valor)
    return ETIQUETA_ESTADO_TABLERO.get(clave, str(valor).strip() or "Sin estado")


def normalizar_generico(valor) -> str:
    """Normalizacion de texto local, por si plan.py viene desactualizado."""
    if hasattr(plan, "normalizar"):
        return plan.normalizar(valor)
    texto_valor = str(valor if valor is not None else "").strip().upper()
    return " ".join(texto_valor.split())



# =========================================================================
# PLAN DE TRABAJO — hoja Casos_Ven
# =========================================================================
#
# La pestaña muestra UNICAMENTE el contenido de la hoja Casos_Ven del archivo
# del Plan de Trabajo: los casos vencidos del mes, con su detalle completo.
#
# No se usan las hojas diarias (23_Septiembre, 22_Septiembre...). Aquellas
# respondian a otra pregunta ("que habia abierto ese dia") y confundian la
# lectura: el usuario espera ver su hoja, con sus 145 filas.

def _dibujar_conciliacion(df_completo, plan_bytes) -> None:
    """
    Dibuja el panel de conciliacion si el modulo vista lo trae.

    El servidor de Streamlit Cloud puede quedar con una version vieja de un modulo
    en cache (ya paso con plan.py). Si eso ocurre, se avisa con la solucion en vez
    de tumbar todo el tablero con un ImportError: la conciliacion es UNA seccion de
    UNA pestana, no el tablero entero.
    """
    dibujar = getattr(vista_panel, "render_conciliacion", None)
    if dibujar is None:
        st.warning(
            "⚠️ **El servidor tiene una versión antigua de `vista.py`**: falta el "
            "panel de conciliación. En Streamlit Cloud, *Manage app* → menú ⋮ → "
            "**Reboot**; si sigue igual, *Delete app* y vuelve a desplegar.",
            icon="🧩",
        )
        return
    dibujar(df_completo, plan_bytes)


def render_plan_trabajo(df_completo: pd.DataFrame | None = None) -> None:
    """
    Pestaña del Plan de Trabajo: todo el contenido de la hoja Casos_Ven.

    Muestra los totales (por mes, tecnico, culpa y categoria) y el detalle caso
    a caso, con filtros y descargas.
    """
    faltantes = _faltantes_de_plan()
    if faltantes:
        _aviso_plan_desactualizado(faltantes)
        return

    st.markdown("#### 🗂️ Plan de Trabajo — hoja Casos_Ven")
    st.caption(
        "Casos vencidos del mes, tal como están en la hoja `Casos_Ven`."
    )

    with st.sidebar:
        st.divider()
        st.subheader("🗂️ Cargar Plan de Trabajo")
        st.caption(
            "El archivo mensual del plan (el que trae la hoja `Casos_Ven`). Se "
            "procesa en memoria; no se guarda en el servidor."
        )
        subido = st.file_uploader(
            "Plan de Trabajo (.xlsx)", type=["xlsx"], key="uploader_plan"
        )

    if subido is not None:
        contenido = subido.getvalue()
        nombre = subido.name
    elif st.session_state.get("plan_bytes"):
        contenido = st.session_state["plan_bytes"]
        nombre = st.session_state.get("plan_nombre", "Plan de Trabajo.xlsx")
        st.success(f"📄 Usando el Plan de Trabajo cargado arriba: `{nombre}`", icon="✅")
    else:
        st.info(
            "**Suba el archivo del Plan de Trabajo** desde la barra lateral "
            "(el que trae la hoja `Casos_Ven`).",
            icon="📤",
        )
        st.markdown(
            """
            **Qué verá aquí**

            | | |
            |---|---|
            | 🔢 **Casos** | Las filas de la hoja `Casos_Ven` |
            | 📅 **Por mes** | Reparto por mes de creación |
            | 👷 **Por técnico** | Cuántos lleva cada uno |
            | 🚧 **Por culpa** | Técnico, logístico, aliado, banco… |
            | 📋 **Detalle** | Caso, ubicación, fecha, justificación |
            """
        )
        return

    with st.spinner("Leyendo la hoja Casos_Ven..."):
        try:
            vista = plan.vista_vencidos(contenido)
        except plan.ErrorPlan as exc:
            st.error("❌ No se pudo leer la hoja Casos_Ven.")
            st.code(str(exc))
            return
        except Exception as exc:
            st.error("❌ Error inesperado al leer el archivo.")
            st.code(f"{type(exc).__name__}: {exc}")
            return

    casos = vista["casos"]
    meta = vista["meta"]
    conc = vista["conciliacion"]

    if casos.empty:
        st.warning("La hoja Casos_Ven no tiene casos utilizables.")
        return

    # --- De donde sale el total ------------------------------------------
    if conc["descartados"]:
        st.info(
            f"La hoja trae **{conc['filas_leidas']} filas** y quedan "
            f"**{vista['total']} casos únicos**: "
            f"{conc['descartados']} se descartaron por estar repetidos o con los "
            "datos corridos.",
            icon="🔎",
        )

    # --- KPIs -------------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Filas en la hoja", conc["filas_leidas"])
    k2.metric("Casos únicos", vista["total"])
    k3.metric("Técnicos", meta["tecnicos"])
    k4.metric("Meses", meta["meses"])

    st.divider()

    # --- Graficos ---------------------------------------------------------
    g1, g2 = st.columns(2)
    with g1:
        datos_mes = [
            {"etiqueta": m["etiqueta"].split()[0], "casos": m["casos"]}
            for m in vista["por_mes"]
        ]
        grafico = _grafico_barras(
            datos_mes, "etiqueta", "casos",
            "Casos por mes de creación", "Mes", "Casos", "#1F4E78",
        )
        if grafico is not None:
            st.altair_chart(grafico, width="stretch")
    with g2:
        datos_culpa = [
            {"etiqueta": c["culpa"], "casos": c["casos"]}
            for c in reversed(vista["por_culpa"])
        ]
        grafico = _grafico_horizontal(datos_culpa, "Casos por culpa", "#C2410C")
        if grafico is not None:
            st.altair_chart(grafico, width="stretch")

    st.markdown("##### 👷 Casos por técnico")
    datos_tec = [
        {"etiqueta": t["tecnico"], "casos": t["casos"]}
        for t in reversed(vista["por_tecnico"])
    ]
    grafico = _grafico_horizontal(datos_tec, "Todos los técnicos", "#2E75B6")
    if grafico is not None:
        st.altair_chart(grafico, width="stretch")

    st.divider()

    # --- Validacion de fechas --------------------------------------------
    _panel_validacion_fechas(contenido)

    # --- Conciliacion con la plantilla SLA --------------------------------
    # Cruza los dos archivos y muestra donde se contradicen (conciliacion.py).
    st.divider()
    _dibujar_conciliacion(df_completo, contenido)

    # --- Detalle ----------------------------------------------------------
    st.markdown("##### 📋 Detalle caso a caso")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        meses = ["(todos)"] + [m["etiqueta"] for m in vista["por_mes"]]
        mes_sel = st.selectbox("Mes de creación", meses, key="ven_mes")
    with c2:
        tecnicos = ["(todos)"] + [t["tecnico"] for t in vista["por_tecnico"] if t["tecnico"]]
        tec_sel = st.selectbox("Técnico", tecnicos, key="ven_tec")
    with c3:
        culpas = ["(todas)"] + [c["culpa"] for c in vista["por_culpa"] if c["culpa"]]
        culpa_sel = st.selectbox("Culpa", culpas, key="ven_culpa")
    with c4:
        categorias = ["(todas)"] + [
            c["categoria"] for c in vista["por_categoria"] if c["categoria"]
        ]
        cat_sel = st.selectbox("Categoría", categorias, key="ven_cat")

    filtrado = casos.copy()
    if mes_sel != "(todos)":
        etiqueta_a_mes = {m["etiqueta"]: m["mes"] for m in vista["por_mes"]}
        filtrado = filtrado[filtrado[plan.COL_MES_CREACION] == etiqueta_a_mes[mes_sel]]
    if tec_sel != "(todos)":
        filtrado = filtrado[filtrado[plan.COL_TECNICO] == tec_sel]
    if culpa_sel != "(todas)":
        filtrado = filtrado[filtrado[plan.COL_CULPA] == culpa_sel]
    if cat_sel != "(todas)":
        filtrado = filtrado[filtrado[plan.COL_CATEGORIA] == cat_sel]

    st.caption(f"Mostrando **{len(filtrado)}** de {vista['total']} casos.")

    columnas = {
        plan.COL_CASO: "Caso",
        plan.COL_UBICACION: "Ubicación",
        plan.COL_FECHA_CREACION: "Fecha creación",
        plan.COL_TECNICO: "Técnico",
        plan.COL_CULPA: "Culpa",
        plan.COL_CATEGORIA: "Categoría",
        plan.COL_JUSTIFICACION: "Justificación del vencimiento",
        plan.COL_DIAS_ABIERTO: "Días abierto",
    }
    presentes = [c for c in columnas if c in filtrado.columns]
    detalle = filtrado[presentes].rename(columns=columnas)
    if "Fecha creación" in detalle.columns:
        detalle["Fecha creación"] = pd.to_datetime(
            detalle["Fecha creación"], errors="coerce"
        ).dt.strftime("%d/%m/%Y")
    if "Días abierto" in detalle.columns:
        detalle["Días abierto"] = detalle["Días abierto"].round(0).astype("Int64")

    st.dataframe(detalle, width="stretch", hide_index=True, height=460)

    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "⬇️ Descargar los casos filtrados (CSV)",
            data=detalle.to_csv(index=False, na_rep="").encode("utf-8-sig"),
            file_name="casos_ven.csv",
            mime="text/csv",
            key="descarga_vencidos",
        )
    with d2:
        st.download_button(
            "⬇️ Descargar los totales (CSV)",
            data=_csv_totales(vista),
            file_name="casos_ven_totales.csv",
            mime="text/csv",
            key="descarga_vencidos_totales",
        )


def _csv_totales(vista: dict) -> bytes:
    """Los cuatro totales en un solo CSV, cada uno con su encabezado."""
    partes = []
    for titulo, clave, etiqueta in (
        ("TOTAL POR MES", "por_mes", "etiqueta"),
        ("TOTAL POR TECNICO", "por_tecnico", "tecnico"),
        ("TOTAL POR CULPA", "por_culpa", "culpa"),
        ("TOTAL POR CATEGORIA", "por_categoria", "categoria"),
    ):
        datos = vista.get(clave) or []
        if not datos:
            continue
        partes.append(titulo)
        partes.append(pd.DataFrame(datos).rename(
            columns={etiqueta: "Concepto", "casos": "Casos"}
        )[["Concepto", "Casos"]].to_csv(index=False).strip())
        partes.append("")
    return "\n".join(partes).encode("utf-8-sig")


def _panel_validacion_fechas(contenido: bytes) -> None:
    """
    Contraste de las fechas con una fuente independiente.

    La fecha de creacion se compara con la fecha de apertura que el banco
    registra en las hojas diarias: si coinciden, la fecha esta confirmada.
    """
    with st.expander("✅ ¿Están bien las fechas?"):
        st.caption(
            "Cada fecha se contrasta con la fecha de apertura que el banco "
            "registra en las hojas diarias del mismo archivo. Son dos fuentes "
            "independientes: si coinciden, la fecha está confirmada."
        )
        try:
            val = plan.validar_fechas_vencidos(contenido)
        except Exception as exc:
            st.warning("No se pudo hacer el contraste con las hojas diarias.")
            st.code(f"{type(exc).__name__}: {exc}")
            return

        if val["comparables"] == 0:
            st.info("No hay casos comparables con las hojas diarias.")
        elif val["difieren"] == 0:
            st.success(
                f"**Las {val['comparables']} fechas comparables coinciden "
                "exactamente** con la fecha de apertura de las hojas diarias. "
                "No hay nada que corregir.",
                icon="✅",
            )
        else:
            st.warning(
                f"Coinciden **{val['coinciden']} de {val['comparables']}** "
                f"({val['pct']}%). Hay **{val['difieren']}** que no cuadran:",
                icon="⚠️",
            )
            st.dataframe(
                pd.DataFrame(val["ejemplos"]).rename(columns={
                    "caso": "Caso", "vencidos": "Casos_Ven",
                    "diario": "Hoja diaria",
                }),
                width="stretch", hide_index=True,
            )

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Por mes — fuente diaria**")
            st.dataframe(
                pd.DataFrame(val["por_mes_diario"])[["etiqueta", "casos"]]
                .rename(columns={"etiqueta": "Mes", "casos": "Casos"}),
                width="stretch", hide_index=True,
            )
        with c2:
            st.markdown("**Por mes — Casos_Ven**")
            st.dataframe(
                pd.DataFrame(val["por_mes_vencidos"])[["etiqueta", "casos"]]
                .rename(columns={"etiqueta": "Mes", "casos": "Casos"}),
                width="stretch", hide_index=True,
            )


def main() -> None:
    st.set_page_config(
        page_title="Torre de Control SLA - Colsof",
        page_icon="🛰️",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # --- Control de acceso ------------------------------------------------
    auth.puente_secretos()
    sesion = auth.exigir_login()

    # --- Inyección de Sistema de Diseño CSS -------------------------------
    estilos_css.inyectar_estilos_css()

    st.title("🛰️ Torre de Control SLA — Colsof / Banco Agrario")
    st.caption("Monitoreo de Acuerdos de Nivel de Servicio para soporte técnico en sitio.")

    # --- Barra lateral ----------------------------------------------------
    with st.sidebar:
        st.header("⚙️ Control")
        if st.button("🔄 Recalcular ahora", width="stretch"):
            st.cache_data.clear()
            st.rerun()

        autorrefresco = st.checkbox(
            f"Auto-actualizar cada {INTERVALO_AUTOREFRESCO_S} s", value=False
        )
        st.divider()
        st.subheader("🚦 Semáforo v2")
        for estado in TODOS_LOS_ESTADOS:
            st.markdown(
                f"<span style='background-color:{COLORES_FONDO[estado]};"
                f"color:{COLORES_TEXTO[estado]};padding:2px 6px;border-radius:4px;"
                f"font-weight:600'>{ICONO_ESTADO[estado]} {estado}</span>",
                unsafe_allow_html=True,
            )
            st.caption(AYUDA_ESTADO[estado])
        st.divider()
        st.caption(f"🛠️ **Versión del tablero:** `{VERSION_APP}`")
        st.caption(
            "Si después de una actualización sigue viendo una versión anterior, "
            "el servidor no ha tomado el último cambio: *Manage app* → ⋮ → **Reboot**."
        )

    # --- Localización del archivo o subida manual -------------------------
    momento = ahora_colombia()

    origen_subido = None
    with st.sidebar:
        st.divider()
        st.subheader("📤 Cargar archivo")
        st.caption(
            "Suba la **plantilla de seguimiento** o el **Plan de Trabajo**: el "
            "sistema reconoce cuál es por sus hojas y lo envía a la pestaña "
            "correcta. Se procesa en memoria; no se guarda en el servidor."
        )
        archivo_subido = st.file_uploader(
            "Plantilla o Plan de Trabajo (.xlsx)",
            type=["xlsx"],
            key="uploader_plantilla",
            help="Si no sube nada, se usa la plantilla que esté junto al sistema.",
        )
        if archivo_subido is not None:
            origen_subido = archivo_subido

    if origen_subido is not None:
        contenido = origen_subido.getvalue()
        import hashlib

        firma = int(hashlib.md5(contenido).hexdigest()[:8], 16)
        ruta = f"(subido) {origen_subido.name}"

        # Reconocimiento automatico: no se le exige al usuario saber cual subir.
        tipo = plan.reconocer_tipo(contenido)
        st.session_state["tipo_archivo"] = tipo["tipo"]
        st.session_state["detalle_archivo"] = tipo["detalle"]

        if tipo["tipo"] == plan.TIPO_PLAN:
            # Es el Plan de Trabajo: se enruta a su pestana y se sigue usando la
            # plantilla SLA que ya estuviera disponible, sin descartarla.
            st.session_state["plan_bytes"] = contenido
            st.session_state["plan_nombre"] = origen_subido.name
            origen_subido = None
        elif tipo["tipo"] == plan.TIPO_PLANTILLA:
            # Es la plantilla SLA. Se limpia un Plan anterior para no mezclar
            # dos archivos de meses distintos.
            st.session_state.pop("plan_bytes", None)
            st.session_state.pop("plan_nombre", None)

    hay_plantilla = False
    plantilla_error = None
    if origen_subido is not None:
        lector = lambda: cargar_desde_bytes(
            origen_subido.getvalue(), float(firma), momento.isoformat(), origen_subido.name
        )
        hay_plantilla = True
    else:
        try:
            ruta = localizar_excel(os.environ.get("TORRE_EXCEL"))
        except ErrorLecturaExcel as exc:
            ruta = "(ninguna)"
            lector = None
            plantilla_error = exc
        else:
            lector = lambda: cargar(ruta, firma_archivo(ruta), momento.isoformat())
            hay_plantilla = True

    # --- Lectura de Datos -------------------------------------------------
    df = pd.DataFrame()
    df_completo = pd.DataFrame()
    df_crudo = None
    nombre_archivo = "plantilla.xlsx"
    total_hoja = total_activos = 0
    avisos = []

    if hay_plantilla:
        try:
            with st.spinner("Leyendo la plantilla y calculando SLA..."):
                devuelto = lector()
                (df, df_completo, total_hoja, total_activos,
                 dias_ventana, avisos) = devuelto[:6]
                # Los dos ultimos son opcionales (compatibilidad si alguna version
                # de cargar() no los devuelve todavia)
                df_crudo = devuelto[6] if len(devuelto) > 6 else None
                nombre_archivo = devuelto[7] if len(devuelto) > 7 else "plantilla.xlsx"
        except ErrorLecturaExcel as exc:
            st.error("❌ No se pudo leer el archivo Excel.")
            st.warning(
                "Causa más frecuente: **el archivo está abierto en Excel por otro "
                "usuario**. Ciérrelo y pulse *Recalcular ahora*."
            )
            st.code(str(exc))
            hay_plantilla = False
        except Exception as exc:
            st.error("❌ No se pudo procesar el archivo.")
            tipo_info = plan.reconocer_tipo(
                origen_subido.getvalue() if origen_subido is not None else ruta
            )
            st.info(
                f"**Reconocimiento del archivo:** {tipo_info['detalle']}\n\n"
                "Si es un **Plan de Trabajo**, ábralo en la pestaña "
                "**🗂️ Plan de Trabajo & ANS** (se enruta solo desde el cargador). "
                "Si es la **plantilla de seguimiento**, debe traer la hoja "
                "`PLANTILLA`."
            )
            st.code(f"{type(exc).__name__}: {exc}")
            hay_plantilla = False

    # --- Anonimización ----------------------------------------------------
    aviso_privacidad = anonimizar.aviso_para_la_interfaz()
    if aviso_privacidad:
        st.warning(aviso_privacidad, icon="🔒")
        df = anonimizar.anonimizar(df)
        df_completo = anonimizar.anonimizar(df_completo)

    modo_datos = anonimizar.modo_actual()
    badge_modo = (
        "<span class='badge-timestamp'>🔓 Datos Reales</span>"
        if modo_datos == anonimizar.MODO_NO
        else f"<span class='badge-timestamp'>🔒 {anonimizar.descripcion(modo_datos)}</span>"
    )

    st.markdown(
        f"📄 `{os.path.basename(ruta)}` · "
        f"Cálculo: **{momento:%Y-%m-%d %H:%M:%S}** · "
        f"Filas: **{total_hoja}** · Activos: **{total_activos}** · "
        f"En ventana: **{len(df)}** · {badge_modo} · "
        f"<span class='badge-timestamp'>🛠️ v{VERSION_APP}</span>",
        unsafe_allow_html=True,
    )

    for aviso in avisos:
        st.warning(f"⚠️ {aviso}", icon="⚠️")

    # --- Casos repetidos en la plantilla ----------------------------------
    # Un N° DE CASO en varias filas es lo que hacía que un caso ya cerrado siguiera
    # alertando como abierto (caso real: IM3238158). Se avisa aquí arriba, además de
    # marcarlo fila por fila en el Explorador.
    repetidos_plantilla = resumen_duplicados(df_completo)
    if not repetidos_plantilla.empty:
        st.error(
            f"🔁 **{len(repetidos_plantilla)} caso(s) están repetidos en la plantilla** "
            f"({int(repetidos_plantilla['VECES'].sum())} filas). Se marcan **DUPLICADO** y "
            "quedan **fuera del semáforo y de las métricas** hasta que se unifiquen. "
            "Detalle en **📋 Explorador de Casos & SLA**."
        )

    # --- Cargar Historial SQLite ------------------------------------------
    historial, hist_tecnicos, hist_detalle, hist_resumen, hist_error = leer_historial()

    # --- Conteos globales -------------------------------------------------
    conteo_completo = (
        df_completo["ESTADO"].value_counts().to_dict() if not df_completo.empty else {}
    )
    total_completo = len(df_completo)

    # Alcance de los avisos para la torre de control.
    # Si no hay plantilla valida (por ejemplo, se subio el Plan de Trabajo en su
    # lugar) los DataFrame vienen vacios y sin columnas: sin este corte se
    # accedia a df_completo["CERRADO"] y saltaba un KeyError.
    incluir_proximos = bool(st.session_state.get("notif_incluir_proximos", True))
    if df_completo.empty:
        conjunto_notificar = pd.DataFrame(columns=["TECNICO", "ESTADO", "CERRADO"])
        estados_notificar = ESTADOS_ALERTA
    elif incluir_proximos:
        estados_notificar = (ROJO, NARANJA, AMARILLO, VERDE)
        conjunto_notificar = df_completo[
            (~df_completo["CERRADO"])
            & df_completo["ESTADO"].isin(estados_notificar)
        ].copy()
    else:
        estados_notificar = ESTADOS_ALERTA
        conjunto_notificar = df[df["ESTADO"].isin(estados_notificar)].copy()

    tecnicos_notificables = {
        str(t).strip()
        for t in conjunto_notificar["TECNICO"].unique()
        if str(t).strip()
    }

    # =====================================================================
    # NAVEGACIÓN PRINCIPAL EN PESTAÑAS (ST.TABS)
    # =====================================================================
    (tab_despacho, tab_explorador, tab_analitica, tab_historial,
     tab_plan, tab_validacion, tab_integridad) = st.tabs([
        "🎯 Despacho Operativo",
        "📋 Explorador de Casos & SLA",
        "📊 Analítica & Técnicos",
        "📜 Historial & Auditoría",
        "🗂️ Plan de Trabajo & ANS",
        "⚖️ Validación de datos",
        "🔍 Integridad de datos",
    ])

    if hay_plantilla:
        # ---------------------------------------------------------------------
        # PESTAÑA 1: 🎯 DESPACHO OPERATIVO
        # ---------------------------------------------------------------------
        with tab_despacho:
            control = vista.barra_control(df_completo, df, momento)
            df_vista = control["datos"]

            # --- Lente por coordinacion: cambia TODO el turno ----------------
            # No reemplaza los filtros de arriba: es el corte Bogota / Regionales
            # que la torre usa para reportarle a cada coordinadora.
            reparto_completo = parte_por_region(df_completo)
            lente = selector_vista(
                ["Todas", "🏢 Bogotá", "🌎 Regionales"],
                "Todas",
                "lente_turno",
                "Muestra solo los casos que le corresponden a cada coordinación.",
            )
            if lente.endswith("Bogotá"):
                base = reparto_completo[BOGOTA]
            elif lente.endswith("Regionales"):
                base = reparto_completo[REGIONALES]
            else:
                base = df_completo

            def _reparto(marco) -> tuple[int, int]:
                """(casos de Bogota, casos de Regionales) de una lista."""
                if marco is None or marco.empty or "REGION_TECNICO" not in marco.columns:
                    return (0, 0)
                es_bogota = (
                    marco["REGION_TECNICO"].astype(str).str.strip().str.upper() == BOGOTA
                )
                return (int(es_bogota.sum()), int((~es_bogota).sum()))

            st.divider()

            # --- Los dos bloques: avance de hoy y acumulado del mes ---------
            # La culpa documentada sale del Plan de Trabajo (si esta cargado).
            plan_bytes_turno = st.session_state.get("plan_bytes")
            culpa_turno = (
                turno.culpa_por_caso(plan_bytes_turno) if plan_bytes_turno else {}
            )
            resumen_turno = turno.avance_y_acumulado(base, momento, culpa_turno)
            vista.bloques_turno(resumen_turno)

            st.markdown("##### 🚦 Reparto del semáforo (todos los casos)")
            vista.barra_semaforo(conteo_completo, total_completo)

            st.divider()
            st.caption(
                "**Orden de atención:** primero lo que todavía se puede salvar "
                "(🟠 y 🟡) y de último los ya vencidos (🔴), que no dependen de una "
                "acción del turno de hoy."
            )

            # --- Los 4 grupos, completos y en el orden pedido ---------------
            grupos_turno = turno.grupos_alerta(base)
            definicion = (
                (turno.GRUPO_INMINENTE, "1️⃣ 🟠 Se vencen en menos de 1 hora",
                 "Emergencia del turno: se atienden primero.",
                 "✅ Nada se vence en la próxima hora.", "turno_inm1h"),
                (turno.GRUPO_PREVENTIVO, "2️⃣ 🟡 Se vencen entre 1 y 4 horas",
                 "Todavía hay margen para gestionarlos.",
                 "✅ Nada se vence entre 1 y 4 horas.", "turno_prev4h"),
                (turno.GRUPO_VENCIDOS, "3️⃣ 🔴 Ya vencidos",
                 "El más atrasado primero: ya no dependen de una acción de hoy.",
                 "✅ No hay casos vencidos sin cerrar.", "turno_vencidos"),
                (turno.GRUPO_PROXIMOS, "4️⃣ 📅 Próximos 3 días",
                 "Aún se pueden salvar, pero sin la urgencia del turno.",
                 "✅ No hay casos que venzan en los próximos 3 días.", "turno_prox3d"),
            )
            for clave, titulo, subtitulo, vacio, clave_ui in definicion:
                marco = grupos_turno[clave]
                vista.lista_accion(
                    marco,
                    titulo,
                    subtitulo,
                    vacio,
                    clave_ui,
                    tecnicos_notificables=tecnicos_notificables,
                    df_completo=df_completo,
                    historial=historial,
                    reparto=_reparto(marco),
                )
                st.divider()

            # Contexto: lo que no es alerta del turno, pero existe y se cuenta
            # para que ningun caso quede invisible.
            mas_adelante = len(grupos_turno[turno.GRUPO_MAS_ADELANTE])
            sin_fecha = len(grupos_turno[turno.GRUPO_SIN_FECHA])
            abiertos_base = int((~base["CERRADO"]).sum()) if not base.empty else 0
            st.caption(
                f"Contexto: **{mas_adelante}** caso(s) vencen a más de 3 días y "
                f"**{sin_fecha}** no tienen fecha en la plantilla. Con los 4 grupos "
                f"suman los **{abiertos_base}** casos abiertos: nada queda oculto."
            )

            # --- Cierre del dia para cada coordinacion ----------------------
            st.divider()
            vencidos_region = {
                BOGOTA: turno.grupos_alerta(reparto_completo[BOGOTA])[turno.GRUPO_VENCIDOS],
                REGIONALES: turno.grupos_alerta(
                    reparto_completo[REGIONALES]
                )[turno.GRUPO_VENCIDOS],
            }
            vista.cierre_del_dia(
                reparto_completo, momento, culpa_turno, vencidos_region
            )

        # ---------------------------------------------------------------------
        # PESTAÑA 2: 📋 EXPLORADOR DE CASOS & SLA
        # ---------------------------------------------------------------------
        with tab_explorador:
            # Recuperar df_vista o calcular si se cambia de pestaña
            if "df_vista" not in locals():
                df_vista = df_completo

            # --- Conjunto recibido desde una tarjeta de Despacho ----------------
            # Los botones "Ver los N en el Explorador" dejan aqui la lista exacta,
            # para que se vean JUSTO esos casos sin volver a filtrar a mano.
            conjunto = st.session_state.get("set_explorador")
            if conjunto is not None and not conjunto.empty:
                origen = st.session_state.get("origen_explorador", "lista seleccionada")
                av1, av2 = st.columns([5, 1.4])
                with av1:
                    st.info(
                        f"🎯 Mostrando el conjunto cargado desde **{origen}**: "
                        f"**{len(conjunto)}** caso(s). Los filtros de abajo se aplican "
                        "sobre este conjunto."
                    )
                with av2:
                    if st.button(
                        "✖️ Quitar conjunto",
                        key="quitar_conjunto_explorador",
                        width="stretch",
                        help="Volver a ver todos los casos del sistema.",
                    ):
                        st.session_state.pop("set_explorador", None)
                        st.session_state.pop("origen_explorador", None)
                        st.rerun()
                df_vista = conjunto.copy()

            # --- Casos repetidos: aviso y detalle, antes de los filtros ------
            panel_duplicados(df_vista)

            st.subheader("🔎 Filtros avanzados de casos")
            f1, f2, f3, f4 = st.columns([2, 2, 2, 2])

            tecnicos = sorted(t for t in df_vista["TECNICO"].unique() if t)
            regiones = sorted(df_vista["REGION_TECNICO"].unique())
            estados = [e for e in sorted(df_vista["ESTADO"].unique(), key=lambda x: ORDEN_ESTADO[x])]

            with f1:
                sel_tecnicos = st.multiselect("👷 Técnico", tecnicos, placeholder="Todos los técnicos", key="exp_tecnicos")
            with f2:
                sel_regiones = st.multiselect("📍 Región", regiones, placeholder="Todas las regiones", key="exp_regiones")
            with f3:
                sel_estados = st.multiselect("🚦 Estado", estados, placeholder="Todos los estados", key="exp_estados")
            with f4:
                busqueda = st.text_input("🔍 Buscar caso / ciudad", placeholder="Ej: IM3237396", key="exp_busqueda")

            filtrado = df_vista.copy()
            if sel_tecnicos:
                filtrado = filtrado[filtrado["TECNICO"].isin(sel_tecnicos)]
            if sel_regiones:
                filtrado = filtrado[filtrado["REGION_TECNICO"].isin(sel_regiones)]
            if sel_estados:
                filtrado = filtrado[filtrado["ESTADO"].isin(sel_estados)]
            if busqueda.strip():
                patron = busqueda.strip()
                mascara = (
                    filtrado[COL_CASO].astype(str).str.contains(patron, case=False, na=False)
                    | filtrado[COL_CIUDAD].astype(str).str.contains(patron, case=False, na=False)
                )
                filtrado = filtrado[mascara]

            if REGION_DESCONOCIDA in set(filtrado["REGION_TECNICO"]):
                st.warning(
                    f"Hay casos con técnico no registrado en el diccionario de regiones "
                    f"({REGION_DESCONOCIDA}). Revise el mapeo en `core.py`."
                )

            st.markdown(f"#### 📋 Tabla de Casos ({len(filtrado)} de {len(df_vista)})")

            if filtrado.empty:
                st.info("Ningún caso coincide con los filtros seleccionados.")
            else:
                visible = filtrado[list(COLUMNAS_TABLA)].rename(columns=COLUMNAS_TABLA)
                venc = pd.to_datetime(visible["Vencimiento"], errors="coerce")
                visible["Vencimiento"] = (
                    venc.dt.strftime("%Y-%m-%d %H:%M").fillna("Sin fecha de vencimiento")
                )
                for col_txt in ("Ciudad", "Tecnico", "Caso", "Regional", "Departamento"):
                    if col_txt in visible.columns:
                        visible[col_txt] = visible[col_txt].astype(str)

                st.dataframe(
                    visible.style.apply(estilizar, axis=None).format(
                        {"Horas restantes": "{:,.2f}"}, na_rep="—"
                    ),
                    width="stretch",
                    hide_index=True,
                    height=min(650, 40 + 35 * len(visible)),
                    column_config={
                        "Horas restantes": st.column_config.NumberColumn(
                            "Horas restantes", help="Negativo = vencido", format="%.2f"
                        ),
                        " ": st.column_config.TextColumn(" ", width="small"),
                    },
                )

                st.download_button(
                    "⬇️ Descargar casos filtrados (CSV)",
                    data=visible.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"casos_sla_{momento:%Y%m%d_%H%M}.csv",
                    mime="text/csv",
                    key="descarga_casos_tab",
                )

        # ---------------------------------------------------------------------
        # PESTAÑA 3: 📊 ANALÍTICA & TÉCNICOS
        # ---------------------------------------------------------------------
        with tab_analitica:
            st.subheader("📈 Análisis de Gestión y Distribución de Carga")
            if "filtrado" not in locals():
                filtrado = df_completo

            # Gráficos de Altair con colores semánticos SLA
            dibujar_graficos_altair(filtrado)

            st.divider()

            # Métricas de Gestión por Técnico (Rankings, Cumplimiento, Matriz)
            st.subheader("🏆 Rankings de Gestión y Matriz de Cumplimiento")
            render_metricas_por_tecnico(df_completo, momento)

        # ---------------------------------------------------------------------
        # PESTAÑA 4: 📜 HISTORIAL & AUDITORÍA
        # ---------------------------------------------------------------------
        with tab_historial:
            st.subheader("🔔 Panel General de Notificaciones y Auditoría")

            st.checkbox(
                "Incluir también los próximos a vencer (no solo los vencidos y urgentes)",
                value=True,
                key="notif_incluir_proximos",
                help="Marcado: incluye casos que aún no vencen. Desmarcado: solo alerta inmediata.",
            )

            tecnico_pedido = st.session_state.pop("tecnico_a_avisar", None)
            if tecnico_pedido:
                st.success(f"Técnico seleccionado desde la lista: **{tecnico_pedido}**", icon="👉")

            render_notificaciones_pendientes(conjunto_notificar, historial, estados_notificar)

            st.divider()

            # Historial SQLite
            st.subheader("🗂️ Registro de Auditoría (SQLite)")
            if hist_error:
                st.warning(f"No se pudo leer el historial de notificaciones: {hist_error}")
            else:
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("🔔 Notificaciones", int(hist_resumen.get("total", 0)))
                k2.metric("✅ Enviadas", int(hist_resumen.get("enviadas", 0)))
                k3.metric("❌ Fallidas", int(hist_resumen.get("fallidas", 0)))
                k4.metric("📄 Casos", int(hist_resumen.get("casos", 0)))
                k5.metric("👷 Técnicos", int(hist_resumen.get("tecnicos", 0)))

                if hist_detalle is not None and not hist_detalle.empty:
                    st.dataframe(
                        hist_detalle.head(MAX_FILAS_DETALLE),
                        width="stretch",
                        hide_index=True,
                        height=min(450, 40 + 35 * len(hist_detalle)),
                    )
                    st.download_button(
                        "⬇️ Descargar Historial Completo (CSV)",
                        data=hist_detalle.to_csv(index=False).encode("utf-8-sig"),
                        file_name=f"historial_notificaciones_{momento:%Y%m%d_%H%M}.csv",
                        mime="text/csv",
                        key="descarga_historial_tab",
                    )

    else:
        st.info(
            "**Falta la plantilla de seguimiento SLA.** Suba el archivo de seguimiento de casos (el que trae la hoja `PLANTILLA`) desde la barra lateral. Mientras tanto, la pestaña **🗂️ Plan de Trabajo & ANS** sí está disponible si cargó ese archivo.",
            icon="📤",
        )
        if plantilla_error is not None:
            st.caption("Detalle de la búsqueda automática:")
            st.code(str(plantilla_error))

    # ---------------------------------------------------------------------
    # PESTANA 5: PLAN DE TRABAJO & ANS
    # ---------------------------------------------------------------------
    with tab_plan:
        render_plan_trabajo(df_completo)

    # ---------------------------------------------------------------------
    # PESTANA 6: VALIDACION DE DATOS (modulo aparte)
    # ---------------------------------------------------------------------
    # Los chequeos son los MISMOS que ya existian (fechas, nombres, repetidos y
    # causa documentada). Viven en su propio modulo para que la pantalla de turno
    # no tenga que bajar por toda la pagina para atender.
    with tab_validacion:
        vista.render_validacion(df_completo)

    # ---------------------------------------------------------------------
    # PESTANA 7: INTEGRIDAD DE DATOS (auditoria de filas: sigue igual)
    # ---------------------------------------------------------------------
    with tab_integridad:
        render_integridad(df_crudo, df_completo, nombre_archivo)

    if autorrefresco:
        import time
        time.sleep(INTERVALO_AUTOREFRESCO_S)
        st.rerun()


if __name__ == "__main__":
    main()
