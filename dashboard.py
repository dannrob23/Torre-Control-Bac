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

import pandas as pd
import streamlit as st

from core import (
    AMARILLO,
    CERRADO_OK,
    CERRADO_TARDE,
    COL_CASO,
    COL_CIUDAD,
    COL_DEPARTAMENTO,
    COL_REGIONAL,
    COL_VENCIMIENTO,
    ESTADOS_ALERTA,
    ICONO_ESTADO,
    NARANJA,
    ORDEN_ESTADO,
    REGION_DESCONOCIDA,
    ROJO,
    SIN_VENCIMIENTO,
    TODOS_LOS_ESTADOS,
    VERDE,
    ErrorLecturaExcel,
    calcular_tablero,
    localizar_excel,
    metricas_por_tecnico,
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

# Refresco automatico cada 60 s (opcional, desactivado por defecto para no
# interrumpir al usuario mientras filtra).
INTERVALO_AUTOREFRESCO_S = 60

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
}
COLORES_TEXTO = {
    ROJO: "#A30000",
    CERRADO_TARDE: "#FFFFFF",
    NARANJA: "#8A4B00",
    AMARILLO: "#7A5B00",
    VERDE: "#136B13",
    SIN_VENCIMIENTO: "#555555",
    CERRADO_OK: "#0B5D0B",
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
        (df_ventana, df_completo, total_hoja, total_activos, dias_ventana, avisos)
    """
    resultado = calcular_tablero(ruta, momento=datetime.fromisoformat(momento_iso))
    return (
        resultado.df,
        resultado.df_completo,
        resultado.total_hoja,
        resultado.total_activos,
        resultado.dias_ventana,
        list(resultado.warnings),
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

    momento = datetime.now()
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
    Envia el aviso a Telegram en TEXTO PLANO.

    parse_mode="" es obligatorio: los avisos de avisos.py son texto plano con
    emojis y pueden traer caracteres < > & que rompen el parseo HTML.
    Devuelve cuantos destinos recibieron el mensaje.
    """
    if not texto or not texto.strip():
        return 0
    if todos:
        return int(telegram_notifier.enviar_a_todos(texto, parse_mode=""))

    destinos = telegram_notifier.destinos_configurados()
    if not destinos:
        return 0
    primero = destinos[0]
    return int(
        telegram_notifier.enviar_mensaje(
            texto,
            parse_mode="",
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
        use_container_width=True,
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
            use_container_width=True,
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
            use_container_width=True,
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
            use_container_width=True,
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
        use_container_width=True,
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
        use_container_width=True,
        height=min(500, 40 + 35 * len(matriz)),
    )
    st.caption(
        "Columnas en orden de gravedad: "
        + " · ".join(f"{ICONO_ESTADO[e]} {e}" for e in TODOS_LOS_ESTADOS)
    )


# ---------------------------------------------------------------------------
# Notificaciones pendientes a tecnicos (avisos.py)
# ---------------------------------------------------------------------------

def render_notificaciones_pendientes(casos: pd.DataFrame, historial) -> None:
    """
    Dibuja la seccion "📨 Notificaciones pendientes a tecnicos".

    Flujo de la torre de control: ver la tabla de pendientes por tecnico, elegir
    uno, copiar el aviso con el boton nativo de st.code y pegarlo en WhatsApp, o
    enviarlo por Telegram. En ambos casos queda registro en el Historial.

    Args:
        casos: DataFrame de casos de la ventana (resultado.df). Se filtra aqui
               por ESTADOS_ALERTA, de modo que la funcion no depende del objeto
               Resultado (que main() no conserva).
        historial: instancia de Historial o None si no se pudo abrir.
    """
    st.divider()
    st.header("📨 Notificaciones pendientes a técnicos")
    st.caption(
        "Casos de la ventana que requieren aviso (ROJO, CERRADO TARDE, NARANJA y "
        "AMARILLO) agrupados por técnico. Genere el aviso, cópielo y péguelo en "
        "WhatsApp, o envíelo por Telegram. Cada aviso queda registrado en el historial."
    )

    # Los estados que requieren atencion ya vienen definidos en core.
    pendientes = casos[casos["ESTADO"].isin(ESTADOS_ALERTA)].copy() if casos is not None and not casos.empty else casos
    if pendientes is None or pendientes.empty:
        st.success(
            "✅ No hay casos pendientes de notificar: ningún técnico tiene casos en "
            "estado ROJO, CERRADO TARDE, NARANJA o AMARILLO dentro de la ventana."
        )
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

    tab_envio, tab_resumen, tab_historial = st.tabs(
        [
            "✉️ Aviso al técnico",
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
        texto = avisos.componer_aviso_tecnico(
            tecnico,
            df_tecnico,
            canal=canal,
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
                + " · Se envía en texto plano (`parse_mode=\"\"`) para no romper los emojis."
            )
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
                    alcanzados = enviar_aviso_telegram(texto, todos=todos)
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
    # Pestana 2: tabla de pendientes (a)
    # ==================================================================
    with tab_resumen:
        st.markdown("##### 📊 Técnicos con casos pendientes de notificar")
        st.dataframe(
            tabla_pendientes_visible(resumen),
            use_container_width=True,
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
            file_name=f"pendientes_por_tecnico_{datetime.now():%Y%m%d_%H%M}.csv",
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
                use_container_width=True,
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

def main() -> None:
    st.set_page_config(
        page_title="Torre de Control SLA - Colsof",
        page_icon="🛰️",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.title("🛰️ Torre de Control SLA — Colsof / Banco Agrario")
    st.caption("Monitoreo de Acuerdos de Nivel de Servicio para soporte tecnico en sitio.")

    # --- Barra lateral ----------------------------------------------------
    with st.sidebar:
        st.header("⚙️ Control")
        if st.button("🔄 Recalcular ahora", use_container_width=True):
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

    # --- Localizacion del archivo o subida manual -------------------------
    momento = datetime.now()

    # En la web no siempre hay archivo local: se permite SUBIR el Excel. Los
    # datos se procesan en memoria y no se guardan en el servidor.
    origen_subido = None
    with st.sidebar:
        st.divider()
        st.subheader("📤 Cargar plantilla")
        st.caption(
            "Suba el archivo Excel de seguimiento. Se procesa en memoria; "
            "no se guarda en el servidor."
        )
        archivo_subido = st.file_uploader(
            "Plantilla de casos (.xlsx)",
            type=["xlsx"],
            key="uploader_plantilla",
            help="Si no sube nada, se usa la plantilla que este junto al sistema.",
        )
        if archivo_subido is not None:
            origen_subido = archivo_subido

    if origen_subido is not None:
        contenido = origen_subido.getvalue()
        import hashlib

        firma = int(hashlib.md5(contenido).hexdigest()[:8], 16)
        ruta = f"(subido) {origen_subido.name}"
        lector = lambda: cargar_desde_bytes(
            contenido, float(firma), momento.isoformat(), origen_subido.name
        )
    else:
        try:
            ruta = localizar_excel(os.environ.get("TORRE_EXCEL"))
        except ErrorLecturaExcel as exc:
            st.error("❌ No se encontró la plantilla de seguimiento.")
            st.info(
                "**Suba el archivo Excel** desde la barra lateral, o colóquelo "
                "junto al sistema."
            )
            st.code(str(exc))
            st.stop()
        lector = lambda: cargar(ruta, firma_archivo(ruta), momento.isoformat())

    # --- Lectura ----------------------------------------------------------
    try:
        with st.spinner("Leyendo la plantilla y calculando SLA..."):
            (df, df_completo, total_hoja, total_activos,
             dias_ventana, avisos) = lector()
    except ErrorLecturaExcel as exc:
        st.error("❌ No se pudo leer el archivo Excel.")
        st.warning(
            "Causa más frecuente: **el archivo está abierto en Excel por otro "
            "usuario**. Ciérrelo y pulse *Recalcular ahora*."
        )
        st.code(str(exc))
        st.stop()
    except Exception as exc:
        st.error("❌ No se pudo procesar el archivo.")
        st.info(
            "Verifique que sea la plantilla correcta y que la hoja se llame "
            "**PLANTILLA**."
        )
        st.code(f"{type(exc).__name__}: {exc}")
        st.stop()

    # --- Anonimizacion (protege los datos si el tablero es publico) -------
    aviso_privacidad = anonimizar.aviso_para_la_interfaz()
    if aviso_privacidad:
        st.warning(aviso_privacidad, icon="🔒")
        df = anonimizar.anonimizar(df)
        df_completo = anonimizar.anonimizar(df_completo)

    st.caption(
        f"📄 `{os.path.basename(ruta)}` · "
        f"Calculo: **{momento:%Y-%m-%d %H:%M:%S}** · "
        f"Filas en la hoja: **{total_hoja}** · Casos activos: **{total_activos}** · "
        f"En ventana ({dias_ventana} dias + vencidos): **{len(df)}**"
    )

    for aviso in avisos:
        st.warning(f"⚠️ {aviso}", icon="⚠️")

    # --- KPIs (semaforo v2: 7 estados, sobre TODOS los casos) -------------
    conteo_completo = (
        df_completo["ESTADO"].value_counts().to_dict() if not df_completo.empty else {}
    )
    total_completo = len(df_completo)

    st.subheader("🚦 Semaforo de casos (activos + cerrados)")
    # El orden es el de gravedad de core.ORDEN_ESTADO: ROJO primero, CERRADO OK al final.
    for grupo in (TODOS_LOS_ESTADOS[:4], TODOS_LOS_ESTADOS[4:]):
        columnas_kpi = st.columns(len(grupo))
        for columna, estado in zip(columnas_kpi, grupo):
            cantidad = int(conteo_completo.get(estado, 0))
            columna.metric(
                label=f"{ICONO_ESTADO[estado]} {estado}",
                value=cantidad,
                delta=f"{cantidad / total_completo * 100:.0f}% del total"
                if total_completo
                else None,
                delta_color="off",
                help=AYUDA_ESTADO[estado],
            )

    st.divider()

    if df.empty:
        st.success(
            "✅ No hay casos activos dentro de la ventana de notificacion "
            "(todos los casos abiertos estan resueltos o vencen mas adelante)."
        )
    else:
        # --- Filtros ------------------------------------------------------
        st.subheader("🔎 Filtros")
        f1, f2, f3, f4 = st.columns([2, 2, 2, 2])

        tecnicos = sorted(t for t in df["TECNICO"].unique() if t)
        regiones = sorted(df["REGION_TECNICO"].unique())
        estados = [e for e in sorted(df["ESTADO"].unique(), key=lambda x: ORDEN_ESTADO[x])]

        with f1:
            sel_tecnicos = st.multiselect("👷 Tecnico", tecnicos, placeholder="Todos los tecnicos")
        with f2:
            sel_regiones = st.multiselect("📍 Region", regiones, placeholder="Todas las regiones")
        with f3:
            sel_estados = st.multiselect("🚦 Estado", estados, placeholder="Todos los estados")
        with f4:
            busqueda = st.text_input("🔍 Buscar caso / ciudad", placeholder="Ej: IM3237396")

        filtrado = df.copy()
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

        # Las columnas sin region se resaltan para que no pasen desapercibidas.
        if REGION_DESCONOCIDA in set(filtrado["REGION_TECNICO"]):
            st.warning(
                f"Hay casos con tecnico no registrado en el diccionario de regiones "
                f"({REGION_DESCONOCIDA}). Revise el mapeo en `core.py`."
            )

        # --- Tabla --------------------------------------------------------
        st.subheader(f"📋 Casos en ventana ({len(filtrado)} de {len(df)})")

        if filtrado.empty:
            st.info("Ningun caso coincide con los filtros seleccionados.")
        else:
            visible = filtrado[list(COLUMNAS_TABLA)].rename(columns=COLUMNAS_TABLA)
            visible["Vencimiento"] = pd.to_datetime(visible["Vencimiento"]).dt.strftime(
                "%Y-%m-%d %H:%M"
            )

            st.dataframe(
                visible.style.apply(estilizar, axis=None).format(
                    {"Horas restantes": "{:,.2f}"}, na_rep="—"
                ),
                use_container_width=True,
                hide_index=True,
                height=min(700, 40 + 35 * len(visible)),
                column_config={
                    "Horas restantes": st.column_config.NumberColumn(
                        "Horas restantes", help="Negativo = vencido", format="%.2f"
                    ),
                    " " : st.column_config.TextColumn(" ", width="small"),
                },
            )

            st.download_button(
                "⬇️ Descargar casos filtrados (CSV)",
                data=visible.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"casos_sla_{momento:%Y%m%d_%H%M}.csv",
                mime="text/csv",
                key="descarga_casos",
            )

        # --- Analisis agregado --------------------------------------------
        st.divider()
        g1, g2 = st.columns(2)

        with g1:
            st.subheader("📍 Casos por region")
            por_region = (
                filtrado.groupby(["REGION_TECNICO", "ESTADO"])
                .size()
                .unstack(fill_value=0)
                .reindex(columns=list(TODOS_LOS_ESTADOS), fill_value=0)
            )
            st.bar_chart(por_region, height=320)

        with g2:
            st.subheader("👷 Carga por tecnico (top 15)")
            por_tecnico = (
                filtrado["TECNICO"].value_counts().head(15).rename("Casos activos")
            )
            st.bar_chart(por_tecnico, height=320)

    # ======================================================================
    # SECCION NUEVA: Notificaciones y Metricas
    # ======================================================================
    st.divider()
    st.header("🔔 Notificaciones y Métricas")
    st.caption(
        "Historial de avisos registrado en `historial.db` (SQLite) por el notificador, "
        "mas las metricas de gestion por tecnico calculadas sobre todos los casos."
    )

    historial, hist_tecnicos, hist_detalle, hist_resumen, hist_error = leer_historial()
    if hist_error:
        st.warning(
            "No se pudo leer el historial de notificaciones "
            f"(historial.db): {hist_error}"
        )

    # --- a) KPIs del historial -------------------------------------------
    st.subheader("📊 Indicadores del historial")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("🔔 Total notificaciones", int(hist_resumen.get("total", 0)))
    k2.metric("✅ Enviadas", int(hist_resumen.get("enviadas", 0)))
    k3.metric("❌ Fallidas", int(hist_resumen.get("fallidas", 0)))
    k4.metric("📄 Casos distintos", int(hist_resumen.get("casos", 0)))
    k5.metric("👷 Tecnicos notificados", int(hist_resumen.get("tecnicos", 0)))

    # --- b) Notificaciones por tecnico -----------------------------------
    st.subheader("👥 Notificaciones por tecnico")
    if hist_tecnicos is None or hist_tecnicos.empty:
        st.info("Aun no hay notificaciones registradas en el historial.")
    else:
        st.dataframe(
            hist_tecnicos,
            use_container_width=True,
            hide_index=True,
            height=min(430, 40 + 35 * len(hist_tecnicos)),
        )
        st.download_button(
            "⬇️ Descargar notificaciones por tecnico (CSV)",
            data=hist_tecnicos.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"notificaciones_por_tecnico_{momento:%Y%m%d_%H%M}.csv",
            mime="text/csv",
            key="descarga_notif_tecnico",
        )

    # --- c) Historial detallado con filtros ------------------------------
    st.subheader("🗂️ Historial detallado de notificaciones")
    if hist_detalle is None or hist_detalle.empty:
        st.info(
            "El historial esta vacio: todavia no se ha registrado ningun aviso. "
            "Se llena al ejecutar `alertas_windows.py` o el notificador de Telegram."
        )
    else:
        fechas = pd.to_datetime(hist_detalle["fecha_hora"], errors="coerce")
        fecha_max, fecha_min = fechas.max(), fechas.min()
        hoy = datetime.now().date()
        defecto_hasta = fecha_max.date() if pd.notna(fecha_max) else hoy
        defecto_desde = (
            (fecha_max - pd.Timedelta(days=30)).date() if pd.notna(fecha_max)
            else hoy - timedelta(days=30)
        )
        if pd.notna(fecha_min) and defecto_desde < fecha_min.date():
            defecto_desde = fecha_min.date()

        h1, h2, h3 = st.columns([2, 1, 1])
        with h1:
            tecnicos_notif = sorted(
                str(t) for t in hist_detalle["tecnico"].dropna().unique() if str(t).strip()
            )
            sel_notif_tecnicos = st.multiselect(
                "👷 Tecnico",
                tecnicos_notif,
                placeholder="Todos los tecnicos",
                key="filtro_notif_tecnico",
            )
        with h2:
            desde = st.date_input("📅 Desde", value=defecto_desde, key="filtro_notif_desde")
        with h3:
            hasta = st.date_input("📅 Hasta", value=defecto_hasta, key="filtro_notif_hasta")

        detalle = hist_detalle
        if desde and hasta:
            if desde > hasta:
                st.warning("La fecha 'Desde' es posterior a 'Hasta': se intercambian los limites.")
                desde, hasta = hasta, desde
            # El filtro de fechas se delega en la API del historial.
            if historial is not None:
                detalle = historial.leer(
                    desde=f"{desde} 00:00:00", hasta=f"{hasta} 23:59:59"
                )
            else:
                detalle = detalle[
                    (hist_detalle["fecha_hora"] >= f"{desde} 00:00:00")
                    & (hist_detalle["fecha_hora"] <= f"{hasta} 23:59:59")
                ]
        if sel_notif_tecnicos:
            detalle = detalle[detalle["tecnico"].isin(sel_notif_tecnicos)]

        if detalle.empty:
            st.info("Ninguna notificacion coincide con los filtros seleccionados.")
        else:
            visibles = detalle.head(MAX_FILAS_DETALLE)
            st.caption(
                f"Mostrando {len(visibles)} de {len(detalle)} notificacion(es) "
                f"(maximo {MAX_FILAS_DETALLE})."
            )
            st.dataframe(
                visibles,
                use_container_width=True,
                hide_index=True,
                height=min(500, 40 + 35 * len(visibles)),
            )
            st.download_button(
                "⬇️ Descargar historial filtrado (CSV)",
                data=detalle.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"historial_notificaciones_{momento:%Y%m%d_%H%M}.csv",
                mime="text/csv",
                key="descarga_notif_detalle",
            )

    # ======================================================================
    # SECCION NUEVA: Notificaciones pendientes a tecnicos
    # ======================================================================
    render_notificaciones_pendientes(df, historial)

    # ======================================================================
    # Metricas de gestion por tecnico (core.metricas_por_tecnico)
    # ======================================================================
    st.divider()
    st.subheader("📈 Metricas de gestion por tecnico")
    render_metricas_por_tecnico(df_completo, momento)

    if autorrefresco:
        # Recarga la pagina para recalcular el tiempo restante.
        import time

        time.sleep(INTERVALO_AUTOREFRESCO_S)
        st.rerun()


if __name__ == "__main__":
    main()
