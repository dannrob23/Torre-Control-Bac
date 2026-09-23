"""
vista.py - Componentes del rediseno "Accion primero" (Prototipo A).

QUE RESUELVE:
    El tablero original ponia 7 tarjetas KPI del mismo tamano, asi que
    "Cerrado OK 98" pesaba visualmente igual que "Vencidos 232". Aqui la
    gravedad manda: rojo grande, verde pequeno, y cada fila ya trae el boton
    para actuar.

QUE APORTA:
    1. Barra de control: elegir VISTA (operativa o historico completo) y
       filtrar por FECHA con un desplegable (mas rango personalizado).
    2. Franja de foco: los 3 numeros que cambian el dia.
    3. Barra de semaforo segmentada en vez de 7 tarjetas.
    4. Listas de accion: "Vencen hoy" (primero) y "Ya vencidos" (el mas
       atrasado primero), con boton de avisar en cada fila.

Este modulo NO sabe de Telegram ni de login: solo dibuja y filtra.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

from core import (
    AMARILLO,
    CERRADO_OK,
    CERRADO_TARDE,
    COL_CASO,
    COL_CIUDAD,
    ICONO_ESTADO,
    ahora_colombia,
    NARANJA,
    ORDEN_ESTADO,
    REGION_DESCONOCIDA,
    ROJO,
    SIN_VENCIMIENTO,
    TODOS_LOS_ESTADOS,
    VERDE,
)

# ---------------------------------------------------------------------------
# Colores (mismos del dashboard, para no romper la identidad visual)
# ---------------------------------------------------------------------------
COLOR_SEMAFORO = {
    ROJO: "#B91C1C",
    CERRADO_TARDE: "#8B0000",
    NARANJA: "#C2410C",
    AMARILLO: "#A16207",
    VERDE: "#15803D",
    SIN_VENCIMIENTO: "#9CA3AF",
    CERRADO_OK: "#15803D",
}

# Colores de las etiquetas de tiempo en las listas de accion.
COLOR_ETIQUETA = {
    ROJO: ("#FEE2E2", "#7F1D1D"),
    CERRADO_TARDE: ("#FEE2E2", "#7F1D1D"),
    NARANJA: ("#FFEDD5", "#C2410C"),
    AMARILLO: ("#FEF9C3", "#A16207"),
    VERDE: ("#DCFCE7", "#15803D"),
    SIN_VENCIMIENTO: ("#F3F4F6", "#6B7280"),
    CERRADO_OK: ("#DCFCE7", "#15803D"),
}

# ---------------------------------------------------------------------------
# Filtro por fecha
# ---------------------------------------------------------------------------

PRESET_TODAS = "Todas las fechas"
PRESET_HOY = "Hoy"
PRESET_3D = "Ultimos 3 dias"
PRESET_7D = "Ultimos 7 dias"
PRESET_30D = "Ultimos 30 dias"
PRESET_MES = "Mes en curso"
PRESET_RANGO = "Rango personalizado..."

PRESETS_FECHA = (
    PRESET_TODAS,
    PRESET_HOY,
    PRESET_3D,
    PRESET_7D,
    PRESET_30D,
    PRESET_MES,
    PRESET_RANGO,
)

# Sobre que fecha se filtra. La clave es el nombre de la columna en el dataframe.
BASES_FECHA = {
    "Vencimiento del caso": "FECHA_VENCIMIENTO",
    "Fecha de creacion del caso": "FECHA_CREACION",
    "Fecha de resolucion (cierre)": "FECHA_RESOLUCION",
}

# ---------------------------------------------------------------------------
# Filtro por situacion del caso (abierto / cerrado)
#
# Sirve para VALIDAR los cierres: en el historico se puede aislar lo que ya se
# cerro, y separar lo cerrado a tiempo de lo cerrado tarde (incumplimiento).
# ---------------------------------------------------------------------------
MOSTRAR_TODOS = "Todos los casos"
MOSTRAR_ABIERTOS = "Solo abiertos (pendientes)"
MOSTRAR_CERRADOS = "Solo cerrados (todos)"
MOSTRAR_CERRADOS_OK = "Cerrados a tiempo (CERRADO OK)"
MOSTRAR_CERRADOS_TARDE = "Cerrados tarde (incumplimiento ANS)"

MOSTRAR_OPCIONES = (
    MOSTRAR_TODOS,
    MOSTRAR_ABIERTOS,
    MOSTRAR_CERRADOS,
    MOSTRAR_CERRADOS_OK,
    MOSTRAR_CERRADOS_TARDE,
)



def aplicar_filtro_situacion(
    df: pd.DataFrame, opcion: str
) -> tuple[pd.DataFrame, str]:
    """
    Filtra por abierto/cerrado. Devuelve (dataframe, etiqueta legible).

    'abierto' se calcula con la columna CERRADO que ya trae core, no por el
    estado del semaforo: un caso puede estar vencido y aun asi estar abierto.
    """
    if df is None or df.empty:
        return df, opcion
    if opcion == MOSTRAR_TODOS or "CERRADO" not in df.columns:
        return df, MOSTRAR_TODOS

    cerrado = df["CERRADO"].fillna(False).astype(bool)
    if opcion == MOSTRAR_ABIERTOS:
        return df[~cerrado].copy(), MOSTRAR_ABIERTOS
    if opcion == MOSTRAR_CERRADOS:
        return df[cerrado].copy(), MOSTRAR_CERRADOS
    if opcion == MOSTRAR_CERRADOS_OK:
        return df[cerrado & (df["ESTADO"] == CERRADO_OK)].copy(), MOSTRAR_CERRADOS_OK
    if opcion == MOSTRAR_CERRADOS_TARDE:
        return (
            df[cerrado & (df["ESTADO"] == CERRADO_TARDE)].copy(),
            MOSTRAR_CERRADOS_TARDE,
        )
    return df, opcion


def aplicar_filtro_fecha(
    df: pd.DataFrame,
    columna: str,
    preset: str,
    desde: date | None = None,
    hasta: date | None = None,
    ahora: datetime | None = None,
) -> tuple[pd.DataFrame, str, int]:
    """
    Filtra por fecha y devuelve (dataframe, etiqueta legible, filas sin fecha).

    'filas sin fecha' son los casos cuya columna de fecha esta vacia (NaT):
    no se pueden ubicar en el calendario, asi que se excluyen del filtro pero
    se informan, para que no desaparezcan en silencio.
    """
    if df is None or df.empty:
        return df, preset, 0
    if columna not in df.columns:
        return df, preset, 0

    if preset == PRESET_TODAS:
        return df, PRESET_TODAS, 0

    ahora = ahora or ahora_colombia()
    hoy = ahora.date()
    fechas = pd.to_datetime(df[columna], errors="coerce")
    sin_fecha = int(fechas.isna().sum())

    if preset == PRESET_HOY:
        lo, hi = hoy, hoy
        etiqueta = f"Hoy ({hoy:%d/%m/%Y})"
    elif preset == PRESET_3D:
        lo, hi = hoy - timedelta(days=2), hoy
        etiqueta = f"Ultimos 3 dias ({lo:%d/%m} a {hi:%d/%m})"
    elif preset == PRESET_7D:
        lo, hi = hoy - timedelta(days=6), hoy
        etiqueta = f"Ultimos 7 dias ({lo:%d/%m} a {hi:%d/%m})"
    elif preset == PRESET_30D:
        lo, hi = hoy - timedelta(days=29), hoy
        etiqueta = f"Ultimos 30 dias ({lo:%d/%m} a {hi:%d/%m})"
    elif preset == PRESET_MES:
        lo = hoy.replace(day=1)
        hi = hoy
        etiqueta = f"Mes en curso ({lo:%d/%m/%Y} a {hi:%d/%m/%Y})"
    elif preset == PRESET_RANGO:
        lo = desde or hoy
        hi = hasta or hoy
        if lo > hi:
            lo, hi = hi, lo
        etiqueta = f"Del {lo:%d/%m/%Y} al {hi:%d/%m/%Y}"
    else:  # pragma: no cover - preset desconocido
        return df, preset, 0

    limite_inf = pd.Timestamp(lo)
    limite_sup = pd.Timestamp(hi) + pd.Timedelta(days=1)  # incluye todo el dia final
    mascara = fechas.notna() & (fechas >= limite_inf) & (fechas < limite_sup)
    return df[mascara].copy(), etiqueta, sin_fecha


# ---------------------------------------------------------------------------
# Barra de control (vista + fecha)
# ---------------------------------------------------------------------------

VISTA_OPERATIVA = "operativa"
VISTA_HISTORICO = "historico"


def barra_control(
    df_completo: pd.DataFrame,
    df_operativa: pd.DataFrame,
    ahora: datetime,
) -> dict:
    """
    Dibuja los controles de la parte superior y devuelve los datos a mostrar.

    IMPORTANTE: por defecto se muestran TODOS los casos y TODAS las fechas.
    La app no recorta nada por su cuenta: los filtros son para acotar, no para
    esconder. Antes el arranque mostraba solo la ventana operativa (99 casos) y
    parecia que faltaban datos.

    Devuelve un dict con:
        datos            -> dataframe ya filtrado
        vista            -> "operativa" o "historico"
        columna_fecha    -> columna usada para filtrar por fecha
        etiqueta_fecha   -> texto legible del filtro de fecha
        etiqueta_mostrar -> texto legible del filtro abierto/cerrado
        sin_fecha        -> casos excluidos por no tener fecha
        n_cerrados       -> cuantos cerrados hay en lo mostrado
        n_abiertos       -> cuantos abiertos hay en lo mostrado
    """
    st.markdown("##### 🔎 Qué datos quieres ver")

    fila1 = st.columns([3, 3])
    with fila1[0]:
        vista_etiqueta = st.selectbox(
            "Vista",
            [
                "📚 Todos los casos (histórico completo)",
                "🚦 Solo lo urgente (por vencer + vencidos)",
            ],
            index=0,
            key="vista_datos",
            help=(
                "Todos los casos: los 483 de la plantilla, abiertos y cerrados. "
                "Solo lo urgente: únicamente lo que exige acción hoy (la ventana "
                "que calcula el sistema)."
            ),
        )
    with fila1[1]:
        mostrar = st.selectbox(
            "Mostrar",
            MOSTRAR_OPCIONES,
            index=0,
            key="mostrar_situacion",
            help=(
                "Sirve para VALIDAR cierres: aísla lo abierto o lo cerrado, y "
                "separa lo cerrado a tiempo de lo cerrado tarde."
            ),
        )

    fila2 = st.columns([3, 3])
    with fila2[0]:
        base_etiqueta = st.selectbox(
            "Filtrar fecha por",
            list(BASES_FECHA),
            index=0,
            key="base_fecha",
            help="Sobre qué fecha se aplica el rango de la derecha.",
        )
    columna_fecha = BASES_FECHA[base_etiqueta]

    with fila2[1]:
        preset = st.selectbox(
            "Rango de fecha",
            PRESETS_FECHA,
            index=0,
            key="preset_fecha",
            help="Todas las fechas por defecto. Elige un rango para acotar los datos.",
        )

    desde = hasta = None
    if preset == PRESET_RANGO:
        d1, d2 = st.columns(2)
        with d1:
            desde = st.date_input("Desde", value=ahora.date(), key="fecha_desde")
        with d2:
            hasta = st.date_input("Hasta", value=ahora.date(), key="fecha_hasta")

    # La vista operativa ya viene acotada por core; "todos los casos" es todo.
    vista = VISTA_HISTORICO if vista_etiqueta.startswith("📚") else VISTA_OPERATIVA
    base = df_completo if vista == VISTA_HISTORICO else df_operativa

    # Primero la fecha, despues la situacion (abierto/cerrado).
    datos, etiqueta, sin_fecha = aplicar_filtro_fecha(
        base, columna_fecha, preset, desde, hasta, ahora
    )
    datos, etiqueta_mostrar = aplicar_filtro_situacion(datos, mostrar)

    n_cerrados = 0
    if "CERRADO" in datos.columns and not datos.empty:
        n_cerrados = int(datos["CERRADO"].fillna(False).astype(bool).sum())
    n_abiertos = len(datos) - n_cerrados

    partes = [
        f"**{len(datos)}** de **{len(base)}** caso(s)",
        f"vista **{'todos los casos' if vista == VISTA_HISTORICO else 'solo lo urgente'}**",
        f"situación: **{etiqueta_mostrar}**",
        f"{base_etiqueta.lower()}: **{etiqueta}**",
    ]
    if sin_fecha:
        partes.append(f"⚠️ {sin_fecha} sin esa fecha quedaron fuera")
    st.caption("Mostrando " + " · ".join(partes))
    st.caption(
        f"De esos {len(datos)}: **{n_abiertos}** abierto(s) · "
        f"**{n_cerrados}** cerrado(s). Cambia **Mostrar** para validar los cierres."
    )

    return {
        "datos": datos,
        "vista": vista,
        "columna_fecha": columna_fecha,
        "etiqueta_fecha": etiqueta,
        "etiqueta_mostrar": etiqueta_mostrar,
        "sin_fecha": sin_fecha,
        "base_etiqueta": base_etiqueta,
        "n_cerrados": n_cerrados,
        "n_abiertos": n_abiertos,
    }


# ---------------------------------------------------------------------------
# Franja de foco: los 3 numeros que cambian el dia
# ---------------------------------------------------------------------------

def _tarjeta_foco(numero: int, etiqueta: str, ayuda: str, color: str) -> str:
    return (
        f"<div class='kpi-card-v2' style='border-left: 5px solid {color};'>"
        f"<div class='kpi-label'>{etiqueta}</div>"
        f"<div class='kpi-number' style='color:{color};'>{numero}</div>"
        f"<div class='kpi-help'>{ayuda}</div>"
        f"</div>"
    )


def franja_foco(
    vencidos: int,
    vencen_hoy: int,
    proximos_3d: int,
    total_casos: int,
) -> None:
    """Los tres numeros grandes de arriba, con la gravedad por delante."""
    pct = f"{vencidos / total_casos * 100:.0f}%" if total_casos else "—"
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            _tarjeta_foco(
                vencidos, "Vencidos sin cerrar",
                f"{pct} de los {total_casos} casos · requieren acción inmediata",
                COLOR_SEMAFORO[ROJO],
            ),
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            _tarjeta_foco(
                vencen_hoy, "Vencen en las próximas 24 h",
                "Última oportunidad · avísales antes de que se venzan",
                COLOR_SEMAFORO[NARANJA],
            ),
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            _tarjeta_foco(
                proximos_3d, "Próximos 3 días",
                "Aún hay margen · se pueden salvar",
                "#2563EB",
            ),
            unsafe_allow_html=True,
        )


# ---------------------------------------------------------------------------
# Barra de semaforo segmentada
# ---------------------------------------------------------------------------

def barra_semaforo(conteo: dict[str, int], total: int) -> None:
    """Una sola barra de 12 px con el reparto real, en vez de 7 tarjetas."""
    if not total:
        st.info("Sin casos para el reparto del semáforo.")
        return

    segmentos = []
    for estado in TODOS_LOS_ESTADOS:
        cantidad = int(conteo.get(estado, 0))
        if not cantidad:
            continue
        ancho = cantidad / total * 100
        segmentos.append(
            f"<span title='{estado}: {cantidad} caso(s)' style='width:{ancho:.2f}%;"
            f"background:{COLOR_SEMAFORO[estado]};display:block;height:100%'></span>"
        )

    leyenda = "".join(
        f"<span style='margin-right:14px;white-space:nowrap'>"
        f"<i style='display:inline-block;width:10px;height:10px;border-radius:3px;"
        f"background:{COLOR_SEMAFORO[e]};margin-right:6px'></i>"
        f"{ICONO_ESTADO[e]} {e} <b>{int(conteo.get(e, 0))}</b></span>"
        for e in TODOS_LOS_ESTADOS
        if conteo.get(e)
    )

    st.markdown(
        "<div style='display:flex;height:14px;border-radius:999px;overflow:hidden;"
        "box-shadow:inset 0 1px 2px rgba(0,0,0,0.1);margin:6px 0 12px'>" + "".join(segmentos) + "</div>"
        "<div style='font-size:12.5px;color:#4B5563;line-height:2.0'>"
        + leyenda
        + "</div>",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Listas de accion
# ---------------------------------------------------------------------------

def _resumen_tecnicos(df: pd.DataFrame, top: int = 4) -> str:
    """Etiqueta corta con los eventuales tecnicos del caso (casi siempre 1)."""
    if "TECNICO" not in df.columns:
        return ""
    nombres = [str(t).strip() for t in df["TECNICO"].head(top) if str(t).strip()]
    return " · ".join(nombres)


def _clase_fondo(estado: str) -> str:
    """Clase CSS con el color de fondo semaforico para la fila."""
    return {
        ROJO: "fondo-rojo",
        CERRADO_TARDE: "fondo-cerrado-tarde",
        NARANJA: "fondo-naranja",
        AMARILLO: "fondo-amarillo",
        VERDE: "fondo-verde",
        CERRADO_OK: "fondo-verde",
    }.get(estado, "fondo-gris")


def fila_accion_html(fila: pd.Series) -> str:
    """HTML de una fila de la lista de accion (caso + tiempo + contexto)."""
    estado = str(fila.get("ESTADO", ""))
    icono = ICONO_ESTADO.get(estado, "⚠️")
    caso = str(fila.get(COL_CASO, "S/N")).strip()
    tiempo = str(fila.get("TIEMPO_VENCIDO" if estado in (ROJO, CERRADO_TARDE)
                          else "TIEMPO_RESTANTE", "—"))

    clase_badge = {
        ROJO: "badge-rojo",
        CERRADO_TARDE: "badge-cerrado-tarde",
        NARANJA: "badge-naranja",
        AMARILLO: "badge-amarillo",
        VERDE: "badge-verde",
        SIN_VENCIMIENTO: "badge-gris",
        CERRADO_OK: "badge-cerrado-ok",
    }.get(estado, "badge-gris")

    tecnico = str(fila.get("TECNICO", "") or "").strip() or "(sin técnico)"
    oficina = str(fila.get(COL_CIUDAD, "") or "").strip() or "(sin oficina)"
    region = str(fila.get("REGION_TECNICO", "") or "").strip()

    contexto = f"{tecnico} · {oficina}"
    if region and region != "SIN REGION":
        contexto += f" · {region}"

    return (
        f"<div class='fila-accion-card {_clase_fondo(estado)}'>"
        f"<div style='display:flex;align-items:center;justify-content:space-between;'>"
        f"<span style='font-weight:700;font-size:15px;color:#111827;'>{icono} {caso}</span>"
        f"<span class='badge-sla {clase_badge}'>{tiempo}</span>"
        f"</div>"
        f"<div style='font-size:12.5px;color:#6B7280;margin-top:4px;'>{contexto}</div>"
        f"</div>"
    )


def lista_accion(
    df: pd.DataFrame,
    titulo: str,
    subtitulo: str,
    vacio: str,
    clave: str,
    limite: int = 6,
    permitir_avisar: bool = True,
    tecnicos_notificables: set[str] | None = None,
    df_completo: pd.DataFrame | None = None,
    historial=None,
) -> None:
    """
    Dibuja una lista de casos con boton popover flotante para avisar en la misma fila.
    """
    import avisos
    import telegram_notifier

    if df is None or df.empty:
        st.markdown(f"#### {titulo}")
        st.info(vacio)
        return

    # --- Encabezado con el TOTAL real, sin truncados silenciosos -----------
    n_total = len(df)
    st.markdown(
        f"<div class='encabezado-lista'>"
        f"<span style='font-size:17px;font-weight:700;'>{titulo}</span>"
        f"<span class='conteo-total'>{n_total} caso{'s' if n_total != 1 else ''}</span>"
        f"</div>",
        unsafe_allow_html=True,
    )
    if subtitulo:
        st.caption(subtitulo)

    # --- contador de filas visibles, persistido por lista ------------------
    clave_vista = f"_ver_{clave}"
    if clave_vista not in st.session_state:
        st.session_state[clave_vista] = limite
    visibles = max(int(st.session_state[clave_vista]), limite)
    visibles = min(visibles, n_total)
    st.session_state[clave_vista] = visibles

    menciones = avisos.cargar_menciones()

    for i in range(visibles):
        fila = df.iloc[i]
        tecnico = str(fila.get("TECNICO", "") or "").strip()
        c1, c2 = st.columns([4.8, 1.6])
        with c1:
            st.markdown(fila_accion_html(fila), unsafe_allow_html=True)
        with c2:
            if not permitir_avisar:
                st.caption("—")
            elif not tecnico:
                st.caption("sin técnico\nno se puede avisar")
            elif tecnicos_notificables is not None and tecnico not in tecnicos_notificables:
                st.caption("sin casos\npendientes")
            else:
                # Popover emergente directo en la fila
                with st.popover(
                    "📨 Avisar",
                    width="stretch",
                    help=f"Generar aviso inmediato para {tecnico}",
                ):
                    st.markdown(f"### 📨 Notificación para {tecnico}")

                    base_casos = df_completo if df_completo is not None else df
                    df_tec = (
                        base_casos[
                            (base_casos["TECNICO"] == tecnico) & (~base_casos["CERRADO"])
                        ]
                        if "CERRADO" in base_casos.columns
                        else base_casos[base_casos["TECNICO"] == tecnico]
                    )

                    texto_wa = avisos.componer_aviso_tecnico(
                        tecnico, df_tec, canal="whatsapp", menciones=menciones
                    )
                    texto_tg = avisos.componer_aviso_tecnico_telegram(
                        tecnico, df_tec, menciones=menciones
                    )

                    st.markdown("**Copiar texto para WhatsApp:**")
                    st.code(texto_wa, language=None)
                    st.caption("Usa el icono de copiar arriba a la derecha de la caja.")

                    p1, p2 = st.columns(2)
                    with p1:
                        if st.button(
                            "✅ Marcar WhatsApp",
                            key=f"{clave}_pop_wa_{i}",
                            width="stretch",
                        ):
                            if historial is not None and not df_tec.empty:
                                for _, f_tec in df_tec.iterrows():
                                    historial.registrar(
                                        caso=str(f_tec.get(COL_CASO, "S/N")),
                                        tecnico=tecnico,
                                        region=str(f_tec.get("REGION_TECNICO", "")),
                                        estado=str(f_tec.get("ESTADO", "")),
                                        horas_restantes=f_tec.get("HORAS_RESTANTES"),
                                        horas_vencido=f_tec.get("HORAS_VENCIDO"),
                                        canal="whatsapp",
                                        resultado="generado",
                                        detalle="marcado desde popover flotante",
                                    )
                                st.success("✅ Notificación registrada.")
                                st.session_state["notif_pend_tecnico"] = tecnico
                                st.rerun()
                    with p2:
                        if telegram_notifier.telegram_configurado():
                            if st.button(
                                "🚀 Enviar Telegram",
                                key=f"{clave}_pop_tg_{i}",
                                width="stretch",
                                type="primary",
                            ):
                                ok = telegram_notifier.enviar_a_todos(texto_tg, parse_mode="HTML")
                                if ok and historial is not None:
                                    for _, f_tec in df_tec.iterrows():
                                        historial.registrar(
                                            caso=str(f_tec.get(COL_CASO, "S/N")),
                                            tecnico=tecnico,
                                            region=str(f_tec.get("REGION_TECNICO", "")),
                                            estado=str(f_tec.get("ESTADO", "")),
                                            horas_restantes=f_tec.get("HORAS_RESTANTES"),
                                            horas_vencido=f_tec.get("HORAS_VENCIDO"),
                                            canal="telegram",
                                            resultado="enviado",
                                            detalle="enviado desde popover flotante",
                                        )
                                    st.success("✅ Telegram enviado y registrado.")
                                    st.rerun()
                        else:
                            st.caption("Telegram no configurado")

        if i < visibles - 1:
            st.markdown("<div style='margin-bottom:6px;'></div>", unsafe_allow_html=True)

    # --- Controles de la lista completa ------------------------------------
    # NO se oculta nada: se indica el rango mostrado sobre el total, se permite
    # seguir mostrando de a poco y se ofrece la descarga completa.
    faltan = n_total - visibles
    c_ver, c_filtro, c_csv = st.columns([1.3, 1.5, 1.6])

    with c_ver:
        if faltan > 0:
            if st.button(
                f"➕ Ver {min(limite, faltan)} más",
                key=f"{clave}_vermas",
                width="stretch",
            ):
                st.session_state[clave_vista] = visibles + limite
                st.rerun()
        else:
            st.button(
                "✅ Lista completa",
                key=f"{clave}_completa",
                width="stretch",
                disabled=True,
            )

    with c_filtro:
        if st.button(
            f"🔎 Ver los {n_total} en el Explorador",
            key=f"{clave}_ir_explorador",
            width="stretch",
            help="Abre la pestaña Explorador de Casos con este mismo conjunto.",
        ):
            # Se guarda la lista para que el Explorador la muestre tal cual.
            st.session_state["set_explorador"] = df.copy()
            st.session_state["origen_explorador"] = titulo
            st.success(
                "✅ Conjunto cargado. Abre la pestaña **📋 Explorador de Casos & SLA**."
            )

    with c_csv:
        st.download_button(
            f"⬇️ Descargar los {n_total} (CSV)",
            data=_csv_lista(df),
            file_name=f"{_nombre_archivo(titulo)}_{datetime.now():%Y%m%d_%H%M}.csv",
            mime="text/csv",
            key=f"{clave}_csv",
            width="stretch",
        )

    st.caption(
        f"Mostrando **{visibles}** de **{n_total}** · "
        f"Ordenados del más urgente al menos urgente."
        + (" Nada oculto: usa *Ver más* o descarga el CSV." if faltan else "")
    )


def _nombre_archivo(texto: str) -> str:
    """Convierte un titulo en nombre de archivo valido."""
    limpio = "".join(c if c.isalnum() or c in " -_" else "" for c in str(texto))
    return "_".join(limpio.split())[:60] or "casos"


def _csv_lista(df: pd.DataFrame) -> bytes:
    """
    CSV con las columnas utiles de la lista, listo para abrir en Excel.

    Se usa utf-8-sig para que Excel respete las tildes y la enye.
    """
    columnas = [
        (COL_CASO, "Caso"),
        ("ESTADO", "Estado"),
        ("TIEMPO_VENCIDO", "Tiempo vencido"),
        ("TIEMPO_RESTANTE", "Tiempo restante"),
        ("FECHA_VENCIMIENTO", "Vencimiento"),
        ("TECNICO", "Tecnico"),
        ("REGION_TECNICO", "Region"),
        (COL_CIUDAD, "Oficina"),
        ("HORAS_VENCIDO", "Horas vencido"),
        ("HORAS_RESTANTES", "Horas restantes"),
    ]
    presentes = [(c, n) for c, n in columnas if c in df.columns]
    vista = df[[c for c, _ in presentes]].rename(columns=dict(presentes))
    return vista.to_csv(index=False).encode("utf-8-sig", errors="replace")


# ---------------------------------------------------------------------------
# Construccion de las listas de accion
# ---------------------------------------------------------------------------

def construir_listas(df_completo: pd.DataFrame, ahora: datetime) -> dict:
    """
    Arma los conjuntos de accion a partir de TODOS los casos (no de la ventana),
    para que los proximos a vencer tambien se vean.

    Devuelve dict con: vencidos, vencen_hoy, proximos_3d y los conteos.
    """
    vacio = pd.DataFrame()
    if df_completo is None or df_completo.empty:
        return {
            "vencidos": vacio, "vencen_hoy": vacio, "proximos_3d": vacio,
            "n_vencidos": 0, "n_vencen_hoy": 0, "n_proximos_3d": 0,
        }

    df = df_completo
    abiertos = df[~df["CERRADO"]] if "CERRADO" in df.columns else df
    venc = pd.to_datetime(abiertos["FECHA_VENCIMIENTO"], errors="coerce")

    # Vencidos sin cerrar: el mas atrasado primero.
    vencidos = (
        abiertos[abiertos["ESTADO"] == ROJO]
        .sort_values("HORAS_VENCIDO", ascending=False)
        .reset_index(drop=True)
    )

    # Proximos a vencer: aun no vencen. Hoy = siguientes 24 h.
    futuros = abiertos[
        venc.notna()
        & (venc >= pd.Timestamp(ahora))
        & (venc < pd.Timestamp(ahora) + pd.Timedelta(hours=24))
    ].sort_values("HORAS_RESTANTES")

    tres_dias = abiertos[
        venc.notna()
        & (venc >= pd.Timestamp(ahora))
        & (venc < pd.Timestamp(ahora) + pd.Timedelta(days=3))
    ].sort_values("HORAS_RESTANTES")

    return {
        "vencidos": vencidos.reset_index(drop=True),
        "vencen_hoy": futuros.reset_index(drop=True),
        "proximos_3d": tres_dias.reset_index(drop=True),
        "n_vencidos": len(vencidos),
        "n_vencen_hoy": len(futuros),
        "n_proximos_3d": len(tres_dias),
    }
