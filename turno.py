"""
turno.py - Calculos del turno de la torre de control.

Que resuelve
------------
Lo que la operadora necesita ver y reportar todos los dias:

  * AVANCE DE HOY ..... que entro, que se cerro (a tiempo o tarde), que queda.
  * ACUMULADO DEL MES . como va el mes, y cuanto del incumplimiento es
                         realmente atribuible al tecnico (no a la logistica).
  * GRUPOS DE ALERTA .. naranja (< 1 h), amarillo (1 a 4 h), rojo (vencidos) y
                         proximos 3 dias, EN ESE ORDEN: primero lo que todavia
                         se puede salvar, y de ultimo lo que ya no depende de
                         una accion de hoy.
  * CIERRE DEL DIA ..... el mensaje para cada coordinacion (Bogota y
                         Regionales), con el acumulado por tecnico.

Todo aqui son funciones puras sobre el DataFrame que ya calcula core.py: no lee
archivos ni toca la interfaz, para poder probarlo con datos reales.

Uso:
    import turno
    turno.avance_y_acumulado(df_completo, corte)
    turno.grupos_alerta(df_completo)
    turno.mensaje_cierre("BOGOTA", resumen, por_tecnico, vencidos, "24/09/2026")
"""
from __future__ import annotations

import pandas as pd

import core

# ---------------------------------------------------------------------------
# Grupos de alerta: orden de atencion pedido por la torre
# ---------------------------------------------------------------------------
#
# OJO CON EL ORDEN: no es el orden de gravedad del semaforo. Es el orden de
# ACCION: primero lo que se puede salvar (naranja y amarillo), despues lo que ya
# se vencio (que ya no depende del turno de hoy).

GRUPO_INMINENTE = "inm1h"
GRUPO_PREVENTIVO = "prev4h"
GRUPO_VENCIDOS = "vencidos"
GRUPO_PROXIMOS = "prox3d"
GRUPO_MAS_ADELANTE = "mas_adelante"
GRUPO_SIN_FECHA = "sin_fecha"

ORDEN_GRUPOS = (GRUPO_INMINENTE, GRUPO_PREVENTIVO, GRUPO_VENCIDOS, GRUPO_PROXIMOS)

# Grupos que NO son alerta del turno, pero se cuentan para que la suma de TODOS
# los casos abiertos cuadre y ninguno quede invisible.
SIN_ALERTA = (GRUPO_MAS_ADELANTE, GRUPO_SIN_FECHA)

ETIQUETA_GRUPO = {
    GRUPO_INMINENTE: "Se vencen en menos de 1 hora",
    GRUPO_PREVENTIVO: "Se vencen entre 1 y 4 horas",
    GRUPO_VENCIDOS: "Ya vencidos",
    GRUPO_PROXIMOS: "Próximos 3 días",
    GRUPO_MAS_ADELANTE: "Más adelante (a más de 3 días)",
    GRUPO_SIN_FECHA: "Sin fecha de vencimiento",
}

AYUDA_GRUPO = {
    GRUPO_INMINENTE: "Emergencia del turno: se atienden primero.",
    GRUPO_PREVENTIVO: "Todavía hay margen para gestionarlos.",
    GRUPO_VENCIDOS: "El más atrasado primero: ya no dependen de una acción de hoy.",
    GRUPO_PROXIMOS: "Aún se pueden salvar, pero sin urgencia del turno.",
    GRUPO_MAS_ADELANTE: "Aún les falta más de 3 días: no son alerta del turno.",
    GRUPO_SIN_FECHA: "La plantilla no trae fecha: no se pueden programar.",
}

HORAS_INMINENTE = 1.0
HORAS_PREVENTIVO = 4.0
HORAS_PROXIMOS_DIAS = 72.0


def _utiles(df: pd.DataFrame) -> pd.DataFrame:
    """
    Casos utilizables para los conteos: SIN las filas repetidas.

    Un caso con el N° DE CASO duplicado se marca DUPLICADO y no se cuenta (ver
    core.marcar_duplicados). Si el DataFrame no trae la marca, se devuelve igual.
    """
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    if core.COL_ES_DUPLICADO in df.columns:
        return df[~df[core.COL_ES_DUPLICADO].fillna(False).astype(bool)]
    return df


# ---------------------------------------------------------------------------
# Avance del dia y acumulado del mes
# ---------------------------------------------------------------------------

def avance_y_acumulado(
    df_completo: pd.DataFrame,
    corte,
    culpa: dict | None = None,
) -> dict:
    """
    Los dos bloques que la operadora necesita: hoy y el mes.

    Args:
        df_completo: casos con las metricas de core.calcular_tablero().
        corte:       fecha del corte (date o datetime). El "hoy" se cuenta por
                     dia: creados y cerrados en esa fecha.
        culpa:       mapa opcional ``{caso: culpa}`` del Plan de Trabajo. Con el
                     se calcula el % ATRIBUIBLE al tecnico (lo que no se puede
                     cargar a la logistica ni al aliado ni al banco).

    Returns:
        ``{"hoy": {...}, "mes": {...}}``
    """
    datos = _utiles(df_completo)
    if datos.empty:
        return {
            "hoy": {"recibidos": 0, "cerrados": 0, "a_tiempo": 0, "tarde": 0,
                    "pendientes": 0, "vencidos": 0},
            "mes": {"asignados": 0, "cerrados": 0, "tarde": 0, "pct": 0.0,
                    "documentados": 0, "atribuibles": 0, "pct_atribuible": None},
        }

    corte = pd.Timestamp(corte).date()
    abiertos = datos[~datos["CERRADO"]]
    creados_hoy = datos["FECHA_CREACION"].dt.date == corte
    cerrados_hoy = datos[datos["FECHA_RESOLUCION"].dt.date == corte]

    cerrados = int(datos["CERRADO"].sum())
    tarde = int((datos["ESTADO"] == core.CERRADO_TARDE).sum())

    # --- % atribuible al tecnico (sobre los casos CON causa documentada) ----
    documentados = atribuibles = 0
    pct_atribuible = None
    if culpa:
        claves = datos[core.COL_CASO].map(core.clave_caso)
        causas = claves.map(culpa).fillna("")
        tarde_con_causa = datos[(datos["ESTADO"] == core.CERRADO_TARDE)
                                & causas.ne("")]
        documentados = int(len(tarde_con_causa))
        atribuibles = int(
            causas[tarde_con_causa.index].str.contains("TECNICO").sum()
        )
        if documentados:
            pct_atribuible = round(atribuibles / documentados * 100, 1)

    return {
        "hoy": {
            "recibidos": int(creados_hoy.sum()),
            "cerrados": int(len(cerrados_hoy)),
            "a_tiempo": int((cerrados_hoy["ESTADO"] == core.CERRADO_OK).sum()),
            "tarde": int((cerrados_hoy["ESTADO"] == core.CERRADO_TARDE).sum()),
            "pendientes": int(len(abiertos)),
            "vencidos": int((abiertos["ESTADO"] == core.ROJO).sum()),
        },
        "mes": {
            "asignados": int(len(datos)),
            "cerrados": cerrados,
            "tarde": tarde,
            "pct": round((cerrados - tarde) / cerrados * 100, 1) if cerrados else 0.0,
            "documentados": documentados,
            "atribuibles": atribuibles,
            "pct_atribuible": pct_atribuible,
        },
    }


# ---------------------------------------------------------------------------
# Grupos de alerta
# ---------------------------------------------------------------------------

def grupos_alerta(df_completo: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Los 4 grupos, YA ORDENADOS para atender: naranja, amarillo, rojo, 3 dias.

    Se calculan sobre TODOS los casos abiertos del archivo (no sobre una ventana
    de fechas), para que no se escape ninguno. Devuelve tambien dos grupos que NO
    son alerta pero se cuentan, para que la suma cuadre con los abiertos:
    ``mas_adelante`` (a más de 3 días) y ``sin_fecha`` (sin fecha en la plantilla).
    """
    vacio = pd.DataFrame()
    datos = _utiles(df_completo)
    if datos is None or datos.empty:
        return {clave: vacio for clave in (*ORDEN_GRUPOS, *SIN_ALERTA)}

    abiertos = datos[~datos["CERRADO"]].copy()
    restantes = abiertos["HORAS_RESTANTES"]

    return {
        GRUPO_INMINENTE: abiertos[
            (restantes > 0) & (restantes <= HORAS_INMINENTE)
        ].sort_values("HORAS_RESTANTES"),
        GRUPO_PREVENTIVO: abiertos[
            (restantes > HORAS_INMINENTE) & (restantes <= HORAS_PREVENTIVO)
        ].sort_values("HORAS_RESTANTES"),
        GRUPO_VENCIDOS: abiertos[abiertos["ESTADO"] == core.ROJO].sort_values(
            "HORAS_VENCIDO", ascending=False
        ),
        GRUPO_PROXIMOS: abiertos[
            (restantes > HORAS_PREVENTIVO) & (restantes <= HORAS_PROXIMOS_DIAS)
        ].sort_values("HORAS_RESTANTES"),
        GRUPO_MAS_ADELANTE: abiertos[restantes > HORAS_PROXIMOS_DIAS].sort_values(
            "HORAS_RESTANTES"
        ),
        GRUPO_SIN_FECHA: abiertos[restantes.isna()],
    }


# ---------------------------------------------------------------------------
# Acumulado por tecnico
# ---------------------------------------------------------------------------

def por_tecnico(
    df_completo: pd.DataFrame,
    culpa: dict | None = None,
) -> pd.DataFrame:
    """
    Acumulado del mes por tecnico, con el peor cumplimiento primero.

    Columnas: TECNICO, CASOS, CERRADOS, TARDE, CUMPLIMIENTO_PCT, ATRIBUIBLES,
    CUMPLIMIENTO_ATRIBUIBLE_PCT (esta ultima solo si se pasa ``culpa``).

    Los nombres se unifican con la grafia del diccionario: el mismo tecnico
    escrito de dos formas sale UNA sola vez.
    """
    columnas = ["TECNICO", "CASOS", "CERRADOS", "TARDE", "CUMPLIMIENTO_PCT"]
    datos = _utiles(df_completo)
    if datos is None or datos.empty:
        return pd.DataFrame(columns=columnas)

    tabla = datos.copy()
    tabla["TECNICO"] = tabla["TECNICO"].map(core.nombre_canonico_tecnico)
    if culpa:
        tabla["CULPA"] = tabla[core.COL_CASO].map(core.clave_caso).map(culpa).fillna("")

    filas = []
    for tecnico, grupo in tabla.groupby("TECNICO", dropna=False):
        cerrados = int(grupo["CERRADO"].sum())
        tarde = int((grupo["ESTADO"] == core.CERRADO_TARDE).sum())
        fila = {
            "TECNICO": tecnico if tecnico else "(sin tecnico)",
            "CASOS": int(len(grupo)),
            "CERRADOS": cerrados,
            "TARDE": tarde,
            "CUMPLIMIENTO_PCT": (
                round((cerrados - tarde) / cerrados * 100, 1) if cerrados else None
            ),
        }
        if culpa:
            con_causa = grupo[(grupo["ESTADO"] == core.CERRADO_TARDE)
                              & (grupo["CULPA"].ne(""))]
            atribuibles = int(con_causa["CULPA"].str.contains("TECNICO").sum())
            fila["ATRIBUIBLES"] = atribuibles
            fila["CUMPLIMIENTO_ATRIBUIBLE_PCT"] = (
                round((len(con_causa) - atribuibles) / len(con_causa) * 100, 1)
                if len(con_causa) else None
            )
        filas.append(fila)

    tabla = pd.DataFrame(filas)
    return (
        tabla.sort_values(["CUMPLIMIENTO_PCT", "CASOS"], ascending=[True, False],
                          na_position="last")
        .reset_index(drop=True)
    )


# ---------------------------------------------------------------------------
# Cierre del dia (mensaje para cada coordinacion)
# ---------------------------------------------------------------------------

def culpa_por_caso(plan_bytes) -> dict:
    """
    Mapa ``{caso: culpa}`` tomado de la hoja Casos_Ven del Plan de Trabajo.

    La culpa (TECNICO / LOGISTICO / ALIADO / BANCO / ACTIVOS / MESA-COLSOF) es lo
    que permite no cargarle al tecnico un vencimiento que no controla.
    """
    import plan

    try:
        casos = plan.vista_vencidos(plan_bytes)["casos"]
    except Exception:
        return {}
    if casos is None or casos.empty or plan.COL_CULPA not in casos.columns:
        return {}
    return dict(zip(
        casos[plan.COL_CASO].map(core.clave_caso),
        casos[plan.COL_CULPA].fillna("").astype(str),
    ))


def mensaje_cierre(
    nombre: str,
    resumen: dict,
    tecnicos: pd.DataFrame,
    vencidos: pd.DataFrame,
    fecha: str,
    *,
    tope_tecnicos: int = 8,
    tope_vencidos: int = 5,
) -> str:
    """
    Mensaje del cierre del dia para UNA coordinacion (texto plano, con barras).

    Se envia por Telegram con el mismo mecanismo que ya existe y con el filtro
    por region (--filtro bogota | regionales), para que cada coordinadora reciba
    solo lo suyo.
    """
    hoy, mes = resumen["hoy"], resumen["mes"]
    lineas = [
        f"🏁 CIERRE DEL DÍA — {nombre}",
        f"📅 {fecha} · corte 18:00",
        "",
        "✅ AVANCE DE HOY",
        f"   Recibidos hoy ....... {hoy['recibidos']}",
        f"   Cerrados hoy ........ {hoy['cerrados']}   "
        f"({hoy['a_tiempo']} a tiempo · {hoy['tarde']} tarde)",
        f"   Pendientes .......... {hoy['pendientes']}",
        f"   Vencidos abiertos .... {hoy['vencidos']}",
        "",
        "📈 ACUMULADO DEL MES",
        f"   Asignados ........... {mes['asignados']}",
        f"   Cerrados ............ {mes['cerrados']}   "
        f"({mes['tarde']} con incumplimiento)",
        f"   Cumplimiento ........ {mes['pct']} %",
    ]
    if mes.get("pct_atribuible") is not None:
        cobertura = (mes["documentados"] / mes["tarde"]) if mes["tarde"] else 0.0
        lineas.append(
            f"   Atribuible al técnico  {mes['pct_atribuible']} %   "
            f"(sobre {mes['documentados']} casos con causa escrita)"
        )
        if cobertura < 0.60:
            # Sin causa escrita no se puede repartir la culpa: se avisa en vez de
            # dar un porcentaje que parezca completo.
            lineas.append(
                f"   ⚠️ El Plan de Trabajo solo documenta {mes['documentados']} de "
                f"{mes['tarde']} casos cerrados tarde: el resto no tiene causa escrita."
            )

    if tecnicos is not None and not tecnicos.empty:
        lineas += ["", "👷 ACUMULADO POR TÉCNICO"]
        for _, fila in tecnicos.head(tope_tecnicos).iterrows():
            pct = fila["CUMPLIMIENTO_PCT"]
            pct = 0.0 if pct is None or pd.isna(pct) else float(pct)
            llenos = int(round(pct / 10))
            lineas.append(
                f"   {str(fila['TECNICO'])[:22]:24s}{int(fila['CASOS']):3d}  "
                f"{'█' * llenos}{'░' * (10 - llenos)} {pct:.0f}%"
            )

    if vencidos is not None and not vencidos.empty:
        lineas += ["", f"⚠️ VENCIDOS ABIERTOS AL CIERRE ({len(vencidos)})"]
        for _, fila in vencidos.head(tope_vencidos).iterrows():
            signo = fila.get("TIEMPO_VENCIDO", "—")
            ciudad = str(fila.get(core.COL_CIUDAD, "") or "")[:18]
            lineas.append(f"   🔴 {fila.get(core.COL_CASO, 'S/N')} · {ciudad} · vencido {signo}")

    lineas += [
        "",
        "ℹ️ 'Cerrados hoy' puede superar a 'recibidos hoy': se cerraron casos",
        "   de días anteriores.",
    ]
    return "\n".join(lineas)
