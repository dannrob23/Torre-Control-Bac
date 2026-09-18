"""
formatos.py - Estilos de presentacion de los casos para Telegram y WhatsApp.

IMPORTANTE (limitacion real de las plataformas):
    Ni Telegram ni WhatsApp renderizan tablas markdown (| a | b |). En Telegram
    se ven como texto con barras y sin alineacion; en WhatsApp, igual.
    Lo unico que se alinea de verdad es un bloque monoespaciado (<pre>).

    Por eso hay varios estilos: cada uno resuelve un caso de uso distinto.

Estilos disponibles:

    tabla    - Bloque <pre> alineado (por defecto). Ideal para COPIAR y pegar en
               WhatsApp conservando columnas. Sobrio, informacion densa.

    tarjetas - Una tarjeta por caso con emojis y negrita, dentro de una cita
               expandible (<blockquote expandable>). Lo mas LEGIBLE en el movil:
               se ve el titulo y se expande para ver el detalle.

    por-estado - Agrupa por semaforo: primero los vencidos, luego naranjas,
               amarillos. Cada grupo con su encabezado y contador. Ideal para
               PRIORIZAR de un vistazo.

    markdown - Tabla markdown con tuberias. Se ve sin alinear en Telegram/WhatsApp,
               pero sirve para pegar en Excel, Word, correo o Notion, donde SI
               se convierte en tabla real.

    resumen  - Solo los totales por estado y por region, sin detalle de casos.
               Para un vistazo rapido o para el canal de jefatura.
"""

from __future__ import annotations

from datetime import datetime

import pandas as pd

# Estilos validos (usados por --estilo en la linea de comandos)
ESTILO_TABLA = "tabla"
ESTILO_TARJETAS = "tarjetas"
ESTILO_POR_ESTADO = "por-estado"
ESTILO_MARKDOWN = "markdown"
ESTILO_RESUMEN = "resumen"

ESTILOS = (
    ESTILO_TABLA, ESTILO_TARJETAS, ESTILO_POR_ESTADO, ESTILO_MARKDOWN, ESTILO_RESUMEN
)

# Orden de presentacion por gravedad.
ORDEN_PRESENTACION = [
    "ROJO", "CERRADO TARDE", "NARANJA", "AMARILLO", "VERDE",
    "SIN VENCIMIENTO", "CERRADO OK",
]

ICONO = {
    "ROJO": "🔴",
    "CERRADO TARDE": "⛔",
    "NARANJA": "🟠",
    "AMARILLO": "🟡",
    "VERDE": "🟢",
    "SIN VENCIMIENTO": "⚪",
    "CERRADO OK": "✅",
}

# Texto corto de cada estado, para encabezados de grupo.
TITULO_ESTADO = {
    "ROJO": "VENCIDOS - accion inmediata",
    "CERRADO TARDE": "CERRADOS CON INCUMPLIMIENTO",
    "NARANJA": "CRITICOS - menos de 1 hora",
    "AMARILLO": "EN ALERTA - menos de 4 horas",
    "VERDE": "A TIEMPO",
    "SIN VENCIMIENTO": "SIN FECHA DE VENCIMIENTO",
    "CERRADO OK": "CERRADOS A TIEMPO",
}


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _escapar(texto) -> str:
    """Escapa los caracteres que rompen el parseo HTML de Telegram."""
    return (
        str(texto)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _fecha(valor) -> str:
    if valor is None or pd.isna(valor):
        return "sin fecha"
    try:
        return pd.Timestamp(valor).strftime("%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return "sin fecha"


def _fecha_corta(valor) -> str:
    """Fecha compacta 'dd/mm HH:MM', para la tabla alineada."""
    if valor is None or pd.isna(valor):
        return "—"
    try:
        return pd.Timestamp(valor).strftime("%d/%m %H:%M")
    except (ValueError, TypeError):
        return "—"


def _tiempo(fila) -> str:
    """Texto del tiempo relevante segun el estado."""
    estado = str(fila.get("ESTADO", ""))
    if estado in ("ROJO", "CERRADO TARDE"):
        vencido = fila.get("TIEMPO_VENCIDO")
        if vencido and str(vencido) != "—":
            return f"VENCIDO hace {vencido}"
        horas = fila.get("HORAS_VENCIDO")
        if pd.notna(horas):
            return f"VENCIDO hace {int(round(float(horas)))} h"
        return "VENCIDO"
    return str(fila.get("TIEMPO_RESTANTE", "—"))


def _tecnico(fila) -> str:
    return str(fila.get("TECNICO", "") or "(sin tecnico)")


def _region(fila) -> str:
    region = str(fila.get("REGION_TECNICO", "") or "")
    return region if region and region != "SIN REGION" else "(sin region)"


def _caso(fila) -> str:
    return str(fila.get("N° DE CASO", "S/N")).strip()


def _conteo(df: pd.DataFrame) -> dict[str, int]:
    if df is None or df.empty:
        return {}
    return df["ESTADO"].value_counts().to_dict()


def _linea_resumen(df: pd.DataFrame) -> str:
    """Resumen compacto de totales por estado, con emojis."""
    conteo = _conteo(df)
    partes = [
        f"{ICONO.get(e, '')} {conteo[e]}"
        for e in ORDEN_PRESENTACION
        if conteo.get(e)
    ]
    return "  ".join(partes) if partes else "Sin casos"


def _cabecera(titulo: str, df: pd.DataFrame, con_resumen: bool = True) -> list[str]:
    lineas = [
        f"<b>🛰️ {_escapar(titulo)}</b>",
        f"<i>Colsof / Banco Agrario — {datetime.now():%d/%m/%Y %H:%M}</i>",
    ]
    if con_resumen:
        lineas.append(f"<b>{_linea_resumen(df)}</b>")
    lineas.append(f"<i>Total: {len(df)} caso(s)</i>" if df is not None
                  else "<i>Total: 0</i>")
    return lineas


# ---------------------------------------------------------------------------
# Estilo: tarjetas (el mas legible en el movil)
# ---------------------------------------------------------------------------

def formato_tarjetas(df: pd.DataFrame, titulo: str = "CASOS ACTIVOS",
                     max_visibles: int = 6) -> str:
    """
    Una tarjeta por caso dentro de una cita expandible.

    Los primeros `max_visibles` casos van visibles; el resto dentro de una cita
    colapsada ("expandable") que el usuario abre con un toque. Asi el mensaje no
    se vuelve un muro de texto aunque haya 40 casos.
    """
    if df is None or df.empty:
        return "<b>Sin casos en la ventana.</b>"

    lineas = _cabecera(titulo, df)
    lineas.append("")

    trabajo = df.reset_index(drop=True)
    visibles = trabajo.head(max_visibles)
    resto = trabajo.iloc[max_visibles:]

    for _, fila in visibles.iterrows():
        lineas.extend(_tarjeta(fila, con_cita=False))

    if not resto.empty:
        lineas.append(
            f"<b>➕ {len(resto)} caso(s) mas</b> — toque para ver el detalle"
        )
        cuerpo = []
        for _, fila in resto.iterrows():
            cuerpo.extend(_tarjeta(fila, con_cita=False))
        lineas.append("<blockquote expandable>" + "\n".join(cuerpo) + "</blockquote>")

    return "\n".join(lineas)


def _tarjeta(fila, con_cita: bool = True) -> list[str]:
    """Bloque de un caso con emojis y negrita."""
    estado = str(fila.get("ESTADO", ""))
    icono = ICONO.get(estado, "•")
    lineas = [
        f"{icono} <b>{_escapar(_caso(fila))}</b> · {_escapar(estado)}",
        f"   ⏳ {_escapar(_tiempo(fila))}",
        f"   🗓️ Vence {_escapar(_fecha(fila.get('FECHA_VENCIMIENTO')))}",
        f"   👷 {_escapar(_tecnico(fila))}",
        f"   📍 {_escapar(_region(fila))}",
        "",
    ]
    if con_cita:
        return ["<blockquote>" + "\n".join(lineas).rstrip() + "</blockquote>"]
    return lineas


# ---------------------------------------------------------------------------
# Estilo: por estado (para priorizar)
# ---------------------------------------------------------------------------

def formato_por_estado(df: pd.DataFrame, titulo: str = "CASOS POR PRIORIDAD") -> str:
    """
    Agrupa los casos por semaforo, del mas grave al menos grave.

    Cada grupo tiene su encabezado con contador, y los casos van en lista
    compacta: caso, tiempo, tecnico. Es el formato para decidir por donde
    empezar.
    """
    if df is None or df.empty:
        return "<b>Sin casos en la ventana.</b>"

    lineas = _cabecera(titulo, df)
    conteo = _conteo(df)

    for estado in ORDEN_PRESENTACION:
        grupo = df[df["ESTADO"] == estado]
        if grupo.empty:
            continue

        icono = ICONO.get(estado, "•")
        titulo_grupo = TITULO_ESTADO.get(estado, estado)
        lineas.extend([
            "",
            f"{icono} <b>{_escapar(titulo_grupo)}</b> — {len(grupo)} caso(s)",
        ])

        for _, fila in grupo.iterrows():
            lineas.append(
                f"   • <b>{_escapar(_caso(fila))}</b> · {_escapar(_tiempo(fila))}\n"
                f"     👷 {_escapar(_tecnico(fila))} · 📍 {_escapar(_region(fila))}"
            )

    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Estilo: tabla alineada (la actual, ideal para copiar a WhatsApp)
# ---------------------------------------------------------------------------

def formato_tabla(df: pd.DataFrame, titulo: str = "CASOS ACTIVOS") -> str:
    """
    Tabla alineada dentro de <pre>.

    Es la unica forma de que las columnas queden alineadas en Telegram, y al
    copiarla a WhatsApp conserva la alineacion porque el texto ya viene con los
    espacios calculados.
    """
    if df is None or df.empty:
        return "<b>Sin casos en la ventana.</b>"

    encabezados = ["ESTADO", "CASO", "VENCE", "TIEMPO", "REGION", "TECNICO"]
    filas = []
    for _, f in df.iterrows():
        filas.append([
            _recortar(str(f.get("ESTADO", "")), 14),
            _recortar(_caso(f), 13),
            _recortar(_fecha_corta(f.get("FECHA_VENCIMIENTO")), 12),
            _recortar(_tiempo(f), 18),
            _recortar(_region(f), 14),
            _recortar(_tecnico(f), 30),
        ])

    anchos = [
        max(len(encabezados[i]), max((len(f[i]) for f in filas), default=0))
        for i in range(len(encabezados))
    ]

    def linea(vals):
        return "  ".join(v.ljust(anchos[i]) for i, v in enumerate(vals)).rstrip()

    cuerpo = [linea(encabezados), "  ".join("-" * a for a in anchos)]
    cuerpo += [linea(f) for f in filas]

    cabecera = _cabecera(titulo, df)
    return "\n".join(cabecera) + "\n\n<pre>" + "\n".join(cuerpo) + "</pre>"


def _recortar(texto: str, ancho: int) -> str:
    """Recorta a `ancho` caracteres agregando elipsis (…) si es necesario."""
    texto = str(texto)
    return texto if len(texto) <= ancho else texto[: max(ancho - 1, 1)] + "…"


# ---------------------------------------------------------------------------
# Estilo: markdown (para Excel / Word / correo)
# ---------------------------------------------------------------------------

def formato_markdown(df: pd.DataFrame, titulo: str = "CASOS ACTIVOS") -> str:
    """
    Tabla markdown con tuberias.

    Aviso: en Telegram y WhatsApp se ve como texto con barras SIN alinear. Este
    formato existe para pegarlo en Excel, Word, correo o Notion, donde la tabla
    se reconstruye de verdad.
    """
    if df is None or df.empty:
        return f"**{titulo}**\n\nSin casos."

    lineas = [
        f"**{titulo}**",
        f"Colsof / Banco Agrario — {datetime.now():%d/%m/%Y %H:%M}",
        "",
        "| Estado | Caso | Vencimiento | Tiempo | Tecnico | Region |",
        "|---|---|---|---|---|---|",
    ]
    for _, f in df.iterrows():
        estado = str(f.get("ESTADO", ""))
        lineas.append(
            f"| {ICONO.get(estado, '')} {estado} | {_caso(f)} | "
            f"{_fecha(f.get('FECHA_VENCIMIENTO'))} | {_tiempo(f)} | "
            f"{_tecnico(f)} | {_region(f)} |"
        )
    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Estilo: resumen ejecutivo
# ---------------------------------------------------------------------------

def formato_resumen(df: pd.DataFrame, titulo: str = "RESUMEN DE CASOS") -> str:
    """
    Solo totales: por estado y por region. Sin detalle de casos.

    Util para el canal de jefatura o como mensaje de apertura antes del detalle.
    """
    if df is None or df.empty:
        return "<b>Sin casos en la ventana.</b>"

    lineas = [
        f"<b>🛰️ {_escapar(titulo)}</b>",
        f"<i>Colsof / Banco Agrario — {datetime.now():%d/%m/%Y %H:%M}</i>",
        "",
        f"<b>📊 Total: {len(df)} caso(s)</b>",
        "",
        "<b>Por estado</b>",
    ]
    conteo = _conteo(df)
    for estado in ORDEN_PRESENTACION:
        if conteo.get(estado):
            lineas.append(
                f"   {ICONO.get(estado, '•')} {_escapar(estado)}: <b>{conteo[estado]}</b>"
            )

    if "REGION_TECNICO" in df.columns:
        lineas.extend(["", "<b>Por region</b>"])
        por_region = df["REGION_TECNICO"].value_counts()
        for region, n in por_region.items():
            vencidos = int((df[df["REGION_TECNICO"] == region]["ESTADO"] == "ROJO").sum())
            extra = f" ({vencidos} vencido{'s' if vencidos != 1 else ''})" if vencidos else ""
            lineas.append(f"   📍 {_escapar(region)}: <b>{n}</b>{extra}")

    if "TECNICO" in df.columns:
        tecnicos = df["TECNICO"].replace("", pd.NA).dropna().nunique()
        lineas.extend(["", f"<b>👷 Tecnicos con casos: {tecnicos}</b>"])

    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Selector
# ---------------------------------------------------------------------------

def generar(df: pd.DataFrame, estilo: str = ESTILO_TABLA,
            titulo: str = "CASOS ACTIVOS", **kwargs) -> str:
    """
    Genera el texto segun el estilo pedido.

    Args:
        df:     casos a presentar
        estilo: uno de ESTILOS
        titulo: encabezado del mensaje
        kwargs: parametros extra del estilo (ej. max_visibles en tarjetas)
    """
    estilo = (estilo or ESTILO_TABLA).strip().lower()
    if estilo == ESTILO_TARJETAS:
        return formato_tarjetas(df, titulo, **kwargs)
    if estilo == ESTILO_POR_ESTADO:
        return formato_por_estado(df, titulo)
    if estilo == ESTILO_MARKDOWN:
        return formato_markdown(df, titulo)
    if estilo == ESTILO_RESUMEN:
        return formato_resumen(df, titulo)
    return formato_tabla(df, titulo)


if __name__ == "__main__":
    # Muestra todos los estilos:  python formatos.py
    import warnings

    warnings.filterwarnings("ignore")
    from core import calcular_tablero

    casos = calcular_tablero().requieren_atencion

    for estilo in ESTILOS:
        print("=" * 78)
        print(f" ESTILO: {estilo.upper()}")
        print("=" * 78)
        print(generar(casos.head(4), estilo, "CASOS QUE REQUIEREN ATENCION"))
        print()
