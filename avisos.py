"""
avisos.py - Composicion de avisos de vencimiento para TECNICOS.

Genera el mensaje individual que la persona encargada de la torre de control envia
a cada tecnico con casos pendientes. El mismo texto sirve para:

  - Telegram (con menciones @usuario cuando se conocen)
  - WhatsApp (se copia y pega conservando formato y emojis)

Ejemplo de salida:

    ⚠️⏳🚨 AVISO DE VENCIMIENTO PROXIMO - CASO QT3329957 🚨⏳⚠️

    Hola JONATHAN FELIPE JARABA CHALARCA 👋
    Te informamos que el caso asignado a ti esta proximo a vencer.

    📌 Caso: QT3329957
    🗓️ Fecha de vencimiento: 17/09/2026 15:20
    ⏳ Tiempo restante: 45 min
    🏢 Oficina: 1300-REGIONAL ANTIOQUIA-ANTIOQUIA
    🗺️ Region: MEDELLIN
    🔴 Estado: NARANJA

    ❗ Actualiza el caso antes de la fecha limite para evitar
    incumplimientos en los tiempos de atencion. ⏳
"""

from __future__ import annotations

import html as _html
import json
import os
from datetime import datetime

import pandas as pd

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_MENCIONES = os.path.join(DIR_SCRIPT, "menciones.json")
ARCHIVO_MENCIONES_EJEMPLO = os.path.join(DIR_SCRIPT, "menciones.example.json")

# ---------------------------------------------------------------------------
# Plantilla del aviso
# ---------------------------------------------------------------------------

CABECERA = "⚠️⏳🚨 AVISO DE VENCIMIENTO PROXIMO - CASO {caso} 🚨⏳⚠️"
SALUDO = "Hola {mencion} 👋"
INTRO = "Te informamos que el caso asignado a ti esta proximo a vencer."
CIERRE = (
    "❗ Actualiza el caso antes de la fecha limite para evitar "
    "incumplimientos en los tiempos de atencion. ⏳"
)

# Emoji e instructivo segun el estado del caso.
GUIA_POR_ESTADO = {
    "ROJO": ("🔴", "El caso YA ESTA VENCIDO. Atiendelo de inmediato. 🚨"),
    "CERRADO TARDE": ("⛔", "El caso se cerro fuera del tiempo. Registra la causa. 📝"),
    "NARANJA": ("🟠", "Queda menos de 1 hora. Es prioritario. 🔥"),
    "AMARILLO": ("🟡", "Quedan menos de 4 horas. Preparate para atenderlo. ⏰"),
    "VERDE": ("🟢", "Aun tienes tiempo, pero no lo dejes para el final. ✅"),
    "SIN VENCIMIENTO": ("⚪", "El caso no tiene fecha de vencimiento registrada. ⚠️"),
    "CERRADO OK": ("✅", "Caso cerrado a tiempo. 👏"),
}


# ---------------------------------------------------------------------------
# Menciones: tecnico -> @usuario (Telegram) / telefono (WhatsApp)
# ---------------------------------------------------------------------------

def cargar_menciones(ruta: str | None = None) -> dict[str, dict[str, str]]:
    """
    Lee el mapeo de menciones por tecnico.

    Estructura de menciones.json:
        {
          "JONATHAN FELIPE JARABA CHALARCA": {
              "telegram": "@jonathan",
              "whatsapp": "+573001234567",
              "chat_id": "123456789",
              "nombre_corto": "Jonathan"
          }
        }

    'chat_id' es opcional: permite enviar el aviso en PRIVADO al tecnico (requiere
    que haya pulsado /start una vez). Si se omite, el aviso va al grupo.

    Devuelve un dict indexado por nombre NORMALIZADO del tecnico.
    Si no existe el archivo, devuelve un dict vacio (el aviso se genera igual).
    """
    ruta = ruta or ARCHIVO_MENCIONES
    if not os.path.isfile(ruta):
        return {}
    try:
        with open(ruta, encoding="utf-8-sig") as fh:
            datos = json.load(fh)
    except (OSError, json.JSONDecodeError):
        return {}

    salida: dict[str, dict[str, str]] = {}
    for nombre, info in (datos or {}).items():
        if str(nombre).startswith("_"):
            continue  # claves de comentario
        if isinstance(info, dict):
            salida[_clave_tecnico(nombre)] = {
                "telegram": str(info.get("telegram", "") or "").strip(),
                "whatsapp": str(info.get("whatsapp", "") or "").strip(),
                "nombre_corto": str(info.get("nombre_corto", "") or "").strip(),
                # chat_id propio del tecnico: permite enviarle el aviso en PRIVADO
                # (requiere que haya pulsado /start una vez). Vacio = se usa el grupo.
                "chat_id": str(info.get("chat_id", "") or "").strip(),
            }
        elif isinstance(info, str):
            # Forma corta: "TECNICO": "@usuario"
            salida[_clave_tecnico(nombre)] = {
                "telegram": info.strip(), "whatsapp": "", "nombre_corto": "",
                "chat_id": "",
            }
    return salida


def _clave_tecnico(nombre: str) -> str:
    """
    Clave normalizada de un tecnico, aplicando los alias de core.py.

    Reutiliza ALIAS_TECNICOS para que menciones.json cruce con los nombres tal
    como vienen en la plantilla (que trae variantes como 'JHONATHAN' o un
    apellido faltante).
    """
    from core import ALIAS_TECNICOS, normalizar_texto

    clave = normalizar_texto(nombre)
    alias = {normalizar_texto(k): normalizar_texto(v) for k, v in ALIAS_TECNICOS.items()}
    return alias.get(clave, clave)


def mencion_para(nombre_tecnico: str, canal: str = "whatsapp",
                 menciones: dict | None = None) -> str:
    """
    Devuelve el texto de mencion para un tecnico.

    canal="telegram" -> usa @usuario si existe.
    canal="whatsapp" -> usa +telefono si existe (WhatsApp lo vuelve enlace).
    Si no hay dato, devuelve el nombre completo (nunca queda vacio).

    El nombre corto SIEMPRE se usa como saludo legible; la mencion va aparte para
    que en WhatsApp/Telegram se convierta en notificacion real.
    """
    if menciones is None:
        menciones = cargar_menciones()
    info = menciones.get(_clave_tecnico(nombre_tecnico), {})
    nombre_legible = info.get("nombre_corto") or nombre_tecnico or "tecnico"

    if canal == "telegram":
        usuario = info.get("telegram", "")
        if usuario:
            return f"{nombre_legible} {usuario}"
    elif canal == "whatsapp":
        telefono = info.get("whatsapp", "")
        if telefono:
            return f"{nombre_legible} {telefono}"

    return nombre_legible


# ---------------------------------------------------------------------------
# Datos del caso -> texto
# ---------------------------------------------------------------------------

def _fecha_legible(valor) -> str:
    if valor is None or pd.isna(valor):
        return "sin fecha registrada"
    try:
        return pd.Timestamp(valor).strftime("%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return "sin fecha registrada"


def _oficina_de(fila: pd.Series) -> str:
    """Oficina: prefiere 'OFICINA'; si viene vacia, arma regional-departamento."""
    oficina = str(fila.get("OFICINA", "") or "").strip()
    if oficina and oficina.lower() != "nan":
        return oficina
    partes = [
        str(fila.get(c, "") or "").strip()
        for c in ("REGIONAL", "DEPARTAMENTO")
        if str(fila.get(c, "") or "").strip()
    ]
    return " - ".join(partes) if partes else "(sin oficina)"


def _tiempo_texto(fila: pd.Series) -> str:
    """Tiempo restante o tiempo vencido, segun el estado."""
    estado = str(fila.get("ESTADO", ""))
    if estado in ("ROJO", "CERRADO TARDE"):
        vencido = fila.get("TIEMPO_VENCIDO")
        if vencido and str(vencido) != "—":
            return f"VENCIDO hace {vencido}"
        horas = fila.get("HORAS_VENCIDO")
        if pd.notna(horas):
            from core import formatear_duracion

            return f"VENCIDO hace {formatear_duracion(float(horas))}"
        return "VENCIDO"
    return str(fila.get("TIEMPO_RESTANTE", "—"))


# ---------------------------------------------------------------------------
# Composicion del aviso
# ---------------------------------------------------------------------------

def componer_aviso(fila: pd.Series, canal: str = "whatsapp",
                   menciones: dict | None = None) -> str:
    """
    Genera el texto completo del aviso para UN caso.

    canal: "whatsapp" (por defecto, texto plano) o "telegram" (mismo texto; el
    emoji y los saltos de linea son validos en ambos).
    """
    estado = str(fila.get("ESTADO", ""))
    icono, guia = GUIA_POR_ESTADO.get(estado, ("⚠️", CIERRE))
    caso = str(fila.get("N° DE CASO", "S/N")).strip()
    tecnico = str(fila.get("TECNICO", "") or "").strip()
    tiempo = _tiempo_texto(fila)
    vencimiento = _fecha_legible(fila.get("FECHA_VENCIMIENTO"))
    oficina = _oficina_de(fila)
    region = str(fila.get("REGION_TECNICO", "") or "").strip()

    lineas = [
        CABECERA.format(caso=caso),
        "",
        SALUDO.format(mencion=mencion_para(tecnico, canal, menciones)),
        INTRO,
        "",
        f"📌 Caso: {caso}",
        f"🗓️ Fecha de vencimiento: {vencimiento}",
        f"⏳ Tiempo: {tiempo}",
        f"🏢 Oficina: {oficina}",
    ]
    if region and region != "SIN REGION":
        lineas.append(f"🗺️ Region: {region}")
    lineas.append(f"{icono} Estado: {estado}")
    lineas.append("")
    lineas.append(guia if estado in ("ROJO", "CERRADO TARDE", "NARANJA") else CIERRE)

    return "\n".join(lineas)


def componer_aviso_multiple(df: pd.DataFrame, canal: str = "whatsapp",
                            menciones: dict | None = None) -> str:
    """
    Genera UN solo mensaje con varios casos del mismo tecnico.

    Util cuando un tecnico tiene 5 casos pendientes: mejor un mensaje con lista
    que 5 mensajes separados.
    """
    if df is None or df.empty:
        return ""

    tecnico = str(df.iloc[0].get("TECNICO", "") or "").strip()
    cantidad = len(df)
    momento = datetime.now().strftime("%d/%m/%Y %H:%M")

    lineas = [
        f"⚠️⏳🚨 AVISO DE VENCIMIENTOS PROXIMOS ({cantidad} casos) 🚨⏳⚠️",
        "",
        SALUDO.format(mencion=mencion_para(tecnico, canal, menciones)),
        f"Tienes {cantidad} casos que requieren tu atencion:",
        "",
    ]

    for i, (_, fila) in enumerate(df.iterrows(), start=1):
        estado = str(fila.get("ESTADO", ""))
        icono, _ = GUIA_POR_ESTADO.get(estado, ("⚠️", ""))
        lineas.extend(
            [
                f"{icono} {i}. Caso {str(fila.get('N° DE CASO', 'S/N')).strip()}",
                f"   🗓️ Vence: {_fecha_legible(fila.get('FECHA_VENCIMIENTO'))}",
                f"   ⏳ {_tiempo_texto(fila)}",
                f"   🏢 {_oficina_de(fila)}",
                "",
            ]
        )

    lineas.append(CIERRE)
    lineas.append("")
    lineas.append(f"🕐 Aviso generado: {momento}")
    return "\n".join(lineas)


def componer_aviso_tecnico(tecnico: str, df_tecnico: pd.DataFrame,
                           canal: str = "whatsapp",
                           un_solo_mensaje: bool = True,
                           menciones: dict | None = None) -> str:
    """
    Compone el texto para un tecnico con todos sus casos.

    un_solo_mensaje=True  -> un mensaje con la lista (recomendado para WhatsApp).
    un_solo_mensaje=False -> varios avisos individuales separados.
    """
    if df_tecnico is None or df_tecnico.empty:
        return ""

    if un_solo_mensaje and len(df_tecnico) > 1:
        return componer_aviso_multiple(df_tecnico, canal, menciones)

    separador = "\n\n" + ("─" * 34) + "\n\n"
    avisos = [
        componer_aviso(f, canal, menciones).replace(
            f" - CASO {str(f.get('N° DE CASO', 'S/N')).strip()}", ""
        )
        for _, f in df_tecnico.iterrows()
    ]
    return separador + separador.join(avisos)


# ---------------------------------------------------------------------------
# Aviso para TELEGRAM: HTML con negrillas y orden por urgencia
# ---------------------------------------------------------------------------
# Telegram acepta un subconjunto de HTML: <b>, <i>, <u>, <s>, <code>, <pre>, <a>.
# Se envian con parse_mode="HTML" y por eso TODO dato que venga de la plantilla
# (nombre, oficina, numero de caso) tiene que pasar por _esc(): si trae < > o &,
# el servidor de Telegram rechaza el mensaje completo.

def _esc(valor) -> str:
    """Escapa el texto para que sea seguro dentro de un mensaje HTML."""
    if valor is None:
        return ""
    return _html.escape(str(valor), quote=False)


def _negrita(valor) -> str:
    """Devuelve el texto en negrita y ya escapado."""
    return f"<b>{_esc(valor)}</b>"


# Orden pedido por la torre de control, de mas a menos prioritario:
#   1) PROXIMOS A VENCER  -> el que vence antes, arriba (NARANJA <1h, AMARILLO
#      1-4h, VERDE >4h)
#   2) YA VENCIDOS        -> el mas atrasado, arriba
#   3) SIN VENCIMIENTO    -> no se pueden clasificar
#   4) CERRADOS TARDE     -> incumplimiento ya consumado, solo informativo
GRUPOS_AVISO: tuple[tuple[str, str, tuple[str, ...]], ...] = (
    ("PROXIMO", "⏳ PRÓXIMOS A VENCER", ("NARANJA", "AMARILLO", "VERDE")),
    ("VENCIDO", "🚨 YA VENCIDOS", ("ROJO",)),
    ("SIN_FECHA", "⚪ SIN VENCIMIENTO", ("SIN VENCIMIENTO",)),
    ("CIERRE", "📝 CERRADOS TARDE", ("CERRADO TARDE",)),
)

_ORDEN_GRUPO = {clave: i for i, (clave, _t, _e) in enumerate(GRUPOS_AVISO)}

# ---------------------------------------------------------------------------
# Emojis de alarma y bloque de instrucciones.
#
# El mensaje lo lee el TECNICO en el celular, de pie y con afan, y tambien el
# CONTROLER que vigila la torre. Por eso:
#   - Las alarmas van ARRIBA, bien visibles, para que el mensaje salte a la vista.
#   - Cada caso ocupa 3 lineas cortas (nada de parrafos).
#   - Al final va QUE HACER, con los pasos para gestionar y CERRAR el caso.
# ---------------------------------------------------------------------------
ALARMA_VENCIDOS = "🚨🚨🚨"
ALARMA_PROXIMOS = "⏰⏰⏰"
LINEA = "━" * 20

# Emoji que marca la urgencia de cada caso (va junto al numero de caso).
# OJO: se usan cadenas literales, no las constantes de core, para no crear
# dependencia circular entre avisos.py y core.py.
MARCA_URGENCIA = {
    "ROJO": "❌",
    "CERRADO TARDE": "⛔",
    "NARANJA": "🔥",
    "AMARILLO": "⏰",
    "VERDE": "⏳",
    "SIN VENCIMIENTO": "⚪",
    "CERRADO OK": "✅",
}

# Franja de estado con alarma, para los encabezados de seccion.
ENCABEZADO_ALARMA = {
    "PROXIMO": ("⏰", "⏰"),
    "VENCIDO": ("🚨", "🚨"),
    "SIN_FECHA": ("", ""),
    "CIERRE": ("", ""),
}

QUE_HACER_HTML = "\n".join(
    [
        "🛠️ <b>¿QUÉ HACER CON ESTOS CASOS?</b>",
        "",
        "1️⃣ <b>Atiende</b> el caso en la oficina.",
        "2️⃣ <b>Registra</b> la solución en la herramienta.",
        "3️⃣ <b>Cierra</b> el caso antes de la fecha de vencimiento.",
        "4️⃣ Si <b>ya lo resolviste</b>, ciérralo <b>YA</b>: mientras siga "
        "abierto sigue contando como incumplido.",
        "",
        "🔔 <b>Recuerda:</b> cerrar a tiempo evita incumplir el ANS.",
        "📣 <b>¿Dudas o necesitas apoyo?</b> Escribe a la <b>Torre de Control</b>.",
    ]
)


def grupo_de_estado(estado: str) -> str:
    """Grupo de prioridad al que pertenece un estado del semaforo."""
    for clave, _titulo, estados in GRUPOS_AVISO:
        if estado in estados:
            return clave
    return "OTROS"


def orden_para_aviso(df: pd.DataFrame) -> pd.DataFrame:
    """
    Reordena los casos con la prioridad de la torre:
    primero los PROXIMOS A VENCER (el mas urgente arriba) y luego los VENCIDOS
    (el mas atrasado arriba). No modifica el DataFrame original.
    """
    if df is None or df.empty:
        return df

    trabajo = df.copy()
    restantes = pd.to_numeric(
        trabajo.get("HORAS_RESTANTES"), errors="coerce"
    ) if "HORAS_RESTANTES" in trabajo else pd.Series(pd.NA, index=trabajo.index)
    vencido = pd.to_numeric(
        trabajo.get("HORAS_VENCIDO"), errors="coerce"
    ) if "HORAS_VENCIDO" in trabajo else pd.Series(pd.NA, index=trabajo.index)

    grupos = trabajo["ESTADO"].map(grupo_de_estado)
    orden = grupos.map(_ORDEN_GRUPO).fillna(len(GRUPOS_AVISO))

    # Clave ascendente unica: para "por vencer" el tiempo restante tal cual;
    # para "vencidos" se invierte el signo para que el mas atrasado quede arriba.
    clave = pd.Series(0.0, index=trabajo.index, dtype="float64")
    por_vencer = grupos == "PROXIMO"
    clave.loc[por_vencer] = restantes.loc[por_vencer].fillna(1e9).astype(float)
    ya_vencidos = grupos.isin(("VENCIDO", "CIERRE"))
    clave.loc[ya_vencidos] = (-vencido.loc[ya_vencidos].fillna(0.0)).astype(float)

    trabajo["_ORDEN_GRUPO"] = orden
    trabajo["_CLAVE_URGENCIA"] = clave
    trabajo = trabajo.sort_values(
        ["_ORDEN_GRUPO", "_CLAVE_URGENCIA"], kind="stable"
    )
    return trabajo.drop(columns=["_ORDEN_GRUPO", "_CLAVE_URGENCIA"])


def _frase_tiempo(fila: pd.Series) -> tuple[str, str]:
    """
    Devuelve (frase_urgente, verbo_de_fecha) de un caso.

    Se calcula desde las columnas NUMERICAS (HORAS_RESTANTES / HORAS_VENCIDO) y
    NO desde TIEMPO_RESTANTE: ese texto ya trae la palabra "restantes", y al
    combinarlo quedaba "VENCE EN 12 h 24 min restantes".
    """
    from core import formatear_duracion  # import local: evita ciclos

    estado = str(fila.get("ESTADO", ""))
    if estado in ("ROJO", "CERRADO TARDE"):
        horas = fila.get("HORAS_VENCIDO")
        if pd.notna(horas):
            return f"VENCIDO HACE {formatear_duracion(float(horas))}", "Venció"
        return "VENCIDO", "Venció"
    if estado == "SIN VENCIMIENTO":
        return "sin fecha en la plantilla", "Sin vencimiento"
    horas = fila.get("HORAS_RESTANTES")
    if pd.notna(horas):
        return f"VENCE EN {formatear_duracion(float(horas))}", "Vence"
    return "sin tiempo calculado", "Vence"


def _bloque_caso_html(fila: pd.Series, con_numero: int | None = None) -> list[str]:
    """
    Tres líneas cortas por caso optimizadas para lectura en Telegram.
    El número de caso usa <code> para permitir COPIADO TÁCTIL con 1 solo toque en el celular.
    """
    estado = str(fila.get("ESTADO", ""))
    icono, _ = GUIA_POR_ESTADO.get(estado, ("⚠️", "ATENDER"))
    marca = MARCA_URGENCIA.get(estado, "⚠️")
    caso_raw = str(fila.get("N° DE CASO", "S/N")).strip()
    caso = _esc(caso_raw)
    vencimiento = _esc(_fecha_legible(fila.get("FECHA_VENCIMIENTO")))
    oficina = _esc(_oficina_de(fila))
    region = str(fila.get("REGION_TECNICO", "") or "").strip()

    frase, verbo = _frase_tiempo(fila)
    frase = _esc(frase)

    prefijo = f"{con_numero}. " if con_numero is not None else ""
    linea_lugar = f"🏢 {oficina}"
    if region and region != "SIN REGION":
        linea_lugar += f" · 🗺️ {_esc(region)}"

    return [
        f"{icono} {prefijo}Caso <code>{caso}</code> — {marca} <b>{frase}</b>",
        f"🗓️ {verbo}: {vencimiento}",
        linea_lugar,
    ]


def componer_aviso_consolidado_telegram(df: pd.DataFrame, menciones: dict | None = None) -> str:
    """
    Genera un informe CONSOLIDADO GLOBAL con TODOS los casos pendientes de todos los técnicos.
    Formateado en HTML para Telegram con tags <code> para copiar rápido.
    """
    if df is None or df.empty:
        return "✅ <b>No hay casos pendientes en la Torre de Control SLA.</b>"

    momento = datetime.now().strftime("%d/%m/%Y %H:%M")
    ordenados = orden_para_aviso(df)
    total = len(ordenados)
    n_vencidos = int((ordenados["ESTADO"] == "ROJO").sum())
    n_por_vencer = int(ordenados["ESTADO"].isin(("NARANJA", "AMARILLO", "VERDE")).sum())

    tecnicos_unicos = ordenados["TECNICO"].dropna().unique()

    lineas = [
        "🚨 <b>REPORTE CONSOLIDADO GLOBAL DE SLA</b> 🚨",
        "🔔 <b>TORRE DE CONTROL</b> · Colsof / Banco Agrario",
        f"📊 Total Casos Alerta: <b>{total}</b> · 🔴 Vencidos: <b>{n_vencidos}</b> · ⏳ Por vencer: <b>{n_por_vencer}</b>",
        f"👥 Técnicos involucrados: <b>{len(tecnicos_unicos)}</b>",
        f"🕐 <i>Generado: {momento}</i>",
        "━" * 22,
        "",
    ]

    for tec, grp in ordenados.groupby("TECNICO", sort=False):
        mencion = mencion_para(tec, canal="telegram", menciones=menciones)
        lineas.append(f"👷 <b>{_esc(tec)}</b> ({_esc(mencion)}) — <b>{len(grp)} caso(s)</b>:")
        for idx, (_, fila) in enumerate(grp.iterrows(), start=1):
            caso = _esc(str(fila.get("N° DE CASO", "S/N")).strip())
            estado = str(fila.get("ESTADO", ""))
            marca = MARCA_URGENCIA.get(estado, "⚠️")
            frase, _ = _frase_tiempo(fila)
            oficina = _esc(_oficina_de(fila))
            lineas.append(f"  {idx}. <code>{caso}</code> — {marca} <b>{_esc(frase)}</b> | 🏢 {oficina}")
        lineas.append("")

    lineas.append("━" * 22)
    lineas.append(QUE_HACER_HTML)
    return "\n".join(lineas)


def componer_aviso_consolidado_whatsapp(df: pd.DataFrame, menciones: dict | None = None) -> str:
    """
    Genera un reporte consolidado global en texto plano para copiar a WhatsApp.
    """
    if df is None or df.empty:
        return "✅ No hay casos pendientes en la Torre de Control SLA."

    momento = datetime.now().strftime("%d/%m/%Y %H:%M")
    ordenados = orden_para_aviso(df)
    total = len(ordenados)
    n_vencidos = int((ordenados["ESTADO"] == "ROJO").sum())
    n_por_vencer = int(ordenados["ESTADO"].isin(("NARANJA", "AMARILLO", "VERDE")).sum())

    lineas = [
        "🚨 REPORTE CONSOLIDADO GLOBAL DE SLA 🚨",
        "TORRE DE CONTROL · Colsof / Banco Agrario",
        f"📊 Total Casos: {total} | 🔴 Vencidos: {n_vencidos} | ⏳ Por vencer: {n_por_vencer}",
        f"🕐 Generado: {momento}",
        "========================================",
        "",
    ]

    for tec, grp in ordenados.groupby("TECNICO", sort=False):
        mencion = mencion_para(tec, canal="whatsapp", menciones=menciones)
        lineas.append(f"👷 *{tec}* ({mencion}) — {len(grp)} caso(s):")
        for idx, (_, fila) in enumerate(grp.iterrows(), start=1):
            caso = str(fila.get("N° DE CASO", "S/N")).strip()
            estado = str(fila.get("ESTADO", ""))
            icono, _ = GUIA_POR_ESTADO.get(estado, ("⚠️", ""))
            frase, _ = _frase_tiempo(fila)
            oficina = _oficina_de(fila)
            lineas.append(f"  {idx}. Caso {caso} — {icono} *{frase}* | {oficina}")
        lineas.append("")

    lineas.append("========================================")
    lineas.append(CIERRE)
    return "\n".join(lineas)


def _usuario_mencion(tecnico: str, menciones: dict | None = None) -> str:
    """
    Devuelve SOLO el @usuario de Telegram del tecnico, o "" si no esta configurado.

    Sin @usuario, Telegram no notifica a nadie: el mensaje llega al grupo pero
    el tecnico no recibe aviso en su celular. El @usuario se configura en
    menciones.json (campo "telegram").
    """
    if menciones is None:
        menciones = cargar_menciones()
    info = menciones.get(_clave_tecnico(tecnico), {})
    return str(info.get("telegram", "") or "").strip()


def _encabezado_html(
    total: int,
    n_vencidos: int,
    n_por_vencer: int,
    tecnico: str = "",
    usuario: str = "",
) -> list[str]:
    """
    Primera parte del mensaje: NOMBRE DEL TECNICO + alarmas + resumen.

    IMPORTANTE: el nombre va en la PRIMERA linea. En el celular la notificacion
    solo muestra esa linea; antes ponia "¡CASOS VENCIDOS SIN CERRAR!" y en un
    grupo con 22 tecnicos nadie sabia de quien eran los casos sin abrir el
    mensaje. Y sin @usuario, Telegram ni siquiera avisa al tecnico.
    """
    alarma = ALARMA_VENCIDOS if n_vencidos else ALARMA_PROXIMOS
    etiqueta = "CASOS VENCIDOS SIN CERRAR" if n_vencidos else "CASOS POR VENCER"

    nombre = _esc(tecnico) if str(tecnico or "").strip() else "TECNICO SIN ASIGNAR"
    # El @usuario se agrega tal cual: Telegram lo convierte en mencion (y dentro
    # de <b> tambien funciona, asi que el tecnico queda etiquetado).
    tageo = f" {_esc(usuario)}" if str(usuario or "").startswith("@") else ""

    return [
        f"{alarma} <b>{nombre}</b>{tageo} {alarma}",
        f"⚠️ <b>{etiqueta}</b> · 🔴 <b>{n_vencidos}</b> vencidos · "
        f"⏳ <b>{n_por_vencer}</b> por vencer · Total <b>{total}</b> caso(s)",
        "🔔 <b>TORRE DE CONTROL SLA</b> · Colsof / Banco Agrario 🔔",
    ]


def _saludo_html(tecnico: str, canal: str, menciones: dict | None) -> str:
    """
    Saludo con el nombre en negrita.

    No repite el @usuario: ya va en la primera linea para que la notificacion
    del celular muestre el nombre.
    """
    if menciones is None:
        menciones = cargar_menciones()
    info = menciones.get(_clave_tecnico(tecnico), {})
    nombre = info.get("nombre_corto") or tecnico or "tecnico"
    return f"👋 Hola <b>{_esc(nombre)}</b>, estos son tus casos:"


def componer_aviso_telegram(df: pd.DataFrame, canal: str = "telegram",
                            menciones: dict | None = None,
                            un_solo_mensaje: bool = True,
                            maximo: int = 3900) -> str:
    """
    Mensaje listo para Telegram (parse_mode="HTML"), pensado para leerse facil
    en el celular y ordenado por urgencia: primero lo que esta por vencer y
    despues lo ya vencido.

    Estructura:
        1. Alarmas (🚨 si hay vencidos, ⏰ si solo hay por vencer) + resumen.
        2. Saludo con el nombre del tecnico.
        3. Secciones por prioridad, cada caso en 3 lineas cortas.
        4. QUE HACER: los pasos para gestionar y CERRAR los casos.

    un_solo_mensaje=True  -> un mensaje con los casos agrupados por prioridad.
    un_solo_mensaje=False -> un aviso individual por caso.

    maximo: se respeta el limite de Telegram (4096); se recortan casos si el
    mensaje se pasa de largo, dejando constancia de cuantos quedaron fuera.
    """
    if df is None or df.empty:
        return ""

    tecnico = str(df.iloc[0].get("TECNICO", "") or "").strip()
    momento = datetime.now().strftime("%d/%m/%Y %H:%M")
    orden = orden_para_aviso(df)
    total = len(orden)

    n_vencidos = int((orden["ESTADO"] == "ROJO").sum())
    n_por_vencer = int(
        orden["ESTADO"].isin(("NARANJA", "AMARILLO", "VERDE")).sum()
    )

    # @usuario del tecnico (vacio si no esta en menciones.json). Va en la
    # primera linea para que Telegram lo notifique de verdad.
    usuario = _usuario_mencion(tecnico, menciones)
    nombre = _esc(tecnico) if tecnico else "TECNICO SIN ASIGNAR"

    # --- Avisos individuales (un mensaje por caso) -------------------------
    if not un_solo_mensaje and total > 1:
        bloques: list[str] = []
        mostrados = 0
        for i, (_, fila) in enumerate(orden.iterrows(), start=1):
            bloque = "\n".join(
                [
                    f"{ALARMA_VENCIDOS} <b>{nombre}</b>"
                    f"{f' {_esc(usuario)}' if usuario.startswith('@') else ''} "
                    f"{ALARMA_VENCIDOS}",
                    f"⚠️ <b>AVISO DE VENCIMIENTO</b> · caso <b>{i}</b> de "
                    f"<b>{total}</b>",
                    "",
                    *_bloque_caso_html(fila),
                    "",
                    LINEA,
                    QUE_HACER_HTML,
                ]
            )
            # Telegram rechaza mensajes de mas de 4096 caracteres con HTTP 400,
            # y enviar_mensaje NO los parte en trozos. Por eso se corta aqui:
            # si no, un tecnico con 30 casos genera 20.000 caracteres y el
            # mensaje no se envia NUNCA.
            candidato = ("\n\n" + LINEA + "\n\n").join(bloques + [bloque])
            if len(candidato) > maximo - 300:  # margen para el aviso de corte
                break
            bloques.append(bloque)
            mostrados += 1

        if not bloques:
            return ""
        texto = ("\n\n" + LINEA + "\n\n").join(bloques)
        if mostrados < total:
            texto += (
                f"\n\n⚠️ <i>Se incluyeron {mostrados} de {total} casos por el "
                "límite de longitud de Telegram. El resto va en el siguiente "
                "aviso.</i>"
            )
        return texto

    # --- Un solo mensaje, agrupado por prioridad ---------------------------
    lineas: list[str] = _encabezado_html(
        total, n_vencidos, n_por_vencer, tecnico, usuario
    )
    lineas.append("")
    lineas.append(_saludo_html(tecnico, canal, menciones))

    # El cierre (QUE HACER + pie) se reserva ANTES de meter casos. Si se deja
    # para el final, 30 casos + el cierre pasaban de los 4096 caracteres que
    # acepta Telegram y el mensaje era rechazado con HTTP 400.
    cola = ["", LINEA, QUE_HACER_HTML, LINEA]
    presupuesto = maximo - len("\n".join(cola)) - 220

    mostrados = 0
    cubiertos: set[str] = set()
    for clave, titulo, estados in GRUPOS_AVISO:
        sub = orden[orden["ESTADO"].isin(estados)]
        if sub.empty:
            continue
        cubiertos.update(estados)
        izq, der = ENCABEZADO_ALARMA.get(clave, ("", ""))
        lineas.extend(
            ["", LINEA, f"{izq} <b>{titulo} ({len(sub)})</b> {der}".strip(), LINEA]
        )
        primero = True
        for _, fila in sub.iterrows():
            nuevo = _bloque_caso_html(fila)
            if len("\n".join(lineas + nuevo)) > presupuesto:
                continue
            if not primero:
                # Linea en blanco entre casos: sin esto la lista se ve como un
                # bloque y cuesta encontrar cada caso.
                lineas.append("")
            lineas.extend(nuevo)
            mostrados += 1
            primero = False

    # Estados no contemplados: nunca se pierden en silencio.
    resto = orden[~orden["ESTADO"].isin(cubiertos)]
    if not resto.empty:
        lineas.extend(["", LINEA, f"<b>📋 OTROS ({len(resto)})</b>", LINEA])
        primero = True
        for _, fila in resto.iterrows():
            nuevo = _bloque_caso_html(fila)
            if len("\n".join(lineas + nuevo)) > presupuesto:
                continue
            if not primero:
                lineas.append("")
            lineas.extend(nuevo)
            mostrados += 1
            primero = False

    lineas.extend(cola)
    if mostrados < total:
        lineas.append(
            f"⚠️ <i>Se listaron {mostrados} de {total} casos por el límite de "
            "longitud de Telegram.</i>"
        )
    lineas.extend(["", f"🕐 <i>Aviso generado: {momento}</i>"])
    return "\n".join(lineas)


def componer_aviso_tecnico_telegram(tecnico: str, df_tecnico: pd.DataFrame,
                                    un_solo_mensaje: bool = True,
                                    menciones: dict | None = None) -> str:
    """Version HTML (Telegram) de componer_aviso_tecnico."""
    if df_tecnico is None or df_tecnico.empty:
        return ""
    return componer_aviso_telegram(
        df_tecnico, canal="telegram", menciones=menciones,
        un_solo_mensaje=un_solo_mensaje,
    )


# ---------------------------------------------------------------------------
# Agrupacion y resumen
# ---------------------------------------------------------------------------

def agrupar_por_tecnico(df: pd.DataFrame, solo_con_tecnico: bool = True) -> dict[str, pd.DataFrame]:
    """
    Agrupa los casos por tecnico, ordenando cada grupo por gravedad.

    Los casos sin tecnico asignado se agrupan bajo "(sin tecnico)" salvo que
    solo_con_tecnico=True (por defecto se incluyen, para no perderlos de vista).
    """
    if df is None or df.empty:
        return {}

    trabajo = df.copy()
    if "TECNICO" not in trabajo.columns:
        return {}

    trabajo["TECNICO"] = trabajo["TECNICO"].fillna("").replace("", "(sin tecnico)")
    grupos: dict[str, pd.DataFrame] = {}
    for tecnico, grupo in trabajo.groupby("TECNICO", sort=True):
        if solo_con_tecnico and tecnico == "(sin tecnico)":
            continue
        grupos[str(tecnico)] = grupo.copy()
    return grupos


def resumen_pendientes(df: pd.DataFrame) -> pd.DataFrame:
    """
    Tabla de pendientes por tecnico, para la vista de la torre de control.

    Columnas: TECNICO, REGION, PENDIENTES, ROJOS, NARANJAS, AMARILLOS,
              PROXIMO_VENCIMIENTO, ULTIMA_NOTIFICACION, VECES_HOY
    """
    columnas = [
        "TECNICO", "REGION", "PENDIENTES", "ROJOS", "NARANJAS", "AMARILLOS",
        "PROXIMO_VENCIMIENTO", "ULTIMA_NOTIFICACION", "VECES_HOY",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columnas)

    filas = []
    for tecnico, grupo in agrupar_por_tecnico(df, solo_con_tecnico=False).items():
        vencimientos = grupo["FECHA_VENCIMIENTO"].dropna()
        filas.append(
            {
                "TECNICO": tecnico,
                "REGION": grupo["REGION_TECNICO"].mode().iloc[0]
                if not grupo["REGION_TECNICO"].mode().empty else "",
                "PENDIENTES": len(grupo),
                "ROJOS": int((grupo["ESTADO"] == "ROJO").sum()),
                "NARANJAS": int((grupo["ESTADO"] == "NARANJA").sum()),
                "AMARILLOS": int((grupo["ESTADO"] == "AMARILLO").sum()),
                "PROXIMO_VENCIMIENTO": vencimientos.min() if not vencimientos.empty else pd.NaT,
            }
        )

    tabla = pd.DataFrame(filas)
    # Se completan con el historial de notificaciones
    try:
        from historial import Historial

        hist = Historial()
        ultimas = hist.ultima_notificacion_por_tecnico()
        hoy = hist.notificaciones_de_hoy_por_tecnico()
        tabla["ULTIMA_NOTIFICACION"] = tabla["TECNICO"].map(ultimas).fillna("nunca")
        tabla["VECES_HOY"] = tabla["TECNICO"].map(hoy).fillna(0).astype(int)
    except Exception:
        tabla["ULTIMA_NOTIFICACION"] = "nunca"
        tabla["VECES_HOY"] = 0

    return tabla.sort_values(
        ["ROJOS", "PENDIENTES"], ascending=[False, False]
    ).reset_index(drop=True)[columnas]


if __name__ == "__main__":
    # Demostracion:  python avisos.py
    import warnings

    warnings.filterwarnings("ignore")
    from core import calcular_tablero

    resultado = calcular_tablero()
    pendientes = resultado.requieren_atencion
    print("=" * 74)
    print(" DEMOSTRACION DE AVISOS POR TECNICO")
    print("=" * 74)
    print(f"Casos pendientes en la ventana: {len(pendientes)}")
    print()
    print("--- RESUMEN POR TECNICO (pendientes de notificar) ---")
    print(resumen_pendientes(pendientes).to_string(index=False))
    print()
    print("--- EJEMPLO DE AVISO INDIVIDUAL ---")
    if not pendientes.empty:
        print(componer_aviso(pendientes.iloc[0], canal="whatsapp"))
