"""
telegram_notifier.py - Envio de notificaciones y tablas a Telegram.

Usa la API HTTP de Telegram (metodo sendMessage). No requiere librerias
externas mas alla de `requests`.

CONFIGURACION (dos opciones):

 1) Variables de entorno (recomendado):
        set TELEGRAM_BOT_TOKEN=123456:ABC...
        set TELEGRAM_CHAT_ID=-1001234567890

 2) Archivo config_telegram.json junto a este script:
        { "bot_token": "123456:ABC...", "chat_id": "-1001234567890" }
    (este archivo esta excluido de git mediante .gitignore)

COMO CREAR EL BOT:
  1. En Telegram, busque @BotFather y envie /newbot
  2. Copie el token que le entrega
  3. Agregue el bot al grupo de la torre de control
  4. Obtenga el chat_id del grupo (vea docs/TUTORIAL_TELEGRAM.md)

Uso por consola:
    python telegram_notifier.py --test
    python telegram_notifier.py --tabla
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import time
from datetime import datetime

import pandas as pd

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None  # type: ignore

# Estilos de presentacion (formatos.py). Solo se importan los nombres, para no
# crear dependencia circular con el modulo de notificaciones.
try:
    from formatos import ESTILOS
except ImportError:  # pragma: no cover
    ESTILOS = ("tabla", "tarjetas", "por-estado", "markdown", "resumen")

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_CONFIG = os.path.join(DIR_SCRIPT, "config_telegram.json")
ARCHIVO_CONFIG_EJEMPLO = os.path.join(DIR_SCRIPT, "config_telegram.example.json")

API_URL = "https://api.telegram.org/bot{token}/sendMessage"
LONGITUD_MAXIMA = 4000      # limite de Telegram es 4096; se deja margen
FILAS_POR_MENSAJE = 20      # filas por tabla antes de dividir el mensaje
REINTENTOS = 3
ESPERA_BASE = 2.0

# --- Filtros de destino por region -----------------------------------------
FILTRO_TODO = "todo"              # el destino recibe todos los casos
FILTRO_BOGOTA = "bogota"          # solo casos de Bogota
FILTRO_REGIONALES = "regionales"  # todo lo que NO sea Bogota

# Etiquetas legibles de cada grupo, para titulos y logs.
ETIQUETA_GRUPO = {
    FILTRO_BOGOTA: "BOGOTA",
    FILTRO_REGIONALES: "REGIONALES",
    FILTRO_TODO: "",
}


def _normalizar_filtro(valor) -> str:
    """
    Normaliza el campo 'filtro' de un destino.

    Acepta: None/vacio -> "todo"; "bogota"; "regionales"; "regional" (alias);
    o una lista de regiones, que se devuelve como texto separado por comas
    (se compara contra REGION_TECNICO en core.parte_por_region / filtrar_por_filtro).
    """
    if valor is None or valor == "":
        return FILTRO_TODO
    if isinstance(valor, (list, tuple)):
        return ",".join(str(v).strip().upper() for v in valor if str(v).strip()) or FILTRO_TODO

    texto = str(valor).strip().lower()
    if texto in ("todo", "todos", "all", "*"):
        return FILTRO_TODO
    if texto in ("bogota", "bogotá", "bog"):
        return FILTRO_BOGOTA
    if texto in ("regionales", "regional", "regions", "resto"):
        return FILTRO_REGIONALES
    # Lista explicita de regiones escrita como texto: "CALI, NEIVA"
    return ",".join(p.strip().upper() for p in texto.split(",") if p.strip())

log = logging.getLogger("telegram")


# ---------------------------------------------------------------------------
# Configuracion
# ---------------------------------------------------------------------------

def cargar_config(ruta: str | None = None) -> tuple[str, str]:
    """
    Devuelve (bot_token, chat_id). Prioriza variables de entorno.

    Para MULTIPLES destinos use cargar_destinos(); esta funcion se mantiene por
    compatibilidad y devuelve el primer destino.

    Lanza ValueError con instrucciones claras si falta la configuracion.
    """
    token, destinos = cargar_destinos(ruta)
    if not destinos:
        raise ValueError(_mensaje_falta_config())
    return token, destinos[0]["chat_id"]


def _mensaje_falta_config() -> str:
    return (
        "Falta la configuracion de Telegram.\n"
        "Defina las variables de entorno TELEGRAM_BOT_TOKEN y TELEGRAM_CHAT_ID, "
        f"o cree el archivo {os.path.basename(ARCHIVO_CONFIG)} basandose en "
        f"{os.path.basename(ARCHIVO_CONFIG_EJEMPLO)}.\n"
        "Vealo detallado en docs/TUTORIAL_TELEGRAM.md"
    )


def cargar_destinos(ruta: str | None = None) -> tuple[str, list[dict[str, str]]]:
    """
    Devuelve (bot_token, [destinos]).

    Cada destino es {"nombre": str, "chat_id": str, "tema_id": str, "filtro": str}.

    'filtro' permite dirigir cada destino a un subconjunto de los casos:
        "todo"       (por defecto) -> recibe todos los casos
        "bogota"     -> solo los casos cuya region es BOGOTA
        "regionales" -> todo lo que NO sea Bogota
    Tambien admite una lista de regiones:  "filtro": ["CALI", "MEDELLÍN"]

    Configuracion admitida en config_telegram.json (todas validas):

      1) Un solo destino, formato clasico:
             {"bot_token": "...", "chat_id": "-100123"}

      2) Varios destinos como lista de chat_id:
             {"bot_token": "...", "chat_id": ["-100123", "-100456"]}

      3) Varios destinos con nombre y filtro por region (RECOMENDADO para
         distribuir Bogota y Regionales en grupos distintos):
             {"bot_token": "...",
              "destinos": [
                 {"nombre": "Bogota",     "chat_id": "-100111", "filtro": "bogota"},
                 {"nombre": "Regionales", "chat_id": "-100222", "filtro": "regionales"}
              ]}

    Ademas, la variable de entorno TELEGRAM_CHAT_ID admite varios ids separados
    por coma:  TELEGRAM_CHAT_ID="-100123,-100456"

    'tema_id' es opcional y sirve para foros (temas) de Telegram.
    """
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    env_chat = os.environ.get("TELEGRAM_CHAT_ID", "").strip()
    destinos: list[dict[str, str]] = []

    # --- 1. Destinos desde variables de entorno (tienen prioridad) ---
    if env_chat:
        for cid in env_chat.replace(";", ",").split(","):
            cid = cid.strip()
            if cid:
                destinos.append({
                    "nombre": "variable de entorno", "chat_id": cid,
                    "tema_id": "", "filtro": FILTRO_TODO,
                })

    # --- 2. Destinos desde el archivo de configuracion ---
    ruta = ruta or ARCHIVO_CONFIG
    if os.path.isfile(ruta):
        try:
            # utf-8-sig: tolera el BOM que agregan el Bloc de notas de Windows y
            # PowerShell al guardar como UTF-8. Sin esto, json.load falla con
            # "Unexpected UTF-8 BOM" y el usuario no entiende por que.
            with open(ruta, encoding="utf-8-sig") as fh:
                datos = json.load(fh)
        except (OSError, json.JSONDecodeError) as exc:
            raise ValueError(
                f"No se pudo leer {os.path.basename(ruta)}: {exc}\n"
                "Verifique que el archivo tenga formato JSON valido "
                "(llaves, comillas dobles y comas)."
            ) from exc

        token = token or str(datos.get("bot_token", "")).strip()

        # Forma 3: lista de objetos con nombre y filtro opcional
        for d in datos.get("destinos", []) or []:
            if isinstance(d, dict) and str(d.get("chat_id", "")).strip():
                destinos.append(
                    {
                        "nombre": str(d.get("nombre", "")).strip() or "grupo",
                        "chat_id": str(d["chat_id"]).strip(),
                        "tema_id": str(d.get("tema_id", "") or "").strip(),
                        "filtro": _normalizar_filtro(d.get("filtro")),
                    }
                )

        # Formas 1 y 2: "chat_id" simple o lista
        if not destinos:
            crudo = datos.get("chat_id", "")
            if isinstance(crudo, (list, tuple)):
                for cid in crudo:
                    if str(cid).strip():
                        destinos.append({
                            "nombre": "grupo", "chat_id": str(cid).strip(),
                            "tema_id": "", "filtro": FILTRO_TODO,
                        })
            elif str(crudo).strip():
                destinos.append({
                    "nombre": "grupo", "chat_id": str(crudo).strip(),
                    "tema_id": "", "filtro": FILTRO_TODO,
                })

    if not token or not destinos:
        raise ValueError(_mensaje_falta_config())

    return token, destinos


def telegram_configurado() -> bool:
    """True si hay credenciales disponibles (sin lanzar excepcion)."""
    try:
        cargar_config()
        return True
    except ValueError:
        return False


# ---------------------------------------------------------------------------
# Envio
# ---------------------------------------------------------------------------

def _explicar_error(respuesta, chat_id: str = "") -> str:
    """
    Traduce los errores de Telegram a mensajes accionables en espanol.

    Telegram responde en ingles ("chat not found"), lo que no ayuda al usuario
    que esta configurando el bot por primera vez. Se distingue si el destino es
    un grupo (chat_id negativo) o una persona, porque la solucion es distinta.
    """
    try:
        detalle = respuesta.json().get("description", "") or respuesta.text
    except Exception:
        detalle = respuesta.text or ""

    bajo = detalle.lower()
    if "chat not found" in bajo:
        if str(chat_id).startswith("-"):
            return (
                f"{detalle}\n"
                "        -> El chat_id del GRUPO es incorrecto, o el bot no esta en el grupo.\n"
                "           Solucion: agregue el bot al grupo, envie un mensaje al grupo\n"
                "           y ejecute:  python telegram_notifier.py --chats"
            )
        return (
            f"{detalle}\n"
            "        -> Esa PERSONA nunca le ha escrito al bot, o el id no es el real.\n"
            "           Telegram NO usa el numero de telefono como chat_id, y un bot\n"
            "           NO puede iniciar una conversacion.\n"
            "           Solucion (30 seg): la persona abre el chat del bot en Telegram,\n"
            "           pulsa START (o escribe 'hola') y luego ejecute:\n"
            "               python telegram_notifier.py --chats\n"
            "           El chat_id real son 9-10 digitos, SIN el 57 del indicativo."
        )
    if "not enough rights" in bajo:
        return (
            f"{detalle}\n"
            "        -> El bot no tiene permiso para escribir en el grupo.\n"
            "           Solucion: haga administrador al bot, o permita enviar mensajes."
        )
    if "bot was blocked" in bajo or "user is deactivated" in bajo:
        return (
            f"{detalle}\n"
            "        -> El destinatario bloqueo al bot, o nunca pulso /start.\n"
            "           Para mensajes privados, la persona debe abrir el chat del bot\n"
            "           y pulsar Start una vez."
        )
    if "unauthorized" in bajo:
        return (
            f"{detalle}\n"
            "        -> El TOKEN del bot es invalido o fue revocado.\n"
            "           Solucion: verifiquelo con: python telegram_notifier.py --verificar"
        )
    if "parse" in bajo or "entities" in bajo:
        return (
            f"{detalle}\n"
            "        -> Problema de formato. Los avisos con emojis deben enviarse\n"
            "           con parse_mode vacio (texto plano)."
        )
    return f"{detalle[:300]}"


def enviar_mensaje(
    texto: str,
    parse_mode: str = "HTML",
    token: str | None = None,
    chat_id: str | None = None,
    silencioso: bool = True,
    tema_id: str | None = None,
) -> bool:
    """
    Envia un mensaje a UN destino de Telegram. Devuelve True si se envio.

    Reintenta con espera creciente ante fallos de red o errores 5xx.
    Si se omiten token/chat_id, se resuelve el primer destino configurado
    (para envios a un unico grupo use enviar_a_todos()).
    """
    if requests is None:
        log.error("Falta el paquete 'requests'. Instale: python -m pip install requests")
        return False

    try:
        if token is None or chat_id is None:
            token, chat_id = cargar_config()
    except ValueError as exc:
        log.error("%s", exc)
        return False

    if not texto or not texto.strip():
        log.warning("Mensaje vacio: no se envia nada.")
        return False

    carga = {
        "chat_id": chat_id,
        "text": texto,
        "parse_mode": parse_mode,
        "disable_web_page_preview": True,
        "disable_notification": silencioso,
    }
    if tema_id:
        carga["message_thread_id"] = tema_id

    for intento in range(1, REINTENTOS + 1):
        try:
            respuesta = requests.post(
                API_URL.format(token=token), json=carga, timeout=20
            )
            if respuesta.status_code == 200:
                return True

            # 400 suele ser formato/parse_mode incorrecto: no tiene sentido reintentar.
            if 400 <= respuesta.status_code < 500:
                log.error(
                    "Telegram rechazo el mensaje (HTTP %s): %s",
                    respuesta.status_code, _explicar_error(respuesta, str(chat_id)),
                )
                return False

            log.warning(
                "Intento %d/%d fallo (HTTP %s). Reintentando...",
                intento, REINTENTOS, respuesta.status_code,
            )
        except Exception as exc:
            log.warning(
                "Intento %d/%d fallo (%s: %s). Reintentando...",
                intento, REINTENTOS, type(exc).__name__, exc,
            )

        if intento < REINTENTOS:
            time.sleep(ESPERA_BASE * intento)

    log.error("No se pudo enviar el mensaje a Telegram tras %d intentos.", REINTENTOS)
    return False


def enviar_a_todos(texto: str, parse_mode: str = "HTML", silencioso: bool = True) -> int:
    """
    Envia el mismo mensaje a TODOS los destinos configurados.

    Returns:
        Cantidad de destinos a los que se envio correctamente.
    """
    try:
        token, destinos = cargar_destinos()
    except ValueError as exc:
        log.error("%s", exc)
        return 0

    enviados = 0
    for destino in destinos:
        ok = enviar_mensaje(
            texto,
            parse_mode=parse_mode,
            token=token,
            chat_id=destino["chat_id"],
            silencioso=silencioso,
            tema_id=destino.get("tema_id") or None,
        )
        if ok:
            enviados += 1
            log.info("Enviado a '%s'", destino["nombre"])
        else:
            log.error("FALLO el envio a '%s' (%s)", destino["nombre"], destino["chat_id"])
        time.sleep(0.5)  # cortesia con el limite de la API
    return enviados


def destinos_configurados() -> list[dict[str, str]]:
    """Lista de destinos configurados (vacia si no hay configuracion)."""
    try:
        _, destinos = cargar_destinos()
        return destinos
    except ValueError:
        return []


def verificar_bot(token: str | None = None) -> dict:
    """
    Consulta a la API de Telegram si el bot responde y devuelve su informacion.

    Solo necesita el TOKEN (no requiere destinos), para poder validar las
    credenciales antes de tener el chat_id del grupo.

    Returns: {"ok": bool, "bot": str, "error": str}
    """
    if requests is None:
        return {"ok": False, "bot": "", "error": "Falta el paquete 'requests'."}

    if not token:
        # Se intenta obtener el token aunque aun no haya destinos configurados.
        token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
        if not token and os.path.isfile(ARCHIVO_CONFIG):
            try:
                with open(ARCHIVO_CONFIG, encoding="utf-8-sig") as fh:
                    token = str(json.load(fh).get("bot_token", "")).strip()
            except (OSError, json.JSONDecodeError):
                token = ""
        if not token:
            return {"ok": False, "bot": "", "error": _mensaje_falta_config()}

    try:
        r = requests.get(f"https://api.telegram.org/bot{token}/getMe", timeout=15)
        datos = r.json()
        if r.status_code == 200 and datos.get("ok"):
            b = datos["result"]
            return {"ok": True, "bot": f"@{b.get('username')} ({b.get('first_name')})", "error": ""}
        return {"ok": False, "bot": "", "error": f"HTTP {r.status_code}: {datos.get('description', '')}"}
    except Exception as exc:
        return {"ok": False, "bot": "", "error": f"{type(exc).__name__}: {exc}"}


def obtener_chat_ids_recientes(token: str | None = None, limite: int = 20) -> list[dict[str, str]]:
    """
    Lista los chats que el bot ha visto recientemente (via getUpdates).

    Es la forma mas comoda de obtener el chat_id: se agrega el bot al grupo, se
    envia un mensaje al grupo y esta funcion devuelve el id y el nombre.

    Returns: [{"tipo": str, "nombre": str, "chat_id": str}]
    """
    if requests is None:
        return []
    if not token:
        token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
        if not token and os.path.isfile(ARCHIVO_CONFIG):
            try:
                with open(ARCHIVO_CONFIG, encoding="utf-8-sig") as fh:
                    token = str(json.load(fh).get("bot_token", "")).strip()
            except (OSError, json.JSONDecodeError):
                token = ""
    if not token:
        return []

    try:
        r = requests.get(
            f"https://api.telegram.org/bot{token}/getUpdates", timeout=15
        )
        datos = r.json()
        if not datos.get("ok"):
            return []
    except Exception:
        return []

    vistos: dict[str, dict[str, str]] = {}
    for actualizacion in datos.get("result", [])[-limite:]:
        for clave in ("message", "channel_post", "edited_message", "my_chat_member"):
            obj = actualizacion.get(clave) or {}
            chat = obj.get("chat") or {}
            cid = chat.get("id")
            if cid is None:
                continue
            nombre = (
                chat.get("title")
                or " ".join(
                    x for x in [chat.get("first_name"), chat.get("last_name")] if x
                )
                or chat.get("username")
                or "(sin nombre)"
            )
            vistos[str(cid)] = {
                "tipo": str(chat.get("type", "")),
                "nombre": str(nombre),
                "chat_id": str(cid),
            }
    return list(vistos.values())


# ---------------------------------------------------------------------------
# Formato de tablas
# ---------------------------------------------------------------------------

def _recortar(texto: str, ancho: int) -> str:
    """Recorta a `ancho` caracteres agregando elipsis si hace falta."""
    texto = str(texto)
    if len(texto) <= ancho:
        return texto
    return texto[: max(ancho - 1, 1)] + "…"


def _filas_tabla(df: pd.DataFrame) -> list[str]:
    """
    Construye las celdas de cada fila.

    El orden de columnas pone TECNICO al final a proposito: es el campo mas
    largo y variable, y al ser la ultima columna no necesita relleno. Asi todos
    los separadores internos quedan alineados aunque el cliente de mensajeria
    recorte los espacios finales de linea al copiar.
    """
    filas = []
    for _, f in df.iterrows():
        filas.append(
            [
                _recortar(str(f.get("ESTADO", "")), 14),
                _recortar(str(f.get("N° DE CASO", "S/N")), 13),
                _recortar(_formatear_fecha(f.get("FECHA_VENCIMIENTO")), 16),
                _recortar(str(f.get("TIEMPO_CLAVE", "")), 18),
                _recortar(str(f.get("REGION_TECNICO", "")), 14),
                _recortar(str(f.get("TECNICO", "")) or "(sin tecnico)", 30),
            ]
        )
    return filas


def _formatear_fecha(valor) -> str:
    """Fecha legible 'dd/mm HH:MM' o '—'."""
    if valor is None or pd.isna(valor):
        return "—"
    try:
        return pd.Timestamp(valor).strftime("%d/%m %H:%M")
    except (ValueError, TypeError):
        return "—"


def _columna_tiempo(df: pd.DataFrame) -> pd.Series:
    """
    Texto del tiempo relevante segun el estado: cuanto lleva vencido o cuanto
    le queda. Es lo que la torre necesita ver de un vistazo.
    """
    def _texto(fila) -> str:
        estado = str(fila.get("ESTADO", ""))
        if estado in ("ROJO", "CERRADO TARDE"):
            hv = fila.get("HORAS_VENCIDO")
            return f"VENCIDO {hv:.0f} h" if pd.notna(hv) else "VENCIDO"
        return str(fila.get("TIEMPO_RESTANTE", "—"))

    return df.apply(_texto, axis=1)


def construir_tabla(df: pd.DataFrame, titulo: str = "ESTADO DE CASOS") -> str:
    """
    Tabla en bloque <pre>: se copia desde Telegram conservando la alineacion.

    Si hay mas filas que FILAS_POR_MENSAJE, el texto resultante se trocea luego
    con dividir_mensaje().
    """
    if df is None or df.empty:
        return f"<b>{titulo}</b>\n\nSin casos en la ventana de notificacion."

    trabajo = df.copy()
    trabajo["TIEMPO_CLAVE"] = _columna_tiempo(trabajo)
    filas = _filas_tabla(trabajo)

    encabezados = ["ESTADO", "CASO", "VENCE", "TIEMPO", "REGION", "TECNICO"]
    # Ancho de cada columna = maximo entre encabezado y contenido
    anchos = [
        max(len(encabezados[i]), max((len(f[i]) for f in filas), default=0))
        for i in range(len(encabezados))
    ]

    def linea(valores: list[str]) -> str:
        return "  ".join(v.ljust(anchos[i]) for i, v in enumerate(valores)).rstrip()

    separador = "  ".join("-" * a for a in anchos)
    cuerpo = [linea(encabezados), separador] + [linea(f) for f in filas]

    momento = datetime.now().strftime("%d/%m/%Y %H:%M")
    resumen = _resumen_texto(df)

    return (
        f"<b>🛰️ {titulo}</b>\n"
        f"<i>Colsof / Banco Agrario — {momento}</i>\n"
        f"{resumen}\n\n"
        f"<pre>{chr(10).join(cuerpo)}</pre>"
    )


def _resumen_texto(df: pd.DataFrame) -> str:
    """Linea de totales por estado."""
    if df is None or df.empty:
        return "Sin casos."
    conteo = df["ESTADO"].value_counts().to_dict()
    orden = ["ROJO", "CERRADO TARDE", "NARANJA", "AMARILLO", "VERDE", "SIN VENCIMIENTO"]
    partes = [f"{e}: {conteo[e]}" for e in orden if conteo.get(e)]
    return " | ".join(partes) if partes else "Sin casos."


def dividir_mensaje(texto: str, maximo: int = LONGITUD_MAXIMA) -> list[str]:
    """
    Divide un mensaje largo respetando la estructura <pre> de cada bloque.

    Se corta por lineas completas para no romper la alineacion ni las etiquetas.
    """
    if len(texto) <= maximo:
        return [texto]

    partes: list[str] = []
    acumulado = ""
    for linea in texto.split("\n"):
        if len(acumulado) + len(linea) + 1 > maximo:
            partes.append(acumulado.rstrip())
            acumulado = ""
        acumulado += linea + "\n"
    if acumulado.strip():
        partes.append(acumulado.rstrip())
    return partes


def construir_tabla_plana(df: pd.DataFrame, titulo: str = "ESTADO DE CASOS") -> str:
    """
    Formato B: texto plano con un caso por linea separado por '|'.

    Sirve para pegar en WhatsApp, correo o cualquier medio sin formato. No usa
    <pre> ni HTML.
    """
    if df is None or df.empty:
        return f"{titulo}\nSin casos en la ventana de notificacion."

    trabajo = df.copy()
    trabajo["TIEMPO_CLAVE"] = _columna_tiempo(trabajo)
    momento = datetime.now().strftime("%d/%m/%Y %H:%M")

    lineas = [
        f"{titulo} — Colsof / Banco Agrario",
        f"Calculado: {momento}",
        _resumen_texto(df),
        "",
    ]
    for _, f in trabajo.iterrows():
        lineas.append(
            " | ".join(
                [
                    str(f.get("ESTADO", "")),
                    str(f.get("N° DE CASO", "S/N")),
                    f"VENCE {_formatear_fecha(f.get('FECHA_VENCIMIENTO'))}",
                    str(f.get("TIEMPO_CLAVE", "")),
                    str(f.get("TECNICO", "")) or "(sin tecnico)",
                    str(f.get("REGION_TECNICO", "")),
                ]
            )
        )
    return "\n".join(lineas)


def filtrar_casos_por_filtro(df: pd.DataFrame, filtro: str) -> pd.DataFrame:
    """
    Aplica el filtro de region de un destino a un DataFrame de casos.

    filtro:
        "todo"       -> devuelve todo
        "bogota"     -> solo REGION_TECNICO == BOGOTA
        "regionales" -> todo lo que NO sea Bogota
        "CALI,NEIVA" -> solo esas regiones
    """
    if df is None or df.empty:
        return df

    filtro = _normalizar_filtro(filtro)
    if filtro == FILTRO_TODO or "REGION_TECNICO" not in df.columns:
        return df

    regiones = df["REGION_TECNICO"].astype(str).str.strip().str.upper()

    if filtro == FILTRO_BOGOTA:
        return df[regiones == "BOGOTA"].copy()
    if filtro == FILTRO_REGIONALES:
        return df[regiones != "BOGOTA"].copy()

    permitidas = {p.strip().upper() for p in filtro.split(",") if p.strip()}
    return df[regiones.isin(permitidas)].copy()


def _enviar_tabla_destino(
    df: pd.DataFrame, destino: dict, token: str, titulo: str,
    max_filas: int, plano: bool, estilo: str | None = None,
) -> int:
    """Envia la tabla a UN destino. Devuelve cuantos mensajes se enviaron."""
    if max_filas and len(df) > max_filas:
        trozos = [df.iloc[i:i + max_filas] for i in range(0, len(df), max_filas)]
    else:
        trozos = [df]

    total_trozos = len(trozos)
    enviados = 0

    for indice, trozo in enumerate(trozos, start=1):
        etiqueta = titulo if total_trozos == 1 else f"{titulo} ({indice}/{total_trozos})"
        if plano:
            texto = construir_tabla_plana(trozo, etiqueta)
            partes = [texto[i:i + LONGITUD_MAXIMA]
                      for i in range(0, len(texto), LONGITUD_MAXIMA)]
        elif estilo:
            # Estilo elegido por el usuario (formatos.py): tarjetas, por-estado, etc.
            import formatos

            partes = dividir_mensaje(
                formatos.generar(trozo, estilo=estilo, titulo=etiqueta)
            )
        else:
            partes = dividir_mensaje(construir_tabla(trozo, etiqueta))

        for parte in partes:
            # Red de seguridad: Telegram rechaza mensajes de mas de 4096
            # caracteres. Si una parte se pasa, se trocea por lineas completas
            # para no perder el aviso.
            sub_partes = dividir_mensaje(parte) if len(parte) > LONGITUD_MAXIMA else [parte]
            for sub in sub_partes:
                if enviar_mensaje(
                    sub, token=token, chat_id=destino["chat_id"],
                    tema_id=destino.get("tema_id") or None,
                ):
                    enviados += 1
                else:
                    log.error(
                        "Fallo el envio del bloque %d/%d a '%s' (%d caracteres)",
                        indice, total_trozos, destino["nombre"], len(sub),
                    )
                time.sleep(1.0)

    return enviados


def enviar_tabla_por_region(
    df: pd.DataFrame,
    max_filas: int = FILAS_POR_MENSAJE,
    plano: bool = False,
    solo_destino: str | None = None,
    estilo: str | None = None,
) -> dict[str, int]:
    """
    Envia la tabla a cada destino RESPETANDO su filtro de region.

    Es la forma de distribuir Bogota y Regionales en grupos distintos: cada
    destino recibe solo los casos que le corresponden.

    Args:
        solo_destino: si se indica, envia UNICAMENTE a los destinos cuyo nombre
                      contenga ese texto (util para probar un grupo puntual).
        estilo:       estilo de presentacion de formatos.py (tabla, tarjetas,
                      por-estado, markdown, resumen). None = tabla clasica.

    Returns: {nombre_destino: mensajes_enviados}
    """
    if df is None or df.empty:
        log.info("No hay casos que enviar a Telegram.")
        return {}

    try:
        token, destinos = cargar_destinos()
    except ValueError as exc:
        log.error("%s", exc)
        return {}

    if solo_destino:
        objetivo = solo_destino.strip().lower()
        destinos = [d for d in destinos if objetivo in d["nombre"].lower()]
        if not destinos:
            log.error("Ningun destino coincide con '%s'.", solo_destino)
            return {}

    resultados: dict[str, int] = {}

    for destino in destinos:
        filtro = destino.get("filtro", FILTRO_TODO)
        subconjunto = filtrar_casos_por_filtro(df, filtro)
        etiqueta = ETIQUETA_GRUPO.get(_normalizar_filtro(filtro), "")
        titulo = (f"CASOS QUE REQUIEREN ATENCION - {etiqueta}" if etiqueta
                  else "CASOS QUE REQUIEREN ATENCION")

        if subconjunto.empty:
            log.info(
                "'%s': sin casos que le correspondan (%s). No se envia nada.",
                destino["nombre"], filtro,
            )
            resultados[destino["nombre"]] = 0
            continue

        log.info("Enviando a '%s' (filtro=%s): %d caso(s)",
                 destino["nombre"], filtro, len(subconjunto))
        enviados = _enviar_tabla_destino(
            subconjunto, destino, token, titulo, max_filas, plano, estilo
        )
        resultados[destino["nombre"]] = enviados
        log.info("  '%s': %d mensaje(s) enviado(s)", destino["nombre"], enviados)

    return resultados


def enviar_tabla(
    df: pd.DataFrame,
    titulo: str = "ESTADO DE CASOS",
    max_filas: int = FILAS_POR_MENSAJE,
    plano: bool = False,
    estilo: str | None = None,
) -> int:
    """
    Envia la tabla a TODOS los destinos configurados, dividiendo en varios
    mensajes si es necesario. Cada destino recibe SOLO los casos de su filtro.

    Returns:
        Cantidad total de mensajes enviados con exito (sumando todos los destinos).
    """
    resultados = enviar_tabla_por_region(
        df, max_filas=max_filas, plano=plano, estilo=estilo
    )
    total = sum(resultados.values())
    log.info(
        "Total de mensajes enviados a Telegram: %d (destinos: %d)",
        total, len(resultados),
    )
    return total


# ---------------------------------------------------------------------------
# Consola
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Envio de notificaciones y tablas a Telegram (Torre de Control SLA)."
    )
    parser.add_argument("--test", action="store_true",
                        help="Envia un mensaje de prueba a TODOS los destinos.")
    parser.add_argument("--verificar", action="store_true",
                        help="Comprueba el token contra la API de Telegram sin enviar nada.")
    parser.add_argument("--destinos", action="store_true",
                        help="Lista los destinos configurados (grupos/personas).")
    parser.add_argument("--chats", action="store_true",
                        help="Muestra los chats que el bot ha visto, para obtener el chat_id.")
    parser.add_argument("--tabla", action="store_true",
                        help="Envia la tabla de casos de la ventana de notificacion.")
    parser.add_argument("--grupo", type=str, default=None,
                        help="Envia solo a los destinos cuyo nombre contenga este texto "
                             "(ej. --grupo bogota).")
    parser.add_argument("--filtro", choices=[FILTRO_TODO, FILTRO_BOGOTA, FILTRO_REGIONALES],
                        default=None,
                        help="Fuerza el filtro de region de los casos a enviar.")
    parser.add_argument("--plano", action="store_true",
                        help="Usa formato de texto plano (para reenviar a WhatsApp).")
    parser.add_argument("--estilo", choices=list(ESTILOS), default=None,
                        help="Estilo de presentacion: tabla (alineada, ideal para copiar), "
                             "tarjetas (mas legible en el movil), por-estado (para priorizar), "
                             "markdown (para Excel/Word), resumen (solo totales).")
    parser.add_argument("--listar-estilos", action="store_true",
                        help="Muestra los estilos disponibles y sale.")
    parser.add_argument("--max-filas", type=int, default=FILAS_POR_MENSAJE,
                        help=f"Filas por mensaje (por defecto {FILAS_POR_MENSAJE}).")
    parser.add_argument("--excel", type=str, default=None, help="Ruta del Excel.")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    # --- Estilos disponibles ----------------------------------------------
    if args.listar_estilos:
        print("\nEstilos de presentacion disponibles:\n")
        print(f"  {'tabla':<12} Bloque alineado. Ideal para COPIAR a WhatsApp (por defecto).")
        print(f"  {'tarjetas':<12} Una tarjeta por caso con emojis. Lo mas legible en el movil.")
        print(f"  {'por-estado':<12} Agrupado por semaforo. Ideal para PRIORIZAR.")
        print(f"  {'markdown':<12} Tabla con tuberias. Para pegar en Excel, Word o correo.")
        print(f"  {'resumen':<12} Solo totales por estado y region. Para jefatura.")
        print("\nEjemplo:  python telegram_notifier.py --tabla --estilo tarjetas\n")
        return 0

    if not telegram_configurado():
        try:
            cargar_config()
        except ValueError as exc:
            print(f"\n[ERROR] {exc}\n")
            return 2

    # --- Diagnostico: listar destinos -------------------------------------
    if args.destinos:
        destinos = destinos_configurados()
        print(f"\nDestinos configurados: {len(destinos)}")
        for i, d in enumerate(destinos, 1):
            tema = f" (tema {d['tema_id']})" if d.get("tema_id") else ""
            filtro = d.get("filtro", FILTRO_TODO)
            etiqueta = f"  [filtro: {filtro}]" if filtro != FILTRO_TODO else "  [todos los casos]"
            print(f"  {i}. {d['nombre']}")
            print(f"       chat_id: {d['chat_id']}{tema}{etiqueta}")
        print()
        return 0

    # --- Diagnostico: verificar el bot ------------------------------------
    if args.verificar:
        info = verificar_bot()
        destinos = destinos_configurados()
        print()
        if info["ok"]:
            print(f"[OK] El bot responde: {info['bot']}")
            print(f"[OK] Destinos configurados: {len(destinos)}")
            for d in destinos:
                print(f"       - {d['nombre']}: {d['chat_id']}")
            print("\nSiguiente paso: python telegram_notifier.py --test")
        else:
            print(f"[FALLA] No se pudo validar el bot: {info['error']}")
            print("        Revise el token en docs/TUTORIAL_TELEGRAM.md")
        print()
        return 0 if info["ok"] else 1

    # --- Diagnostico: chats vistos por el bot (para obtener el chat_id) -----
    if args.chats:
        chats = obtener_chat_ids_recientes()
        print()
        if not chats:
            print("El bot no ha visto ningun chat todavia.")
            print()
            print("Un bot NO puede iniciar una conversacion: necesita recibir el")
            print("primer mensaje. Elija el caso que le corresponda.")
            print()
            print("--- CASO A: la torre de control (una o varias personas) ---")
            print("  1. La persona abre Telegram y busca el bot:  @TC_Colsof_Bac_bot")
            print("  2. Pulsa START (o escribe 'hola')")
            print("  3. Vuelva a ejecutar:  python telegram_notifier.py --chats")
            print()
            print("--- CASO B: un grupo (recomendado si son varias personas) ---")
            print("  1. Cree el grupo y agregue a las personas")
            print("  2. Agregue el bot:  menu del grupo -> Anadir miembros -> busque el bot")
            print("  3. Envie cualquier mensaje AL GRUPO")
            print("  4. Vuelva a ejecutar:  python telegram_notifier.py --chats")
            print()
            print("NOTA: el numero de telefono NO sirve como chat_id. Telegram usa un")
            print("      id interno (persona: positivo, grupo: empieza por -100).")
            print("      Vea docs/COMO_ENVIAR_A_UN_NUMERO.md")
            print()
            return 1
        print(f"Chats vistos por el bot ({len(chats)}):")
        print()
        for c in chats:
            tipo = c["tipo"]
            etiqueta = {
                "group": "GRUPO", "supergroup": "GRUPO", "private": "PERSONA",
                "channel": "CANAL",
            }.get(tipo, tipo.upper())
            print(f"  [{etiqueta}] {c['nombre']}")
            print(f"            chat_id: {c['chat_id']}")
        print()
        print("Copie el chat_id deseado en config_telegram.json -> destinos")
        print("  - PERSONA -> id positivo (ej. 8522163745)")
        print("  - GRUPO   -> id que empieza por -100 (copielo tal cual)")
        print()
        return 0

    # --- Prueba de envio --------------------------------------------------
    if args.test:
        destinos = destinos_configurados()
        print(f"\nEnviando prueba a {len(destinos)} destino(s)...")
        n = enviar_a_todos(
            "<b>✅ Prueba de conexion</b>\n"
            "Torre de Control SLA — Colsof / Banco Agrario\n"
            f"<i>{datetime.now():%d/%m/%Y %H:%M:%S}</i>\n\n"
            "Si ve este mensaje, el bot esta bien configurado."
        )
        print(f"\nDestinos alcanzados: {n}/{len(destinos)}")
        return 0 if n else 1

    if args.tabla:
        from core import calcular_tablero

        resultado = calcular_tablero(args.excel)
        pendientes = resultado.requieren_atencion
        if pendientes.empty:
            print("No hay casos que requieran atencion.")
            return 0

        if args.filtro:
            pendientes = filtrar_casos_por_filtro(pendientes, args.filtro)
            print(f"Filtro '{args.filtro}': {len(pendientes)} caso(s)")
            if pendientes.empty:
                print("Ningun caso coincide con el filtro.")
                return 0

        resultados = enviar_tabla_por_region(
            pendientes, max_filas=args.max_filas, plano=args.plano,
            solo_destino=args.grupo, estilo=args.estilo,
        )
        print()
        for nombre, n in resultados.items():
            print(f"  {nombre}: {n} mensaje(s)")
        total = sum(resultados.values())
        print(f"\nMensajes enviados: {total}")
        return 0 if total else 1

    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
