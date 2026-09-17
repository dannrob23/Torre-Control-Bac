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

    return "\n\n" + ("─" * 34) + "\n\n".join(
        componer_aviso(f, canal, menciones).replace(
            f" - CASO {str(f.get('N° DE CASO', 'S/N')).strip()}", ""
        )
        for _, f in df_tecnico.iterrows()
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
