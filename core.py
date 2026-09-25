"""
core.py - Nucleo de la Torre de Control de SLA (Colsof - Banco Agrario).

Contiene toda la logica de negocio reutilizable por:
    - alertas_windows.py  (notificaciones de escritorio via plyer)
    - dashboard.py        (dashboard interactivo en Streamlit)

Responsabilidades:
    1. Localizar y leer el Excel de seguimiento de casos  (robusto ante archivo abierto/bloqueado)
    2. Normalizar nombres de tecnicos (tildes, mayusculas, espacios) y cruzar la region
    3. Calcular el tiempo restante contra datetime.now()
    4. Clasificar el semaforo: ROJO / AMARILLO / VERDE / SIN VENCIMIENTO
"""

from __future__ import annotations

import glob
import logging
import os
import sys
import time
import unicodedata
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone

if sys.platform == "win32":
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import pandas as pd

# ---------------------------------------------------------------------------
# Configuracion
# ---------------------------------------------------------------------------

# Nombre esperado del archivo. Si no coincide exactamente se autodescubre por
# patron (el nombre real puede variar en espacios o en el mes).
NOMBRE_ARCHIVO_EXCEL = "PLANTILLA DE SEGUIMIENTO DE CASOS SEPTIEMBRE.xlsx"
HOJA_EXCEL = "PLANTILLA"
PATRON_EXCEL = "*SEGUIMIENTO*CASOS*.xlsx"

# Umbrales del semaforo, en horas restantes.
HORAS_CRITICO = 1.0  # menos de 1 hora  -> ROJO
HORAS_ALERTA = 4.0   # entre 1 y 4 horas -> AMARILLO ; mas de 4 -> VERDE

# --- Version 2: ventana de notificacion ---
DIAS_VENTANA = 3

# Modos de ventana (que casos entran a la tabla y a las notificaciones):
#
#   MODO_DIAS   (POR DEFECTO) - "ultimos N dias calendario, acotado".
#       Incluye solo los casos cuya FECHA DE VENCIMIENTO cae entre
#       (hoy - N + 1) y hoy, ambos inclusive.  Con N=3 y hoy=17/09 cubre
#       EXACTAMENTE el 15, 16 y 17 de septiembre, y NADA mas.
#       Es el modo que pidio la torre de control para revisar el dia.
#
#   MODO_ACUMULADO - "proximos N dias + todos los vencidos sin cerrar".
#       Incluye lo que vence en los proximos N dias y ADEMAS todos los vencidos
#       sin importar su antiguedad. Sirve para barridos de cartera completa,
#       pero muestra fechas antiguas (ej. casos del 1 y 2 de septiembre).
MODO_DIAS = "dias"
MODO_ACUMULADO = "acumulado"
MODO_VENTANA_POR_DEFECTO = MODO_DIAS

# Nombres canonicos de las columnas (con sus espacios reales tal como vienen en el Excel).
COL_CASO = "N° DE CASO"
COL_VENCIMIENTO = "HORA Y FECHA DE VENCIMIENTO DE LA HERRAMIENTA"
COL_RESUELTO = "FECHA/ HORA QUE SE ATENDIO Y SE DIO POR RESUELTO"
COL_TECNICO = "TECNICO COLSOF ASIGNADO INICIALMENTE"
COL_TECNICO_RESOLUTOR = "TECNICO COLSOF RESOLUTOR"
COL_CIUDAD = "OFICINA"
COL_REGIONAL = "REGIONAL"
COL_DEPARTAMENTO = "DEPARTAMENTO"
COL_ANS = "ANS"
COL_ANS_INCIDENTE = "ANS INCIDENTE"
COL_ANS_REQUERIMIENTO = "ANS REQUERIMIENTO"
COL_CREACION = "FECHA CREACION CASO POR BANCO"

# --- Estados del semaforo (v2, sin solapamiento) ---------------------------
# Activos: se clasifican por tiempo restante al vencimiento.
VERDE = "VERDE"              # > 4 h restantes
AMARILLO = "AMARILLO"        # entre 1 h y 4 h
NARANJA = "NARANJA"          # entre 0 y 1 h
ROJO = "ROJO"                # vencido sin cerrar
# Cerrados: se clasifican comparando resolucion contra vencimiento.
CERRADO_OK = "CERRADO OK"    # cerrado a tiempo  (resolucion <= vencimiento)
CERRADO_TARDE = "CERRADO TARDE"  # incumplimiento (resolucion > vencimiento)
SIN_VENCIMIENTO = "SIN VENCIMIENTO"  # activo sin fecha de vencimiento en la plantilla
# Fila cuyo N° DE CASO aparece MAS DE UNA VEZ en la plantilla (duplicado).
# No es un estado del SLA: es un problema de calidad de datos. Las filas marcadas
# asi NO se cuentan en el semaforo ni en las metricas (ver marcar_duplicados).
DUPLICADO = "DUPLICADO"

# Orden de gravedad: primero lo que exige accion inmediata.
ORDEN_ESTADO = {
    ROJO: 0,           # vencido abierto: lo mas grave
    CERRADO_TARDE: 1,  # incumplimiento ya consumado
    NARANJA: 2,
    AMARILLO: 3,
    VERDE: 4,
    SIN_VENCIMIENTO: 5,
    CERRADO_OK: 6,     # cumplido: al final
    DUPLICADO: 7,      # fila repetida: se muestra, pero fuera de los conteos
}

ICONO_ESTADO = {
    ROJO: "🔴",
    CERRADO_TARDE: "⛔",
    NARANJA: "🟠",
    AMARILLO: "🟡",
    VERDE: "🟢",
    SIN_VENCIMIENTO: "⚪",
    CERRADO_OK: "✅",
    DUPLICADO: "🔁",
}

# Estados que requieren atencion / se notifican.
ESTADOS_ALERTA = (ROJO, CERRADO_TARDE, NARANJA, AMARILLO)
# Todos los estados en orden de presentacion.
TODOS_LOS_ESTADOS = tuple(ORDEN_ESTADO.keys())

# ---------------------------------------------------------------------------
# Diccionario oficial de tecnicos -> region
# ---------------------------------------------------------------------------

TECNICOS_REGION = {
    "JESUS ALFREDO JIMENEZ": "BOGOTA", "JEAN RAFAEL PIÑEROS GARAVITO": "BOGOTA",
    "JOHAN SEBASTIAN MORENO VALENCIA": "BOGOTA", "SANTIAGO ANDRES NOEL JIMENEZ": "BOGOTA",
    "IBETH KARINA ACOSTA CRUZ": "BOGOTA", "JORGE ALEXANDER GUIO OROZCO": "BOGOTA",
    "DAVID DABINSON CARDENAS AMARIS": "BOGOTA", "BRAYAN DUVAN ALVAREZ VEGA": "BOGOTA",
    "JAIDER ALEJANDRO BONILLA": "BOGOTA", "YEISON YAMIR IBARGUEN LEUDO": "BOGOTA",
    "CARLOS MIGUEL JIMENEZ SERRATO": "BUCARAMANGA", "CARLOS ANDRES BRAVO MARTINEZ": "CALI",
    "DIEGO ALEJANDRO HERNANDEZ OROZCO": "MANIZALES", "VICTOR ALFONSO ULLOA ACOSTA": "NEIVA",
    "CARLOS ENRIQUE CUCHIVAGUEN FORERO": "TUNJA", "WILLIAM FERNANDO ALDANA NIÑO": "IBAGUÉ",
    "ANDREY TIMOTHY RODRIGUEZ MARTINEZ": "VALLEDUPAR", "JOSE LUIS CONTRERAS VALERA": "BARRANQUILLA",
    "MABEL BOTINA IBARRA": "POPAYAN", "CRISTIAN VARGAS RODRIGUEZ": "CÚCUTA",
    "HAROL ARLEY BETANCUR MARTINEZ": "MEDELLÍN", "JONATHAN FELIPE JARABA CHALARCA": "MEDELLÍN",
    "CRISTIAN CAMILO TOBON GONZALEZ": "MEDELLÍN", "CESAR AUGUSTO PORRAS MONCAYO": "VILLAVICENCIO",
    "MIGUEL ANGEL VEGA ARRIETA": "MONTERIA", "DANIEL ANDRES TORRES OVIEDO": "SINCELEJO",
    "JIMMY FABIEN ERAZO CERÓN": "PASTO",
}

# Alias para variantes reales detectadas en la plantilla que no coinciden
# literalmente con el diccionario oficial.
#   - 'DIEGO ALEJANDRO HERNANDEZ'  -> en la plantilla falta el apellido 'OROZCO'
#   - 'JHONATHAN FELIPE JARABA CHALARCA' -> en la plantilla hay un typo ('JHO' por 'JO')
ALIAS_TECNICOS = {
    "DIEGO ALEJANDRO HERNANDEZ": "DIEGO ALEJANDRO HERNANDEZ OROZCO",
    "JHONATHAN FELIPE JARABA CHALARCA": "JONATHAN FELIPE JARABA CHALARCA",
}


def _cargar_regiones_anonimas() -> None:
    """
    Carga el mapa ETIQUETA -> REGION usado en despliegues anonimizados.

    Cuando el tablero se publica en la web, los nombres de los tecnicos se
    reemplazan por etiquetas (TECNICO-01, ...). Para no perder el corte entre
    Bogota y Regionales, el empaquetador guarda la region de cada etiqueta en
    data/regiones_etiquetas.json. Ese archivo NO tiene nombres reales.

    Se ejecuta una sola vez, al importar el modulo.
    """
    import json

    for base in _rutas_candidatas_excel():
        ruta = os.path.join(base, "regiones_etiquetas.json")
        if not os.path.isfile(ruta):
            continue
        try:
            with open(ruta, encoding="utf-8") as fh:
                datos = json.load(fh)
            for etiqueta, region in (datos.get("regiones") or {}).items():
                if etiqueta and region:
                    TECNICOS_REGION.setdefault(str(etiqueta), str(region))
        except (OSError, json.JSONDecodeError):
            pass
        break

REGION_DESCONOCIDA = "SIN REGION"

# Colombia (America/Bogota) usa UTC-5 todo el anio, sin horario de verano.
# Se usa un offset fijo en lugar de zoneinfo para no depender de la base de
# datos de zonas horarias (tzdata), que no esta instalada en Streamlit Cloud.
ZONA_COLOMBIA = timezone(timedelta(hours=-5))


def ahora_colombia() -> datetime:
    """Hora local de Colombia, naive y compatible con fechas de Excel."""
    return datetime.now(ZONA_COLOMBIA).replace(tzinfo=None)


log = logging.getLogger("torre_control")

# openpyxl emite este warning por validaciones de datos del Excel; es ruido, no un error.
warnings.filterwarnings("ignore", message=".*Data Validation extension.*")


# ---------------------------------------------------------------------------
# Normalizacion de texto y nombres
# ---------------------------------------------------------------------------

def normalizar_texto(valor) -> str:
    """
    Normaliza un texto para comparaciones robustas:
      - trim de espacios (incluye NBSP \xa0)
      - colapsa espacios internos multiples
      - elimina tildes/diacriticos (PIÑEROS -> PINEROS)
      - pasa a mayusculas y quita puntos

    Nota: no se elimina la 'Ñ' como letra, solo su tilde; 'NIÑO' -> 'NINO',
    lo que es consistente para ambos lados de la comparacion.
    """
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    texto = str(valor).replace("\xa0", " ").strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.upper().replace(".", "")
    return " ".join(texto.split())


# Indice de regiones normalizado -> region canonica
_MAPA_REGION_NORMALIZADO = {normalizar_texto(k): v for k, v in TECNICOS_REGION.items()}
# Alias normalizados -> nombre canonico del diccionario
_MAPA_ALIAS_NORMALIZADO = {normalizar_texto(k): v for k, v in ALIAS_TECNICOS.items()}


def obtener_region(nombre_tecnico) -> str:
    """
    Devuelve la region del tecnico. Cadena vacia / desconocido => 'SIN REGION'.
    Aplica trim + normalizacion + tabla de alias antes de buscar.
    """
    clave = normalizar_texto(nombre_tecnico)
    if not clave:
        return REGION_DESCONOCIDA

    # Resolver alias hacia el nombre canonico del diccionario
    clave = normalizar_texto(_MAPA_ALIAS_NORMALIZADO.get(clave, clave))

    return _MAPA_REGION_NORMALIZADO.get(clave, REGION_DESCONOCIDA)


def tecnicos_sin_region(df: pd.DataFrame) -> list[str]:
    """Lista de tecnicos presentes en el DataFrame que no se pudieron mapear."""
    if df.empty or "REGION_TECNICO" not in df.columns:
        return []
    crudos = df.loc[df["REGION_TECNICO"] == REGION_DESCONOCIDA, "TECNICO"].dropna().unique()
    return sorted({str(t).strip() for t in crudos if str(t).strip()})


# ---------------------------------------------------------------------------
# Localizacion y lectura robusta del Excel
# ---------------------------------------------------------------------------

class ErrorLecturaExcel(Exception):
    """Error controlado al abrir/leer la plantilla, con mensaje apto para el usuario."""


def _rutas_candidatas_excel() -> list[str]:
    """
    Carpetas donde se busca la plantilla, en orden de prioridad.

    Incluye 'data/', 'datos/' y '.streamlit/' porque en un despliegue web
    (Streamlit Cloud) es habitual guardar los archivos de datos en subcarpetas,
    no junto al codigo.
    """
    base = os.path.dirname(os.path.abspath(__file__))
    subcarpetas = ("data", "datos", "datos_entrada", ".streamlit", "plantillas")
    return [base] + [os.path.join(base, s) for s in subcarpetas]


def localizar_excel(ruta: str | None = None) -> str:
    """
    Resuelve la ruta del Excel. Orden de busqueda:
      1. ruta explicita (argumento)
      2. variable de entorno TORRE_EXCEL
      3. secreto de Streamlit (st.secrets["TORRE_EXCEL"]), si esta disponible
      4. nombre esperado o patron *SEGUIMIENTO*CASOS*.xlsx en la carpeta del
         codigo y en las subcarpetas habituales (data/, datos/, ...)
    """
    if ruta is None:
        ruta = os.environ.get("TORRE_EXCEL")

    # Secreto de Streamlit: se consulta sin fallar si no estamos en Streamlit.
    if ruta is None:
        try:
            import streamlit as st  # import local: core no depende de Streamlit

            if "TORRE_EXCEL" in st.secrets:
                ruta = str(st.secrets["TORRE_EXCEL"])
        except Exception:
            pass

    if ruta:
        if not os.path.isfile(ruta):
            raise ErrorLecturaExcel(f"No existe el archivo indicado: {ruta}")
        return ruta

    bases = _rutas_candidatas_excel()

    # 1. Nombre exacto esperado
    for base in bases:
        candidato = os.path.join(base, NOMBRE_ARCHIVO_EXCEL)
        if os.path.isfile(candidato):
            return candidato

    # 2. Autodescubrimiento por patron (ignora temporales ~$ de Excel)
    encontrados: list[str] = []
    for base in bases:
        encontrados.extend(
            f for f in glob.glob(os.path.join(base, PATRON_EXCEL))
            if not os.path.basename(f).startswith("~$")
        )
    if not encontrados:
        raise ErrorLecturaExcel(
            f"No se encontro la plantilla '{NOMBRE_ARCHIVO_EXCEL}'.\n"
            f"Carpetas buscadas:\n  "
            + "\n  ".join(bases)
            + "\n\nUbique el archivo en una de esas carpetas, o defina la variable "
            "de entorno TORRE_EXCEL, o el secreto TORRE_EXCEL en Streamlit."
        )

    # Si hay varios, se toma el modificado mas recientemente
    return max(encontrados, key=os.path.getmtime)


def _limpiar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza los encabezados: trim (incluye NBSP) y colapso de espacios."""
    df = df.copy()
    df.columns = [
        " ".join(str(c).replace("\xa0", " ").split()) if c is not None else ""
        for c in df.columns
    ]
    return df


def _a_datetime(serie: pd.Series) -> pd.Series:
    """
    Convierte una columna mixta (datetime / texto / numeros de Excel) a datetime.

    Se usa format="mixed" porque la MISMA columna puede traer formatos distintos
    (p. ej. "2026-09-17 10:00" y "17/09/2026 15:00"). Sin esta opcion pandas
    infiere un unico formato del primer valor y devuelve NaT en el resto, lo que
    haria que casos validos aparecieran como "SIN VENCIMIENTO".
    """
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    try:
        return pd.to_datetime(serie, errors="coerce", dayfirst=True, format="mixed")
    except (ValueError, TypeError):
        # Respaldo para versiones de pandas que no soportan format="mixed".
        return pd.to_datetime(serie, errors="coerce", dayfirst=True)


def leer_casos(
    ruta: str | None = None,
    *,
    reintentos: int = 3,
    espera_segundos: float = 2.0,
    hoja: str = HOJA_EXCEL,
) -> pd.DataFrame:
    """
    Lee la hoja PLANTILLA y devuelve el DataFrame crudo con columnas limpias.

    Manejo de errores:
      - PermissionError  -> el Excel esta abierto/bloqueado por otro usuario: reintenta
      - FileNotFoundError-> ruta inexistente
      - Otros errores de openpyxl/zip -> archivo corrupto o no es un xlsx valido

    Lanza ErrorLecturaExcel con un mensaje claro si agota los reintentos.
    """
    archivo = localizar_excel(ruta)
    ultimo_error: Exception | None = None

    for intento in range(1, reintentos + 1):
        try:
            # Se lee todo en memoria y se libera el handle de inmediato.
            with open(archivo, "rb") as fh:
                df = pd.read_excel(fh, sheet_name=hoja, engine="openpyxl")
            log.info("Excel leido correctamente: %s (%d filas)", archivo, len(df))
            return _limpiar_columnas(df)

        except PermissionError as exc:
            ultimo_error = exc
            log.warning(
                "Intento %d/%d: el archivo esta en uso (PermissionError): %s",
                intento, reintentos, archivo,
            )
            if intento < reintentos:
                time.sleep(espera_segundos)

        except FileNotFoundError as exc:
            raise ErrorLecturaExcel(f"No se encontro el archivo: {archivo}") from exc

        except Exception as exc:  # archivo corrupto, hoja inexistente, etc.
            raise ErrorLecturaExcel(
                f"No se pudo leer '{os.path.basename(archivo)}' "
                f"(hoja '{hoja}'): {type(exc).__name__}: {exc}"
            ) from exc

    raise ErrorLecturaExcel(
        f"El archivo '{os.path.basename(archivo)}' esta abierto o bloqueado por otro "
        f"programa (Excel). Cierrelo e intente de nuevo.\nDetalle: {ultimo_error}"
    )


def _validar_columnas(df: pd.DataFrame) -> None:
    """Verifica que existan las columnas minimas requeridas."""
    requeridas = [COL_RESUELTO, COL_VENCIMIENTO, COL_TECNICO]
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ErrorLecturaExcel(
            "La hoja no contiene las columnas esperadas. Faltan: "
            + "; ".join(repr(c) for c in faltantes)
            + "\nColumnas encontradas: "
            + "; ".join(repr(c) for c in df.columns)
        )



# ---------------------------------------------------------------------------
# Duplicados de la plantilla (el mismo N° DE CASO en varias filas)
# ---------------------------------------------------------------------------

# Columnas que agrega marcar_duplicados().
COL_FILA_EXCEL = "FILA_EXCEL"
COL_FILAS_REPETIDAS = "FILAS_REPETIDAS"
COL_ES_DUPLICADO = "ES_DUPLICADO"
COL_FILAS_DEL_CASO = "FILAS_DEL_CASO"
COL_ESTADO_CALCULADO = "ESTADO_CALCULADO"


def clave_caso(valor) -> str:
    """
    Clave de comparacion de un N° DE CASO.

    La plantilla trae el mismo caso con espacios, saltos de linea o NBSP de mas.
    Se normaliza (NBSP -> espacio, trim, colapso de espacios, mayusculas) para que
    dos filas del MISMO caso se reconozcan como repetidas.

    Devuelve "" cuando la celda esta vacia: una fila sin numero de caso NO cuenta
    como duplicado (es otra cosa: fila en blanco).
    """
    if valor is None or pd.isna(valor):
        return ""
    texto = str(valor).replace("\xa0", " ").strip().upper()
    return " ".join(texto.split())


def marcar_duplicados(marco: pd.DataFrame, col_caso: str = COL_CASO) -> pd.DataFrame:
    """
    Marca las filas cuyo N° DE CASO aparece MAS DE UNA VEZ en la plantilla.

    Por que existe: la plantilla real tiene casos repetidos porque alguien agrega
    una fila nueva para registrar el cierre en vez de completar la original.
    Caso real: IM3238158 aparece dos veces, una SIN fecha de resolucion (abierto y
    vencido -> ROJO) y otra cerrada el 24/09 (-> CERRADO TARDE). Sin esta marca, la
    misma fila contaba como ROJO y como CERRADO TARDE a la vez, disparaba alertas
    de un caso ya cerrado e inflaba los conteos.

    NO borra ni elige ninguna fila: solo avisa. Agrega las columnas:
        FILA_EXCEL      -> fila que ocupa en el Excel (2 = primera fila de datos)
        FILAS_REPETIDAS -> cuantas veces aparece ese N° DE CASO
        ES_DUPLICADO    -> True si aparece mas de una vez
        FILAS_DEL_CASO  -> texto con TODAS las filas Excel de ese caso repetido
    """
    marco = marco.copy()
    marco[COL_FILA_EXCEL] = range(2, len(marco) + 2)

    if col_caso not in marco.columns:
        marco[COL_FILAS_REPETIDAS] = 1
        marco[COL_ES_DUPLICADO] = False
        marco[COL_FILAS_DEL_CASO] = ""
        return marco

    clave = marco[col_caso].apply(clave_caso)
    conteo = clave[clave != ""].value_counts()
    # Las filas sin numero de caso quedan en 1: no son duplicados.
    marco[COL_FILAS_REPETIDAS] = clave.map(conteo).fillna(1).astype(int)
    marco[COL_ES_DUPLICADO] = marco[COL_FILAS_REPETIDAS] > 1
    marco[COL_FILAS_DEL_CASO] = clave.map(
        {
            cl: ", ".join(str(int(f)) for f in sorted(grupo[COL_FILA_EXCEL]))
            for cl, grupo in marco[marco[COL_ES_DUPLICADO]].groupby(
                clave[marco[COL_ES_DUPLICADO].index]
            )
        }
    ).fillna("")
    return marco


def resumen_duplicados(df_completo: pd.DataFrame) -> pd.DataFrame:
    """
    Tabla legible de los casos repetidos, para el aviso del tablero.

    Una fila por caso repetido, con sus filas del Excel y el estado que le habria
    correspondido a cada copia (antes de marcarla DUPLICADO). Asi se ve de un golpe
    POR QUE el caso parecia estar abierto y cerrado a la vez.
    """
    columnas = ["CASO", "VECES", "FILAS_EXCEL", "ESTADO_POR_FILA", "TECNICO", "VENCIMIENTO"]
    if (df_completo is None or df_completo.empty
            or COL_ES_DUPLICADO not in df_completo.columns):
        return pd.DataFrame(columns=columnas)

    repetidas = df_completo[df_completo[COL_ES_DUPLICADO].fillna(False).astype(bool)]
    if repetidas.empty:
        return pd.DataFrame(columns=columnas)

    clave = repetidas[COL_CASO].apply(clave_caso)
    filas = []
    for cl, grupo in repetidas.groupby(clave):
        grupo = grupo.sort_values(COL_FILA_EXCEL)
        estados = " | ".join(
            "fila {}: {}".format(
                int(f[COL_FILA_EXCEL]),
                f.get(COL_ESTADO_CALCULADO) or f.get("ESTADO") or "?",
            )
            for _, f in grupo.iterrows()
        )
        primera = grupo.iloc[0]
        vencimiento = primera.get("FECHA_VENCIMIENTO")
        filas.append({
            "CASO": cl,
            "VECES": int(len(grupo)),
            "FILAS_EXCEL": str(primera.get(COL_FILAS_DEL_CASO, "")),
            "ESTADO_POR_FILA": estados,
            "TECNICO": str(primera.get("TECNICO", "") or ""),
            "VENCIMIENTO": (
                pd.Timestamp(vencimiento).strftime("%d/%m/%Y %H:%M")
                if pd.notna(vencimiento) else "sin fecha"
            ),
        })
    return (
        pd.DataFrame(filas, columns=columnas)
        .sort_values("CASO")
        .reset_index(drop=True)
    )


# ---------------------------------------------------------------------------
# Clasificacion del semaforo
# ---------------------------------------------------------------------------

def clasificar_semaforo(horas_restantes: float | None) -> str:
    """
    Clasifica un caso ABIERTO segun su tiempo restante al vencimiento.

        None              -> SIN VENCIMIENTO  (la plantilla no trae fecha)
        <= 0              -> ROJO      (ya vencido)
        0 < h <= 1        -> NARANJA   (menos de 1 hora)
        1 < h <= 4        -> AMARILLO
        > 4 horas         -> VERDE

    Nota: los umbrales son fijos (4 h / 1 h). El ANS por tipo de caso quedo
    pendiente por decision del usuario; cuando se active, el calculo del
    porcentaje de consumo se hara sobre el ANS aplicable (columnas G/H).
    """
    if horas_restantes is None or pd.isna(horas_restantes):
        return SIN_VENCIMIENTO
    if horas_restantes <= 0:
        return ROJO
    if horas_restantes <= HORAS_CRITICO:
        return NARANJA
    if horas_restantes <= HORAS_ALERTA:
        return AMARILLO
    return VERDE


def clasificar_cierre(horas_desviacion: float | None) -> str:
    """
    Clasifica un caso CERRADO comparando resolucion contra vencimiento.

        None  -> SIN VENCIMIENTO (cerrado pero sin fecha de vencimiento)
        <= 0  -> CERRADO OK      (se cerro a tiempo)
        > 0   -> CERRADO TARDE   (incumplimiento: se cerro despues de vencer)
    """
    if horas_desviacion is None or pd.isna(horas_desviacion):
        return SIN_VENCIMIENTO
    return CERRADO_OK if horas_desviacion <= 0 else CERRADO_TARDE


def formatear_duracion(horas: float | None) -> str:
    """
    Convierte horas decimales a texto legible.
    Ej: 4.42 -> '4 h 25 min'    |    -3.5 -> '3 h 30 min'
    """
    if horas is None or pd.isna(horas):
        return "—"
    total_minutos = int(round(abs(horas) * 60))
    horas_ent, minutos = divmod(total_minutos, 60)
    if horas_ent and minutos:
        return f"{horas_ent} h {minutos} min"
    if horas_ent:
        return f"{horas_ent} h"
    return f"{minutos} min"


def formatear_horas(horas: float | None) -> str:
    """
    Texto legible del tiempo restante.
    Ej: '4 h 25 min restantes'  |  'Vencido hace 384 h 12 min'
    """
    if horas is None or pd.isna(horas):
        return "Sin fecha de vencimiento"
    duracion = formatear_duracion(horas)
    if horas < 0:
        return f"Vencido hace {duracion}"
    return f"{duracion} restantes"


def formatear_prediccion(horas: float | None, etiqueta: str) -> str:
    """
    Texto de prediccion: cuanto falta para que el caso cambie de color.
    Ej: 'Pasa a AMARILLO en 2 h 10 min'
    """
    if horas is None or pd.isna(horas):
        return "—"
    if horas <= 0:
        return f"Ya en {etiqueta}"
    return f"Pasa a {etiqueta} en {formatear_duracion(horas)}"


# ---------------------------------------------------------------------------
# Construccion del tablero
# ---------------------------------------------------------------------------

@dataclass
class Resultado:
    """
    Contenedor del calculo de SLA (v2).

    df_completo : TODOS los casos (activos + cerrados), con metricas.
                  Base para el dashboard de metricas por tecnico.
    df          : Vista de trabajo: solo los casos DENTRO de la ventana de
                  notificacion (vencidos + proximos DIAS_VENTANA), ordenados
                  por gravedad. Es lo que usan alertas y notificaciones.
    """
    df_completo: pd.DataFrame
    df: pd.DataFrame
    momentos: datetime
    total_hoja: int = 0
    total_activos: int = 0
    total_en_ventana: int = 0
    dias_ventana: int = DIAS_VENTANA
    warnings: list[str] = field(default_factory=list)

    def _conteo_de(self, marco: pd.DataFrame) -> dict[str, int]:
        if marco is None or marco.empty:
            return {k: 0 for k in TODOS_LOS_ESTADOS}
        c = marco["ESTADO"].value_counts().to_dict()
        return {k: int(c.get(k, 0)) for k in TODOS_LOS_ESTADOS}

    @property
    def conteo(self) -> dict[str, int]:
        """Conteo por estado de los casos en la ventana de notificacion."""
        return self._conteo_de(self.df)

    @property
    def conteo_completo(self) -> dict[str, int]:
        """Conteo por estado de TODOS los casos (activos + cerrados)."""
        return self._conteo_de(self.df_completo)

    def _mascara_duplicado(self) -> pd.Series:
        """
        True en las filas cuyo N° DE CASO esta repetido en la plantilla.

        Si el DataFrame no trae la marca (por ejemplo, un resultado viejo en
        cache), se devuelve todo False para no romper el calculo.
        """
        if COL_ES_DUPLICADO not in self.df_completo.columns:
            return pd.Series(False, index=self.df_completo.index)
        return self.df_completo[COL_ES_DUPLICADO].fillna(False).astype(bool)

    @property
    def conteo_activos(self) -> dict[str, int]:
        """Conteo solo de casos ABIERTOS (excluye cerrados y filas repetidas)."""
        if self.df_completo.empty:
            return {k: 0 for k in TODOS_LOS_ESTADOS}
        abiertos = self.df_completo[
            (~self.df_completo["CERRADO"]) & (~self._mascara_duplicado())
        ]
        return self._conteo_de(abiertos)

    @property
    def duplicados(self) -> pd.DataFrame:
        """
        Un renglon por caso repetido (N° DE CASO en varias filas), con sus filas del
        Excel y el estado que le habria correspondido a cada copia.
        """
        return resumen_duplicados(self.df_completo)

    @property
    def df_duplicados(self) -> pd.DataFrame:
        """TODAS las filas marcadas DUPLICADO (las copias, sin borrar ninguna)."""
        if self.df_completo.empty or COL_ES_DUPLICADO not in self.df_completo.columns:
            return self.df_completo
        return self.df_completo[self._mascara_duplicado()].copy()

    @property
    def requieren_atencion(self) -> pd.DataFrame:
        """Casos de la ventana que exigen accion, ya ordenados por gravedad."""
        if self.df.empty:
            return self.df
        return self.df[self.df["ESTADO"].isin(ESTADOS_ALERTA)].copy()

    @property
    def vencidos(self) -> pd.DataFrame:
        """Casos vencidos sin cerrar (el estado mas grave)."""
        if self.df_completo.empty:
            return self.df_completo
        m = self.df_completo
        return m[(~m["CERRADO"]) & (m["ESTADO"] == ROJO)].copy()

    @property
    def proximos_a_vencer(self) -> pd.DataFrame:
        """Casos abiertos que aun no vencen (verde, amarillo o naranja)."""
        if self.df_completo.empty:
            return self.df_completo
        m = self.df_completo
        return m[(~m["CERRADO"]) & (m["ESTADO"].isin([VERDE, AMARILLO, NARANJA]))].copy()

    @property
    def incumplimientos(self) -> pd.DataFrame:
        """Casos que se cerraron despues de vencer."""
        if self.df_completo.empty:
            return self.df_completo
        m = self.df_completo
        return m[(m["CERRADO"]) & (m["ESTADO"] == CERRADO_TARDE)].copy()


def calcular_tablero(
    ruta: str | None = None,
    momento: datetime | None = None,
    df_crudo: pd.DataFrame | None = None,
    dias_ventana: int = DIAS_VENTANA,
    modo_ventana: str = MODO_VENTANA_POR_DEFECTO,
) -> Resultado:
    """
    Pipeline completo (v2):
    lectura -> metricas por caso -> semaforo -> region -> ventana de notificacion.

    Args:
        ruta:         ruta del Excel (None = autodescubrir)
        momento:      fecha/hora de referencia (por defecto hora colombiana)
        df_crudo:     DataFrame ya leido (evita releer el archivo)
        dias_ventana: cuantos dias calendario abarca la ventana
        modo_ventana: MODO_DIAS (por defecto) = ultimos N dias ACOTADO.
                      MODO_ACUMULADO = proximos N dias + todos los vencidos.

    Returns:
        Resultado con df_completo (todo) y df (ventana de notificacion).
    """
    momento = momento or ahora_colombia()
    avisos: list[str] = []

    df = df_crudo if df_crudo is not None else leer_casos(ruta)
    total_hoja = len(df)
    _validar_columnas(df)

    # --- 1. Normalizar columnas de fecha -----------------------------------
    marco = df.copy()
    marco["FECHA_VENCIMIENTO"] = _a_datetime(marco[COL_VENCIMIENTO])
    marco["FECHA_RESOLUCION"] = _a_datetime(marco[COL_RESUELTO])
    marco["FECHA_CREACION"] = _a_datetime(marco.get(COL_CREACION))

    # --- 2. Clasificacion abierto / cerrado --------------------------------
    # Cerrado = la columna O (fecha/hora de resolucion) tiene valor.
    marco["CERRADO"] = marco["FECHA_RESOLUCION"].notna()

    # --- 2b. Duplicados de la plantilla ------------------------------------
    # El mismo N° DE CASO en varias filas (una sin cerrar y otra ya cerrada) hacia
    # que el caso contara A LA VEZ como abierto y como cerrado. Se marcan las copias
    # y quedan FUERA del semaforo y de las metricas; se siguen mostrando para que
    # alguien las unifique en el Excel.
    marco = marcar_duplicados(marco)

    # --- 3. Tecnico y region ----------------------------------------------
    marco["TECNICO"] = marco[COL_TECNICO].apply(
        lambda v: " ".join(str(v).split()) if pd.notna(v) else ""
    )
    marco["REGION_TECNICO"] = marco[COL_TECNICO].apply(obtener_region)
    if COL_TECNICO_RESOLUTOR in marco.columns:
        marco["TECNICO_RESOLUTOR"] = marco[COL_TECNICO_RESOLUTOR].apply(
            lambda v: " ".join(str(v).split()) if pd.notna(v) else ""
        )
    else:
        marco["TECNICO_RESOLUTOR"] = ""

    # --- 4. Metricas de tiempo --------------------------------------------
    ahora = pd.Timestamp(momento)

    # Tiempo restante (solo tiene sentido para abiertos; en cerrados es informativo).
    delta = marco["FECHA_VENCIMIENTO"] - ahora
    marco["HORAS_RESTANTES"] = (delta.dt.total_seconds() / 3600.0).round(2)

    # Horas vencido: cuanto lleva vencido.
    #   - abierto  -> ahora - vencimiento (crece con el tiempo)
    #   - cerrado  -> resolucion - vencimiento (quedo fijo al cerrar)
    transcurrido = ahora - marco["FECHA_VENCIMIENTO"]
    desviacion_cierre = marco["FECHA_RESOLUCION"] - marco["FECHA_VENCIMIENTO"]
    horas_vencido = transcurrido.where(
        ~marco["CERRADO"], desviacion_cierre
    )
    marco["HORAS_VENCIDO"] = (horas_vencido.dt.total_seconds() / 3600.0).round(2)
    # Nunca negativo: un caso cerrado antes de vencer no "estuvo vencido".
    marco.loc[marco["HORAS_VENCIDO"] < 0, "HORAS_VENCIDO"] = 0.0

    # Desviacion al cerrar (negativa = cerro antes de vencer).
    marco["HORAS_DESVIACION_CIERRE"] = (
        desviacion_cierre.dt.total_seconds() / 3600.0
    ).round(2)

    # Predicciones: cuanto falta para el proximo cambio de color (solo abiertos).
    marco["PASA_A_AMARILLO_EN"] = (marco["HORAS_RESTANTES"] - HORAS_ALERTA).round(2)
    marco["PASA_A_NARANJA_EN"] = (marco["HORAS_RESTANTES"] - HORAS_CRITICO).round(2)
    marco["PASA_A_ROJO_EN"] = marco["HORAS_RESTANTES"]

    # --- 5. Semaforo -------------------------------------------------------
    # Abiertos: por tiempo restante.  Cerrados: comparando contra el vencimiento.
    estado_abierto = marco["HORAS_RESTANTES"].apply(clasificar_semaforo)
    estado_cerrado = marco["HORAS_DESVIACION_CIERRE"].apply(clasificar_cierre)
    marco["ESTADO"] = estado_abierto.where(~marco["CERRADO"], estado_cerrado)
    # Se guarda el estado que le habria correspondido a cada fila (asi se ve si la
    # copia estaba abierta o cerrada) y encima las repetidas pasan a DUPLICADO: se
    # siguen viendo en el Explorador, pero no entran en el semaforo ni en las alertas.
    marco[COL_ESTADO_CALCULADO] = marco["ESTADO"]
    marco.loc[marco[COL_ES_DUPLICADO], "ESTADO"] = DUPLICADO
    marco["ICONO"] = marco["ESTADO"].map(ICONO_ESTADO)
    marco["TIEMPO_RESTANTE"] = marco["HORAS_RESTANTES"].apply(formatear_horas)
    marco["TIEMPO_VENCIDO"] = marco["HORAS_VENCIDO"].apply(
        lambda h: formatear_duracion(h) if pd.notna(h) and h > 0 else "—"
    )
    marco["PREDICCION"] = marco.apply(
        lambda f: formatear_prediccion(f["PASA_A_AMARILLO_EN"], AMARILLO)
        if f["ESTADO"] == VERDE
        else (
            formatear_prediccion(f["PASA_A_NARANJA_EN"], NARANJA)
            if f["ESTADO"] == AMARILLO
            else (
                formatear_prediccion(f["PASA_A_ROJO_EN"], ROJO)
                if f["ESTADO"] == NARANJA
                else "—"
            )
        ),
        axis=1,
    )

    # --- 6. Ventana de notificacion ---------------------------------------
    # Ver la explicacion de los modos junto a MODO_DIAS / MODO_ACUMULADO.
    hoy = ahora.normalize()

    if modo_ventana == MODO_ACUMULADO:
        # Proximos N dias calendario + TODOS los vencidos sin cerrar.
        desde = ahora
        hasta = hoy + pd.Timedelta(days=max(dias_ventana - 1, 0)) + pd.Timedelta(days=1)
        es_vencido = (~marco["CERRADO"]) & (marco["FECHA_VENCIMIENTO"] < ahora)
        dentro_de_n_dias = (
            (~marco["CERRADO"])
            & (marco["FECHA_VENCIMIENTO"] >= ahora)
            & (marco["FECHA_VENCIMIENTO"] < hasta)
        )
        marco["EN_VENTANA"] = es_vencido | dentro_de_n_dias
        rango_texto = (
            f"proximos {dias_ventana} dias (hasta {hasta - pd.Timedelta(days=1):%d/%m}) "
            f"+ TODOS los vencidos sin cerrar"
        )
    else:
        # MODO_DIAS (por defecto): ultimos N dias ACOTADO.
        # Solo lo que vence entre (hoy - N + 1) y el final de hoy. Nada anterior.
        desde = hoy - pd.Timedelta(days=max(dias_ventana - 1, 0))
        hasta = hoy + pd.Timedelta(days=1)  # exclusivo: incluye todo el dia de hoy
        en_rango = (
            (marco["FECHA_VENCIMIENTO"] >= desde)
            & (marco["FECHA_VENCIMIENTO"] < hasta)
        )
        # Se muestran TANTO abiertos como cerrados dentro del rango, para que los
        # casos cerrados recientes aparezcan como CERRADO OK / CERRADO TARDE.
        marco["EN_VENTANA"] = en_rango
        rango_texto = (
            f"del {desde:%d/%m} al {hoy:%d/%m} inclusive "
            f"(ultimos {dias_ventana} dias, acotado)"
        )

    # --- 7. Orden por gravedad --------------------------------------------
    marco["_orden"] = marco["ESTADO"].map(ORDEN_ESTADO).fillna(99)
    marco = marco.sort_values(
        by=["_orden", "HORAS_RESTANTES"], ascending=[True, True], na_position="last"
    ).drop(columns=["_orden"])

    en_ventana = marco[marco["EN_VENTANA"]].reset_index(drop=True)
    completo = marco.reset_index(drop=True)

    # --- 8. Avisos de calidad de datos ------------------------------------
    # Las filas repetidas no cuentan como activas: inflarian el total de activos.
    abiertos = completo[(~completo["CERRADO"]) & (~completo[COL_ES_DUPLICADO])]
    sin_fecha = int(abiertos["FECHA_VENCIMIENTO"].isna().sum())
    if sin_fecha:
        avisos.append(
            f"{sin_fecha} caso(s) activo(s) sin fecha de vencimiento: no se pueden "
            "clasificar ni notificar."
        )

    sin_tecnico = int((abiertos["TECNICO"] == "").sum())
    if sin_tecnico:
        avisos.append(
            f"{sin_tecnico} caso(s) activo(s) sin tecnico asignado: quedan como "
            f"'{REGION_DESCONOCIDA}'."
        )

    sin_region = tecnicos_sin_region(completo)
    if sin_region:
        avisos.append(
            "Tecnico(s) con nombre no registrado en el diccionario de regiones: "
            + ", ".join(sin_region)
        )

    repetidos = resumen_duplicados(completo)
    if not repetidos.empty:
        detalle = "; ".join(
            "{} (filas {}: {})".format(
                fila["CASO"], fila["FILAS_EXCEL"], fila["ESTADO_POR_FILA"]
            )
            for _, fila in repetidos.iterrows()
        )
        avisos.append(
            "{} fila(s) de {} caso(s) con el N° DE CASO repetido en la plantilla: se "
            "marcan DUPLICADO y NO se cuentan en el semaforo ni en las metricas (hay "
            "que unificar las filas). Detalle: {}".format(
                int(repetidos["VECES"].sum()), len(repetidos), detalle
            )
        )

    avisos.append(
        f"Ventana de notificacion: {rango_texto}. "
        f"{len(en_ventana)} de {len(marco)} casos quedan dentro "
        f"({len(abiertos)} activos en total)."
    )

    if abiertos.empty:
        avisos.append("No hay casos activos (todos tienen fecha de resolucion).")

    return Resultado(
        df_completo=completo,
        df=en_ventana,
        momentos=momento,
        total_hoja=total_hoja,
        total_activos=len(abiertos),
        total_en_ventana=len(en_ventana),
        dias_ventana=dias_ventana,
        warnings=avisos,
    )


# ---------------------------------------------------------------------------
# Metricas por tecnico (dashboard de notificaciones)
# ---------------------------------------------------------------------------

def metricas_por_tecnico(df_completo: pd.DataFrame) -> pd.DataFrame:
    """
    Tabla de metricas por tecnico, base del dashboard de notificaciones.

    Columnas devueltas:
        TECNICO, REGION, ASIGNADOS, ABIERTOS, VENCIDOS, PROXIMOS_VENCER,
        NARANJA, AMARILLO, CERRADOS, CERRADOS_TARDE, CUMPLIMIENTO_PCT,
        HORAS_VENCIDO_TOTAL, HORAS_VENCIDO_PROMEDIO
    """
    if df_completo is None or df_completo.empty:
        return pd.DataFrame(
            columns=[
                "TECNICO", "REGION", "ASIGNADOS", "ABIERTOS", "VENCIDOS",
                "PROXIMOS_VENCER", "NARANJA", "AMARILLO", "CERRADOS",
                "CERRADOS_TARDE", "CUMPLIMIENTO_PCT", "HORAS_VENCIDO_TOTAL",
                "HORAS_VENCIDO_PROMEDIO",
            ]
        )

    m = df_completo.copy()
    # Las filas repetidas (ver marcar_duplicados) no entran en las metricas: si no,
    # un mismo caso se contaria dos veces a favor o en contra del tecnico.
    if COL_ES_DUPLICADO in m.columns:
        m = m[~m[COL_ES_DUPLICADO].fillna(False).astype(bool)]
    m["_abierto"] = ~m["CERRADO"]

    filas = []
    for tecnico, grupo in m.groupby("TECNICO", dropna=False):
        abiertos = grupo[grupo["_abierto"]]
        cerrados = grupo[~grupo["_abierto"]]
        cerrados_tarde = cerrados[cerrados["ESTADO"] == CERRADO_TARDE]

        vencidos = abiertos[abiertos["ESTADO"] == ROJO]
        # "Proximos a vencer": abiertos que aun no vencieron.
        proximos = abiertos[abiertos["ESTADO"].isin([VERDE, AMARILLO, NARANJA])]

        horas_vencido_total = float(vencidos["HORAS_VENCIDO"].fillna(0).sum())
        filas.append(
            {
                "TECNICO": tecnico if tecnico else "(sin tecnico)",
                "REGION": grupo["REGION_TECNICO"].mode().iloc[0]
                if not grupo["REGION_TECNICO"].mode().empty
                else REGION_DESCONOCIDA,
                "ASIGNADOS": int(len(grupo)),
                "ABIERTOS": int(len(abiertos)),
                "VENCIDOS": int(len(vencidos)),
                "PROXIMOS_VENCER": int(len(proximos)),
                "NARANJA": int((abiertos["ESTADO"] == NARANJA).sum()),
                "AMARILLO": int((abiertos["ESTADO"] == AMARILLO).sum()),
                "CERRADOS": int(len(cerrados)),
                "CERRADOS_TARDE": int(len(cerrados_tarde)),
                "CUMPLIMIENTO_PCT": round(
                    (len(cerrados) - len(cerrados_tarde)) / len(cerrados) * 100, 1
                ) if len(cerrados) else None,
                "HORAS_VENCIDO_TOTAL": round(horas_vencido_total, 1),
                "HORAS_VENCIDO_PROMEDIO": round(
                    horas_vencido_total / len(vencidos), 1
                ) if len(vencidos) else 0.0,
            }
        )

    tabla = pd.DataFrame(filas)
    if tabla.empty:
        return tabla
    # Tecnicos reales primero (alfabetico), "(sin tecnico)" al final.
    tabla["_sin_asignar"] = (tabla["TECNICO"] == "(sin tecnico)").astype(int)
    return (
        tabla.sort_values(["_sin_asignar", "VENCIDOS"], ascending=[True, False])
        .drop(columns=["_sin_asignar"])
        .reset_index(drop=True)
    )


def construir_mensaje_notificacion(fila: pd.Series) -> tuple[str, str]:
    """Devuelve (titulo, cuerpo) del aviso de escritorio para un caso."""
    estado = fila.get("ESTADO", "")
    icono = ICONO_ESTADO.get(estado, "")
    caso = str(fila.get(COL_CASO, "S/N")).strip()
    tecnico = fila.get("TECNICO", "") or "Sin tecnico"
    region = fila.get("REGION_TECNICO", REGION_DESCONOCIDA)
    ciudad = fila.get(COL_CIUDAD, "") or ""
    vencimiento = fila.get("FECHA_VENCIMIENTO")
    prediccion = fila.get("PREDICCION", "—")

    titulo = f"{icono} {estado} - Caso {caso}"

    if estado == ROJO:
        tiempo = "VENCIDO hace " + str(fila.get("TIEMPO_VENCIDO", "—"))
    elif estado == CERRADO_TARDE:
        tiempo = "Cerrado con incumplimiento (vencido " + str(
            fila.get("TIEMPO_VENCIDO", "—")
        ) + ")"
    else:
        tiempo = str(fila.get("TIEMPO_RESTANTE", "—"))

    vence_txt = (
        pd.Timestamp(vencimiento).strftime("%d/%m/%Y %H:%M")
        if pd.notna(vencimiento)
        else "sin fecha"
    )

    cuerpo = (
        f"{tiempo}\n"
        f"Vence: {vence_txt}\n"
        f"Tecnico: {tecnico}\n"
        f"Ciudad: {ciudad} | Region: {region}"
    )
    if estado in (VERDE, AMARILLO, NARANJA) and prediccion != "—":
        cuerpo += f"\n{prediccion}"
    return titulo, cuerpo


# ---------------------------------------------------------------------------
# Separacion por region: BOGOTA vs REGIONALES
# ---------------------------------------------------------------------------

# Etiqueta canonica de Bogota. Todo lo que no sea Bogota se considera "regional",
# que es exactamente la regla pedida por la torre de control.
BOGOTA = "BOGOTA"
REGIONALES = "REGIONALES"


def parte_por_region(casos: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Divide los casos en dos grupos excluyentes:

        BOGOTA    -> casos cuyo REGION_TECNICO es BOGOTA
        REGIONALES -> TODO lo demas (Cali, Medellin, Barranquilla, ...,
                      e incluso los que no tienen region asignada)

    La regla es la pedida: "lo que no diga Bogota es regionales".

    Args:
        casos: DataFrame con la columna REGION_TECNICO (resultado.df o similar).

    Returns:
        {"BOGOTA": DataFrame, "REGIONALES": DataFrame}
    """
    vacio = casos.iloc[0:0] if casos is not None else pd.DataFrame()
    if casos is None or casos.empty:
        return {BOGOTA: vacio, REGIONALES: vacio}

    es_bogota = (
        casos["REGION_TECNICO"].astype(str).str.strip().str.upper() == BOGOTA
    )
    return {
        BOGOTA: casos[es_bogota].copy(),
        REGIONALES: casos[~es_bogota].copy(),
    }


def resumen_consola(resultado: Resultado) -> str:
    """Resumen legible del tablero para consola (texto plano, sin emojis)."""
    lineas = [
        f"Momento de calculo : {resultado.momentos:%Y-%m-%d %H:%M:%S}",
        f"Filas en la hoja   : {resultado.total_hoja}",
        f"Casos activos      : {resultado.total_activos}",
        f"En ventana ({resultado.dias_ventana} dias + vencidos): {resultado.total_en_ventana}",
        "",
        "--- TODOS LOS CASOS (activos + cerrados) ---",
    ]
    for estado, n in resultado.conteo_completo.items():
        lineas.append(f"  {estado:<16}: {n}")
    lineas.append("")
    lineas.append("--- SOLO CASOS ABIERTOS ---")
    for estado, n in resultado.conteo_activos.items():
        if n:
            lineas.append(f"  {estado:<16}: {n}")
    return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Carga del mapa ETIQUETA -> REGION de los despliegues anonimizados.
#
# Va al FINAL del modulo a proposito: necesita que _rutas_candidatas_excel ya
# este definida. Si existe data/regiones_etiquetas.json (lo genera el
# empaquetador de Streamlit), se registran las regiones de las etiquetas
# TECNICO-NN para no perder el corte Bogota vs Regionales.
#
# OJO: hay que REHACER el indice normalizado. Si solo se añaden a
# TECNICOS_REGION, obtener_region() no las vera, porque
# _MAPA_REGION_NORMALIZADO se construyo antes, al inicio del modulo.
# ---------------------------------------------------------------------------
_cargar_regiones_anonimas()
_MAPA_REGION_NORMALIZADO = {normalizar_texto(k): v for k, v in TECNICOS_REGION.items()}


if __name__ == "__main__":
    # Prueba rapida por consola:  python core.py
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    resultado = calcular_tablero()
    print()
    print(resumen_consola(resultado))
    print()
    print("--- METRICAS POR TECNICO (top 8 por vencidos) ---")
    tabla = metricas_por_tecnico(resultado.df_completo)
    columnas = ["TECNICO", "REGION", "ASIGNADOS", "VENCIDOS", "PROXIMOS_VENCER", "CUMPLIMIENTO_PCT"]
    print(tabla[columnas].head(8).to_string(index=False))
    print()
    for aviso in resultado.warnings:
        print(f"AVISO: {aviso}")
