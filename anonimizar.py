"""
anonimizar.py - Oculta los datos sensibles antes de publicar el tablero.

PARA QUE SIRVE:
    Si el tablero se publica en internet (Streamlit Cloud exige repositorio
    publico en el plan gratuito), NO se deben exponer los datos reales del
    banco: numeros de caso, nombres de tecnicos ni oficinas.

    Este modulo los reemplaza por etiquetas genericas CONSERVANDO la estructura,
    de modo que el tablero sigue siendo util para demostrar y validar el
    sistema, pero sin filtrar informacion.

MODOS (variable de entorno TORRE_ANONIMIZAR):
    "no"       -> sin cambios (uso interno, datos reales)
    "completo" (POR DEFECTO si se detecta Streamlit Cloud) -> oculta casos,
                  tecnicos y oficinas
    "parcial"  -> oculta numeros de caso y oficinas, pero CONSERVA los nombres
                  de los tecnicos (util para grupos de trabajo internos)

Ejemplo:
    QT3329957 / JHONATHAN FELIPE JARABA CHALARCA / 1300-REGIONAL ANTIOQUIA
        ->
    CASO-0007 / TECNICO-03 / OFICINA-05
"""

from __future__ import annotations

import os

import pandas as pd

MODO_NO = "no"
MODO_COMPLETO = "completo"
MODO_PARCIAL = "parcial"

MODOS = (MODO_NO, MODO_PARCIAL, MODO_COMPLETO)

# Columnas que se anonimizan en cada modo.
# IMPORTANTE: se listan los DOS nombres posibles de cada campo, porque este
# modulo puede recibir:
#   - el Excel CRUDO  -> "TECNICO COLSOF ASIGNADO INICIALMENTE", "OFICINA", ...
#   - el DataFrame YA CALCULADO -> "TECNICO", "REGION_TECNICO", ...
# Si se olvida uno, los datos reales se publican sin anonimizar.
_COL_CASO = ("N° DE CASO",)
_COL_TECNICO = ("TECNICO", "TECNICO COLSOF ASIGNADO INICIALMENTE")
_COL_TECNICO_RESOLUTOR = ("TECNICO_RESOLUTOR", "TECNICO COLSOF RESOLUTOR")
# Columnas con NOMBRES DE PERSONAS que trae la plantilla y no se usan en el
# tablero, pero que expondrian apellidos reales si se publican.
_COL_ASIGNADO_POR = ("Asignado por",)
_COL_CERRADO_POR = ("Cerrado por",)
_COL_OFICINA = ("OFICINA",)
_COL_REGIONAL = ("REGIONAL",)
_COL_DEPARTAMENTO = ("DEPARTAMENTO",)
# Ciudades/regiones legitimas del negocio: NO se anonimizan, hacen falta para
# separar Bogota de Regionales. Una ciudad no es un dato personal.
_COL_REGION_TECNICO = ("REGION_TECNICO",)


def _col(df: pd.DataFrame, nombres: tuple[str, ...]) -> str | None:
    """Devuelve el primer nombre de columna que exista en el DataFrame."""
    for n in nombres:
        if n in df.columns:
            return n
    return None


def en_la_nube() -> bool:
    """
    True si la app parece estar corriendo en un servicio web (no en el PC local).

    Se detecta por la variable SHARING (la define Streamlit Community Cloud) o
    por HOSTNAME con nombre de contenedor.
    """
    if os.environ.get("SHARING", "").lower() in ("true", "1", "yes"):
        return True
    host = os.environ.get("HOSTNAME", "").lower()
    return any(x in host for x in ("streamlit", "codespace", "heroku", "render"))


def modo_actual() -> str:
    """
    Modo de anonimizacion activo.

    Prioridad: variable de entorno TORRE_ANONIMIZAR (o el secreto del mismo
    nombre, que auth.puente_secretos copia al entorno); si no esta definida,
    los datos se muestran REALES.

    Nota de la version actual: antes, si la app corria en la nube, se anonimizaba
    automaticamente. Eso se quito a proposito, porque ahora la proteccion de los
    datos la da el LOGIN obligatorio (auth.py), no el enmascaramiento. Si en
    algun momento se publica el tablero SIN login, defina TORRE_ANONIMIZAR en
    "parcial" o "completo".
    """
    valor = (os.environ.get("TORRE_ANONIMIZAR") or "").strip().lower()
    if valor in MODOS:
        return valor
    if valor in ("si", "sí", "true", "1", "yes"):
        return MODO_COMPLETO
    if valor in ("false", "0", "none"):
        return MODO_NO
    return MODO_NO


def descripcion(modo: str | None = None) -> str:
    """Texto legible del modo activo, para mostrar en la interfaz."""
    modo = modo or modo_actual()
    return {
        MODO_NO: "datos reales (sin anonimizar)",
        MODO_PARCIAL: "datos anonimizados parcialmente (casos y oficinas ocultos)",
        MODO_COMPLETO: "datos anonimizados (casos, tecnicos y oficinas ocultos)",
    }.get(modo, modo)


def anonimizar(df: pd.DataFrame, modo: str | None = None) -> pd.DataFrame:
    """
    Devuelve una copia del DataFrame con los datos sensibles reemplazados.

    Se usa un indice estable: el mismo caso o tecnico recibe siempre la misma
    etiqueta dentro de la misma ejecucion, para no romper agrupaciones ni
    conteos (los totales y las metricas siguen siendo correctos).
    """
    modo = modo or modo_actual()
    if df is None or df.empty or modo == MODO_NO:
        return df

    trabajo = df.copy()

    # --- Numeros de caso ---------------------------------------------------
    col_caso = _col(trabajo, _COL_CASO)
    if col_caso:
        mapa_casos = {
            valor: f"CASO-{i:04d}"
            for i, valor in enumerate(
                sorted(trabajo[col_caso].dropna().astype(str).unique()), start=1
            )
        }
        trabajo[col_caso] = (
            trabajo[col_caso].astype(str).map(mapa_casos).fillna("CASO-0000")
        )

    # --- Oficina (siempre) y regional/departamento (solo en modo completo) --
    # NOTA: NO se anonimiza REGION_TECNICO ni se toca la columna del diccionario
    # de tecnicos, porque el tablero necesita separar Bogota de Regionales.
    # Una ciudad no es un dato personal; lo sensible son el caso, el tecnico y
    # la oficina exacta.
    objetivos: list[tuple[tuple[str, ...], str]] = [(_COL_OFICINA, "OFICINA")]
    if modo == MODO_COMPLETO:
        objetivos += [
            (_COL_REGIONAL, "REGIONAL"),
            (_COL_DEPARTAMENTO, "DEPARTAMENTO"),
        ]

    for nombres, etiqueta in objetivos:
        columna = _col(trabajo, nombres)
        if not columna:
            continue
        valores = trabajo[columna].dropna().astype(str)
        valores = valores[valores.str.strip() != ""]
        mapa = {
            valor: f"{etiqueta}-{i:02d}"
            for i, valor in enumerate(sorted(valores.unique()), start=1)
        }
        if mapa:
            trabajo[columna] = trabajo[columna].astype(str).map(mapa).fillna("")

    # --- Tecnicos (solo en modo completo) ---------------------------------
    # Se aplica a TODAS las columnas que contengan nombres de tecnicos, tanto en
    # el Excel crudo como en el DataFrame calculado.
    if modo == MODO_COMPLETO:
        columnas_tec = [
            c for c in (_col(trabajo, _COL_TECNICO),
                        _col(trabajo, _COL_TECNICO_RESOLUTOR),
                        _col(trabajo, _COL_ASIGNADO_POR),
                        _col(trabajo, _COL_CERRADO_POR))
            if c
        ]
        if columnas_tec:
            # Un unico mapa para que el mismo nombre reciba la misma etiqueta en
            # las dos columnas (asignado y resolutor).
            todos = pd.concat([trabajo[c].dropna().astype(str) for c in columnas_tec])
            todos = todos[todos.str.strip() != ""]
            mapa = _mapa_tecnicos(todos)
            for columna in columnas_tec:
                original = trabajo[columna].astype(str)
                trabajo[columna] = original.map(mapa).fillna(original)

    return trabajo


_MAPA_TECNICOS_GLOBAL: dict[str, str] = {}
# Region de cada etiqueta TECNICO-NN. Permite que, tras anonimizar, se siga
# pudiendo separar Bogota de Regionales sin conocer los nombres reales.
_MAPA_REGIONES_ETIQUETA: dict[str, str] = {}


def regiones_por_etiqueta() -> dict[str, str]:
    """
    Mapa {TECNICO-NN: REGION} construido durante la anonimizacion.

    Usa el diccionario TECNICOS_REGION de core para traducir el nombre real a su
    ciudad ANTES de ocultarlo. Asi el tablero conserva el corte por region.
    """
    return dict(_MAPA_REGIONES_ETIQUETA)


def _mapa_tecnicos(valores: pd.Series) -> dict[str, str]:
    """
    Mapa nombre real -> TECNICO-NN, compartido entre llamadas.

    Se guarda en el modulo para que el mismo tecnico conserve su etiqueta
    aunque se anonimicen dataframes distintos (casos abiertos y cerrados).
    Tambien se registra la region de cada etiqueta, consultando el diccionario
    de core, para no perder el corte Bogota vs Regionales.
    """
    from core import TECNICOS_REGION, obtener_region  # import local: evita ciclos

    for valor in sorted(valores.unique()):
        if valor not in _MAPA_TECNICOS_GLOBAL:
            etiqueta = f"TECNICO-{len(_MAPA_TECNICOS_GLOBAL) + 1:02d}"
            _MAPA_TECNICOS_GLOBAL[valor] = etiqueta
            # Se guarda la region real del tecnico bajo su etiqueta
            region = obtener_region(valor)
            _MAPA_REGIONES_ETIQUETA[etiqueta] = region
            TECNICOS_REGION.setdefault(etiqueta, region)
    return _MAPA_TECNICOS_GLOBAL


def reiniciar() -> None:
    """Limpia los mapas (util en pruebas)."""
    _MAPA_TECNICOS_GLOBAL.clear()
    _MAPA_REGIONES_ETIQUETA.clear()


def aviso_para_la_interfaz(modo: str | None = None) -> str | None:
    """
    Texto de advertencia para mostrar en el tablero, o None si no aplica.

    Es importante que quien vea el tablero SEPA que los datos estan alterados.
    """
    modo = modo or modo_actual()
    if modo == MODO_NO:
        return None
    if modo == MODO_PARCIAL:
        return (
            "🔒 **Datos anonimizados parcialmente.** Los números de caso y las "
            "oficinas fueron reemplazados por etiquetas. Los totales y las "
            "métricas son correctos."
        )
    return (
        "🔒 **Datos anonimizados.** Los números de caso, los técnicos y las "
        "oficinas fueron reemplazados por etiquetas genéricas para no exponer "
        "información del banco. Los totales, estados y métricas son correctos; "
        "solo cambian las etiquetas."
    )


if __name__ == "__main__":
    # Demostracion:  python anonimizar.py
    import warnings

    warnings.filterwarnings("ignore")
    from core import calcular_tablero

    print("=" * 74)
    print(" DEMOSTRACION DE ANONIMIZACION")
    print("=" * 74)
    print(f" Modo detectado: {modo_actual()}  ({descripcion()})")
    print()

    resultado = calcular_tablero()
    original = resultado.requieren_atencion.head(3)
    columnas = [c for c in ["N° DE CASO", "TECNICO", "REGION_TECNICO", "OFICINA"] if c in original.columns]

    print("--- ANTES (datos reales) ---")
    print(original[columnas].to_string(index=False))

    for modo in (MODO_PARCIAL, MODO_COMPLETO):
        reiniciar()
        oculto = anonimizar(original, modo)
        print()
        print(f"--- DESPUES (modo {modo}) ---")
        print(oculto[columnas].to_string(index=False))
