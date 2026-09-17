"""
historial.py - Persistencia en SQLite del historial de notificaciones.

Reemplaza la utilidad de `alertas.log` (texto plano) para lo que necesita
consultas: cuantas veces se notifico a cada tecnico, cuando, con que estado y
por que canal.

Tabla principal: `notificaciones`
    id, fecha_hora, caso, tecnico, region, estado, horas_restantes,
    horas_vencido, canal, resultado, detalle

Uso:
    from historial import Historial
    with Historial() as h:
        h.registrar(caso="IM123", tecnico="...", estado="ROJO", ...)
        df = h.por_tecnico()
"""

from __future__ import annotations

import logging
import os
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timedelta

import pandas as pd

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_DB = os.path.join(DIR_SCRIPT, "historial.db")

log = logging.getLogger("historial")

CANAL_WINDOWS = "windows"
CANAL_TELEGRAM = "telegram"
CANAL_WHATSAPP = "whatsapp"
CANAL_AMBOS = "windows+telegram"
RESULTADO_ENVIADO = "enviado"
RESULTADO_FALLIDO = "fallido"
# El aviso se genero y la torre lo copio para pegarlo a mano (WhatsApp).
# NO es un envio automatico: se registra aparte para no ensuciar el log.
RESULTADO_GENERADO = "generado"

# No repetir notificacion del mismo caso dentro de esta ventana (minutos).
MINUTOS_ANTI_DUPLICADO = 60

_ESQUEMA = """
CREATE TABLE IF NOT EXISTS notificaciones (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha_hora       TEXT    NOT NULL,
    caso             TEXT    NOT NULL,
    tecnico          TEXT,
    region           TEXT,
    estado           TEXT,
    horas_restantes  REAL,
    horas_vencido    REAL,
    canal            TEXT,
    resultado        TEXT,
    detalle          TEXT
);
CREATE INDEX IF NOT EXISTS idx_notif_caso    ON notificaciones (caso);
CREATE INDEX IF NOT EXISTS idx_notif_fecha   ON notificaciones (fecha_hora);
CREATE INDEX IF NOT EXISTS idx_notif_tecnico ON notificaciones (tecnico);
CREATE INDEX IF NOT EXISTS idx_notif_estado  ON notificaciones (estado);
"""


class Historial:
    """Acceso al historial de notificaciones (SQLite)."""

    def __init__(self, ruta: str | None = None) -> None:
        self.ruta = ruta or ARCHIVO_DB
        self._crear_esquema()

    # ------------------------------------------------------------------
    # Conexion
    # ------------------------------------------------------------------

    @contextmanager
    def _conexion(self):
        """Conexion con commit/rollback automatico."""
        conexion = sqlite3.connect(self.ruta, timeout=10)
        conexion.row_factory = sqlite3.Row
        try:
            yield conexion
            conexion.commit()
        except Exception:
            conexion.rollback()
            raise
        finally:
            conexion.close()

    def _crear_esquema(self) -> None:
        try:
            with self._conexion() as cx:
                cx.executescript(_ESQUEMA)
        except sqlite3.Error as exc:
            # No es fatal: el sistema puede seguir notificando sin historial.
            log.error("No se pudo preparar el historial (%s): %s", self.ruta, exc)

    # ------------------------------------------------------------------
    # Escritura
    # ------------------------------------------------------------------

    def registrar(
        self,
        *,
        caso: str,
        tecnico: str = "",
        region: str = "",
        estado: str = "",
        horas_restantes: float | None = None,
        horas_vencido: float | None = None,
        canal: str = CANAL_WINDOWS,
        resultado: str = RESULTADO_ENVIADO,
        detalle: str = "",
        momento: datetime | None = None,
    ) -> None:
        """Inserta una notificacion en el historial."""
        momento = momento or datetime.now()
        fila = (
            momento.strftime("%Y-%m-%d %H:%M:%S"),
            str(caso),
            str(tecnico or ""),
            str(region or ""),
            str(estado or ""),
            _a_float(horas_restantes),
            _a_float(horas_vencido),
            str(canal),
            str(resultado),
            str(detalle or ""),
        )
        try:
            with self._conexion() as cx:
                cx.execute(
                    "INSERT INTO notificaciones "
                    "(fecha_hora, caso, tecnico, region, estado, horas_restantes, "
                    " horas_vencido, canal, resultado, detalle) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    fila,
                )
        except sqlite3.Error as exc:
            log.error("No se pudo registrar la notificacion del caso %s: %s", caso, exc)

    def registrar_df(self, df: pd.DataFrame, canal: str, resultado: str = RESULTADO_ENVIADO,
                     momento: datetime | None = None) -> int:
        """Registra en lote todas las filas de un DataFrame. Devuelve cuantas inserto."""
        if df is None or df.empty:
            return 0
        momento = momento or datetime.now()
        insertadas = 0
        for _, fila in df.iterrows():
            self.registrar(
                caso=str(fila.get("N° DE CASO", "S/N")),
                tecnico=str(fila.get("TECNICO", "")),
                region=str(fila.get("REGION_TECNICO", "")),
                estado=str(fila.get("ESTADO", "")),
                horas_restantes=fila.get("HORAS_RESTANTES"),
                horas_vencido=fila.get("HORAS_VENCIDO"),
                canal=canal,
                resultado=resultado,
                detalle=str(fila.get("PREDICCION", "")),
                momento=momento,
            )
            insertadas += 1
        return insertadas

    # ------------------------------------------------------------------
    # Anti-duplicado
    # ------------------------------------------------------------------

    def casos_notificados_recientemente(
        self, minutos: int = MINUTOS_ANTI_DUPLICADO, momento: datetime | None = None
    ) -> set[str]:
        """
        Casos notificados con exito en los ultimos N minutos.

        Permite que una ejecucion cada 15 min no repita 226 notificaciones
        cada vez: solo avisa de lo nuevo o de lo que cambio de estado.
        """
        momento = momento or datetime.now()
        desde = (momento - timedelta(minutes=minutos)).strftime("%Y-%m-%d %H:%M:%S")
        try:
            with self._conexion() as cx:
                filas = cx.execute(
                    "SELECT DISTINCT caso FROM notificaciones "
                    "WHERE fecha_hora >= ? AND resultado = ?",
                    (desde, RESULTADO_ENVIADO),
                ).fetchall()
            return {str(f["caso"]) for f in filas}
        except sqlite3.Error as exc:
            log.error("No se pudo consultar el anti-duplicado: %s", exc)
            return set()

    def estados_notificados_recientemente(
        self, minutos: int = MINUTOS_ANTI_DUPLICADO, momento: datetime | None = None
    ) -> dict[str, str]:
        """
        Mapa caso -> ultimo estado notificado en la ventana reciente.

        Con esto se evita repetir el aviso si el color no cambio, pero SI se
        vuelve a notificar cuando un caso empeora (VERDE -> AMARILLO -> ...).
        """
        momento = momento or datetime.now()
        desde = (momento - timedelta(minutes=minutos)).strftime("%Y-%m-%d %H:%M:%S")
        try:
            with self._conexion() as cx:
                filas = cx.execute(
                    "SELECT caso, estado, MAX(fecha_hora) AS ultima "
                    "FROM notificaciones WHERE fecha_hora >= ? AND resultado = ? "
                    "GROUP BY caso",
                    (desde, RESULTADO_ENVIADO),
                ).fetchall()
            return {str(f["caso"]): str(f["estado"]) for f in filas}
        except sqlite3.Error as exc:
            log.error("No se pudo consultar estados recientes: %s", exc)
            return {}

    # ------------------------------------------------------------------
    # Lectura / metricas
    # ------------------------------------------------------------------

    def leer(self, desde: str | None = None, hasta: str | None = None) -> pd.DataFrame:
        """Devuelve el historial completo (o por rango de fechas) como DataFrame."""
        consulta = "SELECT * FROM notificaciones"
        parametros: list[str] = []
        condiciones = []
        if desde:
            condiciones.append("fecha_hora >= ?")
            parametros.append(desde)
        if hasta:
            condiciones.append("fecha_hora <= ?")
            parametros.append(hasta)
        if condiciones:
            consulta += " WHERE " + " AND ".join(condiciones)
        consulta += " ORDER BY fecha_hora DESC"

        try:
            with self._conexion() as cx:
                return pd.read_sql_query(consulta, cx, params=parametros)
        except (sqlite3.Error, pd.errors.DatabaseError) as exc:
            log.error("No se pudo leer el historial: %s", exc)
            return pd.DataFrame(
                columns=[
                    "id", "fecha_hora", "caso", "tecnico", "region", "estado",
                    "horas_restantes", "horas_vencido", "canal", "resultado", "detalle",
                ]
            )

    def por_tecnico(self, desde: str | None = None) -> pd.DataFrame:
        """
        Metricas de notificaciones agrupadas por tecnico.

        Columnas: TECNICO, REGION, NOTIFICACIONES, CASOS_DISTINTOS,
                  ULTIMA_NOTIFICACION, ROJOS, NARANJAS, AMARILLOS, FALLIDAS
        """
        df = self.leer(desde=desde)
        if df.empty:
            return pd.DataFrame(
                columns=[
                    "TECNICO", "REGION", "NOTIFICACIONES", "CASOS_DISTINTOS",
                    "ULTIMA_NOTIFICACION", "ROJOS", "NARANJAS", "AMARILLOS", "FALLIDAS",
                ]
            )

        # OJO: las columnas de SQLite vienen en minuscula (tecnico, region, ...).
        agrupado = df.groupby("tecnico", dropna=False)
        salida = pd.DataFrame(
            {
                "REGION": agrupado["region"].agg(
                    lambda s: s.dropna().mode().iloc[0] if not s.dropna().mode().empty else ""
                ),
                "NOTIFICACIONES": agrupado.size(),
                "CASOS_DISTINTOS": agrupado["caso"].nunique(),
                "ULTIMA_NOTIFICACION": agrupado["fecha_hora"].max(),
                "ROJOS": agrupado["estado"].apply(lambda s: int((s == "ROJO").sum())),
                "NARANJAS": agrupado["estado"].apply(lambda s: int((s == "NARANJA").sum())),
                "AMARILLOS": agrupado["estado"].apply(lambda s: int((s == "AMARILLO").sum())),
                "FALLIDAS": agrupado["resultado"].apply(
                    lambda s: int((s == RESULTADO_FALLIDO).sum())
                ),
            }
        )
        salida = salida.reset_index().rename(columns={"tecnico": "TECNICO"})
        salida["TECNICO"] = salida["TECNICO"].replace("", "(sin tecnico)")
        return salida.sort_values("NOTIFICACIONES", ascending=False).reset_index(drop=True)

    def resumen(self) -> dict[str, int]:
        """Totales generales del historial."""
        vacio = {
            "total": 0, "enviadas": 0, "fallidas": 0, "generadas": 0,
            "casos": 0, "tecnicos": 0,
        }
        try:
            with self._conexion() as cx:
                f = cx.execute(
                    "SELECT COUNT(*) AS total, "
                    "SUM(CASE WHEN resultado = ? THEN 1 ELSE 0 END) AS enviadas, "
                    "SUM(CASE WHEN resultado = ? THEN 1 ELSE 0 END) AS fallidas, "
                    "SUM(CASE WHEN resultado = ? THEN 1 ELSE 0 END) AS generadas, "
                    "COUNT(DISTINCT caso) AS casos, "
                    "COUNT(DISTINCT tecnico) AS tecnicos "
                    "FROM notificaciones",
                    (RESULTADO_ENVIADO, RESULTADO_FALLIDO, RESULTADO_GENERADO),
                ).fetchone()
            return {
                "total": int(f["total"] or 0),
                "enviadas": int(f["enviadas"] or 0),
                "fallidas": int(f["fallidas"] or 0),
                "generadas": int(f["generadas"] or 0),
                "casos": int(f["casos"] or 0),
                "tecnicos": int(f["tecnicos"] or 0),
            }
        except sqlite3.Error as exc:
            log.error("No se pudo calcular el resumen: %s", exc)
            return vacio

    def ultima_notificacion_por_tecnico(self, solo_exitosas: bool = True) -> dict[str, str]:
        """
        Mapa tecnico -> fecha/hora de la ULTIMA notificacion.

        Por defecto se ignoran los intentos FALLIDOS: si un envio fallo (por
        ejemplo el chat_id estaba mal), el tecnico NO quedo avisado y por tanto
        no debe contarse como notificado ni bloquear un reintento.
        Use solo_exitosas=False para ver el ultimo intento sin importar el
        resultado (util para diagnostico).
        """
        try:
            with self._conexion() as cx:
                if solo_exitosas:
                    filas = cx.execute(
                        "SELECT tecnico, MAX(fecha_hora) AS ultima FROM notificaciones "
                        "WHERE resultado IN (?, ?) GROUP BY tecnico",
                        (RESULTADO_ENVIADO, RESULTADO_GENERADO),
                    ).fetchall()
                else:
                    filas = cx.execute(
                        "SELECT tecnico, MAX(fecha_hora) AS ultima "
                        "FROM notificaciones GROUP BY tecnico"
                    ).fetchall()
            return {
                str(f["tecnico"] or "(sin tecnico)"): str(f["ultima"]) for f in filas
            }
        except sqlite3.Error as exc:
            log.error("No se pudo consultar la ultima notificacion: %s", exc)
            return {}

    def notificaciones_de_hoy_por_tecnico(self, momento: datetime | None = None,
                                          solo_exitosas: bool = True) -> dict[str, int]:
        """
        Mapa tecnico -> cuantas veces se le notifico HOY (dia calendario).

        Por defecto no cuenta los intentos fallidos: "veces notificado" debe
        reflejar avisos que realmente le llegaron.
        """
        momento = momento or datetime.now()
        inicio = momento.strftime("%Y-%m-%d 00:00:00")
        fin = momento.strftime("%Y-%m-%d 23:59:59")
        try:
            with self._conexion() as cx:
                if solo_exitosas:
                    filas = cx.execute(
                        "SELECT tecnico, COUNT(*) AS veces FROM notificaciones "
                        "WHERE fecha_hora BETWEEN ? AND ? AND resultado IN (?, ?) "
                        "GROUP BY tecnico",
                        (inicio, fin, RESULTADO_ENVIADO, RESULTADO_GENERADO),
                    ).fetchall()
                else:
                    filas = cx.execute(
                        "SELECT tecnico, COUNT(*) AS veces FROM notificaciones "
                        "WHERE fecha_hora BETWEEN ? AND ? GROUP BY tecnico",
                        (inicio, fin),
                    ).fetchall()
            return {
                str(f["tecnico"] or "(sin tecnico)"): int(f["veces"]) for f in filas
            }
        except sqlite3.Error as exc:
            log.error("No se pudo contar las notificaciones de hoy: %s", exc)
            return {}

    def conteo_por_resultado(self) -> dict[str, int]:
        """
        Cuantos registros hay de cada tipo de resultado.

        Permite que la torre distinga de un vistazo:
          - enviado  : el sistema lo envio de verdad (Telegram, escritorio)
          - generado : se redacto y una persona lo pego a mano (WhatsApp)
          - fallido  : se intento enviar y NO llego
        """
        base = {
            RESULTADO_ENVIADO: 0, RESULTADO_GENERADO: 0, RESULTADO_FALLIDO: 0,
        }
        try:
            with self._conexion() as cx:
                filas = cx.execute(
                    "SELECT resultado, COUNT(*) AS n FROM notificaciones "
                    "GROUP BY resultado"
                ).fetchall()
            for f in filas:
                base[str(f["resultado"] or "desconocido")] = int(f["n"])
            return base
        except sqlite3.Error as exc:
            log.error("No se pudo contar por resultado: %s", exc)
            return base

    def por_caso(self) -> pd.DataFrame:
        """
        Mapa caso -> (veces notificado, primera, ultima, ultimo estado).

        Sirve para responder "¿cuantas veces he avisado de este caso?".
        """
        columnas = ["caso", "VECES", "PRIMERA", "ULTIMA", "ULTIMO_ESTADO", "ULTIMO_TECNICO"]
        try:
            with self._conexion() as cx:
                filas = cx.execute(
                    "SELECT caso, COUNT(*) AS veces, MIN(fecha_hora) AS primera, "
                    "MAX(fecha_hora) AS ultima FROM notificaciones GROUP BY caso"
                ).fetchall()
                detalle = {}
                for f in filas:
                    ult = cx.execute(
                        "SELECT estado, tecnico FROM notificaciones WHERE caso = ? "
                        "ORDER BY fecha_hora DESC LIMIT 1",
                        (f["caso"],),
                    ).fetchone()
                    detalle[str(f["caso"])] = (
                        int(f["veces"]), str(f["primera"]), str(f["ultima"]),
                        str(ult["estado"]) if ult else "",
                        str(ult["tecnico"]) if ult else "",
                    )
            return pd.DataFrame(
                [{"caso": k, "VECES": v[0], "PRIMERA": v[1], "ULTIMA": v[2],
                  "ULTIMO_ESTADO": v[3], "ULTIMO_TECNICO": v[4]}
                 for k, v in detalle.items()],
                columns=columnas,
            )
        except sqlite3.Error as exc:
            log.error("No se pudo agrupar por caso: %s", exc)
            return pd.DataFrame(columns=columnas)

    def purgar(self, dias: int = 180) -> int:
        """Elimina registros mas antiguos que N dias. Devuelve cuantas filas borro."""
        limite = (datetime.now() - timedelta(days=dias)).strftime("%Y-%m-%d %H:%M:%S")
        try:
            with self._conexion() as cx:
                cur = cx.execute("DELETE FROM notificaciones WHERE fecha_hora < ?", (limite,))
                return int(cur.rowcount or 0)
        except sqlite3.Error as exc:
            log.error("No se pudo purgar el historial: %s", exc)
            return 0


def _a_float(valor) -> float | None:
    """Convierte a float tolerando None/NaN."""
    try:
        if valor is None or pd.isna(valor):
            return None
        return float(valor)
    except (TypeError, ValueError):
        return None


if __name__ == "__main__":
    # Diagnostico rapido:  python historial.py
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    h = Historial()
    print(f"Base de datos: {h.ruta}")
    print(f"Resumen: {h.resumen()}")
    tabla = h.por_tecnico()
    if tabla.empty:
        print("\nAun no hay notificaciones registradas.")
    else:
        print(f"\nNotificaciones por tecnico ({len(tabla)} tecnicos):")
        print(tabla.head(15).to_string(index=False))
