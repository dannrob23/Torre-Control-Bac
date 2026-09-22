"""
avisos_automaticos.py - Envio AUTOMATICO de avisos de vencimiento a los tecnicos.

Pensado para el Programador de Tareas de Windows: cada X minutos detecta que
tecnicos tienen casos pendientes y envia el aviso por Telegram, sin intervencion
de nadie. Deja todo registrado en historial.db.

Diferencia con alertas_windows.py:
  - alertas_windows.py        -> avisa a LA TORRE (notificaciones de escritorio +
                                 tabla resumen al grupo). Es supervision.
  - avisos_automaticos.py     -> redacta y envia el aviso A CADA TECNICO. Es gestion.

Uso:
    python avisos_automaticos.py                  # envia a todos los tecnicos pendientes
    python avisos_automaticos.py --dry-run         # muestra que enviaria, sin enviar
    python avisos_automaticos.py --resumen         # solo la tabla de pendientes
    python avisos_automaticos.py --tecnico "JUAN PEREZ"
    python avisos_automaticos.py --solo-rojo       # solo casos vencidos
    python avisos_automaticos.py --max 5           # limita a 5 tecnicos por corrida
    python avisos_automaticos.py --horas-anti-duplicado 24

IMPORTANTE (limitacion de Telegram):
    Un bot NO puede iniciar un chat privado con una persona. Para que un tecnico
    reciba el aviso en privado, esa persona debe abrir el chat del bot y pulsar
    /start UNA vez. Por eso el envio automatico escribe al GRUPO con la mencion
    @usuario o +telefono: quien tenga ese usuario o numero en el grupo recibe la
    notificacion. Si prefiere privados, configure "chat_id" por tecnico en
    menciones.json (campo "chat_id") y use --privado.
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import time

import pandas as pd

from avisos import (
    agrupar_por_tecnico,
    cargar_menciones,
    componer_aviso_tecnico,
    resumen_pendientes,
)
from core import (
    MODO_ACUMULADO,
    MODO_DIAS,
    MODO_VENTANA_POR_DEFECTO,
    ErrorLecturaExcel,
    ahora_colombia,
    calcular_tablero,
)
from historial import CANAL_TELEGRAM, RESULTADO_FALLIDO, RESULTADO_GENERADO, Historial

APP_NAME = "Torre de Control SLA - Colsof"
DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_LOG = os.path.join(DIR_SCRIPT, "avisos_automaticos.log")
ARCHIVO_LOCK = os.path.join(os.environ.get("TEMP", DIR_SCRIPT), "torre_avisos.lock")

# Anti-duplicado: no reenviar el mismo aviso al mismo tecnico dentro de N horas.
HORAS_ANTI_DUPLICADO = 12

log = logging.getLogger("avisos_auto")


def _sin_emoji(texto: str) -> str:
    """Version ASCII, para consolas cp1252 del Programador de Tareas."""
    return str(texto).encode("ascii", "ignore").decode("ascii").replace("  ", " ").strip()


def configurar_logging(verbose: bool = False) -> None:
    log.setLevel(logging.DEBUG if verbose else logging.INFO)
    formato = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    log.propagate = False
    if not log.handlers:
        try:
            fh = logging.FileHandler(ARCHIVO_LOG, encoding="utf-8")
            fh.setFormatter(formato)
            log.addHandler(fh)
        except OSError:
            pass
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass
    if not any(isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler)
               for h in log.handlers):
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(formato)
        log.addHandler(ch)


class InstanciaUnica:
    """Evita dos ejecuciones simultaneas (disparos solapados del Programador)."""

    def __init__(self, ruta: str = ARCHIVO_LOCK, max_edad_min: int = 30) -> None:
        self.ruta = ruta
        self.max_edad_min = max_edad_min

    def __enter__(self) -> bool:
        try:
            if os.path.exists(self.ruta) and self._expirado():
                os.remove(self.ruta)
            fd = os.open(self.ruta, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w") as fh:
                fh.write(str(os.getpid()))
            return True
        except FileExistsError:
            return False
        except OSError:
            return True

    def _expirado(self) -> bool:
        try:
            return (time.time() - os.path.getmtime(self.ruta)) / 60 > self.max_edad_min
        except OSError:
            return True

    def __exit__(self, *_) -> None:
        try:
            if os.path.exists(self.ruta):
                os.remove(self.ruta)
        except OSError:
            pass


def seleccionar_tecnicos(pendientes: pd.DataFrame, historial: Historial | None,
                         horas_anti_duplicado: int, tecnico: str | None = None,
                         solo_rojo: bool = False, sin_anti_duplicado: bool = False
                         ) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame]]:
    """
    Decide que tecnicos se avisan en esta corrida.

    Returns:
        (por_enviar, omitidos): dos dicts tecnico -> DataFrame de casos.
    """
    from core import normalizar_texto

    grupos = agrupar_por_tecnico(pendientes, solo_con_tecnico=True)

    if tecnico:
        objetivo = {t: g for t, g in grupos.items()
                    if normalizar_texto(tecnico) in normalizar_texto(t)}
        grupos = objetivo

    if solo_rojo:
        grupos = {
            t: g[g["ESTADO"].isin(["ROJO", "CERRADO TARDE"])]
            for t, g in grupos.items()
        }
        grupos = {t: g for t, g in grupos.items() if not g.empty}

    if sin_anti_duplicado or historial is None:
        return grupos, {}

    # Anti-duplicado por TECNICO: si ya se le envio un aviso hace poco, se omite.
    recientes = historial.ultima_notificacion_por_tecnico()
    limite = ahora_colombia() - pd.Timedelta(hours=horas_anti_duplicado)

    por_enviar: dict[str, pd.DataFrame] = {}
    omitidos: dict[str, pd.DataFrame] = {}
    for nombre, grupo in grupos.items():
        ultima = recientes.get(nombre)
        if ultima:
            try:
                if pd.Timestamp(ultima) > pd.Timestamp(limite):
                    omitidos[nombre] = grupo
                    continue
            except (ValueError, TypeError):
                pass
        por_enviar[nombre] = grupo
    return por_enviar, omitidos


def enviar_aviso(tecnico: str, df_tecnico: pd.DataFrame, *, dry_run: bool,
                 un_solo_mensaje: bool, menciones: dict, historial: Historial | None,
                 canal_telegram: bool, destino_chat_id: str | None = None) -> bool:
    """
    Compone y envia el aviso de UN tecnico. Registra el resultado.

    Returns: True si se envio (o si dry_run).
    """
    texto = componer_aviso_tecnico(
        tecnico, df_tecnico, canal="telegram", un_solo_mensaje=un_solo_mensaje,
        menciones=menciones,
    )
    if not texto.strip():
        return False

    if dry_run:
        log.info("[DRY-RUN] %s (%d caso/s)", tecnico, len(df_tecnico))
        # Se quitan los emojis (consola cp1252) y el espacio que dejan al irse.
        for linea in _sin_emoji(texto).split("\n")[:6]:
            limpia = " ".join(linea.split())
            if limpia:
                log.info("          %s", limpia)
        return True

    enviado = False
    if canal_telegram:
        try:
            import telegram_notifier as tg

            if tg.telegram_configurado():
                if destino_chat_id:
                    ok = tg.enviar_mensaje(texto, token=None,
                                           chat_id=destino_chat_id, parse_mode="")
                    enviado = bool(ok)
                else:
                    enviado = tg.enviar_a_todos(texto, parse_mode="") > 0
            else:
                log.warning("Telegram no esta configurado; no se envia el aviso de %s.", tecnico)
        except Exception as exc:
            log.error("Fallo el envio a %s: %s: %s", tecnico, type(exc).__name__, exc)

    # Registro en el historial
    if historial is not None:
        resultado = "enviado" if enviado else (RESULTADO_FALLIDO if canal_telegram
                                              else RESULTADO_GENERADO)
        detalle = ("aviso automatico por Telegram" if enviado
                   else "generado; no se pudo enviar" if canal_telegram
                   else "generado para envio manual")
        for _, fila in df_tecnico.iterrows():
            historial.registrar(
                caso=str(fila.get("N° DE CASO", "S/N")),
                tecnico=tecnico,
                region=str(fila.get("REGION_TECNICO", "") or ""),
                estado=str(fila.get("ESTADO", "") or ""),
                horas_restantes=fila.get("HORAS_RESTANTES"),
                horas_vencido=fila.get("HORAS_VENCIDO"),
                canal=CANAL_TELEGRAM if canal_telegram else "whatsapp",
                resultado=resultado,
                detalle=detalle,
            )

    if enviado:
        log.info("Aviso ENVIADO a %s (%d caso/s)", tecnico, len(df_tecnico))
    else:
        log.warning("Aviso NO enviado a %s (%d caso/s) - queda registrado como fallido",
                    tecnico, len(df_tecnico))
    return enviado


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Envio automatico de avisos de vencimiento a los tecnicos."
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Muestra lo que enviaria, sin enviar ni registrar.")
    parser.add_argument("--resumen", action="store_true",
                        help="Solo muestra la tabla de pendientes por tecnico.")
    parser.add_argument("--tecnico", type=str, default=None,
                        help="Limita el envio a un tecnico (coincidencia parcial).")
    parser.add_argument("--solo-rojo", action="store_true",
                        help="Solo avisa de casos vencidos (ROJO / CERRADO TARDE).")
    parser.add_argument("--max", type=int, default=0,
                        help="Maximo de tecnicos a avisar en esta corrida (0 = sin limite).")
    parser.add_argument("--multiple", dest="multiple", action="store_true", default=True,
                        help="Un solo mensaje con la lista de casos (por defecto).")
    parser.add_argument("--un-mensaje-por-caso", dest="multiple", action="store_false",
                        help="Un mensaje por cada caso, en vez de una lista.")
    parser.add_argument("--sin-telegram", action="store_true",
                        help="No envia a Telegram; solo genera y registra.")
    parser.add_argument("--privado", action="store_true",
                        help="Envia en privado al chat_id de cada tecnico (menciones.json).")
    parser.add_argument("--horas-anti-duplicado", type=int, default=HORAS_ANTI_DUPLICADO,
                        help=f"No reavisar al mismo tecnico dentro de estas horas "
                             f"(por defecto {HORAS_ANTI_DUPLICADO}).")
    parser.add_argument("--sin-anti-duplicado", action="store_true",
                        help="Avisar a todos siempre, sin filtrar repetidos.")
    parser.add_argument("--dias-ventana", type=int, default=3)
    parser.add_argument("--modo-ventana", choices=[MODO_DIAS, MODO_ACUMULADO],
                        default=MODO_VENTANA_POR_DEFECTO)
    parser.add_argument("--excel", type=str, default=None)
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args(argv)

    configurar_logging(args.verbose)

    try:
        resultado = calcular_tablero(
            args.excel, dias_ventana=args.dias_ventana, modo_ventana=args.modo_ventana
        )
    except ErrorLecturaExcel as exc:
        log.error("No se pudo leer la plantilla: %s", exc)
        return 2

    pendientes = resultado.df[resultado.df["ESTADO"].isin(
        ["ROJO", "CERRADO TARDE", "NARANJA", "AMARILLO"]
    )]

    log.info("Ventana: %d caso(s) | pendientes de aviso: %d",
             resultado.total_en_ventana, len(pendientes))

    if pendientes.empty:
        log.info("No hay casos pendientes: no se envia ningun aviso.")
        return 0

    tabla = resumen_pendientes(pendientes)

    if args.resumen:
        print()
        print(tabla.to_string(index=False))
        return 0

    historial = None if args.dry_run else Historial()
    menciones = cargar_menciones()
    if not menciones:
        log.warning(
            "No existe menciones.json: los avisos saldran SIN @usuario ni telefono, "
            "por lo que los tecnicos no recibiran notificacion real."
        )

    por_enviar, omitidos = seleccionar_tecnicos(
        pendientes, historial,
        horas_anti_duplicado=args.horas_anti_duplicado,
        tecnico=args.tecnico, solo_rojo=args.solo_rojo,
        sin_anti_duplicado=args.sin_anti_duplicado,
    )

    log.info("Tecnicos a avisar: %d | omitidos por anti-duplicado: %d",
             len(por_enviar), len(omitidos))

    if not por_enviar:
        log.info("Nada que enviar en esta corrida.")
        return 0

    if args.max > 0:
        por_enviar = dict(list(por_enviar.items())[: args.max])

    enviados = 0
    for tecnico, df_tecnico in por_enviar.items():
        chat_id = None
        if args.privado:
            from core import normalizar_texto

            info = menciones.get(normalizar_texto(tecnico), {})
            chat_id = info.get("chat_id") or None
            if not chat_id:
                log.warning(
                    "%s: --privado pero no tiene 'chat_id' en menciones.json; "
                    "se omite (no se puede iniciar un chat privado sin su id).", tecnico,
                )
                continue

        if enviar_aviso(
            tecnico, df_tecnico,
            dry_run=args.dry_run,
            un_solo_mensaje=args.multiple,
            menciones=menciones,
            historial=historial,
            canal_telegram=not args.sin_telegram,
            destino_chat_id=chat_id,
        ):
            enviados += 1
        time.sleep(1.0)  # respeta el limite de la API de Telegram

    log.info("Avisos procesados: %d/%d", enviados, len(por_enviar))
    return 0


if __name__ == "__main__":
    with InstanciaUnica() as puede:
        if not puede:
            print("Ya hay una ejecucion de avisos en curso. Se omite esta corrida.")
            sys.exit(0)
        sys.exit(main())
