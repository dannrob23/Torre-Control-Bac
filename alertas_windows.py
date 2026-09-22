"""
alertas_windows.py - Notificaciones de la Torre de Control SLA (version 2).

Notifica los casos que requieren atencion usando DOS canales:
  1. Notificaciones nativas de escritorio de Windows (plyer, con respaldo PowerShell)
  2. Telegram: un mensaje resumen con la tabla de estados (opcional, --telegram)

Logica v2:
  - Solo casos DENTRO de la ventana de notificacion: los que vencen dentro de los
    proximos N dias calendario (por defecto 3), MAS todos los vencidos sin cerrar.
  - Estados notificados: ROJO, CERRADO TARDE, NARANJA y AMARILLO.
  - Anti-duplicado con historial SQLite: no repite el mismo caso con el mismo
    estado dentro de la ventana configurada, pero SI avisa si el caso empeora.

Uso:
    python alertas_windows.py                     # notifica y registra
    python alertas_windows.py --telegram           # ademas envia la tabla al grupo
    python alertas_windows.py --solo-rojo          # solo ROJO y CERRADO TARDE
    python alertas_windows.py --dry-run            # simula, no envia ni registra
    python alertas_windows.py --resumen            # solo resumen por consola
    python alertas_windows.py --test               # notificacion de prueba
    python alertas_windows.py --sin-anti-duplicado # notifica todo siempre
    python alertas_windows.py --max 10             # limita avisos de escritorio
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import tempfile
import time
from datetime import datetime

from core import (
    AMARILLO,
    CERRADO_TARDE,
    COL_CASO,
    ESTADOS_ALERTA,
    ICONO_ESTADO,
    MODO_ACUMULADO,
    MODO_DIAS,
    MODO_VENTANA_POR_DEFECTO,
    NARANJA,
    ROJO,
    SIN_VENCIMIENTO,
    VERDE,
    ErrorLecturaExcel,
    ahora_colombia,
    calcular_tablero,
    construir_mensaje_notificacion,
    resumen_consola,
)
from historial import CANAL_AMBOS, CANAL_TELEGRAM, CANAL_WINDOWS, Historial

APP_NAME = "Torre de Control SLA - Colsof"
DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_LOG = os.path.join(DIR_SCRIPT, "alertas.log")
ARCHIVO_LOCK = os.path.join(tempfile.gettempdir(), "torre_control_colsof.lock")

DURACION_NOTIFICACION = 15

# Casos por bloque al enviar la tabla a Telegram. Telegram rechaza mensajes de
# mas de 4096 caracteres, y la tabla ocupa ~110 caracteres por fila mas cabecera.
FILAS_TELEGRAM = 20

# Anti-duplicado por defecto: no repetir el mismo caso con el mismo estado
# dentro de estas horas. Con la tarea cada 15 min evita avalanchas de avisos.
HORAS_ANTI_DUPLICADO = 12

log = logging.getLogger("alertas")


def _sin_emoji(texto: str) -> str:
    """
    Version ASCII de un texto, para consolas y logs que no soportan UTF-8.

    Evita UnicodeEncodeError en consolas cp1252 (frecuente bajo el Programador
    de Tareas) y evita mojibake al abrir alertas.log con editores ANSI.
    Los emojis si se conservan en la notificacion real, que si los soporta.
    """
    return str(texto).encode("ascii", "ignore").decode("ascii").replace("  ", " ").strip()


# ---------------------------------------------------------------------------
# Logging y control de instancia unica
# ---------------------------------------------------------------------------

def configurar_logging(verbose: bool = False) -> None:
    """Log a archivo (UTF-8) y a consola de forma tolerante a fallos de encoding."""
    log.setLevel(logging.DEBUG if verbose else logging.INFO)
    formato = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")
    log.propagate = False

    if not log.handlers:
        try:
            manejador_archivo = logging.FileHandler(ARCHIVO_LOG, encoding="utf-8")
            manejador_archivo.setFormatter(formato)
            log.addHandler(manejador_archivo)
        except OSError:
            pass  # si no se puede escribir el log, seguimos solo con consola

    # La consola se fuerza a UTF-8: bajo el Programador de Tareas stdout suele ser
    # cp1252 y un emoji provocaria UnicodeEncodeError y la caida del script.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError, OSError):
        pass

    if not any(isinstance(h, logging.StreamHandler) and not isinstance(h, logging.FileHandler)
               for h in log.handlers):
        consola = logging.StreamHandler(sys.stdout)
        consola.setFormatter(formato)
        log.addHandler(consola)


class InstanciaUnica:
    """
    Evita que dos ejecuciones simultaneas dupliquen notificaciones
    (por ejemplo si el Programador de Tareas solapa dos disparos).
    Usa creacion atomica de archivo (O_CREAT|O_EXCL), valida en Windows.
    """

    def __init__(self, ruta: str = ARCHIVO_LOCK, max_edad_min: int = 30) -> None:
        self.ruta = ruta
        self.max_edad_min = max_edad_min

    def __enter__(self) -> bool:
        try:
            if os.path.exists(self.ruta) and self._expirado():
                os.remove(self.ruta)
            descriptor = os.open(self.ruta, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(descriptor, "w") as fh:
                fh.write(f"{os.getpid()}")
            return True
        except FileExistsError:
            return False
        except OSError:
            return True  # ante la duda, no bloquear la ejecucion

    def _expirado(self) -> bool:
        try:
            edad_min = (time.time() - os.path.getmtime(self.ruta)) / 60
            return edad_min > self.max_edad_min
        except OSError:
            return True

    def __exit__(self, *_) -> None:
        try:
            if os.path.exists(self.ruta):
                os.remove(self.ruta)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Notificaciones de escritorio
# ---------------------------------------------------------------------------

def notificar(titulo: str, mensaje: str, timeout: int = DURACION_NOTIFICACION) -> bool:
    """
    Lanza una notificacion nativa de Windows.

    Estrategia en cascada:
      1. plyer.notification
      2. PowerShell + Windows.UI.Notifications (toast nativo, respaldo sin dependencias)

    Devuelve True si alguna via funciono.
    """
    # --- Via 1: plyer ---
    try:
        from plyer import notification

        notification.notify(
            title=titulo,
            message=mensaje,
            app_name=APP_NAME,
            timeout=timeout,
        )
        return True
    except Exception as exc:
        log.debug("plyer fallo (%s: %s); se intenta PowerShell.", type(exc).__name__, exc)

    # --- Via 2: PowerShell toast ---
    return _notificar_powershell(titulo, mensaje)


def _escapar_xml(texto: str) -> str:
    return (
        str(texto)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _notificar_powershell(titulo: str, mensaje: str) -> bool:
    """Respaldo: toast nativo via WinRT desde PowerShell."""
    plantilla = f"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null
$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml(@"
<toast>
  <visual>
    <binding template="ToastGeneric">
      <text>{_escapar_xml(titulo)}</text>
      <text>{_escapar_xml(mensaje)}</text>
    </binding>
  </visual>
</toast>
"@)
$toast = New-Object Windows.UI.Notifications.ToastNotification $xml
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('{_escapar_xml(APP_NAME)}').Show($toast)
"""
    try:
        import subprocess

        resultado = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", plantilla],
            capture_output=True,
            text=True,
            timeout=25,
        )
        if resultado.returncode == 0:
            return True
        log.debug("PowerShell devolvio codigo %s: %s", resultado.returncode, resultado.stderr.strip())
    except Exception as exc:
        log.debug("Respaldo PowerShell fallo: %s: %s", type(exc).__name__, exc)
    return False


# ---------------------------------------------------------------------------
# Seleccion de casos
# ---------------------------------------------------------------------------

def seleccionar_casos(resultado, solo_rojo: bool = False, sin_anti_duplicado: bool = False,
                      historial: Historial | None = None,
                      horas_anti_duplicado: int = HORAS_ANTI_DUPLICADO):
    """
    Devuelve (pendientes, omitidos_por_duplicado).

    pendientes: DataFrame de casos a notificar.
    Se excluye lo ya notificado con el MISMO estado dentro de la ventana de
    anti-duplicado; si el caso empeoro (cambio de estado), se vuelve a notificar.
    """
    import pandas as pd

    # La ventana de notificacion ya viene aplicada en resultado.df.
    pendientes = resultado.df[resultado.df["ESTADO"].isin(ESTADOS_ALERTA)].copy()

    if solo_rojo:
        pendientes = pendientes[pendientes["ESTADO"].isin([ROJO, CERRADO_TARDE])]

    vacio = pendientes.iloc[0:0]
    if pendientes.empty or sin_anti_duplicado or historial is None:
        return pendientes, vacio

    recientes = historial.estados_notificados_recientemente(minutos=horas_anti_duplicado * 60)
    if not recientes:
        return pendientes, vacio

    casos = pendientes[COL_CASO].astype(str)
    ya_igual = [
        recientes.get(c) == e
        for c, e in zip(casos, pendientes["ESTADO"].astype(str))
    ]
    mascara_duplicado = pd.Series(ya_igual, index=pendientes.index, dtype=bool)

    return pendientes[~mascara_duplicado], pendientes[mascara_duplicado]


# ---------------------------------------------------------------------------
# Envio
# ---------------------------------------------------------------------------

def enviar_alerta(fila, dry_run: bool = False) -> bool:
    """Construye y envia la notificacion de escritorio de un caso."""
    titulo, cuerpo = construir_mensaje_notificacion(fila)
    caso = str(fila.get(COL_CASO, "S/N")).strip()

    if dry_run:
        log.info("[DRY-RUN] %s | %s", _sin_emoji(titulo), _sin_emoji(cuerpo.replace("\n", " / ")))
        return True

    if notificar(titulo, cuerpo):
        log.info("Notificado: %s", _sin_emoji(titulo))
        return True

    log.error("No se pudo notificar el caso %s por ninguna via.", caso)
    return False


def registrar_historial(historial: Historial, pendientes, enviados: set[str],
                        canal: str, momento: datetime) -> None:
    """Guarda en el historial cada caso intentado, con su resultado."""
    for _, fila in pendientes.iterrows():
        caso = str(fila.get(COL_CASO, "S/N")).strip()
        historial.registrar(
            caso=caso,
            tecnico=str(fila.get("TECNICO", "")),
            region=str(fila.get("REGION_TECNICO", "")),
            estado=str(fila.get("ESTADO", "")),
            horas_restantes=fila.get("HORAS_RESTANTES"),
            horas_vencido=fila.get("HORAS_VENCIDO"),
            canal=canal,
            resultado="enviado" if caso in enviados else "fallido",
            detalle=str(fila.get("PREDICCION", "")),
            momento=momento,
        )


# ---------------------------------------------------------------------------
# Resumen por consola
# ---------------------------------------------------------------------------

def resumen(resultado) -> str:
    """
    Resumen de los conteos por estado.

    Se usa texto plano en lugar de emojis para que la salida sea segura en
    consolas cp1252 (caso tipico del Programador de Tareas).
    """
    lineas = [
        f"Casos en la hoja : {resultado.total_hoja}",
        f"Casos activos    : {resultado.total_activos}",
        f"En ventana       : {resultado.total_en_ventana} "
        f"(proximos {resultado.dias_ventana} dias + vencidos)",
        "",
        "TODOS LOS CASOS:",
    ]
    for estado, cantidad in resultado.conteo_completo.items():
        lineas.append(f"  [{estado}]".ljust(22) + f": {cantidad}")
    return "\n".join(lineas)


def _registrar_envio_telegram(historial: Historial, tabla_df, momento: datetime,
                              detalle: str) -> None:
    """Registra en el historial los casos enviados a Telegram."""
    for _, fila in tabla_df.iterrows():
        historial.registrar(
            caso=str(fila.get(COL_CASO, "S/N")).strip(),
            tecnico=str(fila.get("TECNICO", "")),
            region=str(fila.get("REGION_TECNICO", "")),
            estado=str(fila.get("ESTADO", "")),
            horas_restantes=fila.get("HORAS_RESTANTES"),
            horas_vencido=fila.get("HORAS_VENCIDO"),
            canal=CANAL_TELEGRAM,
            resultado="enviado",
            detalle=detalle,
            momento=momento,
        )


# ---------------------------------------------------------------------------
# Flujo principal
# ---------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Notificaciones de la Torre de Control SLA (Colsof - Banco Agrario)."
    )
    parser.add_argument("--solo-rojo", action="store_true",
                        help="Notificar unicamente ROJO y CERRADO TARDE.")
    parser.add_argument("--telegram", action="store_true",
                        help="Enviar tambien la tabla de estados al grupo de Telegram.")
    parser.add_argument("--por-region", action="store_true",
                        help="Distribuir la tabla segun el filtro de region de cada destino "
                             "(ej. un grupo para Bogota y otro para Regionales).")
    parser.add_argument("--telegram-plano", action="store_true",
                        help="Enviar la tabla a Telegram en texto plano (para WhatsApp).")
    parser.add_argument("--dry-run", action="store_true",
                        help="Simular sin notificar ni registrar en el historial.")
    parser.add_argument("--test", action="store_true",
                        help="Enviar una notificacion de prueba y salir.")
    parser.add_argument("--resumen", action="store_true",
                        help="Imprimir solo el resumen por consola, sin notificar.")
    parser.add_argument("--max", type=int, default=0,
                        help="Maximo de notificaciones de escritorio a enviar (0 = sin limite).")
    parser.add_argument("--max-telegram", type=int, default=0,
                        help="Maximo de filas en la tabla de Telegram (0 = sin limite).")
    parser.add_argument("--dias-ventana", type=int, default=None,
                        help="Dias calendario de la ventana de notificacion (por defecto 3).")
    parser.add_argument("--modo-ventana", choices=[MODO_DIAS, MODO_ACUMULADO],
                        default=MODO_VENTANA_POR_DEFECTO,
                        help="dias = ultimos N dias ACOTADO (por defecto, ej. 15/16/17). "
                             "acumulado = proximos N dias + todos los vencidos sin cerrar.")
    parser.add_argument("--horas-anti-duplicado", type=int, default=HORAS_ANTI_DUPLICADO,
                        help=f"No repetir el mismo caso y estado dentro de estas horas "
                             f"(por defecto {HORAS_ANTI_DUPLICADO}).")
    parser.add_argument("--sin-anti-duplicado", action="store_true",
                        help="Notificar todos los casos siempre, sin filtrar repetidos.")
    parser.add_argument("--excel", type=str, default=None, help="Ruta explicita del Excel.")
    parser.add_argument("--verbose", action="store_true", help="Log detallado.")
    args = parser.parse_args(argv)

    configurar_logging(args.verbose)

    # --- Modo prueba ------------------------------------------------------
    if args.test:
        log.info("Enviando notificacion de prueba...")
        ok = notificar(
            "[ROJO] PRUEBA - Torre de Control SLA",
            "Si ve este mensaje, las notificaciones funcionan correctamente.\n"
            f"Colsof - Banco Agrario\n{ahora_colombia():%Y-%m-%d %H:%M}",
        )
        print("Notificacion de prueba enviada." if ok else "No se pudo enviar la notificacion.")
        return 0 if ok else 1

    # --- Calculo ----------------------------------------------------------
    try:
        resultado = calcular_tablero(
            args.excel,
            dias_ventana=args.dias_ventana or 3,
            modo_ventana=args.modo_ventana,
        )
    except ErrorLecturaExcel as exc:
        # Caso tipico: el Excel esta abierto por otro usuario.
        log.error("No se pudo leer la plantilla: %s", exc)
        return 2

    momento = ahora_colombia()
    log.info("Calculo ejecutado a las %s", resultado.momentos.strftime("%Y-%m-%d %H:%M:%S"))
    for aviso in resultado.warnings:
        log.warning(_sin_emoji(aviso))

    print(resumen(resultado))
    print()
    print(resumen_consola(resultado))

    if args.resumen:
        return 0

    # --- Historial y anti-duplicado --------------------------------------
    historial = None if args.dry_run else Historial()
    pendientes, omitidos = seleccionar_casos(
        resultado,
        solo_rojo=args.solo_rojo,
        sin_anti_duplicado=args.sin_anti_duplicado,
        historial=historial,
        horas_anti_duplicado=args.horas_anti_duplicado,
    )

    if not omitidos.empty:
        log.info(
            "Omitidos por anti-duplicado (%d h): %d caso(s) ya notificados con el mismo estado.",
            args.horas_anti_duplicado, len(omitidos),
        )

    _vista_previa_anti_duplicado(resultado, args, historial, pendientes, omitidos)

    if pendientes.empty:
        log.info("No hay casos nuevos que requieran notificacion.")
    else:
        log.info("Casos a notificar: %d", len(pendientes))

        if args.max > 0:
            pendientes = pendientes.head(args.max)

        # --- Notificaciones de escritorio --------------------------------
        enviados: set[str] = set()
        for _, fila in pendientes.iterrows():
            caso = str(fila.get(COL_CASO, "S/N")).strip()
            if enviar_alerta(fila, dry_run=args.dry_run):
                enviados.add(caso)
            if not args.dry_run:
                time.sleep(0.6)  # evita saturar el centro de notificaciones

        log.info("Notificaciones de escritorio enviadas: %d/%d", len(enviados), len(pendientes))

        # --- Registro en el historial ------------------------------------
        canal = CANAL_WINDOWS
        if args.telegram or args.telegram_plano:
            canal = CANAL_AMBOS
        if historial is not None:
            registrar_historial(historial, pendientes, enviados, canal, momento)
            log.info("Registrado en el historial: %d fila(s).", len(pendientes))

    # --- Telegram ---------------------------------------------------------
    if args.telegram or args.telegram_plano:
        codigo = enviar_telegram(resultado, args, pendientes, historial, momento)
        return codigo
    return 0


def _vista_previa_anti_duplicado(resultado, args, historial, pendientes, omitidos) -> None:
    """
    Muestra un resumen de lo que se va a notificar, sin enviar nada.

    Ayuda a depurar el anti-duplicado sin disparar notificaciones reales.
    """
    total_alerta = len(resultado.df[resultado.df["ESTADO"].isin(ESTADOS_ALERTA)])
    print()
    print("-" * 70)
    print("SELECCION DE NOTIFICACIONES")
    print("-" * 70)
    print(f"  Casos en ventana de notificacion : {resultado.total_en_ventana}")
    print(f"  Casos que requieren atencion     : {total_alerta}")
    if historial is None:
        print("  Anti-duplicado                   : DESACTIVADO (dry-run)")
    else:
        print(f"  Anti-duplicado                   : {args.horas_anti_duplicado} h"
              + (" (desactivado)" if args.sin_anti_duplicado else ""))
    print(f"  Omitidos por ya notificados      : {len(omitidos)}")
    print(f"  A NOTIFICAR                      : {len(pendientes)}")
    if not pendientes.empty:
        conteo = pendientes["ESTADO"].value_counts().to_dict()
        detalle = "   ".join(f"{k}: {v}" for k, v in conteo.items())
        print(f"  Desglose por estado              : {detalle}")
    print("-" * 70)


def enviar_telegram(resultado, args, pendientes, historial, momento: datetime) -> int:
    """Envia la tabla de estados al grupo de Telegram."""
    try:
        import telegram_notifier as tg
    except ImportError as exc:
        log.error("No se pudo importar telegram_notifier: %s", exc)
        return 3

    if not tg.telegram_configurado():
        try:
            tg.cargar_config()
        except ValueError as exc:
            log.error("Telegram no esta configurado: %s", _sin_emoji(str(exc)))
            print(f"\n[ERROR] {exc}\n")
            return 3

    # Se envia la tabla de los casos que requieren atencion dentro de la ventana.
    # OJO: se debe dividir en bloques porque Telegram rechaza mensajes de mas de
    # 4096 caracteres. Con 76 casos la tabla completa ronda los 8000 caracteres,
    # asi que 0 (sin limite) haria fallar el envio.
    tabla_df = resultado.requieren_atencion
    if args.max_telegram > 0:
        tabla_df = tabla_df.head(args.max_telegram)

    if tabla_df.empty:
        log.info("No hay casos para enviar a Telegram.")
        return 0

    total_casos = len(tabla_df)
    tamano = len(tg.construir_tabla(tabla_df, "CASOS QUE REQUIEREN ATENCION"))
    filas_por_bloque = FILAS_TELEGRAM
    if tamano > tg.LONGITUD_MAXIMA:
        # Se reduce el tamano del bloque hasta que quepa en un mensaje.
        while filas_por_bloque > 5:
            muestra = tabla_df.head(filas_por_bloque)
            if len(tg.construir_tabla(muestra, "X")) <= tg.LONGITUD_MAXIMA:
                break
            filas_por_bloque -= 5
        log.info(
            "La tabla completa ocupa %d caracteres (limite %d): se divide en "
            "bloques de %d casos.",
            tamano, tg.LONGITUD_MAXIMA, filas_por_bloque,
        )

    # --- Distribucion por region (Bogota / Regionales) ---------------------
    if getattr(args, "por_region", False):
        log.info(
            "Distribucion por region: cada destino recibe solo los casos de su filtro."
        )
        resultados = tg.enviar_tabla_por_region(
            tabla_df, max_filas=filas_por_bloque, plano=args.telegram_plano
        )
        enviados = sum(resultados.values())
        for nombre, n in resultados.items():
            log.info("  %s: %d mensaje(s)", nombre, n)
        log.info("Bloques enviados a Telegram: %d (de %d casos)", enviados, total_casos)

        if enviados and historial is not None:
            _registrar_envio_telegram(historial, tabla_df, momento,
                                      "tabla enviada por region")
        return 0 if enviados else 1

    enviados = tg.enviar_tabla(
        tabla_df,
        titulo="CASOS QUE REQUIEREN ATENCION",
        max_filas=filas_por_bloque,
        plano=args.telegram_plano,
    )
    log.info("Bloques enviados a Telegram: %d (de %d casos)", enviados, total_casos)

    if enviados and historial is not None:
        _registrar_envio_telegram(historial, tabla_df, momento, "tabla enviada al grupo")

    return 0 if enviados else 1


if __name__ == "__main__":
    # Instancia unica: si ya hay una ejecucion en curso, se sale sin duplicar avisos.
    with InstanciaUnica() as puede_ejecutar:
        if not puede_ejecutar:
            print("Ya hay una ejecucion de alertas en curso. Se omite esta corrida.")
            sys.exit(0)
        sys.exit(main())


# ---------------------------------------------------------------------------
# PROGRAMADOR DE TAREAS DE WINDOWS
# ---------------------------------------------------------------------------
# Registrar la tarea para que corra cada 15 minutos (ejecutar en PowerShell COMO ADMINISTRADOR):
#
#   $py     = (Get-Command pythonw.exe).Source
#   $script = "C:\Users\darobles\Documents\PROYECTOS\TORRE CONTROL COLSOF\alertas_windows.py"
#   $accion = New-ScheduledTaskAction -Execute $py -Argument "`"$script`"" `
#               -WorkingDirectory "C:\Users\darobles\Documents\PROYECTOS\TORRE CONTROL COLSOF"
#   $disparo = New-ScheduledTaskTrigger -Once -At (Get-Date).Date.AddMinutes(1) `
#               -RepetitionInterval (New-TimeSpan -Minutes 15)
#   $ajustes = New-ScheduledTaskSettingsSet -StartWhenAvailable `
#               -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
#               -MultipleInstances IgnoreNew
#   Register-ScheduledTask -TaskName "TorreControlSLA-Colsof" -Action $accion `
#               -Trigger $disparo -Settings $ajustes -Description "Alertas SLA Colsof - Banco Agrario"
#
# Para incluir Telegram, agregue " --telegram" al final del Argument:
#   -Argument "`"$script`" --telegram"
#
# Probar de inmediato:      Start-ScheduledTask -TaskName "TorreControlSLA-Colsof"
# Ver ultimo resultado:     Get-ScheduledTaskInfo -TaskName "TorreControlSLA-Colsof"
# Eliminar la tarea:        Unregister-ScheduledTask -TaskName "TorreControlSLA-Colsof" -Confirm:$false
# ---------------------------------------------------------------------------
