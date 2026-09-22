"""
app_gui.py - Interfaz grafica de escritorio para Torre de Control SLA.

Ejecutar con:
    python app_gui.py

Requisitos:
    - Python 3.8+ (tkinter incluido)
    - core.py e historial.py en el mismo directorio

Componentes (ventana con pestanas ttk.Notebook):
    Pestana 1 "Casos":
        - Panel de KPIs del semaforo v2 (7 estados)
        - Filtros interactivos (Tecnico, Region, Estado, Busqueda)
        - Tabla de casos con semaforo coloreado
        - Detalle del caso seleccionado
    Pestana 2 "Metricas por tecnico":
        - Tabla de metricas por tecnico (core.metricas_por_tecnico)
    Pestana 3 "Notificaciones":
        - Notificaciones por tecnico (historial.por_tecnico)
        - Historial detallado (historial.leer, ultimas 500 filas)
    Barra inferior (Notificar, Exportar CSV, Dashboard, Refrescar) siempre visible
    y el boton Refrescar actualiza las tres pestanas.
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

import pandas as pd

from core import (
    AMARILLO,
    CERRADO_OK,
    CERRADO_TARDE,
    COL_CASO,
    COL_CIUDAD,
    COL_TECNICO,
    ICONO_ESTADO,
    NARANJA,
    ROJO,
    SIN_VENCIMIENTO,
    TODOS_LOS_ESTADOS,
    VERDE,
    ErrorLecturaExcel,
    ahora_colombia,
    calcular_tablero,
    construir_mensaje_notificacion,
    metricas_por_tecnico,
    obtener_region,
)
from historial import (
    CANAL_WHATSAPP,
    RESULTADO_GENERADO,
    Historial,
)

# Avisos por tecnico (composicion del mensaje para WhatsApp / Telegram).
import avisos

DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_EXCEL = os.path.join(DIR_SCRIPT, "PLANTILLA DE SEGUIMIENTO DE CASOS SEPTIEMBRE.xlsx")

# ---------------------------------------------------------------------------
# Colores del semaforo v2 (7 estados)
# ---------------------------------------------------------------------------
ROJO_BG = "#FFE0E0"            # rojo claro  -> vencido sin cerrar
ROJO_FG = "#A30000"
CERRADO_TARDE_BG = "#8B0000"   # rojo oscuro -> incumplimiento consumado
CERRADO_TARDE_FG = "#FFFFFF"
NARANJA_BG = "#FFE0B2"         # naranja claro -> menos de 1 hora
NARANJA_FG = "#8A4B00"
AMARILLO_BG = "#FFF6D5"        # amarillo claro -> entre 1 y 4 horas
AMARILLO_FG = "#7A5B00"
VERDE_BG = "#E3F7E3"           # verde claro -> mas de 4 horas
VERDE_FG = "#136B13"
SIN_BG = "#EFEFEF"             # gris -> sin fecha de vencimiento
SIN_FG = "#555555"
CERRADO_OK_BG = "#D6F5D6"      # verde claro -> cerrado a tiempo
CERRADO_OK_FG = "#0B5D0B"

COLORES_FONDO = {
    ROJO: ROJO_BG,
    CERRADO_TARDE: CERRADO_TARDE_BG,
    NARANJA: NARANJA_BG,
    AMARILLO: AMARILLO_BG,
    VERDE: VERDE_BG,
    SIN_VENCIMIENTO: SIN_BG,
    CERRADO_OK: CERRADO_OK_BG,
}
COLORES_TEXTO = {
    ROJO: ROJO_FG,
    CERRADO_TARDE: CERRADO_TARDE_FG,
    NARANJA: NARANJA_FG,
    AMARILLO: AMARILLO_FG,
    VERDE: VERDE_FG,
    SIN_VENCIMIENTO: SIN_FG,
    CERRADO_OK: CERRADO_OK_FG,
}

INTERVALO_AUTO_REFRESH_S = 60
MAX_FILAS_NOTIFICACIONES = 500
# Medidas del recuadro de detalle (en caracteres).
ANCHO_ETIQUETA_DETALLE = 16
ANCHO_VALOR_DETALLE = 33
ANCHO_INTERIOR_DETALLE = 2 + ANCHO_ETIQUETA_DETALLE + ANCHO_VALOR_DETALLE + 1

# Columnas de las tablas de notificaciones (para dejar tablas vacias con forma).
COLS_NOTIF_TECNICO = [
    "TECNICO", "REGION", "NOTIFICACIONES", "CASOS_DISTINTOS",
    "ULTIMA_NOTIFICACION", "ROJOS", "NARANJAS", "AMARILLOS", "FALLIDAS",
]
COLS_NOTIF_DETALLE = [
    "id", "fecha_hora", "caso", "tecnico", "region", "estado",
    "horas_restantes", "horas_vencido", "canal", "resultado", "detalle",
]
# Columnas de avisos.resumen_pendientes (para dejar la tabla vacia con forma).
COLS_RESUMEN_PENDIENTES = [
    "TECNICO", "REGION", "PENDIENTES", "ROJOS", "NARANJAS", "AMARILLOS",
    "PROXIMO_VENCIMIENTO", "ULTIMA_NOTIFICACION", "VECES_HOY",
]


def texto_seguro(valor, defecto: str = "—") -> str:
    """Convierte un valor de pandas a texto, usando '—' para vacios/NaN."""
    if valor is None:
        return defecto
    try:
        if pd.isna(valor):
            return defecto
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    return texto if texto and texto.lower() not in ("nan", "nat", "none") else defecto


def formatear_pct(valor) -> str:
    """Porcentaje de cumplimiento: '—' cuando es None/NaN (nunca 'nan')."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return "—"
    try:
        if pd.isna(valor):
            return "—"
    except (TypeError, ValueError):
        pass
    return f"{float(valor):.1f} %"


def formatear_momento(valor, defecto: str = "—") -> str:
    """
    Fecha/hora en formato dd/mm/yyyy HH:MM para la tabla de pendientes.

    'defecto' se usa cuando no hay valor: 'nunca' para ULTIMA_NOTIFICACION y '—'
    para el proximo vencimiento.
    """
    if valor is None:
        return defecto
    try:
        if pd.isna(valor):
            return defecto
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    if not texto or texto.lower() in ("nan", "nat", "none", ""):
        return defecto
    try:
        return pd.Timestamp(valor).strftime("%d/%m/%Y %H:%M")
    except (ValueError, TypeError):
        return texto


class TorreControlApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("🛰️ Torre de Control SLA — Colsof / Banco Agrario")
        self.root.geometry("1280x780")
        self.root.minsize(1024, 640)
        self.root.configure(bg="#F5F5F5")

        self._df: pd.DataFrame | None = None            # casos en ventana (pestana Casos)
        self._df_completo: pd.DataFrame = pd.DataFrame()  # todos los casos (metricas)
        self._conteo_completo: dict[str, int] = {}
        self._metricas: pd.DataFrame = pd.DataFrame()
        self._hist_tecnico: pd.DataFrame = pd.DataFrame()
        self._hist_detalle: pd.DataFrame = pd.DataFrame()
        self._hist_resumen: dict[str, int] = {}
        self._hist_error: str = ""
        self._pendientes: pd.DataFrame = pd.DataFrame()       # casos de ESTADOS_ALERTA
        self._resumen_pendientes: pd.DataFrame = pd.DataFrame()
        self._grupos_pendientes: dict[str, pd.DataFrame] = {}
        self._menciones_pendientes: dict = {}
        self._filas_tabla: dict[str, pd.Series] = {}    # iid -> fila del caso
        self._timer_id: str | None = None
        self._seleccion_filas: list[int] = []

        self._construir_menu()
        self._construir_barra_superior()
        self._construir_tabs()
        self._construir_barra_inferior()

        self._refrescar_datos()

    # ------------------------------------------------------------------
    # Menu
    # ------------------------------------------------------------------

    def _construir_menu(self) -> None:
        menu_bar = tk.Menu(self.root)
        self.root.config(menu=menu_bar)

        menu_archivo = tk.Menu(menu_bar, tearoff=0)
        menu_archivo.add_command(label="🔄 Refrescar datos", command=self._on_refrescar)
        menu_archivo.add_separator()
        menu_archivo.add_command(
            label="🚀 Abrir Dashboard Streamlit", command=self._abrir_dashboard
        )
        menu_archivo.add_command(
            label="📤 Exportar CSV de casos", command=self._exportar_csv
        )
        menu_archivo.add_command(
            label="📊 Exportar CSV de metricas por tecnico",
            command=self._exportar_metricas_csv,
        )
        menu_archivo.add_command(
            label="🔔 Exportar CSV de notificaciones",
            command=self._exportar_notificaciones_csv,
        )
        menu_archivo.add_separator()
        menu_archivo.add_command(label="🚪 Salir", command=self.root.quit)
        menu_bar.add_cascade(label="Archivo", menu=menu_archivo)

        menu_acciones = tk.Menu(menu_bar, tearoff=0)
        menu_acciones.add_command(
            label="🔴🟡 Enviar notificaciones (Rojo/Amarillo)",
            command=self._notificar,
        )
        menu_acciones.add_command(
            label="🔴 Enviar notificaciones (solo Rojo)",
            command=lambda: self._notificar(solo_rojo=True),
        )
        menu_bar.add_cascade(label="Acciones", menu=menu_acciones)

        menu_ayuda = tk.Menu(menu_bar, tearoff=0)
        menu_ayuda.add_command(label="ℹ️ Acerca de", command=self._acerca_de)
        menu_bar.add_cascade(label="Ayuda", menu=menu_ayuda)

    # ------------------------------------------------------------------
    # Barra superior
    # ------------------------------------------------------------------

    def _construir_barra_superior(self) -> None:
        frame = ttk.Frame(self.root, padding=8)
        frame.pack(fill=tk.X)

        ttk.Label(
            frame,
            text="🛰️ Torre de Control SLA",
            font=("Segoe UI", 16, "bold"),
            foreground="#1A237E",
        ).pack(side=tk.LEFT)

        self.lbl_hora = ttk.Label(
            frame,
            text="",
            font=("Consolas", 10),
            foreground="#616161",
        )
        self.lbl_hora.pack(side=tk.RIGHT, padx=8)

        ttk.Button(
            frame, text="🔄 Refrescar", command=self._on_refrescar
        ).pack(side=tk.RIGHT, padx=4)

        ttk.Button(
            frame, text="🚀 Dashboard", command=self._abrir_dashboard
        ).pack(side=tk.RIGHT, padx=4)

    # ------------------------------------------------------------------
    # Pestanas (ttk.Notebook)
    # ------------------------------------------------------------------

    def _construir_tabs(self) -> None:
        """Crea el Notebook con las pestanas y llena cada una."""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=(2, 0))

        self.tab_casos = ttk.Frame(self.notebook)
        self.tab_metricas = ttk.Frame(self.notebook)
        self.tab_notificaciones = ttk.Frame(self.notebook)
        self.tab_pendientes = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_casos, text="  📋 Casos  ")
        self.notebook.add(self.tab_metricas, text="  👷 Metricas por tecnico  ")
        self.notebook.add(self.tab_notificaciones, text="  🔔 Notificaciones  ")
        self.notebook.add(self.tab_pendientes, text="  📨 Notificaciones pendientes  ")

        # --- Pestana 1: todo lo que ya existia -------------------------
        self._construir_kpis(self.tab_casos)
        self._construir_filtros(self.tab_casos)
        self._construir_tabla(self.tab_casos)
        self._construir_detalle(self.tab_casos)

        # --- Pestanas 2, 3 y 4 -----------------------------------------
        self._construir_metricas(self.tab_metricas)
        self._construir_notificaciones(self.tab_notificaciones)
        self._construir_pendientes(self.tab_pendientes)

    # ------------------------------------------------------------------
    # Utilidad: Treeview con barras de desplazamiento
    # ------------------------------------------------------------------

    def _crear_tree(
        self,
        frame: ttk.Frame,
        columns: tuple[str, ...],
        headings: dict[str, str],
        widths: dict[str, int],
        anchors: dict[str, str] | None = None,
    ) -> ttk.Treeview:
        """Crea un Treeview con encabezados y scrollbars, y lo ubica con grid."""
        tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        anchors = anchors or {}
        for col in columns:
            tree.heading(col, text=headings.get(col, col), anchor=tk.W)
            tree.column(
                col,
                width=widths.get(col, 100),
                minwidth=40,
                anchor=anchors.get(col, tk.W),
            )

        scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        return tree

    # ------------------------------------------------------------------
    # KPIs
    # ------------------------------------------------------------------

    def _construir_kpis(self, parent: tk.Widget) -> None:
        frame = ttk.Frame(parent, padding=(8, 4, 8, 4))
        frame.pack(fill=tk.X)

        self.kpi_vars: dict[str, tk.StringVar] = {}
        self.kpi_labels: dict[str, ttk.Label] = {}
        self.kpi_frames: dict[str, ttk.Frame] = {}

        # 7 tarjetas, en el orden de gravedad de core.ORDEN_ESTADO.
        for estado in TODOS_LOS_ESTADOS:
            bg = COLORES_FONDO[estado]
            fg = COLORES_TEXTO[estado]

            # Nota: tk.Frame (Tk clasico) NO admite la opcion 'padding'; esta solo
            # existe en ttk.Frame. El margen interior se logra con padx/pady al
            # empaquetar los hijos (ver lbl_count.pack / lbl_name.pack mas abajo).
            f = tk.Frame(frame, bg=bg, relief="groove", borderwidth=2)
            f.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=3)
            self.kpi_frames[estado] = f

            var = tk.StringVar(value="0")
            lbl_count = tk.Label(
                f,
                textvariable=var,
                font=("Segoe UI", 22, "bold"),
                bg=bg,
                fg=fg,
            )
            lbl_count.pack(padx=8, pady=(8, 0))

            lbl_name = tk.Label(
                f,
                text=f"{ICONO_ESTADO.get(estado, '')} {estado}",
                font=("Segoe UI", 9),
                bg=bg,
                fg=fg,
            )
            lbl_name.pack(padx=8, pady=(0, 8))

            self.kpi_vars[estado] = var
            self.kpi_labels[estado] = lbl_count

    # ------------------------------------------------------------------
    # Filtros
    # ------------------------------------------------------------------

    def _construir_filtros(self, parent: tk.Widget) -> None:
        frame = ttk.LabelFrame(parent, text="🔎 Filtros", padding=8)
        frame.pack(fill=tk.X, padx=8, pady=(4, 4))

        filtros_frame = ttk.Frame(frame)
        filtros_frame.pack(fill=tk.X)

        ttk.Label(filtros_frame, text="👷 Tecnico:").pack(side=tk.LEFT, padx=(0, 4))
        self.combo_tecnico = ttk.Combobox(filtros_frame, width=35, state="readonly")
        self.combo_tecnico.pack(side=tk.LEFT, padx=4)
        self.combo_tecnico.bind("<<ComboboxSelected>>", lambda _: self._aplicar_filtros())

        ttk.Label(filtros_frame, text="📍 Region:").pack(side=tk.LEFT, padx=(16, 4))
        self.combo_region = ttk.Combobox(filtros_frame, width=20, state="readonly")
        self.combo_region.pack(side=tk.LEFT, padx=4)
        self.combo_region.bind("<<ComboboxSelected>>", lambda _: self._aplicar_filtros())

        ttk.Label(filtros_frame, text="🔍 Buscar:").pack(side=tk.LEFT, padx=(16, 4))
        self.entry_buscar = ttk.Entry(filtros_frame, width=25)
        self.entry_buscar.pack(side=tk.LEFT, padx=4)
        self.entry_buscar.bind("<Return>", lambda _: self._aplicar_filtros())

        ttk.Label(filtros_frame, text="Estado:").pack(side=tk.LEFT, padx=(16, 4))
        # Los 7 estados del semaforo v2, mas la opcion "Todos".
        self.combo_estado = ttk.Combobox(
            filtros_frame,
            values=["Todos"] + list(TODOS_LOS_ESTADOS),
            state="readonly",
            width=18,
        )
        self.combo_estado.set("Todos")
        self.combo_estado.pack(side=tk.LEFT, padx=4)
        self.combo_estado.bind("<<ComboboxSelected>>", lambda _: self._aplicar_filtros())

        ttk.Button(
            filtros_frame, text="🗑️ Limpiar", command=self._limpiar_filtros
        ).pack(side=tk.RIGHT, padx=8)

    # ------------------------------------------------------------------
    # Tabla de casos (pestana 1)
    # ------------------------------------------------------------------

    def _construir_tabla(self, parent: tk.Widget) -> None:
        frame = ttk.LabelFrame(parent, text="📋 Casos en ventana", padding=8)
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        columns = (
            "icono",
            "caso",
            "tecnico",
            "region",
            "ciudad",
            "vencimiento",
            "horas",
            "estado",
        )
        headings = {
            "icono": " ",
            "caso": "Caso",
            "tecnico": "Tecnico",
            "region": "Region",
            "ciudad": "Ciudad",
            "vencimiento": "Tiempo restante",
            "horas": "Horas restantes",
            "estado": "Estado",
        }
        widths = {
            "icono": 30,
            "caso": 120,
            "tecnico": 160,
            "region": 120,
            "ciudad": 120,
            "vencimiento": 170,
            "horas": 90,
            "estado": 100,
        }
        self.tree = self._crear_tree(frame, columns, headings, widths)

        # Colores del semaforo v2: 7 estados (se configuran una sola vez).
        for estado in TODOS_LOS_ESTADOS:
            self.tree.tag_configure(
                f"tag_{estado}",
                background=COLORES_FONDO[estado],
                foreground=COLORES_TEXTO[estado],
            )

        self.tree.bind("<<TreeviewSelect>>", self._mostrar_detalle)
        self.tree.bind("<<TreeviewSelect>>", self._actualizar_kpi_seleccion, add="+")

    # ------------------------------------------------------------------
    # Detalle (pestana 1)
    # ------------------------------------------------------------------

    def _construir_detalle(self, parent: tk.Widget) -> None:
        frame = ttk.LabelFrame(parent, text="📌 Detalle del caso", padding=8)
        frame.pack(fill=tk.X, padx=8, pady=(0, 4))

        self.txt_detalle = scrolledtext.ScrolledText(
            frame,
            height=8,
            font=("Consolas", 10),
            wrap=tk.WORD,
            state=tk.DISABLED,
            bg="#FAFAFA",
            relief="sunken",
            borderwidth=1,
        )
        self.txt_detalle.pack(fill=tk.X)

    # ------------------------------------------------------------------
    # Pestana 2: metricas por tecnico
    # ------------------------------------------------------------------

    def _construir_metricas(self, parent: tk.Widget) -> None:
        frame = ttk.LabelFrame(
            parent, text="👷 Metricas por tecnico (todos los casos)", padding=8
        )
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        columns = (
            "tecnico", "region", "asignados", "abiertos", "vencidos", "proximos",
            "naranja", "amarillo", "cerrados", "cerrados_tarde", "cumplimiento",
            "horas_total", "horas_prom",
        )
        headings = {
            "tecnico": "Tecnico",
            "region": "Region",
            "asignados": "Asignados",
            "abiertos": "Abiertos",
            "vencidos": "Vencidos",
            "proximos": "Prox. vencer",
            "naranja": "Naranja",
            "amarillo": "Amarillo",
            "cerrados": "Cerrados",
            "cerrados_tarde": "Cerrados tarde",
            "cumplimiento": "Cumplimiento",
            "horas_total": "Hrs vencido tot.",
            "horas_prom": "Hrs vencido prom.",
        }
        widths = {
            "tecnico": 230, "region": 110, "asignados": 80, "abiertos": 75,
            "vencidos": 75, "proximos": 90, "naranja": 70, "amarillo": 70,
            "cerrados": 75, "cerrados_tarde": 100, "cumplimiento": 95,
            "horas_total": 105, "horas_prom": 110,
        }
        anchors = {
            "asignados": tk.E, "abiertos": tk.E, "vencidos": tk.E, "proximos": tk.E,
            "naranja": tk.E, "amarillo": tk.E, "cerrados": tk.E,
            "cerrados_tarde": tk.E, "cumplimiento": tk.E, "horas_total": tk.E,
            "horas_prom": tk.E,
        }
        self.tree_metricas = self._crear_tree(frame, columns, headings, widths, anchors)
        # Resaltado suave de los tecnicos con casos vencidos.
        self.tree_metricas.tag_configure("tag_vencidos", background=ROJO_BG, foreground=ROJO_FG)

        self.lbl_metricas = ttk.Label(
            parent,
            text="Metricas por tecnico: sin datos.",
            font=("Segoe UI", 9),
            foreground="#616161",
        )
        self.lbl_metricas.pack(fill=tk.X, padx=10, pady=(0, 6))

    # ------------------------------------------------------------------
    # Pestana 3: notificaciones
    # ------------------------------------------------------------------

    def _construir_notificaciones(self, parent: tk.Widget) -> None:
        self.lbl_notif = ttk.Label(
            parent,
            text="Historial de notificaciones: sin datos.",
            font=("Segoe UI", 9),
            foreground="#616161",
        )
        self.lbl_notif.pack(fill=tk.X, padx=10, pady=(6, 0))

        contenedor = ttk.Frame(parent)
        contenedor.pack(fill=tk.BOTH, expand=True)
        contenedor.grid_rowconfigure(0, weight=1)
        contenedor.grid_rowconfigure(1, weight=3)
        contenedor.grid_columnconfigure(0, weight=1)

        # --- Arriba: notificaciones por tecnico -------------------------
        marco_tec = ttk.LabelFrame(
            contenedor, text="🔔 Notificaciones por tecnico", padding=6
        )
        marco_tec.grid(row=0, column=0, sticky="nsew", padx=8, pady=(4, 4))

        cols_tec = (
            "tecnico", "region", "notificaciones", "casos", "ultima",
            "rojos", "naranjas", "amarillos", "fallidas",
        )
        head_tec = {
            "tecnico": "Tecnico", "region": "Region", "notificaciones": "Notificaciones",
            "casos": "Casos distintos", "ultima": "Ultima notificacion",
            "rojos": "Rojos", "naranjas": "Naranjas", "amarillos": "Amarillos",
            "fallidas": "Fallidas",
        }
        anchos_tec = {
            "tecnico": 240, "region": 110, "notificaciones": 100, "casos": 100,
            "ultima": 150, "rojos": 70, "naranjas": 80, "amarillos": 80, "fallidas": 75,
        }
        anchors_tec = {
            "notificaciones": tk.E, "casos": tk.E, "rojos": tk.E, "naranjas": tk.E,
            "amarillos": tk.E, "fallidas": tk.E,
        }
        self.tree_notif_tecnico = self._crear_tree(
            marco_tec, cols_tec, head_tec, anchos_tec, anchors_tec
        )

        # --- Abajo: historial detallado --------------------------------
        marco_det = ttk.LabelFrame(
            contenedor,
            text=(
                "🗂️ Historial detallado de notificaciones "
                f"(ultimas {MAX_FILAS_NOTIFICACIONES} filas)"
            ),
            padding=6,
        )
        marco_det.grid(row=1, column=0, sticky="nsew", padx=8, pady=(4, 8))

        cols_det = (
            "id", "fecha_hora", "caso", "tecnico", "region", "estado",
            "horas_restantes", "horas_vencido", "canal", "resultado", "detalle",
        )
        head_det = {
            "id": "ID", "fecha_hora": "Fecha y hora", "caso": "Caso",
            "tecnico": "Tecnico", "region": "Region", "estado": "Estado",
            "horas_restantes": "Hrs restantes", "horas_vencido": "Hrs vencido",
            "canal": "Canal", "resultado": "Resultado", "detalle": "Detalle",
        }
        anchos_det = {
            "id": 60, "fecha_hora": 145, "caso": 120, "tecnico": 230,
            "region": 110, "estado": 120, "horas_restantes": 95,
            "horas_vencido": 90, "canal": 120, "resultado": 85, "detalle": 260,
        }
        anchors_det = {
            "horas_restantes": tk.E, "horas_vencido": tk.E, "resultado": tk.CENTER,
        }
        self.tree_notif_detalle = self._crear_tree(
            marco_det, cols_det, head_det, anchos_det, anchors_det
        )

    # ------------------------------------------------------------------
    # Pestana 4: notificaciones pendientes a tecnicos (avisos.py)
    # ------------------------------------------------------------------

    def _construir_pendientes(self, parent: tk.Widget) -> None:
        """
        Pestana con los casos pendientes de notificar agrupados por tecnico.

        Arriba: resumen por tecnico (avisos.resumen_pendientes) y selector de canal.
        Abajo: a la izquierda la tabla de tecnicos, a la derecha el aviso generado
        (avisos.componer_aviso_tecnico) y los botones de copiar/registrar.
        """
        self.lbl_pendientes = ttk.Label(
            parent,
            text="Notificaciones pendientes: sin datos.",
            font=("Segoe UI", 9),
            foreground="#616161",
        )
        self.lbl_pendientes.pack(fill=tk.X, padx=10, pady=(6, 0))

        # --- Barra de canal ---------------------------------------------
        barra = ttk.Frame(parent)
        barra.pack(fill=tk.X, padx=8, pady=(4, 2))

        ttk.Label(barra, text="📡 Canal de la mencion:").pack(side=tk.LEFT)
        self.combo_canal_pend = ttk.Combobox(
            barra, values=["whatsapp", "telegram"], state="readonly", width=12
        )
        self.combo_canal_pend.set("whatsapp")
        self.combo_canal_pend.pack(side=tk.LEFT, padx=6)
        self.combo_canal_pend.bind(
            "<<ComboboxSelected>>", lambda _: self._mostrar_aviso_tecnico()
        )

        ttk.Label(
            barra,
            text=(
                "whatsapp = telefono de menciones.json · telegram = @usuario · "
                "el texto del aviso es el mismo"
            ),
            font=("Segoe UI", 8),
            foreground="#757575",
        ).pack(side=tk.LEFT, padx=10)

        # --- Contenedor: tabla (izquierda) y aviso (derecha) ------------
        contenedor = ttk.Frame(parent)
        contenedor.pack(fill=tk.BOTH, expand=True)
        contenedor.grid_rowconfigure(0, weight=1)
        contenedor.grid_columnconfigure(0, weight=0)
        contenedor.grid_columnconfigure(1, weight=1)

        marco_tabla = ttk.LabelFrame(
            contenedor, text="👷 Tecnicos con casos pendientes", padding=6
        )
        marco_tabla.grid(row=0, column=0, sticky="nsew", padx=(8, 4), pady=4)

        cols = (
            "tecnico", "region", "pendientes", "rojos", "naranjas", "amarillos",
            "proximo", "ultima", "hoy",
        )
        head = {
            "tecnico": "Tecnico", "region": "Region", "pendientes": "Pend.",
            "rojos": "Rojos", "naranjas": "Naran.", "amarillos": "Amar.",
            "proximo": "Prox. vencimiento", "ultima": "Ultima notificacion",
            "hoy": "Veces hoy",
        }
        anchos = {
            "tecnico": 210, "region": 105, "pendientes": 55, "rojos": 50,
            "naranjas": 60, "amarillos": 55, "proximo": 130, "ultima": 135,
            "hoy": 65,
        }
        anchors = {
            "pendientes": tk.E, "rojos": tk.E, "naranjas": tk.E, "amarillos": tk.E,
            "hoy": tk.E,
        }
        self.tree_pendientes = self._crear_tree(
            marco_tabla, cols, head, anchos, anchors
        )
        # Resaltado suave de los tecnicos con casos vencidos (ROJO).
        self.tree_pendientes.tag_configure(
            "tag_con_rojos", background=ROJO_BG, foreground=ROJO_FG
        )
        self.tree_pendientes.bind(
            "<<TreeviewSelect>>", self._mostrar_aviso_tecnico
        )

        marco_aviso = ttk.LabelFrame(
            contenedor, text="✉️ Aviso para el tecnico seleccionado", padding=6
        )
        marco_aviso.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=4)

        self.txt_aviso = tk.Text(
            marco_aviso,
            wrap="word",
            font=("Segoe UI", 10),
            bg="#FAFAFA",
            relief="sunken",
            borderwidth=1,
            height=18,
        )
        scroll_aviso = ttk.Scrollbar(
            marco_aviso, orient=tk.VERTICAL, command=self.txt_aviso.yview
        )
        self.txt_aviso.configure(yscrollcommand=scroll_aviso.set)
        self.txt_aviso.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_aviso.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_aviso.insert(
            tk.END,
            "Seleccione un tecnico en la tabla de la izquierda para generar su aviso.",
        )
        self.txt_aviso.config(state=tk.DISABLED)

        # --- Botones ----------------------------------------------------
        botones = ttk.Frame(parent)
        botones.pack(fill=tk.X, padx=8, pady=(2, 8))

        self.btn_copiar_aviso = ttk.Button(
            botones, text="📋 Copiar aviso", command=self._copiar_aviso
        )
        self.btn_copiar_aviso.pack(side=tk.LEFT, padx=4)

        self.btn_marcar_pendiente = ttk.Button(
            botones,
            text="✅ Marcar como notificado",
            command=self._marcar_pendiente_notificado,
        )
        self.btn_marcar_pendiente.pack(side=tk.LEFT, padx=4)

        self.btn_marcar_todos = ttk.Button(
            botones,
            text="📣 Marcar TODOS como notificados",
            command=self._marcar_todos_notificados,
        )
        self.btn_marcar_todos.pack(side=tk.LEFT, padx=4)

    def _texto_aviso_actual(self) -> str:
        """Devuelve el aviso que se esta mostrando en el recuadro de la pestana."""
        return self.txt_aviso.get("1.0", tk.END).strip()

    def _canal_pendiente(self) -> str:
        """Canal elegido en la pestana (whatsapp por defecto)."""
        canal = self.combo_canal_pend.get().strip().lower()
        return canal if canal in ("whatsapp", "telegram") else "whatsapp"

    def _poblar_pendientes(self) -> None:
        """Llena la tabla de tecnicos con casos pendientes de notificar."""
        for item in self.tree_pendientes.get_children():
            self.tree_pendientes.delete(item)

        tabla = self._resumen_pendientes
        if tabla is None or tabla.empty:
            self.lbl_pendientes.config(
                text=(
                    "Notificaciones pendientes: ningun caso requiere aviso "
                    "(ROJO, CERRADO TARDE, NARANJA o AMARILLO) en la ventana."
                )
            )
            self._limpiar_aviso(
                "No hay tecnicos con casos pendientes de notificar."
            )
            return

        for _, fila in tabla.iterrows():
            nombre = texto_seguro(fila.get("TECNICO"), "")
            if not nombre:
                continue
            rojos = int(fila.get("ROJOS", 0) or 0)
            tags = ("tag_con_rojos",) if rojos > 0 else ()
            self.tree_pendientes.insert(
                "",
                tk.END,
                iid=nombre,
                values=(
                    nombre,
                    texto_seguro(fila.get("REGION"), ""),
                    int(fila.get("PENDIENTES", 0) or 0),
                    rojos,
                    int(fila.get("NARANJAS", 0) or 0),
                    int(fila.get("AMARILLOS", 0) or 0),
                    formatear_momento(fila.get("PROXIMO_VENCIMIENTO")),
                    formatear_momento(fila.get("ULTIMA_NOTIFICACION"), "nunca"),
                    int(fila.get("VECES_HOY", 0) or 0),
                ),
                tags=tags,
            )

        total_casos = int(tabla["PENDIENTES"].sum())
        total_rojos = int(tabla["ROJOS"].sum())
        self.lbl_pendientes.config(
            text=(
                f"Notificaciones pendientes: {len(tabla)} tecnico(s) | "
                f"casos por notificar: {total_casos} | rojos sin cerrar: {total_rojos} | "
                "las filas resaltadas tienen casos vencidos (ROJO)"
            )
        )

    def _limpiar_aviso(self, mensaje: str) -> None:
        """Reemplaza el contenido del recuadro del aviso."""
        self.txt_aviso.config(state=tk.NORMAL)
        self.txt_aviso.delete("1.0", tk.END)
        self.txt_aviso.insert(tk.END, mensaje)
        self.txt_aviso.config(state=tk.DISABLED)

    def _mostrar_aviso_tecnico(self, event=None) -> None:
        """Genera y muestra el aviso del tecnico seleccionado en la tabla."""
        seleccion = self.tree_pendientes.selection()
        if not seleccion:
            return

        tecnico = str(seleccion[0])
        df_tecnico = self._grupos_pendientes.get(tecnico)
        if df_tecnico is None or df_tecnico.empty:
            self._limpiar_aviso(
                f"Sin casos pendientes para {tecnico}. Pulse Refrescar."
            )
            return

        try:
            texto = avisos.componer_aviso_tecnico(
                tecnico,
                df_tecnico,
                canal=self._canal_pendiente(),
                un_solo_mensaje=True,
                menciones=self._menciones_pendientes,
            )
        except Exception as exc:  # la GUI nunca debe caerse por un aviso
            self._limpiar_aviso(f"No se pudo generar el aviso: {type(exc).__name__}: {exc}")
            return

        self._limpiar_aviso(texto)
        self.lbl_status.config(
            text=(
                f"Aviso generado para {tecnico} "
                f"({len(df_tecnico)} caso(s), canal {self._canal_pendiente()})"
            )
        )

    def _copiar_aviso(self) -> None:
        """Copia el aviso mostrado al portapapeles del sistema."""
        texto = self._texto_aviso_actual()
        if not texto:
            messagebox.showwarning(
                "Sin aviso", "Seleccione primero un tecnico para generar su aviso."
            )
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(texto)
            self.root.update_idletasks()  # asegura que el portapapeles tome el texto
        except tk.TclError as exc:
            messagebox.showerror("Portapapeles", f"No se pudo copiar el aviso:\n{exc}")
            return

        self.lbl_status.config(text="📋 Aviso copiado al portapapeles: pegue en WhatsApp")
        messagebox.showinfo(
            "Aviso copiado",
            "El aviso se copio al portapapeles.\n\n"
            "Abra WhatsApp, elija el chat del tecnico y pegue con Ctrl+V.\n"
            "El formato y los emojis se conservan.",
        )

    def _registrar_casos_pendientes(self, df_casos: pd.DataFrame) -> int:
        """
        Registra en el historial UN registro por caso (canal whatsapp).

        Es el flujo de copiar y pegar: la torre genero el aviso y lo envio a mano
        por WhatsApp. Por eso se registra con resultado='generado' y NO 'enviado':
        el sistema no lo envio, solo lo redacto. Asi el log distingue lo que la
        aplicacion envio de lo que una persona pego manualmente.
        Devuelve cuantos casos se registraron.
        """
        if df_casos is None or df_casos.empty:
            return 0

        try:
            historial = Historial()
        except Exception as exc:
            messagebox.showerror(
                "Historial", f"No se pudo abrir el historial:\n{type(exc).__name__}: {exc}"
            )
            return 0

        momento = ahora_colombia()
        registrados = 0
        for _, fila in df_casos.iterrows():
            historial.registrar(
                caso=str(fila.get(COL_CASO, "S/N")),
                tecnico=str(fila.get("TECNICO", "") or ""),
                region=str(fila.get("REGION_TECNICO", "") or ""),
                estado=str(fila.get("ESTADO", "") or ""),
                horas_restantes=fila.get("HORAS_RESTANTES"),
                horas_vencido=fila.get("HORAS_VENCIDO"),
                canal=CANAL_WHATSAPP,
                resultado=RESULTADO_GENERADO,
                detalle="aviso copiado y enviado manualmente por WhatsApp",
                momento=momento,
            )
            registrados += 1
        return registrados

    def _marcar_pendiente_notificado(self) -> None:
        """Marca como notificados (en el historial) los casos del tecnico elegido."""
        seleccion = self.tree_pendientes.selection()
        if not seleccion:
            messagebox.showwarning(
                "Sin seleccion",
                "Seleccione un tecnico en la tabla para marcarlo como notificado.",
            )
            return

        tecnico = str(seleccion[0])
        df_tecnico = self._grupos_pendientes.get(tecnico)
        if df_tecnico is None or df_tecnico.empty:
            messagebox.showwarning(
                "Sin casos", f"{tecnico} no tiene casos pendientes en la ventana."
            )
            return

        if not messagebox.askyesno(
            "Confirmar",
            f"¿Marcar {len(df_tecnico)} caso(s) de {tecnico} como notificados "
            "por WhatsApp?\n\nEsto queda registrado en historial.db y no envia nada.",
        ):
            return

        registrados = self._registrar_casos_pendientes(df_tecnico)
        if registrados:
            self.lbl_status.config(
                text=f"✅ {registrados} caso(s) de {tecnico} marcados como notificados"
            )
            self._refrescar_datos()  # recalcula ultima notificacion y veces hoy
            messagebox.showinfo(
                "Registrado",
                f"{registrados} caso(s) de {tecnico} marcados como notificados "
                "por WhatsApp.",
            )

    def _marcar_todos_notificados(self) -> None:
        """Marca como notificados TODOS los casos pendientes de la ventana."""
        pendientes = self._pendientes
        if pendientes is None or pendientes.empty:
            messagebox.showwarning(
                "Sin datos", "No hay casos pendientes de notificar."
            )
            return

        if not messagebox.askyesno(
            "Confirmar",
            f"¿Marcar los {len(pendientes)} caso(s) pendientes de "
            f"{len(self._resumen_pendientes)} tecnico(s) como notificados por "
            "WhatsApp?\n\nEsto queda registrado en historial.db y no envia nada.",
        ):
            return

        registrados = self._registrar_casos_pendientes(pendientes)
        if registrados:
            self.lbl_status.config(
                text=f"✅ {registrados} caso(s) de todos los tecnicos marcados como notificados"
            )
            self._refrescar_datos()
            messagebox.showinfo(
                "Registrado",
                f"{registrados} caso(s) marcados como notificados por WhatsApp.",
            )

    # ------------------------------------------------------------------
    # Barra inferior
    # ------------------------------------------------------------------

    def _construir_barra_inferior(self) -> None:
        frame = ttk.Frame(self.root, padding=6)
        frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.lbl_status = ttk.Label(
            frame,
            text="Listo",
            relief="sunken",
            anchor=tk.W,
            font=("Segoe UI", 9),
            foreground="#616161",
        )
        self.lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(frame, text="📤 Exportar CSV", command=self._exportar_csv).pack(
            side=tk.RIGHT, padx=4
        )
        ttk.Button(
            frame, text="🔴🟡 Notificar", command=self._notificar
        ).pack(side=tk.RIGHT, padx=4)
        ttk.Button(
            frame, text="🚀 Dashboard", command=self._abrir_dashboard
        ).pack(side=tk.RIGHT, padx=4)

    # ------------------------------------------------------------------
    # Datos
    # ------------------------------------------------------------------

    def _refrescar_datos(self) -> None:
        """
        Recalcula el tablero y actualiza TODAS las pestanas.

        Se llama al arrancar, desde el boton/menu Refrescar y desde el
        auto-refresh. Cualquier fallo de una vista secundaria (metricas,
        historial o pendientes) no debe tumbar la ventana.
        """
        try:
            resultado = calcular_tablero(ARCHIVO_EXCEL)
            self._df = resultado.df
            self._df_completo = resultado.df_completo
            self._conteo_completo = dict(resultado.conteo_completo)
            self._pendientes = resultado.requieren_atencion

            self._actualizar_kpis(self._conteo_completo)   # pestana 1
            self._poblar_filtros()
            self._aplicar_filtros()

            self._metricas = metricas_por_tecnico(self._df_completo)
            self._poblar_metricas()                        # pestana 2

            self._cargar_historial()
            self._poblar_notificaciones()                  # pestana 3

            self._cargar_pendientes()
            self._poblar_pendientes()                      # pestana 4

            self._actualizar_hora()
            self.lbl_status.config(
                text=(
                    f"Ultima actualizacion: {ahora_colombia():%Y-%m-%d %H:%M:%S} | "
                    f"Casos activos: {resultado.total_activos} de {resultado.total_hoja} | "
                    f"En ventana ({resultado.dias_ventana} dias + vencidos): "
                    f"{resultado.total_en_ventana}"
                )
            )
            for aviso in resultado.warnings:
                self.lbl_status.config(text=self.lbl_status.cget("text") + f" | ⚠️ {aviso}")
        except ErrorLecturaExcel as exc:
            self.lbl_status.config(text=f"❌ Error: {exc}")
            messagebox.showerror("Error de lectura", str(exc))
        except Exception as exc:  # blindaje: la GUI nunca debe caerse
            self.lbl_status.config(
                text=f"❌ Error inesperado ({type(exc).__name__}): {exc}"
            )

        self._iniciar_auto_refresh()

    def _iniciar_auto_refresh(self) -> None:
        if self._timer_id:
            self.root.after_cancel(self._timer_id)
        self._timer_id = self.root.after(
            INTERVALO_AUTO_REFRESH_S * 1000, self._on_auto_refresh
        )

    def _on_auto_refresh(self) -> None:
        self._refrescar_datos()

    def _on_refrescar(self) -> None:
        self._refrescar_datos()

    # ------------------------------------------------------------------
    # KPIs
    # ------------------------------------------------------------------

    def _actualizar_kpis(self, conteo: dict[str, int]) -> None:
        for estado in TODOS_LOS_ESTADOS:
            if estado in self.kpi_vars:
                self.kpi_vars[estado].set(str(int(conteo.get(estado, 0))))

    def _actualizar_kpi_seleccion(self, event=None) -> None:
        """
        Callback de <<TreeviewSelect>>.

        Si hay filas seleccionadas en la tabla, los KPI reflejan el conteo por
        estado SOLO de la seleccion. Si no hay seleccion, se restauran los
        conteos de TODOS los casos (activos + cerrados).
        """
        seleccion = self.tree.selection()
        self._seleccion_filas = [self.tree.index(item) for item in seleccion]

        if not seleccion:
            # Sin seleccion: volver a los totales globales.
            self._actualizar_kpis(self._conteo_completo)
            return

        conteo = {estado: 0 for estado in TODOS_LOS_ESTADOS}
        for item in seleccion:
            valores = self.tree.item(item).get("values", ())
            # La columna 8 (indice 7) contiene el estado del caso.
            if len(valores) >= 8:
                estado = str(valores[7])
                if estado in conteo:
                    conteo[estado] += 1

        self._actualizar_kpis(conteo)

    # ------------------------------------------------------------------
    # Filtros
    # ------------------------------------------------------------------

    def _poblar_filtros(self) -> None:
        if self._df is None or self._df.empty:
            return

        tecnicos = sorted(t for t in self._df[COL_TECNICO].dropna().unique() if str(t).strip())
        self.combo_tecnico["values"] = tecnicos

        regiones = sorted(
            r for r in self._df["REGION_TECNICO"].dropna().unique() if str(r).strip() and r != "SIN REGION"
        )
        self.combo_region["values"] = regiones

    def _limpiar_filtros(self) -> None:
        self.combo_tecnico.set("")
        self.combo_region.set("")
        self.entry_buscar.delete(0, tk.END)
        self.combo_estado.set("Todos")
        self._aplicar_filtros()

    def _aplicar_filtros(self) -> None:
        if self._df is None or self._df.empty:
            self._limpiar_tabla()
            return

        filtrado = self._df.copy()

        tecnico = self.combo_tecnico.get().strip()
        if tecnico:
            filtrado = filtrado[filtrado[COL_TECNICO].str.strip() == tecnico]

        region = self.combo_region.get().strip()
        if region:
            filtrado = filtrado[filtrado["REGION_TECNICO"] == region]

        estado = self.combo_estado.get()
        if estado and estado != "Todos":
            filtrado = filtrado[filtrado["ESTADO"] == estado]

        busqueda = self.entry_buscar.get().strip()
        if busqueda:
            mascara = (
                filtrado[COL_CASO].astype(str).str.contains(busqueda, case=False, na=False)
                | filtrado[COL_CIUDAD].astype(str).str.contains(busqueda, case=False, na=False)
                | filtrado[COL_TECNICO].astype(str).str.contains(busqueda, case=False, na=False)
            )
            filtrado = filtrado[mascara]

        self._poblar_tabla(filtrado)

    # ------------------------------------------------------------------
    # Tabla de casos
    # ------------------------------------------------------------------

    def _limpiar_tabla(self) -> None:
        self._filas_tabla.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

    def _poblar_tabla(self, df: pd.DataFrame) -> None:
        self._limpiar_tabla()
        if df.empty:
            return

        for _, fila in df.iterrows():
            estado = str(fila.get("ESTADO", ""))
            icono = ICONO_ESTADO.get(estado, "")
            caso = str(fila.get(COL_CASO, "S/N"))
            tecnico = str(fila.get(COL_TECNICO, ""))
            region = str(fila.get("REGION_TECNICO", ""))
            ciudad = str(fila.get(COL_CIUDAD, ""))
            tiempo_restante = texto_seguro(fila.get("TIEMPO_RESTANTE"))
            horas_restantes = fila.get("HORAS_RESTANTES")
            horas = (
                f"{horas_restantes:.2f}" if pd.notna(horas_restantes) else "—"
            )

            tags = (f"tag_{estado}",) if estado in COLORES_FONDO else ()

            iid = str(len(self._filas_tabla))
            self._filas_tabla[iid] = fila
            self.tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    icono, caso, tecnico, region, ciudad,
                    tiempo_restante, horas, estado,
                ),
                tags=tags,
            )

    # ------------------------------------------------------------------
    # Pestana 2: metricas por tecnico
    # ------------------------------------------------------------------

    def _poblar_metricas(self) -> None:
        for item in self.tree_metricas.get_children():
            self.tree_metricas.delete(item)

        tabla = self._metricas
        if tabla is None or tabla.empty:
            self.lbl_metricas.config(text="Metricas por tecnico: sin datos.")
            return

        for _, fila in tabla.iterrows():
            vencidos = int(fila.get("VENCIDOS", 0) or 0)
            tags = ("tag_vencidos",) if vencidos > 0 else ()
            self.tree_metricas.insert(
                "",
                tk.END,
                values=(
                    texto_seguro(fila.get("TECNICO")),
                    texto_seguro(fila.get("REGION")),
                    int(fila.get("ASIGNADOS", 0) or 0),
                    int(fila.get("ABIERTOS", 0) or 0),
                    vencidos,
                    int(fila.get("PROXIMOS_VENCER", 0) or 0),
                    int(fila.get("NARANJA", 0) or 0),
                    int(fila.get("AMARILLO", 0) or 0),
                    int(fila.get("CERRADOS", 0) or 0),
                    int(fila.get("CERRADOS_TARDE", 0) or 0),
                    formatear_pct(fila.get("CUMPLIMIENTO_PCT")),
                    f"{float(fila.get('HORAS_VENCIDO_TOTAL', 0) or 0):.1f}",
                    f"{float(fila.get('HORAS_VENCIDO_PROMEDIO', 0) or 0):.1f}",
                ),
                tags=tags,
            )

        total_asignados = int(tabla["ASIGNADOS"].sum())
        total_vencidos = int(tabla["VENCIDOS"].sum())
        self.lbl_metricas.config(
            text=(
                f"{len(tabla)} tecnico(s) | casos asignados: {total_asignados} | "
                f"vencidos sin cerrar: {total_vencidos} | "
                "las filas resaltadas tienen casos vencidos"
            )
        )

    # ------------------------------------------------------------------
    # Pestana 3: notificaciones
    # ------------------------------------------------------------------

    def _cargar_historial(self) -> None:
        """Lee historial.db; si falla, deja las tablas vacias y guarda el error."""
        try:
            historial = Historial()
            self._hist_tecnico = historial.por_tecnico()
            self._hist_detalle = historial.leer()
            self._hist_resumen = historial.resumen()
            self._hist_error = ""
        except Exception as exc:
            self._hist_tecnico = pd.DataFrame(columns=COLS_NOTIF_TECNICO)
            self._hist_detalle = pd.DataFrame(columns=COLS_NOTIF_DETALLE)
            self._hist_resumen = {}
            self._hist_error = f"{type(exc).__name__}: {exc}"

    def _cargar_pendientes(self) -> None:
        """
        Prepara los datos de la pestana de pendientes (avisos.py).

        Si algo falla (por ejemplo historial.db bloqueado), la pestana queda
        vacia con su mensaje, sin tumbar la ventana.
        """
        try:
            pendientes = self._pendientes
            if pendientes is None or pendientes.empty:
                self._resumen_pendientes = pd.DataFrame(
                    columns=COLS_RESUMEN_PENDIENTES
                )
                self._grupos_pendientes = {}
            else:
                self._resumen_pendientes = avisos.resumen_pendientes(pendientes)
                self._grupos_pendientes = avisos.agrupar_por_tecnico(pendientes)
            self._menciones_pendientes = avisos.cargar_menciones()
        except Exception as exc:
            self._resumen_pendientes = pd.DataFrame(
                columns=COLS_RESUMEN_PENDIENTES
            )
            self._grupos_pendientes = {}
            self._menciones_pendientes = {}
            self.lbl_status.config(
                text=f"⚠️ No se pudieron preparar las notificaciones pendientes: "
                f"{type(exc).__name__}: {exc}"
            )

    def _poblar_notificaciones(self) -> None:
        # --- Notificaciones por tecnico --------------------------------
        for item in self.tree_notif_tecnico.get_children():
            self.tree_notif_tecnico.delete(item)

        tabla_tec = self._hist_tecnico
        if tabla_tec is not None and not tabla_tec.empty:
            for _, fila in tabla_tec.iterrows():
                self.tree_notif_tecnico.insert(
                    "",
                    tk.END,
                    values=(
                        texto_seguro(fila.get("TECNICO")),
                        texto_seguro(fila.get("REGION"), ""),
                        int(fila.get("NOTIFICACIONES", 0) or 0),
                        int(fila.get("CASOS_DISTINTOS", 0) or 0),
                        texto_seguro(fila.get("ULTIMA_NOTIFICACION")),
                        int(fila.get("ROJOS", 0) or 0),
                        int(fila.get("NARANJAS", 0) or 0),
                        int(fila.get("AMARILLOS", 0) or 0),
                        int(fila.get("FALLIDAS", 0) or 0),
                    ),
                )

        # --- Historial detallado: ultimas MAX_FILAS_NOTIFICACIONES ------
        for item in self.tree_notif_detalle.get_children():
            self.tree_notif_detalle.delete(item)

        tabla_det = self._hist_detalle
        if tabla_det is not None and not tabla_det.empty:
            for _, fila in tabla_det.head(MAX_FILAS_NOTIFICACIONES).iterrows():
                estado = texto_seguro(fila.get("estado"), "")
                icono = ICONO_ESTADO.get(estado, "")
                self.tree_notif_detalle.insert(
                    "",
                    tk.END,
                    values=(
                        texto_seguro(fila.get("id"), ""),
                        texto_seguro(fila.get("fecha_hora")),
                        texto_seguro(fila.get("caso")),
                        texto_seguro(fila.get("tecnico")),
                        texto_seguro(fila.get("region"), ""),
                        f"{icono} {estado}".strip(),
                        texto_seguro(fila.get("horas_restantes")),
                        texto_seguro(fila.get("horas_vencido")),
                        texto_seguro(fila.get("canal"), ""),
                        texto_seguro(fila.get("resultado"), ""),
                        texto_seguro(fila.get("detalle"), ""),
                    ),
                )

        # --- Resumen ---------------------------------------------------
        resumen = self._hist_resumen or {}
        texto = (
            f"Historial: {int(resumen.get('total', 0))} notificacion(es) | "
            f"enviadas: {int(resumen.get('enviadas', 0))} | "
            f"fallidas: {int(resumen.get('fallidas', 0))} | "
            f"casos distintos: {int(resumen.get('casos', 0))} | "
            f"tecnicos: {int(resumen.get('tecnicos', 0))} | "
            f"mostrando {min(len(self._hist_detalle), MAX_FILAS_NOTIFICACIONES)} "
            f"de {len(self._hist_detalle)} fila(s) en el detalle"
        )
        if self._hist_error:
            texto += f" | ⚠️ No se pudo leer el historial: {self._hist_error}"
        self.lbl_notif.config(text=texto)

    # ------------------------------------------------------------------
    # Detalle del caso
    # ------------------------------------------------------------------

    def _linea_detalle(self, etiqueta: str, valor) -> str:
        """Linea del recuadro de detalle, recortada al ancho de la columna."""
        texto = texto_seguro(valor)[:ANCHO_VALOR_DETALLE]
        return (
            "║  "
            + f"{etiqueta:<{ANCHO_ETIQUETA_DETALLE}}"
            + f"{texto:<{ANCHO_VALOR_DETALLE}}"
            + " ║\n"
        )

    def _mostrar_detalle(self, event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return

        item = self.tree.item(sel[0])
        values = item.get("values", ())
        fila = self._filas_tabla.get(sel[0])
        if fila is None and (not values or len(values) < 8):
            return

        if fila is not None:
            icono, caso, tecnico, region, ciudad, tiempo_restante, horas, estado = values[:8]
            tiempo_vencido = texto_seguro(fila.get("TIEMPO_VENCIDO"))
            prediccion = texto_seguro(fila.get("PREDICCION"))
            resolucion = fila.get("FECHA_RESOLUCION")
            resolucion = (
                pd.Timestamp(resolucion).strftime("%Y-%m-%d %H:%M")
                if pd.notna(resolucion) else "—"
            )
            resolutor = texto_seguro(fila.get("TECNICO_RESOLUTOR"), "")
        else:
            icono, caso, tecnico, region, ciudad, tiempo_restante, horas, estado = values[:8]
            tiempo_vencido = prediccion = resolucion = resolutor = "—"

        detalles = (
            "╔" + "═" * ANCHO_INTERIOR_DETALLE + "╗\n"
            + "║" + "  DETALLE DEL CASO".ljust(ANCHO_INTERIOR_DETALLE) + "║\n"
            + "╠" + "═" * ANCHO_INTERIOR_DETALLE + "╣\n"
            + self._linea_detalle("Caso:", caso)
            + self._linea_detalle("Estado:", f"{icono} {estado}")
            + self._linea_detalle("Horas rest.:", horas)
            + self._linea_detalle("Restante:", tiempo_restante)
            + self._linea_detalle("Tiempo vencido:", tiempo_vencido)
            + self._linea_detalle("Prediccion:", prediccion)
            + self._linea_detalle("Tecnico:", tecnico)
            + self._linea_detalle("Resolutor:", resolutor)
            + self._linea_detalle("Region:", region)
            + self._linea_detalle("Ciudad:", ciudad)
            + self._linea_detalle("Resolucion:", resolucion)
            + "╚" + "═" * ANCHO_INTERIOR_DETALLE + "╝\n"
        )

        self.txt_detalle.config(state=tk.NORMAL)
        self.txt_detalle.delete(1.0, tk.END)
        self.txt_detalle.insert(tk.END, detalles)
        self.txt_detalle.config(state=tk.DISABLED)

    # ------------------------------------------------------------------
    # Acciones
    # ------------------------------------------------------------------

    def _notificar(self, solo_rojo: bool = False) -> None:
        if self._df is None or self._df.empty:
            messagebox.showwarning("Sin datos", "No hay datos cargados para notificar.")
            return
        script = os.path.join(DIR_SCRIPT, "alertas_windows.py")
        args = []
        if solo_rojo:
            args.append("--solo-rojo")
        else:
            args.append("--dry-run")
        try:
            subprocess.Popen(
                [sys.executable, script] + args,
                cwd=DIR_SCRIPT,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            tipo = "solo Rojo" if solo_rojo else "Rojo y Amarillo"
            self.lbl_status.config(text="Notificaciones (" + tipo + ") en segundo plano")
            messagebox.showinfo(
                "Notificaciones",
                "Envio de notificaciones (" + tipo + ") en curso."
                + "\n\nSe abrio una ventana separada."
                + "\nLas alertas apareceran en el centro de notificaciones.",
            )
        except Exception as exc:
            messagebox.showerror("Error", "No se pudo iniciar las notificaciones:\n" + str(exc))

    def _abrir_dashboard(self) -> None:
        dashboard_path = os.path.join(DIR_SCRIPT, "dashboard.py")
        try:
            subprocess.Popen(
                [sys.executable, "-m", "streamlit", "run", dashboard_path],
                cwd=DIR_SCRIPT,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            self.lbl_status.config(text="🚀 Dashboard Streamlit iniciado en el navegador")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo iniciar el dashboard:\n{exc}")

    def _exportar_csv(self) -> None:
        if self._df is None or self._df.empty:
            messagebox.showwarning("Sin datos", "No hay datos para exportar.")
            return

        timestamp = ahora_colombia().strftime("%Y%m%d_%H%M")
        ruta_salida = os.path.join(DIR_SCRIPT, f"casos_sla_{timestamp}.csv")
        try:
            self._df.to_csv(ruta_salida, index=False, encoding="utf-8-sig")
            self.lbl_status.config(text=f"📥 Exportado a: {ruta_salida}")
            messagebox.showinfo("Exportar", f"Archivo guardado en:\n{ruta_salida}")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo exportar:\n{exc}")

    def _exportar_metricas_csv(self) -> None:
        """Exporta la tabla de metricas por tecnico (pestana 2)."""
        if self._metricas is None or self._metricas.empty:
            messagebox.showwarning("Sin datos", "No hay metricas por tecnico para exportar.")
            return
        timestamp = ahora_colombia().strftime("%Y%m%d_%H%M")
        ruta_salida = os.path.join(DIR_SCRIPT, f"metricas_por_tecnico_{timestamp}.csv")
        try:
            self._metricas.to_csv(ruta_salida, index=False, encoding="utf-8-sig")
            self.lbl_status.config(text=f"📥 Metricas exportadas a: {ruta_salida}")
            messagebox.showinfo("Exportar", f"Archivo guardado en:\n{ruta_salida}")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo exportar:\n{exc}")

    def _exportar_notificaciones_csv(self) -> None:
        """Exporta el historial detallado de notificaciones (pestana 3)."""
        if self._hist_detalle is None or self._hist_detalle.empty:
            messagebox.showwarning(
                "Sin datos", "No hay notificaciones registradas para exportar."
            )
            return
        timestamp = ahora_colombia().strftime("%Y%m%d_%H%M")
        ruta_salida = os.path.join(DIR_SCRIPT, f"notificaciones_{timestamp}.csv")
        try:
            self._hist_detalle.to_csv(ruta_salida, index=False, encoding="utf-8-sig")
            self.lbl_status.config(text=f"📥 Notificaciones exportadas a: {ruta_salida}")
            messagebox.showinfo("Exportar", f"Archivo guardado en:\n{ruta_salida}")
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo exportar:\n{exc}")

    # ------------------------------------------------------------------
    # Hora y acerca de
    # ------------------------------------------------------------------

    def _actualizar_hora(self) -> None:
        self.lbl_hora.config(text=f"⏰ {ahora_colombia():%Y-%m-%d %H:%M:%S}")

    def _acerca_de(self) -> None:
        messagebox.showinfo(
            "Acerca de",
            "🛰️ Torre de Control SLA\n"
            "Colsof / Banco Agrario\n\n"
            "Sistema de monitoreo de Acuerdos de Nivel de Servicio\n"
            "para soporte técnico en sitio.\n\n"
            f"Versión: 2.0.0\n"
            f"Componentes: GUI con pestañas, Dashboard Streamlit, Notificaciones",
        )


# ------------------------------------------------------------------
# Punto de entrada
# ------------------------------------------------------------------

def main() -> None:
    root = tk.Tk()
    app = TorreControlApp(root)

    def on_close() -> None:
        if app._timer_id:
            root.after_cancel(app._timer_id)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
