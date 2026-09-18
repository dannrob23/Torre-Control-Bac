"""
estilos_css.py - Sistema de diseño visual y CSS para la Torre de Control SLA.

Aporta:
    - Variables CSS con la paleta de colores oficial del semáforo SLA.
    - Estilos para tarjetas elevadas (KPI cards), bordes redondeados y tipografía pulida.
    - Componentes visuales para Badges/Chips de estado accesibles (WCAG AA).
    - Mejoras en la apariencia de tablas, botones, pestañas y popovers de Streamlit.
"""

from __future__ import annotations

import streamlit as st

# Paleta de colores oficial del semáforo SLA (con alto contraste y elegancia)
COLOR_ROJO = "#B91C1C"
COLOR_CERRADO_TARDE = "#8B0000"
COLOR_NARANJA = "#C2410C"
COLOR_AMARILLO = "#A16207"
COLOR_VERDE = "#15803D"
COLOR_SIN_VENCIMIENTO = "#6B7280"
COLOR_CERRADO_OK = "#15803D"

CSS_SISTEMA_DISENO = """
<style>
/* Importación de fuente limpia */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

/* Reglas generales */
html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* Reducir padding superior por defecto de Streamlit */
.block-container {
    padding-top: 1.8rem;
    padding-bottom: 2.5rem;
    max-width: 1400px;
}

/* Tarjetas elevadas de KPI y Foco */
.kpi-card-v2 {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 12px;
    padding: 16px 20px;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.04), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
    transition: transform 0.15s ease, box-shadow 0.15s ease;
    height: 100%;
}

.kpi-card-v2:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08), 0 4px 6px -2px rgba(0, 0, 0, 0.04);
}

.kpi-card-v2 .kpi-label {
    font-size: 11.5px;
    color: #6B7280;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    font-weight: 700;
}

.kpi-card-v2 .kpi-number {
    font-size: 34px;
    font-weight: 800;
    line-height: 1.1;
    margin-top: 4px;
}

.kpi-card-v2 .kpi-help {
    font-size: 12.5px;
    color: #4B5563;
    margin-top: 6px;
    font-weight: 500;
}

/* Badges y Chips de Estado */
.badge-sla {
    display: inline-flex;
    align-items: center;
    padding: 3px 10px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.01em;
}

.badge-rojo { background-color: #FEE2E2; color: #991B1B; border: 1px solid #FCA5A5; }
.badge-cerrado-tarde { background-color: #7F1D1D; color: #FFFFFF; border: 1px solid #991B1B; }
.badge-naranja { background-color: #FFEDD5; color: #C2410C; border: 1px solid #FDBA74; }
.badge-amarillo { background-color: #FEF9C3; color: #854D0E; border: 1px solid #FDE047; }
.badge-verde { background-color: #DCFCE7; color: #166534; border: 1px solid #86EFAC; }
.badge-cerrado-ok { background-color: #DCFCE7; color: #15803D; border: 1px solid #86EFAC; }
.badge-gris { background-color: #F3F4F6; color: #4B5563; border: 1px solid #E5E7EB; }

/* Tarjetas de Fila de Acción */
.fila-accion-card {
    background-color: #FAFAFA;
    border: 1px solid #E5E7EB;
    border-radius: 10px;
    padding: 12px 14px;
    margin-bottom: 8px;
    transition: background-color 0.15s ease;
}

.fila-accion-card:hover {
    background-color: #F3F4F6;
}

/* Estilo para pestañas principales */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    border-bottom: 2px solid #E5E7EB;
    padding-bottom: 4px;
}

.stTabs [data-baseweb="tab"] {
    height: 44px;
    white-space: pre;
    border-radius: 8px 8px 0 0;
    padding: 8px 18px;
    font-weight: 600;
    font-size: 14px;
    color: #4B5563;
}

.stTabs [aria-selected="true"] {
    background-color: #F0FDF4 !important;
    color: #166534 !important;
    border-bottom: 3px solid #15803D !important;
}

/* Contenedores de aviso flotante y popover */
div[data-testid="stPopoverBody"] {
    min-width: 420px !important;
    max-width: 520px !important;
    border-radius: 12px;
    box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
}

/* Botones principales con verde corporativo */
button[kind="primary"] {
    background-color: #0B5D0B !important;
    border-color: #0B5D0B !important;
    font-weight: 600 !important;
}

button[kind="primary"]:hover {
    background-color: #084708 !important;
    border-color: #084708 !important;
}

/* Barra de estado / insignias de timestamp */
.badge-timestamp {
    background: #E0F2FE;
    color: #0369A1;
    font-weight: 600;
    font-size: 12px;
    padding: 4px 10px;
    border-radius: 20px;
    border: 1px solid #BAE6FD;
}
</style>
"""


def inyectar_estilos_css() -> None:
    """Inyecta el CSS personalizado del sistema de diseño en la app Streamlit."""
    st.markdown(CSS_SISTEMA_DISENO, unsafe_allow_html=True)
