"""
crear_plantilla_plan.py - Prepara el Plan de Trabajo con las columnas que faltan.

Problema que resuelve
---------------------
La hoja ``Casos_Ven`` trae los casos vencidos con su justificacion y su culpa,
pero NO trae estado (abierto/cerrado) ni fecha de cierre. Sin esos dos datos es
imposible medir lo que la operacion necesita para mejorar:

  * tiempo real de cierre por caso y por tecnico,
  * cumplimiento (cerrado dentro del ANS o fuera de el),
  * dias de desviacion respecto al vencimiento,
  * backlog real (vencidos que siguen abiertos) frente al historico.

Este programa agrega esas dos columnas al archivo y deja una lista
desplegable en ESTADO para evitar que se escriba a mano con variantes
("cerrado", "CERRADO", "cerrado ok", ...), que es exactamente el problema que
hoy tienen las hojas diarias.

Uso
---
    python crear_plantilla_plan.py "ruta\\al\\Plan de Trabajo.xlsx"

Genera un archivo NUEVO junto al original, con sufijo ``_con_estado.xlsx``. El
original no se modifica. Si se omite la ruta, se busca un archivo que contenga
"Plan de Trabajo" en la carpeta actual.

Las columnas se insertan vacias a proposito: no se inventan estados ni fechas
de cierre. Mientras esten vacias, el tablero informa lo que si puede calcular
(antiguedad y ANS) y avisa de que la velocidad de cierre aun no es medible.
"""

from __future__ import annotations

import os
import sys

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

HOJA = "Casos_Ven"

COL_ESTADO = "ESTADO"
COL_FECHA_CIERRE = "FECHA DE CIERRE"

# Valores permitidos en ESTADO. Se mantienen los que ya usa la operacion en las
# hojas diarias, normalizados a una sola grafia.
ESTADOS_VALIDOS = [
    "ABIERTO",
    "EN CURSO",
    "EN RUTA",
    "EN FIRMAS",
    "EN VALIDACION",
    "SUSPENDIDO",
    "CERRADO A TIEMPOS",
    "CERRADO TARDE",
    "CANCELADO",
]

COLOR_ENCABEZADO = "1F4E78"
COLOR_ESTADO = "FFF2CC"
COLOR_CIERRE = "DDEBF7"


def localizar_plan(ruta: str | None = None) -> str:
    """Encuentra el archivo del Plan de Trabajo."""
    if ruta:
        if not os.path.isfile(ruta):
            raise SystemExit(f"No existe el archivo: {ruta}")
        return ruta

    candidatos = [
        n for n in os.listdir(".")
        if n.lower().endswith(".xlsx") and "plan de trabajo" in n.lower()
    ]
    if not candidatos:
        raise SystemExit(
            "No se encontro ningun archivo con 'Plan de Trabajo' en el nombre.\n"
            "Pase la ruta explicitamente:\n"
            '    python crear_plantilla_plan.py "C:\\ruta\\Plan de Trabajo.xlsx"'
        )
    if len(candidatos) > 1:
        raise SystemExit(
            "Hay varios archivos candidatos; indique cual:\n  "
            + "\n  ".join(candidatos)
        )
    return candidatos[0]


def _encabezados(ws) -> dict[str, int]:
    """Mapea encabezado normalizado -> numero de columna (1-based)."""
    indice: dict[str, int] = {}
    for celda in ws[1]:
        if celda.value is None:
            continue
        clave = " ".join(str(celda.value).replace("\xa0", " ").split()).upper()
        indice[clave] = celda.column
    return indice


def preparar(ruta_entrada: str, ruta_salida: str) -> dict:
    """Copia el libro agregando ESTADO y FECHA DE CIERRE a la hoja de vencidos."""
    libro = load_workbook(ruta_entrada)

    if HOJA not in libro.sheetnames:
        raise SystemExit(
            f"El archivo no tiene la hoja '{HOJA}'. Hojas encontradas: "
            + ", ".join(libro.sheetnames)
        )

    ws = libro[HOJA]
    encabezados = _encabezados(ws)

    # Ultima columna con encabezado (se ignoran las vacias tipo 'Unnamed: 7').
    ultima = max(encabezados.values()) if encabezados else 1

    # Si la ultima columna con encabezado es la de CULPA y hay una columna sin
    # encabezado justo despues con datos (la categoria real del caso), se le pone
    # nombre para no perderla.
    siguiente = ultima + 1
    tiene_datos_sin_nombre = any(
        ws.cell(row=r, column=siguiente).value is not None
        for r in range(2, min(ws.max_row, 60) + 1)
    )
    if tiene_datos_sin_nombre:
        celda = ws.cell(row=1, column=siguiente)
        celda.value = "CATEGORIA"
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=COLOR_ENCABEZADO)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        ultima = siguiente
        encabezados["CATEGORIA"] = siguiente

    creadas = []
    for nombre, color in ((COL_ESTADO, COLOR_ESTADO), (COL_FECHA_CIERRE, COLOR_CIERRE)):
        if nombre in encabezados:
            continue
        ultima += 1
        celda = ws.cell(row=1, column=ultima)
        celda.value = nombre
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=COLOR_ENCABEZADO)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        # Formato de fecha para que Excel no vuelva a invertir mes y dia.
        if nombre == COL_FECHA_CIERRE:
            for fila in range(2, ws.max_row + 1):
                ws.cell(row=fila, column=ultima).number_format = "DD/MM/YYYY hh:mm"
        ws.column_dimensions[get_column_letter(ultima)].width = 22
        encabezados[nombre] = ultima
        creadas.append(nombre)

    # Lista desplegable en ESTADO: evita las variantes escritas a mano.
    col_estado = encabezados[COL_ESTADO]
    letra = get_column_letter(col_estado)
    validacion = DataValidation(
        type="list",
        formula1='"' + ",".join(ESTADOS_VALIDOS) + '"',
        allow_blank=True,
        showDropDown=False,
    )
    validacion.error = (
        "Use uno de los valores de la lista. Si el caso se cerro despues del "
        "vencimiento, marque 'CERRADO TARDE'."
    )
    validacion.errorTitle = "Estado no reconocido"
    validacion.prompt = "Seleccione el estado del caso"
    validacion.promptTitle = "Estado"
    ws.add_data_validation(validacion)
    validacion.add(f"{letra}2:{letra}{max(ws.max_row, 2)}")

    libro.save(ruta_salida)

    return {
        "hoja": HOJA,
        "filas": ws.max_row - 1,
        "columnas_creadas": creadas,
        "estado_letra": letra,
        "salida": ruta_salida,
    }


def main(argv: list[str]) -> int:
    ruta = argv[1] if len(argv) > 1 else None
    entrada = localizar_plan(ruta)
    base, ext = os.path.splitext(entrada)
    salida = f"{base}_con_estado{ext}"

    if os.path.abspath(entrada) == os.path.abspath(salida):
        raise SystemExit("La entrada y la salida son el mismo archivo.")

    info = preparar(entrada, salida)

    print("=" * 74)
    print("PLAN DE TRABAJO - columnas de seguimiento")
    print("=" * 74)
    print(f"Entrada : {entrada}")
    print(f"Salida  : {info['salida']}")
    print(f"Hoja    : {info['hoja']} ({info['filas']} casos)")
    if info["columnas_creadas"]:
        print(f"Creadas : {', '.join(info['columnas_creadas'])}")
    else:
        print("Creadas : ninguna (ya existian)")
    print()
    print("La columna ESTADO tiene lista desplegable con:")
    for valor in ESTADOS_VALIDOS:
        print(f"   - {valor}")
    print()
    print("Las dos columnas quedan VACIAS a proposito: no se inventan estados")
    print("ni fechas de cierre. Al llenarlas, el tablero podra calcular el")
    print("tiempo de cierre, el cumplimiento del ANS y el backlog real.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
