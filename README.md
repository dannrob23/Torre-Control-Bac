# 🛰️ Torre de Control SLA — Colsof / Banco Agrario

Sistema local de monitoreo de **Acuerdos de Nivel de Servicio (SLA/ANS)** para el
soporte técnico en sitio de la operación **Colsof – Banco Agrario**.

Tres formas de interactuar con el sistema:

| Componente | Tipo | Cómo usarlo |
|---|---|---|
| **app_gui.py** | 🖥️ Escritorio | `python app_gui.py` — Ventana nativa con KPIs, tabla y filtros |
| **dashboard.py** | 🌐 Web | `streamlit run dashboard.py` — Panel interactivo en el navegador |
| **alertas_windows.py** | 🔔 Background | `pythonw alertas_windows.py` — Notifica casos críticos (Programador de Tareas) |

Todos los componentes comparten la misma lógica de cálculo en `core.py`.

---

## 📋 Objetivo

Monitorear en tiempo real los casos de soporte técnico en sitio cuya fecha
de vencimiento se acerca o ya venció, clasificarlos por urgencia (semaforo)
y notificar automáticamente a los técnicos responsables para que ninguna
incidencia quede sin atención dentro del acuerdo de nivel de servicio.

### ¿Qué hace el sistema?

1. **Lee** la plantilla de seguimiento de casos (`PLANTILLA DE SEGUIMIENTO DE CASOS SEPTIEMBRE.xlsx`)
2. **Filtra** los casos activos (sin fecha de resolución)
3. **Calcula** las horas restantes hasta el vencimiento
4. **Clasifica** cada caso en un semáforo:
   - 🔴 **ROJO** — Vencido o con menos de 1 hora restante (crítico)
   - 🟡 **AMARILLO** — Entre 1 y 4 horas restantes (alerta)
   - 🟢 **VERDE** — Más de 4 horas restantes (a tiempo)
   - ⚪ **SIN VENCIMIENTO** — Sin fecha de vencimiento registrada
5. **Asigna** región a cada técnico mediante cruce de diccionario (27 técnicos → 16 regiones)
6. **Notifica** los casos Rojo y Amarillo vía escritorio (solo en Windows)
7. **Muestra** todo en una interfaz visual interactiva

---

## 🚀 Instalación rápida

### Requisitos

- **Windows 11**
- **Python 3.10+** ([descargar](https://www.python.org/downloads/))
- Excel con la hoja `PLANTILLA` (incluida en el repositorio)

### Paso 1 — Clonar repositorio

```powershell
cd $HOME\Documents\PROYECTOS
git clone https://github.com/dannrob23/Torre-Control-Bac.git
```

### Paso 2 — Instalar dependencias

```powershell
cd Torre-Control-Bac
python -m pip install -r requirements.txt
```

### Paso 3 — Ejecutar

Elige tu forma preferida de interactuar:

```powershell
# Opción A: Interfaz gráfica de escritorio (recomendada)
python app_gui.py

# Opción B: Dashboard web en el navegador
streamlit run dashboard.py

# Opción C: Probar notificaciones
pythonw alertas_windows.py --test
```

---

## 🎯 Tutorial de uso

### Interfaz gráfica de escritorio (`app_gui.py`)

Al iniciar la aplicación verás una ventana con las siguientes secciones:

#### 1. Panel de KPIs (parte superior)

Cuatro tarjetas coloreadas muestran el conteo de casos por estado:

| Tarjeta | Color | Significado |
|---|---|---|
| 🔴 ROJO | Rojo | Casos vencidos o críticos (< 1 hora) |
| 🟡 AMARILLO | Amarillo | Casos en alerta (1-4 horas) |
| 🟢 VERDE | Verde | Casos a tiempo (> 4 horas) |
| ⚪ SIN VENCIMIENTO | Gris | Casos sin fecha de vencimiento |

> Los números se actualizan automáticamente cada 60 segundos.

#### 2. Filtros (barra debajo de los KPIs)

Refina la vista de casos aplicando filtros:

| Filtro | Cómo usarlo |
|---|---|
| **👷 Tecnico** | Selecciona un técnico del desplegable para ver solo sus casos |
| **📍 Region** | Filtra por región (BOGOTA, CALI, MEDELLÍN, etc.) |
| **🔍 Buscar** | Escribe texto para buscar en caso, ciudad o técnico |
| **Estado** | Muestra solo casos de un estado específico |
| **🗑️ Limpiar** | Restablece todos los filtros |

#### 3. Tabla de casos (zona central)

La tabla muestra todos los casos activos con **filas coloreadas** según su estado:

| Columna | Contenido |
|---|---|
| **Caso** | Número de caso (ej: IM3237396) |
| **Tecnico** | Nombre del técnico asignado |
| **Region** | Región asignada al técnico |
| **Ciudad** | Oficina/sucursal |
| **Vencimiento** | Tiempo restante (ej: "Vencido hace 3 h") |
| **Horas restantes** | Valor numérico (negativo = vencido) |
| **Estado** | ROJO / AMARILLO / VERDE |

#### 4. Detalle del caso (parte inferior)

Al hacer clic en una fila de la tabla, se muestra un panel con toda la
información detallada del caso seleccionado.

#### 5. Botones de acción (barra inferior)

| Botón | Función |
|---|---|
| **🔄 Refrescar** | Recalcula los datos desde el Excel |
| **🔴🟡 Notificar** | Abre las alertas de escritorio (Rojo + Amarillo) |
| **🚀 Dashboard** | Abre el dashboard Streamlit en tu navegador |
| **📤 Exportar CSV** | Descarga los casos filtrados como archivo CSV |

---

### Dashboard web (`dashboard.py`)

1. Ejecuta `streamlit run dashboard.py`
2. Se abre automáticamente en tu navegador (puerto 8501)
3. En la barra lateral:
   - **📤 Cargar archivo** — Acepta **la plantilla de seguimiento o el Plan de
     Trabajo**. El sistema reconoce cuál es por sus hojas y lo envía solo a la
     pestaña que corresponde, así que no hay que recordar cuál subir en cada
     casilla. Si el archivo no es ninguno de los dos, lo dice y muestra las
     hojas que encontró.
   - **🔄 Recalcular ahora** — Fuerza el recálculo de datos
   - **Auto-actualizar cada 60 s** — Recarga automática
4. Usa los filtros de la parte superior:
   - **Tecnico** — Selección múltiple
   - **Region** — Selección múltiple
   - **Estado** — Selección múltiple
   - **Buscar caso/ciudad** — Texto libre
5. Desplázate hacia abajo para ver:
   - Gráfico de casos por región
   - Gráfico de carga por técnico (top 15)
   - Botón para descargar CSV

El dashboard tiene seis pestañas: **🎯 Despacho Operativo**, **📋 Explorador de
Casos & SLA**, **📊 Analítica & Técnicos**, **📜 Historial & Auditoría**,
**🗂️ Plan de Trabajo & ANS** y **🔍 Integridad de datos**. La quinta trabaja
sobre el Plan de Trabajo mensual — ver [Plan de Trabajo y ANS](#plan-de-trabajo-y-ans-planpy).

---

### Notificaciones (`alertas_windows.py`)

El script envía notificaciones nativas de Windows **solo para casos críticos**
(Rojo y Amarillo). No notifica casos Verdes ni sin vencimiento.

#### Probar notificaciones

```powershell
# Notificación de prueba (verifica que funcione)
pythonw alertas_windows.py --test

# Simular sin enviar notificaciones (modo seco)
python alertas_windows.py --dry-run

# Solo casos Rojo
python alertas_windows.py --solo-rojo

# Limitar cantidad de notificaciones
python alertas_windows.py --max 5

# Ver resumen sin notificar
python alertas_windows.py --resumen
```

#### Programar en Windows (Programador de Tareas)

Ejecutar en **PowerShell como Administrador**:

```powershell
$py     = (Get-Command pythonw.exe).Source
$carpeta = "C:\Users\darobles\Documents\PROYECTOS\Torre-Control-Bac"
$script = Join-Path $carpeta "alertas_windows.py"

$accion  = New-ScheduledTaskAction -Execute $py -Argument "`"$script`"" -WorkingDirectory $carpeta
$disparo = New-ScheduledTaskTrigger -Once -At (Get-Date).Date.AddMinutes(1) `
             -RepetitionInterval (New-TimeSpan -Minutes 15)
$ajustes = New-ScheduledTaskSettingsSet -StartWhenAvailable `
             -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
             -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName "TorreControlSLA-Colsof" `
   -Action $accion -Trigger $disparo -Settings $ajustes `
   -Description "Alertas SLA Colsof - Banco Agrario"
```

> Se recomienda `pythonw.exe` para que no parpadee la consola.

---

## 🗺️ Mapeo de técnicos y regiones

El sistema asigna región a cada técnico mediante un diccionario de 27 técnicos.
El cruce es **tolerante a variaciones**: aplica `strip()`, colapsa espacios,
elimina tildes y unifica mayúsculas.

### Diccionario oficial

| Técnico | Región | | Técnico | Región |
|---|---|---|---|---|
| JESUS ALFREDO JIMENEZ | BOGOTÁ | | CARLOS MIGUEL JIMENEZ SERRATO | BUCARAMANGA |
| JEAN RAFAEL PIÑEROS GARAVITO | BOGOTÁ | | CARLOS ANDRES BRAVO MARTINEZ | CALI |
| JOHAN SEBASTIAN MORENO VALENCIA | BOGOTÁ | | DIEGO ALEJANDRO HERNANDEZ OROZCO | MANIZALES |
| SANTIAGO ANDRES NOEL JIMENEZ | BOGOTÁ | | VICTOR ALFONSO ULLOA ACOSTA | NEIVA |
| IBETH KARINA ACOSTA CRUZ | BOGOTÁ | | CARLOS ENRIQUE CUCHIVAGUEN FORERO | TUNJA |
| JORGE ALEXANDER GUIO OROZCO | BOGOTÁ | | WILLIAM FERNANDO ALDANA NIÑO | IBAGUÉ |
| DAVID DABINSON CARDENAS AMARIS | BOGOTÁ | | ANDREY TIMOTHY RODRIGUEZ MARTINEZ | VALLEDUPAR |
| BRAYAN DUVAN ALVAREZ VEGA | BOGOTÁ | | JOSE LUIS CONTRERAS VALERA | BARRANQUILLA |
| JAIDER ALEJANDRO BONILLA | BOGOTÁ | | MABEL BOTINA IBARRA | POPAYAN |
| YEISON YAMIR IBARGUEN LEUDO | BOGOTÁ | | CRISTIAN VARGAS RODRIGUEZ | CÚCUTA |
| | | | HAROL ARLEY BETANCUR MARTINEZ | MEDELLÍN |
| | | | JONATHAN FELIPE JARABA CHALARCA | MEDELLÍN |
| | | | CRISTIAN CAMILO TOBON GONZALEZ | MEDELLÍN |
| | | | CESAR AUGUSTO PORRAS MONCAYO | VILLAVICENCIO |
| | | | MIGUEL ANGEL VEGA ARRIETA | MONTERIA |
| | | | DANIEL ANDRES TORRES OVIEDO | SINCELEJO |
| | | | JIMMY FABIEN ERAZO CERÓN | PASTO |

> **Nota:** Si un técnico no está registrado, se marca como `SIN REGION` y se
> reporta en los avisos. No se rompe el sistema.

---

## 🩺 Solución de problemas

| Síntoma | Causa | Solución |
|---|---|---|
| `El archivo ... está abierto o bloqueado` | El Excel está abierto en otra ventana | Cerrar el archivo y volver a intentar |
| `No se encontró la plantilla` | El Excel no está en la carpeta | Mover el archivo o definir variable `TORRE_EXCEL` |
| `streamlit: command not found` | Streamlit no instalado | `python -m pip install -r requirements.txt` |
| Las notificaciones no aparecen | Centro de notificaciones en modo No molestar | Verificar configuración de Windows |
| El dashboard muestra `SIN REGION` | Técnico no registrado en el diccionario | Revisar los avisos en pantalla |

---

## 📁 Estructura del proyecto

```
Torre-Control-Bac/
├── .gitignore              # Archivos excluidos del repositorio
├── README.md               # Este archivo
├── requirements.txt        # Dependencias Python
├── core.py                 # Lógica central (lectura, SLA, regiones)
├── plan.py                 # Analítica del Plan de Trabajo (ANS, envejecimiento)
├── crear_plantilla_plan.py # Agrega ESTADO y FECHA DE CIERRE al Plan de Trabajo
├── alertas_windows.py      # Notificaciones de escritorio
├── app_gui.py              # Interfaz gráfica de escritorio (Tkinter)
├── dashboard.py            # Dashboard web (Streamlit)
├── data/
│   ├── PLANTILLA DE SEGUIMIENTO DE CASOS (EJEMPLO).xlsx
│   └── regiones_etiquetas.json
├── .streamlit/
│   ├── config.toml
│   └── secrets.toml.example
└── docs/
    └── DEPLOY_STREAMLIT.md
```

---

### Plan de Trabajo y ANS (`plan.py`)

La pestaña **🗂️ Plan de Trabajo & ANS** del dashboard analiza el archivo mensual
del Plan de Trabajo (hoja `Casos_Ven`), que es distinto de la plantilla SLA.
Responde otra pregunta: no *qué se vence ahora*, sino *cómo está envejeciendo la
cartera y quién la acumula*.

Sube el archivo desde la barra lateral (**🗂️ Cargar Plan de Trabajo**, o el
cargador principal: se reconoce y enruta solo) y obtendrás:

| Bloque | Qué responde |
|---|---|
| 🔴 ANS | Días vencidos por caso y por técnico, con tramos 1-7 / 8-30 / 31-90 / >90 |
| ⏱️ Envejecimiento | Días abierto desde la creación, con tramos ≤7 / 8-15 / 16-30 / >30 |
| 👷 Por técnico | Volumen **y** gravedad de la cartera de cada uno |
| 🚧 Culpa | Distribución y % de vencimiento evitable (técnico + logístico) |
| ⏱️ Velocidad de cierre | Tiempo de cierre y cumplimiento del ANS (requiere dos columnas) |
| 📋 Calidad de datos | Fechas invertidas, duplicados, filas desalineadas |

#### ⚠️ Las fechas de `Casos_Ven` vienen invertidas

Excel convirtió los textos colombianos `DD/MM/AAAA` a fecha nativa aplicando el
formato `m/d/yy`, así que **cuando el día era ≤ 12 mes y día quedaron
intercambiados** (el 1 de septiembre se guardó como 9 de enero). El problema es
que `09/01/2026` es ambiguo: no se puede saber mirando la celda.

`plan.py` lo resuelve con la **secuencia de IDs de caso**, que es monótona (el ID
máximo de cada hoja diaria crece ~100 por día). Cada fecha se contrasta contra
ese modelo y solo se invierte si así queda más cerca. Medido sobre septiembre
2026: la lectura día/mes da 3.6 días de error medio frente a 75.7 de la lectura
mes/día.

Sin esta corrección la antigüedad promedio sale **~4x inflada** (83 días en lugar
de 22). Si una fecha no tiene referencia, **se conserva tal cual: no se adivina**.

#### ⏱️ Activar la velocidad de cierre

`Casos_Ven` no trae estado ni fecha de cierre, así que la velocidad de cierre no
es medible con el archivo tal como viene. No se deduce de las hojas diarias a
propósito: un caso abierto y uno cerrado desaparecen igual y no son
distinguibles.

Para activarla, ejecuta una vez:

```powershell
python crear_plantilla_plan.py "C:\ruta\Septiembre_2026_Plan de Trabajo.xlsx"
```

Genera una copia `..._con_estado.xlsx` con dos columnas nuevas (el original no se
toca):

- **`ESTADO`** — con lista desplegable (`ABIERTO`, `EN CURSO`, `CERRADO A
  TIEMPOS`, `CERRADO TARDE`, …) para evitar las variantes escritas a mano que hoy
  tienen las hojas diarias.
- **`FECHA DE CIERRE`** — con formato `DD/MM/YYYY` para que Excel no la vuelva a
  invertir.

Quedan **vacías a propósito**. Al llenarlas, el tablero calcula tiempo de cierre,
cumplimiento del ANS, días de desviación y backlog real automáticamente.

#### 🔎 Reconocimiento automático del archivo

El cargador **no exige saber qué archivo es**. Al subir un `.xlsx`, el sistema
mira sus hojas y decide:

| Hojas encontradas | Tipo | Qué hace |
|---|---|---|
| `PLANTILLA` (o `PLANTILLA …`) | Plantilla SLA | Carga el tablero de SLA |
| `Casos_Ven` y/o hojas `N_Mes` | Plan de Trabajo | Lo envía a la pestaña 🗂️ |
| Ninguna de las anteriores | Desconocido | Lo dice y lista las hojas que sí trae |

Esto evita el error clásico de subir el Plan de Trabajo en la casilla de la
plantilla y recibir un `ValueError: Worksheet named 'PLANTILLA' not found` de
openpyxl, que no explica nada. Ahora el archivo se enruta solo y, si no hay
plantilla SLA cargada, las pestañas de SLA avisan en vez de romperse.

---

## 📄 Origen de datos

- **Archivo:** `PLANTILLA DE SEGUIMIENTO DE CASOS SEPTIEMBRE.xlsx`
- **Hoja:** `PLANTILLA`
- El archivo se autodescubre por patrón (`*SEGUIMIENTO*CASOS*.xlsx`) o se
  indica con `--excel "ruta\archivo.xlsx"` o la variable de entorno `TORRE_EXCEL`.
- **Plan de Trabajo (ANEXO):** `*Plan de Trabajo*.xlsx`, hoja `Casos_Ven`. Se sube
  desde la pestaña 🗂️ y usa `plan.py`; es un módulo independiente del núcleo SLA.

---

*Torre de Control SLA — Colsof / Banco Agrario*
