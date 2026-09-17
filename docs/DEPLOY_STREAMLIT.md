# 🚀 Deploy del tablero en Streamlit

**Torre de Control SLA — Colsof / Banco Agrario**

Esta guía despliega **el tablero web** en Streamlit. Es la parte del sistema que
tiene sentido en la web; el resto (avisos automáticos, GUI de escritorio,
notificaciones de Windows) **sigue funcionando en el PC**, porque necesita
procesos en segundo plano y Windows.

---

## 1. Qué se despliega y qué NO

| Componente | ¿En la web? | Por qué |
|---|---|---|
| **`dashboard.py`** | ✅ **Sí** | Portable: solo pandas, openpyxl, requests, streamlit |
| `core.py`, `historial.py`, `avisos.py`, `formatos.py`, `anonimizar.py`, `telegram_notifier.py` | ✅ Sí | Librerías que el tablero necesita |
| `app_gui.py` | ❌ No | Necesita tkinter y una pantalla |
| `alertas_windows.py` | ❌ No | Usa PowerShell y notificaciones de Windows |
| `avisos_automaticos.py` | ❌ No | Necesita ejecución **programada**; la web solo corre cuando alguien abre la página |
| Los `.bat` | ❌ No | Son lanzadores de Windows |

> **Conclusión:** la web sirve para **ver y consultar** el estado de los casos.
> Los **avisos automáticos siguen corriendo en el PC de la torre** con el
> Programador de Tareas. Es una división razonable y muy común.

---

## 2. ⚠️ Lo que debes decidir antes de desplegar

**Streamlit Community Cloud (gratis) exige repositorio PÚBLICO y la app es
PÚBLICA.** Eso significa que cualquiera con la URL vería los datos.

Tus datos incluyen: números de caso, nombres de técnicos, oficinas y los 71 casos
sin asignar del **Banco Agrario**.

**Tienes tres caminos, elige uno:**

| Opción | Datos protegidos | Costo | Esfuerzo |
|---|---|---|---|
| **A. Desplegar con datos anonimizados** (recomendado para probar) | ✅ Se ocultan casos, técnicos y oficinas | Gratis | Bajo |
| **B. Subir el Excel por la interfaz** | ⚠️ El archivo no va al repo, pero quien abra la app ve los datos | Gratis | Bajo |
| **C. Streamlit privado / VM interna** | ✅ Total | Pago o servidor propio | Medio |

El sistema **ya trae las tres** implementadas. Elige con la variable
`TORRE_ANONIMIZAR` (ver sección 5).

---

## 3. Qué contiene la carpeta `DEPLOY_STREAMLIT`

```
DEPLOY_STREAMLIT\
├── dashboard.py              <- la aplicación (entrypoint)
├── core.py                   <- logica de SLA
├── historial.py              <- historial en SQLite
├── avisos.py                 <- avisos por tecnico
├── avisos_automaticos.py
├── telegram_notifier.py
├── formatos.py               <- estilos de presentacion
├── anonimizar.py             <- proteccion de datos
├── requirements.txt          <- dependencias (las instala Streamlit Cloud)
├── plantilla_ejemplo.xlsx    <- datos ANONIMIZADOS, para que arranque
├── .gitignore                <- impide subir datos y credenciales reales
├── .streamlit\
│   ├── config.toml           <- tema y limite de subida (50 MB)
│   └── secrets.toml.example  <- plantilla de secretos
└── docs\DEPLOY_STREAMLIT.md  <- este documento
```

**No incluye** el Excel real, `config_telegram.json`, `menciones.json`, `lib\` ni
`.venv\`.

---

## 4. Pasos para desplegar

### Paso 1 — Probar en local

```powershell
cd DEPLOY_STREAMLIT
python -m streamlit run dashboard.py
```

Se abre en `http://localhost:8501`. Verifica que carga y que puedes subir el Excel.

### Paso 2 — Crear el repositorio en GitHub

```powershell
cd DEPLOY_STREAMLIT
git init -b main
git add .
git status --short          # <-- IMPORTANTE: revisar que NO haya .xlsx reales ni secretos
git commit -m "Torre de Control SLA - tablero Streamlit"
git remote add origin https://github.com/TU_USUARIO/torre-control-sla.git
git push -u origin main
```

> ⚠️ **Antes del `push`, comprueba `git status`.** No debe aparecer ningún
> `.xlsx` real, ni `secrets.toml`, ni `config_telegram.json`.
> El `.gitignore` ya los excluye, pero **verifícalo**.

### Paso 3 — Desplegar en Streamlit Cloud

1. Entra en **https://share.streamlit.io** con tu cuenta de GitHub.
2. **New app** → *Deploy a public app from GitHub*.
3. Rellena:
   - **Repository:** `TU_USUARIO/torre-control-sla`
   - **Branch:** `main`
   - **Main file path:** `dashboard.py`
4. **Advanced settings** → **Python version: 3.12**
5. **Deploy**. Tarda 2–5 minutos.

### Paso 4 — Configurar los secretos

En la app desplegada: **⋮ → Settings → Secrets**, y pega:

```toml
# Proteccion de datos: en la nube SIEMPRE "completo"
TORRE_ANONIMIZAR = "completo"

# Opcional: una plantilla guardada en el repositorio (ya anonimizada)
# TORRE_EXCEL = "plantilla_ejemplo.xlsx"

# Telegram: NO lo configures en un tablero publico
# TELEGRAM_BOT_TOKEN = "..."
# TELEGRAM_CHAT_ID = "..."
```

Guarda. La app se reinicia sola.

---

## 5. Los tres modos de datos

Se controla con `TORRE_ANONIMIZAR`:

| Valor | Qué oculta | Cuándo usarlo |
|---|---|---|
| `"no"` | Nada (datos reales) | Solo local o VM privada |
| `"parcial"` | Números de caso y oficinas | Grupos de trabajo internos |
| `"completo"` | Casos, **técnicos** y oficinas | **Tablero público** |

**Se conserva siempre:** los estados, las fechas, los tiempos y las regiones
(Bogotá, Cali, Medellín…). Por eso **las métricas y los totales siguen siendo
correctos**; solo cambian las etiquetas.

Ejemplo del resultado:

```
ANTES:   QT3329957  | JHONATHAN FELIPE JARABA CHALARCA | 1300-REGIONAL ANTIOQUIA
DESPUES: CASO-0042  | TECNICO-07                       | OFICINA-05
```

> ℹ️ **Detección automática:** si la app corre en Streamlit Cloud (variable
> `SHARING`), se anonimiza en modo `completo` **aunque no configures nada**.
> Es una protección por defecto, no un descuido.

---

## 6. Cómo subir los datos

El tablero acepta **dos orígenes**, en este orden de prioridad:

### A. Subir el archivo por la interfaz (recomendado en la web)

En la **barra lateral** → **📤 Cargar plantilla** → sube el `.xlsx`.

- Se procesa **en memoria**: no se guarda en el servidor ni en el repositorio.
- Debe tener la hoja **`PLANTILLA`**.
- Límite: 50 MB (configurado en `.streamlit/config.toml`).

### B. Plantilla en el repositorio (ya anonimizada)

Si defines `TORRE_EXCEL = "plantilla_ejemplo.xlsx"` en los secretos, la app la
carga sola y no hay que subir nada.

> 💡 La carpeta incluye **`plantilla_ejemplo.xlsx`** con los 483 casos
> **anonimizados**. Sirve para que la app arranque y se pueda evaluar sin exponer
> datos reales.

---

## 7. Limitaciones que debes conocer

### El historial de notificaciones NO persiste

`historial.db` es SQLite en el disco del contenedor. Streamlit Cloud **borra el
disco** en cada reinicio, redespliegue o salida de suspensión.

**Qué significa:** las columnas "última notificación" y "veces hoy" se reinician.
El sistema no se rompe (muestra tablas vacías), pero pierdes el registro.

**Solución si lo necesitas:** usar una base externa (PostgreSQL, Supabase) o una
VM propia con disco persistente. Pídeme el cambio y lo implemento.

### La app se "duerme"

Streamlit Cloud suspende la app tras un rato sin visitas. Al entrar de nuevo tarda
unos segundos en despertar. Es normal.

### Los avisos automáticos siguen en el PC

La web **no envía** los avisos programados. Eso sigue con el Programador de Tareas
en el PC de la torre (`avisos_automaticos.py`).

### El botón de Telegram en un tablero público

Si configuras Telegram en los secretos, **cualquier visitante** podría usar el
botón para enviar mensajes al grupo. Por eso viene **desactivado** y la guía
recomienda dejarlo así.

---

## 8. Actualizar la app

### Actualizar el código

```powershell
# En el PC, tras cambiar algo:
copia los .py modificados a DEPLOY_STREAMLIT\
cd DEPLOY_STREAMLIT
git add .
git commit -m "descripcion del cambio"
git push
```

Streamlit Cloud detecta el `push` y **redespliega solo**.

### Actualizar la plantilla

Si usas el subidor de archivos: **no hay que hacer nada**, cada usuario sube la
suya.

Si usas la plantilla del repositorio: copia el Excel nuevo (anonimizado), `git add`,
`commit` y `push`.

> 🔎 **Recuerda:** si el Excel está en el repositorio, **se puede descargar del
> historial de git para siempre**. Si algún día subiste el archivo real, no basta
> con borrarlo después: hay que reescribir el historial o **crear un repositorio
> nuevo**.

---

## 9. Problemas frecuentes

### "No se encontró la plantilla de seguimiento"

No hay Excel. **Súbelo desde la barra lateral**, o define `TORRE_EXCEL` en los
secretos apuntando a `plantilla_ejemplo.xlsx`.

### "ModuleNotFoundError: No module named 'X'"

Falta una dependencia. Añádela a `requirements.txt`, `git push`, y Streamlit
reinstala.

### "No se pudo procesar el archivo"

Verifica que:
- La hoja se llame exactamente **`PLANTILLA`**.
- Sea un `.xlsx` (no `.xls` ni `.csv`).
- Tenga las columnas de vencimiento, resolución y técnico.

### Los datos se ven como `CASO-0001` / `TECNICO-03`

Es la **anonimización** funcionando. Si es tu VM privada y quieres los datos
reales, pon `TORRE_ANONIMIZAR = "no"` en los secretos.

### Se pierde el historial cada vez que entro

Es la limitación de la sección 7. La solución es una base externa.

---

## 10. Checklist antes de publicar

- [ ] `git status` **no** muestra ningún `.xlsx` real
- [ ] `git status` **no** muestra `secrets.toml` ni `config_telegram.json`
- [ ] `TORRE_ANONIMIZAR = "completo"` en los secretos de Streamlit
- [ ] Telegram **desactivado** en los secretos
- [ ] Python **3.12** seleccionado en Advanced settings
- [ ] Probado en local antes de desplegar
- [ ] La app arranca y muestra el tablero
- [ ] Confirmado que los datos visibles están anonimizados

---

## 11. Resumen

```
1. python preparar_deploy_streamlit.py        <- en el PC
2. cd DEPLOY_STREAMLIT && streamlit run dashboard.py   <- probar
3. git init / add / commit / push              <- repo en GitHub
4. share.streamlit.io -> New app -> dashboard.py
5. Secrets -> TORRE_ANONIMIZAR = "completo"
```

**Lo que ganas:** un tablero central consultable desde cualquier navegador.
**Lo que NO cambia:** los avisos automáticos siguen en el PC de la torre.

---

*Torre de Control SLA — Colsof / Banco Agrario*
