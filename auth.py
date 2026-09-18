"""
auth.py - Puerta de entrada con usuario y contrasena para la Torre de Control SLA.

PARA QUE SIRVE:
    El tablero muestra datos reales del banco (numeros de caso, tecnicos,
    oficinas). En vez de anonimizarlos, se protege el ACCESO: nadie ve nada
    hasta iniciar sesion con usuario y contrasena.

COMO SE GUARDAN LAS CREDENCIALES:
    En .streamlit/secrets.toml, que esta en .gitignore (nunca se sube al repo).
    Las contrasenas se guardan SOLO como hash bcrypt, jamas en texto plano.

    Para crearlas o cambiarlas, ejecute:
        .venv\\Scripts\\python.exe configurar_acceso.py

SEGURIDAD (decisiones tomadas a proposito):
    - FALLA CERRADO: si no hay credenciales configuradas, la app NO se muestra.
      Nunca "suelta" el tablero por falta de configuracion.
    - La cookie de sesion usa una clave aleatoria guardada en secrets.toml.
    - Despues de iniciar sesion hay boton para cerrarla.
"""

from __future__ import annotations

import os

import streamlit as st

try:
    import streamlit_authenticator as stauth
except ImportError as exc:  # pragma: no cover
    raise ImportError(
        "Falta streamlit-authenticator. Instalelo con:\n"
        "    .venv\\Scripts\\python.exe -m pip install streamlit-authenticator"
    ) from exc

# --------------------------------------------------------------------------
# Claves de st.secrets que se copian a os.environ.
#
# IMPORTANTE (arreglo de un fallo del codigo original): el resto del sistema
# (anonimizar.py, core.localizar_excel, telegram_notifier.py) lee estas claves
# de las VARIABLES DE ENTORNO, no de st.secrets. La guia DEPLOY_STREAMLIT.md
# decia configurarlas en los secretos de Streamlit, pero nunca llegaban al
# codigo. Este puente hace que funcionen por las dos vias.
# --------------------------------------------------------------------------
CLAVES_PUENTE = (
    "TORRE_ANONIMIZAR",
    "TORRE_EXCEL",
    "TELEGRAM_BOT_TOKEN",
    "TELEGRAM_CHAT_ID",
)

SECCION_CREDENCIALES = "credenciales"
SECCION_COOKIE = "cookie"

NOMBRE_COOKIE_POR_DEFECTO = "torre_control_sla"
DIAS_COOKIE_POR_DEFECTO = 7.0


# --------------------------------------------------------------------------
# Lectura de secretos
# --------------------------------------------------------------------------

def _leer_secreto(clave: str):
    """Devuelve st.secrets[clave] o None si no existe / no hay secrets.toml."""
    try:
        return st.secrets[clave]
    except Exception:
        return None


def _a_diccionario(seccion) -> dict | None:
    """
    Convierte una seccion de st.secrets (AttrDict) en un dict normal.

    streamlit-authenticator valida con isinstance(x, dict); un AttrDict no
    siempre pasa, asi que se normaliza siempre.
    """
    if seccion is None:
        return None
    if hasattr(seccion, "to_dict"):
        try:
            return seccion.to_dict()
        except Exception:
            pass
    if isinstance(seccion, dict):
        salida: dict = {}
        for clave, valor in seccion.items():
            if hasattr(valor, "to_dict"):
                try:
                    salida[clave] = valor.to_dict()
                    continue
                except Exception:
                    pass
            salida[clave] = valor
        return salida
    return None


def puente_secretos() -> list[str]:
    """
    Copia a os.environ los secretos que el resto del sistema lee del entorno.

    Solo copia si la variable de entorno NO esta ya definida, para que el
    entorno real del PC siga teniendo la ultima palabra.
    Devuelve la lista de claves copiadas (util para diagnosticar).
    """
    copiadas: list[str] = []
    for clave in CLAVES_PUENTE:
        if os.environ.get(clave):
            continue
        valor = _leer_secreto(clave)
        if valor is None or str(valor).strip() == "":
            continue
        os.environ[clave] = str(valor)
        copiadas.append(clave)
    return copiadas


def hay_credenciales() -> bool:
    """True si secrets.toml tiene al menos un usuario y la clave de la cookie."""
    credenciales = _a_diccionario(_leer_secreto(SECCION_CREDENCIALES)) or {}
    usuarios = credenciales.get("usernames") or {}
    cookie = _a_diccionario(_leer_secreto(SECCION_COOKIE)) or {}
    return bool(usuarios) and bool(str(cookie.get("key", "")).strip())


# --------------------------------------------------------------------------
# Pantallas
# --------------------------------------------------------------------------

def _pantalla_sin_configurar() -> None:
    """Se muestra cuando todavia no hay usuario y contrasena creados."""
    st.title("🔐 Torre de Control SLA")
    st.error(
        "**El tablero está bloqueado: todavía no hay un usuario configurado.**",
        icon="🔒",
    )
    st.markdown(
        "Los datos están protegidos y **no se muestran a nadie** hasta que exista "
        "al menos un usuario. Esto es intencional: si no hay credenciales, "
        "el sistema falla cerrado en vez de dejar los datos a la vista."
    )
    st.divider()
    st.subheader("⚙️ Qué hacer (una sola vez)")
    st.markdown(
        "Abra una terminal **en la carpeta del proyecto** y ejecute:\n\n"
        "```bash\n"
        ".venv\\Scripts\\python.exe configurar_acceso.py\n"
        "```\n\n"
        "El programa le pedirá el usuario y la contraseña. **La contraseña la "
        "escribe usted y no se muestra en pantalla**; solo se guarda su hash "
        "(bcrypt) en `.streamlit/secrets.toml`, un archivo que jamás se sube al "
        "repositorio."
    )
    st.info(
        "Cuando termine, recargue esta página (F5) y aparecerá el formulario "
        "de ingreso.",
        icon="ℹ️",
    )
    with st.expander("🔎 Ver detalle técnico del bloqueo"):
        credenciales = _a_diccionario(_leer_secreto(SECCION_CREDENCIALES)) or {}
        usuarios = credenciales.get("usernames") or {}
        cookie = _a_diccionario(_leer_secreto(SECCION_COOKIE)) or {}
        st.write(
            {
                "secrets.toml encontrado": bool(
                    _leer_secreto(SECCION_CREDENCIALES)
                    or _leer_secreto(SECCION_COOKIE)
                ),
                "usuarios definidos": len(usuarios),
                "clave de cookie definida": bool(str(cookie.get("key", "")).strip()),
            }
        )


def _permitir_mayusculas(authenticator) -> bool:
    """
    Hace que el usuario NO distinga mayusculas de minusculas al iniciar sesion.

    Por que: streamlit-authenticator compara el usuario de forma EXACTA
    ('Controller' != 'controller'), y eso es una fuente clasica de "no me deja
    entrar". Aqui se envuelve el metodo check_credentials del modelo para
    buscar el usuario ignorando mayusculas, conservando intacta la verificacion
    de la contrasena (bcrypt) y el registro de intentos fallidos.

    Devuelve True si el ajuste se pudo aplicar. Si una version futura de la
    libreria cambia la estructura interna, simplemente devuelve False y todo
    sigue funcionando como antes (se tendra que escribir el usuario exacto).
    """
    modelo = getattr(authenticator, "authentication_model", None)
    original = getattr(modelo, "check_credentials", None)
    if modelo is None or not callable(original):
        return False

    def check_credentials(usuario, contrasena):
        usuarios = getattr(modelo, "credentials", {}).get("usernames", {}) or {}
        if usuario not in usuarios:
            buscado = str(usuario or "").strip().lower()
            for existente in usuarios:
                if str(existente).strip().lower() == buscado:
                    usuario = existente
                    break
        return original(usuario, contrasena)

    try:
        modelo.check_credentials = check_credentials
    except Exception:
        return False
    return True


def _motor():
    """Construye el Authenticate de streamlit-authenticator o None si falta config."""
    credenciales = _a_diccionario(_leer_secreto(SECCION_CREDENCIALES))
    cookie = _a_diccionario(_leer_secreto(SECCION_COOKIE)) or {}
    if not credenciales or not (credenciales.get("usernames") or {}):
        return None
    clave_cookie = str(cookie.get("key", "")).strip()
    if not clave_cookie:
        return None
    authenticator = stauth.Authenticate(
        credenciales,
        str(cookie.get("name", NOMBRE_COOKIE_POR_DEFECTO)),
        clave_cookie,
        float(cookie.get("expiry_days", DIAS_COOKIE_POR_DEFECTO)),
    )
    _permitir_mayusculas(authenticator)
    return authenticator


def exigir_login() -> dict:
    """
    Muestra el formulario de ingreso y DETIENE la app hasta que haya sesión.

    Devuelve {'nombre': str, 'usuario': str} cuando ya está autenticado.
    Si no hay sesión, no devuelve nada: corta la ejecución con st.stop(),
    de modo que el Excel real nunca llega a leerse ni a renderizarse.
    """
    authenticator = _motor()
    if authenticator is None:
        _pantalla_sin_configurar()
        st.stop()

    resultado = authenticator.login(
        location="main",
        fields={
            "Form name": "🔐 Acceso — Torre de Control SLA",
            "Username": "Usuario",
            "Password": "Contraseña",
            "Login": "Ingresar",
        },
        key="login_torre_control",
    )

    # Compatibilidad: si la libreria no devuelve tupla, se lee el estado de sesion.
    if resultado is None:
        nombre = st.session_state.get("name")
        estado = st.session_state.get("authentication_status")
        usuario = st.session_state.get("username")
    else:
        nombre, estado, usuario = resultado

    if estado is False:
        st.error("Usuario o contraseña incorrectos.", icon="⛔")
        st.caption(
            "Si olvidó la contraseña, vuelva a ejecutar `configurar_acceso.py` "
            "para definir una nueva."
        )
        st.stop()

    if not estado:
        # Todavia no ha enviado el formulario: ya se ve el formulario arriba.
        st.caption(
            "🔒 Los datos del banco no se cargan hasta iniciar sesión."
        )
        st.stop()

    # --- Sesion iniciada ---------------------------------------------------
    with st.sidebar:
        st.success(f"👤 {nombre or usuario}", icon="✅")
        authenticator.logout(
            button_name="🚪 Cerrar sesión",
            location="sidebar",
            use_container_width=True,
            key="logout_torre_control",
        )

    return {"nombre": nombre or usuario, "usuario": usuario}
