"""
configurar_acceso.py - Crea o cambia el usuario y la contrasena del tablero.

COMO SE USA (abrir una terminal en la carpeta del proyecto):

    .venv\\Scripts\\python.exe configurar_acceso.py

POR QUE EXISTE ESTE SCRIPT:
    La contrasena se escribe en SU teclado y NUNCA aparece en pantalla ni queda
    guardada en ningun chat, historial o archivo de texto. Lo unico que se
    escribe en el disco es el HASH bcrypt, del cual no se puede recuperar la
    contrasena original.

QUE ESCRIBE:
    .streamlit/secrets.toml  ->  credenciales + clave de la cookie de sesion.
    Ese archivo esta en .gitignore: nunca se sube al repositorio.

Si lo ejecuta otra vez con un usuario que ya existe, le pedira la contrasena
nueva y la reemplazara. Si usa un usuario nuevo, lo AGREGA (no borra los demas).
"""

from __future__ import annotations

import getpass
import os
import secrets
import sys
from pathlib import Path

try:
    import toml
except ImportError:
    print("Falta la libreria 'toml'. Instalela con:")
    print("    .venv\\Scripts\\python.exe -m pip install toml")
    sys.exit(1)

try:
    from streamlit_authenticator.utilities.hasher import Hasher
except ImportError:
    try:
        from streamlit_authenticator import Hasher
    except ImportError:
        print("Falta la libreria 'streamlit-authenticator'. Instalela con:")
        print(
            "    .venv\\Scripts\\python.exe -m pip install streamlit-authenticator"
        )
        sys.exit(1)


RAIZ = Path(__file__).resolve().parent
RUTA_SECRETS = RAIZ / ".streamlit" / "secrets.toml"

LARGO_MINIMO = 10
LARGO_RECOMPENDADO = 14

# Contrasenas tan obvias que no vale la pena dejarlas pasar.
DEBILES = {
    "1234567890", "12345678910", "contrasena", "contraseña", "password",
    "colsof", "bancoagrario", "qwertyuiop", "123456789012", "administrador",
}

ENCABEZADO = """\
# ---------------------------------------------------------------------------
# ARCHIVO GENERADO por configurar_acceso.py  (NO editar a mano)
#
# Contiene las credenciales del tablero. Este archivo esta en .gitignore:
# NUNCA se sube al repositorio.
#
# Las contrasenas estan guardadas como hash bcrypt. No se pueden leer.
# Para cambiar el usuario o la contrasena, ejecute otra vez:
#     .venv\\Scripts\\python.exe configurar_acceso.py
# ---------------------------------------------------------------------------
"""


def pedir(mensaje: str, por_defecto: str = "") -> str:
    sufijo = f" [{por_defecto}]" if por_defecto else ""
    valor = input(f"{mensaje}{sufijo}: ").strip()
    return valor or por_defecto


def pedir_usuario() -> str:
    while True:
        # NO se pasa a minusculas a proposito: el usuario se guarda tal como se
        # escribe (ej. "Controller"). El tablero acepta mayusculas o minusculas
        # al iniciar sesion, pero se conserva el original para mostrarlo.
        usuario = pedir("Usuario (sin espacios, ej. Controller)")
        if not usuario:
            print("  ⚠️  El usuario no puede estar vacío.")
            continue
        if " " in usuario:
            print("  ⚠️  El usuario no puede llevar espacios.")
            continue
        return usuario


def pedir_contrasena() -> str:
    while True:
        primera = getpass.getpass("Contraseña (no se muestra al escribir): ")
        if len(primera) < LARGO_MINIMO:
            print(f"  ⚠️  Muy corta: use al menos {LARGO_MINIMO} caracteres.")
            continue
        if primera.lower() in DEBILES:
            print("  ⚠️  Esa contraseña es demasiado obvia. Elija otra.")
            continue
        segunda = getpass.getpass("Repita la contraseña: ")
        if primera != segunda:
            print("  ⚠️  No coinciden. Intente de nuevo.")
            continue
        if len(primera) < LARGO_RECOMPENDADO:
            print(
                f"  ℹ️  Aviso: menos de {LARGO_RECOMPENDADO} caracteres. "
                "Lo ideal es una frase larga, por ejemplo "
                "'Torre-SLA-Colsof-2026!'."
            )
        return primera


def cargar_secrets() -> dict:
    if RUTA_SECRETS.exists():
        try:
            return toml.load(RUTA_SECRETS)
        except Exception as exc:
            print(f"⚠️  No se pudo leer {RUTA_SECRETS}: {exc}")
            print("    Se empezará de cero (los usuarios anteriores se pierden).")
            if pedir("¿Continuar? (s/N)", "n").lower() != "s":
                sys.exit(0)
            return {}
    return {}


def main() -> int:
    print("=" * 74)
    print(" CONFIGURAR ACCESO — Torre de Control SLA (Colsof / Banco Agrario)")
    print("=" * 74)
    print()
    print("La contraseña se escribe aquí y no se muestra. Solo se guarda el hash.")
    print()

    datos = cargar_secrets()

    # --- Cookie de sesión --------------------------------------------------
    cookie = datos.setdefault("cookie", {})
    cookie.setdefault("name", "torre_control_sla")
    cookie.setdefault("expiry_days", 7)
    if not str(cookie.get("key", "")).strip():
        # Clave aleatoria de 64 caracteres: firma la cookie de sesión.
        cookie["key"] = secrets.token_hex(32)
        print("🔑 Se generó una clave nueva para la cookie de sesión.")
        print()

    # --- Credenciales ------------------------------------------------------
    credenciales = datos.setdefault("credenciales", {})
    usuarios = credenciales.setdefault("usernames", {})

    if usuarios:
        print(f"Usuarios ya definidos: {', '.join(sorted(usuarios))}")
        print()

    usuario = pedir_usuario()
    nombre = pedir("Nombre para mostrar", usuarios.get(usuario, {}).get("name", ""))
    correo = pedir("Correo (opcional)", usuarios.get(usuario, {}).get("email", ""))

    if usuario in usuarios:
        print(f"\nEl usuario '{usuario}' ya existe: se le cambiará la contraseña.")
    print()

    contrasena = pedir_contrasena()

    usuarios[usuario] = {
        "name": nombre or usuario,
        "email": correo or "",
        "password": Hasher.hash(contrasena),
        "roles": ["admin"],
    }

    # --- Escritura ---------------------------------------------------------
    RUTA_SECRETS.parent.mkdir(parents=True, exist_ok=True)
    with open(RUTA_SECRETS, "w", encoding="utf-8") as fh:
        fh.write(ENCABEZADO)
        fh.write(toml.dumps(datos))

    print()
    print("=" * 74)
    print(" ✅ LISTO")
    print("=" * 74)
    print(f" Archivo : {RUTA_SECRETS}")
    print(f" Usuario : {usuario}")
    print(f" Usuarios configurados: {', '.join(sorted(usuarios))}")
    print()
    print(" La contraseña NO quedó guardada en ningún lado legible.")
    print(" Recargue el tablero (F5) e ingrese con ese usuario y contraseña.")
    print()

    # Aviso si el archivo quedara expuesto por git.
    try:
        import subprocess

        resultado = subprocess.run(
            ["git", "check-ignore", "-q", str(RUTA_SECRETS)],
            cwd=RAIZ,
            capture_output=True,
        )
        if resultado.returncode != 0:
            print(" ⚠️  ATENCIÓN: git NO está ignorando este archivo.")
            print("     Verifique que .gitignore contenga:")
            print("         .streamlit/secrets.toml")
    except Exception:
        pass

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nCancelado. No se cambió nada.")
        sys.exit(1)
