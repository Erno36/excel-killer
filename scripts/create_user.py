from __future__ import annotations

import getpass
import os
import sqlite3
import sys
from pathlib import Path

DESKTOP_APP = Path(__file__).resolve().parents[2] / "cruzador_excel"
if str(DESKTOP_APP) not in sys.path:
    sys.path.insert(0, str(DESKTOP_APP))

from auth.service import UserManager
from data.db import init_db


def main() -> int:
    if len(sys.argv) != 2:
        print("Uso: python scripts/create_user.py <usuario>", file=sys.stderr)
        return 2
    username = sys.argv[1].strip()
    if not username:
        print("El usuario no puede estar vacio.", file=sys.stderr)
        return 2

    password = getpass.getpass("Contrasena nueva: ")
    confirmation = getpass.getpass("Repetir contrasena: ")
    if not password or password != confirmation:
        print("Las contrasenas estan vacias o no coinciden.", file=sys.stderr)
        return 2

    default_data_dir = Path(
        os.getenv("XDG_STATE_HOME", Path.home() / ".local" / "state")
    ) / "banco-piano" / "cruzador-excel-web"
    db_path = Path(os.getenv("CRUZADOR_WEB_DB_PATH", default_data_dir / "cruzador.db"))
    connection = init_db(db_path, is_development=False)
    try:
        user = UserManager(connection).create_user(username, password)
    except sqlite3.IntegrityError:
        print("El usuario ya existe o no cumple las restricciones de la base.", file=sys.stderr)
        return 1
    except (OSError, sqlite3.Error) as exc:
        print(f"No se pudo crear el usuario ({type(exc).__name__}).", file=sys.stderr)
        return 1
    finally:
        connection.close()
    print(f"Usuario '{user.username}' creado.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
