#!/usr/bin/env python3
"""
CLI para administrar usuarios en users.yaml.

Uso:
    python setup_usuario.py add <username>          # añadir o actualizar usuario
    python setup_usuario.py remove <username>       # eliminar usuario
    python setup_usuario.py list                    # listar usuarios
"""
import sys
import getpass
import bcrypt
import yaml
from pathlib import Path

USERS_FILE = Path(__file__).parent / "users.yaml"


def load_users():
    if not USERS_FILE.exists():
        return []
    data = yaml.safe_load(USERS_FILE.read_text(encoding="utf-8")) or {}
    return data.get("users", [])


def save_users(users):
    USERS_FILE.write_text(
        yaml.dump({"users": users}, allow_unicode=True, default_flow_style=False),
        encoding="utf-8",
    )


def cmd_add(username):
    password = getpass.getpass(f"Contraseña para '{username}': ")
    if len(password) < 8:
        print("Error: la contraseña debe tener al menos 8 caracteres.")
        sys.exit(1)
    confirm = getpass.getpass("Confirmar contraseña: ")
    if password != confirm:
        print("Error: las contraseñas no coinciden.")
        sys.exit(1)

    hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt(rounds=12)).decode()
    users = load_users()
    existing = next((u for u in users if u["username"] == username), None)
    if existing:
        existing["password_hash"] = hashed
        print(f"Contraseña de '{username}' actualizada.")
    else:
        users.append({"username": username, "password_hash": hashed})
        print(f"Usuario '{username}' añadido.")
    save_users(users)


def cmd_remove(username):
    users = load_users()
    new_users = [u for u in users if u["username"] != username]
    if len(new_users) == len(users):
        print(f"Usuario '{username}' no encontrado.")
        sys.exit(1)
    save_users(new_users)
    print(f"Usuario '{username}' eliminado.")


def cmd_list():
    users = load_users()
    if not users:
        print("No hay usuarios registrados.")
        return
    print(f"{'Usuario':<20} {'Hash (primeros 20 chars)'}")
    print("-" * 50)
    for u in users:
        print(f"{u['username']:<20} {u.get('password_hash','')[:20]}...")


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in ("add", "remove", "list"):
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == "list":
        cmd_list()
    elif cmd in ("add", "remove"):
        if len(sys.argv) < 3:
            print(f"Uso: python setup_usuario.py {cmd} <username>")
            sys.exit(1)
        if cmd == "add":
            cmd_add(sys.argv[2])
        else:
            cmd_remove(sys.argv[2])


if __name__ == "__main__":
    main()
