"""
Autenticación HTTP Basic Auth contra users.yaml con contraseñas bcrypt.
"""
from pathlib import Path
import bcrypt
import yaml

DEFAULT_USERS_FILE = str(Path(__file__).parent / "data" / "users.yaml")


def verify_user(username: str, password: str, users_file: str = DEFAULT_USERS_FILE) -> bool:
    """
    Verifica que username/password coinciden con una entrada en users_file.

    Args:
        username:   nombre de usuario
        password:   contraseña en texto plano
        users_file: ruta al archivo YAML de usuarios (default: users.yaml junto a auth.py)

    Returns:
        True si las credenciales son válidas, False en cualquier otro caso.
    """
    try:
        data = yaml.safe_load(Path(users_file).read_text(encoding="utf-8"))
        users = data.get("users", [])
    except (FileNotFoundError, yaml.YAMLError):
        return False

    for user in users:
        if user.get("username") == username:
            stored_hash = user.get("password_hash", "").encode()
            try:
                return bcrypt.checkpw(password.encode(), stored_hash)
            except Exception:
                return False
    return False
