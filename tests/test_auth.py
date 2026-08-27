import pytest
import bcrypt
import yaml
from pathlib import Path
from unittest.mock import patch


def _make_users_yaml(tmp_path, users: list) -> Path:
    """Crea un users.yaml temporal con los usuarios dados."""
    data = {"users": []}
    for username, password in users:
        hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        data["users"].append({"username": username, "password_hash": hashed})
    f = tmp_path / "users.yaml"
    f.write_text(yaml.dump(data))
    return f


def test_verify_user_valid(tmp_path):
    from auth import verify_user
    users_file = _make_users_yaml(tmp_path, [("ana", "secreta123")])
    assert verify_user("ana", "secreta123", users_file=str(users_file)) is True


def test_verify_user_wrong_password(tmp_path):
    from auth import verify_user
    users_file = _make_users_yaml(tmp_path, [("ana", "secreta123")])
    assert verify_user("ana", "wrongpass", users_file=str(users_file)) is False


def test_verify_user_unknown_user(tmp_path):
    from auth import verify_user
    users_file = _make_users_yaml(tmp_path, [("ana", "secreta123")])
    assert verify_user("hacker", "secreta123", users_file=str(users_file)) is False


def test_verify_user_empty_file(tmp_path):
    from auth import verify_user
    f = tmp_path / "users.yaml"
    f.write_text("users: []")
    assert verify_user("anyone", "pass", users_file=str(f)) is False
