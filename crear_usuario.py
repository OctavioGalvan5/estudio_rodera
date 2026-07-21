"""
Ejecutar una sola vez para crear el primer usuario:
    cd estudio_rodera
    python crear_usuario.py
"""
from werkzeug.security import generate_password_hash
from sqlalchemy import text
from models.database import engine, init_db

init_db()

username = input("Usuario: ").strip()
password = input("Contraseña: ").strip()
fullname = input("Nombre completo: ").strip()

with engine.begin() as conn:
    conn.execute(
        text("INSERT INTO users (username, password, fullname) VALUES (:u, :p, :f)"),
        {"u": username, "p": generate_password_hash(password), "f": fullname}
    )

print(f"Usuario '{username}' creado correctamente.")
