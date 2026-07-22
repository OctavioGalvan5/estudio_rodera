from dotenv import load_dotenv
load_dotenv()

from models.database import engine, _use_postgres
from sqlalchemy import text

print(f"Motor: {'PostgreSQL' if _use_postgres else 'SQLite'}")

try:
    with engine.connect() as conn:
        print("Conexion OK\n")

        users = conn.execute(text("SELECT id, username, fullname FROM users")).fetchall()
        print(f"=== users ({len(users)} filas) ===")
        for row in users:
            print(f"  {row}")

        clientes = conn.execute(text("SELECT id, nombre_completo, numero_cuil FROM data_clientes")).fetchall()
        print(f"\n=== data_clientes ({len(clientes)} filas) ===")
        for row in clientes:
            print(f"  {row}")

except Exception as e:
    print(f"ERROR: {e}")
