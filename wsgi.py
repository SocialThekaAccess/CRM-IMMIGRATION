import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from legacy_core import app, init_db, db
from werkzeug.security import generate_password_hash

# DB pehle init karo
init_db()

# Render pe fresh password set karo
def _reset_passwords():
    try:
        con = db()

        # Saare existing users ke passwords reset karo aur login IDs standardize karo
        # Role ke basis pe match karo
        role_updates = [
            # (role, company_code, new_login_id, new_password)
            ("MD",         None,   "md.admin",       "Admin@2026"),
            ("GM",         "SCIC", "gm.abc",          "ABC@GM2026#"),
            ("GM",         "WWIC", "gm.xyz",          "XYZ@GM2026#"),
            ("RECEPTION",  "SCIC", "reception.abc",   "Welcome@123"),
            ("RECEPTION",  "WWIC", "reception.xyz",   "Welcome@123"),
        ]

        for role, company, new_login_id, pw in role_updates:
            if company is None:
                user = con.execute(
                    "SELECT id FROM users WHERE role=? ORDER BY id LIMIT 1", (role,)
                ).fetchone()
            else:
                user = con.execute(
                    "SELECT id FROM users WHERE role=? AND company_code=? ORDER BY id LIMIT 1",
                    (role, company)
                ).fetchone()

            if user:
                con.execute(
                    "UPDATE users SET login_id=?, password_hash=?, active=1 WHERE id=?",
                    (new_login_id, generate_password_hash(pw), user["id"])
                )

        con.commit()
        rows = con.execute("SELECT login_id, role, active FROM users").fetchall()
        print(f"[WSGI] Users after reset: {[dict(r) for r in rows]}")
        con.close()
    except Exception as e:
        print(f"Password reset warning: {e}")

_reset_passwords()

# Phir modules register karo
from modules import register_modules
register_modules(app)

application = app
