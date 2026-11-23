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
        resets = [
            ("md.admin",       "Admin@2026"),
            ("gm.smartchoice", "ABC@GM2026#"),
            ("gm.whitewave",   "XYZ@GM2026#"),
            ("reception.scic", "Welcome@123"),
            ("reception.wwic", "Welcome@123"),
        ]
        for login_id, pw in resets:
            con.execute(
                "UPDATE users SET password_hash=? WHERE lower(login_id)=lower(?)",
                (generate_password_hash(pw), login_id)
            )
        con.commit()
        # Debug log
        rows = con.execute("SELECT login_id, active FROM users").fetchall()
        print(f"[WSGI] Users in DB: {[dict(r) for r in rows]}")
        con.close()
    except Exception as e:
        print(f"Password reset warning: {e}")

_reset_passwords()

# Phir modules register karo
from modules import register_modules
register_modules(app)

application = app
