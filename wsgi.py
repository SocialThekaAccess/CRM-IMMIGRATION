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

        # Role ke basis pe login ID, password, full_name sab reset karo
        role_updates = [
            ("MD",         None,   "md.admin",       "Admin@2026",   "Managing Director"),
            ("GM",         "SCIC", "gm.abc",          "ABC@GM2026#",  "General Manager"),
            ("GM",         "WWIC", "gm.xyz",          "XYZ@GM2026#",  "General Manager"),
            ("RECEPTION",  "SCIC", "reception.abc",   "Welcome@123",  "Receptionist"),
            ("RECEPTION",  "WWIC", "reception.xyz",   "Welcome@123",  "Receptionist"),
        ]

        for role, company, new_login_id, pw, name in role_updates:
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
                    "UPDATE users SET login_id=?, password_hash=?, full_name=?, active=1 WHERE id=?",
                    (new_login_id, generate_password_hash(pw), name, user["id"])
                )

        # Companies rename karo aur logos remove karo
        con.execute("UPDATE companies SET name=? WHERE code=?", ("ABC Immigration Consultants", "SCIC"))
        con.execute("UPDATE companies SET name=? WHERE code=?", ("XYZ Immigration Consultants", "WWIC"))

        # Kisi bhi table mein logo/photo data clear karo
        try:
            con.execute("UPDATE companies SET logo_data=NULL, logo_mime='' WHERE 1=1")
        except Exception:
            pass
        try:
            con.execute("UPDATE users SET photo_data=NULL, photo_mime='' WHERE 1=1")
        except Exception:
            pass

        # Extra MD users delete karo
        md_users = con.execute("SELECT id FROM users WHERE role='MD' ORDER BY id").fetchall()
        if len(md_users) > 1:
            for u in md_users[1:]:
                con.execute("DELETE FROM users WHERE id=?", (u["id"],))

        con.commit()
        rows = con.execute("SELECT login_id, full_name, role, active FROM users").fetchall()
        print(f"[WSGI] Users after reset: {[dict(r) for r in rows]}")
        con.close()
    except Exception as e:
        print(f"Password reset warning: {e}")

_reset_passwords()

# Phir modules register karo
from modules import register_modules
register_modules(app)

application = app
