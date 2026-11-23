import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from legacy_core import app, init_db

# DB pehle init karo — tables banana zaroori hai
init_db()

# Phir modules register karo
from modules import register_modules
register_modules(app)

application = app
