
from flask import Flask, request, redirect, url_for, session, render_template, flash, jsonify, send_file
import sqlite3, os, uuid, datetime, re, io, mimetypes, csv, json
from urllib.parse import quote
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
import xlrd
import zipfile
import xml.etree.ElementTree as ET

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
IS_POSTGRES = DATABASE_URL.startswith("postgres://") or DATABASE_URL.startswith("postgresql://")
DB_PATH = os.path.join(APP_DIR, "mini_crm.db")
UPLOAD_DIR = os.path.join(APP_DIR, "uploads")
EMPLOYEE_UPLOAD_DIR = os.path.join(UPLOAD_DIR, "employees")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EMPLOYEE_UPLOAD_DIR, exist_ok=True)

ALLOWED_DOC_EXTENSIONS = {"pdf","jpg","jpeg","png","webp","doc","docx"}
ALLOWED_PHOTO_EXTENSIONS = {"jpg","jpeg","png","webp"}

def allowed_file(filename, allowed):
    return "." in filename and filename.rsplit(".",1)[1].lower() in allowed

app = Flask(__name__)
app.secret_key = os.environ.get("GAUR_CRM_SECRET", "GAUR-CRM-LOCAL-2026-CHANGE-ME")
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "20")) * 1024 * 1024



# === v3.98 Mobile PWA ===
@app.route("/manifest.webmanifest")
def pwa_manifest():
    manifest={
        "name":"THE IMMIGRATION MANAGEMENT SYSTEM",
        "short_name":"THE IMMIGRATION MANAGEMENT SYSTEM",
        "description":"Secure mobile CRM for ABC and XYZ teams.",
        "start_url":"/",
        "scope":"/",
        "display":"standalone",
        "background_color":"#061629",
        "theme_color":"#061629",
        "orientation":"any",
        "icons":[
            {"src":"/static/pwa/icon-192.png","sizes":"192x192","type":"image/png","purpose":"any maskable"},
            {"src":"/static/pwa/icon-512.png","sizes":"512x512","type":"image/png","purpose":"any maskable"}
        ]
    }
    return app.response_class(json.dumps(manifest),mimetype="application/manifest+json")

@app.route("/service-worker.js")
def pwa_service_worker():
    js="""const CACHE='the-gaur-pwa-v346';
const SAFE=['/manifest.webmanifest','/static/pwa/icon-192.png','/static/pwa/icon-512.png'];
self.addEventListener('install',event=>{
  event.waitUntil(caches.open(CACHE).then(c=>c.addAll(SAFE)).then(()=>self.skipWaiting()));
});
self.addEventListener('activate',event=>{
  event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));
});
self.addEventListener('fetch',event=>{
  const u=new URL(event.request.url);
  if(event.request.method!=='GET' || u.origin!==location.origin){return;}
  if(SAFE.includes(u.pathname)){
    event.respondWith(caches.match(event.request).then(r=>r||fetch(event.request)));
    return;
  }
  event.respondWith(fetch(event.request));
});"""
    r=app.response_class(js,mimetype="application/javascript")
    r.headers["Service-Worker-Allowed"]="/"
    r.headers["Cache-Control"]="no-cache"
    return r

# === GAUR CRM v3.9 PREMIUM LOGO BRANDING ===
from jinja2 import ChoiceLoader, DictLoader

ROLE_DISPLAY_NAMES = {
    "MD": "Managing Director",
    "GM": "General Manager",
    "AM": "Assistant Manager",
    "HR": "Human Resources Manager",
    "RECEPTION": "Reception Executive",
    "FILING": "Filing Officer",
    "COUNSELOR": "Visa Counselor",
    "SENIOR_COUNSELOR": "Senior Visa Counselor",
    "TELECALLER": "Telecalling Executive",
    "ACCOUNTS": "Accounts Executive",
    "BRANCH_MANAGER": "Branch Manager",
    "TEAM_LEADER": "Team Leader",
}

def role_display(role, designation=None):
    designation=(designation or "").strip()
    if designation and designation.upper() not in ROLE_DISPLAY_NAMES:
        return designation
    return ROLE_DISPLAY_NAMES.get((role or "").upper(), designation or role or "Employee")

# v3.15: Railway/Gunicorn safe — register only after function exists.
app.jinja_env.globals["role_display"] = role_display

@app.context_processor
def inject_pending_enrollments():
    """GM ke liye pending enrollments count — payment received = 0."""
    try:
        u = current_user()
        if u and u["role"] in ("GM", "MD"):
            con = db()
            try:
                if u["role"] == "GM":
                    count = con.execute(
                        "SELECT COUNT(*) c FROM client_cases WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(total_received,0)=0",
                        (u["company_code"],)
                    ).fetchone()["c"]
                else:
                    count = con.execute(
                        "SELECT COUNT(*) c FROM client_cases WHERE COALESCE(deleted_at,'')='' AND COALESCE(total_received,0)=0"
                    ).fetchone()["c"]
                con.close()
                return {"pending_enrollments": int(count or 0)}
            except Exception:
                try: con.close()
                except Exception: pass
    except Exception:
        pass
    return {"pending_enrollments": 0}


@app.context_processor
def inject_report_brand():
    """Report branding context — ABC / XYZ / BOTH based on request param."""
    u=current_user()
    code=""
    try:
        code=(request.args.get("company") or request.args.get("company_code") or "").strip().upper()
    except Exception:
        code=""
    if code not in ("SCIC","WWIC"):
        if u and u["role"]!="MD" and u.get("company_code") in ("SCIC","WWIC"):
            code=u["company_code"]
        else:
            code="BOTH"
    if code=="SCIC":
        brand={"code":"SCIC","name":"ABC Immigration Consultants","short_name":"ABC",
               "address":os.environ.get("ABC_ADDRESS","SCO 411-412, 2nd Floor, Sector 35-C, Chandigarh – 160035"),
               "gstin":os.environ.get("GSTIN_ABC","").strip(),"scope":"ABC Immigration Consultants"}
    elif code=="WWIC":
        brand={"code":"WWIC","name":os.environ.get("XYZ_LEGAL_NAME","XYZ"),"short_name":"XYZ",
               "address":os.environ.get("XYZ_ADDRESS","").strip(),"gstin":os.environ.get("GSTIN_XYZ","").strip(),
               "scope":"XYZ"}
    else:
        brand={"code":"BOTH","name":"THE IMMIGRATION MANAGEMENT SYSTEM","short_name":"THE IMMIGRATION MANAGEMENT SYSTEM","address":"","gstin":"",
               "scope":"ABC + XYZ"}
    return {"report_brand":brand}


_GAUR_V337_TEMPLATES = {'base.html': """\n<!doctype html>\n<html>\n<head>\n<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">\n<title>{{ title or "ABC | XYZ Mini CRM" }}</title>\n<style>\n:root{--navy:#061629;--panel:#0b2744;--line:#315a7b;--gold:#e6b73f;--white:#fff;--red:#7f2e37;--yellow:#6d5a1b}\n*{box-sizing:border-box}body{margin:0;background:var(--navy);color:#fff;font-family:"Times New Roman",serif}\na{color:inherit;text-decoration:none}.layout{display:grid;grid-template-columns:245px 1fr;min-height:100vh}\naside{background:#031020;border-right:1px solid var(--line);padding:20px}.brand{color:var(--gold);font-size:28px;font-weight:bold}\n.userbox{border:1px solid var(--line);background:var(--panel);border-radius:12px;padding:12px;margin:22px 0}\nnav a{display:block;padding:11px 12px;border:1px solid var(--line);background:#0a223a;border-radius:9px;margin:8px 0}\nmain{padding:24px}.title{color:var(--gold);font-size:34px;margin:0 0 16px}.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px;margin-bottom:16px}\n.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}\ninput,select,textarea{width:100%;padding:10px;background:#071d32;color:#fff;border:1px solid #427398;border-radius:8px;font-family:inherit;font-size:16px}\nlabel{display:block;margin:8px 0 4px;font-weight:bold}.btn{display:inline-block;border:0;border-radius:8px;padding:10px 15px;background:var(--gold);color:#081423;font-weight:bold;cursor:pointer;font-family:inherit;font-size:16px}\ntable{width:100%;border-collapse:collapse;min-width:900px}th,td{padding:9px;border:1px solid var(--line);text-align:left}th{color:var(--gold)}\n.tablewrap{overflow:auto}.same{background:rgba(109,90,27,.4)}.cross{background:rgba(127,46,55,.5)}.invalid{background:rgba(85,85,85,.45)}\n.flash{padding:10px;border-radius:8px;margin-bottom:12px}.success{background:#174f34}.error{background:#6d2831}\n.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.metric{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px}.metric b{display:block;color:var(--gold);font-size:28px}\n\n.global-tools{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 16px}\n.toolbtn{display:inline-block;border:1px solid #427398;border-radius:8px;padding:9px 13px;background:#0a223a;color:#fff;font-weight:bold;cursor:pointer;font-family:inherit;font-size:15px}\n.toolbtn:hover{border-color:#e6b73f;color:#e6b73f}\n.rank1{border-left:5px solid #e6b73f!important}.rank2{border-left:5px solid #b9c1ca!important}.rank3{border-left:5px solid #b67a45!important}\n\n.performance-shell{display:grid;gap:16px}\n.kpi-grid{display:grid;grid-template-columns:repeat(6,minmax(150px,1fr));gap:12px}\n.kpi-card{background:linear-gradient(145deg,#0c2847,#071b30);border:1px solid #315a7b;border-radius:14px;padding:16px}\n.kpi-card .kpi-label{font-size:14px;color:#dce8f2}.kpi-card .kpi-value{font-size:30px;color:#e6b73f;font-weight:bold;margin-top:6px}\n.kpi-card .kpi-sub{font-size:13px;color:#9ec2db;margin-top:5px}\n.performance-grid{display:grid;grid-template-columns:1.5fr 1fr;gap:14px}\n.rank-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;align-items:end}\n.rank-card{background:linear-gradient(180deg,#0b2847,#061528);border:1px solid #315a7b;border-radius:16px;padding:16px;text-align:center}\n.rank-card.first{border-color:#e6b73f;box-shadow:0 0 0 1px rgba(230,183,63,.25) inset}\n.rank-card.second{border-color:#4e8fd4}.rank-card.third{border-color:#b24d97}\n.rank-card img{width:84px;height:84px;object-fit:cover;border-radius:50%;border:3px solid #315a7b}\n.rank-badge{font-size:28px;font-weight:bold;margin-bottom:6px}.rank-card.first .rank-badge{color:#e6b73f}.rank-card.second .rank-badge{color:#6fb1ff}.rank-card.third .rank-badge{color:#ef6ac9}\n.rank-score{display:inline-block;margin:8px 0;padding:6px 14px;border-radius:30px;background:#143d63;font-size:22px;font-weight:bold}\n.chart-panel{background:linear-gradient(145deg,#0b2744,#06172a);border:1px solid #315a7b;border-radius:16px;padding:16px}\n.chart-svg{width:100%;height:auto;display:block}\n.donut-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}\n.donut-card{background:#0a223a;border:1px solid #315a7b;border-radius:14px;padding:14px}\n.donut{width:150px;height:150px;border-radius:50%;margin:12px auto;position:relative;background:conic-gradient(#e6b73f 0 28%,#3d8fd1 28% 52%,#d051ac 52% 71%,#314c68 71% 100%)}\n.donut:after{content:"";position:absolute;inset:30px;border-radius:50%;background:#0a223a}\n.donut-center{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;z-index:2;text-align:center;font-size:19px;font-weight:bold}\n.legend-row{display:flex;justify-content:space-between;gap:10px;margin:5px 0;font-size:14px}\n.scorebar{height:9px;border-radius:10px;background:#162c43;overflow:hidden;margin-top:5px}.scorebar>span{display:block;height:100%;background:linear-gradient(90deg,#e6b73f,#5bc0eb)}\n@media(max-width:1200px){.kpi-grid{grid-template-columns:repeat(3,1fr)}.performance-grid{grid-template-columns:1fr}}\n@media(max-width:800px){.kpi-grid,.rank-grid,.donut-grid{grid-template-columns:1fr}}\n\n@media print{\n  aside,.global-tools,.no-print{display:none!important}\n  .layout{display:block!important}\n  main{padding:0!important}\n  body{background:#fff!important;color:#000!important}\n  .card,.metric{background:#fff!important;color:#000!important;border:1px solid #999!important}\n  .title,th,h1,h2,h3{color:#000!important}\n  table{min-width:0!important;font-size:11pt}\n  .tablewrap{overflow:visible!important}\n}\n\n.company-logo-wrap{background:#fff;border-radius:14px;padding:10px;margin:0 auto 8px;text-align:center;overflow:hidden}.company-logo{display:block;max-width:100%;max-height:105px;margin:auto;object-fit:contain}.company-logo-wrap.wwic{padding:4px 8px}.company-logo-wrap.wwic .company-logo{max-height:92px}.lead-link{color:#8fc8ff;text-decoration:underline;font-weight:bold}.lead-link:hover{color:#e6b73f}.status-chip{display:inline-block;padding:4px 9px;border-radius:20px;background:#123c62;border:1px solid #427398}.profile-head{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;flex-wrap:wrap}.profile-meta{display:flex;gap:8px;flex-wrap:wrap}.profile-meta span{background:#0a223a;border:1px solid #315a7b;border-radius:20px;padding:7px 10px}.two-col{display:grid;grid-template-columns:1.4fr .8fr;gap:16px}.activity-item{border-left:3px solid #427398;padding:8px 12px;margin:8px 0;background:#071d32;border-radius:6px}.doc-card{display:flex;justify-content:space-between;gap:12px;align-items:center;border:1px solid #315a7b;border-radius:9px;padding:10px;margin:8px 0;background:#071d32}.allocation-table td b{color:#e6b73f;font-size:19px}@media(max-width:900px){.two-col{grid-template-columns:1fr}}\n@media(max-width:800px){.layout{display:block}aside{position:relative}.grid,.grid3,.metrics{grid-template-columns:1fr}main{padding:14px}}\n\n/* v3.36 Lead Ranking */\n.rank-legend{display:flex;gap:10px;flex-wrap:wrap;margin:10px 0 16px}\n.rank-legend span,.interest-badge{display:inline-flex;align-items:center;gap:6px;border-radius:20px;padding:6px 11px;font-weight:bold;border:1px solid rgba(255,255,255,.18)}\n.legend-hot,.interest-hot{background:#8e2635;color:#fff}\n.legend-warm,.interest-warm{background:#1f6b45;color:#fff}\n.legend-cold,.interest-cold{background:#184c78;color:#fff}\n.lead-hot td{background:rgba(142,38,53,.28)!important}\n.lead-warm td{background:rgba(31,107,69,.28)!important}\n.lead-cold td{background:rgba(24,76,120,.25)!important}\n.lead-hot td:first-child{border-left:6px solid #ff4d61}\n.lead-warm td:first-child{border-left:6px solid #48d58b}\n.lead-cold td:first-child{border-left:6px solid #4aa8ff}\n.interest-score{font-size:19px;font-weight:bold}\n.filterbar{display:grid;grid-template-columns:2fr repeat(3,minmax(150px,1fr));gap:10px;align-items:end}\n.filter-actions{display:flex;gap:8px;align-items:end}\n@media(max-width:1000px){.filterbar{grid-template-columns:1fr 1fr}}\n@media(max-width:700px){.filterbar{grid-template-columns:1fr}}\n\n\n/* v3.36 MD Company Analytics */\n.company-compare{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px}\n.company-panel{border:1px solid var(--line);border-radius:18px;padding:18px;background:var(--panel)}\n.company-panel h2{margin:0 0 6px;color:var(--gold);font-size:30px}\n.company-panel .subtle{opacity:.8;margin-bottom:14px}\n.company-metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}\n.company-metric{border:1px solid var(--line);border-radius:12px;padding:12px;background:#071d34}\n.company-metric span{display:block;font-size:14px;opacity:.85}\n.company-metric b{display:block;font-size:25px;color:var(--gold);margin-top:4px}\n.analytics-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}\n.pie-wrap{display:grid;grid-template-columns:220px 1fr;gap:20px;align-items:center}\n.pie-chart{width:210px;height:210px;border-radius:50%;position:relative;margin:auto;box-shadow:0 0 0 10px rgba(255,255,255,.03)}\n.pie-chart:after{content:"";position:absolute;inset:43px;background:var(--panel);border-radius:50%;box-shadow:inset 0 0 0 1px var(--line)}\n.pie-center{position:absolute;inset:0;z-index:2;display:flex;align-items:center;justify-content:center;flex-direction:column;text-align:center}\n.pie-center b{font-size:28px;color:var(--gold)}\n.pie-legend{display:grid;gap:9px}\n.pie-legend div{display:flex;justify-content:space-between;gap:10px;border-bottom:1px solid rgba(255,255,255,.12);padding-bottom:7px}\n.dot{width:12px;height:12px;border-radius:50%;display:inline-block;margin-right:7px}\n.dot-hot{background:#ff4d61}.dot-warm{background:#48d58b}.dot-cold{background:#4aa8ff}\n.company-title-scic{border-left:6px solid #e6b73f;padding-left:12px}\n.company-title-wwic{border-left:6px solid #55aaff;padding-left:12px}\n.positive-pill{display:inline-block;padding:5px 9px;border-radius:18px;background:#1f6b45;font-weight:bold}\n@media(max-width:1100px){.company-compare,.analytics-grid{grid-template-columns:1fr}.company-metrics{grid-template-columns:1fr 1fr}}\n@media(max-width:650px){.company-metrics{grid-template-columns:1fr}.pie-wrap{grid-template-columns:1fr}}\n\n\n/* v3.36 Monthly Performance & Champions */\n.month-toolbar{display:flex;gap:10px;align-items:end;flex-wrap:wrap;margin:0 0 16px}\n.month-toolbar>div{min-width:190px}\n.champion-banner{display:grid;grid-template-columns:110px 1fr auto;gap:18px;align-items:center;background:linear-gradient(135deg,#5d4510,#102d4d 55%,#071a2d);border:1px solid #e6b73f;border-radius:18px;padding:18px;margin-bottom:16px;box-shadow:0 0 24px rgba(230,183,63,.12)}\n.champion-crown{font-size:64px;text-align:center}.champion-name{font-size:29px;color:#e6b73f;font-weight:bold}.champion-score{font-size:38px;color:#e6b73f;font-weight:bold;text-align:right}\n.champion-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.champion-card{background:linear-gradient(145deg,#0c2847,#071b30);border:1px solid #315a7b;border-radius:15px;padding:14px}\n.champion-card.first{border-color:#e6b73f}.champion-card.second{border-color:#8aa7c0}.champion-card.third{border-color:#b67a45}.champion-card h3{margin:4px 0;color:#e6b73f}\n.month-score{font-size:28px;color:#e6b73f;font-weight:bold}.level-badge{display:inline-block;padding:6px 10px;border-radius:18px;font-weight:bold;margin-top:5px}\n.level-star{background:#725513;color:#fff}.level-strong{background:#1f6b45}.level-good{background:#184c78}.level-build{background:#5c3340}\n.month-strip{display:grid;grid-template-columns:repeat(12,minmax(72px,1fr));gap:6px;overflow:auto}.month-box{background:#071d32;border:1px solid #315a7b;border-radius:9px;padding:8px;text-align:center;min-width:72px}\n.month-box.current{border-color:#e6b73f;box-shadow:0 0 0 1px #e6b73f inset}.month-box b{display:block;color:#e6b73f;font-size:19px}\n.self-performance{border:1px solid #e6b73f!important;background:linear-gradient(145deg,#102d4d,#071b30)!important}\n.role-chip{display:inline-block;padding:4px 8px;border-radius:14px;background:#123c62;border:1px solid #427398;margin-right:5px}\n@media(max-width:950px){.champion-banner{grid-template-columns:1fr;text-align:center}.champion-score{text-align:center}.champion-grid{grid-template-columns:1fr}}\n\n\n/* v3.36 Modern Calendar */\n.gaur-date-wrap{position:relative;width:100%}\n.gaur-date-display{\n  width:100%;box-sizing:border-box;background:#071c31;border:1px solid #3d6b91;color:#fff;\n  border-radius:12px;padding:12px 46px 12px 14px;font:inherit;cursor:pointer;\n  transition:.18s ease;outline:none\n}\n.gaur-date-display:hover,.gaur-date-display:focus{border-color:#e6b73f;box-shadow:0 0 0 3px rgba(230,183,63,.10)}\n.gaur-date-icon{\n  position:absolute;right:12px;top:50%;transform:translateY(-50%);font-size:20px;\n  pointer-events:none;color:#e6b73f\n}\n.gaur-calendar{\n  position:absolute;z-index:99999;top:calc(100% + 8px);left:0;width:330px;max-width:92vw;\n  background:#07192b;border:1px solid #315f85;border-radius:18px;padding:14px;\n  box-shadow:0 22px 60px rgba(0,0,0,.48);display:none\n}\n.gaur-calendar.open{display:block;animation:gaurCalIn .14s ease-out}\n@keyframes gaurCalIn{from{opacity:0;transform:translateY(-5px) scale(.98)}to{opacity:1;transform:none}}\n.gaur-cal-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}\n.gaur-cal-title{font-weight:800;color:#e6b73f;font-size:18px}\n.gaur-cal-nav{\n  width:36px;height:36px;border-radius:10px;border:1px solid #315f85;background:#0d2b49;color:#fff;\n  cursor:pointer;font-size:20px\n}\n.gaur-cal-nav:hover{border-color:#e6b73f;color:#e6b73f}\n.gaur-cal-week,.gaur-cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:5px}\n.gaur-cal-week div{text-align:center;font-size:12px;opacity:.7;padding:5px 0}\n.gaur-day{\n  aspect-ratio:1;border:0;border-radius:10px;background:#0b2946;color:#fff;cursor:pointer;\n  font-weight:700;font-size:13px\n}\n.gaur-day:hover{background:#174873;color:#fff}\n.gaur-day.muted{opacity:.33}\n.gaur-day.today{box-shadow:inset 0 0 0 1px #55aaff}\n.gaur-day.selected{background:#e6b73f;color:#07192b;box-shadow:0 4px 14px rgba(230,183,63,.3)}\n.gaur-cal-actions{display:flex;justify-content:space-between;gap:8px;margin-top:12px;border-top:1px solid rgba(255,255,255,.12);padding-top:10px}\n.gaur-cal-actions button{\n  border:1px solid #315f85;background:#0d2b49;color:#fff;border-radius:9px;padding:8px 12px;cursor:pointer\n}\n.gaur-cal-actions button.primary{background:#e6b73f;color:#07192b;border-color:#e6b73f;font-weight:800}\n@media(max-width:500px){.gaur-calendar{position:fixed;left:50%;top:50%;transform:translate(-50%,-50%);width:min(340px,90vw)}.gaur-calendar.open{animation:none}}\n\n\n/* v3.36 Modern Month Picker */\n.gaur-month-wrap{position:relative;width:100%}\n.gaur-month-display{\n  width:100%;box-sizing:border-box;background:#071c31;border:1px solid #3d6b91;color:#fff;\n  border-radius:12px;padding:12px 46px 12px 14px;font:inherit;cursor:pointer;outline:none;\n  transition:.18s ease\n}\n.gaur-month-display:hover,.gaur-month-display:focus{border-color:#e6b73f;box-shadow:0 0 0 3px rgba(230,183,63,.10)}\n.gaur-month-icon{position:absolute;right:12px;top:50%;transform:translateY(-50%);font-size:20px;color:#e6b73f;pointer-events:none}\n.gaur-month-pop{\n  position:absolute;z-index:100000;top:calc(100% + 8px);left:0;width:330px;max-width:92vw;\n  background:#07192b;border:1px solid #315f85;border-radius:18px;padding:15px;\n  box-shadow:0 22px 60px rgba(0,0,0,.50);display:none\n}\n.gaur-month-pop.open{display:block;animation:gaurMonthIn .14s ease-out}\n@keyframes gaurMonthIn{from{opacity:0;transform:translateY(-5px) scale(.98)}to{opacity:1;transform:none}}\n.gaur-month-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px}\n.gaur-month-year{font-size:20px;font-weight:800;color:#e6b73f}\n.gaur-month-nav{\n  width:38px;height:38px;border-radius:11px;border:1px solid #315f85;background:#0d2b49;color:#fff;\n  cursor:pointer;font-size:20px\n}\n.gaur-month-nav:hover{border-color:#e6b73f;color:#e6b73f}\n.gaur-month-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:9px}\n.gaur-month-btn{\n  border:1px solid #315f85;background:#0b2946;color:#fff;border-radius:12px;padding:13px 7px;\n  cursor:pointer;font-weight:700;transition:.15s ease\n}\n.gaur-month-btn:hover{background:#174873;border-color:#55aaff}\n.gaur-month-btn.current{box-shadow:inset 0 0 0 1px #55aaff}\n.gaur-month-btn.selected{background:#e6b73f;color:#07192b;border-color:#e6b73f;box-shadow:0 5px 16px rgba(230,183,63,.28)}\n.gaur-month-actions{display:flex;justify-content:space-between;gap:8px;margin-top:14px;border-top:1px solid rgba(255,255,255,.12);padding-top:11px}\n.gaur-month-actions button{\n  border:1px solid #315f85;background:#0d2b49;color:#fff;border-radius:9px;padding:8px 12px;cursor:pointer\n}\n.gaur-month-actions button.primary{background:#e6b73f;color:#07192b;border-color:#e6b73f;font-weight:800}\n@media(max-width:500px){\n  .gaur-month-pop{position:fixed;left:50%;top:50%;transform:translate(-50%,-50%);width:min(340px,90vw)}\n  .gaur-month-pop.open{animation:none}\n}\n\n\n/* v3.36 Premium Logo Branding */\n.brand-panel{\n  background:linear-gradient(145deg,#ffffff,#eef4f8);\n  border:2px solid rgba(230,183,63,.65);border-radius:20px;padding:12px;margin:8px 0 14px;\n  box-shadow:0 12px 30px rgba(0,0,0,.28),0 0 0 1px rgba(255,255,255,.22) inset\n}\n.brand-panel img{display:block;width:100%;max-height:145px;object-fit:contain}\n.brand-caption{text-align:center;font-weight:800;letter-spacing:.5px;margin-top:8px;color:#e6b73f}\n.company-logo-title{display:flex;align-items:center;gap:14px}\n.company-logo-title img{\n  width:78px;height:78px;object-fit:contain;background:#fff;border-radius:16px;padding:7px;\n  border:2px solid rgba(230,183,63,.55);box-shadow:0 8px 22px rgba(0,0,0,.24)\n}\n.company-logo-title img.wwic-logo{border-color:#55aaff}\n.page-brand-watermark{\n  position:absolute;right:18px;top:12px;width:82px;height:82px;object-fit:contain;\n  opacity:.12;filter:none;pointer-events:none\n}\n.card,.company-panel{position:relative}\n.logo-strip{\n  display:flex;align-items:center;gap:12px;background:linear-gradient(90deg,#0a223b,#0e3152);\n  border:1px solid #315f85;border-radius:14px;padding:10px 14px;margin-bottom:14px\n}\n.logo-strip img{width:58px;height:58px;object-fit:contain;background:#fff;border-radius:12px;padding:5px}\n.logo-strip strong{font-size:22px;color:#e6b73f}\n@media(max-width:700px){\n  .company-logo-title img{width:60px;height:60px}.company-logo-title{gap:10px}\n  .page-brand-watermark{width:60px;height:60px}\n}\n\n/* v3.11 TRUE-COLOR TRANSPARENT LOGOS */\n.brand-panel{\n  background:transparent!important;\n  border:0!important;\n  box-shadow:none!important;\n  padding:2px!important;\n}\n.brand-panel img,\n.company-logo-title img,\n.logo-strip img{\n  background:transparent!important;\n  border:0!important;\n  box-shadow:none!important;\n  border-radius:0!important;\n  padding:0!important;\n  object-fit:contain!important;\n  mix-blend-mode:normal!important;\n  filter:none!important;\n}\n.logo-strip{\n  background:transparent!important;\n  border:0!important;\n  box-shadow:none!important;\n  padding-left:0!important;\n}\n.company-logo-title img{width:96px!important;height:96px!important}\n.logo-strip img{width:74px!important;height:74px!important}\n\n\n/* v3.36 Unified Branding + User Profile Photos */\n.company-logo-wrap{background:transparent!important;border:0!important;box-shadow:none!important;padding:0!important;margin-bottom:14px!important}\n.company-logo{width:100%!important;max-height:175px!important;object-fit:contain!important;background:transparent!important;border:0!important;box-shadow:none!important;filter:none!important;mix-blend-mode:normal!important}\n.company-logo-title img{width:110px!important;height:110px!important;background:transparent!important;border:0!important;box-shadow:none!important;padding:0!important;filter:none!important;mix-blend-mode:normal!important}\n.logo-strip img{width:88px!important;height:88px!important;background:transparent!important;border:0!important;box-shadow:none!important;padding:0!important;filter:none!important;mix-blend-mode:normal!important}\n.logo-strip{background:transparent!important;border:0!important;box-shadow:none!important;padding-left:0!important}\n.userbox{display:flex;flex-direction:column;align-items:center;text-align:center;gap:10px;padding:18px 12px!important}\n.user-avatar{width:136px;height:136px;border-radius:50%;object-fit:cover;object-position:center;border:4px solid #e6b73f;box-shadow:0 8px 24px rgba(0,0,0,.34);background:#0a223a}\n.user-avatar-placeholder{width:136px;height:136px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#0a223a;border:4px solid #315a7b;color:#e6b73f;font-size:42px;font-weight:bold}\n.userbox>div:last-child{width:100%;line-height:1.25}.userbox>div:last-child>b{display:block;font-size:20px;line-height:1.12;white-space:normal;word-break:normal;margin-bottom:5px}.profile-mini-link{display:inline-block;margin-top:8px;font-size:14px;color:#e6b73f;text-decoration:underline;line-height:1.2}\n.profile-photo-large{width:160px;height:160px;border-radius:50%;object-fit:cover;border:4px solid #e6b73f;box-shadow:0 10px 32px rgba(0,0,0,.35);background:#071d32}\n.profile-shell{display:grid;grid-template-columns:190px 1fr;gap:22px;align-items:start}\n@media(max-width:700px){.profile-shell{grid-template-columns:1fr}.company-logo{max-height:130px!important}}\n\n\n/* v3.36 Professional Employee Identity System */\n.identity-grid{display:grid;grid-template-columns:repeat(4,minmax(170px,1fr));gap:12px}\n.identity-card{background:linear-gradient(145deg,#0c2947,#071a2e);border:1px solid #315f85;border-radius:16px;padding:16px;position:relative;overflow:hidden}\n.identity-card::after{content:"";position:absolute;width:120px;height:120px;border-radius:50%;right:-48px;top:-48px;background:rgba(230,183,63,.06)}\n.identity-card h3{margin:4px 0;color:#e6b73f}\n.identity-avatar{width:92px;height:92px;border-radius:50%;object-fit:cover;border:3px solid #e6b73f;background:#071d32}\n.identity-avatar.big{width:180px;height:180px;border-width:4px}\n.identity-placeholder{width:92px;height:92px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:34px;font-weight:800;color:#e6b73f;border:3px solid #315f85;background:#071d32}\n.identity-placeholder.big{width:180px;height:180px;font-size:64px}\n.identity-badge{display:inline-block;padding:5px 9px;border-radius:18px;border:1px solid #315f85;background:#0b2946;margin:3px 4px 3px 0;font-size:13px}\n.identity-badge.gold{border-color:#e6b73f;color:#e6b73f}\n.identity-badge.green{border-color:#42c98a;color:#71e2ad}\n.identity-badge.red{border-color:#e46767;color:#ff9a9a}\n.identity-profile{display:grid;grid-template-columns:220px 1fr;gap:22px}\n.identity-hero{background:linear-gradient(135deg,#0f3457,#081c31 66%,#4c3b0d);border:1px solid #e6b73f;border-radius:20px;padding:22px;box-shadow:0 12px 34px rgba(0,0,0,.28)}\n.identity-name{font-size:32px;color:#e6b73f;font-weight:800}\n.identity-stat-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:14px}\n.identity-stat{background:#071d32;border:1px solid #315a7b;border-radius:13px;padding:12px}\n.identity-stat b{display:block;color:#e6b73f;font-size:24px;margin-top:3px}\n.id-card-shell{max-width:780px;margin:auto;background:linear-gradient(135deg,#06182a,#10375b 58%,#4d3a0c);border:2px solid #e6b73f;border-radius:24px;padding:24px;box-shadow:0 20px 60px rgba(0,0,0,.42)}\n.id-card-top{display:flex;justify-content:space-between;align-items:center;gap:15px;border-bottom:1px solid rgba(255,255,255,.2);padding-bottom:14px}\n.id-card-body{display:grid;grid-template-columns:190px 1fr;gap:22px;padding-top:20px}\n.id-card-photo{width:170px;height:190px;object-fit:cover;border-radius:18px;border:4px solid #e6b73f;background:#071d32}\n.id-card-code{font-size:26px;letter-spacing:2px;color:#e6b73f;font-weight:800}\n.id-verify-box{border:1px dashed #55aaff;border-radius:12px;padding:12px;margin-top:14px;background:rgba(0,0,0,.16)}\n.readonly-value{padding:11px 12px;background:#071d32;border:1px solid #315a7b;border-radius:10px;min-height:20px}\n.identity-table-photo{width:48px;height:48px;border-radius:50%;object-fit:cover;border:2px solid #e6b73f}\n.identity-actions{display:flex;gap:7px;flex-wrap:wrap}\n@media(max-width:1050px){.identity-grid{grid-template-columns:repeat(2,1fr)}.identity-stat-grid{grid-template-columns:repeat(2,1fr)}}\n@media(max-width:700px){.identity-grid{grid-template-columns:1fr}.identity-profile,.id-card-body{grid-template-columns:1fr}.identity-stat-grid{grid-template-columns:1fr}.id-card-photo{margin:auto}}\n\n\n/* v3.19 Dashboard Employee Identity Header */\n.dashboard-heading-row{display:grid;grid-template-columns:minmax(360px,1fr) minmax(470px,690px);gap:18px;align-items:start;margin-bottom:10px}\n.dashboard-identity-card{display:grid;grid-template-columns:108px 1fr 138px;gap:16px;align-items:center;background:linear-gradient(135deg,#10385d 0%,#0a2845 60%,#362d13 100%);border:1px solid #e6b73f;border-radius:18px;padding:14px 16px;box-shadow:0 10px 28px rgba(0,0,0,.25)}\n.dashboard-identity-photo{width:96px;height:96px;border-radius:50%;object-fit:cover;border:3px solid #e6b73f;background:#071d32}\n.dashboard-identity-placeholder{width:96px;height:96px;border-radius:50%;display:flex;align-items:center;justify-content:center;border:3px solid #315f85;background:#071d32;color:#e6b73f;font-size:38px;font-weight:900}\n.dashboard-identity-name{font-size:24px;font-weight:900;color:#e6b73f;line-height:1.08}\n.dashboard-identity-designation{font-size:16px;font-weight:700;margin-top:4px}\n.dashboard-identity-code{color:#b8cee1;font-size:13px;margin-top:7px}\n.dashboard-identity-meta{display:flex;gap:7px;flex-wrap:wrap;margin-top:8px}\n.dashboard-identity-meta span{border:1px solid #315a7b;background:#071d32;border-radius:18px;padding:4px 8px;font-size:12px}\n.dashboard-rank-box{text-align:center;border-left:1px solid rgba(230,183,63,.35);padding-left:14px}\n.dashboard-rank-number{font-size:34px;font-weight:900;color:#e6b73f}\n.dashboard-rank-label{font-size:12px;color:#b8cee1}\n.dashboard-score{font-size:23px;font-weight:900;color:#55dca0;margin-top:8px}\n.dashboard-profile-link{display:inline-block;margin-top:7px;color:#e6b73f;font-size:12px;text-decoration:underline}\n@media(max-width:1100px){.dashboard-heading-row{grid-template-columns:1fr}.dashboard-identity-card{max-width:760px}}\n@media(max-width:650px){.dashboard-identity-card{grid-template-columns:82px 1fr}.dashboard-identity-photo,.dashboard-identity-placeholder{width:74px;height:74px}.dashboard-rank-box{grid-column:1/-1;border-left:0;border-top:1px solid rgba(230,183,63,.35);padding:10px 0 0;display:flex;gap:14px;justify-content:center;align-items:center}}\n\n\n@media(max-width:900px){.chat-page-grid{grid-template-columns:1fr}.chat-thread{height:560px}.broadcast-form-grid{grid-template-columns:1fr}}\n@media(max-width:520px){}\n\n\n/* v3.19 Chat file sharing */\n.chat-attach-row{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:8px}\n.attach-btn{display:inline-flex;align-items:center;gap:7px;background:#12395d;border:1px solid #55aaff;border-radius:10px;padding:9px 12px;cursor:pointer;font-weight:800}\n.attach-btn:hover{border-color:#e6b73f}.attach-file-name{font-size:12px;opacity:.75;max-width:360px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}\n.official-share-box{margin-top:8px;padding:9px;border:1px dashed #e6b73f;border-radius:10px;background:#0a2138}\n.official-share-box label{color:#e6b73f;display:block;margin-bottom:5px}.official-share-box small{display:block;opacity:.72;margin-top:5px}\n.chat-file-card{display:flex;justify-content:space-between;align-items:center;gap:10px;border:1px solid #55aaff;background:#0a2741;border-radius:12px;padding:10px;margin-top:7px}\n.chat-file-card .toolbtn{white-space:nowrap}\n\n\n/* v3.32 Railway Healthcheck Startup Rescue */\n.priority-strip{\n  margin:0 0 14px 0;border:1px solid #e6b73f;border-radius:14px;\n  background:linear-gradient(90deg,#3b171c 0%,#17324b 58%,#102a42 100%);\n  padding:10px 13px;display:flex;align-items:center;gap:12px;box-shadow:0 6px 18px rgba(0,0,0,.22)\n}\n.priority-strip.urgent{border-color:#ff5f69;background:linear-gradient(90deg,#55181e,#172e45)}\n.priority-strip .priority-icon{font-size:22px;flex:0 0 auto}\n.priority-strip .priority-content{min-width:0;flex:1}\n.priority-strip .priority-title{font-weight:900;color:#ffd05a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}\n.priority-strip .priority-text{font-size:13px;opacity:.9;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}\n.priority-strip .priority-action{white-space:nowrap}\n.notif-bell{\n position:fixed;right:96px;bottom:24px;width:58px;height:58px;border-radius:50%;\n border:1px solid #315f85;background:#0d3151;color:#fff;box-shadow:0 10px 28px rgba(0,0,0,.35);\n font-size:25px;cursor:pointer;z-index:99998\n}\n.notif-count{\n position:absolute;right:-4px;top:-4px;min-width:24px;height:24px;border-radius:14px;background:#ff4f5e;\n color:#fff;font-size:12px;font-weight:900;display:flex;align-items:center;justify-content:center;padding:0 6px\n}\n.notif-panel{\n position:fixed;right:24px;bottom:95px;width:430px;max-width:calc(100vw - 24px);height:620px;\n background:#06182a;border:1px solid #315f85;border-radius:18px;box-shadow:0 22px 60px rgba(0,0,0,.48);\n z-index:99997;display:none;overflow:hidden\n}\n.notif-panel.open{display:grid;grid-template-rows:auto auto 1fr auto}\n.notif-head{display:flex;justify-content:space-between;align-items:center;padding:14px 16px;background:#0b2946;border-bottom:1px solid #315f85}\n.notif-head strong{color:#e6b73f;font-size:18px}\n.notif-filters{display:flex;gap:7px;padding:10px 12px;border-bottom:1px solid #23455f;flex-wrap:wrap}\n.notif-filter{border:1px solid #315f85;background:#0a223a;color:#fff;border-radius:18px;padding:6px 10px;cursor:pointer;font-size:12px}\n.notif-filter.active{background:#e6b73f;color:#071629;border-color:#e6b73f}\n.notif-list{overflow:auto;padding:10px}\n.notif-item{border:1px solid #274c6a;border-radius:13px;padding:11px;margin-bottom:9px;background:#081e34;cursor:pointer}\n.notif-item:hover{border-color:#e6b73f}\n.notif-item.unread{box-shadow:inset 4px 0 0 #e6b73f;background:#0c2945}\n.notif-item.urgent{box-shadow:inset 4px 0 0 #ff4f5e;background:#2b171d}\n.notif-item.lead{box-shadow:inset 4px 0 0 #56aaff}\n.notif-item.target{box-shadow:inset 4px 0 0 #55dca0}\n.notif-item .n-title{font-weight:900;color:#f4c441}\n.notif-item .n-meta{font-size:11px;opacity:.66;margin-top:5px}\n.notif-item .n-body{font-size:13px;margin-top:5px;line-height:1.35}\n.notif-footer{padding:10px 12px;border-top:1px solid #315f85;background:#081e34;display:flex;gap:8px}\n.notification-page-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}\n.notification-big-card{border:1px solid #315f85;border-radius:14px;background:#0a223a;padding:14px}\n.notification-big-card.urgent{border-color:#ff5f69;background:#2a171d}\n.notification-big-card.unread{box-shadow:inset 4px 0 0 #e6b73f}\n@media(max-width:760px){\n .notif-bell{right:86px;bottom:16px}.notif-panel{right:10px;bottom:82px;width:calc(100vw - 20px);height:72vh}\n .notification-page-grid{grid-template-columns:1fr}\n .priority-strip{align-items:flex-start}.priority-strip .priority-text{white-space:normal}\n}\n\n\n/* v3.32 Activity & Security Center */\n.audit-kpis{display:grid;grid-template-columns:repeat(6,minmax(135px,1fr));gap:10px;margin-bottom:14px}\n.audit-kpi{background:#0a223a;border:1px solid #315f85;border-radius:14px;padding:13px}.audit-kpi b{display:block;color:#e6b73f;font-size:25px;margin-top:3px}.audit-kpi span{font-size:12px;opacity:.75}.audit-kpi.critical{border-color:#ff5f69}.audit-kpi.critical b{color:#ff747e}\n.audit-filters{display:grid;grid-template-columns:repeat(6,minmax(145px,1fr));gap:10px;align-items:end}\n.audit-event{border:1px solid #294d69;border-radius:13px;padding:12px;margin-bottom:9px;background:#081e34}.audit-event.critical{border-color:#ff5f69;box-shadow:inset 4px 0 0 #ff5f69}.audit-event.security{box-shadow:inset 4px 0 0 #e6b73f}.audit-event.client{box-shadow:inset 4px 0 0 #55aaff}.audit-event.admin{box-shadow:inset 4px 0 0 #55dca0}\n.audit-title{font-weight:900;color:#e6b73f}.audit-meta{font-size:11px;opacity:.65;margin-top:6px}.audit-badge{display:inline-block;padding:4px 8px;border:1px solid #315f85;border-radius:14px;font-size:11px;margin:4px 4px 0 0}.audit-change{display:grid;grid-template-columns:1fr auto 1fr;gap:8px;align-items:center;margin-top:8px}.audit-value{background:#07192b;border:1px solid #315a7b;border-radius:9px;padding:8px;white-space:pre-wrap;word-break:break-word}.audit-arrow{color:#e6b73f;font-weight:900}\n@media(max-width:1150px){.audit-kpis{grid-template-columns:repeat(3,1fr)}.audit-filters{grid-template-columns:repeat(3,1fr)}}@media(max-width:700px){.audit-kpis,.audit-filters{grid-template-columns:1fr}.audit-change{grid-template-columns:1fr}.audit-arrow{text-align:center}}\n\n.finance-kpis{display:grid;grid-template-columns:repeat(6,minmax(140px,1fr));gap:10px;margin:12px 0 18px}.finance-kpi{background:#0a223a;border:1px solid #315f85;border-radius:14px;padding:13px}.finance-kpi b{display:block;color:#e6b73f;font-size:24px;margin-top:4px}.finance-tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px}.finance-tabs a{padding:9px 12px;border:1px solid #315f85;border-radius:9px;background:#081e34}.finance-form{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.money-in{border-left:5px solid #55dca0}.money-out{border-left:5px solid #ff7a83}.refund-row{border-left:5px solid #e6b73f}.confidential{border:1px solid #e6b73f;background:#231f10;border-radius:10px;padding:10px}.bank-card{border:1px solid #315f85;border-radius:12px;padding:12px;background:#081e34;margin-bottom:8px}@media(max-width:900px){.finance-kpis,.finance-form{grid-template-columns:1fr 1fr}}@media(max-width:650px){.finance-kpis,.finance-form{grid-template-columns:1fr}}\n\n/* v3.32 AI Bank Reconciliation */\n.recon-kpis{display:grid;grid-template-columns:repeat(6,minmax(135px,1fr));gap:10px;margin:12px 0 18px}\n.recon-kpi{background:#0a223a;border:1px solid #315f85;border-radius:14px;padding:13px}.recon-kpi b{display:block;font-size:24px;color:#e6b73f;margin-top:4px}\n.recon-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}.recon-card{border:1px solid #315f85;border-radius:13px;background:#081e34;padding:12px;margin-bottom:10px}\n.recon-card.verified{border-left:5px solid #55dca0}.recon-card.possible{border-left:5px solid #e6b73f}.recon-card.unmatched{border-left:5px solid #ff6b76}.recon-score{font-size:24px;font-weight:900;color:#e6b73f}.recon-detail{font-size:12px;opacity:.78;margin-top:5px}.recon-actions{display:flex;gap:7px;flex-wrap:wrap;margin-top:9px}.recon-table td,.recon-table th{font-size:12px}\n@media(max-width:900px){.recon-kpis{grid-template-columns:repeat(3,1fr)}.recon-grid{grid-template-columns:1fr}}@media(max-width:650px){.recon-kpis{grid-template-columns:1fr}}\n\n.batch-new{display:inline-block;background:#155f44;border:1px solid #55dca0;color:#d9fff0;padding:3px 8px;border-radius:15px;font-size:11px;font-weight:bold}\n.batch-old{display:inline-block;background:#263c59;border:1px solid #6f93b9;color:#e8f4ff;padding:3px 8px;border-radius:15px;font-size:11px;font-weight:bold}\n.batch-legacy{display:inline-block;background:#553e19;border:1px solid #e6b73f;color:#ffe7a3;padding:3px 8px;border-radius:15px;font-size:11px;font-weight:bold}\n.lifecycle-meta{font-size:11px;opacity:.82;line-height:1.5}\n\naside a.nav-active{background:linear-gradient(90deg,rgba(230,183,63,.28),rgba(230,183,63,.07))!important;border-left:4px solid #e6b73f!important;color:#ffe49a!important;box-shadow:inset 0 0 18px rgba(230,183,63,.08),0 0 10px rgba(230,183,63,.08);font-weight:800}\naside a.nav-active::after{content:'  ●';color:#e6b73f;font-size:10px}\n\n/* v3.32 Monthly AM Leaderboard */\n.am-leaderboard-table tr.rank-one{background:linear-gradient(90deg,rgba(230,183,63,.20),rgba(230,183,63,.04));box-shadow:inset 5px 0 #e6b73f}\n.am-leaderboard-table tr.rank-two{background:linear-gradient(90deg,rgba(185,207,224,.15),rgba(185,207,224,.03));box-shadow:inset 5px 0 #b9cfe0}\n.am-leaderboard-table tr.rank-three{background:linear-gradient(90deg,rgba(190,126,69,.14),rgba(190,126,69,.03));box-shadow:inset 5px 0 #be7e45}\n.am-rank-photo{width:44px;height:44px;border-radius:50%;object-fit:cover;border:2px solid #315f85;vertical-align:middle;margin-right:8px}\n.am-rank-avatar{display:inline-flex;width:44px;height:44px;border-radius:50%;align-items:center;justify-content:center;border:2px solid #315f85;color:#e6b73f;background:#071d32;font-weight:900;margin-right:8px}\n.company-champion-wrap{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin:14px 0}\n.company-champion-panel{background:linear-gradient(145deg,#0b2744,#071b30);border:1px solid #315a7b;border-radius:16px;padding:15px}\n.company-champion-panel h2{color:#e6b73f;margin-top:0}\n.mini-winner{display:flex;align-items:center;justify-content:space-between;gap:10px;border:1px solid #315a7b;border-radius:12px;padding:10px;margin:8px 0}\n.mini-winner.first{border-color:#e6b73f}.mini-winner.second{border-color:#b9cfe0}\n@media(max-width:850px){.company-champion-wrap{grid-template-columns:1fr}}\n\n/* v3.32 Back-date Calendar Fix */\n.gaur-cal-jump{display:grid;grid-template-columns:1fr 92px;gap:7px;flex:1;margin:0 8px}\n.gaur-cal-jump select{width:100%;min-width:0;background:#0d2b49;color:#fff;border:1px solid #315f85;border-radius:9px;padding:8px 7px;font:inherit;font-weight:700}\n.gaur-cal-jump select:focus{outline:none;border-color:#e6b73f;box-shadow:0 0 0 2px rgba(230,183,63,.12)}\n.gaur-date-display{cursor:text!important}\n.gaur-date-help{font-size:10px;opacity:.68;margin-top:4px}\n\n/* v3.32 Reception Attendance */\n.attendance-top{display:grid;grid-template-columns:1.2fr .7fr .7fr 1fr auto;gap:10px;align-items:end}\n.attendance-stat{display:inline-block;padding:8px 12px;border-radius:12px;background:#0b2946;border:1px solid #315f85;margin:4px 6px 4px 0}\n.attendance-present{box-shadow:inset 5px 0 #55dca0}\n.attendance-missing{box-shadow:inset 5px 0 #ff6571}\n.wa-mini{background:#1c7b4e!important;color:#fff!important}\n@media(max-width:900px){.attendance-top{grid-template-columns:1fr 1fr}.attendance-top .wide{grid-column:1/-1}}\n\n/* v3.33 MNC Professional Report Engine */\n.report-print-header,.report-print-footer,.report-print-meta{display:none}\n@media print{\n  @page{size:A4 portrait;margin:31mm 11mm 20mm 11mm}\n  html,body{background:#fff!important;color:#111!important;font-family:Arial,Helvetica,sans-serif!important}\n  body{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}\n  aside,.global-tools,.no-print,.notif-bell,.notif-panel,.priority-strip,\n  button,.toolbtn,.profile-mini-link{display:none!important}\n  form:not(.print-keep),.audit-filters,.filterbar,.filters,.finance-form,.attendance-top,.bulkbar{display:none!important}\n  .layout{display:block!important}\n  main{padding:0!important;margin:0!important;width:auto!important}\n  .report-print-header{\n    display:flex!important;position:fixed;left:0;right:0;top:-25mm;height:22mm;\n    align-items:center;justify-content:space-between;gap:10px;border-bottom:2px solid #b98a16;\n    background:#fff;color:#111;padding:0 2mm 3mm;z-index:99999\n  }\n  .report-logo-box{display:flex;align-items:center;gap:7px;min-width:170px}\n  .report-logo{max-width:72px;max-height:18mm;object-fit:contain}\n  .report-logo.dual{max-width:52px;max-height:15mm}\n  .report-company{font-weight:800;font-size:10pt;line-height:1.15;color:#16324b}\n  .report-company small{display:block;font-size:7pt;color:#666;font-weight:600;margin-top:2px}\n  .report-head-center{text-align:center;flex:1}\n  .report-head-center b{display:block;font-size:14pt;color:#111}\n  .report-head-center span{font-size:7.5pt;color:#666}\n  .report-head-right{text-align:right;font-size:7pt;line-height:1.45;color:#555;min-width:150px}\n  .report-print-footer{\n    display:flex!important;position:fixed;left:0;right:0;bottom:-15mm;height:12mm;\n    align-items:flex-start;justify-content:space-between;gap:10px;border-top:1px solid #b98a16;\n    background:#fff;color:#555;padding-top:2.5mm;font-size:7pt;z-index:99999\n  }\n  .report-print-footer .conf{font-weight:800;color:#7d5c09;letter-spacing:.3px}\n  .report-print-footer .footer-mid{text-align:center;flex:1}\n  .report-print-footer .footer-right{text-align:right}\n  .report-print-meta{\n    display:grid!important;grid-template-columns:repeat(4,1fr);gap:7px;border:1px solid #ccd5dc;\n    background:#f7f9fb;padding:8px 10px;margin:0 0 12px;border-radius:6px;font-size:8pt;color:#333\n  }\n  .report-print-meta b{display:block;color:#16324b;font-size:6.8pt;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px}\n  .card,.metric,.company-panel,.company-champion-panel,.audit-event,.finance-kpi,.attendance-stat{\n    background:#fff!important;color:#111!important;border-color:#bac6cf!important;box-shadow:none!important\n  }\n  .title,th,h1,h2,h3,.audit-title{color:#16324b!important}\n  h1{font-size:21pt!important} h2{font-size:14pt!important} h3{font-size:11pt!important}\n  table{min-width:0!important;width:100%!important;border-collapse:collapse!important;font-size:8.2pt!important}\n  thead{display:table-header-group!important}\n  tfoot{display:table-footer-group!important}\n  th{background:#16324b!important;color:#fff!important;border:1px solid #bac6cf!important;padding:6px!important;text-align:left!important}\n  td{color:#111!important;border:1px solid #d5dde3!important;padding:6px!important;background:#fff!important}\n  tbody tr:nth-child(even) td{background:#f7f9fb!important}\n  tr,.audit-event,.notification-big-card,.metric,.finance-kpi{break-inside:avoid!important;page-break-inside:avoid!important}\n  h1,h2,h3{break-after:avoid!important;page-break-after:avoid!important}\n  .tablewrap{overflow:visible!important}\n  a{color:#111!important;text-decoration:none!important}\n  small,.subtle{color:#555!important}\n  .audit-badge,.role-chip,.status-chip,.interest-badge{background:#f2f5f7!important;color:#16324b!important;border-color:#9aaab7!important}\n  .audit-value{background:#f7f9fb!important;color:#111!important;border-color:#ccd5dc!important}\n  .champion-banner,.champion-card,.company-champion-panel{background:#fff!important;color:#111!important;border:1px solid #b98a16!important;box-shadow:none!important}\n  .champion-name,.champion-score,.month-score{color:#9b720f!important}\n}\n\n/* v3.36 Positive Data Pie - Large Company Logo Alignment */\n.analytics-grid>.card .logo-strip{display:flex!important;align-items:center!important;gap:18px!important;width:100%!important;margin:0 0 18px!important;padding:4px 8px!important;min-height:118px!important}\n.analytics-grid>.card .logo-strip img{width:auto!important;height:auto!important;max-width:190px!important;max-height:112px!important;object-fit:contain!important;flex:0 0 auto!important}\n.analytics-grid>.card .logo-strip strong{font-size:28px!important;color:var(--gold)!important;letter-spacing:.5px!important;line-height:1.05!important}\n.analytics-grid>.card:nth-child(1) .logo-strip{justify-content:flex-start!important;text-align:left!important}\n.analytics-grid>.card:nth-child(2) .logo-strip{justify-content:flex-end!important;text-align:right!important;flex-direction:row-reverse!important}\n.analytics-grid>.card:nth-child(2) .logo-strip img{max-width:210px!important;max-height:118px!important}\n@media(max-width:1100px){.analytics-grid>.card:nth-child(2) .logo-strip{justify-content:flex-start!important;text-align:left!important;flex-direction:row!important}}\n@media(max-width:650px){.analytics-grid>.card .logo-strip{min-height:90px!important;gap:10px!important}.analytics-grid>.card .logo-strip img{max-width:135px!important;max-height:82px!important}.analytics-grid>.card .logo-strip strong{font-size:21px!important}}\n@media print{.analytics-grid>.card .logo-strip{display:flex!important;min-height:70px!important}.analytics-grid>.card .logo-strip img{max-width:120px!important;max-height:65px!important}.analytics-grid>.card .logo-strip strong{font-size:18pt!important;color:#16324b!important}}\n\n/* v3.36 Modular Sidebar Navigation */\n.modular-nav{display:flex;flex-direction:column;gap:8px}\n.modular-nav a,.nav-group summary{\n  box-sizing:border-box;width:100%;min-height:48px;border:1px solid #315f85;border-radius:11px;\n  background:#0a2743;color:#fff;text-decoration:none;padding:11px 12px;display:flex;align-items:center;gap:10px;\n  font-weight:700;cursor:pointer;transition:.18s ease\n}\n.modular-nav a:hover,.nav-group summary:hover{border-color:#e6b73f;background:#0d3152;transform:translateX(2px)}\n.nav-num{width:25px;height:25px;min-width:25px;display:inline-flex;align-items:center;justify-content:center;border-radius:50%;\n  background:#e6b73f;color:#071629;font-size:12px;font-weight:900}\n.nav-icon{width:25px;text-align:center}\n.nav-group{border:0;margin:0;padding:0}\n.nav-group>summary{list-style:none}\n.nav-group>summary::-webkit-details-marker{display:none}\n.nav-chevron{margin-left:auto;font-size:24px;line-height:1;transition:.18s ease;color:#9dc9ec}\n.nav-group[open]>.nav-chevron,.nav-group[open]>summary .nav-chevron{transform:rotate(90deg);color:#e6b73f}\n.nav-group[open]>summary{border-color:#e6b73f;background:linear-gradient(90deg,rgba(230,183,63,.18),#0a2743)}\n.nav-submenu{margin:5px 0 2px 17px;padding-left:10px;border-left:2px solid #315f85;display:flex;flex-direction:column;gap:5px}\n.nav-submenu a{min-height:39px;padding:8px 10px;border-radius:9px;font-size:13px;background:#071e34}\n.nav-submenu a::before{content:'└';color:#6ea7d5;font-size:12px}\n.nav-badge{margin-left:auto;background:#d64755;padding:2px 7px;border-radius:12px;font-size:11px}\n.nav-badge.alert{background:#ff4f5e}\n.nav-utility{margin-top:8px}\n.nav-utility summary{background:#071d32}\n.nav-signout{margin-top:8px!important;border-color:#7e4050!important;background:#241925!important}\n.modular-nav a.nav-active{background:linear-gradient(90deg,rgba(230,183,63,.28),rgba(230,183,63,.07))!important;\n  border:1px solid #e6b73f!important;border-left:5px solid #e6b73f!important;color:#ffe49a!important;\n  box-shadow:inset 0 0 18px rgba(230,183,63,.08),0 0 10px rgba(230,183,63,.08);transform:none!important}\n.modular-nav a.nav-active::after{content:'●';margin-left:auto;color:#e6b73f;font-size:9px}\n@media(max-width:800px){\n  .modular-nav{gap:6px}.modular-nav a,.nav-group summary{min-height:44px;padding:9px 10px}\n}\n\n/* v3.36 Bank Manager / Payment Sharing */\n.bank-manager-grid{display:grid;grid-template-columns:repeat(2,minmax(320px,1fr));gap:16px}\n.payment-bank-card{background:linear-gradient(145deg,#0b2948,#06182b);border:1px solid #315f85;border-radius:18px;padding:18px;position:relative;overflow:hidden}\n.payment-bank-head{display:flex;align-items:center;gap:12px;margin-bottom:14px}\n.bank-auto-logo{width:64px;height:64px;border-radius:14px;background:#fff;object-fit:contain;padding:7px}\n.bank-logo-fallback{width:64px;height:64px;border-radius:14px;background:#e6b73f;color:#071629;display:flex;align-items:center;justify-content:center;font-size:22px;font-weight:900}\n.verified-account{display:inline-block;border:1px solid #4fcf93;color:#7ff0b7;border-radius:20px;padding:4px 9px;font-size:12px;font-weight:bold}\n.payment-details{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:12px 0}\n.payment-details div{background:#071d32;border:1px solid #315a7b;border-radius:10px;padding:9px}\n.payment-details small{display:block;opacity:.65}.payment-details b{word-break:break-all}\n.payment-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}\n.qr-panel{display:none;margin-top:12px;padding:14px;background:#fff;color:#111;border-radius:14px;text-align:center}\n.qr-panel.open{display:block}.qr-panel canvas,.qr-panel img{margin:auto}\n.share-mobile{max-width:220px}\n@media(max-width:900px){.bank-manager-grid{grid-template-columns:1fr}.payment-details{grid-template-columns:1fr}}\n\n.competition-livebar{display:none;position:sticky;top:0;z-index:90;margin:-4px 0 14px;border:1px solid #e6b73f;border-radius:14px;background:linear-gradient(90deg,#102c4a,#1b2238,#102c4a);padding:10px 14px}\n.competition-livebar.show{display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:14px}.team-live{text-align:center}.team-live b{display:block;color:#e6b73f}.team-live strong{font-size:23px}.comp-vs{font-weight:900;background:#7d2f3a;padding:7px 10px;border-radius:50%}\n.comp-meter{height:7px;background:#07192c;border-radius:8px;overflow:hidden;margin-top:5px;border:1px solid #315f85}.comp-meter span{display:block;height:100%;background:#e6b73f}\n.competition-popup{position:fixed;inset:0;z-index:9999999;display:flex;align-items:center;justify-content:center;padding:18px;background:radial-gradient(circle at 50% 42%,rgba(36,74,119,.55),rgba(0,7,18,.92) 58%,rgba(0,0,0,.98));opacity:0;visibility:hidden;pointer-events:none;transition:opacity .28s ease,visibility .28s ease;overflow:hidden}\n.competition-popup.show{opacity:1;visibility:visible;pointer-events:auto}.competition-popup:before{content:"";position:absolute;inset:-30%;background:conic-gradient(from 0deg,transparent,rgba(255,214,64,.10),transparent,rgba(76,169,255,.12),transparent);animation:celebrateSpin 9s linear infinite}.celebrate-stage{position:relative;width:min(940px,96vw);min-height:min(690px,92vh);border:2px solid #f5c84b;border-radius:34px;background:linear-gradient(155deg,#071426 0%,#102f51 45%,#071426 100%);box-shadow:0 0 0 7px rgba(245,200,75,.08),0 30px 100px rgba(0,0,0,.72),0 0 60px rgba(245,200,75,.22) inset;padding:26px 30px 24px;text-align:center;overflow:hidden;transform:scale(.72) translateY(35px);transition:transform .48s cubic-bezier(.2,.9,.25,1.18)}\n.competition-popup.show .celebrate-stage{transform:scale(1) translateY(0)}.competition-popup.wwic .celebrate-stage{border-color:#55aaff;box-shadow:0 0 0 7px rgba(85,170,255,.08),0 30px 100px rgba(0,0,0,.72),0 0 60px rgba(85,170,255,.22) inset}.celebrate-glow{position:absolute;left:50%;top:43%;width:520px;height:520px;transform:translate(-50%,-50%);background:radial-gradient(circle,rgba(255,210,64,.30),rgba(255,210,64,.04) 52%,transparent 72%);filter:blur(2px);animation:celebrateGlow 1s ease-in-out infinite alternate}.competition-popup.wwic .celebrate-glow{background:radial-gradient(circle,rgba(85,170,255,.32),rgba(85,170,255,.05) 52%,transparent 72%)}.celebrate-crown{position:relative;z-index:2;font-size:72px;line-height:1;filter:drop-shadow(0 8px 14px rgba(0,0,0,.45));animation:crownBounce .75s ease-in-out infinite alternate}.comp-pop-title{position:relative;z-index:2;margin:6px 0 0;font-family:Arial Black,Arial,sans-serif;font-size:clamp(34px,6vw,76px);line-height:.98;color:#ffd85e;text-shadow:0 3px 0 #8c3517,0 7px 18px rgba(0,0,0,.7);letter-spacing:1px}.competition-popup.wwic .comp-pop-title{color:#d9efff;text-shadow:0 3px 0 #145792,0 7px 18px rgba(0,0,0,.7)}.celebrate-subtitle{position:relative;z-index:2;display:inline-block;margin:10px auto 17px;padding:8px 28px;border-radius:999px;background:linear-gradient(180deg,#1b66c7,#0d3f87);border:2px solid rgba(255,255,255,.34);font-family:Arial Black,Arial,sans-serif;font-size:clamp(21px,3.4vw,39px);letter-spacing:1px;box-shadow:0 8px 20px rgba(0,0,0,.35)}.celebrate-person{position:relative;z-index:2;display:flex;justify-content:center;align-items:center;gap:18px;flex-wrap:wrap;margin:2px auto 14px}.celebrate-avatar-wrap{width:170px;height:170px;border-radius:50%;padding:5px;background:linear-gradient(145deg,#fff7b0,#d99e14);box-shadow:0 0 0 8px rgba(255,210,75,.10),0 0 40px rgba(255,210,75,.38);animation:avatarPop .72s cubic-bezier(.2,.9,.3,1.3)}.competition-popup.wwic .celebrate-avatar-wrap{background:linear-gradient(145deg,#e5f6ff,#4096e8);box-shadow:0 0 0 8px rgba(85,170,255,.10),0 0 40px rgba(85,170,255,.38)}.celebrate-avatar{width:100%;height:100%;object-fit:cover;border-radius:50%;background:#09192a}.celebrate-avatar-fallback{width:100%;height:100%;display:flex;align-items:center;justify-content:center;border-radius:50%;background:#102b49;font-size:70px}.celebrate-person-info{text-align:left;min-width:260px}.celebrate-name{font-family:Arial Black,Arial,sans-serif;font-size:clamp(29px,4vw,48px);color:#fff}.celebrate-role{font-size:20px;color:#ffd85e;font-weight:800;letter-spacing:.6px}.competition-popup.wwic .celebrate-role{color:#7fc4ff}.celebrate-company{margin-top:6px;font-size:16px;opacity:.86}.celebrate-details{position:relative;z-index:2;display:grid;grid-template-columns:repeat(3,1fr);gap:10px;width:min(760px,100%);margin:10px auto}.celebrate-detail{padding:11px 12px;background:rgba(0,0,0,.27);border:1px solid rgba(255,255,255,.16);border-radius:13px}.celebrate-detail small{display:block;opacity:.67;margin-bottom:3px}.celebrate-detail b{display:block;font-size:19px;color:#ffe27b;word-break:break-word}.competition-popup.wwic .celebrate-detail b{color:#9fd1ff}.celebrate-save{position:relative;z-index:2;display:inline-flex;align-items:center;gap:8px;margin-top:7px;padding:10px 21px;border-radius:999px;background:#138b42;border:2px solid #5bea8b;font-size:20px;font-weight:900;box-shadow:0 8px 20px rgba(0,0,0,.3)}.celebrate-close{position:absolute;right:16px;top:12px;z-index:6;width:40px;height:40px;border-radius:50%;border:1px solid rgba(255,255,255,.35);background:rgba(0,0,0,.35);color:#fff;font-size:24px;cursor:pointer}.celebrate-countdown{position:relative;z-index:2;margin-top:8px;font-size:13px;opacity:.7}.confetti-piece{position:absolute;z-index:4;top:-12vh;width:10px;height:20px;border-radius:2px;animation:confettiFall linear forwards}.firework{position:absolute;z-index:1;width:8px;height:8px;border-radius:50%;box-shadow:0 -70px #ffd75e,49px -49px #ff5964,70px 0 #5bc0ff,49px 49px #66e49a,0 70px #ff8ce8,-49px 49px #ffb14e,-70px 0 #78a7ff,-49px -49px #d7ff61;animation:fireworkBurst 1.45s ease-out infinite}.fw1{left:13%;top:20%}.fw2{right:13%;top:18%;animation-delay:.4s}.fw3{left:20%;bottom:21%;animation-delay:.8s}.fw4{right:19%;bottom:20%;animation-delay:1.1s}.celebrate-side{position:absolute;z-index:2;font-size:54px;filter:drop-shadow(0 6px 10px rgba(0,0,0,.5));animation:sideFloat 1s ease-in-out infinite alternate}.side-left{left:4%;top:48%}.side-right{right:4%;top:48%;animation-delay:.35s}@keyframes celebrateSpin{to{transform:rotate(360deg)}}@keyframes celebrateGlow{from{transform:translate(-50%,-50%) scale(.88);opacity:.55}to{transform:translate(-50%,-50%) scale(1.08);opacity:1}}@keyframes crownBounce{from{transform:translateY(0) rotate(-3deg)}to{transform:translateY(-9px) rotate(3deg)}}@keyframes avatarPop{0%{transform:scale(.2) rotate(-16deg)}75%{transform:scale(1.08) rotate(2deg)}100%{transform:scale(1)}}@keyframes confettiFall{0%{transform:translate3d(0,-10vh,0) rotate(0);opacity:1}100%{transform:translate3d(var(--drift),115vh,0) rotate(900deg);opacity:.1}}@keyframes fireworkBurst{0%{transform:scale(.08);opacity:0}18%{opacity:1}70%{transform:scale(1);opacity:.9}100%{transform:scale(1.2);opacity:0}}@keyframes sideFloat{from{transform:translateY(-7px) rotate(-6deg)}to{transform:translateY(7px) rotate(6deg)}}@media(max-width:700px){.celebrate-stage{padding:18px 14px;min-height:auto}.celebrate-crown{font-size:54px}.celebrate-avatar-wrap{width:125px;height:125px}.celebrate-person-info{text-align:center;min-width:0}.celebrate-details{grid-template-columns:1fr}.celebrate-side{display:none}.firework{transform:scale(.7)}}\n.competition-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}.comp-team-card{border:1px solid #315f85;border-radius:16px;background:#081f37;padding:18px}.comp-team-card h2{color:#e6b73f}.comp-score-big{font-size:54px;font-weight:900}.comp-event{padding:11px;border-left:4px solid #e6b73f;background:#071d32;border-radius:9px;margin:7px 0}\n@media(max-width:850px){.competition-grid{grid-template-columns:1fr}.team-live strong{font-size:17px}}\n/* v3.98 Larger Sorted Sidebar Profile - ALL PORTALS */\n.userbox{display:flex!important;flex-direction:column!important;align-items:center!important;text-align:center!important;padding:20px 14px 22px!important;gap:0!important}\n.userbox>div:first-child{width:100%!important;display:flex!important;justify-content:center!important}\n.user-avatar,.user-avatar-placeholder{width:158px!important;height:158px!important;min-width:158px!important;min-height:158px!important;border-width:4px!important}\n.userbox>div:last-child{width:100%!important;margin-top:20px!important;line-height:1.32!important;text-align:center!important}\n.userbox>div:last-child>b{display:block!important;width:100%!important;font-size:20px!important;line-height:1.2!important;margin:0 auto 14px!important;text-align:center!important;white-space:normal!important;word-break:keep-all!important}\n.userbox>div:last-child>span{display:block!important;margin-top:5px!important}\n.profile-mini-link{display:inline-block!important;margin-top:13px!important;font-size:15px!important;line-height:1.2!important}\n@media(max-width:1200px){.user-avatar,.user-avatar-placeholder{width:148px!important;height:148px!important;min-width:148px!important;min-height:148px!important}}\n\n/* =========================================================\n   GAUR CRM v3.98 - APPROVED GOLD LUXURY COMPACT UI\n   Applies to MD / GM / AM / Employees / Accounts / Filing\n   ========================================================= */\n:root{\n  --gaur-bg:#031426;\n  --gaur-panel:#0b2b48;\n  --gaur-panel2:#071d32;\n  --gaur-line:#2d6286;\n  --gaur-gold:#f4bd32;\n  --gaur-gold2:#ffd968;\n  --gaur-text:#f7f8fb;\n  --gaur-muted:#b9c7d6;\n}\n\n/* Compact sidebar profile card */\n.userbox{\n  display:flex!important;\n  flex-direction:column!important;\n  align-items:center!important;\n  justify-content:flex-start!important;\n  text-align:center!important;\n  padding:14px 12px 15px!important;\n  gap:0!important;\n  min-height:0!important;\n  height:auto!important;\n  background:linear-gradient(180deg,#0c3151 0%,#08243d 100%)!important;\n  border:1px solid var(--gaur-line)!important;\n  border-radius:15px!important;\n  overflow:hidden!important;\n}\n.userbox>div:first-child{\n  width:100%!important;\n  display:flex!important;\n  justify-content:center!important;\n  align-items:center!important;\n  margin:0!important;\n  padding:0!important;\n}\n.user-avatar,.user-avatar-placeholder{\n  width:178px!important;\n  height:178px!important;\n  min-width:178px!important;\n  min-height:178px!important;\n  max-width:178px!important;\n  max-height:178px!important;\n  border:4px solid var(--gaur-gold)!important;\n  border-radius:50%!important;\n  object-fit:cover!important;\n  box-shadow:0 0 8px rgba(244,189,50,.7),0 0 20px rgba(244,189,50,.20)!important;\n  margin:0 auto!important;\n}\n.userbox>div:last-child{\n  width:100%!important;\n  margin:10px 0 0!important;\n  padding:0!important;\n  line-height:1.18!important;\n  text-align:center!important;\n}\n.userbox>div:last-child>b{\n  display:block!important;\n  width:100%!important;\n  margin:0 auto 8px!important;\n  padding:0!important;\n  color:var(--gaur-gold2)!important;\n  font-size:19px!important;\n  font-weight:800!important;\n  line-height:1.16!important;\n  text-align:center!important;\n  white-space:normal!important;\n  word-break:normal!important;\n  text-shadow:0 0 5px rgba(255,217,104,.60),0 0 12px rgba(244,189,50,.28)!important;\n}\n.userbox>div:last-child>span{\n  display:block!important;\n  margin:3px 0 0!important;\n  padding:0!important;\n  line-height:1.15!important;\n}\n.profile-mini-link{\n  display:inline-block!important;\n  margin:8px 0 0!important;\n  padding:0!important;\n  color:var(--gaur-gold2)!important;\n  font-size:14px!important;\n  line-height:1.15!important;\n  font-weight:700!important;\n  text-shadow:0 0 5px rgba(244,189,50,.35)!important;\n}\n\n/* Golden/shining names and important headings throughout the app */\nh1,h2,h3,\n.dashboard-title,\n.page-title,\n.card-title,\n.section-title{\n  color:var(--gaur-gold)!important;\n  text-shadow:0 0 5px rgba(255,215,85,.35)!important;\n}\n.userbox b,\n.employee-name,\n.am-name,\n.manager-name,\n.performer-name,\n.profile-name,\n.client-name{\n  color:var(--gaur-gold2)!important;\n  font-weight:800!important;\n  text-shadow:0 0 5px rgba(255,217,104,.55),0 0 12px rgba(244,189,50,.18)!important;\n}\n\n/* Refined cards and controls, without changing app behavior */\n.card,.panel,.stat-card,.metric-card,.dashboard-card{\n  border-color:var(--gaur-line)!important;\n}\nbutton,.btn,a.btn{\n  border-radius:10px!important;\n}\n.sidebar .nav-item,.sidebar a{\n  transition:box-shadow .15s ease,border-color .15s ease!important;\n}\n.sidebar .nav-item:hover,.sidebar a:hover{\n  box-shadow:0 0 10px rgba(244,189,50,.13)!important;\n}\n\n/* Keep mobile/tablet compact and usable */\n@media(max-width:1200px){\n  .user-avatar,.user-avatar-placeholder{\n    width:158px!important;height:158px!important;\n    min-width:158px!important;min-height:158px!important;\n    max-width:158px!important;max-height:158px!important;\n  }\n  .userbox>div:last-child>b{font-size:18px!important}\n}\n@media(max-width:760px){\n  .userbox{padding:10px!important}\n  .user-avatar,.user-avatar-placeholder{\n    width:132px!important;height:132px!important;\n    min-width:132px!important;min-height:132px!important;\n    max-width:132px!important;max-height:132px!important;\n  }\n  .userbox>div:last-child{margin-top:8px!important}\n  .userbox>div:last-child>b{font-size:17px!important;margin-bottom:6px!important}\n  .profile-mini-link{margin-top:6px!important}\n}\n.competition-livebar.lux-score{position:sticky!important;top:0!important;z-index:90!important;margin:0 0 24px!important;padding:14px 20px 25px!important;border:2px solid #d7aa35!important;border-radius:25px!important;background:radial-gradient(circle at 50% 50%,rgba(224,171,42,.14),transparent 18%),linear-gradient(115deg,#020b14,#082039 35%,#020910 50%,#082039 65%,#020b14)!important;box-shadow:0 12px 30px rgba(0,0,0,.4),inset 0 0 28px rgba(35,88,125,.16)!important;overflow:visible!important}.competition-livebar.lux-score.show{display:grid!important;grid-template-columns:1fr 105px 1fr!important;gap:14px!important;align-items:center!important}.lux-score .team-live{display:grid!important;align-items:center!important;gap:15px!important;text-align:center!important}.lux-score .scic{grid-template-columns:145px 1fr!important}.lux-score .wwic{grid-template-columns:1fr 145px!important}.lux-score .brand-logo{display:block;max-width:100%;max-height:135px;object-fit:contain;margin:auto}.lux-score .sc-logo{background:transparent!important;filter:drop-shadow(0 0 8px rgba(241,192,60,.25))}.lux-score .ww-logo-wrap{width:132px;height:132px;border-radius:50%;background:#fff;border:4px solid #d9aa35;padding:8px;display:flex;align-items:center;justify-content:center;box-shadow:0 0 8px #e8ba42,0 0 20px rgba(232,186,66,.28)}.lux-score .ww-logo{background:#fff!important;max-height:110px}.lux-score .score-core i{display:block;color:#f4c13b;font-style:normal;font-size:25px;line-height:1}.lux-score .score-core b{display:block!important;color:#f1c047!important;font-size:22px!important;line-height:1.05!important;text-shadow:0 0 7px rgba(241,192,71,.35)!important}.lux-score .score-core small{display:block;color:#f5c544;letter-spacing:3px;font-size:10px;margin-top:4px}.lux-score .score-core strong{display:block!important;color:#fff2ce!important;font-size:44px!important;line-height:1!important;margin:4px 0 0!important;text-shadow:0 0 8px rgba(255,213,91,.65)!important}.lux-score .score-core em{display:block;color:#eebd3e;font-style:normal;font-size:10px;font-weight:800;letter-spacing:2px;margin:1px 0 5px}.lux-score .comp-vs{width:78px;height:78px;margin:auto!important;padding:0!important;border-radius:50%!important;display:flex;align-items:center;justify-content:center;background:radial-gradient(circle,#10202e 35%,#02080d 72%)!important;border:4px solid #d8a932!important;box-shadow:0 0 0 4px #07111a,0 0 18px #d4a42f!important;color:#f4c64d!important;font-size:27px!important;font-weight:900!important}.lux-score .comp-meter{height:12px!important;margin-top:0!important;background:#020b14!important;border:1px solid #b88c28!important;border-radius:20px!important;overflow:hidden!important}.lux-score .comp-meter span{height:100%!important;border-radius:20px!important;transition:width .5s ease!important}.lux-score .scic .comp-meter span{background:linear-gradient(90deg,#c68d10,#ffe477,#dca421)!important;box-shadow:0 0 10px rgba(255,211,70,.45)!important}.lux-score .wwic .comp-meter span{background:linear-gradient(90deg,#1268dd,#31b9ff,#83ddff)!important;box-shadow:0 0 10px rgba(49,183,255,.55)!important}.lux-score .live-tag{position:absolute;left:50%;bottom:-17px;transform:translateX(-50%);min-width:255px;text-align:center;padding:7px 16px;border:2px solid #d5a42e;border-radius:20px;background:#030d16;color:#efbd3c;font-size:12px;font-weight:900;letter-spacing:1px;box-shadow:0 6px 14px rgba(0,0,0,.4)}@media(max-width:900px){.competition-livebar.lux-score.show{grid-template-columns:1fr 60px 1fr!important;padding:10px 8px 20px!important}.lux-score .scic,.lux-score .wwic{display:flex!important;flex-direction:column!important;gap:4px!important}.lux-score .wwic .score-core{order:2}.lux-score .wwic .ww-logo-wrap{order:1}.lux-score .brand-logo{max-height:65px}.lux-score .ww-logo-wrap{width:64px;height:64px;padding:3px;border-width:2px}.lux-score .ww-logo{max-height:52px}.lux-score .score-core i,.lux-score .score-core small,.lux-score .score-core em{display:none}.lux-score .score-core b{font-size:12px!important}.lux-score .score-core strong{font-size:25px!important}.lux-score .comp-vs{width:48px;height:48px;font-size:16px!important;border-width:2px!important}.lux-score .comp-meter{height:7px!important}.lux-score .live-tag{min-width:165px;font-size:8px;padding:4px 8px;bottom:-12px}}\n</style>\n<link rel="manifest" href="/manifest.webmanifest">\n<meta name="theme-color" content="#061629">\n<meta name="mobile-web-app-capable" content="yes">\n<meta name="apple-mobile-web-app-capable" content="yes">\n<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">\n<meta name="apple-mobile-web-app-title" content="THE IMMIGRATION MANAGEMENT SYSTEM">\n<link rel="apple-touch-icon" href="/static/pwa/icon-192.png">\n<style>\n.mobile-nav-btn,.mobile-nav-overlay{display:none}\n@media(max-width:900px){\n html,body{width:100%;max-width:100%;overflow-x:hidden}\n body{padding-top:env(safe-area-inset-top);padding-bottom:env(safe-area-inset-bottom)}\n .layout{display:block!important;min-height:100dvh}\n aside{position:fixed!important;left:0;top:0;bottom:0;z-index:1000002;width:min(86vw,320px);height:100dvh;overflow-y:auto;transform:translateX(-105%);transition:transform .25s ease;box-shadow:12px 0 40px rgba(0,0,0,.55);padding-top:calc(18px + env(safe-area-inset-top))}\n body.mobile-nav-open aside{transform:translateX(0)}\n .mobile-nav-overlay{display:block;position:fixed;inset:0;z-index:1000001;background:rgba(0,0,0,.62);opacity:0;visibility:hidden;transition:.25s ease}\n body.mobile-nav-open .mobile-nav-overlay{opacity:1;visibility:visible}\n main{padding:12px!important;padding-top:70px!important;width:100%;max-width:100vw}\n .mobile-nav-btn{display:flex;position:fixed;left:12px;top:calc(10px + env(safe-area-inset-top));z-index:1000003;align-items:center;gap:8px;min-height:46px;padding:8px 13px;border-radius:12px;border:1px solid #e6b73f;background:#071d32;color:#fff;font:700 16px "Times New Roman",serif;box-shadow:0 8px 24px rgba(0,0,0,.35)}\n .mobile-nav-btn b{color:#e6b73f;font-size:22px}\n .title{font-size:27px!important;line-height:1.08}\n .card{padding:13px!important;border-radius:12px}\n .grid,.grid3,.metrics,.kpi-grid,.performance-grid,.rank-grid,.donut-grid,.company-compare,.analytics-grid,.company-metrics,.filterbar,.paygrid,.summary-strip{grid-template-columns:1fr!important}\n .btn,.toolbtn,button,input,select,textarea{min-height:44px}\n input,select,textarea{font-size:16px!important}\n .btn,.toolbtn{padding:10px 12px!important}\n .tablewrap{overflow-x:auto!important;-webkit-overflow-scrolling:touch;border-radius:10px;max-width:100%}\n table{min-width:760px}\n th,td{padding:8px 7px;font-size:14px}\n .priority-strip{display:grid!important;grid-template-columns:auto 1fr;gap:8px!important}\n .priority-action{grid-column:1/-1;text-align:center}\n .competition-live-bar{left:8px!important;right:8px!important;top:8px!important}\n .celebrate-stage{max-height:92dvh;overflow-y:auto}\n .celebrate-details{grid-template-columns:1fr!important}\n .report-print-header,.report-print-footer{display:none!important}\n}\n@media(max-width:480px){\n main{padding-left:9px!important;padding-right:9px!important}\n .card{padding:11px!important}\n .title{font-size:24px!important}\n .metric b,.kpi-card .kpi-value{font-size:24px!important}\n .global-tools .toolbtn,.global-tools .btn{flex:1;text-align:center}\n table{min-width:700px}\n}\n@media(display-mode:standalone){\n body{overscroll-behavior-y:none}\n .mobile-nav-btn{top:calc(8px + env(safe-area-inset-top))}\n}\n</style>\n<style>.dashboard-identity-card{display:none!important}.dashboard-heading-row{display:block!important}.dashboard-heading-row h1{margin-bottom:8px!important}</style>\n<style>\n.month-overview-top{margin:8px 0 14px}\n.month-main-kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:12px}\n.month-kpi{background:#0d3151;border:1px solid #315f85;border-radius:14px;padding:15px 18px;box-shadow:0 7px 20px rgba(0,0,0,.18)}\n.month-kpi span{font-weight:800}.month-kpi b{display:block;color:#f3bd2f;font-size:32px;line-height:1.05;margin-top:6px}.month-kpi small{opacity:.65}\n.am-month-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:10px}\n.am-month-card{background:linear-gradient(145deg,#0a2945,#0d3557);border:1px solid #315f85;border-radius:14px;padding:13px;box-shadow:0 6px 20px rgba(0,0,0,.18)}\n.am-card-head{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}.am-card-name{font-size:18px;font-weight:900;color:#f3bd2f}.am-card-head small{opacity:.7}\n.am-conversion{font-size:25px;font-weight:900;color:#55dca0;white-space:nowrap}\n.am-card-stats{display:grid;grid-template-columns:1.5fr .7fr .7fr;gap:7px;margin-top:12px}\n.am-card-stats>div{background:#071d32;border:1px solid #244866;border-radius:9px;padding:8px}.am-card-stats span{display:block;font-size:11px;opacity:.7}.am-card-stats b{font-size:17px}\n.am-progress{height:7px;background:#071629;border-radius:8px;overflow:hidden;margin-top:10px}.am-progress i{display:block;height:100%;background:#55dca0;border-radius:8px}\n.am-note{display:block;margin-top:7px;opacity:.58}\n@media(max-width:900px){.month-main-kpis{grid-template-columns:1fr 1fr 1fr}.am-card-stats{grid-template-columns:1fr 1fr 1fr}}\n@media(max-width:600px){.month-main-kpis{grid-template-columns:1fr}.am-month-grid{grid-template-columns:1fr}}\n</style><style>\n.pro-dash-nav{display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap;margin:2px 0 12px}\n.pro-dash-nav a{position:relative;padding:9px 13px;border:1px solid #315f85;border-radius:10px;background:#0b2946;color:#fff;text-decoration:none;font-weight:800}\n.pro-dash-nav a:hover,.pro-dash-nav a.active{border-color:#e6b73f}.pro-badge{position:absolute;right:-6px;top:-8px;background:#d64755;color:#fff;border-radius:12px;min-width:20px;height:20px;padding:0 5px;display:flex;align-items:center;justify-content:center;font-size:11px}\n.pro-dashboard-hero{border:1px solid #315f85;border-radius:18px;overflow:hidden;background:#0b2946;margin-bottom:14px;box-shadow:0 10px 28px rgba(0,0,0,.2)}\n.pro-cover{height:150px;position:relative;background:linear-gradient(135deg,#0d3151,#163b5a);background-size:cover;background-position:center}.pro-cover.theme-flowers{background:radial-gradient(circle at 20% 30%,#ffd2e8 0 7%,transparent 8%),radial-gradient(circle at 75% 35%,#ffeaa6 0 7%,transparent 8%),linear-gradient(135deg,#7d375f,#233a63)}.pro-cover.theme-nature{background:linear-gradient(135deg,#416f54,#9bc78b)}.pro-cover.theme-minimal{background:linear-gradient(135deg,#dfe7ef,#f7f8fa)}.pro-cover.theme-ocean{background:linear-gradient(135deg,#0b4a6f,#20a6c9,#7ad8dc)}.pro-cover.theme-sunset{background:linear-gradient(135deg,#7f3765,#e3765b,#f4c06a)}\n.pro-edit-cover{position:absolute;right:14px;top:14px;background:rgba(5,22,39,.82);border:1px solid #e6b73f;color:#fff!important;border-radius:9px;padding:7px 10px;text-decoration:none;font-weight:800}\n.pro-profile-body{display:grid;grid-template-columns:130px 1fr auto;gap:16px;align-items:end;padding:0 20px 18px}.pro-profile-photo-wrap{margin-top:-62px}.pro-profile-photo,.pro-profile-photo-ph{width:122px;height:122px;border-radius:50%;border:5px solid #071629;background:#0b2946}.pro-profile-photo{object-fit:cover}.pro-profile-photo-ph{display:flex;align-items:center;justify-content:center;font-size:46px;font-weight:900;color:#e6b73f}.pro-profile-name{font-size:28px;font-weight:900;color:#e6b73f}.pro-profile-headline{font-size:17px;font-weight:700;margin-top:3px}.pro-profile-meta{display:flex;gap:7px;flex-wrap:wrap;margin-top:8px}.pro-profile-meta span{border:1px solid #315f85;background:#071d32;border-radius:16px;padding:4px 8px;font-size:12px}.pro-profile-actions{display:flex;gap:8px;flex-wrap:wrap}\n.pro-dashboard-title{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.due-followup-chip{align-items:center;gap:5px;padding:8px 11px;border:1px solid #d64755;background:#351820;color:#fff;border-radius:18px;font-weight:900}\n.desktop-reminder-stack{position:fixed;right:20px;bottom:20px;z-index:1000010;width:min(365px,calc(100vw - 24px));display:flex;flex-direction:column;gap:10px;pointer-events:none}\n.desktop-reminder{pointer-events:auto;position:relative;background:#1d2228;color:#fff;border:1px solid #4b535c;border-radius:12px;padding:13px 14px;box-shadow:0 16px 40px rgba(0,0,0,.48);transform:translateX(115%);opacity:0;transition:.28s ease}.desktop-reminder.visible{transform:translateX(0);opacity:1}.desktop-reminder.overdue{border-left:4px solid #d64755}.desktop-reminder-close{position:absolute;right:8px;top:6px;background:transparent;border:0;color:#aaa;font-size:20px;cursor:pointer}.desktop-reminder-app{font-size:12px;color:#9dd0ff;font-weight:800;margin-bottom:8px}.desktop-reminder-person{font-size:18px;font-weight:900}.desktop-reminder-text{margin-top:5px;color:#d8dee5;line-height:1.4}.desktop-reminder-actions{display:flex;gap:8px;margin-top:10px}.desktop-reminder-actions a,.desktop-reminder-actions button{border:1px solid #66717d;background:#2b333c;color:#fff;border-radius:8px;padding:7px 10px;text-decoration:none;cursor:pointer;font-weight:700}\nbody.user-light{background:#eef2f6!important;color:#132238!important}body.user-light main{background:#eef2f6!important}body.user-light .pro-dashboard-hero,body.user-light .card,body.user-light .metric,body.user-light .am-month-card{background:#fff!important;color:#132238!important;border-color:#cbd6e2!important}body.user-light .pro-profile-meta span{background:#f3f6f8!important;color:#132238!important}\n@media(max-width:900px){.pro-dash-nav{justify-content:flex-start;overflow-x:auto;flex-wrap:nowrap;padding-bottom:3px}.pro-dash-nav a{white-space:nowrap}.pro-profile-body{grid-template-columns:100px 1fr}.pro-profile-photo,.pro-profile-photo-ph{width:92px;height:92px}.pro-profile-photo-wrap{margin-top:-46px}.pro-profile-actions{grid-column:1/-1}.pro-cover{height:120px}.desktop-reminder-stack{right:8px;bottom:84px;width:calc(100vw - 16px)}}\n</style>\n<style>\n.today-report-wrap.luxury-today-report{border:2px solid #d7aa35;border-radius:24px;background:radial-gradient(circle at 50% 58%,rgba(224,171,42,.12),transparent 22%),linear-gradient(115deg,#020b14,#082039 35%,#020910 50%,#082039 65%,#020b14);padding:14px 18px 18px;margin:8px 0 18px;box-shadow:0 12px 30px rgba(0,0,0,.38),inset 0 0 30px rgba(35,88,125,.16);overflow:hidden}\n.today-report-head{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;border-bottom:1px solid rgba(230,183,63,.35);padding:0 0 10px;margin-bottom:12px}.today-report-head>div:first-child{display:flex;align-items:center;gap:8px;color:#f0bd43;font-size:16px}.today-report-head small{color:#dce8f2;opacity:.72}.today-report-note{font-size:12px;opacity:.68}.today-live-dot{width:9px;height:9px;border-radius:50%;background:#55dca0;box-shadow:0 0 0 4px rgba(85,220,160,.12),0 0 9px rgba(85,220,160,.7);display:inline-block}\n.today-company-grid{display:grid;grid-template-columns:1fr 90px 1fr;gap:14px;align-items:stretch}.today-company-grid.single{grid-template-columns:1fr;max-width:760px;margin:auto}.today-company-card{position:relative;background:linear-gradient(145deg,#08213a,#061626);border:1px solid #9a7626;border-radius:18px;padding:13px 14px 14px;box-shadow:inset 0 0 20px rgba(34,93,133,.12),0 6px 18px rgba(0,0,0,.18)}\n.today-brand-row{display:grid;grid-template-columns:105px 1fr;gap:12px;align-items:center;margin-bottom:10px}.today-brand-row-right{grid-template-columns:1fr 105px}.today-logo{display:flex;align-items:center;justify-content:center;width:96px;height:96px;margin:auto}.today-logo img{max-width:100%;max-height:100%;object-fit:contain;display:block}.today-logo-scic{background:transparent;filter:drop-shadow(0 0 7px rgba(241,192,60,.25))}.today-logo-wwic{background:#fff;border:3px solid #d9aa35;border-radius:50%;padding:6px;box-shadow:0 0 8px #e0b43d,0 0 18px rgba(232,186,66,.22)}.today-logo-wwic img{background:#fff;border-radius:50%}\n.today-brand-title{text-align:center}.today-brand-title i{display:block;color:#f4c13b;font-style:normal;font-size:24px;line-height:1}.today-company-name{text-align:center;color:#f0bd43;font-weight:900;letter-spacing:.5px;font-size:20px;line-height:1.05}.today-brand-title small{display:block;color:#f5c544;letter-spacing:3px;font-size:9px;margin-top:4px}\n.today-stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px}.today-stat-grid>div{background:linear-gradient(180deg,#06182a,#04121f);border:1px solid #315a78;border-radius:12px;padding:10px;text-align:center;box-shadow:inset 0 0 12px rgba(22,74,109,.14)}.today-stat-grid span{display:block;font-size:12px;opacity:.74}.today-stat-grid b{display:block;color:#fff4d2;font-size:25px;margin-top:4px;text-shadow:0 0 7px rgba(255,215,93,.35)}\n.today-vs{display:flex;align-items:center;justify-content:center;width:72px;height:72px;border-radius:50%;align-self:center;justify-self:center;background:radial-gradient(circle,#10202e 35%,#02080d 72%);border:4px solid #d8a932;box-shadow:0 0 0 4px #07111a,0 0 18px #d4a42f;color:#f4c64d;font-size:24px;font-weight:900;text-shadow:0 0 7px rgba(244,198,77,.45)}\n.today-total-strip{display:flex;justify-content:center;gap:20px;flex-wrap:wrap;margin-top:12px;padding-top:10px;border-top:1px solid rgba(230,183,63,.25)}.today-total-strip span{opacity:.72}.today-total-strip b{color:#f0bd43;text-shadow:0 0 5px rgba(240,189,67,.18)}\n@media(max-width:900px){.today-company-grid{grid-template-columns:1fr 58px 1fr;gap:8px}.today-brand-row,.today-brand-row-right{display:flex;flex-direction:column;gap:4px}.today-logo{width:62px;height:62px}.today-logo-wwic{padding:3px;border-width:2px}.today-brand-title i,.today-brand-title small{display:none}.today-company-name{font-size:13px}.today-stat-grid b{font-size:20px}.today-vs{width:48px;height:48px;font-size:16px;border-width:2px}}\n@media(max-width:700px){.today-company-grid{grid-template-columns:1fr}.today-vs{margin:auto}.today-stat-grid{grid-template-columns:1fr 1fr}.today-brand-row-right{flex-direction:column-reverse}}\n@media(max-width:450px){.today-stat-grid{grid-template-columns:1fr}.today-report-note{display:none}}\n</style>\n\\n<style>\\n\n.md-standalone-score-hidden{display:none!important}\n.md-merged-performance{border:2px solid #d7aa35;border-radius:26px;padding:14px 18px 18px;background:radial-gradient(circle at 50% 45%,rgba(226,173,42,.14),transparent 18%),linear-gradient(115deg,#020b14,#082039 35%,#020910 50%,#082039 65%,#020b14);box-shadow:0 12px 32px rgba(0,0,0,.42),inset 0 0 32px rgba(35,88,125,.16)}\n.md-merge-head{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;padding-bottom:10px;border-bottom:1px solid rgba(224,177,57,.35)}.md-merge-head>div:first-child{display:flex;align-items:center;gap:8px;color:#f1c048}.md-merge-head small{color:#c8d6e4}.md-merge-note{font-size:12px;opacity:.7}\n.md-merge-grid{display:grid;grid-template-columns:1fr 100px 1fr;align-items:center;gap:14px;padding:13px 0 8px}.md-team{display:grid;align-items:center;gap:16px}.md-team-scic{grid-template-columns:145px 1fr}.md-team-wwic{grid-template-columns:1fr 145px}\n.md-team-logo{width:136px;height:136px;display:flex;align-items:center;justify-content:center;margin:auto}.md-team-logo img{max-width:100%;max-height:100%;object-fit:contain}.md-team-logo-scic{background:transparent}.md-team-logo-wwic{background:#fff;border:4px solid #d9aa35;border-radius:50%;padding:8px;box-shadow:0 0 8px #e8ba42,0 0 22px rgba(232,186,66,.3)}\n.md-team-content{text-align:center}.md-crown{font-size:25px;color:#f4c13b}.md-team h3{margin:1px 0!important;color:#f2c34a!important;font-size:22px!important}.md-stars{font-size:10px;color:#f4c543;letter-spacing:4px}.md-live-score{font-size:39px;font-weight:900;color:#fff2cf;text-shadow:0 0 9px rgba(255,215,93,.55)}.md-live-label{font-size:9px;color:#eabb3c;letter-spacing:1.6px;font-weight:800;margin:1px 0 5px}\n.md-meter{height:10px;background:#020b14;border:1px solid #a77e22;border-radius:18px;overflow:hidden}.md-meter span{display:block;height:100%;background:linear-gradient(90deg,#bf8610,#ffe679,#dba221)}.md-meter-blue span{background:linear-gradient(90deg,#1165d8,#32b9ff,#84ddff)}\n.md-today-stats{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:9px}.md-today-stats>div{background:#061829;border:1px solid #315a78;border-radius:11px;padding:8px}.md-today-stats span{display:block;font-size:10px;opacity:.78}.md-today-stats b{display:block;font-size:20px;color:#fff4d2}\n.md-merge-vs{width:76px;height:76px;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:auto;background:#07111a;border:4px solid #d8a932;box-shadow:0 0 18px #d4a42f;color:#f4c64d;font-size:25px;font-weight:900}\n.md-merge-footer{display:flex;justify-content:center;gap:18px;flex-wrap:wrap;padding-top:9px;border-top:1px solid rgba(230,183,63,.26)}.md-merge-footer b{color:#f1bd3c}\n@media(max-width:900px){.md-merge-grid{grid-template-columns:1fr 60px 1fr}.md-team-scic{grid-template-columns:90px 1fr}.md-team-wwic{grid-template-columns:1fr 90px}.md-team-logo{width:82px;height:82px}.md-live-score{font-size:28px}.md-team h3{font-size:16px!important}.md-merge-vs{width:50px;height:50px;font-size:17px}}\n@media(max-width:700px){.md-merge-grid{grid-template-columns:1fr}.md-team-scic,.md-team-wwic{display:flex;flex-direction:column;gap:5px}.md-team-wwic .md-team-content{order:2}.md-team-wwic .md-team-logo{order:1}.md-merge-note{display:none}}\n\\n</style>\\n\n<style>\n\n/* v3.98 Smart unread update toast */\n.update-toast{position:fixed;right:18px;bottom:92px;z-index:1002;width:min(390px,calc(100vw - 28px));display:grid;grid-template-columns:42px 1fr;gap:10px;padding:13px 38px 13px 13px;border:1px solid #c79a2d;border-radius:16px;background:linear-gradient(135deg,#24141d 0%,#0b2943 100%);box-shadow:0 14px 34px rgba(0,0,0,.42),0 0 18px rgba(225,178,52,.14);animation:updateToastIn .22s ease-out}\n.update-toast.urgent{border-color:#e0ad2e;box-shadow:0 14px 34px rgba(0,0,0,.44),0 0 22px rgba(225,173,46,.22)}\n.update-toast.closing{opacity:0;transform:translateY(8px);transition:.18s ease}\n.update-toast-icon{width:38px;height:38px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#0c3150;border:1px solid #315f80;font-size:20px}\n.update-toast-kicker{font-size:9px;font-weight:900;letter-spacing:1.8px;color:#edbd43;margin-bottom:2px}\n.update-toast-title{font-size:14px;font-weight:900;color:#ffd46a;line-height:1.15;padding-right:4px}\n.update-toast-text{font-size:12px;color:#eef4fa;opacity:.9;line-height:1.25;margin-top:3px;max-height:46px;overflow:hidden}\n.update-toast-open{display:inline-block;margin-top:8px;padding:6px 10px;border-radius:9px;border:1px solid #39739a;background:#08243c;color:#fff!important;font-size:12px;font-weight:800;text-decoration:none!important}\n.update-toast-close{position:absolute;right:8px;top:7px;width:25px;height:25px;border:0!important;border-radius:50%!important;background:transparent!important;color:#d9e5ee!important;font-size:20px;line-height:20px;padding:0!important;cursor:pointer}\n.update-toast-close:hover{background:rgba(255,255,255,.08)!important}\n@keyframes updateToastIn{from{opacity:0;transform:translateY(12px) scale(.98)}to{opacity:1;transform:none}}\n@media(max-width:600px){.update-toast{right:10px;bottom:82px;width:calc(100vw - 20px)}}\n\n</style>\n\n<style>\n\n/* v3.98 global report controls */\n.global-tools{display:flex!important;gap:8px!important;flex-wrap:wrap!important;align-items:center!important}\n.global-tools .toolbtn{white-space:nowrap}\n@media(max-width:700px){.global-tools{gap:6px!important}.global-tools .toolbtn{font-size:12px!important;padding:8px 9px!important}}\n\n</style>\n\n<style>\n\n/* =========================================================\n   v3.98 GM Dashboard — SAFE CSS-ONLY MERGED DISPLAY\n   Uses the two already-working blocks and visually combines them.\n   ========================================================= */\n.competition-livebar.gm-css-merged-score{\n  margin-bottom:0!important;\n  border-radius:26px 26px 0 0!important;\n  border-bottom-color:rgba(215,170,53,.55)!important;\n  padding-bottom:22px!important;\n  box-shadow:0 12px 26px rgba(0,0,0,.28),inset 0 0 28px rgba(35,88,125,.14)!important;\n}\n\n/* The GM Today's Report following the live scoreboard becomes the lower half\n   of the SAME luxury module. The ~ selector is deliberate so it still works\n   even if another harmless element exists between the two blocks. */\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report{\n  margin-top:0!important;\n  margin-bottom:20px!important;\n  border-top:0!important;\n  border-radius:0 0 26px 26px!important;\n  background:\n    radial-gradient(circle at 50% 16%,rgba(224,171,42,.09),transparent 24%),\n    linear-gradient(115deg,#020b14,#082039 35%,#020910 50%,#082039 65%,#020b14)!important;\n  box-shadow:0 14px 28px rgba(0,0,0,.30),inset 0 0 28px rgba(35,88,125,.12)!important;\n  padding:18px 22px 20px!important;\n  overflow:visible!important;\n}\n\n/* Compact integrated TODAY header */\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-report-head{\n  max-width:980px!important;\n  margin:0 auto 12px!important;\n  padding:2px 4px 10px!important;\n  border-bottom:1px solid rgba(230,183,63,.25)!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-report-head>div:first-child{\n  margin:auto!important;\n  font-size:15px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-report-note{\n  display:none!important;\n}\n\n/* Make the GM company card modular, centered and wider */\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-company-grid.single{\n  width:100%!important;\n  max-width:980px!important;\n  margin:0 auto!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-company-card{\n  display:grid!important;\n  grid-template-columns:minmax(220px,34%) minmax(0,1fr)!important;\n  gap:18px!important;\n  align-items:center!important;\n  min-height:170px!important;\n  padding:16px 18px!important;\n  border-radius:18px!important;\n  border:1px solid #8f7028!important;\n  background:linear-gradient(145deg,#08213a,#061626)!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-brand-row{\n  display:flex!important;\n  flex-direction:column!important;\n  justify-content:center!important;\n  gap:6px!important;\n  margin:0!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-logo{\n  width:90px!important;\n  height:90px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-company-name{\n  font-size:19px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid{\n  grid-template-columns:1fr 1fr!important;\n  gap:12px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid>div{\n  min-height:92px!important;\n  display:flex!important;\n  flex-direction:column!important;\n  justify-content:center!important;\n  border-radius:14px!important;\n  padding:12px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid b{\n  font-size:28px!important;\n}\n\n/* LIVE PERFORMANCE pill sits exactly on the join and makes the two halves read as one module */\n.competition-livebar.gm-css-merged-score .lux-live-badge{\n  bottom:-18px!important;\n  z-index:4!important;\n}\n\n@media(max-width:850px){\n  .competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-company-card{\n    grid-template-columns:1fr!important;\n  }\n  .competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid{\n    grid-template-columns:1fr 1fr!important;\n  }\n}\n@media(max-width:560px){\n  .competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report{\n    padding:16px 9px 14px!important;\n  }\n  .competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid{\n    grid-template-columns:1fr!important;\n  }\n}\n\n</style>\n<style>\n.bin397{background:linear-gradient(180deg,#b92e3d,#791a27)!important;border:1px solid #e45a69!important;color:#fff!important;font-weight:900!important}\n.recycle397{display:inline-flex!important;align-items:center!important;justify-content:center!important;background:#0b3328!important;border:1px solid #42c985!important;color:#a9f2cb!important;text-decoration:none!important;font-weight:900!important}\n.md-recycle397{border:2px solid #42c985!important;box-shadow:0 0 22px rgba(66,201,133,.10)!important;margin-top:18px!important}.md-recycle397 h2{color:#72e6a9!important}\n</style></head>\n\n<body>\n<div class="layout">\n<aside>\n{% if current_user and current_user['role']=='MD' %}\n<div class="brand">THE GAUR</div><div>Immigration CRM <span style="color:#e6b73f">v3.36</span></div>\n{% elif current_user and current_user['company_code']=='SCIC' %}\n<div class="company-logo-wrap"><img class="company-logo" src="" alt="ABC"></div><div style="text-align:center">Immigration CRM <span style="color:#e6b73f">v3.36</span></div>\n{% elif current_user and current_user['company_code']=='WWIC' %}\n<div class="company-logo-wrap wwic"><img class="company-logo" src="" alt="XYZ"></div><div style="text-align:center">Immigration CRM <span style="color:#e6b73f">v3.36</span></div>\n{% else %}<div class="brand">THE GAUR</div>{% endif %}\n{% if current_user %}\n<div class="userbox">\n  <div>\n    {% if current_user['photo_mime'] %}\n      <img class="user-avatar" src="{{ url_for('user_photo', user_id=current_user['id']) }}" alt="User">\n    {% else %}\n      <div class="user-avatar-placeholder">{{ (current_user['full_name'] or '?')[0]|upper }}</div>\n    {% endif %}\n  </div>\n  <div><b>{{ role_display(current_user['role'], current_user['designation'] if current_user['designation'] else '') }}</b>{% if current_user['role']!='MD' %}<br>{{ role_display(current_user['role'],current_user['designation']) }}{% endif %}{% if current_user['company_code'] %}<br><span style="opacity:.82">{{ current_user['company_code'] }}</span>{% endif %}\n  <br><a class="profile-mini-link" href="{{url_for('my_profile')}}">My Profile</a></div>\n</div>\n<nav class="modular-nav">\n<a class="nav-main-link" href="{{ url_for('dashboard') }}"><span class="nav-num">1</span><span>Dashboard</span></a>\n\n{% if current_user['role'] in ['MD','GM'] %}\n<details class="nav-group">\n  <summary><span class="nav-num">2</span><span>Employee Details</span><span class="nav-chevron">›</span></summary>\n  <div class="nav-submenu">\n    <a href="{{ url_for('ams') }}">Assistant Manager</a>\n    <a href="{{ url_for('team_leaders') }}">Team Leaders</a>\n    <a href="{{ url_for('counsellors') }}">Counsellors</a>\n    <a href="{{ url_for('telecallers') }}">Telecallers</a>\n  </div>\n</details>\n\n<a class="nav-main-link" href="{{ url_for('cases') }}"><span class="nav-num">3</span><span>{% if u.role=="MD" %}Enrolment Payment Report{% else %}Enrolment Report{% endif %}</span></a>\n{% if current_user and current_user['role'] in ('MD','GM') %}<details class="nav-group">\n  <summary><span class="nav-num">4</span><span>Accounts</span><span class="nav-chevron">›</span></summary>\n  <div class="nav-submenu">\n    <a href="{{ url_for('finance_banks') }}">🏦 Bank Manager</a>\n    <a href="{{ url_for('finance_center') }}">📒 Daily Passbook</a>\n    <a href="{{ url_for('finance_daily_report') }}">📊 Daily Report</a>\n  </div>\n</details>{% endif %}\n\n<details class="nav-group">\n  <summary><span class="nav-num">5</span><span>Performance Analysis</span><span class="nav-chevron">›</span></summary>\n  <div class="nav-submenu">\n    <a href="{{ url_for('performance_directory') }}">Performance Dashboard</a><a href="{{ url_for('competition_center') }}">🔥 Live Competition Arena</a>\n    <a href="{{ url_for('fb_leads.home') }}">&#128178; Facebook Leads</a>\n    <a href="{{ url_for('reception') }}">Reception</a>\n    <a href="{{ url_for('performance_directory', designation='HR Manager') }}">HR Managers</a>\n  </div>\n</details>\n\n<details class="nav-group">\n  <summary><span class="nav-num">6</span><span>Lead Upload & Allocation</span><span class="nav-chevron">›</span></summary>\n  <div class="nav-submenu">\n    <a href="{{ url_for('leads') }}">Lead Upload</a>\n    <a href="{{ url_for('lead_control_center') }}">Lead Allocation</a>\n    <a href="{{ url_for('lead_control_center', batch='latest') }}">Batch Control</a>\n    <a href="{{ url_for('cases') }}">Filing Department</a>\n    <a href="{{ url_for('attendance_center') }}">Employee Attendance</a>\n  </div>\n</details>\n\n<a class="nav-main-link" href="{{ url_for('cases') }}"><span class="nav-num">7</span><span>Filing Department</span></a>\n<a class="nav-main-link" href="{{ url_for('performance_directory', designation='General Manager') }}"><span class="nav-num">8</span><span>GM Report</span></a>\n<a class="nav-main-link" href="{{ url_for('management_reporting') }}"><span class="nav-icon">📊</span><span>Management Reporting</span></a>\n\n<details class="nav-group nav-utility">\n  <summary><span class="nav-icon">⚙</span><span>Security</span><span class="nav-chevron">›</span></summary>\n  <div class="nav-submenu">\n    <a href="{{ url_for('activity_center') }}">Activity & Security Center</a>\n    \n  </div>\n</details>\n\n{% elif current_user['role']=='RECEPTION' %}\n<a class="nav-main-link" href="{{ url_for('reception') }}"><span class="nav-num">2</span><span>Reception</span></a>\n<a class="nav-main-link" href="{{ url_for('attendance_center') }}"><span class="nav-num">3</span><span>Employee Attendance</span></a>\n{% elif current_user['role']=='AM' %}\n<a class="nav-main-link" href="{{ url_for('my_leads') }}"><span class="nav-num">2</span><span>My Leads</span></a>\n<a class=\"nav-main-link\" href=\"{{ url_for('my_team') }}\"><span class=\"nav-num\">👥</span><span>My Team</span></a>\n<a class=\"nav-main-link\" href=\"{{ url_for("cases") }}\"><span class=\"nav-num\">3</span><span>📋 My Enrollments</span></a>\n{% endif %}\n\n{% if u.role=="MD" %}<a class="nav-main-link" href="{{ url_for('payment_accounts') }}"><span class="nav-icon">🏦</span><span>Company Payment Accounts</span></a><a class="nav-main-link" href="{{url_for('finance_cash_book')}}"><span class="nav-icon">📒</span><span>Daily Cash Book</span></a>{% endif %}\n<a class="nav-signout" href="{{ url_for('logout') }}">Sign Out</a>\n</nav>{% endif %}\n</aside>\n<div id="mobileNavOverlay" class="mobile-nav-overlay"></div>\n<main>\n<button type="button" id="mobileNavBtn" class="mobile-nav-btn" aria-label="Open navigation"><b>☰</b> Menu</button>\n<div class="report-print-header">\n  <div class="report-logo-box">\n    {% if report_brand.code=='SCIC' %}\n      <img class="report-logo" src="" alt="ABC">\n      <div class="report-company">ABC<small>Immigration Consultants</small></div>\n    {% elif report_brand.code=='WWIC' %}\n      <img class="report-logo" src="" alt="XYZ">\n      <div class="report-company">XYZ<small>Corporate Report</small></div>\n    {% else %}\n      <img class="report-logo dual" src="" alt="ABC">\n      <img class="report-logo dual" src="" alt="XYZ">\n      <div class="report-company">THE GAUR<small>Executive Consolidated Report</small></div>\n    {% endif %}\n  </div>\n  <div class="report-head-center">\n    <b id="professionalReportTitle">Detailed Management Report</b>\n    <span>{{report_brand.scope}}</span>\n  </div>\n  <div class="report-head-right">\n    <div><b>Report ID:</b> <span id="professionalReportId">—</span></div>\n    <div><b>Generated:</b> <span id="professionalReportStamp">—</span></div>\n  </div>\n</div>\n<div class="report-print-footer">\n  <div><span class="conf">CONFIDENTIAL • AUTHORIZED USE</span><br>{{report_brand.name}}</div>\n  <div class="footer-mid">{% if report_brand.address %}{{report_brand.address}}{% endif %}{% if report_brand.gstin %}<br>GSTIN: {{report_brand.gstin}}{% endif %}</div>\n  <div class="footer-right">Generated by THE IMMIGRATION MANAGEMENT SYSTEM<br><span id="professionalFooterReportId">—</span></div>\n</div>\n<div class="report-print-meta">\n  <div><b>Report</b><span id="professionalMetaTitle">Detailed Report</span></div>\n  <div><b>Company / Scope</b>{{report_brand.scope}}</div>\n  <div><b>Generated By</b>{{role_display(current_user['role']) if current_user else 'System'}}<br><small>{{role_display(current_user['role']) if current_user else ''}}</small></div>\n  <div><b>Generated On</b><span id="professionalMetaStamp">—</span></div>\n</div>\n{% if current_user %}\n<div class="global-tools no-print" id="managementQuickTools">\n  {% if current_user['role'] in ['MD','GM'] %}<a class="toolbtn" href="{{url_for('dashboard')}}">← Back</a>{% endif %}\n  <button type="button" class="toolbtn" id="safeWhatsAppBtn">WhatsApp Share</button>\n  <button type="button" class="toolbtn" id="safePrintBtn">Print / Save PDF</button>\n  {% if current_user['role'] in ['MD','GM'] %}<button type="button" class="toolbtn" id="safeExcelBtn">Download Excel</button>\n  <a class="toolbtn" href="{{url_for('management_reporting')}}">📊 Reporting</a>{% endif %}\n</div>\n{% endif %}\n{% if current_user %}\n<div id="competitionLiveBar" class="competition-livebar lux-score no-print{% if current_user and current_user['role']=='MD' and request.endpoint=='dashboard' %} md-standalone-score-hidden{% endif %}{% if current_user and current_user['role']=='GM' and request.endpoint=='dashboard' %} gm-css-merged-score{% endif %}"><div class="team-live scic"><img class="brand-logo sc-logo" src=""><div class="score-core"><i>♛</i><b>ABC</b><small>★ ★ ★ ★ ★</small><strong id="scicLiveScore">0</strong><em>SCORE</em><div class="comp-meter"><span id="scicMeter" style="width:0%"></span></div></div></div><div class="comp-vs">VS</div><div class="team-live wwic"><div class="score-core"><i>♛</i><b>XYZ</b><small>★ ★ ★ ★ ★</small><strong id="wwicLiveScore">0</strong><em>SCORE</em><div class="comp-meter"><span id="wwicMeter" style="width:0%"></span></div></div><div class="ww-logo-wrap"><img class="brand-logo ww-logo" src=""></div></div><div class="live-tag">▥ &nbsp; LIVE PERFORMANCE &nbsp; ▥</div></div>\n<div id="competitionPopup" class="competition-popup no-print" aria-live="assertive" aria-label="Enrollment celebration">\n  <div class="celebrate-glow"></div><div class="firework fw1"></div><div class="firework fw2"></div><div class="firework fw3"></div><div class="firework fw4"></div>\n  <div class="celebrate-stage">\n    <button type="button" class="celebrate-close" id="competitionPopupClose" title="Close">×</button>\n    <div class="celebrate-side side-left">🏆🎲</div><div class="celebrate-side side-right">🎯🏆</div>\n    <div class="celebrate-crown">👑</div>\n    <div id="competitionPopupTitle" class="comp-pop-title">CONGRATULATIONS!</div>\n    <div class="celebrate-subtitle">🎉 ENROLLMENT DONE! 🎉</div>\n    <div class="celebrate-person">\n      <div class="celebrate-avatar-wrap"><img id="competitionEmployeePhoto" class="celebrate-avatar" alt="Assistant Manager"><div id="competitionEmployeeFallback" class="celebrate-avatar-fallback" style="display:none">👤</div></div>\n      <div class="celebrate-person-info"><div id="competitionEmployeeName" class="celebrate-name">TEAM MEMBER</div><div id="competitionEmployeeRole" class="celebrate-role">ASSISTANT MANAGER</div><div id="competitionCompanyName" class="celebrate-company"></div></div>\n    </div>\n    <div class="celebrate-details">\n      <div class="celebrate-detail"><small>Client</small><b id="competitionClientName">—</b></div>\n      <div class="celebrate-detail"><small>Enrollment ID</small><b id="competitionCaseId">—</b></div>\n      <div class="celebrate-detail"><small>Amount Received</small><b id="competitionRevenue">₹0</b></div>\n    </div>\n    <div class="celebrate-save">✓ SUCCESSFULLY SAVED!</div>\n    <div id="competitionPopupText" style="position:relative;z-index:2;margin-top:9px;font-weight:bold"></div>\n    <div id="competitionPopupTime" class="celebrate-countdown"></div>\n  </div>\n</div>\n{% endif %}\n{% with messages = get_flashed_messages(with_categories=true) %}\n{% for cat,msg in messages %}<div class="flash {{cat}}">{{msg}}</div>{% endfor %}\n{% endwith %}\n\n\n{% block content %}{% endblock %}\n</main></div>\n\n\n{% if current_user %}\n<button class="notif-bell no-print" id="notifBell" title="Priority Notifications">🔔\n  {% if notification_unread_count %}<span class="notif-count">{{notification_unread_count}}</span>{% endif %}\n</button>\n<div class="notif-panel no-print" id="notifPanel">\n  <div class="notif-head">\n    <strong>Priority Updates</strong>\n    <button class="chat-close" type="button" onclick="toggleNotif(false)">×</button>\n  </div>\n  <div class="notif-filters">\n    <button class="notif-filter active" onclick="loadNotifications('all',this)" type="button">All</button>\n    <button class="notif-filter" onclick="loadNotifications('important',this)" type="button">Important</button>\n    <button class="notif-filter" onclick="loadNotifications('chat',this)" type="button">Messages</button>\n    <button class="notif-filter" onclick="loadNotifications('broadcast',this)" type="button">Broadcasts</button>\n  </div>\n  <div class="notif-list" id="notifList"><div style="opacity:.7;padding:12px">Loading updates...</div></div>\n  <div class="notif-footer">\n    <button class="toolbtn" type="button" onclick="markAllNotificationsRead()">Mark All Read</button>\n  </div>\n</div>\n<script>\nfunction toggleNotif(force){\n const p=document.getElementById('notifPanel'); if(!p)return;\n if(force===false){p.classList.remove('open');return;}\n p.classList.toggle('open'); if(p.classList.contains('open'))loadNotifications('all');\n}\ndocument.addEventListener('DOMContentLoaded',function(){\n const b=document.getElementById('notifBell'); if(b)b.addEventListener('click',()=>toggleNotif());\n});\nfunction nEsc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}\nasync function loadNotifications(filter,btn){\n document.querySelectorAll('.notif-filter').forEach(x=>x.classList.remove('active'));if(btn)btn.classList.add('active');\n const box=document.getElementById('notifList');box.innerHTML='<div style="opacity:.7;padding:12px">Loading...</div>';\n try{\n   const r=await fetch('{{url_for("notifications_api")}}?filter='+encodeURIComponent(filter));const d=await r.json();\n   box.innerHTML=d.items.map(n=>`<div class="notif-item ${n.unread?'unread':''} ${n.priority>=90?'urgent':''} ${n.kind==='LEAD_ALERT'?'lead':''} ${n.kind==='TARGET'?'target':''}" onclick="location.href='${n.url}'">\n     <div class="n-title">${n.priority>=90?'🚨 ':n.kind==='DIRECT_CHAT'?'💬 ':n.kind==='LEAD_ALERT'?'📌 ':'🔔 '}${nEsc(n.title)}</div>\n     <div class="n-body">${nEsc(n.text)}</div>\n     <div class="n-meta">${nEsc(n.sender||'THE IMMIGRATION MANAGEMENT SYSTEM')} • ${nEsc(n.created_at||'')}</div>\n   </div>`).join('')||'<div style="opacity:.7;padding:18px;text-align:center">No updates right now.</div>';\n }catch(e){box.innerHTML='<div style="color:#ff9a9a;padding:12px">Could not load notifications.</div>'}\n}\nasync function markAllNotificationsRead(){\n await fetch('{{url_for("notifications_mark_all")}}',{method:'POST'});\n loadNotifications('all');\n const badge=document.querySelector('.notif-count');if(badge)badge.remove();\n}\n</script>\n{% endif %}\n\n\n<script>\n(function(){\n  function clean(p){return (p||'/').replace(/\\/+$/,'')||'/';}\n  var current=clean(window.location.pathname);\n  var links=[].slice.call(document.querySelectorAll('aside a[href]'));\n  var best=null,bestLen=-1;\n  links.forEach(function(a){\n    try{\n      var p=clean(new URL(a.href,window.location.origin).pathname);\n      var match=(current===p)||(p!=='/' && current.indexOf(p+'/')===0);\n      if(match && p.length>bestLen){best=a;bestLen=p.length;}\n    }catch(e){}\n  });\n  if(best){\n    best.classList.add('nav-active');\n    const parent=best.closest('details.nav-group');\n    if(parent) parent.open=true;\n  }\n})();\n</script>\n<script>\n(function(){let last=Number(sessionStorage.getItem('gaurCompEvent')||0),first=true;const bar=document.getElementById('competitionLiveBar'),pop=document.getElementById('competitionPopup');if(!bar||!pop)return;\nfunction money(n){return '₹'+Number(n||0).toLocaleString('en-IN')}function show(e){pop.className='competition-popup no-print '+(e.company_code==='SCIC'?'scic':'wwic');document.getElementById('competitionPopupTitle').textContent=(e.company_code==='SCIC'?'ABC':'XYZ')+' SCORES! 🎉';document.getElementById('competitionPopupText').textContent=e.message||'New enrollment';document.getElementById('competitionPopupTime').textContent=e.created_at||'';requestAnimationFrame(()=>pop.classList.add('show'));setTimeout(()=>pop.classList.remove('show'),9000)}\nasync function run(){try{const r=await fetch('/api/competition/live?after='+last,{cache:'no-store'});if(!r.ok)return;const d=await r.json();if(!d.enabled){bar.classList.remove('show');return}bar.classList.add('show');const s=d.scores||{};document.getElementById('scicLiveScore').textContent=(s.ABC?.enrollments||0)+' • '+money(s.ABC?.revenue||0);document.getElementById('wwicLiveScore').textContent=(s.XYZ?.enrollments||0)+' • '+money(s.XYZ?.revenue||0);const st=Number(d.scic_target||0),wt=Number(d.wwic_target||0);document.getElementById('scicMeter').style.width=(st?Math.min(100,(s.ABC?.enrollments||0)/st*100):0)+'%';document.getElementById('wwicMeter').style.width=(wt?Math.min(100,(s.XYZ?.enrollments||0)/wt*100):0)+'%';if(first){last=Number(d.latest_id||0);sessionStorage.setItem('gaurCompEvent',last);first=false;return}if(d.popup_enabled&&d.events?.length)d.events.forEach((e,i)=>setTimeout(()=>show(e),i*1200));last=Number(d.latest_id||last);sessionStorage.setItem('gaurCompEvent',last)}catch(e){}}\nrun();setInterval(run,5000)})();\n</script><script>\n(function(){\n const btn=document.getElementById('mobileNavBtn');\n const overlay=document.getElementById('mobileNavOverlay');\n function closeNav(){document.body.classList.remove('mobile-nav-open')}\n if(btn)btn.addEventListener('click',()=>document.body.classList.toggle('mobile-nav-open'));\n if(overlay)overlay.addEventListener('click',closeNav);\n document.querySelectorAll('aside a').forEach(a=>a.addEventListener('click',closeNav));\n window.addEventListener('resize',()=>{if(window.innerWidth>900)closeNav()});\n if('serviceWorker' in navigator){\n   window.addEventListener('load',()=>navigator.serviceWorker.register('/service-worker.js').catch(()=>{}));\n }\n})();\n</script>\n\n<script>\n(function(){\n  function pageTitle(){\n    var t=document.querySelector('h1.title')||document.querySelector('main h1')||document.querySelector('main h2');\n    return (t?t.textContent:document.title).replace(/\\s+/g,' ').trim();\n  }\n\n  var wa=document.getElementById('safeWhatsAppBtn');\n  if(wa){\n    wa.addEventListener('click',function(){\n      var msg='THE IMMIGRATION MANAGEMENT SYSTEM\\nReport/Page: '+pageTitle()+'\\nLink: '+window.location.href;\n      window.open('https://wa.me/?text='+encodeURIComponent(msg),'_blank','noopener');\n    });\n  }\n\n  var pr=document.getElementById('safePrintBtn');\n  if(pr){\n    pr.addEventListener('click',function(){ window.print(); });\n  }\n\n  var ex=document.getElementById('safeExcelBtn');\n  if(ex){\n    ex.addEventListener('click',function(){\n      try{\n        var main=document.querySelector('main');\n        if(!main){ alert('No report content found on this page.'); return; }\n\n        var tables=[].slice.call(main.querySelectorAll('table')).filter(function(t){\n          return t.offsetParent!==null;\n        });\n\n        if(!tables.length){\n          alert('No tabular data is available on this page for Excel export.');\n          return;\n        }\n\n        var name=pageTitle();\n        var html='<html><head><meta charset="UTF-8"><style>table{border-collapse:collapse}td,th{border:1px solid #777;padding:5px}</style></head><body>';\n        html+='<h2>'+name.replace(/[<>&]/g,'')+'</h2>';\n        tables.forEach(function(t,i){\n          html+='<h3>Section '+(i+1)+'</h3>'+t.outerHTML;\n        });\n        html+='</body></html>';\n\n        var blob=new Blob(['\\ufeff',html],{type:'application/vnd.ms-excel'});\n        var a=document.createElement('a');\n        a.href=URL.createObjectURL(blob);\n        a.download=name.replace(/[^a-z0-9]+/gi,'_').replace(/^_|_$/g,'')+'_'+new Date().toISOString().slice(0,10)+'.xls';\n        document.body.appendChild(a);\n        a.click();\n        setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},700);\n      }catch(err){\n        alert('Excel export could not start. Please refresh once and try again.');\n      }\n    });\n  }\n})();\n</script>\n\n</body></html>\n""", 'my_leads.html': """{% extends "base.html" %}{% block content %}\n{% if u['role']!='MD' %}\n  {% if u['company_code']=='SCIC' %}<img class="page-brand-watermark" src="" alt="">\n  {% elif u['company_code']=='WWIC' %}<img class="page-brand-watermark" src="" alt="">{% endif %}\n{% endif %}\n<h1 class="title">My Leads</h1>\n<div class="card">\n  <h2>Lead Ranking & Filters</h2>\n  <div class="rank-legend">\n    <span class="legend-hot">🔴 HOT LEAD • 71–100%</span>\n    <span class="legend-warm">🟢 INTERESTED • 50–70%</span>\n    <span class="legend-cold">🔵 COLD / NOT CONTACTED • 0–49%</span>\n  </div>\n  <form method="get" class="filterbar">\n    <div><label>Search Client / Mobile / Lead ID</label><input name="q" value="{{filters.q}}" placeholder="Search..."></div>\n    <div><label>Ranking</label><select name="rank">\n      <option value="">All Rankings</option>\n      <option value="hot" {% if filters.rank=='hot' %}selected{% endif %}>Hot Lead 71–100%</option>\n      <option value="warm" {% if filters.rank=='warm' %}selected{% endif %}>Interested 50–70%</option>\n      <option value="cold" {% if filters.rank=='cold' %}selected{% endif %}>Cold / Not Contacted 0–49%</option>\n    </select></div>\n    <div><label>Status</label><select name="status">\n      <option value="">All Status</option>\n      {% for s in ['New','Called','Interested','Follow-up','Office Visit','Documents Pending','Enrolled','Not Interested','Closed'] %}\n      <option value="{{s}}" {% if filters.status==s %}selected{% endif %}>{{s}}</option>{% endfor %}\n    </select></div>\n    <div><label>Sort</label><select name="sort">\n      <option value="interest_desc" {% if filters.sort=='interest_desc' %}selected{% endif %}>Interest: High → Low</option>\n      <option value="interest_asc" {% if filters.sort=='interest_asc' %}selected{% endif %}>Interest: Low → High</option>\n      <option value="newest" {% if filters.sort=='newest' %}selected{% endif %}>Newest First</option>\n      <option value="followup" {% if filters.sort=='followup' %}selected{% endif %}>Follow-up Date</option>\n    </select></div>\n    <div class="filter-actions"><button class="btn">Apply Filter</button><a class="toolbtn" href="{{url_for('my_leads')}}">Clear</a></div>\n  </form>\n</div>\n\n<div class="card"><h2>Assigned Client Leads <span style="color:#e6b73f">({{rows|length}})</span></h2>\n<p>Color se interest level turant samajh aayega. Client profile ke andar counselor percentage update karega.</p>\n<div class="tablewrap"><table>\n<tr><th>Lead</th><th>Client</th><th>Mobile</th><th>Country/Visa</th><th>Source</th><th>Interest Ranking</th><th>Status</th><th>Action</th></tr>\n{% for r in rows %}\n{% set score = r['interest_score'] or 0 %}\n<tr class="{% if score >= 71 %}lead-hot{% elif score >= 50 %}lead-warm{% else %}lead-cold{% endif %}">\n<td>{{r['lead_id']}}</td>\n<td><a class="lead-link" href="{{url_for('lead_profile',lead_db_id=r['id'])}}">{{r['client_name'] or 'Open Client'}}</a><br>{{r['email']}}</td>\n<td>{{r['mobile']}}</td>\n<td>{{r['country']}}<br>{{r['visa_type']}}</td>\n<td>{{r['source']}}</td>\n<td>\n  {% if score >= 71 %}<span class="interest-badge interest-hot">🔥 HOT • <span class="interest-score">{{score}}%</span></span>\n  {% elif score >= 50 %}<span class="interest-badge interest-warm">✓ INTERESTED • <span class="interest-score">{{score}}%</span></span>\n  {% else %}<span class="interest-badge interest-cold">● COLD • <span class="interest-score">{{score}}%</span></span>{% endif %}\n</td>\n<td><span class="status-chip">{{r['status']}}</span>{% if r['followup_date'] %}<br>Follow-up: {{r['followup_date']}}{% endif %}</td>\n<td><a class="btn" style="min-width:145px;text-align:center" href="{{url_for('lead_profile',lead_db_id=r['id'])}}">OPEN CLIENT PROFILE</a></td>\n</tr>\n{% else %}<tr><td colspan="8">No leads match this filter.</td></tr>{% endfor %}\n</table></div></div>\n{% endblock %}""", 'lead_profile.html': """{% extends "base.html" %}{% block content %}\n{% if u['role']!='MD' %}\n  {% if u['company_code']=='SCIC' %}<img class="page-brand-watermark" src="" alt="">\n  {% elif u['company_code']=='WWIC' %}<img class="page-brand-watermark" src="" alt="">{% endif %}\n{% endif %}\n<div class="profile-head"><div><h1 class="title">Client Lead Profile</h1><div class="profile-meta"><span>{{lead['lead_id']}}</span><span>{{lead['company_code']}}</span><span>AM: {{lead['am_name'] or 'Unallocated'}}</span><span>Status: {{lead['status']}}</span>\n    {% set score=lead['interest_score'] or 0 %}\n    <span class="{% if score>=71 %}interest-hot{% elif score>=50 %}interest-warm{% else %}interest-cold{% endif %}">\n      Interest: {{score}}%\n    </span></div></div><div><a class="toolbtn" href="{{url_for('my_leads') if u['role']=='AM' else url_for('dashboard')}}">← Back</a></div></div>\n<form method="post" enctype="multipart/form-data">\n<div class="two-col"><div>\n<div class="card"><h2>Client Personal Information</h2><div class="grid3">\n<div><label>Full Name</label><input name="client_name" value="{{lead['client_name']}}"></div><div><label>Date of Birth</label><input type="date" name="date_of_birth" value="{{lead['date_of_birth']}}"></div><div><label>Education</label><input name="education" value="{{lead['education']}}" placeholder="10th / 12th / Graduate etc."></div>\n<div><label>Primary Mobile</label><input name="mobile" value="{{lead['mobile']}}"></div><div><label>Alternate Mobile</label><input name="alternate_mobile" value="{{lead['alternate_mobile']}}"></div><div><label>Email</label><input type="email" name="email" value="{{lead['email']}}"></div>\n<div><label>Occupation / Work Profile</label><input name="occupation" value="{{lead['occupation']}}"></div><div><label>Passport Number</label><input name="passport_no" value="{{lead['passport_no']}}"></div><div><label>Preferred Call Time</label><input name="preferred_call_time" value="{{lead['preferred_call_time']}}"></div>\n</div><label>Full Address</label><textarea name="address" rows="3">{{lead['address']}}</textarea><div class="grid3"><div><label>City</label><input name="city" value="{{lead['city']}}"></div><div><label>State</label><input name="state" value="{{lead['state']}}"></div><div><label>PIN Code</label><input name="pin_code" value="{{lead['pin_code']}}"></div></div></div>\n<div class="card"><h2>Immigration / Counseling Profile</h2><div class="grid3"><div><label>Interested Country</label><input name="country" value="{{lead['country']}}"></div><div><label>Visa Type</label><input name="visa_type" value="{{lead['visa_type']}}"></div><div><label>Approx. Budget</label><input name="budget" value="{{lead['budget']}}"></div></div><div class="grid"><div><label>Travel History</label><textarea name="travel_history" rows="3">{{lead['travel_history']}}</textarea></div><div><label>IELTS / PTE / Language Test</label><textarea name="english_test" rows="3">{{lead['english_test']}}</textarea></div></div></div>\n<div class="card"><h2>Counselor Work, Ranking & Follow-up</h2>\n<div class="rank-legend"><span class="legend-hot">🔴 Hot 71–100%</span><span class="legend-warm">🟢 Interested 50–70%</span><span class="legend-cold">🔵 Cold / Not Contacted 0–49%</span></div>\n<div class="grid3">\n<div><label>Lead Status</label><select name="status"><option>{{lead['status']}}</option><option>Interested</option><option>Not Interested</option><option>Call Back</option><option>Not Picked</option><option>No Plan</option><option>Budget Issue</option><option>Not Connected</option><option>Invalid No.</option><option>No WhatsApp</option><option>Enrolled</option><option>Discussion</option><option>Follow Up</option><option>Payment After Visa</option><option>Closed</option><option>Office Visit</option><option>Docs Received</option></select></div>\n<div><label>Interest Ranking (%)</label><input type="number" name="interest_score" min="0" max="100" step="1" value="{{lead['interest_score'] or 0}}" placeholder="0 to 100"></div>\n<div><label>Next Follow-up Date</label><input type="date" name="followup_date" value="{{lead['followup_date']}}"></div>\n</div>\n<div class="grid"><div><label>Quick Ranking Preset</label><select onchange="this.form.interest_score.value=this.value"><option value="">Choose preset</option><option value="100">🔥 100% - Super Hot</option><option value="85">🔥 85% - Hot Lead</option><option value="70">✓ 70% - Strong Interested</option><option value="60">✓ 60% - Interested</option><option value="50">✓ 50% - Medium Interest</option><option value="25">● 25% - Weak / Not Contacted</option><option value="0">● 0% - Not Interested</option></select></div><div><label>Source</label><input value="{{lead['source']}}" disabled></div></div>\n<label>Detailed Counselor Remarks / Client Discussion</label><textarea name="remarks" rows="7" placeholder="Mention complete discussion, requirements, objections, next action...">{{lead['remarks']}}</textarea></div>\n</div><div>\n<div class="card"><h2>Client Enquiry Form</h2><p>Office enquiry sheet / handwritten client form ka clear photo ya PDF yahan upload karein.</p><input type="file" name="client_form" accept=".jpg,.jpeg,.png,.webp,.pdf"><br><br><button class="btn" type="submit">Save Complete Client Profile</button></div>\n<div class="card"><h2>Uploaded Client Forms</h2>{% for d in docs %}<div class="doc-card"><div><b>{{d['document_type']}}</b><br>{{d['file_name']}}<br><small>{{d['uploaded_by']}} • {{d['uploaded_at']}}</small></div><a class="btn" target="_blank" href="{{url_for('lead_document',document_id=d['id'])}}">View</a></div>{% else %}<p>No client form uploaded yet.</p>{% endfor %}</div>\n<div class="card"><h2>Activity History</h2>{% for a in activity %}<div class="activity-item"><b>{{a['status'] or 'Updated'}}</b>{% if a['followup_date'] %} • Follow-up {{a['followup_date']}}{% endif %}<br>{{a['remarks']}}<br><small>{{a['updated_by']}} • {{a['updated_at']}}</small></div>{% else %}<p>No activity history yet.</p>{% endfor %}</div>\n</div></div></form>\n{% endblock %}""", 'dashboard.html': """{% extends "base.html" %}{% block content %}\n\n<div class="dashboard-heading-row">\n  <div>\n<div class="today-report-wrap luxury-today-report">\n  <div class="today-report-head">\n    <div><span class="today-live-dot"></span><b>TODAY'S REPORT</b><small>{{today_date}}</small></div>\n    <div class="today-report-note">Auto-calculated day-wise from Enrollment Date & Payment Received Date/Time</div>\n  </div>\n  {% if u.role=='MD' %}\n<style>\n/* v3.98 safe MD-only single-panel display.\n   No template cutting: simply suppress duplicate shells visually. */\n.competition-livebar.lux-score,\n.competition-livebar.lux-score.show{display:none!important}\n.today-report-wrap.luxury-today-report{\n  border:0!important;\n  background:transparent!important;\n  box-shadow:none!important;\n  padding:0!important;\n  margin:8px 0 18px!important;\n  overflow:visible!important\n}\n.today-report-wrap.luxury-today-report>.today-report-head{display:none!important}\n</style>\n<div id="compToggleBar" style="margin-bottom:10px;display:flex;align-items:center;justify-content:space-between;background:rgba(2,17,32,.9);border:1px solid rgba(230,183,63,.4);border-radius:12px;padding:10px 16px;flex-wrap:wrap;gap:8px">\n  <div style="display:flex;align-items:center;gap:10px">\n    <span style="color:#e6b73f;font-weight:bold;font-size:14px">&#127942; Competition</span>\n    <span id="compToggleBadge" style="font-size:12px;font-weight:bold;padding:3px 10px;border-radius:10px;background:rgba(230,183,63,.12);border:1px solid #e6b73f;color:#e6b73f">Loading...</span>\n  </div>\n  <form method="post" action="{{ url_for('competition_toggle') }}" style="margin:0">\n    <button type="submit" id="compToggleBtn" style="border:0;padding:8px 20px;border-radius:9px;font-size:14px;font-weight:bold;cursor:pointer;background:linear-gradient(135deg,#22c55e,#16a34a);color:#fff">Loading...</button>\n  </form>\n</div>\n<script>\n(function(){\n  function loadToggleStatus(){\n    fetch('/api/competition/live?after=0',{cache:'no-store'})\n    .then(function(r){return r.json();})\n    .then(function(d){\n      var badge=document.getElementById('compToggleBadge');\n      var btn=document.getElementById('compToggleBtn');\n      if(!badge||!btn)return;\n      if(d.enabled){\n        badge.textContent='ON - ACTIVE';\n        badge.style.background='rgba(34,197,94,.15)';\n        badge.style.borderColor='#22c55e';\n        badge.style.color='#22c55e';\n        btn.textContent='Turn OFF Competition';\n        btn.style.background='linear-gradient(135deg,#ef4444,#b91c1c)';\n      }else{\n        badge.textContent='OFF';\n        badge.style.background='rgba(107,114,128,.15)';\n        badge.style.borderColor='#6b7280';\n        badge.style.color='#9ca3af';\n        btn.textContent='Turn ON Competition';\n        btn.style.background='linear-gradient(135deg,#22c55e,#16a34a)';\n      }\n    }).catch(function(){});\n  }\n  loadToggleStatus();\n})();\n</script>\n<div id="mdMergedPerformance" class="md-merged-performance">\n  <div class="md-merge-head">\n    <div><span class="today-live-dot"></span><b>LIVE PERFORMANCE + TODAY'S REPORT</b><small>{{today_date}}</small></div>\n    <div class="md-merge-note">Overall live score + today's enrollments & received payments</div>\n  </div>\n  <div class="md-merge-grid">\n    <section class="md-team md-team-scic">\n      <div class="md-team-logo md-team-logo-scic"><img class="brand-logo sc-logo" src=""></div>\n      <div class="md-team-content">\n        <div class="md-crown">♛</div><h3>ABC</h3><div class="md-stars">★ ★ ★ ★ ★</div>\n        <div id="mdScicLiveScore" class="md-live-score">0 • ₹0</div>\n        <div class="md-live-label">OVERALL LIVE PERFORMANCE</div>\n        <div class="md-meter"><span id="mdScicMeter" style="width:0%"></span></div>\n        <div class="md-today-stats">\n          <div><span>Today's Enrollments</span><b>{{today_report['SCIC']['enrollments']}}</b></div>\n          <div><span>Today's Revenue</span><b>₹{{"{:,.0f}".format(today_report['SCIC']['revenue'] or 0)}}</b></div>\n        </div>\n      </div>\n    </section>\n    <div class="md-merge-vs">VS</div>\n    <section class="md-team md-team-wwic">\n      <div class="md-team-content">\n        <div class="md-crown">♛</div><h3>XYZ</h3><div class="md-stars">★ ★ ★ ★ ★</div>\n        <div id="mdWwicLiveScore" class="md-live-score">0 • ₹0</div>\n        <div class="md-live-label">OVERALL LIVE PERFORMANCE</div>\n        <div class="md-meter md-meter-blue"><span id="mdWwicMeter" style="width:0%"></span></div>\n        <div class="md-today-stats">\n          <div><span>Today's Enrollments</span><b>{{today_report['WWIC']['enrollments']}}</b></div>\n          <div><span>Today's Revenue</span><b>₹{{"{:,.0f}".format(today_report['WWIC']['revenue'] or 0)}}</b></div>\n        </div>\n      </div>\n      <div class="md-team-logo md-team-logo-wwic"><img class="brand-logo ww-logo" src=""></div>\n    </section>\n  </div>\n  <div class="md-merge-footer">\n    <span>● TODAY</span>\n    <b>{{today_report['SCIC']['enrollments']+today_report['WWIC']['enrollments']}} Enrollments</b>\n    <b>₹{{"{:,.0f}".format((today_report['SCIC']['revenue'] or 0)+(today_report['WWIC']['revenue'] or 0))}} Revenue</b>\n  </div>\n</div>\n<script>\n(function(){\n function sync(){\n  const pairs=[['scicLiveScore','mdScicLiveScore'],['wwicLiveScore','mdWwicLiveScore']];\n  pairs.forEach(p=>{const a=document.getElementById(p[0]),b=document.getElementById(p[1]);if(a&&b)b.textContent=a.textContent;});\n  [['scicMeter','mdScicMeter'],['wwicMeter','mdWwicMeter']].forEach(p=>{const a=document.getElementById(p[0]),b=document.getElementById(p[1]);if(a&&b)b.style.width=a.style.width||'0%';});\n }\n sync(); setInterval(sync,750);\n})();\n</script>\n{% elif u.role=='GM' %}\n  {% set c=u.company_code %}\n  <div class="today-company-grid single">\n    <div class="today-company-card {% if c=='SCIC' %}today-scic-card{% else %}today-wwic-card{% endif %}">\n      <div class="today-brand-row">\n        {% if c=='SCIC' %}<div class="today-logo today-logo-scic"><img src="" alt="ABC"></div>{% endif %}\n        <div class="today-brand-title"><i>♛</i><div class="today-company-name">{{'ABC' if c=='SCIC' else 'XYZ'}}</div><small>★ ★ ★ ★ ★</small></div>\n        {% if c!='SCIC' %}<div class="today-logo today-logo-wwic"><img src="" alt="XYZ"></div>{% endif %}\n      </div>\n      <div class="today-stat-grid">\n        <div><span>Today's Enrollments</span><b>{{today_report[c]['enrollments']}}</b></div>\n        <div><span>Today's Revenue</span><b>₹{{"{:,.0f}".format(today_report[c]['revenue'] or 0)}}</b></div>\n      </div>\n    </div>\n  </div>\n  {% endif %}\n</div>\n<h1 class="title">{% if u['role']=='MD' %}Executive Dashboard{% elif u['role']=='GM' %}General Manager Dashboard{% elif u['role']=='AM' %}Assistant Manager Dashboard{% else %}Reception Dashboard{% endif %}</h1></div>\n  <div class="dashboard-identity-card">\n    <div>\n      {% if u['photo_mime'] %}<img class="dashboard-identity-photo" src="{{url_for('user_photo',user_id=u['id'])}}" alt="{{u['full_name']}}">\n      {% else %}<div class="dashboard-identity-placeholder">{{(u['full_name'] or '?')[0]|upper}}</div>{% endif %}\n    </div>\n    <div>\n      <div class="dashboard-identity-name">{{u['full_name']}}</div>\n      <div class="dashboard-identity-designation">{{role_display(u['role'], u['designation'])}}</div>\n      <div class="dashboard-identity-code">Employee ID: <b>{{dashboard_employee['employee_code'] if dashboard_employee else (u['identity_code'] or '-')}}</b></div>\n      <div class="dashboard-identity-meta">\n        {% if u['company_code'] %}<span>{{u['company_code']}}</span>{% endif %}\n        {% if u['branch'] %}<span>Branch: {{u['branch']}}</span>{% endif %}\n        {% if u['official_mobile'] %}<span>Official: {{u['official_mobile']}}</span>{% endif %}\n      </div>\n      <a class="dashboard-profile-link" href="{{url_for('my_profile')}}">Open Professional Profile</a>\n    </div>\n    <div class="dashboard-rank-box">\n      {% if dashboard_rank %}\n        <div class="dashboard-rank-label">{{month_label}} Rank</div>\n        <div class="dashboard-rank-number">#{{dashboard_rank}}</div>\n      {% else %}\n        <div class="dashboard-rank-label">Employee Profile</div>\n        <div class="dashboard-rank-number">✓</div>\n      {% endif %}\n      {% if dashboard_score is not none %}\n        <div class="dashboard-score">{{dashboard_score}}%</div>\n        <div class="dashboard-rank-label">Enrollment Score</div>\n      {% else %}\n        <div class="dashboard-rank-label" style="margin-top:8px">Professional Identity</div>\n      {% endif %}\n    </div>\n  </div>\n</div>\n\n\n<div class="month-overview-top">\n  <div class="month-main-kpis">\n    <div class="month-kpi"><span>Total Leads of {{month_label}}</span><b>{{total_leads}}</b><small>Selected month</small></div>\n    <div class="month-kpi"><span>Allocated</span><b>{{allocated}}</b><small>Assigned to AMs</small></div>\n    <div class="month-kpi"><span>Unallocated</span><b>{{unallocated}}</b><small>Pending allocation</small></div>\n  </div>\n\n  {% if u['role'] in ['MD','GM'] %}\n  <div class="am-month-grid">\n    {% for a in allocation_summary %}\n    <div class="am-month-card">\n      <div class="am-card-head">\n        <div>\n          <div class="am-card-name">{{a['full_name']}}</div>\n          <small>{{a['company_code']}} • Assistant Manager</small>\n        </div>\n        <div class="am-conversion">{{a['positive_pct'] or 0}}%</div>\n      </div>\n      <div class="am-card-stats">\n        <div><span>Interested / Allocated</span><b>{{a['positive_count'] or 0}} / {{a['allocated_count'] or 0}}</b></div>\n        <div><span>Hot Leads</span><b>{{a['hot_count'] or 0}}</b></div>\n        <div><span>Enrolled</span><b>{{a['enrolled_count'] or 0}}</b></div>\n      </div>\n      <div class="am-progress"><i style="width:{{a['positive_pct'] or 0}}%"></i></div>\n      <small class="am-note">Conversion = Interested (50%+) ÷ Allocated</small>\n    </div>\n    {% else %}\n    <div class="card">No Assistant Managers created yet.</div>\n    {% endfor %}\n  </div>\n  {% endif %}\n</div>\n<form method="get" class="month-toolbar no-print">\n  <div><label>Dashboard Month</label><input type="month" name="month" value="{{selected_month}}"></div>\n  <button class="btn">View Month</button>\n  <a class="toolbtn" href="{{url_for('dashboard')}}">Current Month</a>\n</form>\n<div class="card" style="padding:11px 16px"><b style="color:#e6b73f">{{month_label}}</b> • Dashboard numbers are for the selected month.</div>\n\n{% if month_top_performers %}\n<div class="champion-banner">\n  <div class="champion-crown">👑</div>\n  <div>\n    <div style="letter-spacing:3px">AM ENROLLMENT PERFORMER OF THE MONTH • {{month_label}}</div>\n    <div class="champion-name">{{month_top_performers[0]['full_name']}}</div>\n    <div>{{month_top_performers[0]['company_code']}} • Assistant Manager</div>\n    <div>Enrollments: <b>{{month_top_performers[0]['enrollments']}}</b> • Revenue: <b>₹{{"{:,.0f}".format(month_top_performers[0]['revenue'])}}</b> • Conversion: <b>{{month_top_performers[0]['conversion']}}%</b></div>\n  </div>\n  <div class="champion-score">{{"%.0f"|format(month_top_performers[0]['score'])}}%<br><small>Enrollment Score</small></div>\n</div>\n\n<div class="champion-grid">\n{% for p in month_top_performers[:3] %}\n<div class="champion-card {% if loop.index==1 %}first{% elif loop.index==2 %}second{% else %}third{% endif %}">\n  <div>{% if loop.index==1 %}🥇 1st{% elif loop.index==2 %}🥈 2nd{% else %}🥉 3rd{% endif %}</div>\n  <div style="display:flex;align-items:center;gap:8px;margin-top:8px">\n    {% if p['photo_mime'] %}<img class="am-rank-photo" src="{{url_for('user_photo',user_id=p['id'])}}">{% else %}<span class="am-rank-avatar">{{(p['full_name'] or '?')[0]|upper}}</span>{% endif %}\n    <div><h3 style="margin:0">{{p['full_name']}}</h3><span class="role-chip">{{p['company_code']}}</span><span class="role-chip">AM</span></div>\n  </div>\n  <div class="month-score">{{"%.0f"|format(p['score'])}}%</div>\n  <small><b>{{p['enrollments']}}</b> Enrollments • <b>₹{{"{:,.0f}".format(p['revenue'])}}</b> Revenue • {{p['conversion']}}% Conversion</small>\n</div>\n{% endfor %}\n</div>\n{% endif %}\n\n{% if u['role']=='MD' %}\n<div class="company-champion-wrap">\n  <div class="company-champion-panel"><h2>ABC • {{month_label}} Top AMs</h2>\n  {% for p in scic_am_rankings[:2] %}<div class="mini-winner {% if loop.index==1 %}first{% else %}second{% endif %}">\n    <div><b>{% if loop.index==1 %}🥇{% else %}🥈{% endif %} {{p['full_name']}}</b><br><small>{{p['enrollments']}} Enrollments • ₹{{"{:,.0f}".format(p['revenue'])}}</small></div><strong>{{p['score']}}%</strong>\n  </div>{% else %}<small>No AM data for this month.</small>{% endfor %}</div>\n  <div class="company-champion-panel"><h2>XYZ • {{month_label}} Top AMs</h2>\n  {% for p in wwic_am_rankings[:2] %}<div class="mini-winner {% if loop.index==1 %}first{% else %}second{% endif %}">\n    <div><b>{% if loop.index==1 %}🥇{% else %}🥈{% endif %} {{p['full_name']}}</b><br><small>{{p['enrollments']}} Enrollments • ₹{{"{:,.0f}".format(p['revenue'])}}</small></div><strong>{{p['score']}}%</strong>\n  </div>{% else %}<small>No AM data for this month.</small>{% endfor %}</div>\n</div>\n{% endif %}\n\n<div class="card">\n<h2 style="color:#e6b73f">🏆 {{month_label}} Assistant Manager Leaderboard</h2>\n<p style="opacity:.72">{% if u['role']=='MD' %}All ABC + XYZ Assistant Managers.{% else %}Your company team ranking is visible to keep the whole team motivated.{% endif %}</p>\n<div class="tablewrap"><table class="am-leaderboard-table"><thead><tr><th>Rank</th><th>Assistant Manager</th><th>Company</th><th>Leads Allocated</th><th>Positive Leads</th><th>Enrollments</th><th>Revenue Generated</th><th>Conversion</th><th>Enrollment Score</th></tr></thead><tbody>\n{% for p in company_am_rankings %}\n<tr class="{% if p['rank']==1 %}rank-one{% elif p['rank']==2 %}rank-two{% elif p['rank']==3 %}rank-three{% endif %}">\n<td><b>{% if p['rank']==1 %}🥇{% elif p['rank']==2 %}🥈{% elif p['rank']==3 %}🥉{% endif %} #{{p['rank']}}</b></td>\n<td>{% if p['photo_mime'] %}<img class="am-rank-photo" src="{{url_for('user_photo',user_id=p['id'])}}">{% else %}<span class="am-rank-avatar">{{(p['full_name'] or '?')[0]|upper}}</span>{% endif %}<b>{{p['full_name']}}</b>{% if not p['active'] %}<br><small>Inactive / Former</small>{% endif %}</td>\n<td>{{p['company_code']}}</td><td>{{p['allocated']}}</td><td>{{p['positive']}}</td><td><b>{{p['enrollments']}}</b></td><td><b>₹{{"{:,.0f}".format(p['revenue'])}}</b></td><td>{{p['conversion']}}%</td><td><b>{{p['score']}}%</b></td>\n</tr>\n{% else %}<tr><td colspan="9">No Assistant Manager records available for this month.</td></tr>{% endfor %}\n</tbody></table></div>\n</div>\n\n{% if u['role']=='AM' and my_monthly_performance %}\n<div class="card self-performance">\n<h2>My {{month_label}} Performance</h2>\n<div class="metrics">\n  <div class="metric">My Rank<b>#{{my_monthly_performance['rank']}}</b></div>\n  <div class="metric">Enrollment Score<b>{{my_monthly_performance['score']}}%</b></div>\n  <div class="metric">Leads Allocated<b>{{my_monthly_performance['allocated']}}</b></div>\n  <div class="metric">Positive Leads<b>{{my_monthly_performance['positive']}}</b></div>\n</div>\n<p>Your Level:\n{% if my_monthly_performance['score']>=85 %}<span class="level-badge level-star">⭐ STAR PERFORMER</span>\n{% elif my_monthly_performance['score']>=70 %}<span class="level-badge level-strong">STRONG PERFORMER</span>\n{% elif my_monthly_performance['score']>=50 %}<span class="level-badge level-good">GOOD / IMPROVING</span>\n{% else %}<span class="level-badge level-build">BUILDING PERFORMANCE</span>{% endif %}\n{% if best_am %} • Best AM this month: <b style="color:#e6b73f">{{best_am['full_name']}} ({{best_am['score']}}%)</b>{% endif %}\n</p>\n</div>\n{% endif %}\n\n{% if u['role']=='MD' %}\n<div class="company-compare">\n  {% for c in company_stats %}\n  <div class="company-panel">\n    <div class="company-logo-title">\n      {% if c.code=='SCIC' %}\n        <img src="" alt="ABC Logo">\n      {% else %}\n        <img class="wwic-logo" src="" alt="XYZ Logo">\n      {% endif %}\n      <h2 class="{% if c.code=='SCIC' %}company-title-scic{% else %}company-title-wwic{% endif %}">\n        {% if c.code=='SCIC' %}ABC{% else %}XYZ{% endif %}\n      </h2>\n    </div>\n    <div class="subtle">{{'ABC' if c.code=='SCIC' else 'XYZ'}} • Live Company Performance</div>\n    <div class="company-metrics">\n      <div class="company-metric"><span>Total Leads</span><b>{{c.total}}</b></div>\n      <div class="company-metric"><span>Allocated</span><b>{{c.allocated}}</b></div>\n      <div class="company-metric"><span>Unallocated</span><b>{{c.unallocated}}</b></div>\n      <div class="company-metric"><span>Positive Data</span><b>{{c.positive}}</b></div>\n      <div class="company-metric"><span>Hot Leads</span><b>{{c.hot}}</b></div>\n      <div class="company-metric"><span>Interested</span><b>{{c.warm}}</b></div>\n      <div class="company-metric"><span>Enrolled</span><b>{{c.enrolled}}</b></div>\n      <div class="company-metric"><span>Revenue</span><b>₹{{"%.0f"|format(c.revenue or 0)}}</b></div>\n    </div>\n  </div>\n  {% endfor %}\n</div>\n\n<div class="analytics-grid">\n  {% for c in company_stats %}\n  <div class="card">\n    <div class="logo-strip">\n      {% if c.code=='SCIC' %}<img src="" alt="ABC"><strong>ABC</strong>\n      {% else %}<img src="" alt="XYZ"><strong>XYZ</strong>{% endif %}\n    </div>\n    <h2>{% if c.code=='SCIC' %}ABC{% else %}XYZ{% endif %} • Positive Data Pie</h2>\n    <div class="pie-wrap">\n      <div class="pie-chart" style="background:conic-gradient(#ff4d61 0 {{c.hot_pct}}%, #48d58b {{c.hot_pct}}% {{c.positive_pct}}%, #4aa8ff {{c.positive_pct}}% 100%);">\n        <div class="pie-center"><b>{{c.positive}}</b><span>Positive Leads</span><small>{{c.positive_rate}}%</small></div>\n      </div>\n      <div class="pie-legend">\n        <div><span><i class="dot dot-hot"></i>Hot 71–100%</span><b>{{c.hot}}</b></div>\n        <div><span><i class="dot dot-warm"></i>Interested 50–70%</span><b>{{c.warm}}</b></div>\n        <div><span><i class="dot dot-cold"></i>Cold / Not Contacted</span><b>{{c.cold}}</b></div>\n        <div><span>Average Interest</span><b>{{c.avg_interest}}%</b></div>\n      </div>\n    </div>\n  </div>\n  {% endfor %}\n</div>\n{% endif %}\n\n<div class="card"><h2>Lead Interest Snapshot</h2><div class="metrics">\n<div class="metric" style="border-left:5px solid #ff4d61">Hot Leads<b>{{hot_count}}</b><small>71–100%</small></div>\n<div class="metric" style="border-left:5px solid #48d58b">Interested<b>{{warm_count}}</b><small>50–70%</small></div>\n<div class="metric" style="border-left:5px solid #4aa8ff">Cold / Not Contacted<b>{{cold_count}}</b><small>0–49%</small></div>\n<div class="metric">Average Interest<b>{{avg_interest}}%</b></div></div></div>\n\n<div class="card"><h2>Recent Leads</h2><div class="tablewrap"><table><tr><th>Lead ID</th><th>Client</th><th>Mobile</th><th>Country</th><th>Interest</th><th>Status</th><th>Profile</th></tr>\n{% for r in recent %}{% set score=r['interest_score'] or 0 %}<tr class="{% if score>=71 %}lead-hot{% elif score>=50 %}lead-warm{% else %}lead-cold{% endif %}"><td>{{r['lead_id']}}</td><td>{{r['client_name']}}</td><td>{{r['mobile']}}</td><td>{{r['country']}}</td><td>{% if score>=71 %}<span class="interest-badge interest-hot">🔥 {{score}}%</span>{% elif score>=50 %}<span class="interest-badge interest-warm">✓ {{score}}%</span>{% else %}<span class="interest-badge interest-cold">● {{score}}%</span>{% endif %}</td><td>{{r['status']}}</td><td><a class="btn" href="{{url_for('lead_profile',lead_db_id=r['id'])}}">Open Client</a></td></tr>{% else %}<tr><td colspan="7">No leads yet.</td></tr>{% endfor %}</table></div></div>\n\n<div id="desktopReminderStack" class="desktop-reminder-stack"></div>\n<script>\n(function(){\n  const appearance="{{u['appearance_mode'] or 'dark'}}";\n  if(appearance==='light') document.body.classList.add('user-light');\n  if(appearance==='system' && window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) document.body.classList.add('user-light');\n\n  function escapeHtml(s){return String(s||'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}\n  function showReminder(item,index){\n    const stack=document.getElementById('desktopReminderStack'); if(!stack)return;\n    const box=document.createElement('div');\n    box.className='desktop-reminder '+(item.overdue?'overdue':'');\n    box.style.animationDelay=(index*120)+'ms';\n    box.innerHTML='<button class="desktop-reminder-close">×</button>'+\n      '<div class="desktop-reminder-app">🔔 IMS • Follow-up Reminder</div>'+\n      '<div class="desktop-reminder-person">'+escapeHtml(item.client_name)+'</div>'+\n      '<div class="desktop-reminder-text">'+(item.overdue?'OVERDUE • ':'Due Today • ')+escapeHtml(item.followup_date)+'<br>'+\n      'Interest '+escapeHtml(item.interest)+'%'+(item.am_name?' • '+escapeHtml(item.am_name):'')+'</div>'+\n      '<div class="desktop-reminder-actions"><a href="'+item.url+'">Open Client</a><button type="button">Dismiss</button></div>';\n    box.querySelector('.desktop-reminder-close').onclick=()=>box.remove();\n    box.querySelector('.desktop-reminder-actions button').onclick=()=>box.remove();\n    stack.appendChild(box);\n    if(index===0) setTimeout(()=>{if(box.isConnected)box.classList.add('visible')},50);\n    else setTimeout(()=>{if(box.isConnected)box.classList.add('visible')},index*250+50);\n  }\n\n  fetch('{{url_for("dashboard_reminders_api")}}',{credentials:'same-origin'})\n    .then(r=>r.ok?r.json():null).then(data=>{\n      if(!data)return;\n      const chip=document.getElementById('dueFollowupChip'), count=document.getElementById('dueFollowupCount');\n      if(data.count>0){count.textContent=data.count;chip.style.display='inline-flex'}\n      const key='gaur-reminder-popup-'+new Date().toISOString().slice(0,10)+'-{{u["id"]}}';\n      if(data.items && data.items.length && sessionStorage.getItem(key)!=='1'){\n        data.items.slice(0,3).forEach((x,i)=>showReminder(x,i));\n        sessionStorage.setItem(key,'1');\n      }\n    }).catch(()=>{});\n})();\n</script>\n{% endblock %}""", 'leads.html': """{% extends "base.html" %}{% block content %}\n{% if u.role in ['MD','GM'] %}<div class="card" style="border:2px solid #e6b73f"><h2 style="color:#e6b73f">Management Lead Control</h2><p>Select exact leads, allocate/re-allocate to any authorized Assistant Manager, monitor work, or archive wrong leads.</p><a class="btn" href="{{url_for('lead_control_center')}}">🎯 Open Lead Allocation Control Center</a></div>{% endif %}\n\n{% if u['role']!='MD' %}\n  {% if u['company_code']=='SCIC' %}<img class="page-brand-watermark" src="" alt="">\n  {% elif u['company_code']=='WWIC' %}<img class="page-brand-watermark" src="" alt="">{% endif %}\n{% endif %}\n<h1 class="title">Lead Upload & AM Allocation</h1>\n<div class="card"><h2>Upload Excel Leads</h2>\n<div style="margin-bottom:14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap">\n  <a href="{{ url_for('leads_sample_excel') }}" class="btn" style="background:linear-gradient(135deg,#e6b73f,#b8860b);color:#091423;text-decoration:none;padding:9px 18px;border-radius:9px;font-weight:bold;font-size:14px;display:inline-flex;align-items:center;gap:6px">&#11123; Download Sample Excel</a>\n  <span style="color:#7f9caf;font-size:13px">Download → Fill client data → Upload below</span>\n</div>\n<form method="post" enctype="multipart/form-data"><div class="grid">\n<div><label>Excel File (.xls / .xlsx)</label><input type="file" name="excel_file" accept=".xls,.xlsx" required></div>\n{% if u['role']=='MD' %}<div><label>Company</label><select name="company_code"><option>ABC</option><option>XYZ</option></select></div>{% endif %}\n</div><br><button class="btn">Import Excel</button></form><p>Duplicates are never deleted: same-company = yellow, cross-company = red.</p></div>\n\n<div class="card"><h2>Allocate Unassigned Leads</h2><form method="post" action="{{url_for('allocate')}}"><div class="grid">\n<div><label>Assistant Manager</label><select name="am_id" required><option value="">Select AM</option>\n{% for am in am_rows %}<option value="{{am['id']}}">{{am['company_code']}} - {{am['full_name']}} ({{am['login_id']}})</option>{% endfor %}\n</select></div><div><label>Quantity</label><input type="number" min="1" name="qty" value="20"></div></div><br><button class="btn">Allocate Leads</button></form></div>\n\n<div class="card"><h2>Lead Ranking & Data Filters</h2>\n<div class="rank-legend"><span class="legend-hot">🔴 HOT • 71–100%</span><span class="legend-warm">🟢 INTERESTED • 50–70%</span><span class="legend-cold">🔵 COLD / NOT CONTACTED • 0–49%</span></div>\n<form method="get" class="filterbar">\n<div><label>Search Client / Mobile / Lead ID</label><input name="q" value="{{filters.q}}" placeholder="Search..."></div>\n{% if u['role']=='MD' %}<div><label>Company</label><select name="company"><option value="">Both Companies</option><option value="SCIC" {% if filters.company=='SCIC' %}selected{% endif %}>ABC</option><option value="WWIC" {% if filters.company=='WWIC' %}selected{% endif %}>XYZ</option></select></div>{% endif %}\n<div><label>Assistant Manager</label><select name="am"><option value="">All AMs</option><option value="unallocated" {% if filters.am=='unallocated' %}selected{% endif %}>Unallocated</option>{% for am in am_rows %}<option value="{{am['id']}}" {% if filters.am==am['id']|string %}selected{% endif %}>{{am['full_name']}}</option>{% endfor %}</select></div>\n<div><label>Ranking</label><select name="rank"><option value="">All Rankings</option><option value="hot" {% if filters.rank=='hot' %}selected{% endif %}>Hot 71–100%</option><option value="warm" {% if filters.rank=='warm' %}selected{% endif %}>Interested 50–70%</option><option value="cold" {% if filters.rank=='cold' %}selected{% endif %}>Cold 0–49%</option></select></div>\n<div><label>Status</label><select name="status"><option value="">All Status</option>{% for s in ['New','Called','Interested','Follow-up','Office Visit','Documents Pending','Enrolled','Not Interested','Closed'] %}<option value="{{s}}" {% if filters.status==s %}selected{% endif %}>{{s}}</option>{% endfor %}</select></div>\n<div><label>Sort</label><select name="sort"><option value="interest_desc" {% if filters.sort=='interest_desc' %}selected{% endif %}>Interest High → Low</option><option value="interest_asc" {% if filters.sort=='interest_asc' %}selected{% endif %}>Interest Low → High</option><option value="newest" {% if filters.sort=='newest' %}selected{% endif %}>Newest First</option><option value="followup" {% if filters.sort=='followup' %}selected{% endif %}>Follow-up Date</option></select></div>\n<div class="filter-actions"><button class="btn">Apply Filter</button><a class="toolbtn" href="{{url_for('leads')}}">Clear</a></div>\n</form></div>\n\n<div class="card"><h2>Imported / Working Leads <span style="color:#e6b73f">({{rows|length}})</span></h2><div style="font-size:12px;opacity:.72;margin:6px 0 10px">Upload Date & Time = actual portal import timestamp. Sr No. = current filtered list serial.</div><div class="tablewrap"><table>\n<tr><th>Sr No.</th><th>Upload Date & Time</th><th>Lead ID</th><th>Client</th><th>Mobile</th><th>City</th><th>Country/Visa</th><th>AM</th><th>Interest Ranking</th><th>Status</th><th>Client Profile</th></tr>\n{% for r in rows %}{% set score=r['interest_score'] or 0 %}\n<tr class="{% if score>=71 %}lead-hot{% elif score>=50 %}lead-warm{% else %}lead-cold{% endif %}">\n<td><b>{{loop.index}}</b></td>\n<td><b>{{r['imported_at'][:10] if r['imported_at'] else '-'}}</b><br><small>{{r['imported_at'][11:19] if r['imported_at'] and r['imported_at']|length>=19 else ''}}</small></td>\n<td>{{r['lead_id']}}</td><td>{{r['client_name']}}<br><small>{{r['email']}}</small></td><td>{{r['mobile']}}</td><td>{{r['city']}}</td><td>{{r['country']}}<br>{{r['visa_type']}}</td><td>{{r['am_name'] or 'Unallocated'}}</td>\n<td>{% if score>=71 %}<span class="interest-badge interest-hot">🔥 {{score}}%</span>{% elif score>=50 %}<span class="interest-badge interest-warm">✓ {{score}}%</span>{% else %}<span class="interest-badge interest-cold">● {{score}}%</span>{% endif %}</td>\n<td>{{r['status']}}{% if r['followup_date'] %}<br><small>{{r['followup_date']}}</small>{% endif %}</td>\n<td><a class="btn" href="{{url_for('lead_profile',lead_db_id=r['id'])}}">Open Client</a></td></tr>\n{% else %}<tr><td colspan="11">No leads match this filter.</td></tr>{% endfor %}\n</table></div></div>\n\n<div class="card"><h2>Import History</h2><div class="tablewrap"><table><tr><th>Company</th><th>File</th><th>Total</th><th>New</th><th>Same Dup</th><th>Cross Dup</th><th>Invalid</th><th>By</th><th>At</th></tr>\n{% for r in imports %}<tr><td>{{r['company_code']}}</td><td>{{r['filename']}}</td><td>{{r['total_rows']}}</td><td>{{r['new_rows']}}</td><td>{{r['same_company_duplicates']}}</td><td>{{r['cross_company_duplicates']}}</td><td>{{r['invalid_mobiles']}}</td><td>{{r['imported_by']}}</td><td>{{r['imported_at']}}</td></tr>{% endfor %}</table></div></div>\n{% endblock %}""", 'performance_directory.html': """{% extends "base.html" %}{% block content %}\n<div class="performance-shell">\n  <div style="display:flex;justify-content:space-between;gap:14px;align-items:flex-start;flex-wrap:wrap">\n    <div>\n      <h1 class="title" style="margin-bottom:4px">Employee Performance Dashboard</h1>\n      <div style="color:#e6b73f">Enrollment Based & Revenue Based Performance Analysis</div>\n    </div>\n    <div class="no-print">\n      <form method="get" class="month-toolbar">\n        <div><label>Performance Month</label><input type="month" name="month" value="{{selected_month}}"></div>\n        <div><label>Designation / Role</label><select name="designation">\n          <option value="">All Employees</option>\n          {% for d in designations %}<option value="{{d}}" {% if selected_designation==d %}selected{% endif %}>{{d}}</option>{% endfor %}\n        </select></div>\n        {% if u['role']=='MD' %}<div><label>Company</label><select name="company">\n          <option value="">Both Companies</option><option value="SCIC" {% if selected_company=='SCIC' %}selected{% endif %}>ABC</option>\n          <option value="WWIC" {% if selected_company=='WWIC' %}selected{% endif %}>XYZ</option>\n        </select></div>{% endif %}\n        <button class="btn">Show Month</button>\n      </form>\n    </div>\n  </div>\n\n  <div class="kpi-grid">\n    <div class="kpi-card"><div class="kpi-label">Total Enrollments</div><div class="kpi-value">{{total_enrollments}}</div><div class="kpi-sub">{{label}}</div></div>\n    <div class="kpi-card"><div class="kpi-label">Total Revenue</div><div class="kpi-value">₹{{"{:,.0f}".format(total_revenue)}}</div><div class="kpi-sub">Selected period</div></div>\n    <div class="kpi-card"><div class="kpi-label">Avg. Enrollments / Emp.</div><div class="kpi-value">{{"%.1f"|format(avg_enrollments)}}</div><div class="kpi-sub">Active team average</div></div>\n    <div class="kpi-card"><div class="kpi-label">Avg. Revenue / Emp.</div><div class="kpi-value">₹{{"{:,.0f}".format(avg_revenue)}}</div><div class="kpi-sub">Active team average</div></div>\n    <div class="kpi-card"><div class="kpi-label">Active Employees</div><div class="kpi-value">{{active_employees}}</div><div class="kpi-sub">Included in ranking</div></div>\n    <div class="kpi-card"><div class="kpi-label">Top Performer</div><div class="kpi-value" style="font-size:22px">{{top_performer['full_name'] if top_performer else '-'}}</div><div class="kpi-sub">{% if top_performer %}{{"%.1f"|format(top_performer['score'])}}% Score{% endif %}</div></div>\n  </div>\n\n  <div class="performance-grid">\n    <div class="chart-panel">\n      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap">\n        <h2 style="color:#e6b73f;margin:0">Month-wise Performance Trend (Top 3 Employees)</h2>\n        <span>{{label}}</span>\n      </div>\n      {% if top3 %}\n      <div style="margin:8px 0 14px">\n        {% for r in top3 %}\n          <span style="margin-right:18px">{% if r['rank']==1 %}🥇{% elif r['rank']==2 %}🥈{% else %}🥉{% endif %} {{r['full_name']}}</span>\n        {% endfor %}\n      </div>\n      <svg class="chart-svg" viewBox="0 0 900 420" role="img" aria-label="Monthly top three employee performance trend">\n        <line x1="70" y1="340" x2="860" y2="340" stroke="#315a7b" stroke-width="2"/>\n        <line x1="70" y1="60" x2="70" y2="340" stroke="#315a7b" stroke-width="2"/>\n        {% for y in [0,20,40,60,80,100] %}\n          {% set yy=340-(y*2.8) %}\n          <line x1="70" y1="{{yy}}" x2="860" y2="{{yy}}" stroke="#243f59" stroke-width="1" stroke-dasharray="4 5"/>\n          <text x="35" y="{{yy+5}}" fill="#dce8f2" font-size="15">{{y}}%</text>\n        {% endfor %}\n        {% set x0=95 %}{% set step=63 %}\n        {% for item in trend %}\n          {% set x=x0+(loop.index0*step) %}\n          <text x="{{x}}" y="372" text-anchor="middle" fill="#dce8f2" font-size="15">{{item['month']}}</text>\n          {% if top3|length>0 %}\n          <rect x="{{x-14}}" y="{{340-(item.get('score1',0)*2.8)}}" width="12" height="{{item.get('score1',0)*2.8}}" rx="6" fill="#e6b73f" opacity=".42"/>\n          {% endif %}\n          {% if top3|length>1 %}\n          <rect x="{{x}}" y="{{340-(item.get('score2',0)*2.8)}}" width="12" height="{{item.get('score2',0)*2.8}}" rx="6" fill="#4e8fd4" opacity=".55"/>\n          {% endif %}\n          {% if top3|length>2 %}\n          <rect x="{{x+14}}" y="{{340-(item.get('score3',0)*2.8)}}" width="12" height="{{item.get('score3',0)*2.8}}" rx="6" fill="#d051ac" opacity=".55"/>\n          {% endif %}\n        {% endfor %}\n        {% if top3|length>0 %}\n        <polyline points="{% for item in trend %}{{x0+(loop.index0*step)}},{{340-(item.get('score1',0)*2.8)}} {% endfor %}" fill="none" stroke="#e6b73f" stroke-width="4"/>\n        {% endif %}\n        {% if top3|length>1 %}\n        <polyline points="{% for item in trend %}{{x0+(loop.index0*step)}},{{340-(item.get('score2',0)*2.8)}} {% endfor %}" fill="none" stroke="#4e8fd4" stroke-width="3"/>\n        {% endif %}\n        {% if top3|length>2 %}\n        <polyline points="{% for item in trend %}{{x0+(loop.index0*step)}},{{340-(item.get('score3',0)*2.8)}} {% endfor %}" fill="none" stroke="#d051ac" stroke-width="3"/>\n        {% endif %}\n      </svg>\n      {% else %}\n      <p>No performance data available.</p>\n      {% endif %}\n    </div>\n\n    <div class="chart-panel">\n      <h2 style="color:#e6b73f;margin-top:0">Top Performer Ranking</h2>\n      <div class="rank-grid">\n        {% for r in top3 %}\n        <div class="rank-card {% if r['rank']==1 %}first{% elif r['rank']==2 %}second{% else %}third{% endif %}">\n          <div class="rank-badge">{% if r['rank']==1 %}👑 1st{% elif r['rank']==2 %}🥈 2nd{% else %}🥉 3rd{% endif %}</div>\n          {% if r['photo_path'] %}<img src="{{url_for('employee_photo',employee_id=r['id'])}}" alt="{{r['full_name']}}">{% elif r.get('portal_user_id') %}<img src="{{url_for('user_photo',user_id=r['portal_user_id'])}}" alt="{{r['full_name']}}">{% endif %}\n          <h3 style="margin:8px 0 3px">{{r['full_name']}}</h3>\n          <div>{{r['company_code']}} • {{r['designation']}}</div>\n          <div class="rank-score">{{"%.0f"|format(r['score'])}}%</div>\n          <div class="grid" style="grid-template-columns:1fr 1fr;gap:6px">\n            <div><small>Enrollments</small><br><b>{{r['enrollments']}}</b></div>\n            <div><small>Revenue</small><br><b>₹{{"{:,.0f}".format(r['revenue'])}}</b></div>\n          </div>\n        </div>\n        {% endfor %}\n      </div>\n    </div>\n  </div>\n\n  <div class="performance-grid" style="grid-template-columns:.8fr 1.8fr">\n    <div class="donut-grid">\n      <div class="donut-card">\n        <h3 style="color:#e6b73f;margin-top:0">Enrollment Based Performance</h3>\n        <div class="donut"><div class="donut-center">{{total_enrollments}}<br><small>Total</small></div></div>\n        {% for r in top3 %}\n          <div class="legend-row"><span>{{r['full_name']}}</span><b>{{r['enrollments']}}</b></div>\n        {% endfor %}\n      </div>\n      <div class="donut-card">\n        <h3 style="color:#e6b73f;margin-top:0">Revenue Based Performance</h3>\n        <div class="donut"><div class="donut-center">₹{{"{:,.0f}".format(total_revenue/100000)}}L<br><small>Total</small></div></div>\n        {% for r in top3 %}\n          <div class="legend-row"><span>{{r['full_name']}}</span><b>₹{{"{:,.0f}".format(r['revenue']/100000)}}L</b></div>\n        {% endfor %}\n      </div>\n    </div>\n\n    <div class="chart-panel">\n      <h2 style="color:#e6b73f;margin-top:0">Employee Performance Table</h2>\n      <div class="tablewrap">\n      <table>\n        <tr><th>#</th><th>Employee Name</th><th>Enrollments</th><th>Revenue (₹)</th><th>Enrollment %</th><th>Revenue %</th><th>Performance Score</th><th>Rank</th><th>Report</th></tr>\n        {% set top_enroll=(ranked[0]['enrollments'] if ranked else 0) %}\n        {% set top_revenue=(ranked[0]['revenue'] if ranked else 0) %}\n        {% for r in ranked %}\n        {% set s=summaries.get(r['id'],{}) %}\n        <tr>\n          <td>{{r['rank']}}</td>\n          <td><b>{{r['full_name']}}</b><br><small>{{r['company_code']}} • {{r['designation']}}</small></td>\n          <td>{{r['enrollments']}}</td>\n          <td>₹{{"{:,.0f}".format(r['revenue'])}}</td>\n          <td>{{"%.0f"|format(s.get('enrollment_index',0))}}%</td>\n          <td>{{"%.0f"|format(s.get('revenue_index',0))}}%</td>\n          <td>\n            <b>{{"%.0f"|format(r['score'])}}%</b>\n            <div class="scorebar"><span style="width:{{[r['score'],100]|min}}%"></span></div>\n          </td>\n          <td>{% if r['rank']==1 %}🥇{% elif r['rank']==2 %}🥈{% elif r['rank']==3 %}🥉{% else %}{{r['rank']}}{% endif %}</td>\n          <td><a class="btn" href="{{url_for('employee_performance',employee_id=r['id'],period='month',month=selected_month)}}">Open</a></td>\n        </tr>\n        {% endfor %}\n      </table>\n      </div>\n    </div>\n  </div>\n\n  <div class="card">\n    <h2 style="color:#e6b73f;margin-top:0">12-Month Performance Snapshot</h2>\n    <div class="month-strip">\n      {% for m in year_months %}\n      <div class="month-box {% if m['key']==selected_month %}current{% endif %}">\n        <span>{{m['label']}}</span><b>{{m['top_score']}}%</b><small>{{m['top_name'] or '-'}}</small>\n      </div>\n      {% endfor %}\n    </div>\n  </div>\n\n  <div class="card">\n    <h2 style="color:#e6b73f;margin-top:0">Performance Score Formula</h2>\n    <div>(40% Enrollment Weightage) + (40% Revenue Weightage) + (10% Target Achievement) + (10% Manager Rating)</div>\n  </div>\n</div>\n{% endblock %}""", 'my_profile.html': """{% extends "base.html" %}{% block content %}\n<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap">\n  <h1 class="title">My Profile</h1>\n  <div class="identity-actions">\n    <a class="toolbtn" href="{{url_for('digital_id_card')}}">My Digital ID Card</a>\n  </div>\n</div>\n\n<div class="identity-hero">\n  <div class="identity-profile">\n    <div style="text-align:center">\n      {% if u['photo_mime'] %}\n        <img class="identity-avatar big" src="{{url_for('user_photo',user_id=u['id'])}}" alt="{{u['full_name']}}">\n      {% else %}\n        <div class="identity-placeholder big" style="margin:auto">{{(u['full_name'] or '?')[0]|upper}}</div>\n      {% endif %}\n      <div style="margin-top:12px">\n        <span class="identity-badge gold">{{u['role']}}</span>\n        {% if u['company_code'] %}<span class="identity-badge">{{u['company_code']}}</span>{% endif %}\n        {% if u['mobile_verified'] %}<span class="identity-badge green">✓ Mobile Verified</span>{% else %}<span class="identity-badge">OTP Ready • Verification Pending</span>{% endif %}\n      </div>\n    </div>\n\n    <div>\n      <div class="identity-name">{{u['full_name']}}</div>\n      <div style="font-size:18px;margin:5px 0 12px">{{u['designation'] or u['role']}}{% if u['department'] %} • {{u['department']}}{% endif %}</div>\n      {% if u['profile_headline'] %}<p style="font-size:18px">{{u['profile_headline']}}</p>{% endif %}\n      <div class="identity-stat-grid">\n        <div class="identity-stat">Employee ID<b>{{employee['employee_code'] if employee else (u['identity_code'] or '-')}}</b></div>\n        <div class="identity-stat">Official Mobile<b style="font-size:18px">{{u['official_mobile'] or '-'}}</b></div>\n        <div class="identity-stat">Branch<b style="font-size:18px">{{u['branch'] or '-'}}</b></div>\n        <div class="identity-stat">Joining Date<b style="font-size:18px">{{u['joining_date'] or '-'}}</b></div>\n      </div>\n    </div>\n  </div>\n</div>\n\n<div class="grid" style="grid-template-columns:1fr 1fr;margin-top:16px">\n  <div class="card">\n    <h2 style="color:#e6b73f;margin-top:0">Official Identity</h2>\n    <label>Login ID</label><div class="readonly-value">{{u['login_id']}}</div>\n    <label>Company</label><div class="readonly-value">{{u['company_code'] or 'THE IMMIGRATION MANAGEMENT SYSTEM'}}</div>\n    <label>Designation</label><div class="readonly-value">{{u['designation'] or u['role']}}</div>\n    <label>Department</label><div class="readonly-value">{{u['department'] or '-'}}</div>\n    <label>Official Email</label><div class="readonly-value">{{u['official_email'] or '-'}}</div>\n    <label>Reporting Manager</label><div class="readonly-value">{{reporting_manager['full_name'] if reporting_manager else '-'}}</div>\n    <p style="opacity:.72">Official identity fields are controlled by Management to protect company records.</p>\n  </div>\n\n  <div class="card">\n    <h2 style="color:#e6b73f;margin-top:0">Personalize My Profile</h2>\n    <form method="post" enctype="multipart/form-data">\n      <label>Profile Photo</label>\n      <input type="file" name="profile_photo" accept=".jpg,.jpeg,.png,.webp,image/*">\n      <label>Professional Headline</label>\n      <input name="profile_headline" maxlength="120" value="{{u['profile_headline'] or ''}}" placeholder="e.g. Immigration Counselor • Client Success">\n      <label>About Me / Professional Bio</label>\n      <textarea name="profile_bio" rows="5" maxlength="600" placeholder="Short professional introduction...">{{u['profile_bio'] or ''}}</textarea>\n      <button class="btn">Save My Profile</button>\n    </form>\n    {% if u['photo_mime'] %}\n    <form method="post" action="{{url_for('remove_my_profile_photo')}}" style="margin-top:10px">\n      <button class="toolbtn" type="submit">Remove Current Photo</button>\n    </form>\n    {% endif %}\n  </div>\n</div>\n\n{% if performance %}\n<div class="card" style="margin-top:16px">\n  <h2 style="color:#e6b73f;margin-top:0">My Current Performance Identity</h2>\n  <div class="identity-stat-grid">\n    <div class="identity-stat">Company Rank<b>#{{performance['rank']}}</b></div>\n    <div class="identity-stat">Performance Score<b>{{performance['score']}}%</b></div>\n    <div class="identity-stat">Positive Leads<b>{{performance['positive']}}</b></div>\n    <div class="identity-stat">Hot Leads<b>{{performance['hot']}}</b></div>\n  </div>\n  <p>\n    {% if performance['score']>=85 %}<span class="level-badge level-star">⭐ STAR PERFORMER</span>\n    {% elif performance['score']>=70 %}<span class="level-badge level-strong">STRONG PERFORMER</span>\n    {% elif performance['score']>=50 %}<span class="level-badge level-good">GOOD / IMPROVING</span>\n    {% else %}<span class="level-badge level-build">BUILDING PERFORMANCE</span>{% endif %}\n  </p>\n</div>\n{% endif %}\n{% endblock %}""", 'team_identities.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">Team Identity Center</h1>\n<div class="card">\n  <h2 style="color:#e6b73f;margin-top:0">Create Professional Portal Identity</h2>\n  <p>Create a permanent employee identity first. Login account, official mobile, employee record and future OTP verification stay linked to the same identity.</p>\n  <form method="post" enctype="multipart/form-data">\n    <div class="grid">\n      {% if u['role']=='MD' %}\n      <div><label>Company</label><select name="company_code" required><option value="SCIC">ABC</option><option value="WWIC">XYZ</option></select></div>\n      <div><label>Portal Role</label><select name="portal_role" required><option value="GM">General Manager</option><option value="AM">Assistant Manager</option><option value="RECEPTION">Reception</option></select></div>\n      {% else %}\n      <input type="hidden" name="company_code" value="{{u['company_code']}}">\n      <div><label>Portal Role</label><select name="portal_role" required><option value="AM">Assistant Manager</option><option value="RECEPTION">Reception</option></select></div>\n      {% endif %}\n      <div><label>Full Name</label><input name="full_name" required></div>\n      <div><label>Official Mobile Number</label><input name="official_mobile" required placeholder="10-digit official mobile"></div>\n      <div><label>Official Email</label><input type="email" name="official_email"></div>\n      <div><label>Designation</label><input name="designation" placeholder="AM / Counselor / Telecaller / GM"></div>\n      <div><label>Department</label><input name="department" placeholder="Sales / Counseling / Operations"></div>\n      <div><label>Branch</label><input name="branch" placeholder="Chandigarh / Mohali"></div>\n      <div><label>Joining Date</label><input type="date" name="joining_date" placeholder="dd-mm-yyyy"></div>\n      <div><label>Reporting Manager</label><select name="reporting_user_id"><option value="">Management / None</option>{% for r in managers %}<option value="{{r['id']}}">{{r['full_name']}} • {{r['role']}}</option>{% endfor %}</select></div>\n      <div><label>Initial Password / PIN</label><input type="password" name="password" minlength="6" required placeholder="Minimum 6 characters"></div>\n      <div><label>Profile Photograph</label><input type="file" name="photograph" accept=".jpg,.jpeg,.png,.webp,image/*"></div>\n    </div>\n    <button class="btn">Create Professional Identity</button>\n  </form>\n  <p style="opacity:.75;margin-bottom:0">OTP Status will remain <b>Verification Pending</b> until the SMS/OTP gateway is activated in the next security build.</p>\n</div>\n\n<div class="card">\n  <h2 style="color:#e6b73f;margin-top:0">Professional Team Directory</h2>\n  <div class="tablewrap"><table>\n    <tr><th>Profile</th><th>Employee ID</th><th>Name / Role</th><th>Official Contact</th><th>Branch</th><th>Security</th><th>Status</th><th>Action</th></tr>\n    {% for r in rows %}\n    <tr>\n      <td>{% if r['photo_mime'] %}<img class="identity-table-photo" src="{{url_for('user_photo',user_id=r['id'])}}">{% else %}<div class="identity-placeholder" style="width:48px;height:48px;font-size:18px">{{(r['full_name'] or '?')[0]|upper}}</div>{% endif %}</td>\n      <td><b style="color:#e6b73f">{{r['employee_code'] or r['identity_code'] or '-'}}</b><br><small>{{r['login_id']}}</small></td>\n      <td><b>{{r['full_name']}}</b><br>{{r['designation'] or r['role']}} • {{r['company_code']}}</td>\n      <td>{{r['official_mobile'] or '-'}}<br><small>{{r['official_email'] or ''}}</small></td>\n      <td>{{r['branch'] or '-'}}</td>\n      <td>{% if r['mobile_verified'] %}<span class="identity-badge green">Verified</span>{% else %}<span class="identity-badge">OTP Pending</span>{% endif %}</td>\n      <td>{% if r['active'] %}<span class="identity-badge green">Active</span>{% else %}<span class="identity-badge red">Inactive</span>{% endif %}</td>\n      <td class="identity-actions"><a class="toolbtn" href="{{url_for('management_identity_profile',user_id=r['id'])}}">Open Profile</a><a class="toolbtn" href="{{url_for('digital_id_card',user_id=r['id'])}}">ID Card</a>\n{% if r['active'] %}<form method="post" action="{{url_for('deactivate_team_user',user_id=r['id'])}}" style="display:inline" onsubmit="return confirm('Deactivate this employee login? Their history will remain safe.')"><button class="toolbtn" type="submit">Deactivate ID</button></form>{% else %}<form method="post" action="{{url_for('reactivate_team_user',user_id=r['id'])}}" style="display:inline"><button class="toolbtn" type="submit">Reactivate ID</button></form>{% endif %}\n{% if u['role']=='MD' and not r['active'] %}<form method="post" action="{{url_for('permanent_delete_team_user',user_id=r['id'])}}" style="display:inline" onsubmit="return confirm('Permanently delete this inactive portal ID? This is allowed only when no operational records are linked.')"><button class="toolbtn danger" type="submit">Delete ID</button></form>{% endif %}</td>\n    </tr>\n    {% endfor %}\n  </table></div>\n</div>\n{% endblock %}""", 'management_identity_profile.html': """{% extends "base.html" %}{% block content %}\n<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap">\n  <h1 class="title">Professional Employee Identity</h1>\n  <div class="identity-actions"><a class="toolbtn" href="{{url_for('team_identities')}}">← Team Identity Center</a><a class="btn" href="{{url_for('digital_id_card',user_id=person['id'])}}">Digital ID Card</a></div>\n</div>\n<div class="identity-hero">\n  <div class="identity-profile">\n    <div style="text-align:center">\n      {% if person['photo_mime'] %}<img class="identity-avatar big" src="{{url_for('user_photo',user_id=person['id'])}}">{% else %}<div class="identity-placeholder big" style="margin:auto">{{(person['full_name'] or '?')[0]|upper}}</div>{% endif %}\n    </div>\n    <div>\n      <div class="identity-name">{{person['full_name']}}</div>\n      <div style="font-size:19px">{{person['designation'] or person['role']}} • {{person['company_code']}}</div>\n      <div style="margin-top:8px">\n        <span class="identity-badge gold">{{employee['employee_code'] if employee else person['identity_code']}}</span>\n        <span class="identity-badge">{{person['role']}}</span>\n        {% if person['mobile_verified'] %}<span class="identity-badge green">✓ Mobile Verified</span>{% else %}<span class="identity-badge">OTP Verification Pending</span>{% endif %}\n      </div>\n      {% if person['profile_headline'] %}<p>{{person['profile_headline']}}</p>{% endif %}\n    </div>\n  </div>\n</div>\n\n<div class="card" style="margin-top:16px">\n<h2 style="color:#e6b73f;margin-top:0">Management-Controlled Identity Fields</h2>\n<form method="post" enctype="multipart/form-data">\n<div class="grid">\n  <div><label>Full Name</label><input name="full_name" value="{{person['full_name']}}" required></div>\n  <div><label>Official Mobile</label><input name="official_mobile" value="{{person['official_mobile'] or ''}}" required></div>\n  <div><label>Official Email</label><input type="email" name="official_email" value="{{person['official_email'] or ''}}"></div>\n  <div><label>Designation</label><input name="designation" value="{{person['designation'] or ''}}"></div>\n  <div><label>Department</label><input name="department" value="{{person['department'] or ''}}"></div>\n  <div><label>Branch</label><input name="branch" value="{{person['branch'] or ''}}"></div>\n  <div><label>Joining Date</label><input type="date" name="joining_date" placeholder="dd-mm-yyyy" value="{{person['joining_date'] or ''}}"></div>\n  <div><label>Reporting Manager</label><select name="reporting_user_id"><option value="">None / Management</option>{% for r in managers %}<option value="{{r['id']}}" {% if person['reporting_user_id']==r['id'] %}selected{% endif %}>{{r['full_name']}} • {{r['role']}}</option>{% endfor %}</select></div>\n  <div><label>Replace Profile Photo</label><input type="file" name="photograph" accept=".jpg,.jpeg,.png,.webp,image/*"></div>\n  <div><label>Account Status</label><select name="active"><option value="1" {% if person['active'] %}selected{% endif %}>Active</option><option value="0" {% if not person['active'] %}selected{% endif %}>Inactive</option></select></div>\n</div>\n<button class="btn">Save Official Identity</button>\n</form>\n<div style="margin-top:16px;border-top:1px solid #315f85;padding-top:14px">\n<h3 style="color:#e6b73f">Employee Offboarding</h3>\n{% if person['active'] %}\n<form method="post" action="{{url_for('deactivate_team_user',user_id=person['id'])}}" onsubmit="return confirm('Deactivate this employee login now?')"><button class="btn danger">Deactivate Employee ID</button></form>\n{% else %}\n<form method="post" action="{{url_for('reactivate_team_user',user_id=person['id'])}}" style="display:inline"><button class="btn">Reactivate Employee ID</button></form>\n{% if u['role']=='MD' %}<form method="post" action="{{url_for('permanent_delete_team_user',user_id=person['id'])}}" style="display:inline" onsubmit="return confirm('Permanently delete this inactive ID? Only IDs without linked operational records can be erased.')"><button class="btn danger">Permanently Delete ID</button></form>{% endif %}\n{% endif %}\n<p style="opacity:.7;font-size:12px">Deactivation blocks login immediately but preserves leads, performance and audit history. Permanent ID deletion is MD-only and is blocked when operational records are linked.</p>\n</div>\n</div>\n{% endblock %}""", 'digital_id_card.html': """{% extends "base.html" %}{% block content %}\n<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap" class="no-print">\n  <h1 class="title">Digital Employee ID Card</h1>\n  <button class="btn" onclick="window.print()">Print / Save ID Card</button>\n</div>\n<div class="id-card-shell">\n  <div class="id-card-top">\n    <div><div style="font-size:30px;color:#e6b73f;font-weight:900">THE GAUR</div><div>Verified Enterprise Identity</div></div>\n    <div style="text-align:right"><b>{{person['company_code'] or 'MANAGEMENT'}}</b><br><small>Immigration CRM</small></div>\n  </div>\n  <div class="id-card-body">\n    <div style="text-align:center">\n      {% if person['photo_mime'] %}<img class="id-card-photo" src="{{url_for('user_photo',user_id=person['id'])}}">{% else %}<div class="identity-placeholder big">{{(person['full_name'] or '?')[0]|upper}}</div>{% endif %}\n    </div>\n    <div>\n      <div class="identity-name">{{person['full_name']}}</div>\n      <div style="font-size:20px">{{person['designation'] or person['role']}}</div>\n      <div class="id-card-code">{{employee['employee_code'] if employee else person['identity_code']}}</div>\n      <div class="grid" style="grid-template-columns:1fr 1fr;margin-top:12px">\n        <div><small>Official Mobile</small><br><b>{{person['official_mobile'] or '-'}}</b></div>\n        <div><small>Branch</small><br><b>{{person['branch'] or '-'}}</b></div>\n        <div><small>Joining Date</small><br><b>{{person['joining_date'] or '-'}}</b></div>\n        <div><small>Role</small><br><b>{{person['role']}}</b></div>\n      </div>\n      <div class="id-verify-box">\n        {% if person['mobile_verified'] %}✓ Official mobile verified{% else %}OTP Security Ready • Mobile verification pending activation{% endif %}\n        <br>Account Status: <b>{{'ACTIVE' if person['active'] else 'INACTIVE'}}</b>\n      </div>\n    </div>\n  </div>\n</div>\n{% endblock %}""", 'activity_center.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">Activity & Security Center</h1><p style="opacity:.72">{{scope_label}} • Complete portal audit trail</p>\n<div class="audit-kpis">\n<div class="audit-kpi"><span>Logins Today</span><b>{{kpi.logins}}</b></div><div class="audit-kpi"><span>Logouts Today</span><b>{{kpi.logouts}}</b></div><div class="audit-kpi"><span>Client / Lead Updates</span><b>{{kpi.client}}</b></div><div class="audit-kpi"><span>Chat / Broadcast</span><b>{{kpi.chat}}</b></div><div class="audit-kpi"><span>Admin / Employee Actions</span><b>{{kpi.admin}}</b></div><div class="audit-kpi critical"><span>Critical / Failed</span><b>{{kpi.critical}}</b></div>\n</div>\n<div class="card"><h2 style="color:#e6b73f;margin-top:0">Filters</h2><form method="get" class="audit-filters">\n<div><label>From</label><input type="date" name="date_from" value="{{f.date_from}}"></div><div><label>To</label><input type="date" name="date_to" value="{{f.date_to}}"></div>\n{% if u['role']=='MD' %}<div><label>Company</label><select name="company"><option value="">Both</option><option value="SCIC" {% if f.company=='SCIC' %}selected{% endif %}>ABC</option><option value="WWIC" {% if f.company=='WWIC' %}selected{% endif %}>XYZ</option></select></div>{% endif %}\n<div><label>Employee</label><select name="user_id"><option value="">All Employees</option>{% for p in people %}<option value="{{p.id}}" {% if f.user_id==p.id|string %}selected{% endif %}>{{p.full_name}} • {{role_display(p.role,p.designation)}}</option>{% endfor %}</select></div>\n<div><label>Activity</label><select name="category"><option value="">All</option>{% for c in categories %}<option value="{{c}}" {% if f.category==c %}selected{% endif %}>{{c}}</option>{% endfor %}</select></div>\n<div><label>Severity</label><select name="severity"><option value="">All</option><option value="CRITICAL" {% if f.severity=='CRITICAL' %}selected{% endif %}>Critical</option><option value="WARNING" {% if f.severity=='WARNING' %}selected{% endif %}>Warning</option><option value="INFO" {% if f.severity=='INFO' %}selected{% endif %}>Info</option></select></div>\n<div><label>Search</label><input name="q" value="{{f.q}}" placeholder="Employee, client, lead, action..."></div><div><button class="btn">Apply</button> <a class="toolbtn" href="{{url_for('activity_center')}}">Reset</a></div>\n</form></div>\n<div class="card"><h2 style="color:#e6b73f;margin-top:0">Activity Timeline</h2>\n{% for r in rows %}<div class="audit-event {% if r.severity=='CRITICAL' %}critical{% elif r.category=='SECURITY' %}security{% elif r.category in ['CLIENT','LEAD'] %}client{% elif r.category in ['ADMIN','EMPLOYEE'] %}admin{% endif %}">\n<div class="audit-title">{% if r.severity=='CRITICAL' %}🚨{% elif r.category=='SECURITY' %}🔐{% elif r.category in ['CLIENT','LEAD'] %}📝{% elif r.category=='CHAT' %}💬{% elif r.category in ['ADMIN','EMPLOYEE'] %}👤{% else %}•{% endif %} {{r.action_label}}</div>\n<span class="audit-badge">{{r.category}}</span><span class="audit-badge">{{r.severity}}</span>{% if r.company_code %}<span class="audit-badge">{{r.company_code}}</span>{% endif %}\n<div style="margin-top:7px"><b>{{r.user_name or 'Anonymous / System'}}</b>{% if r.user_role %} • {{role_display(r.user_role,r.user_designation)}}{% endif %}{% if r.entity_label %}<br>{{r.entity_type}}: <b>{{r.entity_label}}</b>{% endif %}{% if r.details %}<br>{{r.details}}{% endif %}</div>\n{% if r.old_value or r.new_value %}<div class="audit-change"><div class="audit-value"><small>OLD</small><br>{{r.old_value or '-'}}</div><div class="audit-arrow">→</div><div class="audit-value"><small>NEW</small><br>{{r.new_value or '-'}}</div></div>{% endif %}\n<div class="audit-meta">{{r.created_at}}{% if r.ip_address %} • IP {{r.ip_address}}{% endif %} • {{r.request_method}} {{r.request_path}}</div></div>\n{% else %}<div style="opacity:.7;padding:20px;text-align:center">No activities found.</div>{% endfor %}</div>{% endblock %}""", 'finance.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">💰 Daily Accounts & Expense</h1>\n<div class="finance-tabs"><a href="{{url_for('finance_center')}}">Ledger</a><a href="{{url_for('finance_banks')}}">Bank & Cash Master</a><a href="{{url_for('finance_export',date_from=f.date_from,date_to=f.date_to,company=f.company)}}">⬇ Excel Export</a></div>\n<div class="finance-kpis">\n<div class="finance-kpi"><span>Money Received</span><b>₹{{'%.2f'|format(k.income)}}</b></div>\n<div class="finance-kpi"><span>Office / Client Payments</span><b>₹{{'%.2f'|format(k.expense)}}</b></div>\n<div class="finance-kpi"><span>Refund Paid</span><b>{% if u.role=='MD' %}₹{{'%.2f'|format(k.refund)}}{% else %}MD Confidential{% endif %}</b></div>\n<div class="finance-kpi"><span>Net Cash Flow</span><b>₹{{'%.2f'|format(k.net)}}</b></div>\n<div class="finance-kpi"><span>Transactions</span><b>{{k.count}}</b></div>\n<div class="finance-kpi"><span>Bill Pending</span><b>{{k.bill_pending}}</b></div>\n</div>\n\n<div class="card"><h2 style="color:#e6b73f">New Transaction</h2>\n<form method="post" enctype="multipart/form-data" class="finance-form">\n<div><label>Transaction</label><select name="direction" required><option value="IN">Money In / Receipt</option><option value="OUT">Money Out / Payment</option><option value="REFUND">Client Refund</option></select></div>\n<div><label>Date</label><input type="date" name="txn_date" value="{{today}}" required></div>\n<div><label>Company</label><select name="company_code" required>{% for c in companies %}<option value="{{c}}">{{'ABC Immigration Consultants' if c=='SCIC' else 'XYZ'}}</option>{% endfor %}</select></div>\n<div><label>Expense / Receipt Head</label><select name="head" required>{% for h in heads %}<option>{{h}}</option>{% endfor %}</select></div>\n<div><label>Client / Payee / Vendor</label><input name="party_name" required></div>\n<div><label>Client / Lead ID</label><input name="client_ref" placeholder="Optional"></div>\n<div><label>Amount ₹</label><input type="number" step="0.01" min="0.01" name="amount" required></div>\n<div><label>Payment Mode</label><select name="payment_mode"><option>Cash</option><option>UPI</option><option>Bank Transfer</option><option>NEFT</option><option>RTGS</option><option>Cheque</option><option>Card</option></select></div>\n<div><label>Bank / Cash Account</label><select name="account_id"><option value="">Cash / Not Selected</option>{% for b in banks %}<option value="{{b.id}}">{{b.bank_name}} • {{b.account_name}} • ****{{b.account_last4}}</option>{% endfor %}</select></div>\n<div><label>UTR / Cheque / Reference</label><input name="reference_no"></div>\n<div><label>GST Taxable Value</label><input type="number" step="0.01" name="taxable_value" value="0"></div>\n<div><label>GST Amount</label><input type="number" step="0.01" name="gst_amount" value="0"></div>\n<div style="grid-column:span 2"><label>Clear Remarks / Particulars</label><textarea name="remarks" placeholder="e.g. Canada Visitor Visa – Embassy Fee – Client XYZ – HDFC – UTR..."></textarea></div>\n<div style="grid-column:span 2"><label>Bill / Receipt / Proof</label><input type="file" name="bill" accept=".jpg,.jpeg,.png,.webp,.pdf"></div>\n<div><button class="btn">Save Transaction</button></div>\n</form></div>\n\n<div class="card"><form method="get" class="finance-form">\n<div><label>From</label><input type="date" name="date_from" value="{{f.date_from}}"></div><div><label>To</label><input type="date" name="date_to" value="{{f.date_to}}"></div>\n{% if u.role=='MD' %}<div><label>Company</label><select name="company"><option value="">Both</option><option value="SCIC" {% if f.company=='SCIC' %}selected{% endif %}>ABC</option><option value="WWIC" {% if f.company=='WWIC' %}selected{% endif %}>XYZ</option></select></div>{% endif %}\n<div><label>Head</label><select name="head"><option value="">All Heads</option>{% for h in heads %}<option {% if f.head==h %}selected{% endif %}>{{h}}</option>{% endfor %}</select></div>\n<div><button class="btn">Filter</button></div></form></div>\n\n<div class="card"><h2 style="color:#e6b73f">Transaction Ledger</h2><div class="tablewrap"><table><thead><tr><th>Voucher</th><th>Date</th><th>Company</th><th>Type</th><th>Head</th><th>Party / Client</th><th>Mode / Account</th><th>Money In</th><th>Money Out</th><th>Remarks</th><th>Proof</th><th>Receipt</th></tr></thead><tbody>\n{% for r in rows %}<tr class="{% if r.direction=='IN' %}money-in{% elif r.direction=='REFUND' %}refund-row{% else %}money-out{% endif %}">\n<td><b>{{r.voucher_no}}</b></td><td>{{r.txn_date}}</td><td>{{r.company_code}}</td><td>{{r.direction}}</td><td>{{r.head}}</td><td>{{r.party_name}}{% if r.client_ref %}<br><small>{{r.client_ref}}</small>{% endif %}</td><td>{{r.payment_mode}}{% if r.bank_name %}<br><small>{{r.bank_name}} • ****{{r.account_last4}}</small>{% endif %}</td>\n<td>{% if r.direction=='IN' %}₹{{'%.2f'|format(r.amount)}}{% endif %}</td><td>{% if r.direction!='IN' %}{% if u.role=='MD' or r.direction!='REFUND' %}₹{{'%.2f'|format(r.amount)}}{% else %}Confidential{% endif %}{% endif %}</td><td>{{r.remarks}}</td>\n<td>{% if r.bill_name %}<a class="toolbtn" href="{{url_for('finance_bill',txn_id=r.id)}}">View Bill</a>{% else %}<span style="color:#ff8d96">Pending</span>{% endif %}</td>\n<td>{% if r.direction=='IN' %}<a class="toolbtn" href="{{url_for('finance_receipt',txn_id=r.id)}}">GST Receipt</a> <a class="toolbtn" target="_blank" href="{{url_for('finance_whatsapp',txn_id=r.id)}}">WhatsApp</a>{% endif %}</td></tr>{% else %}<tr><td colspan="12">No transactions.</td></tr>{% endfor %}\n</tbody></table></div></div>{% endblock %}""", 'finance_banks.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">🏦 Bank Manager</h1>\n<div class="finance-tabs"><a href="{{url_for('finance_center')}}">← Daily Passbook</a> <a href="{{url_for('bank_share_history')}}">Share History</a> <a href="{{url_for('payment_accounts')}}">Employee Payment View</a></div>\n{% if u.role=='MD' %}<div class="card"><h2 style="color:#e6b73f">Add Approved Company Bank Account</h2><form method="post" class="finance-form">\n<div><label>Company</label><select name="company_code"><option value="SCIC">ABC</option><option value="WWIC">XYZ</option></select></div>\n<div><label>Bank Name</label><input name="bank_name" id="bankName" required placeholder="Search / type HDFC Bank, SBI, ICICI..."></div>\n<div><label>Account Nickname</label><input name="account_nickname" placeholder="Main Collection / Embassy Fees"></div>\n<div><label>Account Name</label><input name="account_name" required></div>\n<div><label>Account Number</label><input name="account_number" required></div>\n<div><label>IFSC Code</label><input name="ifsc_code"></div>\n<div><label>Branch</label><input name="branch"></div><div><label>City</label><input name="city"></div>\n<div><label>UPI ID</label><input name="upi_id" placeholder="company@bank"></div>\n<div><label>Account Type</label><select name="account_type"><option>Current</option><option>Savings</option><option>Cash</option></select></div>\n<div><label>Opening Balance</label><input type="number" step="0.01" name="opening_balance" value="0"></div>\n<div><label>Bank RM Name</label><input name="rm_name"></div><div><label>RM Mobile</label><input name="rm_mobile"></div>\n<div><button class="btn">Add Approved Account</button></div></form></div>{% endif %}\n<div class="card"><h2 style="color:#e6b73f">Approved Accounts</h2>\n<div class="bank-manager-grid">{% for b in banks %}<div class="payment-bank-card"><h3>{{b.bank_name}} • {{b.account_nickname or b.account_name}}</h3>\n<span class="verified-account">✓ Approved Company Account</span>\n<div class="payment-details"><div><small>Company</small><b>{{b.company_code}}</b></div><div><small>Account Name</small><b>{{b.account_name}}</b></div>\n<div><small>Account Number</small><b>{{b.account_number if u.role=='MD' else '••••••••'+b.account_last4}}</b></div><div><small>IFSC</small><b>{{b.ifsc_code or '-'}}</b></div>\n<div><small>Branch / City</small><b>{{b.branch or '-'}} {{b.city or ''}}</b></div><div><small>UPI</small><b>{{b.upi_id or '-'}}</b></div></div><div style=\"margin-top:10px;padding-top:10px;border-top:1px solid #315a7b\"><form method=\"post\" action=\"{{url_for('bank_manager.bank_delete_request',bank_id=b.id)}}\"><button type=\"submit\" class=\"toolbtn\" style=\"border-color:#e46767;color:#ff9a9a;font-size:12px\" onclick=\"return confirm('Send OTP to delete this bank account?')\">🗑 Delete (OTP Required)</button></form>{% if request.args.get('pending_delete')==b.id|string %}<form method=\"post\" action=\"{{url_for('bank_manager.bank_delete_confirm',bank_id=b.id)}}\" style=\"margin-top:8px;display:flex;gap:8px;align-items:center\"><input name=\"otp\" maxlength=\"6\" placeholder=\"6-digit OTP from Gmail\" style=\"width:160px;padding:7px;background:#071d32;border:1px solid #e6b73f;color:#fff;border-radius:8px\" required><button class=\"btn\" style=\"padding:7px 12px;font-size:13px;background:#e46767;color:#fff\">✓ Confirm Delete</button></form>{% endif %}</div></div>\n{% else %}No approved bank accounts added.{% endfor %}</div></div>{% endblock %}""", 'finance_receipt.html': """{% extends "base.html" %}{% block content %}<div class="card" style="max-width:850px;margin:auto;background:#fff;color:#111">\n<div style="text-align:center"><h1 style="margin:0">{{company_name}}</h1><p>{{company_address}}</p><b>{% if gstin %}GSTIN: {{gstin}}{% endif %}</b><h2>PAYMENT RECEIPT</h2><p style="font-size:12px">Computer Generated • Verified Transaction Record</p></div><hr>\n<table style="min-width:0;color:#111"><tr><th>Receipt / Voucher No.</th><td>{{r.voucher_no}}</td><th>Date</th><td>{{r.txn_date}}</td></tr><tr><th>Received From</th><td>{{r.party_name}}</td><th>Client Ref.</th><td>{{r.client_ref or '-'}}</td></tr><tr><th>Particulars</th><td colspan="3">{{r.head}} • {{r.remarks}}</td></tr><tr><th>Amount Received</th><td>₹{{'%.2f'|format(r.amount)}}</td><th>Mode</th><td>{{r.payment_mode}}</td></tr><tr><th>Reference</th><td>{{r.reference_no or '-'}}</td><th>Receiving Account</th><td>{{r.bank_name or 'Cash'}} {% if r.account_last4 %}****{{r.account_last4}}{% endif %}</td></tr><tr><th>Taxable Value</th><td>₹{{'%.2f'|format(r.taxable_value or 0)}}</td><th>GST Amount</th><td>₹{{'%.2f'|format(r.gst_amount or 0)}}</td></tr></table>\n<p><b>Remarks:</b> {{r.remarks or '-'}}</p><p style="font-size:12px">Computer-generated receipt from IMS Portal. Tax treatment/GST values are based on the configured transaction entry and should be verified as per applicable accounting requirements.</p>\n<div class="no-print"><button class="btn" onclick="window.print()">Print / Save PDF</button> <a class="toolbtn" target="_blank" href="{{url_for('finance_whatsapp',txn_id=r.id)}}">Share on WhatsApp</a></div></div>{% endblock %}""", 'finance_reconciliation.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">🤖 AI Bank Reconciliation</h1>\n<p style="opacity:.72">Upload bank statement → THE GAUR matches bank credits/debits with portal receipts/payments using Amount, Date, Reference/UTR, Account and Narration.</p>\n\n<div class="recon-kpis">\n<div class="recon-kpi"><span>Statement Transactions</span><b>{{k.statement_count}}</b></div>\n<div class="recon-kpi"><span>Verified</span><b>{{k.verified}}</b></div>\n<div class="recon-kpi"><span>Possible Match</span><b>{{k.possible}}</b></div>\n<div class="recon-kpi"><span>Unmatched Portal Entries</span><b>{{k.portal_unmatched}}</b></div>\n<div class="recon-kpi"><span>Bank Entries Without Receipt</span><b>{{k.bank_unmatched}}</b></div>\n<div class="recon-kpi"><span>Duplicates / Suspicious</span><b>{{k.suspicious}}</b></div>\n</div>\n\n<div class="card">\n<h2 style="color:#e6b73f;margin-top:0">Upload Bank Statement</h2>\n<form method="post" enctype="multipart/form-data">\n<div class="finance-form">\n<div><label>Company</label><select name="company_code">{% for c in companies %}<option value="{{c}}">{{c}}</option>{% endfor %}</select></div>\n<div><label>Bank Account</label><select name="bank_id" required>{% for b in banks %}<option value="{{b.id}}">{{b.company_code}} • {{b.bank_name}} • {{b.account_name}} • ****{{b.account_last4}}</option>{% endfor %}</select></div>\n<div style="grid-column:span 2"><label>Statement File</label><input type="file" name="statement" accept=".xlsx,.xls,.csv" required><small>Best results: Excel/CSV with Date, Narration/Description, Reference/UTR, Debit, Credit, Balance.</small></div>\n<div><button class="btn">Upload & Reconcile</button></div>\n</div></form>\n</div>\n\n{% if run %}\n<div class="card"><h2 style="color:#e6b73f;margin-top:0">Reconciliation Run</h2>\n<b>{{run.file_name}}</b> • {{run.company_code}} • {{run.created_at}} • {{run.total_rows}} bank rows\n<a class="toolbtn" href="{{url_for('finance_reconciliation_export',run_id=run.id)}}">⬇ Reconciliation Excel</a>\n</div>\n{% endif %}\n\n<div class="recon-grid">\n<div>\n<h2 style="color:#e6b73f">Matched / Suggested</h2>\n{% for r in matches %}\n<div class="recon-card {% if r.status=='VERIFIED' %}verified{% elif r.status=='POSSIBLE' %}possible{% else %}unmatched{% endif %}">\n<div style="display:flex;justify-content:space-between;gap:10px"><b>{{r.portal_voucher or 'No Portal Voucher'}}</b><div class="recon-score">{{r.score}}%</div></div>\n<div><b>{{r.status}}</b> • {{r.match_reason}}</div>\n<div class="recon-detail">Portal: {{r.portal_date or '-'}} • ₹{{'%.2f'|format(r.portal_amount or 0)}} • {{r.portal_party or '-'}} • {{r.portal_reference or '-'}}</div>\n<div class="recon-detail">Bank: {{r.bank_date or '-'}}{% if r.bank_time %} {{r.bank_time}}{% endif %} • ₹{{'%.2f'|format(r.bank_amount or 0)}} • {{r.bank_type}} • {{r.bank_reference or '-'}} • {{r.bank_narration or '-'}}</div>\n<div class="recon-actions">\n{% if r.status!='VERIFIED' and r.portal_txn_id and r.bank_row_id %}\n<form method="post" action="{{url_for('finance_reconciliation_confirm')}}"><input type="hidden" name="match_id" value="{{r.id}}"><button class="btn">Confirm Match</button></form>\n{% endif %}\n{% if r.status=='VERIFIED' %}<span class="audit-badge">✓ Bank Verified</span>{% endif %}\n</div></div>\n{% else %}<div style="opacity:.7">No matches yet. Upload a statement.</div>{% endfor %}\n</div>\n\n<div>\n<h2 style="color:#e6b73f">Needs Attention</h2>\n{% for r in attention %}\n<div class="recon-card unmatched"><b>{{r.title}}</b><div class="recon-detail">{{r.details}}</div></div>\n{% else %}<div style="opacity:.7">No unmatched/suspicious items.</div>{% endfor %}\n</div>\n</div>\n{% endblock %}""", 'lead_control.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">🎯 Lead Allocation & Data Batch Control</h1>\n<p style="opacity:.76">Upload date/time, batch/file, uploaded by, allocation date/time/by and actual AM work are now visible together.</p>\n\n<div class="cards">\n<div class="card"><b>Visible Leads</b><h2>{{k.total}}</h2></div>\n<div class="card"><b>Latest Batch Data</b><h2>{{k.latest_batch}}</h2></div>\n<div class="card"><b>Old / Previous Data</b><h2>{{k.old_data}}</h2></div>\n<div class="card"><b>Unallocated</b><h2>{{k.unallocated}}</h2></div>\n<div class="card"><b>Not Worked</b><h2>{{k.not_worked}}</h2></div>\n<div class="card"><b>Overdue</b><h2>{{k.overdue}}</h2></div>\n</div>\n\n<div class="card"><h2 style="color:#e6b73f">Filters</h2><form method="get" class="filters">\n{% if u.role=='MD' %}<select name="company"><option value="">Both Companies</option><option value="SCIC" {% if f.company=='SCIC' %}selected{% endif %}>ABC</option><option value="WWIC" {% if f.company=='WWIC' %}selected{% endif %}>XYZ</option></select>{% endif %}\n<select name="batch"><option value="">All Upload Batches</option><option value="latest" {% if f.batch=='latest' %}selected{% endif %}>Latest Batch / New Data</option><option value="old" {% if f.batch=='old' %}selected{% endif %}>Old / Previous Data</option><option value="legacy" {% if f.batch=='legacy' %}selected{% endif %}>Legacy Data</option>{% for b in batches %}<option value="{{b.batch_code}}" {% if f.batch==b.batch_code %}selected{% endif %}>{{b.batch_code}} • {{b.filename}}</option>{% endfor %}</select>\n<select name="am"><option value="">All AMs</option><option value="unallocated" {% if f.am=='unallocated' %}selected{% endif %}>Unallocated</option>{% for a in am_rows %}<option value="{{a.id}}" {% if f.am==a.id|string %}selected{% endif %}>{{a.company_code}} • {{a.full_name}}</option>{% endfor %}</select>\n<select name="work"><option value="">All Work Status</option><option value="worked" {% if f.work=='worked' %}selected{% endif %}>Worked</option><option value="not_worked" {% if f.work=='not_worked' %}selected{% endif %}>Not Worked</option><option value="overdue" {% if f.work=='overdue' %}selected{% endif %}>Follow-up Overdue</option></select>\n<label>Uploaded From <input type="date" name="upload_from" value="{{f.upload_from}}"></label>\n<label>Uploaded To <input type="date" name="upload_to" value="{{f.upload_to}}"></label>\n<label>Allocated From <input type="date" name="alloc_from" value="{{f.alloc_from}}"></label>\n<label>Allocated To <input type="date" name="alloc_to" value="{{f.alloc_to}}"></label>\n<input name="q" value="{{f.q}}" placeholder="Name / Mobile / Lead ID / Batch"><button class="btn">Apply Filters</button>\n</form></div>\n\n<div class="card"><h2 style="color:#e6b73f">📦 Upload Batch Summary</h2><div class="tablewrap"><table><thead><tr><th>Type</th><th>Batch</th><th>Company</th><th>Uploaded</th><th>Uploaded By</th><th>File</th><th>Total</th><th>Allocated</th><th>Unallocated</th><th>Worked</th><th>Not Worked</th><th>Positive</th></tr></thead><tbody>\n{% for b in batch_summary %}<tr><td>{% if b.is_latest %}<span class="batch-new">NEW / LATEST</span>{% else %}<span class="batch-old">OLD DATA</span>{% endif %}</td><td><b>{{b.batch_code}}</b></td><td>{{b.company_code}}</td><td>{{b.imported_at}}</td><td>{{b.imported_by}}</td><td>{{b.filename}}</td><td>{{b.total}}</td><td>{{b.allocated}}</td><td>{{b.unallocated}}</td><td>{{b.worked}}</td><td>{{b.not_worked}}</td><td>{{b.positive}}</td></tr>\n{% else %}<tr><td colspan="12">No tracked batches yet. The next Excel upload will create one automatically.</td></tr>{% endfor %}</tbody></table></div></div>\n\n<form method="post" action="{{url_for('lead_bulk_action')}}">\n<div class="card" style="position:sticky;top:0;z-index:5"><label><input type="checkbox" id="all"> Select All Visible</label>\n<select name="am_id"><option value="">Choose Assistant Manager</option>{% for a in am_rows %}<option value="{{a.id}}">{{a.company_code}} • {{a.full_name}}</option>{% endfor %}</select>\n<input name="reason" placeholder="Allocation / deletion reason"><button class="btn" name="action" value="allocate">Allocate / Re-allocate Selected</button><button class="btn danger" name="action" value="archive" onclick="return confirm('Move selected leads to Recycle Bin? This will be audited.')">Delete Selected</button></div>\n\n<div class="card tablewrap"><table><thead><tr><th>Select</th><th>Lead / Data Age</th><th>Client</th><th>Upload Details</th><th>Allocation Details</th><th>AM</th><th>Interest</th><th>Work</th><th>Follow-up</th><th>Profile</th></tr></thead><tbody>\n{% for r in rows %}<tr style="{% if r.work_state=='NOT WORKED' %}box-shadow:inset 5px 0 #ff6571{% elif r.work_state=='OVERDUE' %}box-shadow:inset 5px 0 #e6b73f{% else %}box-shadow:inset 5px 0 #55dca0{% endif %}">\n<td><input class="leadbox" type="checkbox" name="lead_ids" value="{{r.id}}"></td>\n<td><b>{{r.lead_id}}</b><br>{% if r.data_age=='LATEST' %}<span class="batch-new">NEW BATCH</span>{% elif r.data_age=='LEGACY' %}<span class="batch-legacy">LEGACY</span>{% else %}<span class="batch-old">OLD DATA</span>{% endif %}</td>\n<td>{{r.client_name}}<br><small>{{r.mobile}}</small></td>\n<td><div class="lifecycle-meta"><b>Batch:</b> {{r.upload_batch or 'Legacy / Untracked'}}<br><b>Uploaded:</b> {{r.imported_at or '-'}}<br><b>By:</b> {{r.imported_by or '-'}}<br><b>File:</b> {{r.import_filename or '-'}}</div></td>\n<td><div class="lifecycle-meta">{% if r.assigned_am %}<b>Allocated:</b> {{r.assigned_at or '-'}}<br><b>Allocated By:</b> {{r.assigned_by or '-'}}{% else %}<b style="color:#e6b73f">NOT ALLOCATED</b>{% endif %}</div></td>\n<td>{{r.am_name or 'UNALLOCATED'}}</td><td>{{r.interest_score or 0}}%</td>\n<td><b>{{r.work_state}}</b><br><small>{{r.activity_count}} update(s)<br>Last: {{r.last_activity_at or '-'}}</small></td><td>{{r.followup_date or '-'}}</td><td><a class="toolbtn" href="{{url_for('lead_profile',lead_db_id=r.id)}}">Open</a></td></tr>\n{% else %}<tr><td colspan="10">No matching leads.</td></tr>{% endfor %}</tbody></table></div></form>\n\n<div class="card"><h2 style="color:#e6b73f">Assistant Manager Workload</h2><div class="tablewrap"><table><thead><tr><th>AM</th><th>Company</th><th>Allocated</th><th>Worked</th><th>Not Worked</th><th>Hot</th><th>Positive</th><th>Overdue</th><th>Work %</th></tr></thead><tbody>\n{% for a in summary %}<tr><td><b>{{a.full_name}}</b></td><td>{{a.company_code}}</td><td>{{a.allocated}}</td><td>{{a.worked}}</td><td>{{a.not_worked}}</td><td>{{a.hot}}</td><td>{{a.positive}}</td><td>{{a.overdue}}</td><td>{{a.work_pct}}%</td></tr>{% endfor %}</tbody></table></div></div>\n\n{% if u.role in ['MD','GM'] %}<div class="card"><h2>🗑 Recycle Bin</h2>\n<p style="opacity:.75">Archived leads are excluded from active dashboard/performance totals. Restore returns them to active calculations. Permanent erase cannot be undone.</p>\n<form method="post" action="{{url_for('recycle_bin_bulk_action')}}">\n<div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:12px">\n<label><input type="checkbox" id="recycleAll"> Select All</label>\n<button class="btn" name="action" value="restore">Restore Selected</button>\n<button class="btn danger" name="action" value="erase" onclick="return confirm('PERMANENT DELETE: Selected lead data will be erased and cannot be restored. Continue?')">Permanently Delete Selected</button>\n{% if archived %}<button class="btn danger" name="action" value="empty" onclick="return confirm('EMPTY RECYCLE BIN: All leads visible to your management scope will be permanently erased. This cannot be undone. Continue?')">Empty Recycle Bin</button>{% endif %}\n</div>\n<div class="tablewrap"><table><thead><tr><th>Select</th><th>Lead</th><th>Batch</th><th>Client</th><th>Company</th><th>Deleted By</th><th>Deleted At</th><th>Reason</th></tr></thead><tbody>\n{% for r in archived %}<tr><td><input class="recyclebox" type="checkbox" name="lead_ids" value="{{r.id}}"></td><td>{{r.lead_id}}</td><td>{{r.upload_batch or 'Legacy'}}</td><td>{{r.client_name}}</td><td>{{r.company_code}}</td><td>{{r.deleted_by}}</td><td>{{r.deleted_at}}</td><td>{{r.deletion_reason}}</td></tr>\n{% else %}<tr><td colspan="8">Recycle Bin empty.</td></tr>{% endfor %}</tbody></table></div></form></div>\n<script>var ra=document.getElementById('recycleAll');if(ra){ra.addEventListener('change',function(){document.querySelectorAll('.recyclebox').forEach(x=>x.checked=this.checked)})}</script>{% endif %}\n<script>document.getElementById('all').addEventListener('change',function(){document.querySelectorAll('.leadbox').forEach(x=>x.checked=this.checked)});</script>\n{% endblock %}""", 'attendance.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">🕘 Employee Attendance • Reception</h1>\n<p style="opacity:.75">Reception can record employee office-entry time. MD/GM can also review the same company attendance.</p>\n\n<div class="card">\n<form method="post" class="attendance-top">\n{% if u['role']=='MD' %}<div><label>Company</label><select name="company_code"><option value="SCIC" {% if company=='SCIC' %}selected{% endif %}>ABC</option><option value="WWIC" {% if company=='WWIC' %}selected{% endif %}>XYZ</option></select></div>{% else %}<input type="hidden" name="company_code" value="{{company}}">{% endif %}\n<div class="wide"><label>Employee Name</label><select name="employee_id" required><option value="">Select Employee</option>{% for e in employees %}<option value="{{e.id}}">{{e.full_name}} • {{e.designation or e.department or 'Employee'}}</option>{% endfor %}</select></div>\n<div><label>Date</label><input type="date" name="attendance_date" value="{{today}}" required></div>\n<div><label>Office Entry Time</label><input type="time" name="entry_time" value="{{now_time}}" required></div>\n<div><label>Remarks</label><input name="remarks" placeholder="Optional"></div>\n<div><button class="btn" name="save_attendance" value="1">Save Attendance</button></div>\n</form>\n</div>\n\n<div class="card">\n<span class="attendance-stat"><b>{{rows|length}}</b> Marked Today</span>\n<span class="attendance-stat"><b>{{employees|length}}</b> Active Employees</span>\n<h2 style="color:#e6b73f">Today's Attendance • {{today}}</h2>\n<div class="tablewrap"><table><thead><tr><th>Sr.</th><th>Employee</th><th>Designation</th><th>Company</th><th>Entry Time</th><th>Marked By</th><th>Remarks</th><th>Good Morning</th></tr></thead><tbody>\n{% for r in rows %}<tr class="attendance-present"><td>{{loop.index}}</td><td><b>{{r.full_name}}</b></td><td>{{r.designation or '-'}}</td><td>{{r.company_code}}</td><td><b>{{r.entry_time}}</b></td><td>{{r.marked_by}}</td><td>{{r.remarks or '-'}}</td>\n<td>{% if r.mobile %}<a class="toolbtn wa-mini" target="_blank" href="{{r.whatsapp_url}}">WhatsApp</a>{% else %}<small>No mobile</small>{% endif %}</td></tr>\n{% else %}<tr><td colspan="8">No attendance marked yet for today.</td></tr>{% endfor %}\n</tbody></table></div>\n</div>\n\n<div class="card">\n<h2 style="color:#e6b73f">Not Yet Marked Today</h2>\n<div class="tablewrap"><table><thead><tr><th>Employee</th><th>Designation</th><th>Official Mobile</th></tr></thead><tbody>\n{% for e in missing %}<tr class="attendance-missing"><td><b>{{e.full_name}}</b></td><td>{{e.designation or '-'}}</td><td>{{e.mobile or '-'}}</td></tr>\n{% else %}<tr><td colspan="3">All active employees are marked.</td></tr>{% endfor %}\n</tbody></table></div>\n</div>\n{% endblock %}""", 'payment_accounts.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">🏦 Company Payment Accounts</h1>\n<p style="opacity:.78">Only approved company accounts are shown here. Use the official payment card / QR and WhatsApp message when sharing details with a client.</p>\n<div class="bank-manager-grid">\n{% for b in banks %}<div class="payment-bank-card" data-bank="{{b.id}}">\n<div class="payment-bank-head">\n{% if b.logo_url %}<img class="bank-auto-logo" src="{{b.logo_url}}" alt="{{b.bank_name}}" onerror="this.style.display='none';this.nextElementSibling.style.display='flex'"><div class="bank-logo-fallback" style="display:none">{{b.bank_name[:2]|upper}}</div>\n{% else %}<div class="bank-logo-fallback">{{b.bank_name[:2]|upper}}</div>{% endif %}\n<div><h2 style="margin:0;color:#e6b73f">{{b.bank_name}}</h2><div>{{b.account_nickname or b.account_name}}</div><span class="verified-account">✓ Verified Company Account</span></div>\n</div>\n<div class="payment-details">\n<div><small>Company</small><b>{{'ABC Immigration Consultants' if b.company_code=='SCIC' else 'XYZ'}}</b></div>\n<div><small>Account Name</small><b>{{b.account_name}}</b></div>\n<div><small>Account Number</small><b>{{b.account_number}}</b></div>\n<div><small>IFSC Code</small><b>{{b.ifsc_code or '-'}}</b></div>\n<div><small>Branch</small><b>{{b.branch or '-'}}{% if b.city %}, {{b.city}}{% endif %}</b></div>\n<div><small>UPI ID</small><b>{{b.upi_id or 'Not configured'}}</b></div>\n</div>\n<div class="payment-actions">\n{% if b.upi_id %}<button class="toolbtn" type="button" onclick="toggleQR({{b.id}},'upi://pay?pa={{b.upi_id|urlencode}}&pn={{b.account_name|urlencode}}&cu=INR')">▦ Show QR</button>{% endif %}\n<input class="share-mobile" id="mobile-{{b.id}}" placeholder="Client WhatsApp No.">\n<button class="btn" type="button" onclick="shareBank({{b.id}})">WhatsApp Payment Details</button>\n</div>\n<div class="qr-panel" id="qr-{{b.id}}"><h3>{{b.account_name}}</h3><div id="qrcode-{{b.id}}"></div><p><b>{{b.upi_id}}</b></p><small>Scan only after verifying company/account name.</small></div>\n</div>{% else %}<div class="card">No approved payment account is currently available for your company.</div>{% endfor %}\n</div>\n<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>\n<script>\nfunction toggleQR(id,text){\n const p=document.getElementById('qr-'+id);p.classList.toggle('open');\n const box=document.getElementById('qrcode-'+id);\n if(p.classList.contains('open') && box && !box.dataset.done){\n   if(window.QRCode){new QRCode(box,{text:text,width:190,height:190});box.dataset.done='1'}\n   else{box.innerHTML='<b>QR service unavailable. Please share bank details manually.</b>'}\n }\n}\nasync function shareBank(id){\n const mobile=document.getElementById('mobile-'+id).value.trim();\n const fd=new FormData();fd.append('mobile',mobile);\n const r=await fetch('/payment-accounts/share/'+id,{method:'POST',body:fd});\n const d=await r.json();\n if(d.ok){window.open(d.whatsapp_url,'_blank')}else{alert(d.error||'Unable to share')}\n}\n</script>\n{% endblock %}""", 'bank_share_history.html': """{% extends "base.html" %}{% block content %}\n<h1 class="title">📤 Bank Details Share History</h1>\n<div class="card"><p>Audit trail of approved company bank details shared through the portal.</p>\n<div class="tablewrap"><table><thead><tr><th>Date / Time</th><th>Company</th><th>Bank</th><th>Account</th><th>Shared By</th><th>Recipient</th><th>Mode</th></tr></thead><tbody>\n{% for r in rows %}<tr><td>{{r.shared_at}}</td><td>{{r.company_code}}</td><td>{{r.bank_name}}</td><td>{{r.account_name}} ••••{{r.account_last4}}</td><td>{{r.shared_by_name}}</td><td>{{r.recipient_mobile}}</td><td>{{r.share_type}}</td></tr>\n{% else %}<tr><td colspan="7">No share activity yet.</td></tr>{% endfor %}</tbody></table></div></div>{% endblock %}""", 'competition.html': """{% extends "base.html" %}{% block content %}\n<style>\n.comp-mgmt-bar{background:rgba(2,17,32,.95);border:1px solid rgba(230,183,63,.5);border-radius:16px;padding:20px 24px;margin-bottom:22px}\n.comp-status-badge{display:inline-flex;align-items:center;gap:7px;padding:6px 16px;border-radius:20px;font-size:13px;font-weight:bold;letter-spacing:.5px}\n.comp-status-badge.active{background:rgba(34,197,94,.15);border:1px solid #22c55e;color:#22c55e}\n.comp-status-badge.finished{background:rgba(107,114,128,.15);border:1px solid #6b7280;color:#9ca3af}\n.comp-status-badge.not_started{background:rgba(230,183,63,.12);border:1px solid #e6b73f;color:#e6b73f}\n.comp-actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px}\n.btn-start{background:linear-gradient(135deg,#22c55e,#16a34a);color:#fff;border:0;padding:10px 22px;border-radius:10px;font-size:15px;font-weight:bold;cursor:pointer;display:flex;align-items:center;gap:7px}\n.btn-finish{background:linear-gradient(135deg,#ef4444,#b91c1c);color:#fff;border:0;padding:10px 22px;border-radius:10px;font-size:15px;font-weight:bold;cursor:pointer;display:flex;align-items:center;gap:7px}\n.btn-start:hover{filter:brightness(1.1)}.btn-finish:hover{filter:brightness(1.1)}\n.comp-details-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-top:14px}\n.comp-detail-item{background:rgba(1,14,28,.7);border:1px solid #1e3a5f;border-radius:10px;padding:12px 14px}\n.comp-detail-item small{color:#7f9caf;font-size:12px;display:block;margin-bottom:3px}\n.comp-detail-item b{color:#e6b73f;font-size:16px}\n/* Modal */\n.comp-modal-overlay{position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,.75);display:flex;align-items:center;justify-content:center;padding:16px;opacity:0;visibility:hidden;transition:.2s}\n.comp-modal-overlay.open{opacity:1;visibility:visible}\n.comp-modal{background:#071626;border:1px solid rgba(230,183,63,.5);border-radius:20px;padding:28px 30px;width:min(520px,96vw);max-height:92vh;overflow-y:auto;box-shadow:0 30px 80px rgba(0,0,0,.7)}\n.comp-modal h2{color:#e6b73f;margin:0 0 18px;font-size:22px;display:flex;align-items:center;gap:10px}\n.comp-modal label{display:block;color:#e6b73f;font-size:13px;font-weight:bold;margin:14px 0 5px;letter-spacing:.3px}\n.comp-modal input,.comp-modal select,.comp-modal textarea{width:100%;padding:11px 13px;border:1px solid #2a4d6e;border-radius:9px;background:rgba(1,14,28,.9);color:#fff;font-size:15px;outline:none;font-family:inherit}\n.comp-modal input:focus,.comp-modal select:focus{border-color:#e6b73f}\n.comp-modal-actions{display:flex;gap:10px;margin-top:22px;justify-content:flex-end}\n.comp-modal-cancel{background:transparent;border:1px solid #2a4d6e;color:#9ca3af;padding:10px 20px;border-radius:9px;cursor:pointer;font-size:14px}\n.comp-modal-submit{background:linear-gradient(135deg,#22c55e,#16a34a);color:#fff;border:0;padding:10px 24px;border-radius:9px;font-size:15px;font-weight:bold;cursor:pointer}\n.comp-modal-submit:hover{filter:brightness(1.1)}\n.comp-row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}\n@media(max-width:600px){.comp-row2{grid-template-columns:1fr}.comp-actions{flex-direction:column}.comp-modal{padding:20px 16px}}\n.comp-history-table{width:100%;border-collapse:collapse;font-size:13px;margin-top:10px}\n.comp-history-table th{color:#e6b73f;text-align:left;padding:8px 10px;border-bottom:1px solid #1e3a5f;font-weight:bold}\n.comp-history-table td{padding:8px 10px;border-bottom:1px solid #0d2035;color:#c8dce9}\n.comp-history-table tr:hover td{background:rgba(30,58,95,.3)}\n</style>\n\n<h1 class="title">🔥 Competition Management</h1>\n\n{% if u.role == 'MD' %}\n<div class="comp-mgmt-bar">\n  <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">\n    <div>\n      <div style="color:#7f9caf;font-size:13px;margin-bottom:6px">COMPETITION STATUS</div>\n      {% set status = settings.get('status','not_started') if settings.get('status') else ('active' if settings.enabled else 'not_started') %}\n      <span class="comp-status-badge {{status}}">\n        {% if status == 'active' %}🟢 ACTIVE{% elif status == 'finished' %}⛔ FINISHED{% else %}⏳ NOT STARTED{% endif %}\n      </span>\n    </div>\n    <div class="comp-actions">\n      {% if status != 'active' %}\n      <button class="btn-start" onclick="document.getElementById('startTestModal').classList.add('open')">\n        🚀 Start Test\n      </button>\n      {% endif %}\n      {% if status == 'active' %}\n      <button class="btn-finish" onclick="document.getElementById('finishConfirmModal').classList.add('open')">\n        🏁 Finish Test\n      </button>\n      <button class="btn-start" onclick="document.getElementById('startTestModal').classList.add('open')" style="background:linear-gradient(135deg,#e6b73f,#b8860b);color:#091423">\n        ✏️ Edit Competition\n      </button>\n      {% endif %}\n    </div>\n  </div>\n\n  {% if status in ('active','finished') %}\n  <div class="comp-details-grid" style="margin-top:16px">\n    <div class="comp-detail-item">\n      <small>Competition Name</small>\n      <b>{{settings.title or '—'}}</b>\n    </div>\n    <div class="comp-detail-item">\n      <small>Month / Year</small>\n      <b>{{settings.competition_month or (settings.start_date[:7] if settings.start_date else '—')}}</b>\n    </div>\n    <div class="comp-detail-item">\n      <small>Target Enrollment</small>\n      <b>{{settings.scic_target or 0}} each</b>\n    </div>\n    <div class="comp-detail-item">\n      <small>Budget</small>\n      <b>{% if settings.budget %}₹{{'{:,.0f}'.format(settings.budget|float)}}{% else %}—{% endif %}</b>\n    </div>\n    <div class="comp-detail-item">\n      <small>Prize</small>\n      <b>{{settings.prize or '—'}}</b>\n    </div>\n    <div class="comp-detail-item">\n      <small>Period</small>\n      <b>{{settings.start_date}} → {{settings.end_date}}</b>\n    </div>\n  </div>\n  {% endif %}\n</div>\n{% endif %}\n\n<div class="competition-grid">\n<div class="comp-team-card">\n  <h2>🏢 ABC</h2>\n  <div class="comp-score-big">{{scores.ABC.enrollments if scores.ABC is defined else 0}}</div>\n  <p>Enrollments</p>\n  {% if u.role == 'MD' %}<p style="color:#e6b73f">₹{{'{:,.0f}'.format(scores.ABC.revenue|float) if scores.ABC is defined else 0}}</p>{% endif %}\n  {% if settings.scic_target %}<p style="color:#7f9caf;font-size:13px">Target: {{settings.scic_target}}</p>{% endif %}\n</div>\n<div class="comp-team-card">\n  <h2>🌊 XYZ</h2>\n  <div class="comp-score-big">{{scores.XYZ.enrollments if scores.XYZ is defined else 0}}</div>\n  <p>Enrollments</p>\n  {% if u.role == 'MD' %}<p style="color:#e6b73f">₹{{'{:,.0f}'.format(scores.XYZ.revenue|float) if scores.XYZ is defined else 0}}</p>{% endif %}\n  {% if settings.wwic_target %}<p style="color:#7f9caf;font-size:13px">Target: {{settings.wwic_target}}</p>{% endif %}\n</div>\n</div>\n\n{% if u.role == 'MD' %}\n<div class="card" style="margin-top:20px">\n  <h2 style="color:#e6b73f;margin-bottom:14px">⚙️ Advanced Settings</h2>\n  <form method="post" action="{{url_for('competition_settings_save')}}">\n    <div class="comp-row2">\n      <div><label>Start Date</label><input type="date" name="start_date" value="{{settings.start_date}}"></div>\n      <div><label>End Date</label><input type="date" name="end_date" value="{{settings.end_date}}"></div>\n    </div>\n    <div class="comp-row2" style="margin-top:10px">\n      <div><label>ABC Target Enrollments</label><input type="number" name="scic_target" value="{{settings.scic_target}}"></div>\n      <div><label>XYZ Target Enrollments</label><input type="number" name="wwic_target" value="{{settings.wwic_target}}"></div>\n    </div>\n    <div class="comp-row2" style="margin-top:10px">\n      <div><label>Competition Status</label>\n        <select name="enabled">\n          <option value="1" {% if settings.enabled %}selected{% endif %}>ON</option>\n          <option value="0" {% if not settings.enabled %}selected{% endif %}>OFF</option>\n        </select>\n      </div>\n      <div><label>Popups</label>\n        <select name="popup_enabled">\n          <option value="1" {% if settings.popup_enabled %}selected{% endif %}>ON</option>\n          <option value="0" {% if not settings.popup_enabled %}selected{% endif %}>OFF</option>\n        </select>\n      </div>\n    </div>\n    <button class="btn btn-gold" style="margin-top:14px" type="submit">💾 Save Settings</button>\n  </form>\n</div>\n{% endif %}\n\n<div class="card" style="margin-top:20px">\n  <h2 style="color:#e6b73f">⚡ Live Competition Feed</h2>\n  {% if events %}\n  <table class="comp-history-table">\n    <thead><tr><th>Time</th><th>Company</th><th>Employee</th><th>Message</th></tr></thead>\n    <tbody>\n    {% for e in events %}\n    <tr>\n      <td>{{e.created_at[:16] if e.created_at else '—'}}</td>\n      <td><span style="color:{{'#e6b73f' if e.company_code=='SCIC' else '#55aaff'}}">{{e.company_code}}</span></td>\n      <td>{{e.employee_name or '—'}}</td>\n      <td>{{e.message or '—'}}</td>\n    </tr>\n    {% endfor %}\n    </tbody>\n  </table>\n  {% else %}\n  <p style="color:#7f9caf;text-align:center;padding:20px">No competition events yet.</p>\n  {% endif %}\n</div>\n\n{% if u.role == 'MD' %}\n<!-- Start Test Modal -->\n<div class="comp-modal-overlay" id="startTestModal" onclick="if(event.target===this)this.classList.remove('open')">\n  <div class="comp-modal">\n    <h2>🚀 Start Competition</h2>\n    <form method="post" action="{{url_for('competition_start')}}" onsubmit="return validateCompForm()">\n      <div class="comp-row2">\n        <div>\n          <label>Month & Year *</label>\n          <input type="date" name="start_date" id="compStartDate" required value="{{settings.start_date or ''}}">\n        </div>\n        <div>\n          <label>End Date *</label>\n          <input type="date" name="end_date" id="compEndDate" required value="{{settings.end_date or ''}}">\n        </div>\n        <div>\n          <label>Target Enrollment (Each) *</label>\n          <input type="number" name="target_enrollment" id="compTarget" min="1" required placeholder="e.g. 50" value="{{settings.scic_target or ''}}">\n        </div>\n      </div>\n      <div>\n        <label>Budget (₹) *</label>\n        <input type="number" name="budget" id="compBudget" step="0.01" min="0" required placeholder="e.g. 50000" value="{{settings.budget or ''}}">\n      </div>\n      <div>\n        <label>Name of Competition *</label>\n        <input type="text" name="comp_name" id="compName" required placeholder="e.g. Grand Q1 Competition" value="{{settings.title or ''}}">\n      </div>\n      <div>\n        <label>Prize of Competition *</label>\n        <input type="text" name="prize" id="compPrize" required placeholder="e.g. iPhone 15, Trophy + ₹10,000 cash" value="{{settings.prize or ''}}">\n      </div>\n      <div class="comp-modal-actions">\n        <button type="button" class="comp-modal-cancel" onclick="document.getElementById('startTestModal').classList.remove('open')">Cancel</button>\n        <button type="submit" class="comp-modal-submit">🚀 Start Competition</button>\n      </div>\n    </form>\n  </div>\n</div>\n\n<!-- Finish Confirmation Modal -->\n<div class="comp-modal-overlay" id="finishConfirmModal" onclick="if(event.target===this)this.classList.remove('open')">\n  <div class="comp-modal" style="text-align:center;max-width:400px">\n    <h2 style="justify-content:center">🏁 Finish Competition?</h2>\n    <p style="color:#c8dce9;margin-bottom:20px">Are you sure you want to finish this competition?<br><span style="color:#ef4444;font-size:13px">This will mark it as CLOSED and hide it from normal portal users.</span></p>\n    <div class="comp-modal-actions" style="justify-content:center">\n      <button type="button" class="comp-modal-cancel" onclick="document.getElementById('finishConfirmModal').classList.remove('open')">Cancel</button>\n      <form method="post" action="{{url_for('competition_finish')}}" style="display:inline">\n        <button type="submit" class="btn-finish">🏁 Yes, Finish It</button>\n      </form>\n    </div>\n  </div>\n</div>\n\n<script>\nfunction validateCompForm(){\n  const m=document.getElementById('compMonth').value;\n  const t=document.getElementById('compTarget').value;\n  const b=document.getElementById('compBudget').value;\n  const n=document.getElementById('compName').value.trim();\n  const p=document.getElementById('compPrize').value.trim();\n  if(!m||!t||!b||!n||!p){alert('Please fill all required fields.');return false;}\n  if(parseInt(t)<1){alert('Target enrollment must be at least 1.');return false;}\n  if(parseFloat(b)<0){alert('Budget cannot be negative.');return false;}\n  return true;\n}\n// Close modal on Escape\ndocument.addEventListener('keydown',function(e){\n  if(e.key==='Escape'){\n    document.getElementById('startTestModal').classList.remove('open');\n    document.getElementById('finishConfirmModal').classList.remove('open');\n  }\n});\n</script>\n{% endif %}\n{% endblock %}"""}
_GAUR_V337_TEMPLATES['finance_daily_report.html']="""{% extends "base.html" %}{% block content %}
<style>
.dr-card{background:#071d32;border:1px solid #315a7b;border-radius:14px;padding:16px;margin-bottom:16px}
.dr-table{width:100%;border-collapse:collapse;font-size:13px}
.dr-table th{background:#0a2540;color:#e6b73f;padding:8px 10px;text-align:left;border:1px solid #1e4060}
.dr-table td{padding:7px 10px;border:1px solid #1e3050;vertical-align:top}
.dr-table tr.in-row td{background:rgba(72,213,139,.07)}
.dr-table tr.out-row td{background:rgba(255,80,80,.07)}
.dr-table tr.total-row td{background:#0a2540;font-weight:bold;color:#e6b73f}
.dr-table tr.cf-row td{background:#0c1e30;font-weight:bold;color:#8fc8ff}
.dr-stat{background:linear-gradient(145deg,#0c2847,#071b30);border:1px solid #315a7b;border-radius:12px;padding:14px;text-align:center}
.dr-stat b{display:block;font-size:24px;color:#e6b73f;margin-top:4px}
.in-val{color:#48d58b;font-weight:bold}
.out-val{color:#ff7a7a;font-weight:bold}
</style>

<div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;margin-bottom:16px">
  <h1 class="title" style="margin:0">📊 Daily Report — {{ sel_company }}</h1>
  <a href="{{url_for('finance_center')}}" class="toolbtn">← Daily Passbook</a>
</div>

<div class="dr-card">
  <form method="get" style="display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end">
    <div><label style="font-size:13px">Date</label><br>
      <input type="date" name="date" value="{{sel_date}}" style="padding:8px;background:#071d32;border:1px solid #427398;color:#fff;border-radius:8px"></div>
    {% if u.role=='MD' %}
    <div><label style="font-size:13px">Company</label><br>
      <select name="company" style="padding:8px;background:#071d32;border:1px solid #427398;color:#fff;border-radius:8px">
        <option value="SCIC" {{'selected' if sel_company=='SCIC' else ''}}>ABC</option>
        <option value="WWIC" {{'selected' if sel_company=='WWIC' else ''}}>XYZ</option>
      </select></div>
    {% endif %}
    <button class="btn" style="padding:8px 16px">🔍 View</button>
  </form>
</div>

<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin-bottom:16px">
  <div class="dr-stat"><span>Total In Flow</span><b class="in-val">₹{{"{:,.0f}".format(total_in)}}</b></div>
  <div class="dr-stat"><span>Total Out Flow</span><b class="out-val">₹{{"{:,.0f}".format(total_out)}}</b></div>
  <div class="dr-stat"><span>Net</span><b style="color:{{'#48d58b' if total_in>=total_out else '#ff7a7a'}}">₹{{"{:,.0f}".format(total_in-total_out)}}</b></div>
  <div class="dr-stat"><span>Entries</span><b>{{rows|length}}</b></div>
</div>

<div class="dr-card" id="add-entry">
  <h2 style="color:#e6b73f;margin:0 0 14px">➕ Add Entry for {{ sel_date }}</h2>
  <form method="post">
    {% if u.role=='MD' %}<input type="hidden" name="company_code" value="{{sel_company}}">{% endif %}
    <div class="grid3">
      <div><label>Date *</label><input type="date" name="txn_date" value="{{sel_date}}" required></div>
      <div><label>Direction *</label>
        <select name="direction">
          <option value="IN">📥 In Flow (Receipt)</option>
          <option value="OUT">📤 Out Flow (Expense)</option>
        </select>
      </div>
      <div><label>Amount (₹) *</label><input type="number" name="amount" step="0.01" min="0.01" required></div>
      <div><label>Party / Client Name *</label><input name="party_name" required placeholder="Client ya vendor naam"></div>
      <div><label>Bank Account</label>
        <select name="account_id">
          <option value="">— Cash —</option>
          {% for b in banks %}<option value="{{b.id}}">{{b.bank_name}} (...{{b.account_last4}})</option>{% endfor %}
        </select>
      </div>
      <div><label>Category</label>
        <select name="head">
          <option>Enrollment Fee</option><option>Second Payment</option><option>Other Receipt</option>
          <option>Office Salaries</option><option>Office Rent</option><option>Miscellaneous Expense</option>
          <option>Travel</option><option>Marketing</option><option>Bank Charges</option><option>Other</option>
        </select>
      </div>
      <div><label>Payment Mode</label>
        <select name="payment_mode">
          <option>Cash</option><option>NEFT</option><option>IMPS</option><option>UPI</option>
          <option>Cheque</option><option>RTGS</option>
        </select>
      </div>
      <div><label>Reference No.</label><input name="reference_no" placeholder="UTR / Cheque No."></div>
      <div><label>Remarks</label><input name="remarks" placeholder="Optional note"></div>
    </div>
    <br><button class="btn">💾 Save Entry</button>
  </form>
</div>

<div class="dr-card">
  <h2 style="color:#e6b73f;margin:0 0 14px">{{ sel_company }} Payment Inflow & Outflow — {{ sel_date }}</h2>
  <div class="tablewrap">
  <table class="dr-table">
    <thead>
      <tr>
        <th>Sr.</th><th>Particulars</th><th>Cash</th>
        {% for b in banks %}<th>{{b.bank_name}}<br><small>...{{b.account_last4}}</small></th>{% endfor %}
      </tr>
    </thead>
    <tbody>
      <tr class="cf-row">
        <td>—</td><td>Balance C/F</td>
        <td class="in-val">₹{{"{:,.0f}".format(balance_cf.get(none_key,0))}}</td>
        {% for b in banks %}<td class="in-val">₹{{"{:,.0f}".format(balance_cf.get(b.id,0))}}</td>{% endfor %}
      </tr>
      {% for i,r in enumerate(inflows) %}
      <tr class="in-row">
        <td>{{i+1}}</td>
        <td><b>{{r.party_name}}</b>{% if r.remarks %}<br><small style="opacity:.6">{{r.remarks}}</small>{% endif %}</td>
        <td>{% if not r.account_id %}<span class="in-val">₹{{"{:,.0f}".format(r.amount)}}</span>{% endif %}</td>
        {% for b in banks %}<td>{% if r.account_id==b.id %}<span class="in-val">₹{{"{:,.0f}".format(r.amount)}}</span>{% endif %}</td>{% endfor %}
      </tr>
      {% endfor %}
      <tr class="total-row">
        <td colspan="2">📥 In Flow Total</td>
        <td class="in-val">₹{{"{:,.0f}".format(bank_in.get(none_key,0))}}</td>
        {% for b in banks %}<td class="in-val">₹{{"{:,.0f}".format(bank_in.get(b.id,0))}}</td>{% endfor %}
      </tr>
      <tr class="total-row" style="color:#8fc8ff">
        <td colspan="2">Running Balance</td>
        <td>₹{{"{:,.0f}".format(balance_cf.get(none_key,0)+bank_in.get(none_key,0))}}</td>
        {% for b in banks %}<td>₹{{"{:,.0f}".format(balance_cf.get(b.id,0)+bank_in.get(b.id,0))}}</td>{% endfor %}
      </tr>
      {% for i,r in enumerate(outflows) %}
      <tr class="out-row">
        <td>{{i+1}}</td>
        <td><b>{{r.party_name}}</b>{% if r.head %}<br><small style="opacity:.6">{{r.head}}</small>{% endif %}</td>
        <td>{% if not r.account_id %}<span class="out-val">₹{{"{:,.0f}".format(r.amount)}}</span>{% endif %}</td>
        {% for b in banks %}<td>{% if r.account_id==b.id %}<span class="out-val">₹{{"{:,.0f}".format(r.amount)}}</span>{% endif %}</td>{% endfor %}
      </tr>
      {% endfor %}
      <tr class="total-row">
        <td colspan="2">📤 Out Flow Total</td>
        <td class="out-val">₹{{"{:,.0f}".format(bank_out.get(none_key,0))}}</td>
        {% for b in banks %}<td class="out-val">₹{{"{:,.0f}".format(bank_out.get(b.id,0))}}</td>{% endfor %}
      </tr>
      <tr class="total-row" style="background:#071828;border-top:2px solid #e6b73f">
        <td colspan="2">💰 Grand Total (Closing Balance)</td>
        <td>₹{{"{:,.0f}".format(balance_cf.get(none_key,0)+bank_in.get(none_key,0)-bank_out.get(none_key,0))}}</td>
        {% for b in banks %}<td>₹{{"{:,.0f}".format(balance_cf.get(b.id,0)+bank_in.get(b.id,0)-bank_out.get(b.id,0))}}</td>{% endfor %}
      </tr>
    </tbody>
  </table>
  </div>
</div>
{% endblock %}"""
app.jinja_loader = ChoiceLoader([DictLoader(_GAUR_V337_TEMPLATES), app.jinja_loader])

# === GAUR CRM v3.38 ENROLLMENT CELEBRATION BLAST ===

_GAUR_V337_TEMPLATES['cases.html'] = r"""{% extends "base.html" %}{% block content %}
<h1 class="title">Enrollment & Payment Report</h1>
<p style="opacity:.78">Structured payment ledger • Package, After Visa, First Payment, Second Payment and Other Payment with date/time audit.</p>

<div class="card">
<h2 style="color:#e6b73f;margin-top:0">Create New Enrollment</h2>
<form method="post">
<div class="grid3">
{% if u['role']=='MD' %}<div><label>Company</label><select name="company_code" required><option value="SCIC">ABC</option><option value="WWIC">XYZ</option></select></div>{% endif %}
<div><label>Client Name</label><input name="client_name" required></div>
<div><label>Mobile</label><input name="mobile"></div>
<div><label>Country</label><input name="country"></div>
<div><label>Visa Type</label><input name="visa_type"></div>
<div><label>Enrollment Date</label><input type="date" name="enrollment_date"></div>
<div><label>Assigned Employee</label><select name="assigned_employee_id"><option value="">Select</option>{% for e in staff %}<option value="{{e.id}}">{{e.full_name}} • {{e.company_code}}</option>{% endfor %}</select></div>
</div>

<h2 style="color:#e6b73f;margin:22px 0 8px">Enrollment & Payment Details</h2>
<div class="grid3">
<div><label>1. Package Amount</label><input type="number" step="0.01" min="0" name="package_amount" value="0"><label>Package Entry Date & Time</label><input type="datetime-local" name="package_amount_datetime"></div>
<div><label>2. After Visa Payment</label><input type="number" step="0.01" min="0" name="after_visa_payment" value="0"><label>After Visa Entry Date & Time</label><input type="datetime-local" name="after_visa_payment_datetime"></div>
<div><label>3. First Payment</label><input type="number" step="0.01" min="0" name="booking_amount" value="0"><label>First Payment Date & Time</label><input type="datetime-local" name="first_payment_datetime"></div>
<div><label>4. First Payment Status</label><select name="first_payment_status"><option>Received</option><option>Partially Received</option><option>Pending</option><option>Cancelled</option></select></div>
<div><label>5. 2nd Payment</label><input type="number" step="0.01" min="0" name="second_payment" value="0"><label>2nd Payment Date & Time</label><input type="datetime-local" name="second_payment_datetime"></div>
<div><label>6. 2nd Payment Status</label><select name="second_payment_status"><option>Pending</option><option>Received</option><option>Partially Received</option><option>Cancelled</option></select></div>
<div><label>7. Other Payment Received</label><input type="number" step="0.01" min="0" name="other_payment_received" value="0"><label>Other Payment Date & Time</label><input type="datetime-local" name="other_payment_datetime"></div>
<div><label>Filing Status</label><select name="filing_status"><option>Documents Pending</option><option>Documents Complete</option><option>Filed</option><option>Visa Granted</option><option>Visa Refused</option></select></div>
<div><label>Remarks</label><textarea name="remarks"></textarea></div>
</div>
<div style="margin-top:16px"><button class="btn">Save Enrollment & Payment Details</button></div>
</form>
</div>

<div class="card">
<h2 style="color:#e6b73f;margin-top:0">Enrollment & Payment Records</h2>
<div class="tablewrap"><table>
<thead><tr><th>Case / Client</th><th>Package</th><th>After Visa</th><th>First Payment</th><th>1st Status</th><th>1st Date/Time</th><th>2nd Payment</th><th>2nd Status</th><th>2nd Date/Time</th><th>Other Received</th><th>Other Date/Time</th><th>Actual Collection</th><th>Update</th></tr></thead>
<tbody>
{% for r in rows %}
<tr>
<td><b>{{r.case_id}}</b><br>{{r.client_name}}<br><small>{{r.employee_name or '-'}}</small></td>
<td>₹{{"{:,.0f}".format(r.package_amount or 0)}}<br><small>{{r.package_amount_datetime or '-'}}</small></td>
<td>₹{{"{:,.0f}".format(r.after_visa_payment or 0)}}<br><small>{{r.after_visa_payment_datetime or '-'}}</small></td>
<form method="post" action="{{url_for('update_case',case_id=r.id)}}">
<td><input type="number" step="0.01" min="0" name="booking_amount" value="{{r.booking_amount or 0}}"></td>
<td><select name="first_payment_status">{% for s in ['Received','Partially Received','Pending','Cancelled'] %}<option {% if (r.first_payment_status or 'Received')==s %}selected{% endif %}>{{s}}</option>{% endfor %}</select></td>
<td><input type="datetime-local" name="first_payment_datetime" value="{{(r.first_payment_datetime or '')[:16]}}"></td>
<td><input type="number" step="0.01" min="0" name="second_payment" value="{{r.second_payment or 0}}"></td>
<td><select name="second_payment_status">{% for s in ['Pending','Received','Partially Received','Cancelled'] %}<option {% if (r.second_payment_status or r.payment_status or 'Pending')==s %}selected{% endif %}>{{s}}</option>{% endfor %}</select></td>
<td><input type="datetime-local" name="second_payment_datetime" value="{{(r.second_payment_datetime or '')[:16]}}"></td>
<td><input type="number" step="0.01" min="0" name="other_payment_received" value="{{r.other_payment_received or 0}}"></td>
<td><input type="datetime-local" name="other_payment_datetime" value="{{(r.other_payment_datetime or '')[:16]}}"></td>
<td><b style="color:#48d58b">₹{{"{:,.0f}".format(r.total_received or 0)}}</b></td>
<td>
<input type="hidden" name="package_amount" value="{{r.package_amount or 0}}">
<input type="hidden" name="package_amount_datetime" value="{{(r.package_amount_datetime or '')[:16]}}">
<input type="hidden" name="after_visa_payment" value="{{r.after_visa_payment or 0}}">
<input type="hidden" name="after_visa_payment_datetime" value="{{(r.after_visa_payment_datetime or '')[:16]}}">
<input type="hidden" name="filing_status" value="{{r.filing_status or 'Documents Pending'}}">
<input type="hidden" name="remarks" value="{{r.remarks or ''}}">
<button class="toolbtn">Save</button>
</td>
</form>
</tr>
{% else %}<tr><td colspan="13">No enrollment records yet.</td></tr>{% endfor %}
</tbody></table></div>
</div>

<div class="card" style="border-color:#e6b73f">
<b>Calculation Rule:</b> Package Amount and After Visa Payment are agreement/due values only. Revenue/Collection includes only First Payment and 2nd Payment whose status is Received/Partially Received, plus Other Payment Received.
</div>
{% endblock %}"""

_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('.competition-popup{position:fixed;right:22px;top:84px;width:min(420px,calc(100vw - 30px));z-index:999999;background:linear-gradient(145deg,#091f37,#152e4a);border:2px solid #e6b73f;border-radius:20px;padding:20px;box-shadow:0 18px 60px rgba(0,0,0,.55);transform:translateX(130%);opacity:0;transition:.45s ease}\n.competition-popup.show{transform:translateX(0);opacity:1}.competition-popup.wwic{border-color:#4da6ff}.comp-pop-title{font-size:26px;color:#e6b73f;font-weight:900}.comp-fire{font-size:40px;animation:compPulse .75s infinite alternate}@keyframes compPulse{from{transform:scale(1)}to{transform:scale(1.15)}}\n','.competition-popup{position:fixed;inset:0;z-index:9999999;display:flex;align-items:center;justify-content:center;padding:18px;background:radial-gradient(circle at 50% 42%,rgba(36,74,119,.55),rgba(0,7,18,.92) 58%,rgba(0,0,0,.98));opacity:0;visibility:hidden;pointer-events:none;transition:opacity .28s ease,visibility .28s ease;overflow:hidden}\n.competition-popup.show{opacity:1;visibility:visible;pointer-events:auto}.competition-popup:before{content:"";position:absolute;inset:-30%;background:conic-gradient(from 0deg,transparent,rgba(255,214,64,.10),transparent,rgba(76,169,255,.12),transparent);animation:celebrateSpin 9s linear infinite}.celebrate-stage{position:relative;width:min(940px,96vw);min-height:min(690px,92vh);border:2px solid #f5c84b;border-radius:34px;background:linear-gradient(155deg,#071426 0%,#102f51 45%,#071426 100%);box-shadow:0 0 0 7px rgba(245,200,75,.08),0 30px 100px rgba(0,0,0,.72),0 0 60px rgba(245,200,75,.22) inset;padding:26px 30px 24px;text-align:center;overflow:hidden;transform:scale(.72) translateY(35px);transition:transform .48s cubic-bezier(.2,.9,.25,1.18)}\n.competition-popup.show .celebrate-stage{transform:scale(1) translateY(0)}.competition-popup.wwic .celebrate-stage{border-color:#55aaff;box-shadow:0 0 0 7px rgba(85,170,255,.08),0 30px 100px rgba(0,0,0,.72),0 0 60px rgba(85,170,255,.22) inset}.celebrate-glow{position:absolute;left:50%;top:43%;width:520px;height:520px;transform:translate(-50%,-50%);background:radial-gradient(circle,rgba(255,210,64,.30),rgba(255,210,64,.04) 52%,transparent 72%);filter:blur(2px);animation:celebrateGlow 1s ease-in-out infinite alternate}.competition-popup.wwic .celebrate-glow{background:radial-gradient(circle,rgba(85,170,255,.32),rgba(85,170,255,.05) 52%,transparent 72%)}.celebrate-crown{position:relative;z-index:2;font-size:72px;line-height:1;filter:drop-shadow(0 8px 14px rgba(0,0,0,.45));animation:crownBounce .75s ease-in-out infinite alternate}.comp-pop-title{position:relative;z-index:2;margin:6px 0 0;font-family:Arial Black,Arial,sans-serif;font-size:clamp(34px,6vw,76px);line-height:.98;color:#ffd85e;text-shadow:0 3px 0 #8c3517,0 7px 18px rgba(0,0,0,.7);letter-spacing:1px}.competition-popup.wwic .comp-pop-title{color:#d9efff;text-shadow:0 3px 0 #145792,0 7px 18px rgba(0,0,0,.7)}.celebrate-subtitle{position:relative;z-index:2;display:inline-block;margin:10px auto 17px;padding:8px 28px;border-radius:999px;background:linear-gradient(180deg,#1b66c7,#0d3f87);border:2px solid rgba(255,255,255,.34);font-family:Arial Black,Arial,sans-serif;font-size:clamp(21px,3.4vw,39px);letter-spacing:1px;box-shadow:0 8px 20px rgba(0,0,0,.35)}.celebrate-person{position:relative;z-index:2;display:flex;justify-content:center;align-items:center;gap:18px;flex-wrap:wrap;margin:2px auto 14px}.celebrate-avatar-wrap{width:170px;height:170px;border-radius:50%;padding:5px;background:linear-gradient(145deg,#fff7b0,#d99e14);box-shadow:0 0 0 8px rgba(255,210,75,.10),0 0 40px rgba(255,210,75,.38);animation:avatarPop .72s cubic-bezier(.2,.9,.3,1.3)}.competition-popup.wwic .celebrate-avatar-wrap{background:linear-gradient(145deg,#e5f6ff,#4096e8);box-shadow:0 0 0 8px rgba(85,170,255,.10),0 0 40px rgba(85,170,255,.38)}.celebrate-avatar{width:100%;height:100%;object-fit:cover;border-radius:50%;background:#09192a}.celebrate-avatar-fallback{width:100%;height:100%;display:flex;align-items:center;justify-content:center;border-radius:50%;background:#102b49;font-size:70px}.celebrate-person-info{text-align:left;min-width:260px}.celebrate-name{font-family:Arial Black,Arial,sans-serif;font-size:clamp(29px,4vw,48px);color:#fff}.celebrate-role{font-size:20px;color:#ffd85e;font-weight:800;letter-spacing:.6px}.competition-popup.wwic .celebrate-role{color:#7fc4ff}.celebrate-company{margin-top:6px;font-size:16px;opacity:.86}.celebrate-details{position:relative;z-index:2;display:grid;grid-template-columns:repeat(3,1fr);gap:10px;width:min(760px,100%);margin:10px auto}.celebrate-detail{padding:11px 12px;background:rgba(0,0,0,.27);border:1px solid rgba(255,255,255,.16);border-radius:13px}.celebrate-detail small{display:block;opacity:.67;margin-bottom:3px}.celebrate-detail b{display:block;font-size:19px;color:#ffe27b;word-break:break-word}.competition-popup.wwic .celebrate-detail b{color:#9fd1ff}.celebrate-save{position:relative;z-index:2;display:inline-flex;align-items:center;gap:8px;margin-top:7px;padding:10px 21px;border-radius:999px;background:#138b42;border:2px solid #5bea8b;font-size:20px;font-weight:900;box-shadow:0 8px 20px rgba(0,0,0,.3)}.celebrate-close{position:absolute;right:16px;top:12px;z-index:6;width:40px;height:40px;border-radius:50%;border:1px solid rgba(255,255,255,.35);background:rgba(0,0,0,.35);color:#fff;font-size:24px;cursor:pointer}.celebrate-countdown{position:relative;z-index:2;margin-top:8px;font-size:13px;opacity:.7}.confetti-piece{position:absolute;z-index:4;top:-12vh;width:10px;height:20px;border-radius:2px;animation:confettiFall linear forwards}.firework{position:absolute;z-index:1;width:8px;height:8px;border-radius:50%;box-shadow:0 -70px #ffd75e,49px -49px #ff5964,70px 0 #5bc0ff,49px 49px #66e49a,0 70px #ff8ce8,-49px 49px #ffb14e,-70px 0 #78a7ff,-49px -49px #d7ff61;animation:fireworkBurst 1.45s ease-out infinite}.fw1{left:13%;top:20%}.fw2{right:13%;top:18%;animation-delay:.4s}.fw3{left:20%;bottom:21%;animation-delay:.8s}.fw4{right:19%;bottom:20%;animation-delay:1.1s}.celebrate-side{position:absolute;z-index:2;font-size:54px;filter:drop-shadow(0 6px 10px rgba(0,0,0,.5));animation:sideFloat 1s ease-in-out infinite alternate}.side-left{left:4%;top:48%}.side-right{right:4%;top:48%;animation-delay:.35s}@keyframes celebrateSpin{to{transform:rotate(360deg)}}@keyframes celebrateGlow{from{transform:translate(-50%,-50%) scale(.88);opacity:.55}to{transform:translate(-50%,-50%) scale(1.08);opacity:1}}@keyframes crownBounce{from{transform:translateY(0) rotate(-3deg)}to{transform:translateY(-9px) rotate(3deg)}}@keyframes avatarPop{0%{transform:scale(.2) rotate(-16deg)}75%{transform:scale(1.08) rotate(2deg)}100%{transform:scale(1)}}@keyframes confettiFall{0%{transform:translate3d(0,-10vh,0) rotate(0);opacity:1}100%{transform:translate3d(var(--drift),115vh,0) rotate(900deg);opacity:.1}}@keyframes fireworkBurst{0%{transform:scale(.08);opacity:0}18%{opacity:1}70%{transform:scale(1);opacity:.9}100%{transform:scale(1.2);opacity:0}}@keyframes sideFloat{from{transform:translateY(-7px) rotate(-6deg)}to{transform:translateY(7px) rotate(6deg)}}@media(max-width:700px){.celebrate-stage{padding:18px 14px;min-height:auto}.celebrate-crown{font-size:54px}.celebrate-avatar-wrap{width:125px;height:125px}.celebrate-person-info{text-align:center;min-width:0}.celebrate-details{grid-template-columns:1fr}.celebrate-side{display:none}.firework{transform:scale(.7)}}\n',1)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('<div id="competitionPopup" class="competition-popup no-print"><div class="comp-fire">🔥🏆</div><div id="competitionPopupTitle" class="comp-pop-title">NEW ENROLLMENT!</div><div id="competitionPopupText"></div><small id="competitionPopupTime"></small></div>','<div id="competitionPopup" class="competition-popup no-print" aria-live="assertive" aria-label="Enrollment celebration">\n  <div class="celebrate-glow"></div><div class="firework fw1"></div><div class="firework fw2"></div><div class="firework fw3"></div><div class="firework fw4"></div>\n  <div class="celebrate-stage">\n    <button type="button" class="celebrate-close" id="competitionPopupClose" title="Close">×</button>\n    <div class="celebrate-side side-left">🏆🎲</div><div class="celebrate-side side-right">🎯🏆</div>\n    <div class="celebrate-crown">👑</div>\n    <div id="competitionPopupTitle" class="comp-pop-title">CONGRATULATIONS!</div>\n    <div class="celebrate-subtitle">🎉 ENROLLMENT DONE! 🎉</div>\n    <div class="celebrate-person">\n      <div class="celebrate-avatar-wrap"><img id="competitionEmployeePhoto" class="celebrate-avatar" alt="Assistant Manager"><div id="competitionEmployeeFallback" class="celebrate-avatar-fallback" style="display:none">👤</div></div>\n      <div class="celebrate-person-info"><div id="competitionEmployeeName" class="celebrate-name">TEAM MEMBER</div><div id="competitionEmployeeRole" class="celebrate-role">ASSISTANT MANAGER</div><div id="competitionCompanyName" class="celebrate-company"></div></div>\n    </div>\n    <div class="celebrate-details">\n      <div class="celebrate-detail"><small>Client</small><b id="competitionClientName">—</b></div>\n      <div class="celebrate-detail"><small>Enrollment ID</small><b id="competitionCaseId">—</b></div>\n      <div class="celebrate-detail"><small>Amount Received</small><b id="competitionRevenue">₹0</b></div>\n    </div>\n    <div class="celebrate-save">✓ SUCCESSFULLY SAVED!</div>\n    <div id="competitionPopupText" style="position:relative;z-index:2;margin-top:9px;font-weight:bold"></div>\n    <div id="competitionPopupTime" class="celebrate-countdown"></div>\n  </div>\n</div>',1)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace("<script>\n(function(){let last=Number(sessionStorage.getItem('gaurCompEvent')||0),first=true;const bar=document.getElementById('competitionLiveBar'),pop=document.getElementById('competitionPopup');if(!bar||!pop)return;\nfunction money(n){return '₹'+Number(n||0).toLocaleString('en-IN')}function show(e){pop.className='competition-popup no-print '+(e.company_code==='SCIC'?'scic':'wwic');document.getElementById('competitionPopupTitle').textContent=(e.company_code==='SCIC'?'ABC':'XYZ')+' SCORES! 🎉';document.getElementById('competitionPopupText').textContent=e.message||'New enrollment';document.getElementById('competitionPopupTime').textContent=e.created_at||'';requestAnimationFrame(()=>pop.classList.add('show'));setTimeout(()=>pop.classList.remove('show'),9000)}\nasync function run(){try{const r=await fetch('/api/competition/live?after='+last,{cache:'no-store'});if(!r.ok)return;const d=await r.json();const compFinished=(d.status==='finished'||!d.enabled);const isWwicOnly=d.is_wwic_only||false;const isScicOnly=d.is_scic_only||false;if(compFinished&&(isWwicOnly||isScicOnly)){bar.classList.remove('show');return}if(!d.enabled){bar.classList.remove('show');return}bar.classList.add('show');const s=d.scores||{};document.getElementById('scicLiveScore').textContent=(s.ABC?.enrollments||0)+' • '+money(s.ABC?.revenue||0);document.getElementById('wwicLiveScore').textContent=(s.XYZ?.enrollments||0)+' • '+money(s.XYZ?.revenue||0);const st=Number(d.scic_target||0),wt=Number(d.wwic_target||0);document.getElementById('scicMeter').style.width=(st?Math.min(100,(s.ABC?.enrollments||0)/st*100):0)+'%';document.getElementById('wwicMeter').style.width=(wt?Math.min(100,(s.XYZ?.enrollments||0)/wt*100):0)+'%';if(first){last=Number(d.latest_id||0);sessionStorage.setItem('gaurCompEvent',last);first=false;return}if(d.popup_enabled&&d.events?.length)d.events.forEach((e,i)=>setTimeout(()=>show(e),i*1200));last=Number(d.latest_id||last);sessionStorage.setItem('gaurCompEvent',last)}catch(e){}}\nrun();setInterval(run,5000)})();\n</script>","<script>\n(function(){\n  let last=Number(sessionStorage.getItem('gaurCompEvent')||0),first=true,hideTimer=null;\n  const bar=document.getElementById('competitionLiveBar'),pop=document.getElementById('competitionPopup');\n  if(!bar||!pop)return;\n  const photo=document.getElementById('competitionEmployeePhoto'),fallback=document.getElementById('competitionEmployeeFallback');\n  const closeBtn=document.getElementById('competitionPopupClose');if(closeBtn)closeBtn.addEventListener('click',()=>hide());\n  function money(n){return '₹'+Number(n||0).toLocaleString('en-IN')} function text(id,val){const x=document.getElementById(id);if(x)x.textContent=val??''}\n  function hide(){pop.classList.remove('show');document.querySelectorAll('.confetti-piece').forEach(x=>x.remove());if(hideTimer)clearTimeout(hideTimer)}\n  function confetti(){const colors=['#ffd64d','#ff4d67','#4da6ff','#48d58b','#f78cff','#ffffff'];for(let i=0;i<95;i++){const c=document.createElement('i');c.className='confetti-piece';c.style.left=(Math.random()*100)+'vw';c.style.background=colors[i%colors.length];c.style.setProperty('--drift',((Math.random()-.5)*280)+'px');c.style.animationDuration=(2.7+Math.random()*3)+'s';c.style.animationDelay=(Math.random()*.7)+'s';c.style.transform='rotate('+(Math.random()*360)+'deg)';pop.appendChild(c);setTimeout(()=>c.remove(),7000)}}\n  function victorySound(){try{const A=window.AudioContext||window.webkitAudioContext;if(!A)return;const ctx=new A(),now=ctx.currentTime;[523.25,659.25,783.99,1046.5].forEach((f,i)=>{const o=ctx.createOscillator(),g=ctx.createGain();o.type='triangle';o.frequency.value=f;g.gain.setValueAtTime(.0001,now+i*.13);g.gain.exponentialRampToValueAtTime(.12,now+i*.13+.02);g.gain.exponentialRampToValueAtTime(.0001,now+i*.13+.38);o.connect(g);g.connect(ctx.destination);o.start(now+i*.13);o.stop(now+i*.13+.42)});setTimeout(()=>ctx.close(),1600)}catch(e){}}\n  function show(e,soundOn){pop.className='competition-popup no-print '+(e.company_code==='SCIC'?'scic':'wwic');const brand=e.company_code==='SCIC'?'ABC Immigration Consultants':'XYZ';text('competitionPopupTitle','CONGRATULATIONS!');text('competitionEmployeeName',(e.employee_name||'TEAM MEMBER').toUpperCase());text('competitionEmployeeRole',(e.employee_role||'ASSISTANT MANAGER').toUpperCase());text('competitionCompanyName',brand+' • +1 ENROLLMENT');text('competitionClientName',e.client_name||'New Client');text('competitionCaseId',e.case_id||'—');text('competitionRevenue',money(e.revenue||0));text('competitionPopupText',e.message||brand+' SCORES!');text('competitionPopupTime',(e.created_at||'')+' • Celebration closes automatically');if(photo){photo.onerror=()=>{photo.style.display='none';fallback.style.display='flex'};if(e.photo_url){fallback.style.display='none';photo.style.display='block';photo.src=e.photo_url+'?v='+encodeURIComponent(e.id||Date.now())}else{photo.style.display='none';fallback.style.display='flex'}}requestAnimationFrame(()=>pop.classList.add('show'));confetti();if(soundOn)victorySound();if(hideTimer)clearTimeout(hideTimer);hideTimer=setTimeout(hide,9000)}\n  async function run(){try{const r=await fetch('/api/competition/live?after='+last,{cache:'no-store'});if(!r.ok)return;const d=await r.json();if(!d.enabled){bar.classList.remove('show');return}bar.classList.add('show');const s=d.scores||{};text('scicLiveScore',(s.ABC?.enrollments||0)+' • '+money(s.ABC?.revenue||0));text('wwicLiveScore',(s.XYZ?.enrollments||0)+' • '+money(s.XYZ?.revenue||0));const st=Number(d.scic_target||0),wt=Number(d.wwic_target||0);document.getElementById('scicMeter').style.width=(st?Math.min(100,(s.ABC?.enrollments||0)/st*100):0)+'%';document.getElementById('wwicMeter').style.width=(wt?Math.min(100,(s.XYZ?.enrollments||0)/wt*100):0)+'%';if(first){last=Number(d.latest_id||0);sessionStorage.setItem('gaurCompEvent',last);first=false;return}if(d.popup_enabled&&d.events?.length)d.events.forEach((e,i)=>setTimeout(()=>show(e,!!d.sound_enabled),i*9800));last=Number(d.latest_id||last);sessionStorage.setItem('gaurCompEvent',last)}catch(e){}}\n  run();setInterval(run,3000);\n})();\n</script>",1)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('CRM v3.36 • Large Company Logos','CRM v3.98 • Quick Multi Lead Allocation')
# === END GAUR CRM v3.9 PREMIUM LOGO BRANDING ===
if os.environ.get("CLOUD_MODE", "").lower() in ("1","true","yes") or os.environ.get("RAILWAY_ENVIRONMENT"):
    app.config.update(
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=True,
    )

class CursorProxy:
    def __init__(self, cur, lastrowid=None, buffered_row=None):
        self.cur=cur
        self.lastrowid=lastrowid
        self._buffered_row=buffered_row
    def fetchone(self):
        if self._buffered_row is not None:
            r=self._buffered_row; self._buffered_row=None; return r
        return self.cur.fetchone()
    def fetchall(self):
        rows=[]
        if self._buffered_row is not None:
            rows.append(self._buffered_row); self._buffered_row=None
        rows.extend(self.cur.fetchall())
        return rows

class PostgresConnection:
    def __init__(self, url):
        import psycopg
        from psycopg.rows import dict_row
        self.con=psycopg.connect(url, row_factory=dict_row, connect_timeout=15)
    def _sql(self, sql):
        # App code uses SQLite-style ? placeholders; psycopg uses %s.
        sql=sql.replace("?", "%s")
        sql=re.sub(r"INSERT\s+OR\s+IGNORE\s+INTO\s+companies\(code,name\)\s+VALUES\(%s,%s\)",
                   "INSERT INTO companies(code,name) VALUES(%s,%s) ON CONFLICT (code) DO NOTHING", sql, flags=re.I)
        return sql
    def execute(self, sql, params=()):
        sql2=self._sql(sql)
        cur=self.con.cursor()
        # Preserve SQLite's cursor.lastrowid behavior used by Employee creation.
        need_id=bool(re.match(r"\s*INSERT\s+INTO\s+employee_master\b", sql2, re.I)) and "RETURNING" not in sql2.upper()
        if need_id:
            sql2=sql2.rstrip().rstrip(';')+" RETURNING id"
        cur.execute(sql2, params or ())
        if need_id:
            row=cur.fetchone()
            rid=row["id"] if row else None
            return CursorProxy(cur,lastrowid=rid)
        return CursorProxy(cur)
    def cursor(self):
        return self
    def executescript(self, script):
        cur=self.con.cursor()
        for stmt in script.split(';'):
            if stmt.strip(): cur.execute(stmt)
        return CursorProxy(cur)
    def commit(self): self.con.commit()
    def rollback(self): self.con.rollback()
    def close(self): self.con.close()

def db():
    if IS_POSTGRES:
        return PostgresConnection(DATABASE_URL)
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.row_factory = sqlite3.Row
    try:
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("PRAGMA synchronous=NORMAL")
        con.execute("PRAGMA busy_timeout=30000")
    except Exception:
        pass
    return con

def init_db():
    con=db(); cur=con.cursor()
    if IS_POSTGRES:
        cur.executescript("""
        CREATE TABLE IF NOT EXISTS companies(
          id BIGSERIAL PRIMARY KEY, code TEXT UNIQUE NOT NULL, name TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS users(
          id BIGSERIAL PRIMARY KEY, login_id TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
          full_name TEXT NOT NULL, role TEXT NOT NULL, company_code TEXT, active INTEGER DEFAULT 1,
          photo_data BYTEA, photo_mime TEXT DEFAULT '', photo_name TEXT DEFAULT '',
          employee_id BIGINT, identity_code TEXT DEFAULT '', official_mobile TEXT DEFAULT '',
          official_email TEXT DEFAULT '', designation TEXT DEFAULT '', department TEXT DEFAULT '',
          branch TEXT DEFAULT '', joining_date TEXT DEFAULT '', reporting_user_id BIGINT,
          mobile_verified INTEGER DEFAULT 0, profile_headline TEXT DEFAULT '', profile_bio TEXT DEFAULT '',
          account_created_at TEXT DEFAULT '', last_login_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS visitors(
          id BIGSERIAL PRIMARY KEY, visitor_id TEXT UNIQUE, company_code TEXT NOT NULL, visit_date TEXT,
          client_name TEXT, mobile TEXT, email TEXT, city TEXT, country TEXT, visa_type TEXT,
          coordinating_with TEXT, purpose TEXT, status TEXT, remarks TEXT, created_by TEXT
        );
        CREATE TABLE IF NOT EXISTS employee_attendance(
          id BIGSERIAL PRIMARY KEY, attendance_date TEXT NOT NULL, employee_id BIGINT NOT NULL,
          company_code TEXT NOT NULL, entry_time TEXT NOT NULL, remarks TEXT DEFAULT '',
          marked_by TEXT DEFAULT '', marked_at TEXT DEFAULT '',
          UNIQUE(attendance_date,employee_id)
        );
        CREATE TABLE IF NOT EXISTS leads(
          id BIGSERIAL PRIMARY KEY, lead_id TEXT UNIQUE, company_code TEXT NOT NULL, client_name TEXT,
          mobile TEXT, email TEXT, city TEXT, country TEXT, visa_type TEXT, source TEXT,
          duplicate_type TEXT DEFAULT '', assigned_am BIGINT, status TEXT DEFAULT 'New',
          followup_date TEXT DEFAULT '', remarks TEXT DEFAULT '', imported_by TEXT, imported_at TEXT,
          assigned_at TEXT DEFAULT '', assigned_by TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS imports(
          id BIGSERIAL PRIMARY KEY, company_code TEXT, filename TEXT, total_rows INTEGER, new_rows INTEGER,
          same_company_duplicates INTEGER, cross_company_duplicates INTEGER, invalid_mobiles INTEGER,
          imported_by TEXT, imported_at TEXT
        );
        CREATE TABLE IF NOT EXISTS employee_master(
          id BIGSERIAL PRIMARY KEY, employee_code TEXT UNIQUE NOT NULL, company_code TEXT NOT NULL,
          full_name TEXT NOT NULL, father_spouse_name TEXT DEFAULT '', designation TEXT DEFAULT '',
          department TEXT DEFAULT '', mobile TEXT DEFAULT '', alternate_mobile TEXT DEFAULT '', email TEXT DEFAULT '',
          date_of_birth TEXT DEFAULT '', joining_date TEXT DEFAULT '', address TEXT DEFAULT '', city TEXT DEFAULT '',
          state TEXT DEFAULT '', pin_code TEXT DEFAULT '', emergency_contact_name TEXT DEFAULT '',
          emergency_contact_number TEXT DEFAULT '', aadhaar_last4 TEXT DEFAULT '', pan_number TEXT DEFAULT '',
          bank_name TEXT DEFAULT '', account_last4 TEXT DEFAULT '', ifsc_code TEXT DEFAULT '',
          salary_reference TEXT DEFAULT '', employment_status TEXT DEFAULT 'Active', remarks TEXT DEFAULT '',
          photo_path TEXT DEFAULT '', photo_data BYTEA, photo_mime TEXT DEFAULT '', portal_user_id BIGINT, reporting_employee_id BIGINT, created_by TEXT DEFAULT '',
          created_at TEXT DEFAULT '', updated_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS employee_documents(
          id BIGSERIAL PRIMARY KEY, employee_id BIGINT NOT NULL, document_type TEXT NOT NULL,
          document_name TEXT DEFAULT '', file_name TEXT NOT NULL, stored_path TEXT NOT NULL,
          file_data BYTEA, mime_type TEXT DEFAULT '', uploaded_by TEXT DEFAULT '', uploaded_at TEXT DEFAULT '',
          FOREIGN KEY(employee_id) REFERENCES employee_master(id)
        );
        CREATE TABLE IF NOT EXISTS employee_performance(
          id BIGSERIAL PRIMARY KEY, employee_id BIGINT NOT NULL, performance_date TEXT NOT NULL,
          leads_assigned INTEGER DEFAULT 0, calls_done INTEGER DEFAULT 0, followups_done INTEGER DEFAULT 0,
          office_visits INTEGER DEFAULT 0, interested_clients INTEGER DEFAULT 0, enrollments INTEGER DEFAULT 0,
          collection_amount REAL DEFAULT 0, target_amount REAL DEFAULT 0, attendance_status TEXT DEFAULT 'Present',
          manager_rating REAL DEFAULT 0, remarks TEXT DEFAULT '', created_by TEXT DEFAULT '', created_at TEXT DEFAULT '',
          FOREIGN KEY(employee_id) REFERENCES employee_master(id)
        );
        CREATE TABLE IF NOT EXISTS allocation_history(
          id BIGSERIAL PRIMARY KEY, company_code TEXT NOT NULL, am_user_id BIGINT NOT NULL,
          quantity INTEGER DEFAULT 0, allocated_by TEXT DEFAULT '', allocated_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS client_cases(
          id BIGSERIAL PRIMARY KEY, case_id TEXT UNIQUE NOT NULL, company_code TEXT NOT NULL,
          client_name TEXT NOT NULL, mobile TEXT DEFAULT '', country TEXT DEFAULT '', visa_type TEXT DEFAULT '',
          enrollment_date TEXT DEFAULT '',
          package_amount REAL DEFAULT 0, package_amount_datetime TEXT DEFAULT '',
          after_visa_payment REAL DEFAULT 0, after_visa_payment_datetime TEXT DEFAULT '',
          booking_amount REAL DEFAULT 0, first_payment_status TEXT DEFAULT 'Received', first_payment_datetime TEXT DEFAULT '',
          second_payment REAL DEFAULT 0, second_payment_status TEXT DEFAULT 'Pending', second_payment_datetime TEXT DEFAULT '',
          other_payment_received REAL DEFAULT 0, other_payment_datetime TEXT DEFAULT '',
          total_received REAL DEFAULT 0, payment_status TEXT DEFAULT 'Pending',
          filing_status TEXT DEFAULT 'Documents Pending', assigned_employee_id BIGINT, remarks TEXT DEFAULT '',
          created_by TEXT DEFAULT '', created_at TEXT DEFAULT '', updated_at TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS lead_documents(
          id BIGSERIAL PRIMARY KEY, lead_id BIGINT NOT NULL, document_type TEXT NOT NULL DEFAULT 'Client Form',
          file_name TEXT NOT NULL, file_data BYTEA, mime_type TEXT DEFAULT '', uploaded_by TEXT DEFAULT '',
          uploaded_at TEXT DEFAULT '', FOREIGN KEY(lead_id) REFERENCES leads(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS lead_activity(
          id BIGSERIAL PRIMARY KEY, lead_id BIGINT NOT NULL, status TEXT DEFAULT '', followup_date TEXT DEFAULT '',
          remarks TEXT DEFAULT '', updated_by TEXT DEFAULT '', updated_at TEXT DEFAULT '',
          FOREIGN KEY(lead_id) REFERENCES leads(id) ON DELETE CASCADE
        );
        CREATE INDEX IF NOT EXISTS idx_users_role_company ON users(role,company_code);
        CREATE INDEX IF NOT EXISTS idx_visitors_company ON visitors(company_code,id);
        CREATE INDEX IF NOT EXISTS idx_leads_company_am ON leads(company_code,assigned_am,id);
        CREATE INDEX IF NOT EXISTS idx_cases_company_employee ON client_cases(company_code,assigned_employee_id,id);
        """)
        # Cloud-safe migrations for deployments upgraded in place.
        cur.execute("ALTER TABLE employee_master ADD COLUMN IF NOT EXISTS photo_data BYTEA")
        cur.execute("ALTER TABLE employee_master ADD COLUMN IF NOT EXISTS photo_mime TEXT DEFAULT ''")
        cur.execute("ALTER TABLE employee_documents ADD COLUMN IF NOT EXISTS file_data BYTEA")
        cur.execute("ALTER TABLE employee_documents ADD COLUMN IF NOT EXISTS mime_type TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS assigned_at TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS assigned_by TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS date_of_birth TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS education TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS occupation TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS address TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS state TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS pin_code TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS alternate_mobile TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS passport_no TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS travel_history TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS english_test TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS budget TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS preferred_call_time TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS interest_score INTEGER DEFAULT 0")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS deleted_at TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS deleted_by TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS deletion_reason TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS upload_batch TEXT DEFAULT ''")
        cur.execute("ALTER TABLE leads ADD COLUMN IF NOT EXISTS import_filename TEXT DEFAULT ''")
        cur.execute("ALTER TABLE imports ADD COLUMN IF NOT EXISTS batch_code TEXT DEFAULT ''")
    else:
        cur.executescript("""
        CREATE TABLE IF NOT EXISTS companies(id INTEGER PRIMARY KEY AUTOINCREMENT,code TEXT UNIQUE NOT NULL,name TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT,login_id TEXT UNIQUE NOT NULL,password_hash TEXT NOT NULL,full_name TEXT NOT NULL,role TEXT NOT NULL,company_code TEXT,active INTEGER DEFAULT 1,photo_data BLOB,photo_mime TEXT DEFAULT '',photo_name TEXT DEFAULT '',employee_id INTEGER,
identity_code TEXT DEFAULT '',official_mobile TEXT DEFAULT '',official_email TEXT DEFAULT '',
designation TEXT DEFAULT '',department TEXT DEFAULT '',branch TEXT DEFAULT '',joining_date TEXT DEFAULT '',
reporting_user_id INTEGER,mobile_verified INTEGER DEFAULT 0,profile_headline TEXT DEFAULT '',
profile_bio TEXT DEFAULT '',account_created_at TEXT DEFAULT '',last_login_at TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS visitors(id INTEGER PRIMARY KEY AUTOINCREMENT,visitor_id TEXT UNIQUE,company_code TEXT NOT NULL,visit_date TEXT,client_name TEXT,mobile TEXT,email TEXT,city TEXT,country TEXT,visa_type TEXT,coordinating_with TEXT,purpose TEXT,status TEXT,remarks TEXT,created_by TEXT);
        CREATE TABLE IF NOT EXISTS employee_attendance(id INTEGER PRIMARY KEY AUTOINCREMENT,attendance_date TEXT NOT NULL,employee_id INTEGER NOT NULL,company_code TEXT NOT NULL,entry_time TEXT NOT NULL,remarks TEXT DEFAULT '',marked_by TEXT DEFAULT '',marked_at TEXT DEFAULT '',UNIQUE(attendance_date,employee_id));
        CREATE TABLE IF NOT EXISTS leads(id INTEGER PRIMARY KEY AUTOINCREMENT,lead_id TEXT UNIQUE,company_code TEXT NOT NULL,client_name TEXT,mobile TEXT,email TEXT,city TEXT,country TEXT,visa_type TEXT,source TEXT,duplicate_type TEXT DEFAULT '',assigned_am INTEGER,status TEXT DEFAULT 'New',followup_date TEXT DEFAULT '',remarks TEXT DEFAULT '',imported_by TEXT,imported_at TEXT,assigned_at TEXT DEFAULT '',assigned_by TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS imports(id INTEGER PRIMARY KEY AUTOINCREMENT,company_code TEXT,filename TEXT,total_rows INTEGER,new_rows INTEGER,same_company_duplicates INTEGER,cross_company_duplicates INTEGER,invalid_mobiles INTEGER,imported_by TEXT,imported_at TEXT);
        CREATE TABLE IF NOT EXISTS employee_master(id INTEGER PRIMARY KEY AUTOINCREMENT,employee_code TEXT UNIQUE NOT NULL,company_code TEXT NOT NULL,full_name TEXT NOT NULL,father_spouse_name TEXT DEFAULT '',designation TEXT DEFAULT '',department TEXT DEFAULT '',mobile TEXT DEFAULT '',alternate_mobile TEXT DEFAULT '',email TEXT DEFAULT '',date_of_birth TEXT DEFAULT '',joining_date TEXT DEFAULT '',address TEXT DEFAULT '',city TEXT DEFAULT '',state TEXT DEFAULT '',pin_code TEXT DEFAULT '',emergency_contact_name TEXT DEFAULT '',emergency_contact_number TEXT DEFAULT '',aadhaar_last4 TEXT DEFAULT '',pan_number TEXT DEFAULT '',bank_name TEXT DEFAULT '',account_last4 TEXT DEFAULT '',ifsc_code TEXT DEFAULT '',salary_reference TEXT DEFAULT '',employment_status TEXT DEFAULT 'Active',remarks TEXT DEFAULT '',photo_path TEXT DEFAULT '',photo_data BLOB,photo_mime TEXT DEFAULT '',portal_user_id INTEGER,reporting_employee_id INTEGER,created_by TEXT DEFAULT '',created_at TEXT DEFAULT '',updated_at TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS employee_documents(id INTEGER PRIMARY KEY AUTOINCREMENT,employee_id INTEGER NOT NULL,document_type TEXT NOT NULL,document_name TEXT DEFAULT '',file_name TEXT NOT NULL,stored_path TEXT NOT NULL,file_data BLOB,mime_type TEXT DEFAULT '',uploaded_by TEXT DEFAULT '',uploaded_at TEXT DEFAULT '',FOREIGN KEY(employee_id) REFERENCES employee_master(id));
        CREATE TABLE IF NOT EXISTS employee_performance(id INTEGER PRIMARY KEY AUTOINCREMENT,employee_id INTEGER NOT NULL,performance_date TEXT NOT NULL,leads_assigned INTEGER DEFAULT 0,calls_done INTEGER DEFAULT 0,followups_done INTEGER DEFAULT 0,office_visits INTEGER DEFAULT 0,interested_clients INTEGER DEFAULT 0,enrollments INTEGER DEFAULT 0,collection_amount REAL DEFAULT 0,target_amount REAL DEFAULT 0,attendance_status TEXT DEFAULT 'Present',manager_rating REAL DEFAULT 0,remarks TEXT DEFAULT '',created_by TEXT DEFAULT '',created_at TEXT DEFAULT '',FOREIGN KEY(employee_id) REFERENCES employee_master(id));
        CREATE TABLE IF NOT EXISTS allocation_history(id INTEGER PRIMARY KEY AUTOINCREMENT,company_code TEXT NOT NULL,am_user_id INTEGER NOT NULL,quantity INTEGER DEFAULT 0,allocated_by TEXT DEFAULT '',allocated_at TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS client_cases(id INTEGER PRIMARY KEY AUTOINCREMENT,case_id TEXT UNIQUE NOT NULL,company_code TEXT NOT NULL,client_name TEXT NOT NULL,mobile TEXT DEFAULT '',country TEXT DEFAULT '',visa_type TEXT DEFAULT '',enrollment_date TEXT DEFAULT '',package_amount REAL DEFAULT 0,package_amount_datetime TEXT DEFAULT '',after_visa_payment REAL DEFAULT 0,after_visa_payment_datetime TEXT DEFAULT '',booking_amount REAL DEFAULT 0,first_payment_status TEXT DEFAULT 'Received',first_payment_datetime TEXT DEFAULT '',second_payment REAL DEFAULT 0,second_payment_status TEXT DEFAULT 'Pending',second_payment_datetime TEXT DEFAULT '',other_payment_received REAL DEFAULT 0,other_payment_datetime TEXT DEFAULT '',total_received REAL DEFAULT 0,payment_status TEXT DEFAULT 'Pending',filing_status TEXT DEFAULT 'Documents Pending',assigned_employee_id INTEGER,remarks TEXT DEFAULT '',created_by TEXT DEFAULT '',created_at TEXT DEFAULT '',updated_at TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS lead_documents(id INTEGER PRIMARY KEY AUTOINCREMENT,lead_id INTEGER NOT NULL,document_type TEXT NOT NULL DEFAULT 'Client Form',file_name TEXT NOT NULL,file_data BLOB,mime_type TEXT DEFAULT '',uploaded_by TEXT DEFAULT '',uploaded_at TEXT DEFAULT '',FOREIGN KEY(lead_id) REFERENCES leads(id) ON DELETE CASCADE);
        CREATE TABLE IF NOT EXISTS lead_activity(id INTEGER PRIMARY KEY AUTOINCREMENT,lead_id INTEGER NOT NULL,status TEXT DEFAULT '',followup_date TEXT DEFAULT '',remarks TEXT DEFAULT '',updated_by TEXT DEFAULT '',updated_at TEXT DEFAULT '',FOREIGN KEY(lead_id) REFERENCES leads(id) ON DELETE CASCADE);
        """)
        cols=[r[1] for r in cur.execute("PRAGMA table_info(employee_master)").fetchall()]
        for col,ddl in [("photo_path","TEXT DEFAULT ''"),("photo_data","BLOB"),("photo_mime","TEXT DEFAULT ''")]:
            if col not in cols: cur.execute(f"ALTER TABLE employee_master ADD COLUMN {col} {ddl}")
        dcols=[r[1] for r in cur.execute("PRAGMA table_info(employee_documents)").fetchall()]
        for col,ddl in [("file_data","BLOB"),("mime_type","TEXT DEFAULT ''")]:
            if col not in dcols: cur.execute(f"ALTER TABLE employee_documents ADD COLUMN {col} {ddl}")
        lead_cols=[r[1] for r in cur.execute("PRAGMA table_info(leads)").fetchall()]
        if "assigned_at" not in lead_cols: cur.execute("ALTER TABLE leads ADD COLUMN assigned_at TEXT DEFAULT ''")
        if "assigned_by" not in lead_cols: cur.execute("ALTER TABLE leads ADD COLUMN assigned_by TEXT DEFAULT ''")
        for col,ddl in [("date_of_birth","TEXT DEFAULT ''"),("education","TEXT DEFAULT ''"),("occupation","TEXT DEFAULT ''"),("address","TEXT DEFAULT ''"),("state","TEXT DEFAULT ''"),("pin_code","TEXT DEFAULT ''"),("alternate_mobile","TEXT DEFAULT ''"),("passport_no","TEXT DEFAULT ''"),("travel_history","TEXT DEFAULT ''"),("english_test","TEXT DEFAULT ''"),("budget","TEXT DEFAULT ''"),("preferred_call_time","TEXT DEFAULT ''"),("interest_score","INTEGER DEFAULT 0")]:
            if col not in lead_cols: cur.execute(f"ALTER TABLE leads ADD COLUMN {col} {ddl}")
        for col,ddl in [("deleted_at","TEXT DEFAULT ''"),("deleted_by","TEXT DEFAULT ''"),("deletion_reason","TEXT DEFAULT ''")]:
            if col not in lead_cols: cur.execute(f"ALTER TABLE leads ADD COLUMN {col} {ddl}")
        for col,ddl in [("upload_batch","TEXT DEFAULT ''"),("import_filename","TEXT DEFAULT ''")]:
            if col not in lead_cols: cur.execute(f"ALTER TABLE leads ADD COLUMN {col} {ddl}")
        import_cols=[r[1] for r in cur.execute("PRAGMA table_info(imports)").fetchall()]
        if "batch_code" not in import_cols: cur.execute("ALTER TABLE imports ADD COLUMN batch_code TEXT DEFAULT ''")


    # v3.98 structured enrollment/payment ledger
    payment_case_cols=[
        ("package_amount","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
        ("package_amount_datetime","TEXT DEFAULT ''"),
        ("after_visa_payment","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
        ("after_visa_payment_datetime","TEXT DEFAULT ''"),
        ("first_payment_status","TEXT DEFAULT 'Received'"),
        ("first_payment_datetime","TEXT DEFAULT ''"),
        ("second_payment_status","TEXT DEFAULT 'Pending'"),
        ("second_payment_datetime","TEXT DEFAULT ''"),
        ("other_payment_received","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
        ("other_payment_datetime","TEXT DEFAULT ''")
    ]
    if IS_POSTGRES:
        for col,ctype in payment_case_cols:
            cur.execute(f"ALTER TABLE client_cases ADD COLUMN IF NOT EXISTS {col} {ctype}")
    else:
        case_cols=[r["name"] for r in cur.execute("PRAGMA table_info(client_cases)").fetchall()]
        for col,ctype in payment_case_cols:
            if col not in case_cols:
                cur.execute(f"ALTER TABLE client_cases ADD COLUMN {col} {ctype}")

    # Backward compatibility: old booking_amount is now First Payment.
    # Old payment_status represented the old second-payment state.
    cur.execute("""UPDATE client_cases SET first_payment_status='Received'
                   WHERE COALESCE(booking_amount,0)>0
                     AND COALESCE(TRIM(first_payment_status),'') IN ('','Pending')""")
    cur.execute("""UPDATE client_cases SET second_payment_status=COALESCE(NULLIF(TRIM(payment_status),''),'Pending')
                   WHERE COALESCE(TRIM(second_payment_status),'') IN ('','Pending')
                     AND COALESCE(TRIM(payment_status),'')<>''""")
    # Historical records may not have individual timestamps; preserve the best available audit time.
    cur.execute("""UPDATE client_cases SET first_payment_datetime=COALESCE(NULLIF(first_payment_datetime,''),NULLIF(created_at,''),NULLIF(enrollment_date,''))
                   WHERE COALESCE(booking_amount,0)>0 AND COALESCE(first_payment_datetime,'')=''""")
    cur.execute("""UPDATE client_cases SET second_payment_datetime=COALESCE(NULLIF(second_payment_datetime,''),NULLIF(updated_at,''),NULLIF(enrollment_date,''))
                   WHERE COALESCE(second_payment,0)>0 AND COALESCE(second_payment_datetime,'')=''""")
    cur.execute("""UPDATE client_cases SET package_amount_datetime=COALESCE(NULLIF(package_amount_datetime,''),NULLIF(created_at,''),NULLIF(enrollment_date,''))
                   WHERE COALESCE(package_amount,0)>0 AND COALESCE(package_amount_datetime,'')=''""")
    cur.execute("""UPDATE client_cases SET after_visa_payment_datetime=COALESCE(NULLIF(after_visa_payment_datetime,''),NULLIF(created_at,''),NULLIF(enrollment_date,''))
                   WHERE COALESCE(after_visa_payment,0)>0 AND COALESCE(after_visa_payment_datetime,'')=''""")


    # v3.98 payment-channel tracking for bank-wise Accounts reports
    v361_case_cols=[
      ("first_payment_method","TEXT DEFAULT ''"),
      ("second_payment_method","TEXT DEFAULT ''"),
      ("other_payment_method","TEXT DEFAULT ''"),
      ("other_payment_bank","TEXT DEFAULT ''")
    ]
    if IS_POSTGRES:
        for col,ctype in v361_case_cols:
            cur.execute(f"ALTER TABLE client_cases ADD COLUMN IF NOT EXISTS {col} {ctype}")
    else:
        _v361cols=[r["name"] for r in cur.execute("PRAGMA table_info(client_cases)").fetchall()]
        for col,ctype in v361_case_cols:
            if col not in _v361cols:
                cur.execute(f"ALTER TABLE client_cases ADD COLUMN {col} {ctype}")

    # v3.98 Excel-style split payments + lead/counselor interlink
    v343_case_cols=[
      ("lead_db_id","BIGINT" if IS_POSTGRES else "INTEGER"),("counselor_name","TEXT DEFAULT ''"),
      ("am_name","TEXT DEFAULT ''"),("interaction_date","TEXT DEFAULT ''"),
      ("first_cash","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("first_rbl","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("first_yes_bank","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("first_au_bank","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("first_indusind","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("second_cash","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("second_rbl","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("second_yes_bank","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("second_au_bank","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"),
      ("second_indusind","DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0")]
    if IS_POSTGRES:
        for col,ctype in v343_case_cols: cur.execute(f"ALTER TABLE client_cases ADD COLUMN IF NOT EXISTS {col} {ctype}")
    else:
        existing=[r["name"] for r in cur.execute("PRAGMA table_info(client_cases)").fetchall()]
        for col,ctype in v343_case_cols:
            if col not in existing: cur.execute(f"ALTER TABLE client_cases ADD COLUMN {col} {ctype}")


    # v3.98 profile personalization
    v358_user_cols=[
      ("profile_cover_data","BYTEA" if IS_POSTGRES else "BLOB"),
      ("profile_cover_mime","TEXT DEFAULT ''"),
      ("profile_cover_name","TEXT DEFAULT ''"),
      ("profile_theme","TEXT DEFAULT 'royal'"),
      ("appearance_mode","TEXT DEFAULT 'dark'")
    ]
    if IS_POSTGRES:
        for col,ctype in v358_user_cols:
            cur.execute(f"ALTER TABLE users ADD COLUMN IF NOT EXISTS {col} {ctype}")
    else:
        existing_v358=[r["name"] for r in cur.execute("PRAGMA table_info(users)").fetchall()]
        for col,ctype in v358_user_cols:
            if col not in existing_v358:
                cur.execute(f"ALTER TABLE users ADD COLUMN {col} {ctype}")

    # v3.15 professional identity fields for existing databases
    if IS_POSTGRES:
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS photo_data BYTEA")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS photo_mime TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS photo_name TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS employee_id BIGINT")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS identity_code TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS official_mobile TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS official_email TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS designation TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS department TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS branch TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS joining_date TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS reporting_user_id BIGINT")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS mobile_verified INTEGER DEFAULT 0")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS profile_headline TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS profile_bio TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS account_created_at TEXT DEFAULT ''")
        cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS last_login_at TEXT DEFAULT ''")
        cur.execute("ALTER TABLE employee_master ADD COLUMN IF NOT EXISTS portal_user_id BIGINT")
        cur.execute("ALTER TABLE employee_master ADD COLUMN IF NOT EXISTS reporting_employee_id BIGINT")
    else:
        existing_cols=[r["name"] for r in cur.execute("PRAGMA table_info(users)").fetchall()]
        user_cols=[
            ("photo_data","BLOB"),("photo_mime","TEXT DEFAULT ''"),("photo_name","TEXT DEFAULT ''"),
            ("employee_id","INTEGER"),("identity_code","TEXT DEFAULT ''"),("official_mobile","TEXT DEFAULT ''"),
            ("official_email","TEXT DEFAULT ''"),("designation","TEXT DEFAULT ''"),("department","TEXT DEFAULT ''"),
            ("branch","TEXT DEFAULT ''"),("joining_date","TEXT DEFAULT ''"),("reporting_user_id","INTEGER"),
            ("mobile_verified","INTEGER DEFAULT 0"),("profile_headline","TEXT DEFAULT ''"),("profile_bio","TEXT DEFAULT ''"),
            ("account_created_at","TEXT DEFAULT ''"),("last_login_at","TEXT DEFAULT ''")
        ]
        for col,ctype in user_cols:
            if col not in existing_cols:
                cur.execute(f"ALTER TABLE users ADD COLUMN {col} {ctype}")
        emp_cols=[r["name"] for r in cur.execute("PRAGMA table_info(employee_master)").fetchall()]
        for col,ctype in [("portal_user_id","INTEGER"),("reporting_employee_id","INTEGER")]:
            if col not in emp_cols:
                cur.execute(f"ALTER TABLE employee_master ADD COLUMN {col} {ctype}")

    # v3.17 internal communication tables
    if IS_POSTGRES:
        cur.execute("""CREATE TABLE IF NOT EXISTS chat_messages(
          id BIGSERIAL PRIMARY KEY,sender_id BIGINT NOT NULL,recipient_id BIGINT NOT NULL,
          message TEXT NOT NULL DEFAULT '',lead_id BIGINT,created_at TEXT DEFAULT '',
          read_at TEXT DEFAULT '',deleted_by_sender INTEGER DEFAULT 0,deleted_by_recipient INTEGER DEFAULT 0
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS broadcasts(
          id BIGSERIAL PRIMARY KEY,sender_id BIGINT NOT NULL,audience TEXT NOT NULL DEFAULT '',
          company_scope TEXT DEFAULT '',message_type TEXT DEFAULT 'ANNOUNCEMENT',
          message TEXT NOT NULL DEFAULT '',created_at TEXT DEFAULT ''
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS broadcast_reads(
          id BIGSERIAL PRIMARY KEY,broadcast_id BIGINT NOT NULL,user_id BIGINT NOT NULL,read_at TEXT DEFAULT ''
        )""")
    else:
        cur.execute("""CREATE TABLE IF NOT EXISTS chat_messages(
          id INTEGER PRIMARY KEY AUTOINCREMENT,sender_id INTEGER NOT NULL,recipient_id INTEGER NOT NULL,
          message TEXT NOT NULL DEFAULT '',lead_id INTEGER,created_at TEXT DEFAULT '',
          read_at TEXT DEFAULT '',deleted_by_sender INTEGER DEFAULT 0,deleted_by_recipient INTEGER DEFAULT 0
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS broadcasts(
          id INTEGER PRIMARY KEY AUTOINCREMENT,sender_id INTEGER NOT NULL,audience TEXT NOT NULL DEFAULT '',
          company_scope TEXT DEFAULT '',message_type TEXT DEFAULT 'ANNOUNCEMENT',
          message TEXT NOT NULL DEFAULT '',created_at TEXT DEFAULT ''
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS broadcast_reads(
          id INTEGER PRIMARY KEY AUTOINCREMENT,broadcast_id INTEGER NOT NULL,user_id INTEGER NOT NULL,read_at TEXT DEFAULT ''
        )""")

    # v3.18 employee-uploaded chat attachments (separate from official CRM data)
    if IS_POSTGRES:
        cur.execute("""CREATE TABLE IF NOT EXISTS chat_attachments(
          id BIGSERIAL PRIMARY KEY,uploader_id BIGINT NOT NULL,original_name TEXT NOT NULL DEFAULT '',
          mime_type TEXT DEFAULT '',file_bytes BYTEA NOT NULL,created_at TEXT DEFAULT ''
        )""")
        # Idempotent PostgreSQL migration; avoids transaction rollback on every restart.
        cur.execute("ALTER TABLE chat_messages ADD COLUMN IF NOT EXISTS attachment_id BIGINT")
    else:
        cur.execute("""CREATE TABLE IF NOT EXISTS chat_attachments(
          id INTEGER PRIMARY KEY AUTOINCREMENT,uploader_id INTEGER NOT NULL,original_name TEXT NOT NULL DEFAULT '',
          mime_type TEXT DEFAULT '',file_bytes BLOB NOT NULL,created_at TEXT DEFAULT ''
        )""")
        try: cur.execute("ALTER TABLE chat_messages ADD COLUMN attachment_id INTEGER")
        except Exception: pass

    # v3.19 notification read tracking
    if IS_POSTGRES:
        cur.execute("""CREATE TABLE IF NOT EXISTS notification_reads(
          id BIGSERIAL PRIMARY KEY,user_id BIGINT NOT NULL,notification_key TEXT NOT NULL,read_at TEXT DEFAULT ''
        )""")
    else:
        cur.execute("""CREATE TABLE IF NOT EXISTS notification_reads(
          id INTEGER PRIMARY KEY AUTOINCREMENT,user_id INTEGER NOT NULL,notification_key TEXT NOT NULL,read_at TEXT DEFAULT ''
        )""")

    # v3.21 safe indexes for chat/notification lookups
    try:
        cur.execute("CREATE INDEX IF NOT EXISTS idx_chat_recipient_read ON chat_messages(recipient_id,read_at)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_chat_sender_recipient ON chat_messages(sender_id,recipient_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_broadcast_reads_user ON broadcast_reads(user_id,broadcast_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_notification_reads_user ON notification_reads(user_id,notification_key)")
    except Exception:
        # Index creation should never prevent the web service from starting.
        try: con.rollback()
        except Exception: pass

    # v3.22 Activity & Security audit trail
    if IS_POSTGRES:
        cur.execute("""CREATE TABLE IF NOT EXISTS activity_logs(
          id BIGSERIAL PRIMARY KEY,user_id BIGINT,user_name TEXT DEFAULT '',user_role TEXT DEFAULT '',user_designation TEXT DEFAULT '',
          company_code TEXT DEFAULT '',category TEXT DEFAULT 'GENERAL',action_code TEXT DEFAULT '',action_label TEXT DEFAULT '',
          entity_type TEXT DEFAULT '',entity_id TEXT DEFAULT '',entity_label TEXT DEFAULT '',details TEXT DEFAULT '',
          old_value TEXT DEFAULT '',new_value TEXT DEFAULT '',severity TEXT DEFAULT 'INFO',ip_address TEXT DEFAULT '',
          request_method TEXT DEFAULT '',request_path TEXT DEFAULT '',user_agent TEXT DEFAULT '',created_at TEXT DEFAULT ''
        )""")
    else:
        cur.execute("""CREATE TABLE IF NOT EXISTS activity_logs(
          id INTEGER PRIMARY KEY AUTOINCREMENT,user_id INTEGER,user_name TEXT DEFAULT '',user_role TEXT DEFAULT '',user_designation TEXT DEFAULT '',
          company_code TEXT DEFAULT '',category TEXT DEFAULT 'GENERAL',action_code TEXT DEFAULT '',action_label TEXT DEFAULT '',
          entity_type TEXT DEFAULT '',entity_id TEXT DEFAULT '',entity_label TEXT DEFAULT '',details TEXT DEFAULT '',
          old_value TEXT DEFAULT '',new_value TEXT DEFAULT '',severity TEXT DEFAULT 'INFO',ip_address TEXT DEFAULT '',
          request_method TEXT DEFAULT '',request_path TEXT DEFAULT '',user_agent TEXT DEFAULT '',created_at TEXT DEFAULT ''
        )""")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_activity_created ON activity_logs(created_at)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_activity_user ON activity_logs(user_id,created_at)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_activity_company ON activity_logs(company_code,created_at)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_activity_category ON activity_logs(category,created_at)")

    # v3.23 Finance, banking, bills and GST receipt ledger
    finance_tables=[
    """CREATE TABLE IF NOT EXISTS finance_banks(id {pk},company_code TEXT NOT NULL,bank_name TEXT NOT NULL,account_name TEXT NOT NULL,account_number TEXT NOT NULL,account_last4 TEXT DEFAULT '',ifsc_code TEXT DEFAULT '',branch TEXT DEFAULT '',account_type TEXT DEFAULT 'Current',opening_balance REAL DEFAULT 0,active INTEGER DEFAULT 1,created_by TEXT DEFAULT '',created_at TEXT DEFAULT '')""",
    """CREATE TABLE IF NOT EXISTS finance_transactions(id {pk},voucher_no TEXT UNIQUE NOT NULL,txn_date TEXT NOT NULL,company_code TEXT NOT NULL,direction TEXT NOT NULL,head TEXT NOT NULL,party_name TEXT NOT NULL,client_ref TEXT DEFAULT '',amount REAL DEFAULT 0,payment_mode TEXT DEFAULT '',account_id INTEGER,reference_no TEXT DEFAULT '',taxable_value REAL DEFAULT 0,gst_amount REAL DEFAULT 0,remarks TEXT DEFAULT '',bill_name TEXT DEFAULT '',bill_mime TEXT DEFAULT '',bill_bytes {blob},status TEXT DEFAULT 'POSTED',created_by_id INTEGER,created_by_name TEXT DEFAULT '',created_at TEXT DEFAULT '')"""
    ]
    for q in finance_tables:
        cur.execute(q.format(pk=("BIGSERIAL PRIMARY KEY" if IS_POSTGRES else "INTEGER PRIMARY KEY AUTOINCREMENT"),blob=("BYTEA" if IS_POSTGRES else "BLOB")))
    cur.execute("CREATE INDEX IF NOT EXISTS idx_finance_txn_date ON finance_transactions(txn_date,company_code)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_finance_txn_head ON finance_transactions(head,direction)")
    # v3.36.1: Bank Manager schema upgrade is lazy; do not block Railway startup.\n\n    # v3.24 AI-style bank reconciliation
    if IS_POSTGRES:
        cur.execute("""CREATE TABLE IF NOT EXISTS bank_recon_runs(
          id BIGSERIAL PRIMARY KEY,company_code TEXT NOT NULL,bank_id BIGINT,file_name TEXT DEFAULT '',
          total_rows INTEGER DEFAULT 0,created_by_id BIGINT,created_by_name TEXT DEFAULT '',created_at TEXT DEFAULT ''
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS bank_statement_rows(
          id BIGSERIAL PRIMARY KEY,run_id BIGINT NOT NULL,txn_date TEXT DEFAULT '',txn_time TEXT DEFAULT '',
          narration TEXT DEFAULT '',reference_no TEXT DEFAULT '',debit REAL DEFAULT 0,credit REAL DEFAULT 0,
          balance REAL DEFAULT 0,raw_json TEXT DEFAULT ''
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS bank_recon_matches(
          id BIGSERIAL PRIMARY KEY,run_id BIGINT NOT NULL,portal_txn_id BIGINT,bank_row_id BIGINT,
          score REAL DEFAULT 0,status TEXT DEFAULT 'POSSIBLE',match_reason TEXT DEFAULT '',confirmed_by_id BIGINT,
          confirmed_at TEXT DEFAULT ''
        )""")
    else:
        cur.execute("""CREATE TABLE IF NOT EXISTS bank_recon_runs(
          id INTEGER PRIMARY KEY AUTOINCREMENT,company_code TEXT NOT NULL,bank_id INTEGER,file_name TEXT DEFAULT '',
          total_rows INTEGER DEFAULT 0,created_by_id INTEGER,created_by_name TEXT DEFAULT '',created_at TEXT DEFAULT ''
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS bank_statement_rows(
          id INTEGER PRIMARY KEY AUTOINCREMENT,run_id INTEGER NOT NULL,txn_date TEXT DEFAULT '',txn_time TEXT DEFAULT '',
          narration TEXT DEFAULT '',reference_no TEXT DEFAULT '',debit REAL DEFAULT 0,credit REAL DEFAULT 0,
          balance REAL DEFAULT 0,raw_json TEXT DEFAULT ''
        )""")
        cur.execute("""CREATE TABLE IF NOT EXISTS bank_recon_matches(
          id INTEGER PRIMARY KEY AUTOINCREMENT,run_id INTEGER NOT NULL,portal_txn_id INTEGER,bank_row_id INTEGER,
          score REAL DEFAULT 0,status TEXT DEFAULT 'POSSIBLE',match_reason TEXT DEFAULT '',confirmed_by_id INTEGER,
          confirmed_at TEXT DEFAULT ''
        )""")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_recon_rows_run ON bank_statement_rows(run_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_recon_matches_run ON bank_recon_matches(run_id,status)")
    for code,name in [("SCIC","ABC Immigration Consultants"),("WWIC","XYZ Immigration Consultants")]:
        if IS_POSTGRES:
            cur.execute("INSERT INTO companies(code,name) VALUES(?,?) ON CONFLICT (code) DO NOTHING",(code,name))
        else:
            cur.execute("INSERT OR IGNORE INTO companies(code,name) VALUES(?,?)",(code,name))

    defaults=[
      ("md.admin","Managing Director","MD",None,"Admin@2026"),
      ("gm.smartchoice","ABC General Manager","GM","SCIC","ABC@GM2026#"),
      ("gm.whitewave","XYZ General Manager","GM","WWIC","XYZ@GM2026#"),
      ("reception.scic","ABC Receptionist","RECEPTION","SCIC","Welcome@123"),
      ("reception.wwic","XYZ Receptionist","RECEPTION","WWIC","Welcome@123"),
    ]
    legacy_map={"gm.scic":"gm.smartchoice","gm.wwic":"gm.whitewave"}
    for old_id,new_id in legacy_map.items():
        oldrow=cur.execute("SELECT id FROM users WHERE lower(login_id)=?",(old_id,)).fetchone()
        newrow=cur.execute("SELECT id FROM users WHERE lower(login_id)=?",(new_id,)).fetchone()
        if oldrow and not newrow: cur.execute("UPDATE users SET login_id=? WHERE id=?",(new_id,oldrow["id"]))

    for login_id,name,role,company,pw in defaults:
        existing=cur.execute("SELECT id FROM users WHERE lower(login_id)=lower(?)",(login_id,)).fetchone()
        if not existing:
            if company is None:
                existing=cur.execute("SELECT id FROM users WHERE role=? AND company_code IS NULL ORDER BY id LIMIT 1",(role,)).fetchone()
            elif role in ("GM","RECEPTION"):
                existing=cur.execute("SELECT id FROM users WHERE role=? AND company_code=? ORDER BY id LIMIT 1",(role,company)).fetchone()
        if existing:
            # Preserve current Login ID and password across redeploys.
            cur.execute("UPDATE users SET full_name=?,role=?,company_code=?,active=1 WHERE id=?",
                        (name,role,company,existing["id"]))
        else:
            cur.execute("INSERT INTO users(login_id,password_hash,full_name,role,company_code,active) VALUES(?,?,?,?,?,1)",
                        (login_id,generate_password_hash(pw),name,role,company))
    con.commit(); con.close()

def current_user():
    uid = session.get("uid")
    if not uid: return None
    con=db()
    u=con.execute("SELECT * FROM users WHERE id=? AND active=1",(uid,)).fetchone()
    con.close()
    return u

def require_roles(*roles):
    def deco(fn):
        from functools import wraps
        @wraps(fn)
        def wrap(*a, **kw):
            u=current_user()
            if not u: return redirect(url_for("login"))
            if roles and u["role"] not in roles:
                flash("Access denied","error")
                return redirect(url_for("dashboard"))
            return fn(*a,**kw)
        return wrap
    return deco

@app.route("/healthz")
def healthz():
    # Railway liveness endpoint: if Flask/Waitress can answer HTTP, the service is alive.
    # Do NOT make liveness depend on PostgreSQL; temporary DB latency should not kill deployment.
    return jsonify({"status":"ok","service":"the-gaur-crm","version":"3.37"}),200

@app.route("/readyz")
def readyz():
    # Separate readiness/database diagnostic endpoint.
    try:
        con=db()
        con.execute("SELECT 1").fetchone()
        con.close()
        return jsonify({"status":"ready","database":"postgresql" if IS_POSTGRES else "sqlite"}),200
    except Exception as exc:
        return jsonify({"status":"degraded","detail":str(exc)[:200]}),503

@app.route("/", methods=["GET","POST"])
def login():
    if current_user():
        return redirect(url_for("dashboard"))

    if request.method=="POST":
        lid=request.form.get("login_id","").strip().lower()
        pw=request.form.get("password","")
        con=db()
        u=con.execute("SELECT * FROM users WHERE lower(login_id)=? AND active=1",(lid,)).fetchone()
        con.close()

        if u and check_password_hash(u["password_hash"],pw):
            # No company or role is revealed before authentication.
            # Company/role access is resolved only AFTER valid credentials.
            session["uid"]=u["id"]
            try:
                con=db()
                con.execute("UPDATE users SET last_login_at=? WHERE id=?",(datetime.datetime.now().isoformat(timespec="seconds"),u["id"]))
                con.commit(); con.close()
            except Exception:
                pass
            log_activity("LOGIN_SUCCESS","Successful Login","SECURITY",details="Portal login successful",actor=u)
            return redirect(url_for("dashboard"))

        flash("Access denied. Invalid Login ID or password.","error")

    return render_template("login_gaur.html")



@app.route("/user-cover/<int:user_id>")
def user_cover(user_id):
    if not current_user(): return ("",401)
    con=db()
    row=con.execute("SELECT profile_cover_data,profile_cover_mime FROM users WHERE id=?",(user_id,)).fetchone()
    con.close()
    if not row or not row["profile_cover_data"]: return ("",404)
    data=row["profile_cover_data"]
    if isinstance(data,memoryview): data=data.tobytes()
    resp=send_file(io.BytesIO(bytes(data)),mimetype=row["profile_cover_mime"] or "image/jpeg",max_age=0)
    resp.headers["Cache-Control"]="no-store, no-cache, must-revalidate, max-age=0"
    return resp

@app.route("/user-photo/<int:user_id>")
def user_photo(user_id):
    viewer=current_user()
    if not viewer:
        return redirect(url_for("login"))
    con=db()
    row=con.execute("SELECT id,photo_data,photo_mime FROM users WHERE id=?",(user_id,)).fetchone()
    con.close()
    if not row or not row["photo_data"]:
        return ("",404)
    resp=send_file(io.BytesIO(bytes(row["photo_data"])),mimetype=row["photo_mime"] or "image/jpeg",
                   download_name="profile-photo",max_age=0)
    resp.headers["Cache-Control"]="no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"]="no-cache"
    resp.headers["Expires"]="0"
    return resp


@app.route("/my-profile/account",methods=["POST"])
def my_profile_account():
    u=current_user()
    if not u:
        return redirect(url_for("login"))

    current_password=request.form.get("current_password","")
    new_login_id=(request.form.get("new_login_id") or u["login_id"] or "").strip().lower()
    new_password=request.form.get("new_password","")
    confirm_password=request.form.get("confirm_password","")

    con=db()
    fresh=con.execute("SELECT * FROM users WHERE id=?",(u["id"],)).fetchone()
    if not fresh or not check_password_hash(fresh["password_hash"],current_password):
        con.close()
        flash("Current password is incorrect. No account changes were made.","error")
        return redirect(url_for("my_profile"))

    if not re.fullmatch(r"[a-z0-9._-]{4,50}",new_login_id or ""):
        con.close()
        flash("Login ID must be 4-50 characters and may use letters, numbers, dot, underscore or hyphen.","error")
        return redirect(url_for("my_profile"))

    duplicate=con.execute("SELECT id FROM users WHERE lower(login_id)=lower(?) AND id<>?",(new_login_id,u["id"])).fetchone()
    if duplicate:
        con.close()
        flash("This Login ID is already in use. Please choose another Login ID.","error")
        return redirect(url_for("my_profile"))

    password_change=bool(new_password or confirm_password)
    if password_change:
        if len(new_password)<8:
            con.close()
            flash("New password must be at least 8 characters.","error")
            return redirect(url_for("my_profile"))
        if new_password!=confirm_password:
            con.close()
            flash("New password and Confirm Password do not match.","error")
            return redirect(url_for("my_profile"))

    old_login=fresh["login_id"]
    if password_change:
        con.execute("UPDATE users SET login_id=?,password_hash=? WHERE id=?",
                    (new_login_id,generate_password_hash(new_password),u["id"]))
    else:
        con.execute("UPDATE users SET login_id=? WHERE id=?",(new_login_id,u["id"]))
    con.commit()
    con.close()

    try:
        log_activity("PROFILE_LOGIN_UPDATED","Profile Login / Password Updated","SECURITY",
                     "User Account",u["id"],new_login_id,
                     details={"old_login_id":old_login,"login_id_changed":old_login!=new_login_id,
                              "password_changed":password_change},
                     actor=current_user(),company_code=u["company_code"])
    except Exception:
        pass

    flash("Login / security settings updated successfully.","success")
    return redirect(url_for("my_profile"))


@app.route("/my-profile",methods=["GET","POST"])
def my_profile():
    u=current_user()
    if not u:
        return redirect(url_for("login"))
    con=db()
    if request.method=="POST":
        profile_name=(request.form.get("profile_name") or u["full_name"] or "").strip()[:100]
        if len(profile_name)<2:
            con.close(); flash("Profile name must be at least 2 characters.","error")
            return redirect(url_for("my_profile"))
        headline=request.form.get("profile_headline","").strip()[:120]
        bio=request.form.get("profile_bio","").strip()[:600]
        official_mobile=normalize_mobile(request.form.get("official_mobile",""))
        official_email=request.form.get("official_email","").strip()[:160]
        department=request.form.get("department","").strip()[:100]
        branch=request.form.get("branch","").strip()[:120]
        profile_theme=(request.form.get("profile_theme") or "royal").strip()[:30]
        appearance_mode=(request.form.get("appearance_mode") or "dark").strip()[:20]
        if profile_theme not in ("royal","flowers","nature","minimal","ocean","sunset"): profile_theme="royal"
        if appearance_mode not in ("dark","light","system"): appearance_mode="dark"
        if official_mobile and len(official_mobile)!=10:
            con.close(); flash("Official mobile must be a valid 10-digit number.","error")
            return redirect(url_for("my_profile"))
        if official_mobile:
            duplicate=con.execute("SELECT id FROM users WHERE official_mobile=? AND id<>?",(official_mobile,u["id"])).fetchone()
            if duplicate:
                con.close(); flash("This official mobile number is already assigned to another portal user.","error")
                return redirect(url_for("my_profile"))
        cover=request.files.get("profile_cover")
        if cover and cover.filename:
            if not allowed_file(cover.filename,ALLOWED_PHOTO_EXTENSIONS):
                con.close(); flash("Cover image must be JPG, JPEG, PNG or WEBP.","error")
                return redirect(url_for("my_profile"))
            cdata=cover.read()
            if cdata:
                cmime=cover.mimetype or mimetypes.guess_type(cover.filename)[0] or "image/jpeg"
                cname=secure_filename(cover.filename)
                con.execute("UPDATE users SET profile_cover_data=?,profile_cover_mime=?,profile_cover_name=? WHERE id=?",
                            (cdata,cmime,cname,u["id"]))
        f=request.files.get("profile_photo")
        if f and f.filename:
            if not allowed_file(f.filename,ALLOWED_PHOTO_EXTENSIONS):
                con.close(); flash("Only JPG, JPEG, PNG or WEBP photos are allowed.","error")
                return redirect(url_for("my_profile"))
            data=f.read()
            if not data:
                con.close(); flash("Selected photo file is empty. Please choose the image again.","error")
                return redirect(url_for("my_profile"))
            mime=f.mimetype or mimetypes.guess_type(f.filename)[0] or "image/jpeg"
            name=secure_filename(f.filename)
            con.execute("UPDATE users SET photo_data=?,photo_mime=?,photo_name=? WHERE id=?",(data,mime,name,u["id"]))
            emp=linked_employee(con,u)
            if emp:
                con.execute("UPDATE employee_master SET photo_data=?,photo_mime=?,photo_path=?,updated_at=? WHERE id=?",
                            (data,mime,name,datetime.datetime.now().isoformat(timespec="seconds"),emp["id"]))
        con.execute("""UPDATE users SET full_name=?,profile_headline=?,profile_bio=?,official_mobile=?,official_email=?,department=?,branch=?,
                       profile_theme=?,appearance_mode=? WHERE id=?""",
                    (profile_name,headline,bio,official_mobile,official_email,department,branch,profile_theme,appearance_mode,u["id"]))
        emp=linked_employee(con,u)
        if emp:
            con.execute("""UPDATE employee_master SET full_name=?,mobile=?,email=?,department=?,updated_at=? WHERE id=?""",
                        (profile_name,official_mobile,official_email,department,datetime.datetime.now().isoformat(timespec="seconds"),emp["id"]))
        con.commit(); con.close()
        flash("Professional profile updated successfully.","success")
        return redirect(url_for("my_profile",photo_refresh=int(datetime.datetime.now().timestamp())))

    employee=linked_employee(con,u)
    manager=None
    if u["reporting_user_id"]:
        manager=con.execute("SELECT id,full_name,role FROM users WHERE id=?",(u["reporting_user_id"],)).fetchone()
    performance=None
    if u["role"]=="AM" and u["company_code"]:
        _,start_date,end_date,_=selected_month_bounds(None)
        rankings=am_month_rankings(con,start_date,end_date,u["company_code"])
        performance=next((r for r in rankings if int(r["id"])==int(u["id"])),None)
    today=datetime.date.today().isoformat()
    reminder_rows=[]
    if u["role"]=="AM":
        reminder_rows=con.execute("""SELECT id,lead_id,client_name,mobile,followup_date,status,interest_score
                                     FROM leads WHERE assigned_am=? AND COALESCE(deleted_at,'')=''
                                     AND COALESCE(followup_date,'')<>'' AND followup_date<=?
                                     ORDER BY followup_date ASC,COALESCE(interest_score,0) DESC LIMIT 25""",(u["id"],today)).fetchall()
    elif u["role"]=="GM":
        reminder_rows=con.execute("""SELECT id,lead_id,client_name,mobile,followup_date,status,interest_score
                                     FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')=''
                                     AND COALESCE(followup_date,'')<>'' AND followup_date<=?
                                     ORDER BY followup_date ASC,COALESCE(interest_score,0) DESC LIMIT 25""",(u["company_code"],today)).fetchall()
    elif u["role"]=="MD":
        reminder_rows=con.execute("""SELECT id,lead_id,client_name,mobile,followup_date,status,interest_score
                                     FROM leads WHERE COALESCE(deleted_at,'')='' AND COALESCE(followup_date,'')<>''
                                     AND followup_date<=? ORDER BY followup_date ASC,COALESCE(interest_score,0) DESC LIMIT 25""",(today,)).fetchall()
    con.close()
    return render_template("my_profile.html",u=current_user(),employee=employee,reporting_manager=manager,
                           performance=performance,reminder_rows=reminder_rows,today=today)

@app.route("/my-profile/remove-photo",methods=["POST"])
def remove_my_profile_photo():
    u=current_user()
    if not u:
        return redirect(url_for("login"))
    con=db()
    con.execute("UPDATE users SET photo_data=NULL,photo_mime='',photo_name='' WHERE id=?",(u["id"],))
    emp=linked_employee(con,u)
    if emp:
        con.execute("UPDATE employee_master SET photo_data=NULL,photo_mime='',photo_path='' WHERE id=?",(emp["id"],))
    con.commit(); con.close()
    flash("Profile photo removed.","success")
    return redirect(url_for("my_profile"))

@app.route("/logout")
def logout():
    log_activity("LOGOUT","Logout","SECURITY",details="User logged out")
    session.clear()
    return redirect(url_for("login"))


@app.route("/md/open-portal/<int:user_id>")
@require_roles("MD")
def md_open_employee_portal(user_id):
    """MD can view any employee's portal without their password."""
    u = current_user()
    con = db()
    target = con.execute("SELECT * FROM users WHERE id=? AND active=1", (user_id,)).fetchone()
    con.close()
    if not target:
        flash("Employee portal not found or inactive.", "error")
        return redirect(url_for("employees"))
    # Store MD's original session so they can return
    session["md_original_uid"] = u["id"]
    session["uid"] = target["id"]
    log_activity("MD_PORTAL_VIEW", f"MD opened portal as {target['full_name']}", "ADMIN",
                 details=f"MD {u['full_name']} opened portal as {target['full_name']} (id={target['id']})")
    return redirect(url_for("dashboard"))


@app.route("/md/return-to-md")
def md_return_to_portal():
    """Return MD back to their own session after viewing employee portal."""
    orig = session.get("md_original_uid")
    if orig:
        session["uid"] = orig
        session.pop("md_original_uid", None)
        flash("Returned to MD portal.", "success")
    return redirect(url_for("dashboard"))


@app.route("/central-status")
def central_status():
    u=current_user()
    if not u:
        return redirect(url_for("login"))
    con=db()
    visitors=con.execute("SELECT COUNT(*) c FROM visitors").fetchone()["c"]
    ams=con.execute("SELECT COUNT(*) c FROM users WHERE role='AM'").fetchone()["c"]
    leads=con.execute("SELECT COUNT(*) c FROM leads").fetchone()["c"]
    con.close()
    return jsonify({
        "status":"online",
        "mode":"CENTRAL_SHARED_SERVER",
        "database":os.path.basename(DB_PATH),
        "visitors":visitors,
        "assistant_managers":ams,
        "leads":leads,
        "logged_in_as":u["login_id"],
        "company":u["company_code"] or "BOTH"
    })


def selected_month_bounds(month_value=None):
    today=datetime.date.today()
    month_value=(month_value or "").strip()
    try:
        year,month=[int(x) for x in month_value.split("-")]
        start=datetime.date(year,month,1)
    except Exception:
        start=datetime.date(today.year,today.month,1)
    if start.month==12:
        end=datetime.date(start.year+1,1,1)-datetime.timedelta(days=1)
    else:
        end=datetime.date(start.year,start.month+1,1)-datetime.timedelta(days=1)
    return start.strftime("%Y-%m"),start.isoformat(),end.isoformat(),start.strftime("%B %Y")

def employee_month_rankings(con,start_date,end_date,company_code=None,designation=None):
    q="SELECT * FROM employee_master WHERE employment_status='Active'"
    params=[]
    if company_code:
        q+=" AND company_code=?"; params.append(company_code)
    if designation:
        q+=" AND designation=?"; params.append(designation)
    employees=con.execute(q+" ORDER BY company_code,full_name",params).fetchall()
    raw=[]
    for e in employees:
        perf=con.execute("""SELECT COALESCE(SUM(leads_assigned),0) leads_assigned,
                     COALESCE(SUM(target_amount),0) target_amount,
                     COALESCE(AVG(NULLIF(manager_rating,0)),0) avg_rating
                     FROM employee_performance
                     WHERE employee_id=? AND performance_date BETWEEN ? AND ?""",
                     (e["id"],start_date,end_date)).fetchone()
        cases=con.execute("""SELECT COUNT(*) enrollments,COALESCE(SUM(COALESCE(total_received,0)),0) revenue
                             FROM client_cases WHERE assigned_employee_id=? AND enrollment_date BETWEEN ? AND ?""",
                          (e["id"],start_date,end_date)).fetchone()
        raw.append({"employee":e,"leads_assigned":perf["leads_assigned"] or 0,
                    "target_amount":perf["target_amount"] or 0,"avg_rating":perf["avg_rating"] or 0,
                    "enrollments":cases["enrollments"] or 0,"revenue":cases["revenue"] or 0})
    max_enroll=max([r["enrollments"] for r in raw],default=0)
    max_revenue=max([r["revenue"] for r in raw],default=0)
    ranked=[]
    for r in raw:
        e=r["employee"]
        enroll_index=(r["enrollments"]/max_enroll*100) if max_enroll else 0
        revenue_index=(r["revenue"]/max_revenue*100) if max_revenue else 0
        target_achievement=(r["revenue"]/r["target_amount"]*100) if r["target_amount"] else 0
        score=(min(enroll_index,100)*0.40+min(revenue_index,100)*0.40+
               min(target_achievement,100)*0.10+min((r["avg_rating"] or 0)*20,100)*0.10)
        ranked.append({"id":e["id"],"full_name":e["full_name"],"company_code":e["company_code"],
                       "designation":e["designation"],"department":e["department"],"enrollments":r["enrollments"],
                       "revenue":r["revenue"],"score":round(score,1)})
    ranked.sort(key=lambda x:(x["score"],x["revenue"],x["enrollments"]),reverse=True)
    for i,r in enumerate(ranked,1): r["rank"]=i
    return ranked


def am_business_month_rankings(con,start_date,end_date,company_code=None):
    """v3.98: ranking is based ONLY on actual enrollment count."""
    q="""SELECT u.id,u.full_name,u.company_code,u.designation,u.photo_mime,u.active,u.employee_id
         FROM users u WHERE u.role='AM'"""
    params=[]
    if company_code:
        q+=" AND u.company_code=?"; params.append(company_code)
    q+=" ORDER BY u.company_code,u.full_name"
    ams=con.execute(q,params).fetchall()

    rows=[]
    for r in ams:
        enrollments=0
        if r["employee_id"]:
            c=con.execute("""SELECT COUNT(*) enrollments
                             FROM client_cases
                             WHERE assigned_employee_id=? AND enrollment_date BETWEEN ? AND ?""",
                          (r["employee_id"],start_date,end_date)).fetchone()
            enrollments=int(c["enrollments"] or 0)
        rows.append({
            "id":r["id"],"full_name":r["full_name"],"company_code":r["company_code"],
            "designation":r["designation"] or "Assistant Manager","photo_mime":r["photo_mime"],
            "active":int(r["active"] or 0),"employee_id":r["employee_id"],
            "allocated":0,"positive":0,"hot":0,"revenue":0.0,"conversion":0.0,
            "enrollments":enrollments
        })

    max_enroll=max([x["enrollments"] for x in rows],default=0)
    for x in rows:
        x["score"]=round((x["enrollments"]/max_enroll*100) if max_enroll else 0,1)

    rows.sort(key=lambda x:(-x["enrollments"],x["full_name"].lower()))
    dense_rank=0; previous=None
    for idx,x in enumerate(rows,1):
        if previous is None or x["enrollments"]!=previous:
            dense_rank=idx
            previous=x["enrollments"]
        x["rank"]=dense_rank
    return rows

def am_month_rankings(con,start_date,end_date,company_code=None):
    q="""SELECT u.id,u.full_name,u.company_code,
         COUNT(l.id) allocated,
         SUM(CASE WHEN COALESCE(l.interest_score,0)>=50 THEN 1 ELSE 0 END) positive,
         SUM(CASE WHEN COALESCE(l.interest_score,0)>=71 THEN 1 ELSE 0 END) hot,
         SUM(CASE WHEN l.status='Enrolled' THEN 1 ELSE 0 END) enrolled,
         COALESCE(AVG(COALESCE(l.interest_score,0)),0) avg_interest
         FROM users u LEFT JOIN leads l ON l.assigned_am=u.id
         AND COALESCE(NULLIF(l.assigned_at,''),l.imported_at,'') BETWEEN ? AND ?
         WHERE u.role='AM' AND u.active=1"""
    params=[start_date,end_date+"T23:59:59"]
    if company_code:
        q+=" AND u.company_code=?"; params.append(company_code)
    q+=" GROUP BY u.id,u.full_name,u.company_code"
    rows=con.execute(q,params).fetchall()
    ranked=[]
    for r in rows:
        allocated=int(r["allocated"] or 0); positive=int(r["positive"] or 0); hot=int(r["hot"] or 0); enrolled=int(r["enrolled"] or 0)
        positive_rate=(positive/allocated*100) if allocated else 0
        hot_rate=(hot/allocated*100) if allocated else 0
        enrollment_rate=(enrolled/allocated*100) if allocated else 0
        score=min(100,positive_rate*0.45+hot_rate*0.25+enrollment_rate*0.20+float(r["avg_interest"] or 0)*0.10)
        ranked.append({"id":r["id"],"full_name":r["full_name"],"company_code":r["company_code"],
                       "allocated":allocated,"positive":positive,"hot":hot,"enrolled":enrolled,
                       "avg_interest":round(float(r["avg_interest"] or 0)), "score":round(score,1)})
    ranked.sort(key=lambda x:(x["score"],x["positive"],x["hot"]),reverse=True)
    for i,r in enumerate(ranked,1): r["rank"]=i
    return ranked


def normalize_mobile(value):
    digits=re.sub(r"\D","",value or "")
    if len(digits)>10 and digits.startswith("91"):
        digits=digits[-10:]
    return digits

def next_identity_code(con,company,role):
    prefix=(company or "GAUR").upper()
    role_key={"GM":"GM","AM":"AM","RECEPTION":"REC","MD":"MD"}.get((role or "").upper(),"EMP")
    n=con.execute("SELECT COUNT(*) c FROM users WHERE company_code=? AND role=?",(company,role)).fetchone()["c"]+1
    code=f"{prefix}-{role_key}-{n:03d}"
    while con.execute("SELECT id FROM users WHERE identity_code=?",(code,)).fetchone():
        n+=1; code=f"{prefix}-{role_key}-{n:03d}"
    return code





def _notif_is_read(con,user_id,key):
    return bool(con.execute("SELECT id FROM notification_reads WHERE user_id=? AND notification_key=? ORDER BY id DESC LIMIT 1",
                            (user_id,key)).fetchone())



def _audit_txt(v,n=1200):
    if v is None: return ""
    try:
        if isinstance(v,(dict,list,tuple)): v=json.dumps(v,ensure_ascii=False,default=str)
    except Exception: pass
    return str(v)[:n]

def log_activity(action_code,action_label,category="GENERAL",entity_type="",entity_id="",entity_label="",details="",
                 old_value="",new_value="",severity="INFO",actor=None,company_code=None):
    con=None
    try:
        actor=actor if actor is not None else current_user()
        con=db()
        uid=actor["id"] if actor else None; uname=actor["full_name"] if actor else "Anonymous"; role=actor["role"] if actor else ""
        designation=(actor["designation"] if actor and "designation" in actor.keys() else "")
        company=company_code if company_code is not None else (actor["company_code"] if actor else "")
        ip=(request.headers.get("X-Forwarded-For") or request.remote_addr or "").split(",")[0].strip()[:80]
        con.execute("""INSERT INTO activity_logs(user_id,user_name,user_role,user_designation,company_code,category,action_code,
          action_label,entity_type,entity_id,entity_label,details,old_value,new_value,severity,ip_address,request_method,request_path,user_agent,created_at)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
          (uid,uname,role,designation,company or "",category,action_code,action_label,entity_type,_audit_txt(entity_id,120),
           _audit_txt(entity_label,240),_audit_txt(details),_audit_txt(old_value),_audit_txt(new_value),severity,ip,
           request.method[:12],request.path[:240],(request.headers.get("User-Agent") or "")[:400],datetime.datetime.now().isoformat(timespec="seconds")))
        con.commit()
    except Exception:
        try:
            if con: con.rollback()
        except Exception: pass
    finally:
        try:
            if con: con.close()
        except Exception: pass

def _audit_classify_request():
    p=request.path
    if p.startswith("/lead/"): return ("CLIENT","CLIENT_UPDATE","Client / Lead Updated")
    if p.startswith("/leads"): return ("LEAD","LEAD_ACTION","Lead Management Action")
    if p.startswith("/chat/broadcast"): return ("CHAT","BROADCAST_SENT","Management Broadcast Sent")
    if p.startswith("/chat/send"): return ("CHAT","CHAT_MESSAGE_SENT","Chat Message Sent")
    if p.startswith("/team-identit") or p.startswith("/ams"): return ("EMPLOYEE","EMPLOYEE_ADMIN","Employee / Portal Identity Action")
    if p.startswith("/employee"): return ("EMPLOYEE","EMPLOYEE_UPDATE","Employee Information Updated")
    if p.startswith("/reception"): return ("ADMIN","RECEPTION_UPDATE","Reception / Visitor Action")
    if p.startswith("/notifications"): return ("ADMIN","NOTIFICATION_ACTION","Notification Action")
    return ("ADMIN","PORTAL_CHANGE","Portal Data Change")

@app.before_request
def audit_before_request():
    try:
        if request.method=="POST" and request.path.startswith("/lead/"):
            x=request.path.rstrip("/").split("/")[-1]
            if x.isdigit():
                con=db(); r=con.execute("SELECT id,lead_id,status,interest_score,next_followup,remarks,company_code FROM leads WHERE id=?",(int(x),)).fetchone(); con.close()
                if r: request.environ["GAUR_AUDIT_OLD"]=dict(r)
    except Exception: pass

@app.after_request
def audit_after_request(response):
    try:
        if request.method=="POST" and response.status_code<400 and request.path not in ("/","/logout"):
            actor=current_user()
            if actor:
                cat,code,label=_audit_classify_request(); entity_type="";entity_id="";entity_label="";oldv="";newv=""
                if request.path.startswith("/lead/"):
                    entity_id=request.path.rstrip("/").split("/")[-1];entity_type="Lead";old=request.environ.get("GAUR_AUDIT_OLD") or {}
                    if old:
                        entity_label=old.get("lead_id","");oldv={"status":old.get("status"),"interest":old.get("interest_score"),"followup":old.get("next_followup")}
                    try:
                        con=db();fresh=con.execute("SELECT lead_id,status,interest_score,next_followup FROM leads WHERE id=?",(int(entity_id),)).fetchone();con.close()
                        if fresh: entity_label=fresh["lead_id"];newv={"status":fresh["status"],"interest":fresh["interest_score"],"followup":fresh["next_followup"]}
                    except Exception: pass
                safe={}
                for k,v in request.form.items():
                    if "password" in k.lower() or "otp" in k.lower(): safe[k]="***"
                    elif k.lower() not in ("message","profile_bio"): safe[k]=str(v)[:180]
                log_activity(code,label,cat,entity_type,entity_id,entity_label,safe,oldv,newv,"WARNING" if request.path.startswith("/chat/broadcast") else "INFO",actor)
    except Exception: pass
    return response

def linked_employee(con,user_row):
    if not user_row: return None
    if user_row["employee_id"]:
        e=con.execute("SELECT * FROM employee_master WHERE id=?",(user_row["employee_id"],)).fetchone()
        if e: return e
    return con.execute("SELECT * FROM employee_master WHERE portal_user_id=?",(user_row["id"],)).fetchone()



@app.route("/api/dashboard/reminders")
def dashboard_reminders_api():
    u=current_user()
    if not u:
        return jsonify({"items":[],"count":0}),401
    con=db()
    today=datetime.date.today().isoformat()
    params=[]
    where=["COALESCE(l.deleted_at,'')=''","COALESCE(l.followup_date,'')<>''","l.followup_date<=?"]
    params.append(today)
    if u["role"]=="AM":
        where.append("l.assigned_am=?"); params.append(u["id"])
    elif u["role"]!="MD":
        where.append("l.company_code=?"); params.append(u["company_code"])
    rows=con.execute("""SELECT l.id,l.lead_id,l.client_name,l.mobile,l.followup_date,l.status,l.interest_score,
                               au.full_name am_name
                        FROM leads l LEFT JOIN users au ON au.id=l.assigned_am
                        WHERE """+" AND ".join(where)+"""
                        ORDER BY l.followup_date ASC,COALESCE(l.interest_score,0) DESC,l.id DESC LIMIT 12""",params).fetchall()
    out=[]
    for r in rows:
        out.append({
          "id":r["id"],"lead_id":r["lead_id"],"client_name":r["client_name"] or r["lead_id"],
          "mobile":r["mobile"] or "","followup_date":r["followup_date"],"status":r["status"] or "",
          "interest":int(r["interest_score"] or 0),"am_name":r["am_name"] or "",
          "url":url_for("lead_profile",lead_db_id=r["id"]),
          "overdue":bool(r["followup_date"]<today)
        })
    con.close()
    return jsonify({"items":out,"count":len(out),"chat_unread":0})


@app.route("/dashboard")
@require_roles("MD","GM","RECEPTION","AM")
def dashboard():
    # v3.98 emergency fix: live DB may pre-date Enrollment Recycle Bin columns.
    ensure_enrollment_recycle_schema()
    u=current_user(); con=db()
    _active_case_sql=active_client_case_predicate(con)
    selected_month,month_start,month_end,month_label=selected_month_bounds(request.args.get("month"))
    month_end_ts=month_end+"T23:59:59"

    # v3.98 TODAY REPORT: strict day-wise calculation.
    today_date=datetime.date.today().isoformat()
    today_report={}
    for _co in ("SCIC","WWIC"):
        try:
            _enr=con.execute("SELECT COUNT(*) c FROM client_cases WHERE company_code=? AND enrollment_date=? AND "+_active_case_sql,(_co,today_date)).fetchone()["c"]
        except Exception:
            try: con.rollback()
            except Exception: pass
            _enr=0
        try:
            _r=con.execute("""SELECT
                COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(first_payment_datetime,''),1,10)=?
                     THEN COALESCE(first_cash,0)+COALESCE(first_rbl,0)+COALESCE(first_yes_bank,0)+COALESCE(first_au_bank,0)+COALESCE(first_indusind,0) ELSE 0 END),0) f,
                COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(second_payment_datetime,''),1,10)=?
                     THEN COALESCE(second_cash,0)+COALESCE(second_rbl,0)+COALESCE(second_yes_bank,0)+COALESCE(second_au_bank,0)+COALESCE(second_indusind,0) ELSE 0 END),0) s,
                COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(other_payment_datetime,''),1,10)=?
                     THEN COALESCE(other_payment_received,0) ELSE 0 END),0) o
                FROM client_cases WHERE company_code=? AND """+_active_case_sql,
                (today_date,today_date,today_date,_co)).fetchone()
            _rev=float((_r["f"] or 0)+(_r["s"] or 0)+(_r["o"] or 0))
        except Exception:
            try: con.rollback()
            except Exception: pass
            _rev=0.0
        today_report[_co]={"enrollments":int(_enr or 0),"revenue":_rev}

    # v4.1 MONTHLY REPORT: company-wise enrollments & revenue for selected month
    monthly_report={}
    for _co in ("SCIC","WWIC"):
        try:
            _menr=con.execute(
                "SELECT COUNT(*) c FROM client_cases WHERE company_code=? AND enrollment_date BETWEEN ? AND ? AND "+_active_case_sql,
                (_co,month_start,month_end)
            ).fetchone()["c"]
        except Exception:
            try: con.rollback()
            except Exception: pass
            _menr=0
        try:
            _mr=con.execute("""SELECT
                COALESCE(SUM(COALESCE(first_cash,0)+COALESCE(first_rbl,0)+COALESCE(first_yes_bank,0)+COALESCE(first_au_bank,0)+COALESCE(first_indusind,0)+
                             COALESCE(second_cash,0)+COALESCE(second_rbl,0)+COALESCE(second_yes_bank,0)+COALESCE(second_au_bank,0)+COALESCE(second_indusind,0)+
                             COALESCE(other_payment_received,0)),0) total
                FROM client_cases WHERE company_code=? AND enrollment_date BETWEEN ? AND ? AND """+_active_case_sql,
                (_co,month_start,month_end)).fetchone()
            _mrev=float(_mr["total"] or 0)
        except Exception:
            try: con.rollback()
            except Exception: pass
            _mrev=0.0
        monthly_report[_co]={"enrollments":int(_menr or 0),"revenue":_mrev}

    where=["COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ?","COALESCE(deleted_at,'')=''"]; params=[month_start,month_end_ts]
    if u["role"]!="MD":
        where.append("company_code=?"); params.append(u["company_code"])
    if u["role"]=="AM":
        where.append("assigned_am=?"); params.append(u["id"])
    lead_where=(" WHERE "+" AND ".join(where)) if where else ""
    total_leads=con.execute("SELECT COUNT(*) c FROM leads"+lead_where,params).fetchone()["c"]
    allocated=con.execute("SELECT COUNT(*) c FROM leads"+lead_where+(" AND " if lead_where else " WHERE ")+"assigned_am IS NOT NULL",params).fetchone()["c"]
    interested=con.execute("SELECT COUNT(*) c FROM leads"+lead_where+(" AND " if lead_where else " WHERE ")+"status='Interested'",params).fetchone()["c"]
    unallocated=con.execute("SELECT COUNT(*) c FROM leads"+lead_where+(" AND " if lead_where else " WHERE ")+"assigned_am IS NULL",params).fetchone()["c"]
    cross_dups=con.execute("SELECT COUNT(*) c FROM leads"+lead_where+(" AND " if lead_where else " WHERE ")+"duplicate_type='cross'",params).fetchone()["c"]
    hot_count=con.execute("SELECT COUNT(*) c FROM leads"+lead_where+(" AND " if lead_where else " WHERE ")+"COALESCE(interest_score,0)>=71",params).fetchone()["c"]
    warm_count=con.execute("SELECT COUNT(*) c FROM leads"+lead_where+(" AND " if lead_where else " WHERE ")+"COALESCE(interest_score,0) BETWEEN 50 AND 70",params).fetchone()["c"]
    cold_count=con.execute("SELECT COUNT(*) c FROM leads"+lead_where+(" AND " if lead_where else " WHERE ")+"COALESCE(interest_score,0)<50",params).fetchone()["c"]
    avg_interest=round(float(con.execute("SELECT COALESCE(AVG(COALESCE(interest_score,0)),0) a FROM leads"+lead_where,params).fetchone()["a"] or 0))

    visitors_where=" WHERE visit_date BETWEEN ? AND ?"
    visitors_params=[month_start,month_end]
    if u["role"]!="MD":
        visitors_where+=" AND company_code=?"; visitors_params.append(u["company_code"])
    visitors=con.execute("SELECT COUNT(*) c FROM visitors"+visitors_where,visitors_params).fetchone()["c"]
    recent=con.execute("SELECT * FROM leads"+lead_where+" ORDER BY id DESC LIMIT 8",params).fetchall()
    case_where=" WHERE enrollment_date BETWEEN ? AND ? AND "+_active_case_sql
    case_params=[month_start,month_end]
    if u["role"]!="MD":
        case_where+=" AND company_code=?"; case_params.append(u["company_code"])
    cases=con.execute("SELECT COUNT(*) c FROM client_cases"+case_where,case_params).fetchone()["c"]
    revenue=con.execute("SELECT COALESCE(SUM(COALESCE(total_received,0)),0) s FROM client_cases"+case_where,case_params).fetchone()["s"] if u["role"]=="MD" else 0
    company_stats=[]
    if u["role"]=="MD":
        for code in ("SCIC","WWIC"):
            total_c=con.execute("SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ?",(code,month_start,month_end_ts)).fetchone()["c"]
            allocated_c=con.execute("SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ? AND assigned_am IS NOT NULL",(code,month_start,month_end_ts)).fetchone()["c"]
            unallocated_c=con.execute("SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ? AND assigned_am IS NULL",(code,month_start,month_end_ts)).fetchone()["c"]
            hot_c=con.execute("SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ? AND COALESCE(interest_score,0)>=71",(code,month_start,month_end_ts)).fetchone()["c"]
            warm_c=con.execute("SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ? AND COALESCE(interest_score,0) BETWEEN 50 AND 70",(code,month_start,month_end_ts)).fetchone()["c"]
            cold_c=con.execute("SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ? AND COALESCE(interest_score,0)<50",(code,month_start,month_end_ts)).fetchone()["c"]
            positive_c=hot_c+warm_c
            enrolled_c=con.execute("SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ? AND status='Enrolled'",(code,month_start,month_end_ts)).fetchone()["c"]
            avg_c=round(float(con.execute("SELECT COALESCE(AVG(COALESCE(interest_score,0)),0) a FROM leads WHERE company_code=? AND COALESCE(deleted_at,'')='' AND COALESCE(NULLIF(assigned_at,''),imported_at,'') BETWEEN ? AND ?",(code,month_start,month_end_ts)).fetchone()["a"] or 0))
            rev_c=con.execute("SELECT COALESCE(SUM(COALESCE(total_received,0)),0) s FROM client_cases WHERE company_code=? AND enrollment_date BETWEEN ? AND ? AND "+_active_case_sql,(code,month_start,month_end)).fetchone()["s"]
            denom=max(total_c,1)
            hot_pct=round((hot_c/denom)*100,2)
            positive_pct=round((positive_c/denom)*100,2)
            positive_rate=round((positive_c/denom)*100)
            company_stats.append({
                "code":code,"total":total_c,"allocated":allocated_c,"unallocated":unallocated_c,
                "hot":hot_c,"warm":warm_c,"cold":cold_c,"positive":positive_c,
                "enrolled":enrolled_c,"avg_interest":avg_c,"revenue":rev_c,
                "hot_pct":hot_pct,"positive_pct":positive_pct,"positive_rate":positive_rate
            })
    allocation_summary=[]
    if u["role"] in ("MD","GM"):
        aq="""SELECT usr.id,usr.login_id,usr.full_name,usr.company_code,usr.active,
                     COUNT(l.id) allocated_count,
                     SUM(CASE WHEN l.status='New' THEN 1 ELSE 0 END) new_count,
                     SUM(CASE WHEN l.status='Called' THEN 1 ELSE 0 END) called_count,
                     SUM(CASE WHEN l.status='Interested' THEN 1 ELSE 0 END) interested_count,
                     SUM(CASE WHEN l.status='Follow-up' THEN 1 ELSE 0 END) followup_count,
                     SUM(CASE WHEN l.status='Enrolled' THEN 1 ELSE 0 END) enrolled_count,
                     SUM(CASE WHEN COALESCE(l.interest_score,0)>=71 THEN 1 ELSE 0 END) hot_count,
                     SUM(CASE WHEN COALESCE(l.interest_score,0) BETWEEN 50 AND 70 THEN 1 ELSE 0 END) warm_count,
                     SUM(CASE WHEN COALESCE(l.interest_score,0)>=50 THEN 1 ELSE 0 END) positive_count
              FROM users usr LEFT JOIN leads l ON l.assigned_am=usr.id
              AND COALESCE(NULLIF(l.assigned_at,''),l.imported_at,'') BETWEEN ? AND ?
              WHERE usr.role='AM'"""
        ap=[month_start,month_end_ts]
        if u["role"]=="GM": aq+=" AND usr.company_code=?"; ap.append(u["company_code"])
        aq+=" GROUP BY usr.id,usr.login_id,usr.full_name,usr.company_code,usr.active ORDER BY usr.company_code,allocated_count DESC,usr.full_name"
        allocation_rows=con.execute(aq,ap).fetchall()
        allocation_summary=[]
        for a in allocation_rows:
            d=dict(a)
            allocated_n=int(d.get("allocated_count") or 0)
            positive_n=int(d.get("positive_count") or 0)
            d["positive_pct"]=round((positive_n/allocated_n)*100) if allocated_n else 0
            allocation_summary.append(d)
    # v3.30: business-focused AM leaderboard is visible on every portal dashboard.
    if u["role"]=="MD":
        month_top_performers=am_business_month_rankings(con,month_start,month_end,None)
        scic_am_rankings=am_business_month_rankings(con,month_start,month_end,"SCIC")
        wwic_am_rankings=am_business_month_rankings(con,month_start,month_end,"WWIC")
        company_am_rankings=month_top_performers
    else:
        company_am_rankings=am_business_month_rankings(con,month_start,month_end,u["company_code"])
        month_top_performers=company_am_rankings
        scic_am_rankings=[]
        wwic_am_rankings=[]
    my_monthly_performance=None; best_am=None
    if u["role"]=="AM":
        best_am=company_am_rankings[0] if company_am_rankings else None
        my_monthly_performance=next((x for x in company_am_rankings if int(x["id"])==int(u["id"])),None)
    dashboard_employee=linked_employee(con,u)
    dashboard_rank=None
    dashboard_score=None
    if u["role"]=="AM" and my_monthly_performance:
        dashboard_rank=my_monthly_performance.get("rank")
        dashboard_score=my_monthly_performance.get("score")
    elif dashboard_employee and u["role"] in ("GM","RECEPTION"):
        try:
            staff_rankings=employee_month_rankings(con,month_start,month_end,u["company_code"])
            for idx,item in enumerate(staff_rankings,1):
                if int(item.get("id") or 0)==int(dashboard_employee["id"]):
                    dashboard_rank=idx
                    dashboard_score=round(float(item.get("score") or 0),1)
                    break
        except Exception:
            pass
    con.close()
    return render_template("dashboard.html",u=u,total_leads=total_leads,allocated=allocated,interested=interested,
                           unallocated=unallocated,cross_dups=cross_dups,visitors=visitors,recent=recent,cases=cases,revenue=revenue,allocation_summary=allocation_summary,
                           hot_count=hot_count,warm_count=warm_count,cold_count=cold_count,avg_interest=avg_interest,
                           company_stats=company_stats,selected_month=selected_month,month_label=month_label,
                           month_top_performers=month_top_performers,company_am_rankings=company_am_rankings,
                           scic_am_rankings=scic_am_rankings,wwic_am_rankings=wwic_am_rankings,
                           my_monthly_performance=my_monthly_performance,best_am=best_am,
                           dashboard_employee=dashboard_employee,dashboard_rank=dashboard_rank,dashboard_score=dashboard_score,
                           today_report=today_report,today_date=today_date,monthly_report=monthly_report)


# ── Temporary debug endpoint ──────────────────────────────────────────────
@app.route("/debug/perf-check")
@require_roles("MD")
def debug_perf_check():
    con = db()
    # Check AM users and their employee_id links
    users = con.execute("SELECT id, login_id, full_name, employee_id FROM users WHERE role='AM' ORDER BY full_name").fetchall()
    # Check recent cases
    cases = con.execute("""SELECT c.case_id, c.lead_db_id, c.assigned_employee_id, c.enrollment_date,
                           l.assigned_am, l.client_name
                           FROM client_cases c LEFT JOIN leads l ON l.id=c.lead_db_id
                           WHERE c.enrollment_date >= '2026-08-01' LIMIT 20""").fetchall()
    out = "<h2>AM Users</h2><table border=1>"
    for u in users:
        out += f"<tr><td>{u['id']}</td><td>{u['login_id']}</td><td>{u['full_name']}</td><td>employee_id: {u['employee_id']}</td></tr>"
    out += "</table><h2>Recent Cases (Aug 2026)</h2><table border=1>"
    out += "<tr><th>case_id</th><th>lead_db_id</th><th>assigned_employee_id</th><th>enrollment_date</th><th>assigned_am (from lead)</th><th>client_name</th></tr>"
    for c in cases:
        out += f"<tr><td>{c['case_id']}</td><td>{c['lead_db_id']}</td><td>{c['assigned_employee_id']}</td><td>{c['enrollment_date']}</td><td>{c['assigned_am']}</td><td>{c['client_name']}</td></tr>"
    out += "</table>"
    con.close()
    return out
# ── End debug ──────────────────────────────────────────────────────────────

# ── Facebook Lead Ads Webhook ─────────────────────────────────────────────
FB_VERIFY_TOKEN = os.environ.get("FB_VERIFY_TOKEN", "gaur-crm-fb-verify-2026")

# Page ID → Company Code mapping (set via environment variables)
FB_PAGE_MAP = {}
_wwic_page = os.environ.get("FB_PAGE_ID_XYZ", "")
_scic_page = os.environ.get("FB_PAGE_ID_ABC", "")
if _wwic_page: FB_PAGE_MAP[_wwic_page] = "WWIC"
if _scic_page: FB_PAGE_MAP[_scic_page] = "SCIC"

@app.route("/webhooks/facebook", methods=["GET", "POST"])
def facebook_webhook():
    # Webhook verification (GET)
    if request.method == "GET":
        mode = request.args.get("hub.mode")
        token = request.args.get("hub.verify_token")
        challenge = request.args.get("hub.challenge")
        if mode == "subscribe" and token == FB_VERIFY_TOKEN:
            return challenge, 200
        return "Forbidden", 403

    # Lead event (POST)
    data = request.get_json(silent=True) or {}
    try:
        for entry in data.get("entry", []):
            page_id = str(entry.get("id", ""))
            company_code = FB_PAGE_MAP.get(page_id, "WWIC")  # default XYZ
            for change in entry.get("changes", []):
                if change.get("field") != "leadgen":
                    continue
                value = change.get("value", {})
                lead_id_fb = str(value.get("leadgen_id", ""))
                form_id = str(value.get("form_id", ""))
                if not lead_id_fb:
                    continue
                # Fetch lead details from Facebook Graph API
                access_token = os.environ.get("FB_PAGE_ACCESS_TOKEN_" + company_code, "")
                if not access_token:
                    access_token = os.environ.get("FB_PAGE_ACCESS_TOKEN", "")
                lead_data = _fetch_fb_lead(lead_id_fb, access_token)
                if lead_data:
                    _save_fb_lead(lead_data, company_code, lead_id_fb)
    except Exception as e:
        app.logger.error(f"Facebook webhook error: {e}")
    return jsonify({"status": "ok"}), 200

def _fetch_fb_lead(lead_id, access_token):
    """Fetch lead field data from Facebook Graph API."""
    if not access_token:
        return None
    try:
        import urllib.request
        url = f"https://graph.facebook.com/v19.0/{lead_id}?access_token={access_token}"
        with urllib.request.urlopen(url, timeout=10) as r:
            data = json.loads(r.read())
        fields = {f["name"]: f["values"][0] if f.get("values") else "" for f in data.get("field_data", [])}
        return {
            "full_name": fields.get("full_name") or fields.get("name") or "",
            "mobile": fields.get("phone_number") or fields.get("mobile") or fields.get("phone") or "",
            "email": fields.get("email") or "",
            "city": fields.get("city") or "",
            "country": fields.get("country") or "India",
            "visa_type": fields.get("visa_type") or fields.get("interested_country") or "",
            "source": "Facebook",
        }
    except Exception as e:
        app.logger.error(f"FB lead fetch error: {e}")
        return None

def _save_fb_lead(lead, company_code, fb_lead_id):
    """Save Facebook lead into CRM leads table."""
    mobile = re.sub(r"[^\d+]", "", str(lead.get("mobile", "")))
    if not mobile or not lead.get("full_name"):
        return
    con = db()
    try:
        # Duplicate check
        existing = con.execute("SELECT id FROM leads WHERE mobile=? AND company_code=?",
                               (mobile, company_code)).fetchone()
        if existing:
            con.close()
            return
        now = datetime.datetime.now().isoformat(timespec="seconds")
        lead_id = f"{company_code}-FB-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}-{fb_lead_id[-4:]}"
        batch = f"{company_code}-FB-{datetime.date.today().strftime('%Y%m%d')}"
        con.execute("""INSERT INTO leads
            (lead_id, company_code, client_name, mobile, email, city, country, visa_type,
             source, status, imported_by, imported_at, upload_batch)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (lead_id, company_code, lead["full_name"], mobile,
             lead.get("email",""), lead.get("city",""), lead.get("country","India"),
             lead.get("visa_type",""), "Facebook", "New", "FB-WEBHOOK", now, batch))
        con.commit()
        app.logger.info(f"FB lead saved: {lead['full_name']} / {mobile} / {company_code}")
    except Exception as e:
        app.logger.error(f"FB lead save error: {e}")
        try: con.rollback()
        except: pass
    finally:
        con.close()

@app.route("/webhooks/facebook/status")
@require_roles("MD", "GM")
def facebook_webhook_status():
    """Check Facebook webhook configuration status."""
    wwic_page = os.environ.get("FB_PAGE_ID_XYZ", "NOT SET")
    scic_page = os.environ.get("FB_PAGE_ID_ABC", "NOT SET")
    token_wwic = "SET" if os.environ.get("FB_PAGE_ACCESS_TOKEN_XYZ") else "NOT SET"
    token_scic = "SET" if os.environ.get("FB_PAGE_ACCESS_TOKEN_ABC") else "NOT SET"
    verify = FB_VERIFY_TOKEN
    webhook_url = request.host_url.rstrip("/") + "/webhooks/facebook"
    return jsonify({
        "webhook_url": webhook_url,
        "verify_token": verify,
        "wwic_page_id": wwic_page,
        "scic_page_id": scic_page,
        "wwic_access_token": token_wwic,
        "scic_access_token": token_scic,
    })
# ── End Facebook Webhook ──────────────────────────────────────────────────

# ── One-time setup: create meta operator user ─────────────────────────────
@app.route("/setup/create-meta-operator")
def create_meta_operator():
    con = db()
    existing = con.execute("SELECT id FROM users WHERE login_id='manjul.nagar'").fetchone()
    if existing:
        # Update designation if needed
        con.execute("UPDATE users SET designation='Meta Lead Operator' WHERE login_id='manjul.nagar'")
        con.commit()
        con.close()
        return "User manjul.nagar already exists (designation updated).", 200
    pw = generate_password_hash("socialtheka")
    con.execute("INSERT INTO users (login_id, password_hash, full_name, role, company_code, active, designation) VALUES (?,?,?,?,?,1,?)",
                ("manjul.nagar", pw, "Manjul Nagar", "MD", "SCIC", "Meta Lead Operator"))
    con.commit()
    con.close()
    return "Meta Operator created. Login: manjul.nagar / socialtheka", 200
# ── End setup ─────────────────────────────────────────────────────────────

@app.route("/api/notifications")
def notifications_api():
    u = current_user()
    if not u: return jsonify({"ok": False}), 401
    return jsonify({"ok": True, "notifications": [], "unread": 0})

@app.route("/notifications")
def notification_center():
    u = current_user()
    if not u: return redirect(url_for("login"))
    return redirect(url_for("dashboard"))

@app.route("/notifications/mark-all", methods=["POST"])
def notifications_mark_all():
    u = current_user()
    if not u: return jsonify({"ok": False}), 401
    return jsonify({"ok": True})

@app.route("/notifications/open", methods=["POST"])
def notification_open():
    u = current_user()
    if not u: return jsonify({"ok": False}), 401
    return jsonify({"ok": True})
# ── End Notification stubs ─────────────────────────────────────────────────













def _management_am_report_rows(con,u,report_date,company_filter=""):
    """AM-wise daily + month-to-date operational report."""
    try:
        d=datetime.date.fromisoformat(report_date)
    except Exception:
        d=datetime.date.today()
        report_date=d.isoformat()

    month_start=d.replace(day=1).isoformat()
    month_end=report_date
    day_start=report_date+"T00:00:00"
    day_end=report_date+"T23:59:59"
    month_start_ts=month_start+"T00:00:00"
    month_end_ts=month_end+"T23:59:59"

    q="SELECT id,full_name,company_code,designation,employee_id,active FROM users WHERE role='AM'"
    params=[]
    if u["role"]=="GM":
        q+=" AND company_code=?";params.append(u["company_code"])
    elif company_filter in ("SCIC","WWIC"):
        q+=" AND company_code=?";params.append(company_filter)
    q+=" ORDER BY company_code,full_name"
    ams=con.execute(q,params).fetchall()

    rows=[]
    for am in ams:
        # Allocations are based on lead assigned_at.
        today_alloc=con.execute("""SELECT COUNT(*) c FROM leads
            WHERE assigned_am=? AND COALESCE(deleted_at,'')='' AND COALESCE(assigned_at,'') BETWEEN ? AND ?""",
            (am["id"],day_start,day_end)).fetchone()["c"]
        mtd_alloc=con.execute("""SELECT COUNT(*) c FROM leads
            WHERE assigned_am=? AND COALESCE(deleted_at,'')='' AND COALESCE(assigned_at,'') BETWEEN ? AND ?""",
            (am["id"],month_start_ts,month_end_ts)).fetchone()["c"]

        # Interested = current interest score 50%+ among leads allocated in the respective period.
        today_int=con.execute("""SELECT COUNT(*) c FROM leads
            WHERE assigned_am=? AND COALESCE(deleted_at,'')='' AND COALESCE(assigned_at,'') BETWEEN ? AND ?
              AND COALESCE(interest_score,0)>=50""",(am["id"],day_start,day_end)).fetchone()["c"]
        mtd_int=con.execute("""SELECT COUNT(*) c FROM leads
            WHERE assigned_am=? AND COALESCE(deleted_at,'')='' AND COALESCE(assigned_at,'') BETWEEN ? AND ?
              AND COALESCE(interest_score,0)>=50""",(am["id"],month_start_ts,month_end_ts)).fetchone()["c"]

        today_enr=mtd_enr=0
        today_rev=mtd_rev=0.0
        emp_id=am["employee_id"]
        if emp_id:
            try:
                today_enr=con.execute("""SELECT COUNT(*) c FROM client_cases
                    WHERE assigned_employee_id=? AND enrollment_date=? AND COALESCE(deleted_at,'')=''""",
                    (emp_id,report_date)).fetchone()["c"]
                mtd_enr=con.execute("""SELECT COUNT(*) c FROM client_cases
                    WHERE assigned_employee_id=? AND enrollment_date BETWEEN ? AND ? AND COALESCE(deleted_at,'')=''""",
                    (emp_id,month_start,month_end)).fetchone()["c"]
            except Exception:
                # Backward compatibility if deleted_at is not present on a live DB.
                try: con.rollback()
                except Exception: pass
                today_enr=con.execute("SELECT COUNT(*) c FROM client_cases WHERE assigned_employee_id=? AND enrollment_date=?",
                                      (emp_id,report_date)).fetchone()["c"]
                mtd_enr=con.execute("SELECT COUNT(*) c FROM client_cases WHERE assigned_employee_id=? AND enrollment_date BETWEEN ? AND ?",
                                    (emp_id,month_start,month_end)).fetchone()["c"]

            # Revenue follows actual payment received timestamps, not enrollment date.
            r=con.execute("""SELECT
              COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(first_payment_datetime,''),1,10)=?
                   THEN COALESCE(first_cash,0)+COALESCE(first_rbl,0)+COALESCE(first_yes_bank,0)+COALESCE(first_au_bank,0)+COALESCE(first_indusind,0) ELSE 0 END),0) f,
              COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(second_payment_datetime,''),1,10)=?
                   THEN COALESCE(second_cash,0)+COALESCE(second_rbl,0)+COALESCE(second_yes_bank,0)+COALESCE(second_au_bank,0)+COALESCE(second_indusind,0) ELSE 0 END),0) s,
              COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(other_payment_datetime,''),1,10)=?
                   THEN COALESCE(other_payment_received,0) ELSE 0 END),0) o
              FROM client_cases WHERE assigned_employee_id=?""",
              (report_date,report_date,report_date,emp_id)).fetchone()
            today_rev=float((r["f"] or 0)+(r["s"] or 0)+(r["o"] or 0))

            r=con.execute("""SELECT
              COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(first_payment_datetime,''),1,10) BETWEEN ? AND ?
                   THEN COALESCE(first_cash,0)+COALESCE(first_rbl,0)+COALESCE(first_yes_bank,0)+COALESCE(first_au_bank,0)+COALESCE(first_indusind,0) ELSE 0 END),0) f,
              COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(second_payment_datetime,''),1,10) BETWEEN ? AND ?
                   THEN COALESCE(second_cash,0)+COALESCE(second_rbl,0)+COALESCE(second_yes_bank,0)+COALESCE(second_au_bank,0)+COALESCE(second_indusind,0) ELSE 0 END),0) s,
              COALESCE(SUM(CASE WHEN SUBSTR(COALESCE(other_payment_datetime,''),1,10) BETWEEN ? AND ?
                   THEN COALESCE(other_payment_received,0) ELSE 0 END),0) o
              FROM client_cases WHERE assigned_employee_id=?""",
              (month_start,month_end,month_start,month_end,month_start,month_end,emp_id)).fetchone()
            mtd_rev=float((r["f"] or 0)+(r["s"] or 0)+(r["o"] or 0))

        conv=round((float(mtd_enr)/float(mtd_alloc)*100),1) if mtd_alloc else 0.0
        rows.append({
            "am_id":am["id"],"am_name":am["full_name"],"company_code":am["company_code"],
            "today_allocated":int(today_alloc or 0),"today_interested":int(today_int or 0),
            "today_enrolled":int(today_enr or 0),"today_revenue":today_rev,
            "mtd_allocated":int(mtd_alloc or 0),"mtd_interested":int(mtd_int or 0),
            "mtd_enrolled":int(mtd_enr or 0),"mtd_revenue":mtd_rev,"conversion":conv
        })

    rows.sort(key=lambda x:(x["company_code"],-x["mtd_enrolled"],-x["mtd_revenue"],x["am_name"].lower()))
    return rows,month_start,month_end


@app.route("/management-reporting")
@require_roles("MD","GM")
def management_reporting():
    u=current_user();con=db()
    report_date=(request.args.get("date") or datetime.date.today().isoformat()).strip()
    company=(request.args.get("company") or "").strip().upper()
    if u["role"]=="GM": company=u["company_code"]
    if company not in ("","SCIC","WWIC"): company=""

    rows,month_start,month_end=_management_am_report_rows(con,u,report_date,company)
    con.close()

    totals={
      "today_allocated":sum(x["today_allocated"] for x in rows),
      "today_interested":sum(x["today_interested"] for x in rows),
      "today_enrolled":sum(x["today_enrolled"] for x in rows),
      "today_revenue":sum(x["today_revenue"] for x in rows),
      "mtd_allocated":sum(x["mtd_allocated"] for x in rows),
      "mtd_interested":sum(x["mtd_interested"] for x in rows),
      "mtd_enrolled":sum(x["mtd_enrolled"] for x in rows),
      "mtd_revenue":sum(x["mtd_revenue"] for x in rows)
    }

    lines=[
      "IMS • AM WISE MANAGEMENT REPORT",
      "Date: "+report_date,
      "Scope: "+(("ABC" if company=="SCIC" else "XYZ") if company else "Both Companies"),
      "",
      "TODAY:"
    ]
    for x in rows:
        lines.append(f"{x['company_code']} • {x['am_name']}: Alloc {x['today_allocated']} | Interested {x['today_interested']} | Enrolled {x['today_enrolled']} | Revenue ₹{x['today_revenue']:,.0f}")
    lines.extend([
      "",
      f"TODAY TOTAL: Alloc {totals['today_allocated']} | Interested {totals['today_interested']} | Enrolled {totals['today_enrolled']} | Revenue ₹{totals['today_revenue']:,.0f}",
      "",
      f"MONTH TILL DATE ({month_start} to {month_end}):"
    ])
    for x in rows:
        lines.append(f"{x['company_code']} • {x['am_name']}: Alloc {x['mtd_allocated']} | Interested {x['mtd_interested']} | Enrolled {x['mtd_enrolled']} | Revenue ₹{x['mtd_revenue']:,.0f} | Conv {x['conversion']:.1f}%")
    lines.append("")
    lines.append(f"MTD TOTAL: Alloc {totals['mtd_allocated']} | Interested {totals['mtd_interested']} | Enrolled {totals['mtd_enrolled']} | Revenue ₹{totals['mtd_revenue']:,.0f}")
    wa_text="\n".join(lines)

    return render_template("management_reporting.html",u=u,rows=rows,totals=totals,
                           report_date=report_date,company=company,month_start=month_start,
                           month_end=month_end,wa_text=wa_text)


@app.route("/management-reporting/export.xlsx")
@require_roles("MD","GM")
def management_reporting_export():
    u=current_user();con=db()
    report_date=(request.args.get("date") or datetime.date.today().isoformat()).strip()
    company=(request.args.get("company") or "").strip().upper()
    if u["role"]=="GM": company=u["company_code"]
    if company not in ("","SCIC","WWIC"): company=""
    rows,month_start,month_end=_management_am_report_rows(con,u,report_date,company)
    con.close()

    wb=Workbook()
    ws=wb.active
    ws.title="AM Management Report"
    gold="E6B73F";navy="0B2744";white="FFFFFF";green="28C77A"
    ws.merge_cells("A1:L1")
    ws["A1"]="IMS • AM WISE MANAGEMENT REPORT"
    ws["A1"].font=Font(bold=True,size=16,color=gold)
    ws["A2"]="Report Date";ws["B2"]=report_date
    ws["D2"]="Month Till Date";ws["E2"]=month_start+" to "+month_end
    ws["G2"]="Scope";ws["H2"]=("ABC" if company=="SCIC" else "XYZ" if company=="WWIC" else "Both Companies")

    headers=["Company","AM Name","Today Allocated","Today Interested","Today Enrolled","Today Revenue",
             "MTD Allocated","MTD Interested","MTD Enrolled","MTD Revenue","MTD Conversion %","Report Date"]
    for c,h in enumerate(headers,1):
        cell=ws.cell(row=4,column=c,value=h);cell.font=Font(bold=True,color=white);cell.fill=__import__("openpyxl").styles.PatternFill("solid",fgColor=navy)

    r=5
    for x in rows:
        vals=[x["company_code"],x["am_name"],x["today_allocated"],x["today_interested"],x["today_enrolled"],x["today_revenue"],
              x["mtd_allocated"],x["mtd_interested"],x["mtd_enrolled"],x["mtd_revenue"],x["conversion"],report_date]
        for c,v in enumerate(vals,1): ws.cell(row=r,column=c,value=v)
        ws.cell(r,6).number_format='₹#,##0'
        ws.cell(r,10).number_format='₹#,##0'
        ws.cell(r,11).number_format='0.0%'
        ws.cell(r,11).value=(x["conversion"]/100.0)
        r+=1

    # Totals row
    ws.cell(r,1,"TOTAL");ws.cell(r,1).font=Font(bold=True,color=gold)
    for col,key in [(3,"today_allocated"),(4,"today_interested"),(5,"today_enrolled"),(6,"today_revenue"),
                    (7,"mtd_allocated"),(8,"mtd_interested"),(9,"mtd_enrolled"),(10,"mtd_revenue")]:
        ws.cell(r,col,sum(x[key] for x in rows));ws.cell(r,col).font=Font(bold=True)
    ws.cell(r,6).number_format='₹#,##0';ws.cell(r,10).number_format='₹#,##0'

    widths=[12,24,15,16,14,16,14,15,13,16,16,13]
    for i,w in enumerate(widths,1): ws.column_dimensions[__import__("openpyxl").utils.get_column_letter(i)].width=w
    ws.freeze_panes="A5"
    bio=io.BytesIO();wb.save(bio);bio.seek(0)
    scope=company or "BOTH"
    return send_file(bio,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,download_name=f"THE_GAUR_AM_Report_{scope}_{report_date}.xlsx")


@app.route("/competition")
@require_roles("MD","GM")
def competition_center():
    ensure_competition_schema();u=current_user();con=db()
    settings=get_competition_settings(con);scores=competition_score(con,settings)
    events=con.execute("SELECT * FROM competition_events ORDER BY id DESC LIMIT 100").fetchall()
    con.close();return render_template("competition.html",u=u,settings=settings,scores=scores,events=events)

@app.route("/competition/start",methods=["POST"])
@require_roles("MD")
def competition_start():
    ensure_competition_schema()
    u=current_user();con=db()
    comp_name=request.form.get("comp_name","").strip() or "ABC vs XYZ"
    start_date=request.form.get("start_date","").strip()
    end_date=request.form.get("end_date","").strip()
    target=int(request.form.get("target_enrollment") or 0)
    budget=float(request.form.get("budget") or 0)
    prize=request.form.get("prize","").strip()
    # Derive competition_month from start_date
    comp_month=start_date[:7] if start_date else ""
    if not start_date or not end_date:
        start_date,end_date=_current_month_range()
        comp_month=start_date[:7]
    now=datetime.datetime.now().isoformat(timespec="seconds")
    settings=get_competition_settings(con)
    con.execute("""UPDATE competition_settings SET title=?,start_date=?,end_date=?,enabled=1,
                   scic_target=?,wwic_target=?,status=?,budget=?,prize=?,competition_month=?,
                   popup_enabled=1,created_by=?,updated_at=? WHERE id=?""",
                (comp_name,start_date,end_date,target,target,"active",budget,prize,comp_month,
                 u["login_id"],now,settings["id"]))
    con.commit();con.close()
    log_activity("COMPETITION_STARTED","Competition Started","PERFORMANCE","Competition",settings["id"],comp_name,actor=u)
    flash(f"Competition '{comp_name}' is now ACTIVE!","success")
    return redirect(url_for("competition_center"))

@app.route("/competition/toggle",methods=["POST"])
@require_roles("MD")
def competition_toggle():
    ensure_competition_schema()
    u=current_user();con=db()
    settings=get_competition_settings(con)
    now=datetime.datetime.now().isoformat(timespec="seconds")
    new_enabled=0 if int(settings.get("enabled") or 0) else 1
    new_status="active" if new_enabled else "finished"
    con.execute("UPDATE competition_settings SET enabled=?,status=?,updated_at=? WHERE id=?",
                (new_enabled,new_status,now,settings["id"]))
    con.commit();con.close()
    action="Started" if new_enabled else "Stopped"
    log_activity(f"COMPETITION_{action.upper()}",f"Competition {action}","PERFORMANCE","Competition",settings["id"],settings.get("title",""),actor=u)
    flash(f"Competition {action}!","success")
    return redirect(request.referrer or url_for("dashboard"))

@app.route("/competition/finish",methods=["POST"])
@require_roles("MD")
def competition_finish():
    ensure_competition_schema()
    u=current_user();con=db()
    settings=get_competition_settings(con)
    now=datetime.datetime.now().isoformat(timespec="seconds")
    con.execute("UPDATE competition_settings SET enabled=0,status=?,updated_at=? WHERE id=?",
                ("finished",now,settings["id"]))
    con.commit();con.close()
    log_activity("COMPETITION_FINISHED","Competition Finished","PERFORMANCE","Competition",settings["id"],settings.get("title",""),actor=u)
    flash("Competition has been marked as FINISHED.","success")
    return redirect(url_for("competition_center"))

@app.route("/competition/settings",methods=["POST"])
@require_roles("MD")
def competition_settings_save():
    ensure_competition_schema();u=current_user();con=db();settings=get_competition_settings(con)
    con.execute("""UPDATE competition_settings SET title=?,start_date=?,end_date=?,enabled=?,scic_target=?,wwic_target=?,
                   scic_revenue_target=?,wwic_revenue_target=?,popup_enabled=?,sound_enabled=?,created_by=?,updated_at=? WHERE id=?""",
                (request.form.get("title","ABC vs XYZ").strip(),request.form.get("start_date"),request.form.get("end_date"),
                 1 if request.form.get("enabled")=="1" else 0,int(request.form.get("scic_target") or 0),int(request.form.get("wwic_target") or 0),
                 float(request.form.get("scic_revenue_target") or 0),float(request.form.get("wwic_revenue_target") or 0),
                 1 if request.form.get("popup_enabled")=="1" else 0,1 if request.form.get("sound_enabled")=="1" else 0,
                 u["login_id"],datetime.datetime.now().isoformat(timespec="seconds"),settings["id"]))
    con.commit();con.close()
    log_activity("COMPETITION_SETTINGS_UPDATED","Monthly Competition Settings Updated","PERFORMANCE","Competition",settings["id"],request.form.get("title",""),actor=u)
    flash("Competition settings updated.","success");return redirect(url_for("competition_center"))

@app.route("/api/competition/live")
def competition_live_api():
    u=current_user()
    if not u:return jsonify({"enabled":False}),401
    ensure_competition_schema();con=db();settings=get_competition_settings(con);scores=competition_score(con,settings)
    after=int(request.args.get("after") or 0)
    events=[]
    for r in con.execute("SELECT * FROM competition_events WHERE id>? ORDER BY id ASC LIMIT 20",(after,)).fetchall():
        e=dict(r)
        case=con.execute("SELECT client_name,assigned_employee_id FROM client_cases WHERE case_id=?",(e.get("case_id"),)).fetchone()
        e["client_name"]=case["client_name"] if case else ""
        e["employee_id"]=case["assigned_employee_id"] if case else None
        e["employee_role"]="Assistant Manager"
        if case and case["assigned_employee_id"]:
            er=con.execute("SELECT full_name,designation,department FROM employee_master WHERE id=?",(case["assigned_employee_id"],)).fetchone()
            if er:
                e["employee_name"]=er["full_name"] or e.get("employee_name") or "Team Member"
                e["employee_role"]=er["designation"] or er["department"] or "Assistant Manager"
                e["photo_url"]=url_for("competition_employee_photo",employee_id=case["assigned_employee_id"])
        events.append(e)
    latest_id=con.execute("SELECT COALESCE(MAX(id),0) m FROM competition_events").fetchone()["m"]
    show_payment=(u["role"]=="MD")
    # SECURITY: Company isolation — but GM sees BOTH companies in competition tab
    user_company=u.get("company_code","")
    # GM can see both companies in competition (it's a cross-company competition view)
    is_wwic_only=False
    is_scic_only=False
    if not show_payment:
        for code_key in scores:
            scores[code_key]["revenue"]=0
        for e in events:
            e["revenue"]=0
            e["payment_visible"]=False
    else:
        for e in events:
            e["payment_visible"]=True
    con.close()
    comp_status=settings.get("status","not_started") or ("active" if settings.get("enabled") else "not_started")
    return jsonify({"enabled":bool(settings["enabled"]),"popup_enabled":bool(settings["popup_enabled"]),
                    "sound_enabled":bool(settings["sound_enabled"]),"title":settings["title"],"start_date":settings["start_date"],
                    "end_date":settings["end_date"],"scic_target":settings["scic_target"],"wwic_target":settings["wwic_target"],
                    "scores":scores,"events":events,"latest_id":latest_id,"show_payment":show_payment,
                    "status":comp_status,"budget":settings.get("budget",0),"prize":settings.get("prize",""),
                    "competition_month":settings.get("competition_month",""),
                    "is_wwic_only":is_wwic_only,"is_scic_only":is_scic_only})

@app.route("/competition/employee-photo/<int:employee_id>")
def competition_employee_photo(employee_id):
    # Same enrollment celebration is visible across every authenticated portal.
    if not current_user(): return ("",401)
    con=db();row=con.execute("SELECT photo_path,photo_data,photo_mime FROM employee_master WHERE id=?",(employee_id,)).fetchone();con.close()
    if not row: return ("",404)
    pdata=row["photo_data"]
    if pdata:
        if isinstance(pdata,memoryview): pdata=pdata.tobytes()
        resp=send_file(io.BytesIO(pdata),mimetype=row["photo_mime"] or "image/jpeg",max_age=0)
        resp.headers["Cache-Control"]="no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"]="no-cache"
        resp.headers["Expires"]="0"
        return resp
    if not row["photo_path"]: return ("",404)
    path=os.path.join(APP_DIR,row["photo_path"])
    if not os.path.exists(path): return ("",404)
    return send_file(path)


@app.route("/api/celebrations/birthdays")
def birthday_celebrations_api():
    u=current_user()
    if not u:
        return jsonify({"birthdays":[]}),401
    today=datetime.date.today()
    con=db()
    rows=con.execute("""SELECT id,employee_code,company_code,full_name,designation,department,date_of_birth,
                               employment_status
                        FROM employee_master
                        WHERE COALESCE(date_of_birth,'')<>'' AND COALESCE(employment_status,'Active')<>'Inactive'
                        ORDER BY company_code,full_name""").fetchall()
    con.close()
    birthdays=[]
    for r in rows:
        raw=(r["date_of_birth"] or "").strip()
        parsed=None
        for fmt in ("%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%m/%d/%Y"):
            try:
                parsed=datetime.datetime.strptime(raw,fmt).date()
                break
            except Exception:
                pass
        if not parsed or (parsed.month,parsed.day)!=(today.month,today.day):
            continue
        birthdays.append({
            "event_type":"BIRTHDAY",
            "id":f"BIRTHDAY-{today.isoformat()}-{r['id']}",
            "employee_id":r["id"],
            "employee_code":r["employee_code"] or "",
            "company_code":r["company_code"] or "",
            "employee_name":r["full_name"] or "Team Member",
            "employee_role":r["designation"] or r["department"] or "Team Member",
            "photo_url":url_for("competition_employee_photo",employee_id=r["id"]),
            "date":today.isoformat()
        })
    return jsonify({"birthdays":birthdays,"date":today.isoformat()})









FINANCE_HEADS=["Client Payment / Service Fee","Embassy Fee","VFS / Biometric Fee","Hotel Booking","Flight / Ticket","Funds Arrangement","Client Refund","Office Rent","Electricity","Internet / Telephone","Salary / Staff Advance","Marketing / Advertisement","Travel / Fuel","Refreshment","Stationery / Printing","Courier","Legal / Professional Fee","Vendor Payment","Other Income","Miscellaneous Expense"]



def _finance_companies(u):
    return ["SCIC","WWIC"] if u["role"]=="MD" else [u["company_code"]]




def ensure_competition_schema():
    con=db();cur=con.cursor()
    try:
        if IS_POSTGRES:
            try: cur.execute("SET lock_timeout TO '3000ms'")
            except Exception: pass
            cur.execute("""CREATE TABLE IF NOT EXISTS competition_settings(
              id BIGSERIAL PRIMARY KEY,title TEXT DEFAULT 'ABC vs XYZ',
              start_date TEXT DEFAULT '',end_date TEXT DEFAULT '',enabled INTEGER DEFAULT 1,
              scic_target INTEGER DEFAULT 0,wwic_target INTEGER DEFAULT 0,
              scic_revenue_target REAL DEFAULT 0,wwic_revenue_target REAL DEFAULT 0,
              popup_enabled INTEGER DEFAULT 1,sound_enabled INTEGER DEFAULT 0,
              created_by TEXT DEFAULT '',updated_at TEXT DEFAULT '',
              status TEXT DEFAULT 'not_started',budget REAL DEFAULT 0,
              prize TEXT DEFAULT '',competition_month TEXT DEFAULT '')""")
            cur.execute("""CREATE TABLE IF NOT EXISTS competition_events(
              id BIGSERIAL PRIMARY KEY,event_type TEXT DEFAULT 'ENROLLMENT',company_code TEXT DEFAULT '',
              case_id TEXT DEFAULT '',employee_name TEXT DEFAULT '',enrollment_count INTEGER DEFAULT 0,
              revenue REAL DEFAULT 0,message TEXT DEFAULT '',created_at TEXT DEFAULT '')""")
            # Add new columns if table existed before
            for col,ctype in [("status","TEXT DEFAULT 'not_started'"),("budget","REAL DEFAULT 0"),
                               ("prize","TEXT DEFAULT ''"),("competition_month","TEXT DEFAULT ''")]:
                try:
                    cur.execute(f"ALTER TABLE competition_settings ADD COLUMN IF NOT EXISTS {col} {ctype}")
                    con.commit()
                except Exception:
                    try: con.rollback()
                    except Exception: pass
        else:
            cur.execute("""CREATE TABLE IF NOT EXISTS competition_settings(
              id INTEGER PRIMARY KEY AUTOINCREMENT,title TEXT DEFAULT 'ABC vs XYZ',
              start_date TEXT DEFAULT '',end_date TEXT DEFAULT '',enabled INTEGER DEFAULT 1,
              scic_target INTEGER DEFAULT 0,wwic_target INTEGER DEFAULT 0,
              scic_revenue_target REAL DEFAULT 0,wwic_revenue_target REAL DEFAULT 0,
              popup_enabled INTEGER DEFAULT 1,sound_enabled INTEGER DEFAULT 0,
              created_by TEXT DEFAULT '',updated_at TEXT DEFAULT '',
              status TEXT DEFAULT 'not_started',budget REAL DEFAULT 0,
              prize TEXT DEFAULT '',competition_month TEXT DEFAULT '')""")
            cur.execute("""CREATE TABLE IF NOT EXISTS competition_events(
              id INTEGER PRIMARY KEY AUTOINCREMENT,event_type TEXT DEFAULT 'ENROLLMENT',company_code TEXT DEFAULT '',
              case_id TEXT DEFAULT '',employee_name TEXT DEFAULT '',enrollment_count INTEGER DEFAULT 0,
              revenue REAL DEFAULT 0,message TEXT DEFAULT '',created_at TEXT DEFAULT '')""")
            # Add new columns if table existed before
            try:
                existing=[r["name"] for r in cur.execute("PRAGMA table_info(competition_settings)").fetchall()]
            except Exception:
                existing=[]
            for col,ctype in [("status","TEXT DEFAULT 'not_started'"),("budget","REAL DEFAULT 0"),
                               ("prize","TEXT DEFAULT ''"),("competition_month","TEXT DEFAULT ''")]:
                if col not in existing:
                    try:
                        cur.execute(f"ALTER TABLE competition_settings ADD COLUMN {col} {ctype}")
                        con.commit()
                    except Exception:
                        try: con.rollback()
                        except Exception: pass
        try: cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_competition_case_event ON competition_events(case_id,event_type)")
        except Exception: pass
        try: cur.execute("CREATE INDEX IF NOT EXISTS idx_competition_events_created ON competition_events(created_at,id)")
        except Exception: pass
        con.commit()
    except Exception:
        try: con.rollback()
        except Exception: pass
    finally: con.close()

def _current_month_range():
    today=datetime.date.today();first=today.replace(day=1)
    nxt=datetime.date(today.year+1,1,1) if today.month==12 else datetime.date(today.year,today.month+1,1)
    return first.isoformat(),(nxt-datetime.timedelta(days=1)).isoformat()

def get_competition_settings(con):
    row=con.execute("SELECT * FROM competition_settings ORDER BY id DESC LIMIT 1").fetchone()
    if row:return dict(row)
    start,end=_current_month_range();now=datetime.datetime.now().isoformat(timespec="seconds")
    con.execute("""INSERT INTO competition_settings(title,start_date,end_date,enabled,popup_enabled,sound_enabled,created_by,updated_at)
                   VALUES(?,?,?,?,?,?,?,?)""",
                ("ABC vs XYZ • Monthly Competition",start,end,1,1,0,"SYSTEM",now))
    con.commit()
    return dict(con.execute("SELECT * FROM competition_settings ORDER BY id DESC LIMIT 1").fetchone())

def competition_score(con,settings):
    scores={}
    for code in ("SCIC","WWIC"):
        r=con.execute("""SELECT COUNT(*) enrollments,COALESCE(SUM(COALESCE(total_received,0)),0) revenue
                         FROM client_cases WHERE company_code=? AND enrollment_date BETWEEN ? AND ? AND COALESCE(deleted_at,'')=''""",
                      (code,settings["start_date"],settings["end_date"])).fetchone()
        scores[code]={"enrollments":int(r["enrollments"] or 0),"revenue":float(r["revenue"] or 0)}
    return scores

def record_competition_enrollment(con,case_id,company_code,employee_name,revenue):
    try:
        settings=get_competition_settings(con)
        today=datetime.date.today().isoformat()
        if not int(settings.get("enabled") or 0) or not(settings["start_date"]<=today<=settings["end_date"]):return
        if con.execute("SELECT id FROM competition_events WHERE case_id=? AND event_type='ENROLLMENT'",(case_id,)).fetchone():return
        scores=competition_score(con,settings);count=scores[company_code]["enrollments"]
        other="WWIC" if company_code=="SCIC" else "SCIC";lead=count-scores[other]["enrollments"]
        brand="ABC" if company_code=="SCIC" else "XYZ"
        hype=(f"{brand} TAKES THE LEAD BY {lead}!" if lead>0 else ("SCORES ARE LEVEL — GAME ON!" if lead==0 else f"{brand} CLOSES THE GAP!"))
        msg=f"{brand} SCORES! {employee_name or 'Team Member'} • +1 Enrollment • Monthly {count} • {hype}"
        con.execute("""INSERT INTO competition_events(event_type,company_code,case_id,employee_name,enrollment_count,revenue,message,created_at)
                       VALUES('ENROLLMENT',?,?,?,?,?,?,?)""",
                    (company_code,case_id,employee_name or "Team Member",count,float(revenue or 0),msg,datetime.datetime.now().isoformat(timespec="seconds")))
        con.commit()
    except Exception:
        try: con.rollback()
        except Exception: pass



def ensure_enrollment_payment_schema():
    """Best-effort production-safe enrollment ledger migration."""
    con=db(); cur=con.cursor()
    money="DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"
    integer="BIGINT" if IS_POSTGRES else "INTEGER"
    cols=[
      ("package_amount",money),("package_amount_datetime","TEXT DEFAULT ''"),
      ("after_visa_payment",money),("after_visa_payment_datetime","TEXT DEFAULT ''"),
      ("first_payment_status","TEXT DEFAULT 'Pending'"),("first_payment_datetime","TEXT DEFAULT ''"),
      ("second_payment_status","TEXT DEFAULT 'Pending'"),("second_payment_datetime","TEXT DEFAULT ''"),
      ("other_payment_received",money),("other_payment_datetime","TEXT DEFAULT ''"),
      ("lead_db_id",integer),("counselor_name","TEXT DEFAULT ''"),("am_name","TEXT DEFAULT ''"),
      ("interaction_date","TEXT DEFAULT ''"),
      ("first_cash",money),("first_rbl",money),("first_yes_bank",money),("first_au_bank",money),("first_indusind",money),
      ("second_cash",money),("second_rbl",money),("second_yes_bank",money),("second_au_bank",money),("second_indusind",money),
      ("first_payment_method","TEXT DEFAULT ''"),("second_payment_method","TEXT DEFAULT ''"),
      ("other_payment_method","TEXT DEFAULT ''"),("other_payment_bank","TEXT DEFAULT ''"),
      ("first_payment_bank_id",integer),("second_payment_bank_id",integer),("other_payment_bank_id",integer),
      ("deleted_at","TEXT DEFAULT ''"),("deleted_by","TEXT DEFAULT ''"),("deletion_reason","TEXT DEFAULT ''"),
      ("am_submitted","INTEGER DEFAULT 0"),("am_submitted_by","TEXT DEFAULT ''"),("am_submitted_at","TEXT DEFAULT ''")
    ]
    try:
        if IS_POSTGRES:
            try:
                cur.execute("SET lock_timeout TO '2000ms'"); con.commit()
            except Exception:
                try: con.rollback()
                except Exception: pass
            for c,t in cols:
                try:
                    cur.execute(f"ALTER TABLE client_cases ADD COLUMN IF NOT EXISTS {c} {t}")
                    con.commit()
                except Exception:
                    try: con.rollback()
                    except Exception: pass
            try:
                cur.execute("""CREATE TABLE IF NOT EXISTS client_payment_proofs(
                  id BIGSERIAL PRIMARY KEY,case_id BIGINT NOT NULL,payment_stage TEXT DEFAULT '',
                  file_name TEXT DEFAULT '',mime_type TEXT DEFAULT '',file_data BYTEA,
                  uploaded_by_id BIGINT,uploaded_by_name TEXT DEFAULT '',uploaded_at TEXT DEFAULT '')""")
                con.commit()
            except Exception:
                try: con.rollback()
                except Exception: pass
        else:
            try:
                existing=[r["name"] for r in cur.execute("PRAGMA table_info(client_cases)").fetchall()]
            except Exception:
                existing=[]
            for c,t in cols:
                if c not in existing:
                    try:
                        cur.execute(f"ALTER TABLE client_cases ADD COLUMN {c} {t}")
                        con.commit()
                    except Exception:
                        try: con.rollback()
                        except Exception: pass
            try:
                cur.execute("""CREATE TABLE IF NOT EXISTS client_payment_proofs(
                  id INTEGER PRIMARY KEY AUTOINCREMENT,case_id INTEGER NOT NULL,payment_stage TEXT DEFAULT '',
                  file_name TEXT DEFAULT '',mime_type TEXT DEFAULT '',file_data BLOB,
                  uploaded_by_id INTEGER,uploaded_by_name TEXT DEFAULT '',uploaded_at TEXT DEFAULT '')""")
                con.commit()
            except Exception:
                try: con.rollback()
                except Exception: pass
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_client_payment_proofs_case ON client_payment_proofs(case_id,payment_stage,id)")
            con.commit()
        except Exception:
            try: con.rollback()
            except Exception: pass
    finally:
        con.close()




def active_client_case_predicate(con):
    """Safe active-enrollment condition for old/new databases."""
    try:
        cols=_table_columns(con,"client_cases")
        if "deleted_at" in cols:
            return "COALESCE(deleted_at,'')=''"
    except Exception:
        try: con.rollback()
        except Exception: pass
    return "1=1"

def ensure_enrollment_recycle_schema():
    con=db(); cur=con.cursor()
    money="DOUBLE PRECISION DEFAULT 0" if IS_POSTGRES else "REAL DEFAULT 0"
    integer="BIGINT" if IS_POSTGRES else "INTEGER"
    cols=[
        ("deleted_at","TEXT DEFAULT ''"),("deleted_by","TEXT DEFAULT ''"),("deletion_reason","TEXT DEFAULT ''"),
        ("first_cash",money),("first_rbl",money),("first_yes_bank",money),("first_au_bank",money),("first_indusind",money),
        ("second_cash",money),("second_rbl",money),("second_yes_bank",money),("second_au_bank",money),("second_indusind",money),
        ("first_payment_method","TEXT DEFAULT ''"),("second_payment_method","TEXT DEFAULT ''"),
        ("other_payment_method","TEXT DEFAULT ''"),("other_payment_bank","TEXT DEFAULT ''"),
        ("first_payment_status","TEXT DEFAULT 'Pending'"),("first_payment_datetime","TEXT DEFAULT ''"),
        ("second_payment_status","TEXT DEFAULT 'Pending'"),("second_payment_datetime","TEXT DEFAULT ''"),
        ("other_payment_received",money),("other_payment_datetime","TEXT DEFAULT ''"),
        ("lead_db_id",integer),("counselor_name","TEXT DEFAULT ''"),("am_name","TEXT DEFAULT ''"),
        ("interaction_date","TEXT DEFAULT ''"),
        ("first_payment_bank_id",integer),("second_payment_bank_id",integer),("other_payment_bank_id",integer),
    ]
    try:
        if IS_POSTGRES:
            for c,t in cols:
                try:
                    cur.execute(f"ALTER TABLE client_cases ADD COLUMN IF NOT EXISTS {c} {t}")
                    con.commit()
                except Exception:
                    try: con.rollback()
                    except Exception: pass
        else:
            try:
                existing=[r["name"] for r in cur.execute("PRAGMA table_info(client_cases)").fetchall()]
            except Exception:
                existing=[]
            for c,t in cols:
                if c not in existing:
                    try:
                        cur.execute(f"ALTER TABLE client_cases ADD COLUMN {c} {t}")
                        con.commit()
                    except Exception:
                        try: con.rollback()
                        except Exception: pass
    finally:
        con.close()

def _table_columns(con,table_name):
    try:
        if IS_POSTGRES:
            rows=con.execute("""SELECT column_name FROM information_schema.columns
                                WHERE table_schema=current_schema() AND table_name=?""",(table_name,)).fetchall()
            return {r["column_name"] for r in rows}
        return {r["name"] for r in con.execute(f"PRAGMA table_info({table_name})").fetchall()}
    except Exception:
        try: con.rollback()
        except Exception: pass
        return set()

def _table_exists(con,table_name):
    try:
        if IS_POSTGRES:
            r=con.execute("""SELECT 1 ok FROM information_schema.tables
                             WHERE table_schema=current_schema() AND table_name=? LIMIT 1""",(table_name,)).fetchone()
            return bool(r)
        r=con.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?",(table_name,)).fetchone()
        return bool(r)
    except Exception:
        try: con.rollback()
        except Exception: pass
        return False

def ensure_bank_manager_schema():
    """Lazy finance migration so Railway can bind /healthz before PostgreSQL DDL."""
    con=db();cur=con.cursor()
    try:
        if IS_POSTGRES:
            try: cur.execute("SET lock_timeout TO '3000ms'")
            except Exception: pass
        for col,ddl in [
            ("upi_id","TEXT DEFAULT ''"),("city","TEXT DEFAULT ''"),("rm_name","TEXT DEFAULT ''"),
            ("rm_mobile","TEXT DEFAULT ''"),("account_nickname","TEXT DEFAULT ''"),("share_enabled","INTEGER DEFAULT 1")
        ]:
            try:
                if IS_POSTGRES: cur.execute(f"ALTER TABLE finance_banks ADD COLUMN IF NOT EXISTS {col} {ddl}")
                else: cur.execute(f"ALTER TABLE finance_banks ADD COLUMN {col} {ddl}")
                con.commit()
            except Exception:
                try: con.rollback()
                except Exception: pass
        try:
            if IS_POSTGRES:
                cur.execute("""CREATE TABLE IF NOT EXISTS bank_share_logs(
                  id BIGSERIAL PRIMARY KEY,bank_id BIGINT,company_code TEXT DEFAULT '',shared_by_id BIGINT,
                  shared_by_name TEXT DEFAULT '',recipient_mobile TEXT DEFAULT '',share_type TEXT DEFAULT 'WHATSAPP',
                  shared_at TEXT DEFAULT '')""")
            else:
                cur.execute("""CREATE TABLE IF NOT EXISTS bank_share_logs(
                  id INTEGER PRIMARY KEY AUTOINCREMENT,bank_id INTEGER,company_code TEXT DEFAULT '',shared_by_id INTEGER,
                  shared_by_name TEXT DEFAULT '',recipient_mobile TEXT DEFAULT '',share_type TEXT DEFAULT 'WHATSAPP',
                  shared_at TEXT DEFAULT '')""")
            con.commit()
        except Exception:
            try: con.rollback()
            except Exception: pass
        try:
            cur.execute("CREATE INDEX IF NOT EXISTS idx_bank_share_logs ON bank_share_logs(bank_id,shared_at)")
            con.commit()
        except Exception:
            try: con.rollback()
            except Exception: pass
    finally:
        con.close()

BANK_LOGO_DOMAINS={
    "hdfc":"hdfcbank.com","icici":"icicibank.com","axis":"axisbank.com","state bank":"sbi.co.in",
    "sbi":"sbi.co.in","punjab national":"pnbindia.in","pnb":"pnbindia.in","kotak":"kotak.com",
    "yes bank":"yesbank.in","indusind":"indusind.com","idfc":"idfcfirstbank.com","federal":"federalbank.co.in",
    "bank of baroda":"bankofbaroda.in","canara":"canarabank.com","union bank":"unionbankofindia.co.in"
}
def _bank_logo_url(name):
    n=(name or "").lower()
    for key,domain in BANK_LOGO_DOMAINS.items():
        if key in n:return "https://logo.clearbit.com/"+domain
    return ""

def _payment_message(bank,company_name):
    return (
        f"Greetings from {company_name}\\n\\n"
        "Thank you for choosing us. For your convenience, please find our official payment details below.\\n\\n"
        f"Bank: {bank['bank_name']}\\n"
        f"Account Name: {bank['account_name']}\\n"
        f"Account No.: {bank['account_number']}\\n"
        f"IFSC: {bank['ifsc_code'] or '-'}\\n"
        f"Branch: {bank['branch'] or '-'}"
        + (f"\\nUPI ID: {bank['upi_id']}" if bank.get('upi_id') else "") +
        "\\n\\nPlease verify the account name before making payment. After payment, kindly share the transaction screenshot / UTR for verification.\\n\\n"
        "“Your trust is our responsibility, and your journey is our commitment.”\\n\\n"
        f"Regards,\\n{company_name}"
    )


def _finance_voucher(con,company,direction,date_text):
    ym=(date_text or datetime.date.today().isoformat())[:7].replace("-","")
    tag="RCP" if direction=="IN" else ("REF" if direction=="REFUND" else "PAY")
    prefix=f"{company}-{tag}-{ym}-"
    row=con.execute("SELECT COUNT(*) c FROM finance_transactions WHERE voucher_no LIKE ?",(prefix+"%",)).fetchone()
    return prefix+str(int(row["c"])+1).zfill(4)


def _norm_text(v):
    return re.sub(r"[^A-Z0-9]","",(str(v or "").upper()))

def _norm_date(v):
    if not v: return ""
    if isinstance(v,datetime.datetime): return v.date().isoformat()
    if isinstance(v,datetime.date): return v.isoformat()
    text=str(v).strip()
    for fmt in ("%Y-%m-%d","%d-%m-%Y","%d/%m/%Y","%m/%d/%Y","%d.%m.%Y"):
        try:return datetime.datetime.strptime(text[:10],fmt).date().isoformat()
        except Exception:pass
    try:
        if text.replace(".","",1).isdigit():
            return (datetime.datetime(1899,12,30)+datetime.timedelta(days=float(text))).date().isoformat()
    except Exception:pass
    return text[:10]

def _money(v):
    if v is None or v=="": return 0.0
    try:return float(str(v).replace(",","").replace("₹","").strip())
    except Exception:return 0.0

def _statement_rows(file_obj):
    name=(file_obj.filename or "").lower()
    rows=[]
    if name.endswith(".csv"):
        text=file_obj.read().decode("utf-8-sig",errors="ignore")
        data=list(csv.reader(io.StringIO(text)))
    else:
        wb=load_workbook(io.BytesIO(file_obj.read()),data_only=True,read_only=True)
        ws=wb.active
        data=[list(r) for r in ws.iter_rows(values_only=True)]
    if not data:return []
    # discover header row in first 12 rows
    header_idx=0
    keywords=("date","narration","description","particular","reference","utr","debit","credit","withdrawal","deposit","balance")
    best=-1
    for i,row in enumerate(data[:12]):
        score=sum(1 for c in row if any(k in str(c or "").lower() for k in keywords))
        if score>best:best=score;header_idx=i
    headers=[str(x or "").strip().lower() for x in data[header_idx]]
    def findcol(*names):
        for i,h in enumerate(headers):
            if any(n in h for n in names):return i
        return None
    date_i=findcol("date")
    narr_i=findcol("narration","description","particular")
    ref_i=findcol("reference","ref no","utr","transaction id","cheque")
    debit_i=findcol("debit","withdrawal","dr amount")
    credit_i=findcol("credit","deposit","cr amount")
    amount_i=findcol("amount")
    type_i=findcol("type","dr/cr")
    bal_i=findcol("balance")
    time_i=findcol("time")
    for row in data[header_idx+1:]:
        if not any(x not in (None,"") for x in row):continue
        date=_norm_date(row[date_i] if date_i is not None and date_i<len(row) else "")
        narr=str(row[narr_i] if narr_i is not None and narr_i<len(row) else "")
        ref=str(row[ref_i] if ref_i is not None and ref_i<len(row) else "")
        debit=_money(row[debit_i] if debit_i is not None and debit_i<len(row) else 0)
        credit=_money(row[credit_i] if credit_i is not None and credit_i<len(row) else 0)
        if amount_i is not None and amount_i<len(row) and debit==0 and credit==0:
            amt=_money(row[amount_i]);typ=str(row[type_i] if type_i is not None and type_i<len(row) else "").upper()
            if "DR" in typ or "DEBIT" in typ:debit=amt
            else:credit=amt
        balance=_money(row[bal_i] if bal_i is not None and bal_i<len(row) else 0)
        tm=str(row[time_i] if time_i is not None and time_i<len(row) else "")
        if date or debit or credit:
            rows.append({"date":date,"time":tm[:12],"narration":narr[:500],"reference":ref[:180],"debit":debit,"credit":credit,"balance":balance,"raw":row})
    return rows

def _match_score(portal,bank):
    amount=float(portal["amount"] or 0)
    bank_amount=float(bank["credit"] if portal["direction"]=="IN" else bank["debit"])
    if amount<=0 or bank_amount<=0:return (0,"No comparable amount")
    diff=abs(amount-bank_amount)
    score=0;reasons=[]
    if diff<0.01:score+=55;reasons.append("Exact amount")
    elif diff<=1:score+=45;reasons.append("Amount within ₹1")
    elif diff/max(amount,1)<=0.01:score+=30;reasons.append("Amount within 1%")
    else:return (0,"Amount mismatch")
    pd=_norm_date(portal["txn_date"]);bd=bank["txn_date"]
    try:
        dd=abs((datetime.date.fromisoformat(pd)-datetime.date.fromisoformat(bd)).days)
        if dd==0:score+=25;reasons.append("Same date")
        elif dd==1:score+=15;reasons.append("±1 day")
        elif dd<=3:score+=6;reasons.append("Within 3 days")
    except Exception:pass
    pref=_norm_text(portal["reference_no"]);bref=_norm_text(bank["reference_no"]+" "+bank["narration"])
    if pref and pref in bref:score+=20;reasons.append("Reference/UTR matched")
    elif pref and len(pref)>=6 and any(pref[i:i+6] in bref for i in range(max(1,len(pref)-5))):score+=10;reasons.append("Reference partial match")
    party=_norm_text(portal["party_name"])
    narr=_norm_text(bank["narration"])
    if party and len(party)>=5 and party[:8] in narr:score+=5;reasons.append("Party/narration match")
    return (min(100,score),", ".join(reasons) or "Heuristic match")

def _build_reconciliation(con,run_id,company):
    con.execute("DELETE FROM bank_recon_matches WHERE run_id=?",(run_id,))
    bank_rows=con.execute("SELECT * FROM bank_statement_rows WHERE run_id=? ORDER BY id",(run_id,)).fetchall()
    portals=con.execute("SELECT * FROM finance_transactions WHERE company_code=? AND status='POSTED' ORDER BY txn_date,id",(company,)).fetchall()
    candidates=[]
    for p in portals:
        best=None
        for b in bank_rows:
            score,reason=_match_score(p,b)
            if score and (best is None or score>best[0]):best=(score,reason,b["id"])
        if best:
            status="VERIFIED" if best[0]>=90 else ("POSSIBLE" if best[0]>=65 else "REVIEW")
            con.execute("INSERT INTO bank_recon_matches(run_id,portal_txn_id,bank_row_id,score,status,match_reason) VALUES(?,?,?,?,?,?)",(run_id,p["id"],best[2],best[0],status,best[1]))
    con.commit()

@app.route("/finance/reconciliation",methods=["GET","POST"])
@require_roles("MD")
def finance_reconciliation():
    u=current_user();con=db();companies=_finance_companies(u)
    banks=con.execute("SELECT * FROM finance_banks WHERE active=1 AND company_code IN ("+",".join(["?"]*len(companies))+") ORDER BY company_code,bank_name",companies).fetchall()
    if request.method=="POST":
        company=request.form.get("company_code","")
        bank_id=request.form.get("bank_id","")
        file=request.files.get("statement")
        if company not in companies or not bank_id.isdigit() or not file or not file.filename:
            con.close();flash("Company, bank account and statement file are required.","error");return redirect(url_for("finance_reconciliation"))
        if not file.filename.lower().endswith((".xlsx",".xls",".csv")):
            con.close();flash("Upload Excel or CSV bank statement.","error");return redirect(url_for("finance_reconciliation"))
        parsed=_statement_rows(file)
        if not parsed:
            con.close();flash("No usable bank transactions found in statement.","error");return redirect(url_for("finance_reconciliation"))
        now=datetime.datetime.now().isoformat(timespec="seconds")
        if IS_POSTGRES:
            run_id=con.execute("INSERT INTO bank_recon_runs(company_code,bank_id,file_name,total_rows,created_by_id,created_by_name,created_at) VALUES(?,?,?,?,?,?,?) RETURNING id",(company,int(bank_id),secure_filename(file.filename),len(parsed),u["id"],u["full_name"],now)).fetchone()["id"]
        else:
            cur=con.execute("INSERT INTO bank_recon_runs(company_code,bank_id,file_name,total_rows,created_by_id,created_by_name,created_at) VALUES(?,?,?,?,?,?,?)",(company,int(bank_id),secure_filename(file.filename),len(parsed),u["id"],u["full_name"],now));run_id=cur.lastrowid
        for r in parsed:
            con.execute("INSERT INTO bank_statement_rows(run_id,txn_date,txn_time,narration,reference_no,debit,credit,balance,raw_json) VALUES(?,?,?,?,?,?,?,?,?)",(run_id,r["date"],r["time"],r["narration"],r["reference"],r["debit"],r["credit"],r["balance"],json.dumps(r["raw"],default=str)[:3000]))
        con.commit();_build_reconciliation(con,run_id,company)
        con.close()
        log_activity("BANK_RECON_UPLOAD","Bank Statement Reconciled","FINANCE","Reconciliation Run",run_id,str(run_id),details={"company":company,"rows":len(parsed),"file":file.filename},actor=u,company_code=company)
        return redirect(url_for("finance_reconciliation",run_id=run_id))
    rid=request.args.get("run_id","").strip()
    run=None;matches=[];attention=[]
    if rid.isdigit():
        run=con.execute("SELECT * FROM bank_recon_runs WHERE id=?",(int(rid),)).fetchone()
        if run and run["company_code"] in companies:
            matches=con.execute("""SELECT m.*,p.voucher_no portal_voucher,p.txn_date portal_date,p.amount portal_amount,p.party_name portal_party,p.reference_no portal_reference,
            b.txn_date bank_date,b.txn_time bank_time,b.reference_no bank_reference,b.narration bank_narration,
            CASE WHEN p.direction='IN' THEN b.credit ELSE b.debit END bank_amount,
            CASE WHEN p.direction='IN' THEN 'CREDIT' ELSE 'DEBIT' END bank_type
            FROM bank_recon_matches m JOIN finance_transactions p ON p.id=m.portal_txn_id JOIN bank_statement_rows b ON b.id=m.bank_row_id
            WHERE m.run_id=? ORDER BY m.score DESC,m.id""",(run["id"],)).fetchall()
            matched_portal={x["portal_txn_id"] for x in matches};matched_bank={x["bank_row_id"] for x in matches if x["score"]>=65}
            portals=con.execute("SELECT * FROM finance_transactions WHERE company_code=? AND status='POSTED'",(run["company_code"],)).fetchall()
            for p in portals:
                if p["id"] not in matched_portal:attention.append({"title":"Portal Entry Not Verified","details":f"{p['voucher_no']} • {p['txn_date']} • ₹{p['amount']:.2f} • {p['party_name']}"})
            brows=con.execute("SELECT * FROM bank_statement_rows WHERE run_id=?",(run["id"],)).fetchall()
            for b in brows:
                if b["id"] not in matched_bank:
                    amt=b["credit"] or b["debit"];typ="Credit" if b["credit"] else "Debit"
                    attention.append({"title":"Bank Entry Without Portal Match","details":f"{b['txn_date']} • {typ} ₹{amt:.2f} • {b['reference_no']} • {b['narration'][:120]}"})
    k={"statement_count":run["total_rows"] if run else 0,"verified":sum(1 for x in matches if x["status"]=="VERIFIED"),"possible":sum(1 for x in matches if x["status"] in ("POSSIBLE","REVIEW")),"portal_unmatched":sum(1 for x in attention if x["title"].startswith("Portal")),"bank_unmatched":sum(1 for x in attention if x["title"].startswith("Bank")),"suspicious":0}
    # duplicate references in bank rows
    if run:
        dup=con.execute("SELECT COUNT(*) c FROM (SELECT reference_no FROM bank_statement_rows WHERE run_id=? AND reference_no<>'' GROUP BY reference_no HAVING COUNT(*)>1) x",(run["id"],)).fetchone()["c"]
        k["suspicious"]=dup
    con.close()
    return render_template("finance_reconciliation.html",u=u,companies=companies,banks=banks,run=run,matches=matches,attention=attention,k=k)

@app.route("/finance/reconciliation/confirm",methods=["POST"])
@require_roles("MD")
def finance_reconciliation_confirm():
    u=current_user();mid=request.form.get("match_id","")
    if not mid.isdigit():return redirect(url_for("finance_reconciliation"))
    con=db();m=con.execute("SELECT m.*,r.company_code FROM bank_recon_matches m JOIN bank_recon_runs r ON r.id=m.run_id WHERE m.id=?",(int(mid),)).fetchone()
    if not m or m["company_code"] not in _finance_companies(u):con.close();return "Access denied",403
    con.execute("UPDATE bank_recon_matches SET status='VERIFIED',confirmed_by_id=?,confirmed_at=? WHERE id=?",(u["id"],datetime.datetime.now().isoformat(timespec="seconds"),m["id"]))
    con.commit();run_id=m["run_id"];con.close()
    log_activity("BANK_MATCH_CONFIRMED","Bank Match Confirmed","FINANCE","Reconciliation Match",m["id"],str(m["id"]),details="Manual confirmation by management",actor=u,company_code=m["company_code"])
    return redirect(url_for("finance_reconciliation",run_id=run_id))

@app.route("/finance/reconciliation/export.xlsx")
@require_roles("MD")
def finance_reconciliation_export():
    u=current_user();rid=request.args.get("run_id","")
    if not rid.isdigit():return redirect(url_for("finance_reconciliation"))
    con=db();run=con.execute("SELECT * FROM bank_recon_runs WHERE id=?",(int(rid),)).fetchone()
    if not run or run["company_code"] not in _finance_companies(u):con.close();return "Access denied",403
    rows=con.execute("""SELECT m.status,m.score,m.match_reason,p.voucher_no,p.txn_date portal_date,p.amount portal_amount,p.party_name,p.reference_no portal_ref,
      b.txn_date bank_date,b.txn_time,b.narration,b.reference_no bank_ref,b.debit,b.credit
      FROM bank_recon_matches m JOIN finance_transactions p ON p.id=m.portal_txn_id JOIN bank_statement_rows b ON b.id=m.bank_row_id
      WHERE m.run_id=? ORDER BY m.score DESC""",(run["id"],)).fetchall();con.close()
    wb=Workbook();ws=wb.active;ws.title="Reconciliation"
    ws.append(["Status","Score %","Reason","Portal Voucher","Portal Date","Portal Amount","Party","Portal Reference","Bank Date","Bank Time","Bank Narration","Bank Reference","Debit","Credit"])
    for c in ws[1]:c.font=Font(bold=True)
    for r in rows:ws.append([r["status"],r["score"],r["match_reason"],r["voucher_no"],r["portal_date"],r["portal_amount"],r["party_name"],r["portal_ref"],r["bank_date"],r["txn_time"],r["narration"],r["bank_ref"],r["debit"],r["credit"]])
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    bio=io.BytesIO();wb.save(bio);bio.seek(0)
    return send_file(bio,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",download_name=f"THE_GAUR_Bank_Reconciliation_{run['id']}.xlsx",as_attachment=True)



@app.route("/accounts/reports")
@require_roles("MD","GM")
def accounts_report_center():
    u=current_user(); con=db(); today=datetime.date.today()
    mode=(request.args.get("mode") or "daily").lower()
    if mode not in ("daily","weekly","monthly","yearly"): mode="daily"
    allowed=_finance_companies(u)
    company=(request.args.get("company") or "").strip()
    if company not in allowed: company=allowed[0] if len(allowed)==1 else ""

    if mode=="daily":
        d=datetime.datetime.strptime(request.args.get("date") or today.isoformat(),"%Y-%m-%d").date(); start=end=d; label=d.strftime("%d %b %Y")
    elif mode=="weekly":
        d=datetime.datetime.strptime(request.args.get("date") or today.isoformat(),"%Y-%m-%d").date()
        start=d-datetime.timedelta(days=d.weekday()); end=start+datetime.timedelta(days=6); label=f"{start.strftime('%d %b')} – {end.strftime('%d %b %Y')}"
    elif mode=="monthly":
        y,m=map(int,(request.args.get("month") or today.strftime("%Y-%m")).split("-")); start=datetime.date(y,m,1)
        nxt=datetime.date(y+1,1,1) if m==12 else datetime.date(y,m+1,1); end=nxt-datetime.timedelta(days=1); label=start.strftime("%B %Y")
    else:
        y=int(request.args.get("year") or today.year); start=datetime.date(y,1,1); end=datetime.date(y,12,31); label=str(y)
    s,e=start.isoformat(),end.isoformat()

    cwhere=""; cp=[]
    if company: cwhere=" AND company_code=?"; cp=[company]
    cases=con.execute("SELECT * FROM client_cases WHERE COALESCE(deleted_at,'')=''"+cwhere,cp).fetchall()
    inflow=[]
    def add_in(c,amount,bank,method,dt,stage):
        amount=float(amount or 0)
        if amount<=0:return
        date=(dt or c["enrollment_date"] or c["created_at"] or "")[:10]
        if not date or not(s<=date<=e):return
        inflow.append({"date":date,"company":c["company_code"],"party":c["client_name"] or c["case_id"],
                       "ref":c["case_id"],"stage":stage,"bank":bank or "Cash",
                       "method":method or ("Cash" if (bank or "").lower()=="cash" else "Net Banking"),
                       "amount":amount,"source":"Enrollment"})
    for c in cases:
        fm=c["first_payment_method"] or "Net Banking"; sm=c["second_payment_method"] or "Net Banking"
        add_in(c,c["first_cash"],"Cash","Cash",c["first_payment_datetime"],"1st Payment")
        add_in(c,c["first_rbl"],"RBL Bank",fm,c["first_payment_datetime"],"1st Payment")
        add_in(c,c["first_yes_bank"],"Yes Bank",fm,c["first_payment_datetime"],"1st Payment")
        add_in(c,c["first_au_bank"],"AU Bank",fm,c["first_payment_datetime"],"1st Payment")
        add_in(c,c["second_cash"],"Cash","Cash",c["second_payment_datetime"],"2nd Payment")
        add_in(c,c["second_rbl"],"RBL Bank",sm,c["second_payment_datetime"],"2nd Payment")
        add_in(c,c["second_yes_bank"],"Yes Bank",sm,c["second_payment_datetime"],"2nd Payment")
        add_in(c,c["second_au_bank"],"AU Bank",sm,c["second_payment_datetime"],"2nd Payment")
        add_in(c,c["other_payment_received"],c["other_payment_bank"] or "Other / Unspecified",
               c["other_payment_method"] or "Other",c["other_payment_datetime"],"Other Payment")

    wh=["t.status='POSTED'","t.txn_date BETWEEN ? AND ?"]; pa=[s,e]
    if company: wh.append("t.company_code=?"); pa.append(company)
    manual=con.execute("""SELECT t.*,b.bank_name,b.account_name,b.account_last4 FROM finance_transactions t
                          LEFT JOIN finance_banks b ON b.id=t.account_id WHERE """+" AND ".join(wh)+" ORDER BY t.txn_date,t.id",pa).fetchall()
    outflow=[]; manual_income=[]
    for r in manual:
        item={"date":r["txn_date"],"company":r["company_code"],"party":r["party_name"],"ref":r["voucher_no"],
              "stage":r["head"],"bank":r["bank_name"] or ("Cash" if r["payment_mode"]=="Cash" else "Unspecified"),
              "method":r["payment_mode"] or "Other","amount":float(r["amount"] or 0),"source":"Accounts"}
        (manual_income if r["direction"]=="IN" else outflow).append(item)
    all_in=inflow+manual_income
    bank_totals={}; method_totals={}
    for x in all_in:
        bank_totals[x["bank"]]=bank_totals.get(x["bank"],0)+x["amount"]
        method_totals[x["method"]]=method_totals.get(x["method"],0)+x["amount"]
    total_in=sum(x["amount"] for x in all_in); total_out=sum(x["amount"] for x in outflow); net=total_in-total_out
    paired=[{"in":all_in[i] if i<len(all_in) else None,"out":outflow[i] if i<len(outflow) else None} for i in range(max(len(all_in),len(outflow)))]
    con.close()
    return render_template("accounts_reports.html",u=u,mode=mode,label=label,start=s,end=e,company=company,companies=allowed,
                           inflow=all_in,outflow=outflow,paired=paired,bank_totals=sorted(bank_totals.items(),key=lambda x:-x[1]),
                           method_totals=sorted(method_totals.items(),key=lambda x:-x[1]),total_in=total_in,total_out=total_out,net=net)


@app.route("/finance/cash-book")
@require_roles("MD","GM")
def finance_cash_book():
    u=current_user(); con=db()
    day=(request.args.get("date") or datetime.date.today().isoformat()).strip()
    month=(request.args.get("month") or day[:7]).strip()
    company=(request.args.get("company") or "").strip()
    if company not in ("SCIC","WWIC"): company=""

    dw=["t.status='POSTED'","t.txn_date=?"]; dp=[day]
    if company:
        dw.append("t.company_code=?"); dp.append(company)
    daily=con.execute("""SELECT t.*,b.bank_name,b.account_name,b.account_last4
                         FROM finance_transactions t
                         LEFT JOIN finance_banks b ON b.id=t.account_id
                         WHERE """+" AND ".join(dw)+" ORDER BY t.id",dp).fetchall()

    mw=["t.status='POSTED'","t.txn_date LIKE ?"]; mp=[month+"%"]
    if company:
        mw.append("t.company_code=?"); mp.append(company)
    monthly=con.execute("""SELECT t.*,b.bank_name,b.account_name,b.account_last4
                           FROM finance_transactions t
                           LEFT JOIN finance_banks b ON b.id=t.account_id
                           WHERE """+" AND ".join(mw)+" ORDER BY t.txn_date,t.id",mp).fetchall()

    daily_in=sum(float(r["amount"] or 0) for r in daily if r["direction"]=="IN")
    daily_out=sum(float(r["amount"] or 0) for r in daily if r["direction"] in ("OUT","REFUND"))
    monthly_in=sum(float(r["amount"] or 0) for r in monthly if r["direction"]=="IN")
    monthly_out=sum(float(r["amount"] or 0) for r in monthly if r["direction"] in ("OUT","REFUND"))
    con.close()
    return render_template("finance_cash_book.html",u=u,daily=daily,monthly=monthly,day=day,month=month,company=company,
                           daily_in=daily_in,daily_out=daily_out,daily_net=daily_in-daily_out,
                           monthly_in=monthly_in,monthly_out=monthly_out,monthly_net=monthly_in-monthly_out)

@app.route("/finance/daily-report", methods=["GET","POST"])
@require_roles("MD","GM")
def finance_daily_report():
    """Daily bank-wise Inflow & Outflow report — ABC Payment Inflow & Outflow style."""
    ensure_bank_manager_schema()
    u=current_user(); con=db(); companies=_finance_companies(u)

    # POST — add a new daily entry
    if request.method=="POST":
        company=request.form.get("company_code","").strip()
        if company not in companies:
            con.close(); flash("Company access denied.","error"); return redirect(url_for("finance_daily_report"))
        direction=request.form.get("direction","IN").strip()
        if direction not in ("IN","OUT"): direction="IN"
        txn_date=request.form.get("txn_date") or datetime.date.today().isoformat()
        amount=float(request.form.get("amount") or 0)
        if amount<=0:
            con.close(); flash("Amount must be greater than zero.","error"); return redirect(url_for("finance_daily_report"))
        voucher=_finance_voucher(con,company,direction,txn_date)
        aid=request.form.get("account_id","").strip()
        aid=int(aid) if aid.isdigit() else None
        con.execute("""INSERT INTO finance_transactions(voucher_no,txn_date,company_code,direction,head,party_name,
            amount,payment_mode,account_id,reference_no,remarks,bill_name,bill_mime,bill_bytes,
            status,created_by_id,created_by_name,created_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,'POSTED',?,?,?)""",
            (voucher,txn_date,company,direction,
             request.form.get("head","Enrollment Fee") if direction=="IN" else request.form.get("head","Miscellaneous Expense"),
             request.form.get("party_name","").strip(),amount,
             request.form.get("payment_mode","Cash"),aid,
             request.form.get("reference_no","").strip(),
             request.form.get("remarks","").strip(),
             "","",None,u["id"],u["full_name"],
             datetime.datetime.now().isoformat(timespec="seconds")))
        con.commit()
        log_activity("DAILY_REPORT_ENTRY","Daily Report Entry Posted","FINANCE",
                     details={"direction":direction,"amount":amount,"party":request.form.get("party_name")},
                     actor=u,company_code=company)
        flash(f"Entry saved: {voucher}","success")
        return redirect(url_for("finance_daily_report",date=txn_date,company=company))

    sel_date=(request.args.get("date") or datetime.date.today().isoformat()).strip()
    sel_company=(request.args.get("company") or (u["company_code"] if u["role"]=="GM" else "SCIC")).strip()
    if sel_company not in companies: sel_company=list(companies)[0]

    # Fetch all active banks for selected company
    banks=con.execute("SELECT * FROM finance_banks WHERE active=1 AND company_code=? ORDER BY bank_name",(sel_company,)).fetchall()
    bank_ids=[b["id"] for b in banks]

    # Fetch transactions for selected date
    rows=con.execute("""SELECT t.*,b.bank_name,b.account_last4,b.id AS bid
        FROM finance_transactions t
        LEFT JOIN finance_banks b ON b.id=t.account_id
        WHERE t.company_code=? AND t.txn_date=? AND t.status='POSTED'
        ORDER BY t.direction DESC,t.id""",
        (sel_company,sel_date)).fetchall()

    # Compute balance C/F — sum all prior days
    prior=con.execute("""SELECT account_id, direction, SUM(amount) as total
        FROM finance_transactions
        WHERE company_code=? AND txn_date<? AND status='POSTED'
        GROUP BY account_id,direction""",(sel_company,sel_date)).fetchall()
    balance_cf={}  # bank_id -> balance
    for p in prior:
        bid=p["account_id"]
        if bid not in balance_cf: balance_cf[bid]=0.0
        if p["direction"]=="IN": balance_cf[bid]+=float(p["total"] or 0)
        else: balance_cf[bid]-=float(p["total"] or 0)
    # Add opening_balance
    for b in banks:
        ob=float(b["opening_balance"] or 0)
        balance_cf[b["id"]]=balance_cf.get(b["id"],0.0)+ob

    inflows=[dict(r) for r in rows if r["direction"]=="IN"]
    outflows=[dict(r) for r in rows if r["direction"] in ("OUT","REFUND")]

    # Bank-wise totals for today
    bank_in={}; bank_out={}
    for r in rows:
        bid=r["account_id"]
        if bid not in bank_in: bank_in[bid]=0.0; bank_out[bid]=0.0
        if r["direction"]=="IN": bank_in[bid]+=float(r["amount"] or 0)
        else: bank_out[bid]+=float(r["amount"] or 0)

    total_in=sum(float(r["amount"] or 0) for r in rows if r["direction"]=="IN")
    total_out=sum(float(r["amount"] or 0) for r in rows if r["direction"] in ("OUT","REFUND"))

    con.close()

    return render_template("finance_daily_report.html",
        u=u, banks=banks, rows=rows,
        inflows=inflows, outflows=outflows,
        balance_cf=balance_cf, bank_in=bank_in, bank_out=bank_out,
        total_in=total_in, total_out=total_out,
        sel_date=sel_date, sel_company=sel_company,
        enumerate=enumerate, none_key=None)


@app.route("/finance",methods=["GET","POST"])
@require_roles("MD","GM")
def finance_center():
    u=current_user();con=db();companies=_finance_companies(u)
    if request.method=="POST":
        company=request.form.get("company_code","").strip()
        if company not in companies:
            con.close();flash("Company access denied.","error");return redirect(url_for("finance_center"))
        direction=request.form.get("direction","IN").strip()
        if direction not in ("IN","OUT","REFUND"): direction="OUT"
        txn_date=request.form.get("txn_date") or datetime.date.today().isoformat()
        amount=float(request.form.get("amount") or 0)
        if amount<=0:
            con.close();flash("Amount must be greater than zero.","error");return redirect(url_for("finance_center"))
        voucher=_finance_voucher(con,company,direction,txn_date)
        bill=request.files.get("bill");bn="";bm="";bb=None
        if bill and bill.filename:
            bn=secure_filename(bill.filename)[:180];bb=bill.read()
            if len(bb)>10*1024*1024:
                con.close();flash("Bill attachment must be 10 MB or smaller.","error");return redirect(url_for("finance_center"))
            bm=bill.mimetype or "application/octet-stream"
        aid=request.form.get("account_id","").strip()
        aid=int(aid) if aid.isdigit() else None
        con.execute("""INSERT INTO finance_transactions(voucher_no,txn_date,company_code,direction,head,party_name,client_ref,amount,payment_mode,account_id,reference_no,taxable_value,gst_amount,remarks,bill_name,bill_mime,bill_bytes,status,created_by_id,created_by_name,created_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'POSTED',?,?,?)""",(voucher,txn_date,company,direction,request.form.get("head","Miscellaneous Expense"),request.form.get("party_name","").strip(),request.form.get("client_ref","").strip(),amount,request.form.get("payment_mode","Cash"),aid,request.form.get("reference_no","").strip(),float(request.form.get("taxable_value") or 0),float(request.form.get("gst_amount") or 0),request.form.get("remarks","").strip(),bn,bm,bb,u["id"],u["full_name"],datetime.datetime.now().isoformat(timespec="seconds")))
        con.commit();con.close()
        log_activity("FINANCE_TRANSACTION","Finance Transaction Posted","FINANCE","Voucher",voucher,voucher,details={"direction":direction,"head":request.form.get("head"),"amount":amount,"party":request.form.get("party_name")},severity="WARNING" if direction=="REFUND" else "INFO",actor=u,company_code=company)
        flash("Transaction saved: "+voucher,"success");return redirect(url_for("finance_center"))
    df=request.args.get("date_from","");dt=request.args.get("date_to","");company=request.args.get("company","");head=request.args.get("head","")
    wh=["t.company_code IN ("+",".join(["?"]*len(companies))+")"];pa=list(companies)
    if df:wh.append("t.txn_date>=?");pa.append(df)
    if dt:wh.append("t.txn_date<=?");pa.append(dt)
    if company in companies:wh.append("t.company_code=?");pa.append(company)
    if head:wh.append("t.head=?");pa.append(head)
    rows=con.execute("""SELECT t.*,b.bank_name,b.account_name,b.account_last4 FROM finance_transactions t LEFT JOIN finance_banks b ON b.id=t.account_id WHERE """+" AND ".join(wh)+" ORDER BY t.txn_date DESC,t.id DESC",pa).fetchall()
    banks=con.execute("SELECT * FROM finance_banks WHERE active=1 AND company_code IN ("+",".join(["?"]*len(companies))+") ORDER BY company_code,bank_name",companies).fetchall()
    vals=[dict(x) for x in rows];income=sum(x["amount"] for x in vals if x["direction"]=="IN");expense=sum(x["amount"] for x in vals if x["direction"]=="OUT");refund=sum(x["amount"] for x in vals if x["direction"]=="REFUND")
    k={"income":income,"expense":expense,"refund":refund,"net":income-expense-refund,"count":len(vals),"bill_pending":sum(1 for x in vals if x["direction"]!="IN" and not x["bill_name"])}
    con.close()
    return render_template("finance.html",u=u,rows=rows,banks=banks,companies=companies,heads=FINANCE_HEADS,k=k,today=datetime.date.today().isoformat(),f={"date_from":df,"date_to":dt,"company":company,"head":head})

@app.route("/finance/banks",methods=["GET","POST"])
@require_roles("MD","GM")
def finance_banks():
    ensure_bank_manager_schema()
    u=current_user();con=db();companies=_finance_companies(u)
    if request.method=="POST":
        company=request.form.get("company_code") or u["company_code"]
        if company not in companies:
            con.close();flash("You can add bank accounts only for your authorized company.","error");return redirect(url_for("finance_banks"))
        num=request.form.get("account_number","").strip()
        bank_name=request.form.get("bank_name","").strip()
        account_name=request.form.get("account_name","").strip()
        if not bank_name or not account_name or not num:
            con.close();flash("Bank Name, Account Name and Account Number are required.","error");return redirect(url_for("finance_banks"))
        con.execute("""INSERT INTO finance_banks(company_code,bank_name,account_name,account_number,account_last4,ifsc_code,branch,account_type,opening_balance,created_by,created_at,upi_id,city,rm_name,rm_mobile,account_nickname,share_enabled)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(company,bank_name,account_name,num,num[-4:],request.form.get("ifsc_code","").strip().upper(),request.form.get("branch","").strip(),request.form.get("account_type","Current"),float(request.form.get("opening_balance") or 0),u["full_name"],datetime.datetime.now().isoformat(timespec="seconds"),request.form.get("upi_id","").strip(),request.form.get("city","").strip(),request.form.get("rm_name","").strip(),request.form.get("rm_mobile","").strip(),request.form.get("account_nickname","").strip(),1))
        con.commit();con.close()
        log_activity("BANK_ACCOUNT_ADDED","Bank / Cash Account Added","FINANCE",
                     details={"company":company,"bank":bank_name,"last4":num[-4:]},actor=u,company_code=company)
        flash("Bank account added successfully.","success");return redirect(url_for("finance_banks"))
    rows=con.execute("SELECT * FROM finance_banks WHERE active=1 AND company_code IN ("+",".join(["?"]*len(companies))+") ORDER BY company_code,bank_name",companies).fetchall();con.close()
    return render_template("finance_banks.html",u=u,banks=rows)



@app.route("/finance/banks/delete/<int:bank_id>",methods=["POST"])
@require_roles("MD","GM")
def finance_bank_delete(bank_id):
    ensure_bank_manager_schema()
    u=current_user(); con=db(); companies=_finance_companies(u)
    row=con.execute("SELECT * FROM finance_banks WHERE id=? AND active=1",(bank_id,)).fetchone()
    if not row or row["company_code"] not in companies:
        con.close(); flash("Bank account not found or access denied.","error")
        return redirect(url_for("finance_banks"))

    # Soft delete keeps historical transactions safe.
    con.execute("UPDATE finance_banks SET active=0,share_enabled=0 WHERE id=?",(bank_id,))
    con.commit(); con.close()
    log_activity("BANK_ACCOUNT_REMOVED","Bank Account Removed","FINANCE","Bank Account",bank_id,
                 row["bank_name"],details={"company":row["company_code"],"last4":row["account_last4"]},
                 severity="WARNING",actor=u,company_code=row["company_code"])
    flash("Bank account removed from active Bank Manager. Historical transactions remain safe.","success")
    return redirect(url_for("finance_banks"))


@app.route("/payment-accounts")
@require_roles("MD","GM")
def payment_accounts():
    ensure_bank_manager_schema()
    u=current_user();con=db()
    companies=_finance_companies(u)
    rows=con.execute("SELECT * FROM finance_banks WHERE active=1 AND COALESCE(share_enabled,1)=1 AND company_code IN ("+",".join(["?"]*len(companies))+") ORDER BY company_code,bank_name",companies).fetchall()
    banks=[]
    for r in rows:
        d=dict(r);d["logo_url"]=_bank_logo_url(d.get("bank_name",""));banks.append(d)
    con.close()
    return render_template("payment_accounts.html",u=u,banks=banks)

@app.route("/payment-accounts/share/<int:bank_id>",methods=["POST"])
@require_roles("MD")
def payment_account_share(bank_id):
    ensure_bank_manager_schema()
    u=current_user();con=db();companies=_finance_companies(u)
    b=con.execute("SELECT * FROM finance_banks WHERE id=? AND active=1 AND COALESCE(share_enabled,1)=1",(bank_id,)).fetchone()
    if not b or b["company_code"] not in companies:
        con.close();return jsonify({"ok":False,"error":"Account unavailable"}),403
    mobile=re.sub(r"\\D","",request.form.get("mobile",""))
    if len(mobile)==10:mobile="91"+mobile
    company_name="ABC Immigration Consultants" if b["company_code"]=="SCIC" else "XYZ"
    msg=_payment_message(dict(b),company_name)
    con.execute("INSERT INTO bank_share_logs(bank_id,company_code,shared_by_id,shared_by_name,recipient_mobile,share_type,shared_at) VALUES(?,?,?,?,?,'WHATSAPP',?)",
                (b["id"],b["company_code"],u["id"],u["full_name"],mobile,datetime.datetime.now().isoformat(timespec="seconds")))
    con.commit();con.close()
    log_activity("BANK_DETAILS_SHARED","Approved Bank Details Shared","FINANCE","Bank Account",bank_id,b["bank_name"],
                 details={"recipient_mobile":mobile,"account_last4":b["account_last4"]},actor=u,company_code=b["company_code"])
    wa=("https://wa.me/"+mobile+"?text="+quote(msg)) if mobile else ("https://wa.me/?text="+quote(msg))
    return jsonify({"ok":True,"whatsapp_url":wa,"message":msg})

@app.route("/finance/banks/shares")
@require_roles("MD")
def bank_share_history():
    ensure_bank_manager_schema()
    u=current_user();con=db();companies=_finance_companies(u)
    rows=con.execute("""SELECT l.*,b.bank_name,b.account_name,b.account_last4 FROM bank_share_logs l
                       LEFT JOIN finance_banks b ON b.id=l.bank_id
                       WHERE l.company_code IN ("""+",".join(["?"]*len(companies))+") ORDER BY l.id DESC LIMIT 1000",companies).fetchall()
    con.close();return render_template("bank_share_history.html",u=u,rows=rows)

@app.route("/finance/bill/<int:txn_id>")
@require_roles("MD")
def finance_bill(txn_id):
    u=current_user();con=db();r=con.execute("SELECT * FROM finance_transactions WHERE id=?",(txn_id,)).fetchone()
    if not r or r["company_code"] not in _finance_companies(u) or not r["bill_bytes"]:
        con.close();return "Bill not found",404
    data=bytes(r["bill_bytes"]);name=r["bill_name"];mime=r["bill_mime"];con.close()
    return send_file(io.BytesIO(data),mimetype=mime,download_name=name,as_attachment=False)

@app.route("/finance/receipt/<int:txn_id>")
@require_roles("MD")
def finance_receipt(txn_id):
    u=current_user();con=db();r=con.execute("""SELECT t.*,b.bank_name,b.account_last4 FROM finance_transactions t LEFT JOIN finance_banks b ON b.id=t.account_id WHERE t.id=?""",(txn_id,)).fetchone()
    if not r or r["direction"]!="IN":con.close();return "Receipt not found",404
    if u["role"]!="MD" and r["company_code"]!=u["company_code"]:con.close();return "Access denied",403
    company_name="ABC Immigration Consultants" if r["company_code"]=="SCIC" else "XYZ"
    address="SCO 411-412, 2nd Floor, Sector 35-C, Chandigarh – 160035" if r["company_code"]=="SCIC" else "Company address to be configured"
    gstin=os.environ.get("GSTIN_"+r["company_code"],"")
    con.close();return render_template("finance_receipt.html",u=u,r=r,company_name=company_name,company_address=address,gstin=gstin)

@app.route("/finance/whatsapp/<int:txn_id>")
@require_roles("MD")
def finance_whatsapp(txn_id):
    u=current_user();con=db();r=con.execute("SELECT * FROM finance_transactions WHERE id=?",(txn_id,)).fetchone();con.close()
    if not r or r["direction"]!="IN":return "Receipt not found",404
    if u["role"]!="MD" and r["company_code"]!=u["company_code"]:return "Access denied",403
    text=f"Payment Receipt {r['voucher_no']} | {r['company_code']} | Received ₹{r['amount']:.2f} from {r['party_name']} on {r['txn_date']}. Please open the secure portal receipt link shared by our team."
    import urllib.parse
    return redirect("https://wa.me/?text="+urllib.parse.quote(text))

@app.route("/finance/export.xlsx")
@require_roles("MD")
def finance_export():
    u=current_user();con=db();companies=_finance_companies(u);df=request.args.get("date_from","");dt=request.args.get("date_to","");company=request.args.get("company","")
    wh=["t.company_code IN ("+",".join(["?"]*len(companies))+")"];pa=list(companies)
    if df:wh.append("t.txn_date>=?");pa.append(df)
    if dt:wh.append("t.txn_date<=?");pa.append(dt)
    if company in companies:wh.append("t.company_code=?");pa.append(company)
    rows=con.execute("""SELECT t.*,b.bank_name,b.account_name,b.account_last4,b.ifsc_code,b.branch FROM finance_transactions t LEFT JOIN finance_banks b ON b.id=t.account_id WHERE """+" AND ".join(wh)+" ORDER BY t.txn_date,t.id",pa).fetchall();con.close()
    wb=Workbook();ws=wb.active;ws.title="All Transactions"
    headers=["Voucher","Date","Company","Direction","Head","Party / Client","Client Ref","Amount","Payment Mode","Bank","Account Name","Account Last4","IFSC","Branch","Reference","Taxable Value","GST Amount","Remarks","Bill Status","Created By","Created At"]
    ws.append(headers)
    for r in rows:ws.append([r["voucher_no"],r["txn_date"],r["company_code"],r["direction"],r["head"],r["party_name"],r["client_ref"],r["amount"],r["payment_mode"],r["bank_name"] or "Cash",r["account_name"] or "",r["account_last4"] or "",r["ifsc_code"] or "",r["branch"] or "",r["reference_no"],r["taxable_value"],r["gst_amount"],r["remarks"],"Attached" if r["bill_name"] else "Pending",r["created_by_name"],r["created_at"]])
    for cell in ws[1]:cell.font=Font(bold=True)
    ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions
    for col in ws.columns:
        letter=col[0].column_letter;ws.column_dimensions[letter].width=min(max(12,max(len(str(c.value or "")) for c in col)+2),35)
    # modular sheets
    data=[dict(r) for r in rows]
    specs=[("Money Received",lambda x:x["direction"]=="IN"),("Money Paid",lambda x:x["direction"]=="OUT"),("Client Refunds",lambda x:x["direction"]=="REFUND"),("Embassy VFS",lambda x:x["head"] in ("Embassy Fee","VFS / Biometric Fee")),("Hotel Travel",lambda x:x["head"] in ("Hotel Booking","Flight / Ticket","Travel / Fuel")),("Funds",lambda x:x["head"]=="Funds Arrangement")]
    for title,fn in specs:
        sh=wb.create_sheet(title);sh.append(headers)
        for c in sh[1]:c.font=Font(bold=True)
        for r in data:
            if fn(r):sh.append([r["voucher_no"],r["txn_date"],r["company_code"],r["direction"],r["head"],r["party_name"],r["client_ref"],r["amount"],r["payment_mode"],r["bank_name"] or "Cash",r["account_name"] or "",r["account_last4"] or "",r["ifsc_code"] or "",r["branch"] or "",r["reference_no"],r["taxable_value"],r["gst_amount"],r["remarks"],"Attached" if r["bill_name"] else "Pending",r["created_by_name"],r["created_at"]])
        sh.freeze_panes="A2";sh.auto_filter.ref=sh.dimensions
    sm=wb.create_sheet("Summary",0);sm.append(["THE GAUR - FINANCE EXPORT"]);sm.append(["Period",df or "All",dt or "All"]);sm.append(["Company",company or ("Both" if u["role"]=="MD" else u["company_code"])]);sm.append([]);sm.append(["Metric","Amount"])
    income=sum(x["amount"] for x in data if x["direction"]=="IN");out=sum(x["amount"] for x in data if x["direction"]=="OUT");ref=sum(x["amount"] for x in data if x["direction"]=="REFUND")
    for a,b in [("Money Received",income),("Money Paid",out),("Client Refunds",ref),("Net Cash Flow",income-out-ref)]:sm.append([a,b])
    bio=io.BytesIO();wb.save(bio);bio.seek(0)
    log_activity("FINANCE_EXPORT","Finance Excel Exported","FINANCE",details={"from":df,"to":dt,"company":company or "scope"},actor=u)
    return send_file(bio,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",download_name="THE_GAUR_Finance_Export.xlsx",as_attachment=True)

@app.route("/activity-center")
@require_roles("MD","GM")
def activity_center():
    u=current_user();con=db()
    df=request.args.get("date_from","").strip();dt=request.args.get("date_to","").strip();company=request.args.get("company","").strip();uid=request.args.get("user_id","").strip();category=request.args.get("category","").strip();severity=request.args.get("severity","").strip();q=request.args.get("q","").strip()
    where=[];params=[]
    if u["role"]=="GM": where.append("company_code=?");params.append(u["company_code"]);scope_label=u["company_code"]+" Team"
    else:
        scope_label="Both Companies"
        if company in ("SCIC","WWIC"): where.append("company_code=?");params.append(company)
    if df: where.append("created_at>=?");params.append(df+"T00:00:00")
    if dt: where.append("created_at<=?");params.append(dt+"T23:59:59")
    if uid.isdigit(): where.append("user_id=?");params.append(int(uid))
    if category: where.append("category=?");params.append(category)
    if severity: where.append("severity=?");params.append(severity)
    if q:
        like="%"+q+"%";where.append("(user_name LIKE ? OR action_label LIKE ? OR details LIKE ? OR entity_label LIKE ? OR entity_id LIKE ?)");params.extend([like]*5)
    ws=(" WHERE "+" AND ".join(where)) if where else ""
    rows=con.execute("SELECT * FROM activity_logs"+ws+" ORDER BY id DESC LIMIT 1200",params).fetchall()
    today=datetime.date.today().isoformat();kw=["created_at>=?"];kp=[today+"T00:00:00"]
    if u["role"]=="GM":kw.append("company_code=?");kp.append(u["company_code"])
    def count(extra):return con.execute("SELECT COUNT(*) c FROM activity_logs WHERE "+" AND ".join(kw)+" AND "+extra,kp).fetchone()["c"]
    kpi={"logins":count("action_code='LOGIN_SUCCESS'"),"logouts":count("action_code='LOGOUT'"),"client":count("category IN ('CLIENT','LEAD')"),"chat":count("category='CHAT'"),"admin":count("category IN ('ADMIN','EMPLOYEE')"),"critical":count("severity='CRITICAL'")}
    pq="SELECT * FROM users WHERE active=1";pp=[]
    if u["role"]=="GM":pq+=" AND company_code=?";pp.append(u["company_code"])
    people=con.execute(pq+" ORDER BY company_code,full_name",pp).fetchall();categories=[x["category"] for x in con.execute("SELECT DISTINCT category FROM activity_logs ORDER BY category").fetchall()]
    con.close();f={"date_from":df,"date_to":dt,"company":company,"user_id":uid,"category":category,"severity":severity,"q":q}
    return render_template("activity_center.html",u=u,rows=rows,people=people,categories=categories,kpi=kpi,scope_label=scope_label,f=f)


@app.route("/attendance",methods=["GET","POST"])
@require_roles("MD","GM","RECEPTION")
def attendance_center():
    u=current_user(); con=db()
    today=datetime.date.today().isoformat()
    now_time=datetime.datetime.now().strftime("%H:%M")
    company=(request.form.get("company_code") or request.args.get("company") or (u["company_code"] if u["role"]!="MD" else "SCIC")).strip()
    if u["role"]!="MD":
        company=u["company_code"]
    if company not in ("SCIC","WWIC"):
        company="SCIC"

    if request.method=="POST" and request.form.get("save_attendance")=="1":
        eid=request.form.get("employee_id","")
        adate=request.form.get("attendance_date") or today
        etime=request.form.get("entry_time") or now_time
        remarks=request.form.get("remarks","").strip()
        emp=con.execute("SELECT * FROM employee_master WHERE id=? AND company_code=? AND employment_status='Active'",(int(eid) if eid.isdigit() else -1,company)).fetchone()
        if not emp:
            flash("Select a valid active employee.","error")
        else:
            existing=con.execute("SELECT id FROM employee_attendance WHERE attendance_date=? AND employee_id=?",(adate,emp["id"])).fetchone()
            stamp=datetime.datetime.now().isoformat(timespec="seconds")
            if existing:
                con.execute("UPDATE employee_attendance SET entry_time=?,remarks=?,marked_by=?,marked_at=? WHERE id=?",(etime,remarks,u["login_id"],stamp,existing["id"]))
            else:
                con.execute("""INSERT INTO employee_attendance(attendance_date,employee_id,company_code,entry_time,remarks,marked_by,marked_at)
                               VALUES(?,?,?,?,?,?,?)""",(adate,emp["id"],company,etime,remarks,u["login_id"],stamp))
            con.commit()
            log_activity("EMPLOYEE_ATTENDANCE_MARKED","Employee Attendance Marked","HR","Employee",emp["id"],emp["full_name"],
                         details={"date":adate,"entry_time":etime,"remarks":remarks},actor=u,company_code=company)
            flash(f"Attendance saved for {emp['full_name']} at {etime}.","success")

    employees=con.execute("""SELECT * FROM employee_master WHERE company_code=? AND employment_status='Active'
                             ORDER BY full_name""",(company,)).fetchall()
    raw=con.execute("""SELECT a.*,e.full_name,e.designation,e.mobile
                       FROM employee_attendance a JOIN employee_master e ON e.id=a.employee_id
                       WHERE a.company_code=? AND a.attendance_date=? ORDER BY a.entry_time,e.full_name""",(company,today)).fetchall()
    company_name="ABC Immigration Consultants" if company=="SCIC" else "XYZ"
    rows=[]
    for r in raw:
        d=dict(r)
        mobile=re.sub(r"\D","",d.get("mobile") or "")
        if len(mobile)==10: mobile="91"+mobile
        elif mobile.startswith("0") and len(mobile)==11: mobile="91"+mobile[1:]
        msg=f"Good Morning {d['full_name']}! Welcome to office. Your attendance has been marked at {d['entry_time']} on {today}. Have a productive day! - {company_name}"
        d["whatsapp_url"]="https://wa.me/"+mobile+"?text="+quote(msg) if mobile else ""
        rows.append(d)
    marked_ids={int(r["employee_id"]) for r in rows}
    missing=[e for e in employees if int(e["id"]) not in marked_ids]
    con.close()
    return render_template("attendance.html",u=u,company=company,today=today,now_time=now_time,employees=employees,rows=rows,missing=missing,company_name=company_name)

@app.route("/reception", methods=["GET","POST"])
@require_roles("MD","GM","RECEPTION")
def reception():
    u=current_user()
    con=db()
    if request.method=="POST":
        company=request.form.get("company_code") if u["role"]=="MD" else u["company_code"]
        vid=f"{company}-VIS-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
        con.execute("""INSERT INTO visitors(visitor_id,company_code,visit_date,client_name,mobile,email,city,country,visa_type,coordinating_with,purpose,status,remarks,created_by)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (vid,company,request.form.get("visit_date"),request.form.get("client_name"),request.form.get("mobile"),
                     request.form.get("email"),request.form.get("city"),request.form.get("country"),request.form.get("visa_type"),
                     request.form.get("coordinating_with"),request.form.get("purpose"),request.form.get("status"),request.form.get("remarks"),u["login_id"]))
        con.commit()
        flash("Visitor saved","success")
    if u["role"]=="MD":
        rows=con.execute("SELECT * FROM visitors ORDER BY id DESC LIMIT 100").fetchall()
    else:
        rows=con.execute("SELECT * FROM visitors WHERE company_code=? ORDER BY id DESC LIMIT 100",(u["company_code"],)).fetchall()
    con.close()
    return render_template("reception.html",u=u,rows=rows)

def normalize_header(x):
    return re.sub(r'[^a-z0-9]+','',str(x or '').lower())

FIELD_ALIASES = {
    "client_name":{"name","clientname","fullname","leadname","customername"},
    "mobile":{"mobile","mobileno","mobilephone","phone","phoneno","contact","contactno","contactnumber"},
    "email":{"email","emailid","emailaddress"},
    "city":{"city","place","location","district"},
    "country":{"country","interestedcountry","destination","destinationcountry"},
    "visa_type":{"visa","visatype","product","category"},
    "source":{"source","campaign","leadsource","area","state"}
}

def map_headers(headers):
    result={}
    for i,h in enumerate(headers):
        n=normalize_header(h)
        for field,aliases in FIELD_ALIASES.items():
            if n in aliases and field not in result:
                result[field]=i
    return result

def read_excel(path):
    """Read real XLS/XLSX and Excel 2003 XML files, even when extension is misleading."""
    with open(path, "rb") as fh:
        head = fh.read(512)

    # Modern .xlsx is a ZIP container. Detect by content, not only filename.
    if head.startswith(b"PK") or zipfile.is_zipfile(path):
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        return list(ws.iter_rows(values_only=True))

    # Legacy binary .xls uses the OLE Compound File signature.
    if head.startswith(bytes.fromhex("D0CF11E0A1B11AE1")):
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_index(0)
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]

    # Some banking/CRM exports are Excel 2003 XML but are saved with .xls extension.
    stripped = head.lstrip(b"\xef\xbb\xbf\x00\t\r\n ")
    if stripped.startswith(b"<?xml") or stripped.startswith(b"<Workbook"):
        try:
            root = ET.parse(path).getroot()
            rows = []
            for row in root.findall('.//{*}Row'):
                values = []
                col = 1
                for cell in row.findall('{*}Cell'):
                    idx = cell.attrib.get('{urn:schemas-microsoft-com:office:spreadsheet}Index')
                    if idx:
                        target = int(idx)
                        while col < target:
                            values.append(None); col += 1
                    data = cell.find('{*}Data')
                    values.append(data.text if data is not None and data.text is not None else '')
                    col += 1
                rows.append(values)
            if rows:
                return rows
        except ET.ParseError as e:
            raise ValueError(f"Excel XML file could not be read: {e}")

    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xls", ".xlsx"):
        raise ValueError("Only .xls and .xlsx are supported")
    raise ValueError("Unsupported or damaged Excel file. Please open it in Excel and Save As .xlsx, then upload again.")



@app.route("/lead-daily-report")
@require_roles("MD","GM")
def lead_daily_report():
    u=current_user()
    report_date=(request.args.get("date") or datetime.date.today().isoformat()).strip()
    company=(request.args.get("company") or "").strip().upper()
    if u["role"]=="GM":
        company=u["company_code"]
    if company not in ("","SCIC","WWIC"):
        company=""

    con=db()
    start=report_date+"T00:00:00"
    end=report_date+"T23:59:59"

    # Upload summary. "New Leads" = actual lead records created/imported on the selected day.
    csql=""; cparams=[]
    if company:
        csql=" AND company_code=?"; cparams=[company]

    company_rows=[]
    company_codes=[company] if company else ["SCIC","WWIC"]
    for cc in company_codes:
        actual=con.execute("""SELECT COUNT(*) c FROM leads
                              WHERE COALESCE(imported_at,'') BETWEEN ? AND ?
                                AND company_code=?""",(start,end,cc)).fetchone()["c"]
        active=con.execute("""SELECT COUNT(*) c FROM leads
                              WHERE COALESCE(imported_at,'') BETWEEN ? AND ?
                                AND company_code=? AND COALESCE(deleted_at,'')=''""",(start,end,cc)).fetchone()["c"]
        in_bin=int(actual or 0)-int(active or 0)

        # Import file stats provide uploaded rows / duplicate picture.
        imp=con.execute("""SELECT
                            COALESCE(SUM(total_rows),0) total_rows,
                            COALESCE(SUM(new_rows),0) new_rows,
                            COALESCE(SUM(same_company_duplicates),0) same_dup,
                            COALESCE(SUM(cross_company_duplicates),0) cross_dup,
                            COALESCE(SUM(invalid_mobiles),0) invalid,
                            COUNT(*) files
                           FROM imports
                           WHERE company_code=? AND COALESCE(imported_at,'') BETWEEN ? AND ?""",
                        (cc,start,end)).fetchone()
        company_rows.append({
            "company_code":cc,
            "company_name":"ABC" if cc=="SCIC" else "XYZ",
            "uploaded_rows":int(imp["total_rows"] or 0),
            "new_leads":int(actual or 0),
            "active_new_leads":int(active or 0),
            "moved_to_bin":int(in_bin or 0),
            "same_dup":int(imp["same_dup"] or 0),
            "cross_dup":int(imp["cross_dup"] or 0),
            "invalid":int(imp["invalid"] or 0),
            "files":int(imp["files"] or 0),
        })

    # AM-wise allocations based on actual allocation_history actions of the selected day.
    q="""SELECT ah.company_code,ah.am_user_id,u2.full_name,
                COALESCE(SUM(ah.quantity),0) allocated_qty,
                COUNT(ah.id) allocation_actions
         FROM allocation_history ah
         LEFT JOIN users u2 ON u2.id=ah.am_user_id
         WHERE COALESCE(ah.allocated_at,'') BETWEEN ? AND ?"""
    params=[start,end]
    if company:
        q+=" AND ah.company_code=?"; params.append(company)
    q+=" GROUP BY ah.company_code,ah.am_user_id,u2.full_name ORDER BY ah.company_code,allocated_qty DESC,u2.full_name"
    allocation_rows=[dict(r) for r in con.execute(q,params).fetchall()]

    # Current leads whose latest assigned_at falls on selected date.
    q2="""SELECT l.company_code,l.assigned_am,u2.full_name,COUNT(*) current_qty
          FROM leads l LEFT JOIN users u2 ON u2.id=l.assigned_am
          WHERE COALESCE(l.assigned_at,'') BETWEEN ? AND ?
            AND COALESCE(l.deleted_at,'')=''"""
    p2=[start,end]
    if company:
        q2+=" AND l.company_code=?";p2.append(company)
    q2+=" GROUP BY l.company_code,l.assigned_am,u2.full_name ORDER BY l.company_code,current_qty DESC"
    current_alloc=[dict(r) for r in con.execute(q2,p2).fetchall()]

    # Import file detail.
    q3="""SELECT company_code,filename,total_rows,new_rows,same_company_duplicates,
                 cross_company_duplicates,invalid_mobiles,imported_by,imported_at
          FROM imports WHERE COALESCE(imported_at,'') BETWEEN ? AND ?"""
    p3=[start,end]
    if company:
        q3+=" AND company_code=?";p3.append(company)
    q3+=" ORDER BY imported_at DESC"
    import_files=[dict(r) for r in con.execute(q3,p3).fetchall()]
    con.close()

    totals={
        "uploaded_rows":sum(x["uploaded_rows"] for x in company_rows),
        "new_leads":sum(x["new_leads"] for x in company_rows),
        "active_new_leads":sum(x["active_new_leads"] for x in company_rows),
        "moved_to_bin":sum(x["moved_to_bin"] for x in company_rows),
        "allocated":sum(int(x["allocated_qty"] or 0) for x in allocation_rows),
    }

    wa=[
        "IMS • DAILY LEAD REPORT",
        "Date: "+report_date,
        "Scope: "+(("ABC" if company=="SCIC" else "XYZ") if company else "Both Companies"),
        ""
    ]
    for x in company_rows:
        wa.append(f"{x['company_name']}: Uploaded Rows {x['uploaded_rows']} | New Leads {x['new_leads']} | Active {x['active_new_leads']} | Bin {x['moved_to_bin']}")
    wa.append("")
    wa.append("AM-WISE ALLOCATION:")
    if allocation_rows:
        for x in allocation_rows:
            wa.append(f"{x['company_code']} • {x['full_name'] or 'Unknown AM'}: {int(x['allocated_qty'] or 0)} leads")
    else:
        wa.append("No allocation recorded on this date.")
    wa.append("")
    wa.append(f"TOTAL NEW LEADS: {totals['new_leads']} | TOTAL ALLOCATED: {totals['allocated']}")
    wa_text="\n".join(wa)

    return render_template("lead_daily_report.html",u=u,report_date=report_date,company=company,
                           company_rows=company_rows,allocation_rows=allocation_rows,
                           current_alloc=current_alloc,import_files=import_files,totals=totals,
                           wa_text=wa_text)


@app.route("/lead-daily-report/export.xlsx")
@require_roles("MD","GM")
def lead_daily_report_export():
    u=current_user()
    report_date=(request.args.get("date") or datetime.date.today().isoformat()).strip()
    company=(request.args.get("company") or "").strip().upper()
    if u["role"]=="GM": company=u["company_code"]
    if company not in ("","SCIC","WWIC"): company=""
    start=report_date+"T00:00:00"; end=report_date+"T23:59:59"
    con=db()

    wb=Workbook()
    ws=wb.active; ws.title="Daily Lead Summary"
    headers=["Company","Uploaded Rows","New Leads","Active New Leads","Moved to Bin"]
    for c,h in enumerate(headers,1):
        ws.cell(1,c,h).font=Font(bold=True)
    r=2
    for cc in ([company] if company else ["SCIC","WWIC"]):
        imp=con.execute("""SELECT COALESCE(SUM(total_rows),0) total_rows FROM imports
                           WHERE company_code=? AND COALESCE(imported_at,'') BETWEEN ? AND ?""",(cc,start,end)).fetchone()
        newc=con.execute("""SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(imported_at,'') BETWEEN ? AND ?""",(cc,start,end)).fetchone()["c"]
        active=con.execute("""SELECT COUNT(*) c FROM leads WHERE company_code=? AND COALESCE(imported_at,'') BETWEEN ? AND ? AND COALESCE(deleted_at,'')=''""",(cc,start,end)).fetchone()["c"]
        vals=[cc,int(imp["total_rows"] or 0),int(newc or 0),int(active or 0),int(newc or 0)-int(active or 0)]
        for c,v in enumerate(vals,1):ws.cell(r,c,v)
        r+=1

    wa=wb.create_sheet("AM Allocation")
    for c,h in enumerate(["Company","AM Name","Allocated Qty","Allocation Actions"],1):
        wa.cell(1,c,h).font=Font(bold=True)
    q="""SELECT ah.company_code,u2.full_name,COALESCE(SUM(ah.quantity),0) qty,COUNT(ah.id) actions
         FROM allocation_history ah LEFT JOIN users u2 ON u2.id=ah.am_user_id
         WHERE COALESCE(ah.allocated_at,'') BETWEEN ? AND ?"""
    p=[start,end]
    if company:q+=" AND ah.company_code=?";p.append(company)
    q+=" GROUP BY ah.company_code,ah.am_user_id,u2.full_name ORDER BY ah.company_code,qty DESC"
    rows=con.execute(q,p).fetchall()
    for i,x in enumerate(rows,2):
        wa.cell(i,1,x["company_code"]);wa.cell(i,2,x["full_name"] or "Unknown AM")
        wa.cell(i,3,int(x["qty"] or 0));wa.cell(i,4,int(x["actions"] or 0))

    wf=wb.create_sheet("Import Files")
    hdr=["Company","File","Total Rows","New Rows","Same Dup","Cross Dup","Invalid","Uploaded By","Uploaded At"]
    for c,h in enumerate(hdr,1):wf.cell(1,c,h).font=Font(bold=True)
    qf="""SELECT company_code,filename,total_rows,new_rows,same_company_duplicates,cross_company_duplicates,
                 invalid_mobiles,imported_by,imported_at FROM imports
          WHERE COALESCE(imported_at,'') BETWEEN ? AND ?"""
    pf=[start,end]
    if company:qf+=" AND company_code=?";pf.append(company)
    qf+=" ORDER BY imported_at DESC"
    for i,x in enumerate(con.execute(qf,pf).fetchall(),2):
        vals=[x["company_code"],x["filename"],x["total_rows"],x["new_rows"],x["same_company_duplicates"],
              x["cross_company_duplicates"],x["invalid_mobiles"],x["imported_by"],x["imported_at"]]
        for c,v in enumerate(vals,1):wf.cell(i,c,v)

    con.close()
    for sh in (ws,wa,wf):
        for col in range(1,sh.max_column+1):
            sh.column_dimensions[get_column_letter(col)].width=20
        sh.freeze_panes="A2"
    bio=io.BytesIO();wb.save(bio);bio.seek(0)
    scope=company or "BOTH"
    return send_file(bio,mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True,download_name=f"Daily_Lead_Report_{scope}_{report_date}.xlsx")


@app.route("/lead-control")
@require_roles("MD","GM")
def lead_control_center():
    u=current_user(); con=db()
    company=request.args.get("company","").strip()
    batch=request.args.get("batch","").strip()
    am=request.args.get("am","").strip()
    area=request.args.get("area","").strip()
    location=request.args.get("location","").strip()
    work=request.args.get("work","").strip()
    search=request.args.get("q","").strip()
    upload_from=request.args.get("upload_from","").strip()
    upload_to=request.args.get("upload_to","").strip()
    alloc_from=request.args.get("alloc_from","").strip()
    alloc_to=request.args.get("alloc_to","").strip()

    latest_rows=con.execute("""SELECT company_code,batch_code FROM imports
        WHERE COALESCE(batch_code,'')<>'' ORDER BY company_code,id DESC""").fetchall()
    latest_by_company={}
    for z in latest_rows:
        if z["company_code"] not in latest_by_company:
            latest_by_company[z["company_code"]]=z["batch_code"]

    clauses=["COALESCE(l.deleted_at,'')=''"]; p=[]
    if u["role"]=="GM":
        clauses.append("l.company_code=?"); p.append(u["company_code"])
    elif company in ("SCIC","WWIC"):
        clauses.append("l.company_code=?"); p.append(company)

    scope_latest=[v for k,v in latest_by_company.items() if (u["role"]=="MD" and (not company or k==company)) or (u["role"]=="GM" and k==u["company_code"])]
    if batch=="legacy":
        clauses.append("COALESCE(l.upload_batch,'')=''")
    elif batch=="latest":
        if scope_latest:
            clauses.append("l.upload_batch IN ("+",".join(["?"]*len(scope_latest))+")"); p.extend(scope_latest)
        else:
            clauses.append("1=0")
    elif batch=="old":
        if scope_latest:
            clauses.append("(COALESCE(l.upload_batch,'')='' OR l.upload_batch NOT IN ("+",".join(["?"]*len(scope_latest))+"))"); p.extend(scope_latest)
    elif batch:
        clauses.append("l.upload_batch=?"); p.append(batch)

    if am=="unallocated":
        clauses.append("l.assigned_am IS NULL")
    elif am.isdigit():
        clauses.append("l.assigned_am=?"); p.append(int(am))
    # v3.98: Area uses imported Source/Area field; Location uses City/Location field.
    if area:
        clauses.append("LOWER(TRIM(COALESCE(l.source,'')))=LOWER(?)"); p.append(area)
    if location:
        clauses.append("LOWER(TRIM(COALESCE(l.city,'')))=LOWER(?)"); p.append(location)
    if upload_from: clauses.append("l.imported_at>=?"); p.append(upload_from+"T00:00:00")
    if upload_to: clauses.append("l.imported_at<=?"); p.append(upload_to+"T23:59:59")
    if alloc_from: clauses.append("COALESCE(l.assigned_at,'')>=?"); p.append(alloc_from+"T00:00:00")
    if alloc_to: clauses.append("COALESCE(l.assigned_at,'')<=?"); p.append(alloc_to+"T23:59:59")
    if search:
        like="%"+search.lower()+"%"; digits=re.sub(r"\D","",search)
        clauses.append("(LOWER(COALESCE(l.client_name,'')) LIKE ? OR COALESCE(l.mobile,'') LIKE ? OR LOWER(COALESCE(l.lead_id,'')) LIKE ? OR LOWER(COALESCE(l.upload_batch,'')) LIKE ?)")
        p.extend([like, "%"+digits+"%" if digits else like, like, like])

    raw=con.execute("""SELECT l.*,u.full_name am_name,
      (SELECT COUNT(*) FROM lead_activity a WHERE a.lead_id=l.id) activity_count,
      (SELECT MAX(a.updated_at) FROM lead_activity a WHERE a.lead_id=l.id) last_activity_at
      FROM leads l LEFT JOIN users u ON u.id=l.assigned_am WHERE """+" AND ".join(clauses)+""" ORDER BY COALESCE(l.interest_score,0) DESC, l.id DESC LIMIT 1500""",p).fetchall()

    today=datetime.date.today(); rows=[]
    for x in raw:
        d=dict(x); cnt=int(d["activity_count"] or 0); over=False
        try: over=bool(d["followup_date"]) and datetime.date.fromisoformat(d["followup_date"])<today
        except: pass
        d["work_state"]="OVERDUE" if over and cnt else ("WORKED" if cnt else "NOT WORKED")
        if not d.get("upload_batch"):
            d["data_age"]="LEGACY"
        elif d["upload_batch"]==latest_by_company.get(d["company_code"]):
            d["data_age"]="LATEST"
        else:
            d["data_age"]="OLD"
        if work=="worked" and not cnt: continue
        if work=="not_worked" and cnt: continue
        if work=="overdue" and not over: continue
        rows.append(d)

    if u["role"]=="MD":
        am_rows=con.execute("SELECT * FROM users WHERE role='AM' AND active=1 ORDER BY company_code,full_name").fetchall()
        batches=con.execute("SELECT * FROM imports WHERE COALESCE(batch_code,'')<>'' ORDER BY id DESC LIMIT 100").fetchall()
    else:
        am_rows=con.execute("SELECT * FROM users WHERE role='AM' AND active=1 AND company_code=? ORDER BY full_name",(u["company_code"],)).fetchall()
        batches=con.execute("SELECT * FROM imports WHERE COALESCE(batch_code,'')<>'' AND company_code=? ORDER BY id DESC LIMIT 100",(u["company_code"],)).fetchall()

    batch_summary=[]
    for b in batches[:30]:
        z=con.execute("""SELECT COUNT(*) total,
          SUM(CASE WHEN assigned_am IS NOT NULL THEN 1 ELSE 0 END) allocated,
          SUM(CASE WHEN assigned_am IS NULL THEN 1 ELSE 0 END) unallocated,
          SUM(CASE WHEN EXISTS(SELECT 1 FROM lead_activity a WHERE a.lead_id=leads.id) THEN 1 ELSE 0 END) worked,
          SUM(CASE WHEN NOT EXISTS(SELECT 1 FROM lead_activity a WHERE a.lead_id=leads.id) THEN 1 ELSE 0 END) not_worked,
          SUM(CASE WHEN COALESCE(interest_score,0)>=50 THEN 1 ELSE 0 END) positive
          FROM leads WHERE upload_batch=? AND COALESCE(deleted_at,'')=''""",(b["batch_code"],)).fetchone()
        d=dict(b); d.update(dict(z)); d["is_latest"]=b["batch_code"]==latest_by_company.get(b["company_code"]); batch_summary.append(d)

    summary=[]
    for a in am_rows:
        z=con.execute("""SELECT COUNT(*) allocated,
        SUM(CASE WHEN EXISTS(SELECT 1 FROM lead_activity x WHERE x.lead_id=l.id) THEN 1 ELSE 0 END) worked,
        SUM(CASE WHEN NOT EXISTS(SELECT 1 FROM lead_activity x WHERE x.lead_id=l.id) THEN 1 ELSE 0 END) not_worked,
        SUM(CASE WHEN COALESCE(l.interest_score,0)>=71 THEN 1 ELSE 0 END) hot,
        SUM(CASE WHEN COALESCE(l.interest_score,0)>=50 THEN 1 ELSE 0 END) positive,
        SUM(CASE WHEN COALESCE(l.followup_date,'')<>'' AND l.followup_date<? THEN 1 ELSE 0 END) overdue
        FROM leads l WHERE l.assigned_am=? AND COALESCE(l.deleted_at,'')=''""",(today.isoformat(),a["id"])).fetchone()
        d=dict(a); d.update(dict(z)); d["work_pct"]=round((int(d["worked"] or 0)/int(d["allocated"] or 1))*100) if d["allocated"] else 0; summary.append(d)

    # Fast filter option lists, scoped by company/role.
    opt_clauses=["COALESCE(deleted_at,'')=''"]; opt_params=[]
    if u["role"]=="GM":
        opt_clauses.append("company_code=?"); opt_params.append(u["company_code"])
    elif company in ("SCIC","WWIC"):
        opt_clauses.append("company_code=?"); opt_params.append(company)
    opt_where=" WHERE "+" AND ".join(opt_clauses)
    area_options=[r["v"] for r in con.execute(
        "SELECT DISTINCT TRIM(source) v FROM leads"+opt_where+" AND COALESCE(TRIM(source),'')<>'' ORDER BY v",
        opt_params).fetchall()]
    location_options=[r["v"] for r in con.execute(
        "SELECT DISTINCT TRIM(city) v FROM leads"+opt_where+" AND COALESCE(TRIM(city),'')<>'' ORDER BY v",
        opt_params).fetchall()]

    archived=con.execute("SELECT * FROM leads WHERE COALESCE(deleted_at,'')<>'' ORDER BY deleted_at DESC LIMIT 1000").fetchall() if u["role"]=="MD" else []
    k={"total":len(rows),
       "latest_batch":sum(1 for r in rows if r["data_age"]=="LATEST"),
       "old_data":sum(1 for r in rows if r["data_age"]!="LATEST"),
       "unallocated":sum(not r["assigned_am"] for r in rows),
       "not_worked":sum(not r["activity_count"] for r in rows),
       "overdue":sum(r["work_state"]=="OVERDUE" for r in rows)}
    con.close()
    return render_template("lead_control.html",u=u,rows=rows,am_rows=am_rows,summary=summary,archived=archived,
                           batches=batches,batch_summary=batch_summary,k=k,
                           area_options=area_options,location_options=location_options,
                           f={"company":company,"batch":batch,"am":am,"area":area,"location":location,"work":work,"q":search,
                              "upload_from":upload_from,"upload_to":upload_to,"alloc_from":alloc_from,"alloc_to":alloc_to})

@app.route("/lead-control/bulk",methods=["POST"])
@require_roles("MD","GM")
def lead_bulk_action():
    u=current_user();con=db();ids=[int(x) for x in request.form.getlist("lead_ids") if x.isdigit()]
    if not ids:con.close();flash("Select at least one lead.","error");return redirect(url_for("lead_control_center"))
    ph=",".join(["?"]*len(ids));p=list(ids);scope=""
    if u["role"]=="GM":scope=" AND company_code=?";p.append(u["company_code"])
    selected=con.execute(f"SELECT * FROM leads WHERE id IN ({ph}) AND COALESCE(deleted_at,'')=''"+scope,p).fetchall()
    action=request.form.get("action");reason=request.form.get("reason","").strip();now=datetime.datetime.now().isoformat(timespec="seconds")
    if action=="allocate":
        aid=request.form.get("am_id","");am=con.execute("SELECT * FROM users WHERE id=? AND role='AM' AND active=1",(int(aid) if aid.isdigit() else -1,)).fetchone()
        if not am or (u["role"]=="GM" and am["company_code"]!=u["company_code"]):con.close();flash("Choose an authorized Assistant Manager.","error");return redirect(url_for("lead_control_center"))
        valid=[r for r in selected if r["company_code"]==am["company_code"]]
        for r in valid:con.execute("UPDATE leads SET assigned_am=?,assigned_at=?,assigned_by=? WHERE id=?",(am["id"],now,u["login_id"],r["id"]))
        con.execute("INSERT INTO allocation_history(company_code,am_user_id,quantity,allocated_by,allocated_at) VALUES(?,?,?,?,?)",(am["company_code"],am["id"],len(valid),u["login_id"],now));con.commit();con.close()
        log_activity("LEADS_BULK_ALLOCATED","Bulk Lead Allocation / Reallocation","LEADS","Lead Batch","",str(len(valid)),details={"assistant_manager":am["full_name"],"quantity":len(valid),"reason":reason},actor=u,company_code=am["company_code"])
        flash(f"{len(valid)} selected leads allocated/re-allocated to {am['full_name']}.","success")
    elif action=="archive":
        delete_reason=reason or ("Moved to Bin by "+u["login_id"])
        for r in selected:
            con.execute("UPDATE leads SET deleted_at=?,deleted_by=?,deletion_reason=? WHERE id=?",(now,u["login_id"],delete_reason,r["id"]))
        con.commit();n=len(selected);con.close()
        log_activity("LEADS_MOVED_TO_BIN","Selected Leads Moved to MD Recycle Bin","LEADS","Recycle Bin","",str(n),
                     details={"quantity":n,"reason":delete_reason,"restore_authority":"MD only"},severity="WARNING",actor=u)
        flash(f"{n} lead(s) moved to MD Recycle Bin and removed from active data.","success")
    else:con.close();flash("Invalid action.","error")
    return redirect(url_for("lead_control_center"))


@app.route("/lead-control/recycle-bulk",methods=["POST"])
@require_roles("MD")
def recycle_bin_bulk_action():
    u=current_user(); con=db(); action=request.form.get("action","")
    ids=[int(x) for x in request.form.getlist("lead_ids") if str(x).isdigit()]
    scope=""; params=[]
    if action=="empty":
        rows=con.execute("SELECT * FROM leads WHERE COALESCE(deleted_at,'')<>''"+scope,params).fetchall()
    else:
        if not ids:
            con.close(); flash("Select at least one Recycle Bin lead.","error"); return redirect(url_for("lead_control_center"))
        ph=",".join(["?"]*len(ids))
        rows=con.execute(f"SELECT * FROM leads WHERE id IN ({ph}) AND COALESCE(deleted_at,'')<>''"+scope,ids+params).fetchall()
    if action=="restore":
        for r in rows:
            con.execute("UPDATE leads SET deleted_at='',deleted_by='',deletion_reason='' WHERE id=?",(r["id"],))
        con.commit(); n=len(rows); con.close()
        log_activity("LEADS_BULK_RESTORED","Recycle Bin Leads Restored","LEADS","Recycle Bin","",str(n),details={"quantity":n},actor=u)
        flash(f"{n} lead(s) restored to active data.","success")
    elif action in ("erase","empty"):
        audit=[{"lead_id":r["lead_id"],"company":r["company_code"],"batch":r["upload_batch"] or "Legacy","deleted_at":r["deleted_at"]} for r in rows]
        for r in rows:
            con.execute("DELETE FROM lead_activity WHERE lead_id=?",(r["id"],))
            con.execute("DELETE FROM leads WHERE id=?",(r["id"],))
        con.commit(); n=len(rows); con.close()
        log_activity("LEADS_PERMANENTLY_ERASED","Recycle Bin Permanent Erase","SECURITY","Recycle Bin","",str(n),
                     details={"quantity":n,"action":action,"erased_lead_refs":audit[:100]},severity="WARNING",actor=u)
        flash(f"{n} lead(s) permanently erased. Audit event retained.","success")
    else:
        con.close(); flash("Invalid Recycle Bin action.","error")
    return redirect(url_for("lead_control_center"))

@app.route("/team-identity/<int:user_id>/deactivate",methods=["POST"])
@require_roles("MD","GM")
def deactivate_team_user(user_id):
    u=current_user(); con=db(); person=con.execute("SELECT * FROM users WHERE id=?",(user_id,)).fetchone()
    if not person or person["role"]=="MD" or (u["role"]=="GM" and person["company_code"]!=u["company_code"]):
        con.close(); flash("Employee ID cannot be deactivated from this account.","error"); return redirect(url_for("team_identities"))
    con.execute("UPDATE users SET active=0 WHERE id=?",(user_id,))
    if person["employee_id"]:
        con.execute("UPDATE employee_master SET employment_status='Inactive',updated_at=? WHERE id=?",(datetime.datetime.now().isoformat(timespec="seconds"),person["employee_id"]))
    con.commit(); con.close()
    log_activity("EMPLOYEE_ID_DEACTIVATED","Employee Portal ID Deactivated","SECURITY","Employee",user_id,person["login_id"],details={"employee":person["full_name"],"role":person["role"]},severity="WARNING",actor=u,company_code=person["company_code"])
    flash(f"{person['full_name']} login deactivated. Historical records remain safe.","success")
    return redirect(url_for("team_identities"))

@app.route("/team-identity/<int:user_id>/reactivate",methods=["POST"])
@require_roles("MD","GM")
def reactivate_team_user(user_id):
    u=current_user(); con=db(); person=con.execute("SELECT * FROM users WHERE id=?",(user_id,)).fetchone()
    if not person or person["role"]=="MD" or (u["role"]=="GM" and person["company_code"]!=u["company_code"]):
        con.close(); flash("Employee ID cannot be reactivated from this account.","error"); return redirect(url_for("team_identities"))
    con.execute("UPDATE users SET active=1 WHERE id=?",(user_id,))
    if person["employee_id"]:
        con.execute("UPDATE employee_master SET employment_status='Active',updated_at=? WHERE id=?",(datetime.datetime.now().isoformat(timespec="seconds"),person["employee_id"]))
    con.commit(); con.close()
    log_activity("EMPLOYEE_ID_REACTIVATED","Employee Portal ID Reactivated","SECURITY","Employee",user_id,person["login_id"],actor=u,company_code=person["company_code"])
    flash(f"{person['full_name']} login reactivated.","success")
    return redirect(url_for("team_identities"))

@app.route("/team-identity/<int:user_id>/permanent-delete",methods=["POST"])
@require_roles("MD")
def permanent_delete_team_user(user_id):
    u=current_user(); con=db(); person=con.execute("SELECT * FROM users WHERE id=?",(user_id,)).fetchone()
    if not person or person["role"]=="MD":
        con.close(); flash("This portal ID cannot be deleted.","error"); return redirect(url_for("team_identities"))
    if person["active"]:
        con.close(); flash("Deactivate the employee ID before permanent deletion.","error"); return redirect(url_for("team_identities"))
    linked=con.execute("SELECT COUNT(*) c FROM leads WHERE assigned_am=?",(user_id,)).fetchone()["c"]
    if linked:
        con.close(); flash(f"Permanent delete blocked: {linked} lead(s) are linked to this employee. Re-allocate/archive those leads first. The ID remains safely inactive.","error"); return redirect(url_for("team_identities"))
    emp_id=person["employee_id"]
    login_id=person["login_id"]; name=person["full_name"]; company=person["company_code"]; role=person["role"]
    if emp_id:
        con.execute("UPDATE employee_master SET portal_user_id=NULL,employment_status='Former / ID Deleted',updated_at=? WHERE id=?",(datetime.datetime.now().isoformat(timespec="seconds"),emp_id))
    con.execute("DELETE FROM users WHERE id=?",(user_id,))
    con.commit(); con.close()
    log_activity("EMPLOYEE_ID_PERMANENTLY_DELETED","Employee Portal ID Permanently Deleted","SECURITY","Employee",user_id,login_id,details={"employee":name,"role":role},severity="WARNING",actor=u,company_code=company)
    flash(f"Portal ID for {name} permanently deleted. Employee master/audit history retained.","success")
    return redirect(url_for("team_identities"))

@app.route("/lead-control/restore/<int:lead_id>",methods=["POST"])
@require_roles("MD")
def restore_lead(lead_id):
    u=current_user();con=db();r=con.execute("SELECT * FROM leads WHERE id=? AND COALESCE(deleted_at,'')<>''",(lead_id,)).fetchone()
    if not r:con.close();flash("Archived lead not found.","error");return redirect(url_for("lead_control_center"))
    con.execute("UPDATE leads SET deleted_at='',deleted_by='',deletion_reason='' WHERE id=?",(lead_id,));con.commit();con.close()
    log_activity("LEAD_RESTORED","Archived Lead Restored","LEADS","Lead",lead_id,r["lead_id"],actor=u,company_code=r["company_code"]);flash("Lead restored.","success")
    return redirect(url_for("lead_control_center"))

@app.route("/leads/sample-excel")
@require_roles("MD","GM")
def leads_sample_excel():
    """Download a sample Excel file showing the required format for lead upload."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    headers = ["client_name","mobile","email","city","country","visa_type","source"]
    display = ["Client Name *","Mobile *","Email","City","Country","Visa Type","Source / Platform"]
    # Style header row
    gold = PatternFill("solid", fgColor="1a3a5c")
    for col, (h, d) in enumerate(zip(headers, display), 1):
        cell = ws.cell(row=1, column=col, value=d)
        cell.font = Font(bold=True, color="E6B73F")
        cell.fill = gold
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[cell.column_letter].width = 20
    # No sample data rows - user will fill their own data
    # Just freeze the header row for convenience
    ws.freeze_panes = "A2"
    # Notes row at bottom
    ws.cell(row=3, column=1, value="NOTE: Client Name and Mobile are required fields. Fill data from row 2 onwards.")
    ws.cell(row=3, column=1).font = Font(italic=True, color="888888")
    ws.merge_cells("A3:G3")
    import io
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(bio, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="Gaur_CRM_Lead_Upload_Sample.xlsx")

@app.route("/leads", methods=["GET","POST"])
@require_roles("MD","GM")
def leads():
    u=current_user()
    con=db()
    if request.method=="POST":
        f=request.files.get("excel_file")
        if not f or not f.filename:
            flash("Select an Excel file","error")
            return redirect(url_for("leads"))
        company=request.form.get("company_code") if u["role"]=="MD" else u["company_code"]
        fn=f"{uuid.uuid4().hex}_{f.filename}"
        path=os.path.join(UPLOAD_DIR,fn); f.save(path)
        try:
            rows=read_excel(path)
            if not rows: raise ValueError("Excel is empty")
            header_idx=0
            best={}
            for idx in range(min(10,len(rows))):
                m=map_headers(rows[idx])
                if len(m)>len(best):
                    best=m; header_idx=idx
            mapping=best
            if "mobile" not in mapping:
                raise ValueError("Mobile/Contact column not found. Please ensure your Excel has a column named: mobile, phone, contact, or phoneno. Download the Sample Excel for the correct format.")
            total=new=same=cross=invalid=0
            batch_time=datetime.datetime.now()
            batch_stamp=batch_time.isoformat(timespec="seconds")
            batch_code=f"{company}-{batch_time.strftime('%Y%m%d-%H%M')}-{uuid.uuid4().hex[:4].upper()}"
            for raw in rows[header_idx+1:]:
                if not any(str(x).strip() for x in raw if x is not None): continue
                total+=1
                def val(field):
                    i=mapping.get(field)
                    if i is None or i>=len(raw): return ""
                    x=raw[i]
                    return str(x).strip() if x is not None else ""
                mobile=re.sub(r'\D','',val("mobile"))
                dtype=""
                if len(mobile)<8:
                    dtype="invalid"; invalid+=1
                else:
                    same_hit=con.execute("SELECT id FROM leads WHERE company_code=? AND REPLACE(REPLACE(mobile,' ',''),'-','')=? LIMIT 1",(company,mobile)).fetchone()
                    cross_hit=con.execute("SELECT id FROM leads WHERE company_code<>? AND REPLACE(REPLACE(mobile,' ',''),'-','')=? LIMIT 1",(company,mobile)).fetchone()
                    if cross_hit: dtype="cross"; cross+=1
                    elif same_hit: dtype="same"; same+=1
                    else: new+=1
                lead_id=f"{company}-LEAD-{datetime.datetime.now().strftime('%y%m%d')}-{uuid.uuid4().hex[:6].upper()}"
                con.execute("""INSERT INTO leads(lead_id,company_code,client_name,mobile,email,city,country,visa_type,source,duplicate_type,imported_by,imported_at,upload_batch,import_filename)
                               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (lead_id,company,val("client_name"),mobile,val("email"),val("city"),val("country"),val("visa_type"),
                             val("source"),dtype,u["login_id"],batch_stamp,batch_code,f.filename))
            con.execute("""INSERT INTO imports(company_code,filename,total_rows,new_rows,same_company_duplicates,cross_company_duplicates,invalid_mobiles,imported_by,imported_at,batch_code)
                           VALUES(?,?,?,?,?,?,?,?,?,?)""",(company,f.filename,total,new,same,cross,invalid,u["login_id"],batch_stamp,batch_code))
            con.commit()
            log_activity("LEAD_BATCH_UPLOADED","Lead Upload Batch Created","LEADS","Upload Batch",batch_code,batch_code,
                         details={"file":f.filename,"total":total,"new":new,"same_duplicates":same,"cross_duplicates":cross,"invalid":invalid},
                         actor=u,company_code=company)
            flash(f"Batch {batch_code} | Imported {total} rows | New {new} | Same duplicate {same} | Cross duplicate {cross} | Invalid {invalid}","success")
        except Exception as e:
            flash(str(e),"error")
        finally:
            try: os.remove(path)
            except: pass
    company_filter=request.args.get("company","").strip()
    search_text=request.args.get("q","").strip()
    rank_filter=request.args.get("rank","").strip()
    status_filter=request.args.get("status","").strip()
    am_filter=request.args.get("am","").strip()
    sort_filter=request.args.get("sort","interest_desc").strip()
    q="SELECT l.*, u.full_name am_name FROM leads l LEFT JOIN users u ON u.id=l.assigned_am"
    params=[]; clauses=["COALESCE(l.deleted_at,'')=''"]
    if u["role"]=="GM":
        clauses.append("l.company_code=?"); params.append(u["company_code"])
    elif company_filter:
        clauses.append("l.company_code=?"); params.append(company_filter)
    if search_text:
        clauses.append("(LOWER(COALESCE(l.client_name,'')) LIKE ? OR COALESCE(l.mobile,'') LIKE ? OR LOWER(COALESCE(l.lead_id,'')) LIKE ?)")
        like="%"+search_text.lower()+"%"; params.extend([like,"%"+re.sub(r"\D","",search_text)+"%" if re.sub(r"\D","",search_text) else like,like])
    if rank_filter=="hot": clauses.append("COALESCE(l.interest_score,0)>=71")
    elif rank_filter=="warm": clauses.append("COALESCE(l.interest_score,0) BETWEEN 50 AND 70")
    elif rank_filter=="cold": clauses.append("COALESCE(l.interest_score,0)<50")
    if status_filter:
        clauses.append("l.status=?"); params.append(status_filter)
    if am_filter=="unallocated":
        clauses.append("l.assigned_am IS NULL")
    elif am_filter.isdigit():
        clauses.append("l.assigned_am=?"); params.append(int(am_filter))
    if clauses: q+=" WHERE "+" AND ".join(clauses)
    order_map={"interest_desc":"COALESCE(l.interest_score,0) DESC,l.id DESC","interest_asc":"COALESCE(l.interest_score,0) ASC,l.id DESC","newest":"l.id DESC","followup":"CASE WHEN COALESCE(l.followup_date,'')='' THEN 1 ELSE 0 END,l.followup_date ASC,l.id DESC"}
    q+=" ORDER BY "+order_map.get(sort_filter,order_map["interest_desc"])+" LIMIT 500"
    rows=con.execute(q,params).fetchall()
    imports=con.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 20").fetchall()
    if u["role"]=="MD":
        am_rows=con.execute("SELECT * FROM users WHERE role='AM' AND active=1 ORDER BY company_code,full_name").fetchall()
    else:
        am_rows=con.execute("SELECT * FROM users WHERE role='AM' AND active=1 AND company_code=? ORDER BY full_name",(u["company_code"],)).fetchall()
    con.close()
    filters={"company":company_filter,"q":search_text,"rank":rank_filter,"status":status_filter,"am":am_filter,"sort":sort_filter}
    return render_template("leads.html",u=u,rows=rows,imports=imports,am_rows=am_rows,filters=filters)


@app.route("/team-identities",methods=["GET","POST"])
@require_roles("MD","GM")
def team_identities():
    u=current_user(); con=db()
    if request.method=="POST":
        company=request.form.get("company_code","").strip() if u["role"]=="MD" else u["company_code"]
        role=request.form.get("portal_role","").strip().upper()
        allowed_roles=("GM","AM","RECEPTION") if u["role"]=="MD" else ("AM","RECEPTION")
        name=request.form.get("full_name","").strip()
        mobile=normalize_mobile(request.form.get("official_mobile",""))
        email=request.form.get("official_email","").strip().lower()
        designation=request.form.get("designation","").strip() or role
        department=request.form.get("department","").strip()
        branch=request.form.get("branch","").strip()
        joining_date=request.form.get("joining_date","").strip()
        manager_id=request.form.get("reporting_user_id","").strip()
        manager_id=int(manager_id) if manager_id.isdigit() else None
        pw=request.form.get("password","")
        if role not in allowed_roles:
            flash("This portal role cannot be created from your account.","error")
        elif not name or len(mobile)!=10 or len(pw)<6:
            flash("Full name, valid 10-digit official mobile and minimum 6-character password are required.","error")
        elif con.execute("SELECT id FROM users WHERE official_mobile=?",(mobile,)).fetchone():
            flash("This official mobile number is already assigned to another portal identity.","error")
        else:
            identity_code=next_identity_code(con,company,role)
            # Login ID remains professional and predictable.
            login_id=identity_code.lower().replace("-",".")
            counter=1
            original_login=login_id
            while con.execute("SELECT id FROM users WHERE lower(login_id)=?",(login_id.lower(),)).fetchone():
                counter+=1; login_id=f"{original_login}.{counter}"
            now=datetime.datetime.now().isoformat(timespec="seconds")
            cur=con.execute("""INSERT INTO users(
                login_id,password_hash,full_name,role,company_code,active,identity_code,official_mobile,
                official_email,designation,department,branch,joining_date,reporting_user_id,mobile_verified,
                account_created_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (login_id,generate_password_hash(pw),name,role,company,1,identity_code,mobile,email,designation,
             department,branch,joining_date,manager_id,0,now))
            user_id=cur.lastrowid

            seq=con.execute("SELECT COUNT(*) c FROM employee_master WHERE company_code=?",(company,)).fetchone()["c"]+1
            employee_code=f"{company}-EMP-{seq:04d}"
            while con.execute("SELECT id FROM employee_master WHERE employee_code=?",(employee_code,)).fetchone():
                seq+=1; employee_code=f"{company}-EMP-{seq:04d}"
            ecur=con.execute("""INSERT INTO employee_master(
                employee_code,company_code,full_name,designation,department,mobile,email,joining_date,
                employment_status,remarks,photo_path,portal_user_id,created_by,created_at,updated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (employee_code,company,name,designation,department,mobile,email,joining_date,"Active",
             "Professional portal identity created through Team Identity Center","",user_id,u["login_id"],now,now))
            employee_id=ecur.lastrowid
            con.execute("UPDATE users SET employee_id=? WHERE id=?",(employee_id,user_id))

            photo=request.files.get("photograph")
            if photo and photo.filename and allowed_file(photo.filename,ALLOWED_PHOTO_EXTENSIONS):
                pdata=photo.read()
                mime=photo.mimetype or mimetypes.guess_type(photo.filename)[0] or "image/jpeg"
                pname=secure_filename(photo.filename)
                con.execute("UPDATE users SET photo_data=?,photo_mime=?,photo_name=? WHERE id=?",(pdata,mime,pname,user_id))
                con.execute("UPDATE employee_master SET photo_data=?,photo_mime=?,photo_path=? WHERE id=?",(pdata,mime,pname,employee_id))
            con.commit()
            flash(f"Professional identity created: {employee_code} • Login ID: {login_id} • OTP verification pending","success")

    managers_q="SELECT id,full_name,role FROM users WHERE active=1 AND role IN ('MD','GM')"
    manager_params=[]
    if u["role"]=="GM":
        managers_q+=" AND (company_code=? OR role='MD')"; manager_params.append(u["company_code"])
    managers=con.execute(managers_q+" ORDER BY role,full_name",manager_params).fetchall()

    rows_q="""SELECT u.*,e.employee_code FROM users u LEFT JOIN employee_master e ON e.id=u.employee_id
              WHERE u.role IN ('GM','AM','RECEPTION')"""
    params=[]
    if u["role"]=="GM":
        rows_q+=" AND u.company_code=?"; params.append(u["company_code"])
    rows=con.execute(rows_q+" ORDER BY u.company_code,u.role,u.full_name",params).fetchall()
    con.close()
    return render_template("team_identities.html",u=u,rows=rows,managers=managers)

@app.route("/team-identity/<int:user_id>",methods=["GET","POST"])
@require_roles("MD","GM")
def management_identity_profile(user_id):
    viewer=current_user(); con=db()
    person=con.execute("SELECT * FROM users WHERE id=?",(user_id,)).fetchone()
    if not person:
        con.close(); flash("Portal identity not found.","error"); return redirect(url_for("team_identities"))
    if viewer["role"]=="GM" and person["company_code"]!=viewer["company_code"]:
        con.close(); flash("Access denied.","error"); return redirect(url_for("team_identities"))
    if request.method=="POST":
        name=request.form.get("full_name","").strip()
        mobile=normalize_mobile(request.form.get("official_mobile",""))
        email=request.form.get("official_email","").strip().lower()
        designation=request.form.get("designation","").strip()
        department=request.form.get("department","").strip()
        branch=request.form.get("branch","").strip()
        joining=request.form.get("joining_date","").strip()
        rid=request.form.get("reporting_user_id","").strip()
        rid=int(rid) if rid.isdigit() else None
        active=1 if request.form.get("active","1")=="1" else 0
        duplicate=con.execute("SELECT id FROM users WHERE official_mobile=? AND id<>?",(mobile,user_id)).fetchone()
        if not name or len(mobile)!=10:
            flash("Name and valid 10-digit official mobile are required.","error")
        elif duplicate:
            flash("Official mobile is already assigned to another identity.","error")
        else:
            con.execute("""UPDATE users SET full_name=?,official_mobile=?,official_email=?,designation=?,department=?,
                           branch=?,joining_date=?,reporting_user_id=?,active=? WHERE id=?""",
                        (name,mobile,email,designation,department,branch,joining,rid,active,user_id))
            emp=linked_employee(con,person)
            if emp:
                con.execute("""UPDATE employee_master SET full_name=?,mobile=?,email=?,designation=?,department=?,
                               joining_date=?,employment_status=?,updated_at=? WHERE id=?""",
                            (name,mobile,email,designation,department,joining,"Active" if active else "Inactive",
                             datetime.datetime.now().isoformat(timespec="seconds"),emp["id"]))
            photo=request.files.get("photograph")
            if photo and photo.filename and allowed_file(photo.filename,ALLOWED_PHOTO_EXTENSIONS):
                pdata=photo.read(); mime=photo.mimetype or mimetypes.guess_type(photo.filename)[0] or "image/jpeg"
                pname=secure_filename(photo.filename)
                con.execute("UPDATE users SET photo_data=?,photo_mime=?,photo_name=? WHERE id=?",(pdata,mime,pname,user_id))
                if emp:
                    con.execute("UPDATE employee_master SET photo_data=?,photo_mime=?,photo_path=? WHERE id=?",(pdata,mime,pname,emp["id"]))
            con.commit()
            flash("Official identity updated.","success")
            person=con.execute("SELECT * FROM users WHERE id=?",(user_id,)).fetchone()
    employee=linked_employee(con,person)
    mq="SELECT id,full_name,role FROM users WHERE active=1 AND role IN ('MD','GM') AND id<>?"
    mp=[user_id]
    if viewer["role"]=="GM":
        mq+=" AND (company_code=? OR role='MD')"; mp.append(viewer["company_code"])
    managers=con.execute(mq+" ORDER BY role,full_name",mp).fetchall()
    con.close()
    return render_template("management_identity_profile.html",u=viewer,person=person,employee=employee,managers=managers)

@app.route("/digital-id-card")
@app.route("/digital-id-card/<int:user_id>")
def digital_id_card(user_id=None):
    viewer=current_user()
    if not viewer: return redirect(url_for("login"))
    target_id=user_id or viewer["id"]
    if target_id!=viewer["id"] and viewer["role"] not in ("MD","GM"):
        flash("Access denied.","error"); return redirect(url_for("dashboard"))
    con=db()
    person=con.execute("SELECT * FROM users WHERE id=?",(target_id,)).fetchone()
    if not person:
        con.close(); return ("Not found",404)
    if viewer["role"]=="GM" and target_id!=viewer["id"] and person["company_code"]!=viewer["company_code"]:
        con.close(); return ("Access denied",403)
    employee=linked_employee(con,person)
    con.close()
    return render_template("digital_id_card.html",u=viewer,person=person,employee=employee)

@app.route("/ams", methods=["GET","POST"])
@require_roles("MD","GM")
def ams():
    u=current_user(); con=db()
    if request.method=="POST":
        company=request.form.get("company_code") if u["role"]=="MD" else u["company_code"]
        name=request.form.get("full_name","").strip()
        pw=request.form.get("password","").strip()
        mobile=normalize_mobile(request.form.get("official_mobile",""))
        if not name or len(pw)<6 or len(mobile)!=10:
            flash("Name, 10-digit official mobile and minimum 6-character password required. For full profile use Team Identity Center.","error")
        elif con.execute("SELECT id FROM users WHERE official_mobile=?",(mobile,)).fetchone():
            flash("Official mobile number already assigned.","error")
        else:
            identity_code=next_identity_code(con,company,"AM")
            lid=identity_code.lower().replace("-",".")
            now=datetime.datetime.now().isoformat(timespec="seconds")
            cur=con.execute("""INSERT INTO users(login_id,password_hash,full_name,role,company_code,identity_code,
                               official_mobile,designation,mobile_verified,account_created_at)
                               VALUES(?,?,?,?,?,?,?,?,?,?)""",
                        (lid,generate_password_hash(pw),name,"AM",company,identity_code,mobile,"Assistant Manager",0,now))
            uid=cur.lastrowid
            seq=con.execute("SELECT COUNT(*) c FROM employee_master WHERE company_code=?",(company,)).fetchone()["c"]+1
            employee_code=f"{company}-EMP-{seq:04d}"
            while con.execute("SELECT id FROM employee_master WHERE employee_code=?",(employee_code,)).fetchone():
                seq+=1; employee_code=f"{company}-EMP-{seq:04d}"
            ec=con.execute("""INSERT INTO employee_master(employee_code,company_code,full_name,designation,mobile,
                              employment_status,portal_user_id,created_by,created_at,updated_at)
                              VALUES(?,?,?,?,?,?,?,?,?,?)""",
                           (employee_code,company,name,"Assistant Manager",mobile,"Active",uid,u["login_id"],now,now))
            con.execute("UPDATE users SET employee_id=? WHERE id=?",(ec.lastrowid,uid))
            new_emp_id=ec.lastrowid
            con.commit(); con.close()

            # ── OTP + credentials email to MD ────────────────────────────
            otp_sent=False; otp_code=None
            try:
                from modules.security_settings.service import send_am_creation_otp, email_configured
                if email_configured():
                    new_user_obj={"full_name":name,"company_code":company,"login_id":lid}
                    otp_code=send_am_creation_otp(new_user_obj, pw, u)
                    otp_sent=True
            except Exception:
                pass  # email failure should not block creation

            if otp_sent:
                flash(f"✅ AM Portal Created! Login ID: {lid} | Password: {pw} | Employee ID: {employee_code} | Company: {company} — OTP confirmation sent to MD email.","success")
            else:
                flash(f"✅ AM Portal Created! Login ID: {lid} | Password: {pw} | Employee ID: {employee_code} | Company: {company} — (Email not configured, note these credentials)","success")

            return redirect(url_for("edit_employee",employee_id=new_emp_id))
    if u["role"]=="MD":
        rows=con.execute("SELECT * FROM users WHERE role='AM' ORDER BY id DESC").fetchall()
    else:
        rows=con.execute("SELECT * FROM users WHERE role='AM' AND company_code=? ORDER BY id DESC",(u["company_code"],)).fetchall()
    # Build emp_map via single JOIN — no N+1
    row_ids = [r["id"] for r in rows]
    emp_map = {}
    if row_ids:
        ph = ",".join("?" * len(row_ids))
        for emp in con.execute(f"SELECT id,full_name,portal_user_id FROM employee_master WHERE portal_user_id IN ({ph})", row_ids).fetchall():
            emp_map[emp["portal_user_id"]] = dict(emp)
    con.close()
    return render_template("ams.html",u=u,rows=rows,emp_map=emp_map)


# ── Staff helpers (Team Leader / Counsellor / Telecaller) ────────────────────

def _get_am_list(con, u):
    """Return list of AMs with their employee_master id + current team counts for limit display."""
    if u["role"]=="MD":
        ams = con.execute("""
            SELECT u.id AS user_id, u.full_name, u.company_code, e.id AS emp_id
            FROM users u
            LEFT JOIN employee_master e ON e.portal_user_id = u.id
            WHERE u.role='AM' AND u.active=1
            ORDER BY u.company_code, u.full_name
        """).fetchall()
    else:
        ams = con.execute("""
            SELECT u.id AS user_id, u.full_name, u.company_code, e.id AS emp_id
            FROM users u
            LEFT JOIN employee_master e ON e.portal_user_id = u.id
            WHERE u.role='AM' AND u.active=1 AND u.company_code=?
            ORDER BY u.full_name
        """, (u["company_code"],)).fetchall()
    result = []
    for am in ams:
        emp_id = am["emp_id"]
        tl=co=tc=0
        if emp_id:
            team = con.execute("SELECT designation FROM employee_master WHERE reporting_employee_id=?",(emp_id,)).fetchall()
            for t in team:
                dl = (t["designation"] or "").lower()
                if "team leader" in dl: tl+=1
                elif "counsel" in dl: co+=1
                elif "telecall" in dl: tc+=1
        result.append({"emp_id": emp_id, "user_id": am["user_id"],
                        "full_name": am["full_name"], "company_code": am["company_code"],
                        "tl": tl, "co": co, "tc": tc})
    return result


def _save_staff_employee(con, u, designation):
    """Save a new staff employee (TL / Counsellor / Telecaller) from POST form.
    Optionally assign to an AM immediately with limit enforcement."""
    company = request.form.get("company_code") if u["role"]=="MD" else u["company_code"]
    full_name = request.form.get("full_name","").strip()
    if not full_name:
        flash("Employee name is required","error"); return

    # AM assignment (optional) — form sends user_id, resolve to employee_master.id
    am_emp_id = None
    assign_am_id = request.form.get("assign_am_id","").strip()
    if assign_am_id and assign_am_id.isdigit():
        am_user_id = int(assign_am_id)
        # Look up employee_master record for this AM user
        am_emp_row = con.execute("SELECT id FROM employee_master WHERE portal_user_id=?",(am_user_id,)).fetchone()
        if not am_emp_row:
            flash("Selected AM ka employee record nahi mila. Pehle AM ka employee profile create karo.","error"); return
        am_emp_id = am_emp_row["id"]
        # Validate limits before saving
        desig_lower = designation.lower()
        team = con.execute("SELECT designation FROM employee_master WHERE reporting_employee_id=?",(am_emp_id,)).fetchall()
        desigs = [t["designation"].lower() for t in team]
        am_name_row = con.execute("SELECT full_name FROM users WHERE id=?",(am_user_id,)).fetchone()
        am_name = am_name_row["full_name"] if am_name_row else "this AM"
        if "team leader" in desig_lower:
            if sum(1 for d in desigs if "team leader" in d) >= 1:
                flash(f"Is AM ({am_name}) ke paas pehle se ek Team Leader assigned hai. Pehle remove karo.","error"); return
        elif "counsel" in desig_lower:
            if sum(1 for d in desigs if "counsel" in d) >= 2:
                flash(f"Is AM ({am_name}) ke paas pehle se 2 Counsellors assigned hain. Limit full hai.","error"); return
        elif "telecall" in desig_lower:
            if sum(1 for d in desigs if "telecall" in d) >= 2:
                flash(f"Is AM ({am_name}) ke paas pehle se 2 Telecallers assigned hain. Limit full hai.","error"); return

    seq = con.execute("SELECT COUNT(*) c FROM employee_master WHERE company_code=?",(company,)).fetchone()["c"]+1
    employee_code = f"{company}-EMP-{seq:04d}"
    while con.execute("SELECT id FROM employee_master WHERE employee_code=?",(employee_code,)).fetchone():
        seq += 1; employee_code = f"{company}-EMP-{seq:04d}"
    now = datetime.datetime.now().isoformat(timespec="seconds")
    cur = con.execute("""INSERT INTO employee_master(
        employee_code,company_code,full_name,father_spouse_name,designation,department,
        mobile,alternate_mobile,email,date_of_birth,joining_date,address,city,state,pin_code,
        emergency_contact_name,emergency_contact_number,aadhaar_last4,pan_number,bank_name,
        account_last4,ifsc_code,salary_reference,employment_status,remarks,photo_path,
        reporting_employee_id,created_by,created_at,updated_at
    ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
        employee_code,company,full_name,
        request.form.get("father_spouse_name",""),designation,
        request.form.get("department",""),request.form.get("mobile",""),
        request.form.get("alternate_mobile",""),request.form.get("email",""),
        request.form.get("date_of_birth",""),request.form.get("joining_date",""),
        request.form.get("address",""),request.form.get("city",""),
        request.form.get("state",""),request.form.get("pin_code",""),
        request.form.get("emergency_contact_name",""),request.form.get("emergency_contact_number",""),
        request.form.get("aadhaar_last4",""),request.form.get("pan_number",""),
        request.form.get("bank_name",""),request.form.get("account_last4",""),
        request.form.get("ifsc_code",""),request.form.get("salary_reference",""),
        request.form.get("employment_status","Active"),request.form.get("remarks",""),
        "",am_emp_id,u["login_id"],now,now
    ))
    employee_id = cur.lastrowid
    photo = request.files.get("photograph")
    if photo and photo.filename and allowed_file(photo.filename, ALLOWED_PHOTO_EXTENSIONS):
        mime = photo.mimetype or mimetypes.guess_type(photo.filename)[0] or "image/jpeg"
        if IS_POSTGRES or os.environ.get("STORE_FILES_IN_DB","1")=="1":
            con.execute("UPDATE employee_master SET photo_path='DB',photo_data=?,photo_mime=? WHERE id=?",
                        (photo.read(),mime,employee_id))
        else:
            emp_dir = os.path.join(EMPLOYEE_UPLOAD_DIR,employee_code); os.makedirs(emp_dir,exist_ok=True)
            ext = photo.filename.rsplit(".",1)[1].lower()
            full_path = os.path.join(emp_dir,f"photograph.{ext}"); photo.save(full_path)
            con.execute("UPDATE employee_master SET photo_path=? WHERE id=?",
                        (os.path.relpath(full_path,APP_DIR).replace("\\","/"),employee_id))
    con.commit()
    assigned_msg = ""
    if am_emp_id:
        am_name_r = con.execute("SELECT full_name FROM employee_master WHERE id=?",(am_emp_id,)).fetchone()
        if am_name_r:
            assigned_msg = f" → Assigned to {am_name_r['full_name']}"
    flash(f"{designation} saved. Employee Code: {employee_code}{assigned_msg}","success")


def _get_staff_rows(con, u, designations):
    """Fetch staff rows with assigned AM name."""
    ph = tuple("?" for _ in designations)
    lower_desigs = tuple(d.lower() for d in designations)
    ph_str = ",".join(ph)
    if u["role"]=="MD":
        rows = con.execute(f"""
            SELECT e.*, am.full_name AS am_name
            FROM employee_master e
            LEFT JOIN employee_master am ON am.id = e.reporting_employee_id
            WHERE LOWER(e.designation) IN ({ph_str})
            ORDER BY e.company_code, e.full_name
        """, lower_desigs).fetchall()
    else:
        rows = con.execute(f"""
            SELECT e.*, am.full_name AS am_name
            FROM employee_master e
            LEFT JOIN employee_master am ON am.id = e.reporting_employee_id
            WHERE e.company_code=? AND LOWER(e.designation) IN ({ph_str})
            ORDER BY e.full_name
        """, (u["company_code"],) + lower_desigs).fetchall()
    return [dict(r) for r in rows]


@app.route("/team-leaders", methods=["GET","POST"])
@require_roles("MD","GM")
def team_leaders():
    u=current_user(); con=db()
    if request.method=="POST":
        _save_staff_employee(con, u, "Team Leader")
    rows=_get_staff_rows(con, u, ('team leader','teamleader','team lead'))
    ams=_get_am_list(con, u)
    con.close()
    return render_template("staff_directory.html",u=u,rows=rows,ams=ams,page_title="Team Leaders",role_label="Team Leader",back_url=url_for("employees"))


@app.route("/counsellors", methods=["GET","POST"])
@require_roles("MD","GM")
def counsellors():
    u=current_user(); con=db()
    if request.method=="POST":
        _save_staff_employee(con, u, "Counsellor")
    rows=_get_staff_rows(con, u, ('counselor','counsellor','senior counselor','visa counselor','sr. counselor'))
    ams=_get_am_list(con, u)
    con.close()
    return render_template("staff_directory.html",u=u,rows=rows,ams=ams,page_title="My Counsellors",role_label="Counsellor",back_url=url_for("employees"))


@app.route("/telecallers", methods=["GET","POST"])
@require_roles("MD","GM")
def telecallers():
    u=current_user(); con=db()
    if request.method=="POST":
        _save_staff_employee(con, u, "Telecaller")
    rows=_get_staff_rows(con, u, ('telecaller','tele caller','telecalling executive','telecalling'))
    ams=_get_am_list(con, u)
    con.close()
    return render_template("staff_directory.html",u=u,rows=rows,ams=ams,page_title="My Telecallers",role_label="Telecaller",back_url=url_for("employees"))


# ── AM Team Assignment (GM assigns team to each AM) ──────────────────────────

@app.route("/am-team")
@require_roles("MD","GM")
def am_team():
    """GM sees all AMs with their assigned team members."""
    u=current_user(); con=db()
    p="%s" if IS_POSTGRES else "?"
    if u["role"]=="MD":
        ams=con.execute("SELECT * FROM users WHERE role='AM' ORDER BY company_code,full_name").fetchall()
    else:
        ams=con.execute(f"SELECT * FROM users WHERE role='AM' AND company_code={p} ORDER BY full_name",(u["company_code"],)).fetchall()

    # Build team for each AM
    am_teams={}
    for am in ams:
        emp=con.execute(f"SELECT id FROM employee_master WHERE portal_user_id={p}",(am["id"],)).fetchone()
        am_emp_id=emp["id"] if emp else None
        if am_emp_id:
            team=con.execute(f"""SELECT e.id,e.full_name,e.designation,e.mobile,e.employment_status
                                  FROM employee_master e
                                  WHERE e.reporting_employee_id={p}
                                  ORDER BY e.designation,e.full_name""",(am_emp_id,)).fetchall()
        else:
            team=[]
        am_teams[am["id"]]={"emp_id":am_emp_id,"team":[dict(t) for t in team]}

    # Available employees to assign (not yet assigned to anyone in same company)
    company=u["company_code"] if u["role"]!="MD" else None
    if company:
        available=con.execute(f"""SELECT id,full_name,designation,mobile FROM employee_master
                                   WHERE company_code={p}
                                   AND COALESCE(reporting_employee_id,'')=''
                                   AND employment_status='Active'
                                   AND LOWER(designation) NOT IN ('assistant manager','general manager')
                                   ORDER BY designation,full_name""",(company,)).fetchall()
    else:
        available=con.execute("""SELECT id,full_name,designation,mobile,company_code FROM employee_master
                                  WHERE COALESCE(reporting_employee_id,'')=''
                                  AND employment_status='Active'
                                  AND LOWER(designation) NOT IN ('assistant manager','general manager')
                                  ORDER BY company_code,designation,full_name""").fetchall()
    con.close()
    return render_template("am_team.html",u=u,ams=[dict(a) for a in ams],am_teams=am_teams,available=[dict(a) for a in available])


@app.route("/am-team/assign",methods=["POST"])
@require_roles("MD","GM")
def am_team_assign():
    """Assign an employee to an AM's team."""
    u=current_user(); con=db()
    p="%s" if IS_POSTGRES else "?"
    am_user_id=int(request.form.get("am_user_id","0") or 0)
    emp_id=int(request.form.get("emp_id","0") or 0)
    action=request.form.get("action","assign")

    am_emp=con.execute(f"SELECT id,company_code FROM employee_master WHERE portal_user_id={p}",(am_user_id,)).fetchone()
    emp=con.execute(f"SELECT id,full_name,company_code FROM employee_master WHERE id={p}",(emp_id,)).fetchone()

    if not am_emp or not emp:
        con.close(); flash("Invalid selection.","error"); return redirect(url_for("am_team"))
    if u["role"]=="GM" and emp["company_code"]!=u["company_code"]:
        con.close(); flash("Access denied.","error"); return redirect(url_for("am_team"))

    now=datetime.datetime.now().isoformat(timespec="seconds")
    if action=="remove":
        con.execute(f"UPDATE employee_master SET reporting_employee_id=NULL,updated_at={p} WHERE id={p}",(now,emp_id))
        con.commit(); con.close()
        flash(f"{emp['full_name']} team se remove kar diya.","success")
    else:
        # Count current team members by role
        team=con.execute(f"""SELECT designation FROM employee_master WHERE reporting_employee_id={p}""",(am_emp["id"],)).fetchall()
        desigs=[t["designation"].lower() for t in team]
        new_desig=(emp["company_code"] and con.execute(f"SELECT designation FROM employee_master WHERE id={p}",(emp_id,)).fetchone()["designation"] or "").lower()

        # Limits: 1 team leader, 2 counsellors, 2 telecallers
        tl_count=sum(1 for d in desigs if "team leader" in d)
        co_count=sum(1 for d in desigs if "counsell" in d)
        tc_count=sum(1 for d in desigs if "telecall" in d)

        if "team leader" in new_desig and tl_count>=1:
            con.close(); flash("Ek AM ke paas sirf 1 Team Leader ho sakta hai.","error"); return redirect(url_for("am_team"))
        if "counsell" in new_desig and co_count>=2:
            con.close(); flash("Ek AM ke paas sirf 2 Counsellors ho sakte hain.","error"); return redirect(url_for("am_team"))
        if "telecall" in new_desig and tc_count>=2:
            con.close(); flash("Ek AM ke paas sirf 2 Telecallers ho sakte hain.","error"); return redirect(url_for("am_team"))

        con.execute(f"UPDATE employee_master SET reporting_employee_id={p},updated_at={p} WHERE id={p}",(am_emp["id"],now,emp_id))
        con.commit(); con.close()
        flash(f"{emp['full_name']} ko AM ki team mein add kar diya.","success")
    return redirect(url_for("am_team"))


@app.route("/my-team")
@require_roles("AM")
def my_team():
    """AM sees their assigned team."""
    u=current_user(); con=db()
    p="%s" if IS_POSTGRES else "?"
    emp=con.execute(f"SELECT id FROM employee_master WHERE portal_user_id={p}",(u["id"],)).fetchone()
    team=[]
    if emp:
        team=[dict(t) for t in con.execute(f"""SELECT id,full_name,designation,mobile,email,employment_status
                                                FROM employee_master WHERE reporting_employee_id={p}
                                                ORDER BY designation,full_name""",(emp["id"],)).fetchall()]
    con.close()
    return render_template("my_team.html",u=u,team=team)

@app.route("/employee-file/<int:document_id>")
@require_roles("MD","GM")
def employee_file(document_id):
    u=current_user(); con=db()
    row=con.execute("""SELECT d.*,e.company_code FROM employee_documents d
                       JOIN employee_master e ON e.id=d.employee_id WHERE d.id=?""",(document_id,)).fetchone()
    con.close()
    if not row:
        flash("Document not found","error"); return redirect(url_for("employees"))
    if u["role"]=="GM" and row["company_code"]!=u["company_code"]:
        flash("Access denied","error"); return redirect(url_for("employees"))
    if row.get("file_data") if isinstance(row,dict) else row["file_data"]:
        data=row["file_data"]
        if isinstance(data,memoryview): data=data.tobytes()
        return send_file(io.BytesIO(data),mimetype=row["mime_type"] or "application/octet-stream",
                         download_name=row["file_name"],as_attachment=False)
    path=os.path.join(APP_DIR,row["stored_path"])
    if not os.path.exists(path):
        flash("File missing from storage","error"); return redirect(url_for("employees"))
    return send_file(path,as_attachment=False)

@app.route("/employee-photo/<int:employee_id>")
@require_roles("MD","GM")
def employee_photo(employee_id):
    u=current_user(); con=db()
    row=con.execute("SELECT company_code,photo_path,photo_data,photo_mime FROM employee_master WHERE id=?",(employee_id,)).fetchone()
    con.close()
    if not row: return ("",404)
    if u["role"]=="GM" and row["company_code"]!=u["company_code"]: return ("",403)
    pdata=row["photo_data"]
    if pdata:
        if isinstance(pdata,memoryview): pdata=pdata.tobytes()
        resp=send_file(io.BytesIO(pdata),mimetype=row["photo_mime"] or "image/jpeg",max_age=0)
        resp.headers["Cache-Control"]="no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"]="no-cache"
        resp.headers["Expires"]="0"
        return resp
    if not row["photo_path"]: return ("",404)
    path=os.path.join(APP_DIR,row["photo_path"])
    if not os.path.exists(path): return ("",404)
    return send_file(path)


@app.route("/employees", methods=["GET","POST"])
@require_roles("MD","GM")
def employees():
    u=current_user()
    con=db()

    if request.method=="POST":
        company=request.form.get("company_code") if u["role"]=="MD" else u["company_code"]
        full_name=request.form.get("full_name","").strip()
        if not full_name:
            flash("Employee name is required","error")
        else:
            seq=con.execute("SELECT COUNT(*) c FROM employee_master WHERE company_code=?",(company,)).fetchone()["c"]+1
            employee_code=f"{company}-EMP-{seq:04d}"
            while con.execute("SELECT id FROM employee_master WHERE employee_code=?",(employee_code,)).fetchone():
                seq += 1
                employee_code=f"{company}-EMP-{seq:04d}"
            now=datetime.datetime.now().isoformat(timespec="seconds")
            cur=con.execute("""INSERT INTO employee_master(
                employee_code,company_code,full_name,father_spouse_name,designation,department,
                mobile,alternate_mobile,email,date_of_birth,joining_date,address,city,state,pin_code,
                emergency_contact_name,emergency_contact_number,aadhaar_last4,pan_number,bank_name,
                account_last4,ifsc_code,salary_reference,employment_status,remarks,photo_path,created_by,created_at,updated_at
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
                employee_code,company,full_name,request.form.get("father_spouse_name",""),
                request.form.get("designation",""),request.form.get("department",""),
                request.form.get("mobile",""),request.form.get("alternate_mobile",""),
                request.form.get("email",""),request.form.get("date_of_birth",""),
                request.form.get("joining_date",""),request.form.get("address",""),
                request.form.get("city",""),request.form.get("state",""),request.form.get("pin_code",""),
                request.form.get("emergency_contact_name",""),request.form.get("emergency_contact_number",""),
                request.form.get("aadhaar_last4",""),request.form.get("pan_number",""),
                request.form.get("bank_name",""),request.form.get("account_last4",""),
                request.form.get("ifsc_code",""),request.form.get("salary_reference",""),
                request.form.get("employment_status","Active"),request.form.get("remarks",""),
                "",u["login_id"],now,now
            ))
            employee_id=cur.lastrowid

            photo=request.files.get("photograph")
            if photo and photo.filename and allowed_file(photo.filename,ALLOWED_PHOTO_EXTENSIONS):
                mime=photo.mimetype or mimetypes.guess_type(photo.filename)[0] or "image/jpeg"
                if IS_POSTGRES or os.environ.get("STORE_FILES_IN_DB","1")=="1":
                    con.execute("UPDATE employee_master SET photo_path='DB',photo_data=?,photo_mime=? WHERE id=?",
                                (photo.read(),mime,employee_id))
                else:
                    employee_dir=os.path.join(EMPLOYEE_UPLOAD_DIR,employee_code); os.makedirs(employee_dir,exist_ok=True)
                    ext=photo.filename.rsplit(".",1)[1].lower(); full_path=os.path.join(employee_dir,f"photograph.{ext}")
                    photo.save(full_path); rel_path=os.path.relpath(full_path,APP_DIR).replace("\\","/")
                    con.execute("UPDATE employee_master SET photo_path=? WHERE id=?",(rel_path,employee_id))

            con.commit()
            flash(f"Employee saved successfully. Employee Code: {employee_code}","success")

    if u["role"]=="MD":
        rows=con.execute("SELECT * FROM employee_master ORDER BY id DESC").fetchall()
    else:
        rows=con.execute("SELECT * FROM employee_master WHERE company_code=? ORDER BY id DESC",(u["company_code"],)).fetchall()
    doc_counts={r["employee_id"]:r["c"] for r in con.execute("SELECT employee_id,COUNT(*) c FROM employee_documents GROUP BY employee_id").fetchall()}
    con.close()
    return render_template("employees.html",u=u,rows=rows,doc_counts=doc_counts)



def performance_range(period, month_value=None):
    today=datetime.date.today()
    selected_month,selected_start,selected_end,selected_label=selected_month_bounds(month_value)
    if period=="week":
        start=today-datetime.timedelta(days=today.weekday())
        end=start+datetime.timedelta(days=6)
        label=f"Weekly Report ({start.strftime('%d-%b-%Y')} to {end.strftime('%d-%b-%Y')})"
    elif period=="year":
        start=datetime.date(today.year,1,1)
        end=datetime.date(today.year,12,31)
        label=f"Yearly Report {today.year}"
    else:
        start=datetime.date.fromisoformat(selected_start)
        end=datetime.date.fromisoformat(selected_end)
        label="Monthly Report - "+start.strftime("%B %Y")
        period="month"
    return period,start.isoformat(),end.isoformat(),label


@app.route("/performance")
@require_roles("MD","GM")
def performance_directory():
    u=current_user()
    con=db()
    selected_month,_,_,_=selected_month_bounds(request.args.get("month"))
    selected_company=request.args.get("company","").strip()
    selected_designation=request.args.get("designation","").strip()
    eq="SELECT * FROM employee_master WHERE employment_status='Active'"; ep=[]
    if u["role"]=="GM":
        eq+=" AND company_code=?"; ep.append(u["company_code"])
    elif selected_company:
        eq+=" AND company_code=?"; ep.append(selected_company)
    if selected_designation:
        eq+=" AND designation=?"; ep.append(selected_designation)
    employees=con.execute(eq+" ORDER BY company_code,full_name",ep).fetchall()
    dq="SELECT DISTINCT designation FROM employee_master WHERE employment_status='Active' AND COALESCE(designation,'')<>''"
    dp=[]
    if u["role"]=="GM": dq+=" AND company_code=?"; dp.append(u["company_code"])
    designations=[r["designation"] for r in con.execute(dq+" ORDER BY designation",dp).fetchall()]

    period,start_date,end_date,label=performance_range(request.args.get("period","month"),selected_month)
    raw=[]
    for e in employees:
        s=con.execute("""SELECT
              COALESCE(SUM(leads_assigned),0) leads_assigned,
              COALESCE(SUM(calls_done),0) calls_done,
              COALESCE(SUM(followups_done),0) followups_done,
              COALESCE(SUM(office_visits),0) office_visits,
              COALESCE(SUM(interested_clients),0) interested_clients,
              COALESCE(SUM(enrollments),0) enrollments,
              COALESCE(SUM(collection_amount),0) collection_amount,
              COALESCE(SUM(target_amount),0) target_amount,
              COALESCE(AVG(NULLIF(manager_rating,0)),0) avg_rating
            FROM employee_performance
            WHERE employee_id=? AND performance_date BETWEEN ? AND ?""",
            (e["id"],start_date,end_date)).fetchone()
        d=dict(s)
        # Find user linked to this employee — try employee_id first, then name match
        user_row=con.execute("SELECT id FROM users WHERE employee_id=?",(e["id"],)).fetchone()
        if not user_row:
            user_row=con.execute("SELECT id FROM users WHERE full_name=? AND company_code=?",(e["full_name"],e["company_code"])).fetchone()
        linked_enr=0;linked_rev=0.0
        if user_row:
            lr=con.execute("""SELECT COUNT(DISTINCT c.id) enrollments,COALESCE(SUM(COALESCE(c.total_received,0)),0) revenue
                              FROM client_cases c JOIN leads l ON l.id=c.lead_db_id
                              WHERE c.enrollment_date BETWEEN ? AND ? AND l.assigned_am=?""",
                           (start_date,end_date,user_row["id"])).fetchone()
            linked_enr=int(lr["enrollments"] or 0);linked_rev=float(lr["revenue"] or 0)
        # Also check direct company cases (no lead link)
        legacy=con.execute("""SELECT COUNT(DISTINCT id) enrollments,COALESCE(SUM(COALESCE(total_received,0)),0) revenue
                              FROM client_cases WHERE enrollment_date BETWEEN ? AND ?
                              AND lead_db_id IS NULL AND assigned_employee_id=?""",
                           (start_date,end_date,e["id"])).fetchone()
        d["enrollments"]=linked_enr+int(legacy["enrollments"] or 0)
        d["collection_amount"]=linked_rev+float(legacy["revenue"] or 0)
        # ── Live leads count for performance directory ──────────────────────
        live_leads_dir = 0
        live_calls_dir = 0
        live_followups_dir = 0
        live_office_dir = 0
        live_interested_dir = 0
        if user_row:
            uid = user_row["id"]
            am_login = con.execute("SELECT login_id FROM users WHERE id=?", (uid,)).fetchone()
            am_login_id = am_login["login_id"] if am_login else ""

            lr_live = con.execute("""
                SELECT COUNT(*) c FROM leads
                WHERE assigned_am=?
                  AND COALESCE(deleted_at,'')=''
                  AND COALESCE(NULLIF(assigned_at,''), imported_at, '') BETWEEN ? AND ?
            """, (uid, start_date+"T00:00:00", end_date+"T23:59:59")).fetchone()
            live_leads_dir = int(lr_live["c"] or 0)

            if am_login_id:
                lc2 = con.execute("""
                    SELECT COUNT(*) c FROM lead_activity
                    WHERE updated_by=?
                      AND status IN ('Called','Not Picked','Not Connected','Call Back','No Plan','Budget Issue','No WhatsApp','Invalid No.')
                      AND SUBSTR(updated_at,1,10) BETWEEN ? AND ?
                """, (am_login_id, start_date, end_date)).fetchone()
                live_calls_dir = int(lc2["c"] or 0)

                lf2 = con.execute("""
                    SELECT COUNT(*) c FROM lead_activity
                    WHERE updated_by=?
                      AND status IN ('Follow Up','Follow-up','Call Back','Discussion')
                      AND SUBSTR(updated_at,1,10) BETWEEN ? AND ?
                """, (am_login_id, start_date, end_date)).fetchone()
                live_followups_dir = int(lf2["c"] or 0)

                lv2 = con.execute("""
                    SELECT COUNT(*) c FROM lead_activity
                    WHERE updated_by=?
                      AND status='Office Visit'
                      AND SUBSTR(updated_at,1,10) BETWEEN ? AND ?
                """, (am_login_id, start_date, end_date)).fetchone()
                live_office_dir = int(lv2["c"] or 0)

            li2 = con.execute("""
                SELECT COUNT(*) c FROM leads
                WHERE assigned_am=?
                  AND COALESCE(deleted_at,'')=''
                  AND COALESCE(interest_score,0) >= 50
                  AND COALESCE(NULLIF(assigned_at,''), imported_at, '') BETWEEN ? AND ?
            """, (uid, start_date+"T00:00:00", end_date+"T23:59:59")).fetchone()
            live_interested_dir = int(li2["c"] or 0)

        d["leads_assigned"] = live_leads_dir
        d["calls_done"] = live_calls_dir
        d["followups_done"] = live_followups_dir
        d["office_visits"] = live_office_dir
        d["interested_clients"] = live_interested_dir
        # ───────────────────────────────────────────────────────────────────
        d["employee"]=e
        raw.append(d)

    max_enrollments=max([r["enrollments"] or 0 for r in raw], default=0)
    max_revenue=max([r["collection_amount"] or 0 for r in raw], default=0)
    total_enrollments=sum(r["enrollments"] or 0 for r in raw)
    total_revenue=sum(r["collection_amount"] or 0 for r in raw)
    summaries={}
    ranked=[]

    for r in raw:
        e=r["employee"]
        enrollments=r["enrollments"] or 0
        revenue=r["collection_amount"] or 0
        target=r["target_amount"] or 0
        rating=r["avg_rating"] or 0
        leads=r["leads_assigned"] or 0

        # Enrollment-only performance rule.
        enrollment_index=(enrollments/max_enrollments*100) if max_enrollments else 0
        revenue_index=0
        target_achievement=0
        conversion=0
        score=min(enrollment_index,100)

        r.update({
            "enrollment_index":enrollment_index,
            "revenue_index":revenue_index,
            "target_achievement":target_achievement,
            "conversion":conversion,
            "score":score
        })
        summaries[e["id"]]=r
        ranked.append({
            "id":e["id"],
            "employee_code":e["employee_code"],
            "full_name":e["full_name"],
            "company_code":e["company_code"],
            "designation":e["designation"],
            "photo_path":e["photo_path"],
            "score":score,
            "enrollments":enrollments,
            "revenue":revenue,
            "target_achievement":target_achievement,
            "conversion":conversion,
            "rating":rating
        })

    ranked.sort(key=lambda x:(-x["enrollments"],x["full_name"].lower()))
    for idx,r in enumerate(ranked,1):
        r["rank"]=idx
    top3=ranked[:3]

    active_employees=len(employees)
    avg_enrollments=(total_enrollments/active_employees) if active_employees else 0
    avg_revenue=(total_revenue/active_employees) if active_employees else 0
    top_performer=top3[0] if top3 else None

    # Build month-by-month trend for selected calendar year using the same ranking formula.
    year=datetime.date.fromisoformat(start_date).year
    trend=[]
    # Take up to 3 current top performers and show their monthly performance scores.
    trend_employees=top3
    for month in range(1,13):
        month_start=datetime.date(year,month,1)
        month_end=(datetime.date(year+1,1,1)-datetime.timedelta(days=1)) if month==12 else (datetime.date(year,month+1,1)-datetime.timedelta(days=1))
        month_raw=[]
        for emp in trend_employees:
            s=con.execute("""SELECT
                  COALESCE(SUM(enrollments),0) enrollments,
                  COALESCE(SUM(collection_amount),0) collection_amount,
                  COALESCE(SUM(target_amount),0) target_amount,
                  COALESCE(AVG(NULLIF(manager_rating,0)),0) avg_rating
                FROM employee_performance
                WHERE employee_id=? AND performance_date BETWEEN ? AND ?""",
                (emp["id"],month_start.isoformat(),month_end.isoformat())).fetchone()
            md=dict(s)
            case_stats=con.execute("""SELECT COUNT(*) enrollments, COALESCE(SUM(COALESCE(total_received,0)),0) revenue
                                      FROM client_cases WHERE assigned_employee_id=? AND enrollment_date BETWEEN ? AND ?""",
                                   (emp["id"],month_start.isoformat(),month_end.isoformat())).fetchone()
            md["enrollments"]=case_stats["enrollments"]
            md["collection_amount"]=case_stats["revenue"]
            month_raw.append({"id":emp["id"],**md})
        month_max_enroll=max([m["enrollments"] or 0 for m in month_raw],default=0)
        month_max_revenue=max([m["collection_amount"] or 0 for m in month_raw],default=0)
        item={"month":month_start.strftime("%b")}
        for idx,m in enumerate(month_raw,1):
            enrollment_index=(m["enrollments"]/month_max_enroll*100) if month_max_enroll else 0
            revenue_index=(m["collection_amount"]/month_max_revenue*100) if month_max_revenue else 0
            target_achievement=(m["collection_amount"]/m["target_amount"]*100) if m["target_amount"] else 0
            rating=m["avg_rating"] or 0
            mscore=(min(enrollment_index,100)*0.40 +
                    min(revenue_index,100)*0.40 +
                    min(target_achievement,100)*0.10 +
                    min(rating*20,100)*0.10)
            item[f"score{idx}"]=round(mscore,1)
        trend.append(item)

    year_months=[]
    for mm in range(1,13):
        ms=datetime.date(year,mm,1)
        me=(datetime.date(year+1,1,1)-datetime.timedelta(days=1)) if mm==12 else (datetime.date(year,mm+1,1)-datetime.timedelta(days=1))
        mr=employee_month_rankings(con,ms.isoformat(),me.isoformat(),u["company_code"] if u["role"]=="GM" else (selected_company or None),selected_designation or None)
        year_months.append({"key":ms.strftime("%Y-%m"),"label":ms.strftime("%b"),"top_name":mr[0]["full_name"] if mr else "","top_score":round(mr[0]["score"]) if mr else 0})
    con.close()
    return render_template("performance_directory.html",u=u,employees=employees,summaries=summaries,
                           ranked=ranked,top3=top3,period=period,label=label,start_date=start_date,end_date=end_date,
                           total_enrollments=total_enrollments,total_revenue=total_revenue,
                           avg_enrollments=avg_enrollments,avg_revenue=avg_revenue,
                           active_employees=active_employees,top_performer=top_performer,trend=trend,
                           selected_month=selected_month,selected_company=selected_company,
                           selected_designation=selected_designation,designations=designations,year_months=year_months)


@app.route("/employees/<int:employee_id>/performance", methods=["GET","POST"])
@require_roles("MD","GM")
def employee_performance(employee_id):
    u=current_user()
    con=db()
    employee=con.execute("SELECT * FROM employee_master WHERE id=?",(employee_id,)).fetchone()
    if not employee:
        con.close()
        flash("Employee record not found","error")
        return redirect(url_for("employees"))
    if u["role"]=="GM" and employee["company_code"]!=u["company_code"]:
        con.close()
        flash("Access denied","error")
        return redirect(url_for("employees"))

    if request.method=="POST":
        pdate=request.form.get("performance_date") or datetime.date.today().isoformat()
        try:
            rating=float(request.form.get("manager_rating","0") or 0)
        except:
            rating=0
        rating=max(0,min(5,rating))
        con.execute("""INSERT INTO employee_performance(
            employee_id,performance_date,leads_assigned,calls_done,followups_done,office_visits,
            interested_clients,enrollments,collection_amount,target_amount,attendance_status,
            manager_rating,remarks,created_by,created_at
        ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",(
            employee_id,pdate,
            int(request.form.get("leads_assigned","0") or 0),
            int(request.form.get("calls_done","0") or 0),
            int(request.form.get("followups_done","0") or 0),
            int(request.form.get("office_visits","0") or 0),
            int(request.form.get("interested_clients","0") or 0),
            int(request.form.get("enrollments","0") or 0),
            float(request.form.get("collection_amount","0") or 0),
            float(request.form.get("target_amount","0") or 0),
            request.form.get("attendance_status","Present"),
            rating,request.form.get("remarks",""),u["login_id"],
            datetime.datetime.now().isoformat(timespec="seconds")
        ))
        con.commit()
        flash("Performance entry saved","success")
        return redirect(url_for("employee_performance",employee_id=employee_id,period=request.args.get("period","month"),month=request.args.get("month","")))

    period,start_date,end_date,label=performance_range(request.args.get("period","month"),request.args.get("month"))
    rows=con.execute("""SELECT * FROM employee_performance
                        WHERE employee_id=? AND performance_date BETWEEN ? AND ?
                        ORDER BY performance_date DESC,id DESC""",
                     (employee_id,start_date,end_date)).fetchall()

    totals=dict(con.execute("""SELECT
          COALESCE(SUM(leads_assigned),0) leads_assigned,
          COALESCE(SUM(calls_done),0) calls_done,
          COALESCE(SUM(followups_done),0) followups_done,
          COALESCE(SUM(office_visits),0) office_visits,
          COALESCE(SUM(interested_clients),0) interested_clients,
          COALESCE(SUM(enrollments),0) enrollments,
          COALESCE(SUM(collection_amount),0) collection_amount,
          COALESCE(SUM(target_amount),0) target_amount,
          COALESCE(AVG(NULLIF(manager_rating,0)),0) avg_rating,
          COUNT(*) entries
        FROM employee_performance
        WHERE employee_id=? AND performance_date BETWEEN ? AND ?""",
        (employee_id,start_date,end_date)).fetchone())
    case_stats=con.execute("""SELECT COUNT(*) enrollments,COALESCE(SUM(COALESCE(total_received,0)),0) revenue
                              FROM client_cases WHERE assigned_employee_id=? AND enrollment_date BETWEEN ? AND ?""",
                           (employee_id,start_date,end_date)).fetchone()
    # Two-path enrollment lookup
    emp_row=con.execute("SELECT id FROM users WHERE employee_id=?",(employee_id,)).fetchone()
    if not emp_row:
        emp_master=con.execute("SELECT full_name,company_code FROM employee_master WHERE id=?",(employee_id,)).fetchone()
        if emp_master:
            emp_row=con.execute("SELECT id FROM users WHERE full_name=? AND company_code=?",(emp_master["full_name"],emp_master["company_code"])).fetchone()
    linked_enr=0;linked_rev=0.0
    if emp_row:
        lr=con.execute("""SELECT COUNT(DISTINCT c.id) enrollments,COALESCE(SUM(COALESCE(c.total_received,0)),0) revenue
                          FROM client_cases c JOIN leads l ON l.id=c.lead_db_id
                          WHERE c.enrollment_date BETWEEN ? AND ? AND l.assigned_am=?""",
                       (start_date,end_date,emp_row["id"])).fetchone()
        linked_enr=int(lr["enrollments"] or 0);linked_rev=float(lr["revenue"] or 0)
    legacy2=con.execute("""SELECT COUNT(DISTINCT id) enrollments,COALESCE(SUM(COALESCE(total_received,0)),0) revenue
                           FROM client_cases WHERE enrollment_date BETWEEN ? AND ?
                           AND lead_db_id IS NULL AND assigned_employee_id=?""",
                        (start_date,end_date,employee_id)).fetchone()
    totals["enrollments"]=linked_enr+int(legacy2["enrollments"] or 0)
    totals["collection_amount"]=linked_rev+float(legacy2["revenue"] or 0)

    # ── Live leads count (overrides manual entry) ──────────────────────────
    live_leads = 0
    live_calls = 0
    live_followups = 0
    live_office_visits = 0
    live_interested = 0
    if emp_row:
        uid = emp_row["id"]
        # Get AM's login_id for lead_activity lookup
        am_login = con.execute("SELECT login_id FROM users WHERE id=?", (uid,)).fetchone()
        am_login_id = am_login["login_id"] if am_login else ""

        # Leads assigned in period
        lr2 = con.execute("""
            SELECT COUNT(*) c FROM leads
            WHERE assigned_am=?
              AND COALESCE(deleted_at,'')=''
              AND COALESCE(NULLIF(assigned_at,''), imported_at, '') BETWEEN ? AND ?
        """, (uid, start_date+"T00:00:00", end_date+"T23:59:59")).fetchone()
        live_leads = int(lr2["c"] or 0)

        if am_login_id:
            # Calls done from lead_activity
            lc = con.execute("""
                SELECT COUNT(*) c FROM lead_activity
                WHERE updated_by=?
                  AND status IN ('Called','Not Picked','Not Connected','Call Back','No Plan','Budget Issue','No WhatsApp','Invalid No.')
                  AND SUBSTR(updated_at,1,10) BETWEEN ? AND ?
            """, (am_login_id, start_date, end_date)).fetchone()
            live_calls = int(lc["c"] or 0)

            # Follow-ups done
            lf = con.execute("""
                SELECT COUNT(*) c FROM lead_activity
                WHERE updated_by=?
                  AND status IN ('Follow Up','Follow-up','Call Back','Discussion')
                  AND SUBSTR(updated_at,1,10) BETWEEN ? AND ?
            """, (am_login_id, start_date, end_date)).fetchone()
            live_followups = int(lf["c"] or 0)

            # Office visits
            lv = con.execute("""
                SELECT COUNT(*) c FROM lead_activity
                WHERE updated_by=?
                  AND status='Office Visit'
                  AND SUBSTR(updated_at,1,10) BETWEEN ? AND ?
            """, (am_login_id, start_date, end_date)).fetchone()
            live_office_visits = int(lv["c"] or 0)

        # Interested clients (interest_score >= 50 among assigned leads)
        li = con.execute("""
            SELECT COUNT(*) c FROM leads
            WHERE assigned_am=?
              AND COALESCE(deleted_at,'')=''
              AND COALESCE(interest_score,0) >= 50
              AND COALESCE(NULLIF(assigned_at,''), imported_at, '') BETWEEN ? AND ?
        """, (uid, start_date+"T00:00:00", end_date+"T23:59:59")).fetchone()
        live_interested = int(li["c"] or 0)

    totals["leads_assigned"] = live_leads
    totals["calls_done"] = live_calls
    totals["followups_done"] = live_followups
    totals["office_visits"] = live_office_visits
    totals["interested_clients"] = live_interested
    # ───────────────────────────────────────────────────────────────────────

    emp_row_user_id = emp_row["id"] if emp_row else None
    con.close()

    conversion=(totals["enrollments"]/totals["leads_assigned"]*100) if totals["leads_assigned"] else 0
    achievement=(totals["collection_amount"]/totals["target_amount"]*100) if totals["target_amount"] else 0

    share_text=(
        f"ABC | XYZ Employee Performance\\n"
        f"Employee: {employee['full_name']} ({employee['employee_code']})\\n"
        f"Company: {employee['company_code']}\\n"
        f"{label}\\n\\n"
        f"Leads Assigned: {totals['leads_assigned']}\\n"
        f"Calls Done: {totals['calls_done']}\\n"
        f"Follow-ups: {totals['followups_done']}\\n"
        f"Office Visits: {totals['office_visits']}\\n"
        f"Interested Clients: {totals['interested_clients']}\\n"
        f"Enrollments: {totals['enrollments']}\\n"
        f"Conversion: {conversion:.1f}%\\n"
        f"Collection: ₹{totals['collection_amount']:,.0f}\\n"
        f"Target: ₹{totals['target_amount']:,.0f}\\n"
        f"Target Achievement: {achievement:.1f}%\\n"
        f"Manager Rating: {totals['avg_rating']:.1f}/5"
    )
    whatsapp_url="https://wa.me/?text="+quote(share_text)

    return render_template("employee_performance.html",u=u,employee=employee,rows=rows,
                           totals=totals,period=period,label=label,start_date=start_date,end_date=end_date,
                           conversion=conversion,achievement=achievement,share_text=share_text,
                           whatsapp_url=whatsapp_url,emp_row_user_id=emp_row_user_id)


@app.route("/employees/<int:employee_id>/documents", methods=["GET","POST"])
@require_roles("MD","GM")
def employee_documents(employee_id):
    u=current_user()
    con=db()
    employee=con.execute("SELECT * FROM employee_master WHERE id=?",(employee_id,)).fetchone()
    if not employee:
        con.close()
        flash("Employee record not found","error")
        return redirect(url_for("employees"))
    if u["role"]=="GM" and employee["company_code"]!=u["company_code"]:
        con.close()
        flash("Access denied","error")
        return redirect(url_for("employees"))

    if request.method=="POST":
        f=request.files.get("document_file")
        doc_type=request.form.get("document_type","Other").strip()
        doc_name=request.form.get("document_name","").strip()
        if not f or not f.filename:
            flash("Select a document file","error")
        elif not allowed_file(f.filename,ALLOWED_DOC_EXTENSIONS):
            flash("Allowed formats: PDF, JPG, JPEG, PNG, WEBP, DOC, DOCX","error")
        else:
            safe=secure_filename(f.filename)
            mime=f.mimetype or mimetypes.guess_type(safe)[0] or "application/octet-stream"
            now=datetime.datetime.now().isoformat(timespec="seconds")
            if IS_POSTGRES or os.environ.get("STORE_FILES_IN_DB","1")=="1":
                con.execute("""INSERT INTO employee_documents(employee_id,document_type,document_name,file_name,stored_path,file_data,mime_type,uploaded_by,uploaded_at)
                               VALUES(?,?,?,?,?,?,?,?,?)""",
                            (employee_id,doc_type,doc_name,safe,"DB",f.read(),mime,u["login_id"],now))
            else:
                emp_dir=os.path.join(EMPLOYEE_UPLOAD_DIR,employee["employee_code"],"documents"); os.makedirs(emp_dir,exist_ok=True)
                stamp=datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f"); full_path=os.path.join(emp_dir,f"{stamp}_{safe}")
                f.save(full_path); rel_path=os.path.relpath(full_path,APP_DIR).replace("\\","/")
                con.execute("""INSERT INTO employee_documents(employee_id,document_type,document_name,file_name,stored_path,uploaded_by,uploaded_at)
                               VALUES(?,?,?,?,?,?,?)""",(employee_id,doc_type,doc_name,safe,rel_path,u["login_id"],now))
            con.commit()
            flash("Employee document uploaded","success")

    docs=con.execute("SELECT * FROM employee_documents WHERE employee_id=? ORDER BY id DESC",(employee_id,)).fetchall()
    con.close()
    return render_template("employee_documents.html",u=u,employee=employee,docs=docs)


@app.route("/employees/<int:employee_id>/edit", methods=["GET","POST"])
@require_roles("MD","GM")
def edit_employee(employee_id):
    u=current_user()
    con=db()
    row=con.execute("SELECT * FROM employee_master WHERE id=?",(employee_id,)).fetchone()
    if not row:
        con.close()
        flash("Employee record not found","error")
        return redirect(url_for("employees"))
    if u["role"]=="GM" and row["company_code"]!=u["company_code"]:
        con.close()
        flash("Access denied","error")
        return redirect(url_for("employees"))

    if request.method=="POST":
        company=row["company_code"] if u["role"]=="GM" else request.form.get("company_code",row["company_code"])
        now=datetime.datetime.now().isoformat(timespec="seconds")
        con.execute("""UPDATE employee_master SET
            company_code=?,full_name=?,father_spouse_name=?,designation=?,department=?,mobile=?,
            alternate_mobile=?,email=?,date_of_birth=?,joining_date=?,address=?,city=?,state=?,pin_code=?,
            emergency_contact_name=?,emergency_contact_number=?,aadhaar_last4=?,pan_number=?,bank_name=?,
            account_last4=?,ifsc_code=?,salary_reference=?,employment_status=?,remarks=?,updated_at=?
            WHERE id=?""",(
            company,request.form.get("full_name","").strip(),request.form.get("father_spouse_name",""),
            request.form.get("designation",""),request.form.get("department",""),
            request.form.get("mobile",""),request.form.get("alternate_mobile",""),
            request.form.get("email",""),request.form.get("date_of_birth",""),
            request.form.get("joining_date",""),request.form.get("address",""),
            request.form.get("city",""),request.form.get("state",""),request.form.get("pin_code",""),
            request.form.get("emergency_contact_name",""),request.form.get("emergency_contact_number",""),
            request.form.get("aadhaar_last4",""),request.form.get("pan_number",""),
            request.form.get("bank_name",""),request.form.get("account_last4",""),
            request.form.get("ifsc_code",""),request.form.get("salary_reference",""),
            request.form.get("employment_status","Active"),request.form.get("remarks",""),
            now,employee_id
        ))
        photo=request.files.get("photograph")
        if photo and photo.filename:
            if allowed_file(photo.filename,ALLOWED_PHOTO_EXTENSIONS):
                mime=photo.mimetype or mimetypes.guess_type(photo.filename)[0] or "image/jpeg"
                if IS_POSTGRES or os.environ.get("STORE_FILES_IN_DB","1")=="1":
                    con.execute("UPDATE employee_master SET photo_path='DB',photo_data=?,photo_mime=?,updated_at=? WHERE id=?",
                                (photo.read(),mime,now,employee_id))
                else:
                    employee_dir=os.path.join(EMPLOYEE_UPLOAD_DIR,row["employee_code"]); os.makedirs(employee_dir,exist_ok=True)
                    ext=photo.filename.rsplit(".",1)[1].lower(); full_path=os.path.join(employee_dir,f"photograph.{ext}")
                    photo.save(full_path); rel_path=os.path.relpath(full_path,APP_DIR).replace("\\","/")
                    con.execute("UPDATE employee_master SET photo_path=?,updated_at=? WHERE id=?",(rel_path,now,employee_id))
            else:
                flash("Photograph must be JPG, JPEG, PNG or WEBP","error")
        con.commit()
        con.close()
        flash("Employee information updated","success")
        return redirect(url_for("employees"))

    portal_user=None
    if row["portal_user_id"]:
        portal_user=con.execute(
            "SELECT login_id,role,active,identity_code FROM users WHERE id=?",
            (row["portal_user_id"],)
        ).fetchone()
    con.close()
    return render_template("employee_edit.html",u=u,row=row,portal_user=portal_user)


@app.route("/ams/<int:user_id>/employee")
@require_roles("MD","GM")
def am_employee_redirect(user_id):
    """AM Directory se name click karne pe employee edit form pe redirect karta hai."""
    u=current_user(); con=db()
    am=con.execute("SELECT * FROM users WHERE id=? AND role='AM'",(user_id,)).fetchone()
    if not am or (u["role"]=="GM" and am["company_code"]!=u["company_code"]):
        con.close(); flash("Access denied","error"); return redirect(url_for("ams"))
    # Try employee_id link first
    emp=None
    if am["employee_id"]:
        emp=con.execute("SELECT id FROM employee_master WHERE id=?",(am["employee_id"],)).fetchone()
    # Fallback: search by portal_user_id
    if not emp:
        emp=con.execute("SELECT id FROM employee_master WHERE portal_user_id=?",(user_id,)).fetchone()
    con.close()
    if emp:
        return redirect(url_for("edit_employee",employee_id=emp["id"]))
    flash("Is AM ka employee record nahi mila. Pehle Employee Information mein add karein.","error")
    return redirect(url_for("ams"))


@app.route("/ams/<int:user_id>/toggle", methods=["POST"])
@require_roles("MD","GM")
def toggle_am(user_id):
    u=current_user(); con=db()
    am=con.execute("SELECT * FROM users WHERE id=? AND role='AM'",(user_id,)).fetchone()
    if not am or (u["role"]=="GM" and am["company_code"]!=u["company_code"]):
        con.close(); flash("Assistant Manager not found or access denied","error"); return redirect(url_for("ams"))
    new_active=0 if am["active"] else 1
    con.execute("UPDATE users SET active=? WHERE id=?",(new_active,user_id)); con.commit(); con.close()
    flash("Assistant Manager portal activated" if new_active else "Assistant Manager portal deactivated. Data has been preserved.","success")
    return redirect(url_for("ams"))

@app.route("/ams/<int:user_id>/request-delete", methods=["POST"])
@require_roles("MD","GM")
def request_delete_am(user_id):
    """Seedha delete — no OTP."""
    u=current_user(); con=db()
    am=con.execute("SELECT * FROM users WHERE id=? AND role='AM' AND active=0",(user_id,)).fetchone()
    if not am or (u["role"]=="GM" and am["company_code"]!=u["company_code"]):
        con.close(); flash("Only inactive AMs can be deleted, or access denied.","error")
        return redirect(url_for("ams"))
    name=am["full_name"]; login_id=am["login_id"]
    con.execute("DELETE FROM employee_master WHERE portal_user_id=?",(user_id,))
    con.execute("DELETE FROM users WHERE id=?",(user_id,))
    con.commit(); con.close()
    flash(f"AM '{login_id}' ({name}) permanently delete ho gaya.","success")
    return redirect(url_for("ams"))


@app.route("/ams/<int:user_id>/confirm-delete", methods=["POST"])
@require_roles("MD","GM")
def confirm_delete_am(user_id):
    """Legacy route — redirect to ams."""
    return redirect(url_for("ams"))


@app.route("/ams/handover", methods=["POST"])
@require_roles("MD","GM")
def handover_am():
    u=current_user(); con=db()
    old_id=int(request.form.get("old_am_id","0") or 0); new_id=int(request.form.get("new_am_id","0") or 0)
    old=con.execute("SELECT * FROM users WHERE id=? AND role='AM'",(old_id,)).fetchone()
    new=con.execute("SELECT * FROM users WHERE id=? AND role='AM' AND active=1",(new_id,)).fetchone()
    if not old or not new or old_id==new_id or old["company_code"]!=new["company_code"] or (u["role"]=="GM" and old["company_code"]!=u["company_code"]):
        con.close(); flash("Select two valid AMs from the same company","error"); return redirect(url_for("ams"))
    count=con.execute("SELECT COUNT(*) c FROM leads WHERE assigned_am=?",(old_id,)).fetchone()["c"]
    now=datetime.datetime.now().isoformat(timespec="seconds")
    con.execute("UPDATE leads SET assigned_am=?,assigned_at=?,assigned_by=? WHERE assigned_am=?",(new_id,now,u["login_id"],old_id))
    con.execute("UPDATE users SET active=0 WHERE id=?",(old_id,))
    con.execute("INSERT INTO allocation_history(company_code,am_user_id,quantity,allocated_by,allocated_at) VALUES(?,?,?,?,?)",(new["company_code"],new_id,count,u["login_id"]+" HANDOVER",now))
    con.commit(); con.close(); flash(f"Handover complete: {count} leads transferred. Old Assistant Manager portal deactivated.","success")
    return redirect(url_for("ams"))

def payment_status_counts(status):
    """Only statuses confirming money received contribute to collection/revenue."""
    s=(status or "").strip().lower()
    return s in {
        "received","paid","completed","complete","success","successful","cleared",
        "partially received","partial","part payment received"
    }

def effective_case_received(first_payment, first_payment_status, second_payment, second_payment_status, other_payment_received):
    """Package / After-Visa amounts are contractual values, not collection."""
    first=float(first_payment or 0) if payment_status_counts(first_payment_status) else 0.0
    second=float(second_payment or 0) if payment_status_counts(second_payment_status) else 0.0
    other=float(other_payment_received or 0)
    return first+second+other


@app.route("/client-record/<int:lead_db_id>/edit", methods=["GET","POST"])
@require_roles("MD","GM")
def edit_client_record(lead_db_id):
    u=current_user(); con=db()
    lead=con.execute("""SELECT l.*,au.full_name am_name
                        FROM leads l
                        LEFT JOIN users au ON au.id=l.assigned_am
                        WHERE l.id=?""",(lead_db_id,)).fetchone()
    if not lead or (u["role"]=="GM" and lead["company_code"]!=u["company_code"]):
        con.close(); flash("Client record not found or access denied.","error")
        return redirect(url_for("cases"))

    case=con.execute("SELECT * FROM client_cases WHERE lead_db_id=? ORDER BY id DESC LIMIT 1",(lead_db_id,)).fetchone()

    if request.method=="POST":
        client_name=(request.form.get("client_name") or "").strip()
        mobile=(request.form.get("mobile") or "").strip()
        country=(request.form.get("country") or "").strip()
        visa_type=(request.form.get("visa_type") or "").strip()
        remarks=(request.form.get("remarks") or "").strip()
        status=(request.form.get("status") or lead["status"] or "New").strip()
        followup_date=(request.form.get("followup_date") or "").strip()
        now=datetime.datetime.now().isoformat(timespec="seconds")

        if not client_name:
            con.close(); flash("Client name is required.","error")
            return redirect(url_for("edit_client_record",lead_db_id=lead_db_id))

        # Update the master AM-maintained Lead record.
        con.execute("""UPDATE leads
                       SET client_name=?,mobile=?,country=?,visa_type=?,status=?,
                           followup_date=?,remarks=?
                       WHERE id=?""",
                    (client_name,mobile,country,visa_type,status,followup_date,remarks,lead_db_id))

        # Keep the enrolled Client Case synchronized with the SAME client record.
        if case:
            con.execute("""UPDATE client_cases
                           SET client_name=?,mobile=?,country=?,visa_type=?,remarks=?,updated_at=?
                           WHERE lead_db_id=?""",
                        (client_name,mobile,country,visa_type,remarks,now,lead_db_id))

        con.execute("""INSERT INTO lead_activity
                       (lead_id,status,followup_date,remarks,updated_by,updated_at)
                       VALUES(?,?,?,?,?,?)""",
                    (lead_db_id,status,followup_date,
                     "Client details edited from GM/MD portal. "+remarks,
                     u["login_id"],now))
        con.commit(); con.close()
        flash("Client details updated successfully.","success")
        return redirect(url_for("cases"))

    con.close()
    return render_template("edit_client_record.html",u=u,lead=lead,case=case)



@app.route("/accounts/client-ledger-legacy", methods=["GET","POST"])
@require_roles("MD","GM")
def enrollment_report_legacy_redirect():
    # Old builds/bookmarks may still point here. Never execute the legacy ledger.
    return redirect(url_for("enrollment_report"))

@app.route("/accounts/client-ledger", methods=["GET"])
@require_roles("MD","GM")
def enrollment_report():
    u=current_user()
    con=db()

    # v3.98 filters
    company_filter=(request.args.get("company") or "").strip()
    am_filter=(request.args.get("am") or "").strip()
    search_q=(request.args.get("q") or "").strip()
    if u["role"]=="GM":
        company_filter=u["company_code"] or ""

    # 1) Core enrollment rows: SELECT * only. No optional joins can break page opening.
    try:
        if u["role"]=="GM":
            raw=con.execute("SELECT * FROM client_cases WHERE company_code=? AND COALESCE(deleted_at,'')='' ORDER BY id DESC",(u["company_code"],)).fetchall()
        else:
            raw=con.execute("SELECT * FROM client_cases WHERE COALESCE(deleted_at,'')='' ORDER BY id DESC").fetchall()
    except Exception:
        try: con.rollback()
        except Exception: pass
        raw=[]

    rows=[]
    core_defaults={
      "id":0,"case_id":"","company_code":"","client_name":"","mobile":"","country":"","visa_type":"",
      "enrollment_date":"","lead_db_id":None,"counselor_name":"","am_name":"","interaction_date":"",
      "package_amount":0,"after_visa_payment":0,"booking_amount":0,"first_payment_status":"Pending",
      "first_payment_datetime":"","second_payment":0,"second_payment_status":"Pending",
      "second_payment_datetime":"","other_payment_received":0,"other_payment_datetime":"",
      "total_received":0,"payment_status":"Pending","filing_status":"Documents Pending","remarks":"",
      "first_cash":0,"first_rbl":0,"first_yes_bank":0,"first_au_bank":0,"first_indusind":0,
      "second_cash":0,"second_rbl":0,"second_yes_bank":0,"second_au_bank":0,"second_indusind":0,
      "first_payment_method":"","second_payment_method":"","other_payment_method":"","other_payment_bank":"",
      "first_payment_bank_id":None,"second_payment_bank_id":None,"other_payment_bank_id":None,
      "lead_code":"","linked_am_name":"","passport_no":"","assigned_am":None,"first_bank_name":"","second_bank_name":"","other_bank_name":"",
      "first_proof_id":None,"second_proof_id":None,"other_proof_id":None
    }

    # Detect optional objects without changing schema.
    try:
        case_cols=_table_columns(con,"client_cases")
    except Exception:
        case_cols=set()
    try:
        proof_ok=_table_exists(con,"client_payment_proofs")
    except Exception:
        proof_ok=False

    for rr in raw:
        try: d=dict(rr)
        except Exception: d={k:rr[k] for k in rr.keys()}
        for k,v in core_defaults.items():
            if k not in d or d[k] is None: d[k]=v

        # Best-effort Lead ID + AM name.
        if d.get("lead_db_id"):
            try:
                lr=con.execute("""SELECT l.lead_id,l.passport_no,l.assigned_am,au.full_name am_name
                                  FROM leads l LEFT JOIN users au ON au.id=l.assigned_am WHERE l.id=?""",
                               (d["lead_db_id"],)).fetchone()
                if lr:
                    d["lead_code"]=lr["lead_id"] or ""
                    d["linked_am_name"]=lr["am_name"] or ""
                    d["passport_no"]=lr["passport_no"] or ""
                    d["assigned_am"]=lr["assigned_am"]
            except Exception:
                try: con.rollback()
                except Exception: pass

        # Best-effort bank labels.
        for idkey,namekey in (("first_payment_bank_id","first_bank_name"),
                              ("second_payment_bank_id","second_bank_name"),
                              ("other_payment_bank_id","other_bank_name")):
            bid=d.get(idkey)
            if bid:
                try:
                    br=con.execute("SELECT bank_name,account_last4 FROM finance_banks WHERE id=?",(bid,)).fetchone()
                    if br: d[namekey]=(br["bank_name"] or "")+" ****"+(br["account_last4"] or "")
                except Exception:
                    try: con.rollback()
                    except Exception: pass

        if proof_ok:
            for stage,key in (("FIRST","first_proof_id"),("SECOND","second_proof_id"),("OTHER","other_proof_id")):
                try:
                    pr=con.execute("SELECT id FROM client_payment_proofs WHERE case_id=? AND payment_stage=? ORDER BY id DESC LIMIT 1",
                                   (d["id"],stage)).fetchone()
                    if pr:d[key]=pr["id"]
                except Exception:
                    try: con.rollback()
                    except Exception: pass
        rows.append(d)

    # Apply Company / AM / Search filters after safe enrichment.
    if company_filter:
        rows=[x for x in rows if (x.get("company_code") or "")==company_filter]
    if am_filter.isdigit():
        rows=[x for x in rows if str(x.get("assigned_am") or "")==am_filter]
    if search_q:
        _needle=search_q.lower()
        def _matches(x):
            hay=[
              x.get("client_name"),x.get("mobile"),x.get("passport_no"),
              x.get("lead_code"),x.get("case_id")
            ]
            return any(_needle in str(v or "").lower() for v in hay)
        rows=[x for x in rows if _matches(x)]

    # Filter dropdown options are role scoped.
    try:
        if u["role"]=="MD":
            am_options=con.execute("""SELECT id,full_name,company_code FROM users
                                      WHERE role='AM' AND active=1 ORDER BY company_code,full_name""").fetchall()
            company_options=["SCIC","WWIC"]
        else:
            am_options=con.execute("""SELECT id,full_name,company_code FROM users
                                      WHERE role='AM' AND active=1 AND company_code=?
                                      ORDER BY full_name""",(u["company_code"],)).fetchall()
            company_options=[u["company_code"]]
    except Exception:
        try: con.rollback()
        except Exception: pass
        am_options=[]; company_options=[u["company_code"]] if u["role"]=="GM" else ["SCIC","WWIC"]

    # Active banks are optional. Failure returns an empty dropdown, never HTTP 500.
    banks=[]
    try:
        companies=_finance_companies(u)
        placeholders=",".join(["?"]*len(companies))
        banks=con.execute("""SELECT * FROM finance_banks WHERE active=1 AND company_code IN ("""+
                          placeholders+") ORDER BY company_code,bank_name",companies).fetchall()
    except Exception:
        try: con.rollback()
        except Exception: pass
        banks=[]

    con.close()
    return render_template("enrollment_report_safe.html",u=u,rows=rows,banks=banks,
                           company_filter=company_filter,am_filter=am_filter,search_q=search_q,
                           company_options=company_options,am_options=am_options)




@app.route("/enrollment/<int:case_id>/delete",methods=["POST"])
@require_roles("MD","GM")
def enrollment_soft_delete(case_id):
    ensure_enrollment_recycle_schema()
    u=current_user(); con=db()
    row=con.execute("SELECT * FROM client_cases WHERE id=?",(case_id,)).fetchone()
    if not row or (u["role"]=="GM" and row["company_code"]!=u["company_code"]):
        con.close(); flash("Enrollment record not found or access denied.","error")
        return redirect(request.referrer or url_for("enrollment_report"))
    reason=(request.form.get("reason") or "Removed from Enrollment by Management").strip()[:250]
    now=datetime.datetime.now().isoformat(timespec="seconds")
    con.execute("UPDATE client_cases SET deleted_at=?,deleted_by=?,deletion_reason=? WHERE id=?",
                (now,u["full_name"],reason,case_id))
    con.commit(); con.close()
    log_activity("ENROLLMENT_MOVED_TO_RECYCLE_BIN","Enrollment Moved to Recycle Bin","ENROLLMENT",
                 "Client Case",case_id,row["case_id"],
                 details={"client":row["client_name"],"company":row["company_code"],"reason":reason},
                 severity="WARNING",actor=u,company_code=row["company_code"])
    flash(f"{row['client_name']} enrollment moved to Recycle Bin.","success")
    return redirect(request.referrer or url_for("enrollment_report"))

@app.route("/enrollment-recycle-bin")
@require_roles("MD")
def enrollment_recycle_bin():
    ensure_enrollment_recycle_schema()
    u=current_user(); con=db()
    rows=con.execute("""SELECT * FROM client_cases
                        WHERE COALESCE(deleted_at,'')<>''
                        ORDER BY deleted_at DESC,id DESC""").fetchall()
    con.close()
    return render_template("enrollment_recycle_bin.html",u=u,rows=rows)

@app.route("/enrollment-recycle-bin/<int:case_id>/restore",methods=["POST"])
@require_roles("MD")
def enrollment_restore(case_id):
    ensure_enrollment_recycle_schema()
    u=current_user(); con=db()
    row=con.execute("SELECT * FROM client_cases WHERE id=? AND COALESCE(deleted_at,'')<>''",(case_id,)).fetchone()
    if not row:
        con.close(); flash("Enrollment record not found in Recycle Bin.","error")
        return redirect(url_for("enrollment_recycle_bin"))
    con.execute("UPDATE client_cases SET deleted_at='',deleted_by='',deletion_reason='' WHERE id=?",(case_id,))
    con.commit(); con.close()
    log_activity("ENROLLMENT_RESTORED","Enrollment Restored","ENROLLMENT","Client Case",case_id,row["case_id"],
                 details={"client":row["client_name"],"company":row["company_code"]},
                 actor=u,company_code=row["company_code"])
    flash(f"{row['client_name']} enrollment restored successfully.","success")
    return redirect(url_for("enrollment_recycle_bin"))

@app.route("/enrollment-recycle-bin/<int:case_id>/permanent-delete",methods=["POST"])
@require_roles("MD")
def enrollment_permanent_delete(case_id):
    ensure_enrollment_recycle_schema()
    u=current_user(); con=db()
    row=con.execute("SELECT * FROM client_cases WHERE id=? AND COALESCE(deleted_at,'')<>''",(case_id,)).fetchone()
    if not row:
        con.close(); flash("Enrollment must be in Recycle Bin before final deletion.","error")
        return redirect(url_for("enrollment_recycle_bin"))
    try:
        if _table_exists(con,"client_payment_proofs"):
            con.execute("DELETE FROM client_payment_proofs WHERE case_id=?",(case_id,))
    except Exception:
        try: con.rollback()
        except Exception: pass
    try:
        if _table_exists(con,"competition_events"):
            con.execute("DELETE FROM competition_events WHERE case_id=?",(row["case_id"],))
    except Exception:
        try: con.rollback()
        except Exception: pass
    audit={"case_id":row["case_id"],"client":row["client_name"],"company":row["company_code"],
           "deleted_at":row["deleted_at"],"deleted_by":row["deleted_by"]}
    con.execute("DELETE FROM client_cases WHERE id=?",(case_id,))
    con.commit(); con.close()
    log_activity("ENROLLMENT_PERMANENTLY_DELETED","Enrollment Permanently Deleted","SECURITY",
                 "Client Case",case_id,row["case_id"],details=audit,severity="WARNING",
                 actor=u,company_code=row["company_code"])
    flash(f"{row['client_name']} enrollment permanently deleted.","success")
    return redirect(url_for("enrollment_recycle_bin"))


@app.route("/cases", methods=["GET"])
@require_roles("MD","GM","AM")
def cases():
    u=current_user()
    con=db()

    # Core client account rows only.
    try:
        if u["role"]=="AM":
            # AM sirf apne assigned leads ke cases dekh sakta hai
            try:
                raw=con.execute("""SELECT cc.* FROM client_cases cc
                                   JOIN leads l ON l.id=cc.lead_db_id
                                   WHERE COALESCE(cc.deleted_at,'')='' AND l.assigned_am=?
                                   ORDER BY cc.id DESC""",(u["id"],)).fetchall()
            except Exception:
                try: con.rollback()
                except Exception: pass
                raw=[]
            if not raw:
                # fallback: am_name se match karo
                try:
                    raw=con.execute("""SELECT * FROM client_cases
                                       WHERE COALESCE(deleted_at,'')='' AND am_name=?
                                       ORDER BY id DESC""",(u["full_name"],)).fetchall()
                except Exception:
                    try: con.rollback()
                    except Exception: pass
                    raw=[]
        elif u["role"]=="GM":
            raw=con.execute("SELECT * FROM client_cases WHERE company_code=? AND COALESCE(deleted_at,'')='' ORDER BY id DESC",(u["company_code"],)).fetchall()
        else:
            raw=con.execute("SELECT * FROM client_cases WHERE COALESCE(deleted_at,'')='' ORDER BY id DESC").fetchall()
    except Exception:
        try: con.rollback()
        except Exception: pass
        raw=[]

    defaults={
      "id":0,"case_id":"","company_code":"","client_name":"","mobile":"","country":"","visa_type":"",
      "enrollment_date":"","lead_db_id":None,"counselor_name":"","am_name":"","interaction_date":"",
      "package_amount":0,"after_visa_payment":0,"booking_amount":0,
      "first_payment_status":"Pending","first_payment_datetime":"",
      "second_payment":0,"second_payment_status":"Pending","second_payment_datetime":"",
      "other_payment_received":0,"other_payment_datetime":"","total_received":0,
      "payment_status":"Pending","filing_status":"Documents Pending","remarks":"",
      "first_cash":0,"first_rbl":0,"first_yes_bank":0,"first_au_bank":0,"first_indusind":0,
      "second_cash":0,"second_rbl":0,"second_yes_bank":0,"second_au_bank":0,"second_indusind":0,
      "first_payment_method":"","second_payment_method":"","other_payment_method":"","other_payment_bank":"",
      "first_payment_bank_id":None,"second_payment_bank_id":None,"other_payment_bank_id":None,
      "lead_code":"","linked_am_name":"","passport_no":"",
      "first_bank_name":"","second_bank_name":"","other_bank_name":"",
      "first_proof_id":None,"second_proof_id":None,"other_proof_id":None
    }

    rows=[]
    proof_ok=False
    try: proof_ok=_table_exists(con,"client_payment_proofs")
    except Exception: proof_ok=False

    for rr in raw:
        try: d=dict(rr)
        except Exception: d={k:rr[k] for k in rr.keys()}
        for k,v in defaults.items():
            if k not in d or d[k] is None: d[k]=v

        if d.get("lead_db_id"):
            try:
                lr=con.execute("""SELECT l.lead_id,l.passport_no,au.full_name am_name
                                  FROM leads l LEFT JOIN users au ON au.id=l.assigned_am WHERE l.id=?""",
                               (d["lead_db_id"],)).fetchone()
                if lr:
                    d["lead_code"]=lr["lead_id"] or ""
                    d["passport_no"]=lr["passport_no"] or ""
                    d["linked_am_name"]=lr["am_name"] or ""
            except Exception:
                try: con.rollback()
                except Exception: pass

        for idkey,namekey in (("first_payment_bank_id","first_bank_name"),
                              ("second_payment_bank_id","second_bank_name"),
                              ("other_payment_bank_id","other_bank_name")):
            bid=d.get(idkey)
            if bid:
                try:
                    br=con.execute("SELECT bank_name,account_last4 FROM finance_banks WHERE id=?",(bid,)).fetchone()
                    if br: d[namekey]=(br["bank_name"] or "")+" ****"+(br["account_last4"] or "")
                except Exception:
                    try: con.rollback()
                    except Exception: pass

        if proof_ok:
            for stage,key in (("FIRST","first_proof_id"),("SECOND","second_proof_id"),("OTHER","other_proof_id")):
                try:
                    pr=con.execute("SELECT id FROM client_payment_proofs WHERE case_id=? AND payment_stage=? ORDER BY id DESC LIMIT 1",
                                   (d["id"],stage)).fetchone()
                    if pr: d[key]=pr["id"]
                except Exception:
                    try: con.rollback()
                    except Exception: pass
        rows.append(d)

    # Active banks are optional.
    banks=[]
    try:
        companies=_finance_companies(u)
        placeholders=",".join(["?"]*len(companies))
        banks=con.execute("SELECT * FROM finance_banks WHERE active=1 AND company_code IN ("+
                          placeholders+") ORDER BY company_code,bank_name",companies).fetchall()
    except Exception:
        try: con.rollback()
        except Exception: pass
        banks=[]

    con.close()
    return render_template("cases_safe.html",u=u,rows=rows,banks=banks)



@app.route("/cases-legacy-open")
@require_roles("MD","GM")
def cases_legacy_redirect():
    return redirect(url_for("cases"))

@app.route("/cases-legacy", methods=["GET","POST"])
@require_roles("MD","GM")
def cases_legacy():
    ensure_enrollment_payment_schema(); ensure_bank_manager_schema()
    u=current_user();con=db()
    # Repair old Enrolled leads, but NEVER block Enrollment Report if legacy schema is inconsistent.
    if request.method=="GET":
        try:
            if u["role"]=="GM":
                _enrolled=con.execute("""SELECT id FROM leads WHERE company_code=? AND LOWER(TRIM(COALESCE(status,'')))='enrolled'
                                         AND COALESCE(deleted_at,'')='' ORDER BY id""",(u["company_code"],)).fetchall()
            else:
                _enrolled=con.execute("""SELECT id FROM leads WHERE LOWER(TRIM(COALESCE(status,'')))='enrolled'
                                         AND COALESCE(deleted_at,'')='' ORDER BY id""").fetchall()
            _made=0
            for _r in _enrolled:
                try:
                    _c,_new=ensure_enrollment_case_for_lead(con,_r["id"],u["login_id"])
                    if _new:_made+=1
                except Exception:
                    try: con.rollback()
                    except Exception: pass
            if _made: con.commit()
        except Exception:
            try: con.rollback()
            except Exception: pass
    if request.method=="POST":
        company=request.form.get("company_code") if u["role"]=="MD" else u["company_code"]
        lead_db_id=int(request.form.get("lead_db_id") or 0) or None
        lead=None
        if lead_db_id:
            lead=con.execute("""SELECT l.*,au.full_name am_name FROM leads l LEFT JOIN users au ON au.id=l.assigned_am WHERE l.id=?""",(lead_db_id,)).fetchone()
            if not lead or lead["company_code"]!=company:
                con.close();flash("Selected Lead is not available for this company.","error");return redirect(url_for("cases"))
        cname=(request.form.get("client_name") or (lead["client_name"] if lead else "") or "").strip()
        if not cname:
            con.close();flash("Client name is required","error");return redirect(url_for("cases"))
        seq=con.execute("SELECT COUNT(*) c FROM client_cases WHERE company_code=?",(company,)).fetchone()["c"]+1
        case_id=f"{company}-CASE-{datetime.date.today().strftime('%y%m%d')}-{seq:04d}"
        while con.execute("SELECT id FROM client_cases WHERE case_id=?",(case_id,)).fetchone():
            seq+=1;case_id=f"{company}-CASE-{datetime.date.today().strftime('%y%m%d')}-{seq:04d}"
        def amt(k): return float(request.form.get(k,"0") or 0)
        fc,fr,fy,fa,fi=amt("first_cash"),amt("first_rbl"),amt("first_yes_bank"),amt("first_au_bank"),amt("first_indusind")
        sc,sr,sy,sa,si=amt("second_cash"),amt("second_rbl"),amt("second_yes_bank"),amt("second_au_bank"),amt("second_indusind")
        first=fc+fr+fy+fa+fi;second=sc+sr+sy+sa+si;other=amt("other_payment_received")
        pkg=amt("package_amount");avp=amt("after_visa_payment");total=first+second+other
        fs=request.form.get("first_payment_status") or ("Received" if first else "Pending")
        ss=request.form.get("second_payment_status") or ("Received" if second else "Pending")
        now=datetime.datetime.now().isoformat(timespec="seconds")
        fdt=request.form.get("first_payment_datetime","") or (now if first else "")
        sdt=request.form.get("second_payment_datetime","") or (now if second else "")
        odt=request.form.get("other_payment_datetime","") or (now if other else "")
        counselor=request.form.get("counselor_name","").strip()
        am_name=((lead["am_name"] if lead and lead["am_name"] else "") or request.form.get("am_name","")).strip()
        mobile=(request.form.get("mobile") or (lead["mobile"] if lead else "") or "").strip()
        country=(request.form.get("country") or (lead["country"] if lead else "") or "").strip()
        visa=(request.form.get("visa_type") or (lead["visa_type"] if lead else "") or "").strip()
        con.execute("""INSERT INTO client_cases(case_id,company_code,client_name,mobile,country,visa_type,enrollment_date,
          lead_db_id,counselor_name,am_name,interaction_date,package_amount,after_visa_payment,
          first_cash,first_rbl,first_yes_bank,first_au_bank,first_indusind,booking_amount,first_payment_status,first_payment_datetime,
          second_cash,second_rbl,second_yes_bank,second_au_bank,second_indusind,second_payment,second_payment_status,second_payment_datetime,
          other_payment_received,other_payment_datetime,total_received,payment_status,filing_status,remarks,created_by,created_at,updated_at)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
          (case_id,company,cname,mobile,country,visa,request.form.get("enrollment_date") or datetime.date.today().isoformat(),
           lead_db_id,counselor,am_name,request.form.get("interaction_date",""),pkg,avp,
           fc,fr,fy,fa,fi,first,fs,fdt,sc,sr,sy,sa,si,second,ss,sdt,other,odt,total,ss,
           request.form.get("filing_status","Documents Pending"),request.form.get("remarks",""),u["login_id"],now,now))
        _new_case_row=con.execute("SELECT id FROM client_cases WHERE case_id=?",(case_id,)).fetchone()
        if _new_case_row:
            def _bid(name):
                v=(request.form.get(name) or "").strip()
                return int(v) if v.isdigit() else None
            con.execute("""UPDATE client_cases SET first_payment_method=?,second_payment_method=?,
                           other_payment_method=?,other_payment_bank=?,first_payment_bank_id=?,second_payment_bank_id=?,other_payment_bank_id=? WHERE id=?""",
                        (request.form.get("first_payment_method",""),request.form.get("second_payment_method",""),
                         request.form.get("other_payment_method",""),request.form.get("other_payment_bank",""),
                         _bid("first_payment_bank_id"),_bid("second_payment_bank_id"),_bid("other_payment_bank_id"),_new_case_row["id"]))
            for st,field in (("FIRST","first_payment_proof"),("SECOND","second_payment_proof"),("OTHER","other_payment_proof")):
                f=request.files.get(field)
                if f and f.filename:
                    data=f.read()
                    if len(data)<=10*1024*1024:
                        con.execute("""INSERT INTO client_payment_proofs(case_id,payment_stage,file_name,mime_type,file_data,uploaded_by_id,uploaded_by_name,uploaded_at)
                                       VALUES(?,?,?,?,?,?,?,?)""",(_new_case_row["id"],st,secure_filename(f.filename)[:180],f.mimetype or "application/octet-stream",data,u["id"],u["full_name"],now))
        if lead_db_id:
            con.execute("UPDATE leads SET status='Enrolled' WHERE id=?",(lead_db_id,))
            con.execute("INSERT INTO lead_activity(lead_id,status,followup_date,remarks,updated_by,updated_at) VALUES(?,?,?,?,?,?)",
                        (lead_db_id,"Enrolled","",f"Converted to {case_id}",u["login_id"],now))
        con.commit()
        try: record_competition_enrollment(con,case_id,company,am_name,total)
        except Exception: pass
        flash(f"Saved {case_id} • Received ₹{total:,.0f} • Balance ₹{max(pkg-total,0):,.0f}","success")

    wh="";pa=[]
    if u["role"]=="GM":
        wh=" WHERE company_code=?"; pa=[u["company_code"]]

    # Core ledger MUST open even if optional bank/proof schema is unavailable.
    try:
        _raw_rows=con.execute("SELECT * FROM client_cases"+wh+" ORDER BY id DESC",pa).fetchall()
    except Exception:
        try: con.rollback()
        except Exception: pass
        _raw_rows=[]

    # Convert DB rows to normal dictionaries and provide every field expected by Jinja.
    _defaults={
      "lead_db_id":None,"counselor_name":"","am_name":"","interaction_date":"",
      "package_amount":0,"after_visa_payment":0,
      "first_cash":0,"first_rbl":0,"first_yes_bank":0,"first_au_bank":0,"first_indusind":0,
      "second_cash":0,"second_rbl":0,"second_yes_bank":0,"second_au_bank":0,"second_indusind":0,
      "booking_amount":0,"second_payment":0,"other_payment_received":0,"total_received":0,
      "first_payment_status":"Pending","second_payment_status":"Pending",
      "first_payment_datetime":"","second_payment_datetime":"","other_payment_datetime":"",
      "first_payment_method":"","second_payment_method":"","other_payment_method":"","other_payment_bank":"",
      "first_payment_bank_id":None,"second_payment_bank_id":None,"other_payment_bank_id":None,
      "first_bank_name":"","second_bank_name":"","other_bank_name":"",
      "first_proof_id":None,"second_proof_id":None,"other_proof_id":None,
      "lead_code":"","linked_am_name":"","filing_status":"Documents Pending","remarks":""
    }
    rows=[]
    _proof_ok=_table_exists(con,"client_payment_proofs")
    for _rr in _raw_rows:
        try: _d=dict(_rr)
        except Exception:
            _d={k:_rr[k] for k in _rr.keys()}
        for _k,_v in _defaults.items(): _d.setdefault(_k,_v)

        # Lead ID / AM enrichment is best-effort only.
        _lid=_d.get("lead_db_id")
        if _lid:
            try:
                _lr=con.execute("""SELECT l.lead_id,au.full_name linked_am_name
                                   FROM leads l LEFT JOIN users au ON au.id=l.assigned_am WHERE l.id=?""",(_lid,)).fetchone()
                if _lr:
                    _d["lead_code"]=_lr["lead_id"] or ""
                    _d["linked_am_name"]=_lr["linked_am_name"] or ""
            except Exception:
                try: con.rollback()
                except Exception: pass

        # Selected bank labels are best-effort only.
        for _idkey,_namekey in (("first_payment_bank_id","first_bank_name"),
                                ("second_payment_bank_id","second_bank_name"),
                                ("other_payment_bank_id","other_bank_name")):
            _bid=_d.get(_idkey)
            if _bid:
                try:
                    _br=con.execute("SELECT bank_name FROM finance_banks WHERE id=?",(_bid,)).fetchone()
                    if _br: _d[_namekey]=_br["bank_name"] or ""
                except Exception:
                    try: con.rollback()
                    except Exception: pass

        # Latest proof links are best-effort only.
        if _proof_ok:
            for _stage,_key in (("FIRST","first_proof_id"),("SECOND","second_proof_id"),("OTHER","other_proof_id")):
                try:
                    _pr=con.execute("SELECT id FROM client_payment_proofs WHERE case_id=? AND payment_stage=? ORDER BY id DESC LIMIT 1",
                                    (_d.get("id"),_stage)).fetchone()
                    if _pr: _d[_key]=_pr["id"]
                except Exception:
                    try: con.rollback()
                    except Exception: pass
        rows.append(_d)

    # Leads dropdown: also safe and company-scoped.
    try:
        if u["role"]=="MD":
            leads=con.execute("""SELECT l.id,l.lead_id,l.client_name,l.mobile,l.country,l.visa_type,l.company_code,
                                 au.full_name am_name
                                 FROM leads l LEFT JOIN users au ON au.id=l.assigned_am
                                 WHERE COALESCE(l.deleted_at,'')=''
                                 ORDER BY CASE WHEN LOWER(TRIM(COALESCE(l.status,'')))='enrolled' THEN 0 ELSE 1 END,l.id DESC LIMIT 800""").fetchall()
        else:
            leads=con.execute("""SELECT l.id,l.lead_id,l.client_name,l.mobile,l.country,l.visa_type,l.company_code,
                                 au.full_name am_name
                                 FROM leads l LEFT JOIN users au ON au.id=l.assigned_am
                                 WHERE l.company_code=? AND COALESCE(l.deleted_at,'')=''
                                 ORDER BY CASE WHEN LOWER(TRIM(COALESCE(l.status,'')))='enrolled' THEN 0 ELSE 1 END,l.id DESC LIMIT 800""",
                              (u["company_code"],)).fetchall()
    except Exception:
        try: con.rollback()
        except Exception: pass
        leads=[]

    try:
        if u["role"]=="MD":
            counselors=con.execute("SELECT full_name,company_code FROM employee_master WHERE employment_status='Active' ORDER BY full_name").fetchall()
        else:
            counselors=con.execute("SELECT full_name,company_code FROM employee_master WHERE employment_status='Active' AND company_code=? ORDER BY full_name",
                                   (u["company_code"],)).fetchall()
    except Exception:
        try: con.rollback()
        except Exception: pass
        counselors=[]

    # Bank Manager integration is optional; zero banks must never break Enrollment.
    try:
        _companies=_finance_companies(u)
        banks=con.execute("""SELECT id,company_code,bank_name,account_name,account_last4,
                             COALESCE(account_nickname,'') account_nickname,COALESCE(upi_id,'') upi_id
                             FROM finance_banks WHERE active=1 AND company_code IN ("""+
                          ",".join(["?"]*len(_companies))+") ORDER BY company_code,bank_name",_companies).fetchall()
    except Exception:
        try: con.rollback()
        except Exception: pass
        try:
            _companies=_finance_companies(u)
            banks=con.execute("SELECT id,company_code,bank_name,account_name,account_last4,'' account_nickname,'' upi_id FROM finance_banks WHERE active=1 AND company_code IN ("+
                              ",".join(["?"]*len(_companies))+") ORDER BY company_code,bank_name",_companies).fetchall()
        except Exception:
            try: con.rollback()
            except Exception: pass
            banks=[]

    con.close()
    return render_template("cases.html",u=u,rows=rows,leads=leads,counselors=counselors,banks=banks)


@app.route("/cases/payment-proof/<int:proof_id>")
@require_roles("MD","GM")
def client_payment_proof(proof_id):
    ensure_enrollment_payment_schema()
    u=current_user(); con=db()
    try:
        p=con.execute("SELECT p.*,c.company_code FROM client_payment_proofs p JOIN client_cases c ON c.id=p.case_id WHERE p.id=?",(proof_id,)).fetchone()
    except Exception:
        try: con.rollback()
        except Exception: pass
        p=None
    if not p or (u["role"]=="GM" and p["company_code"]!=u["company_code"]):
        con.close(); return "Proof not found or access denied",404
    data=p["file_data"]
    if isinstance(data,memoryview): data=data.tobytes()
    name=p["file_name"] or "payment-proof"; mime=p["mime_type"] or "application/octet-stream"
    con.close(); return send_file(io.BytesIO(bytes(data)),mimetype=mime,download_name=name,as_attachment=False)


@app.route("/cases/<int:case_id>/update", methods=["POST"])
@require_roles("MD","GM","AM")
def update_case(case_id):
    ensure_enrollment_payment_schema(); ensure_bank_manager_schema()
    u=current_user();con=db();r=con.execute("SELECT * FROM client_cases WHERE id=?",(case_id,)).fetchone()
    if not r or (u["role"]=="GM" and r["company_code"]!=u["company_code"]):
        con.close();flash("Case not found or access denied","error");return redirect(url_for("cases"))
    # AM sirf pehli baar add kar sakta hai — agar am_submitted=1 toh block
    if u["role"]=="AM":
        already_submitted = int(r["am_submitted"] if "am_submitted" in r.keys() else 0) == 1
        if already_submitted:
            con.close();flash("Tum pehle hi yeh enrollment submit kar chuke ho. Ab sirf GM/MD edit kar sakte hain.","error");return redirect(url_for("cases"))
    def a(k,old): return float(request.form.get(k,old or 0) or 0)
    fc,fr,fy,fa,fi=a("first_cash",r["first_cash"]),a("first_rbl",r["first_rbl"]),a("first_yes_bank",r["first_yes_bank"]),a("first_au_bank",r["first_au_bank"]),a("first_indusind",r.get("first_indusind",0) if hasattr(r,"keys") else 0)
    sc,sr,sy,sa,si=a("second_cash",r["second_cash"]),a("second_rbl",r["second_rbl"]),a("second_yes_bank",r["second_yes_bank"]),a("second_au_bank",r["second_au_bank"]),a("second_indusind",r.get("second_indusind",0) if hasattr(r,"keys") else 0)
    first=fc+fr+fy+fa+fi;second=sc+sr+sy+sa+si;other=a("other_payment_received",r["other_payment_received"])
    pkg=a("package_amount",r["package_amount"]);avp=a("after_visa_payment",r["after_visa_payment"])
    fs=request.form.get("first_payment_status",r["first_payment_status"] or "Pending");ss=request.form.get("second_payment_status",r["second_payment_status"] or "Pending")
    now=datetime.datetime.now().isoformat(timespec="seconds")
    fdt=request.form.get("first_payment_datetime","") or r["first_payment_datetime"] or (now if first else "")
    sdt=request.form.get("second_payment_datetime","") or r["second_payment_datetime"] or (now if second else "")
    odt=request.form.get("other_payment_datetime","") or r["other_payment_datetime"] or (now if other else "")
    total=first+second+other
    con.execute("""UPDATE client_cases SET package_amount=?,after_visa_payment=?,interaction_date=?,counselor_name=?,am_name=?,
      first_cash=?,first_rbl=?,first_yes_bank=?,first_au_bank=?,first_indusind=?,booking_amount=?,first_payment_status=?,first_payment_datetime=?,
      first_payment_method=?,second_cash=?,second_rbl=?,second_yes_bank=?,second_au_bank=?,second_indusind=?,second_payment=?,second_payment_status=?,second_payment_datetime=?,
      second_payment_method=?,other_payment_received=?,other_payment_datetime=?,other_payment_method=?,other_payment_bank=?,
      total_received=?,payment_status=?,remarks=?,updated_at=? WHERE id=?""",
      (pkg,avp,request.form.get("interaction_date",r["interaction_date"] or ""),request.form.get("counselor_name",r["counselor_name"] or ""),
       request.form.get("am_name",r["am_name"] or ""),fc,fr,fy,fa,fi,first,fs,fdt,
       request.form.get("first_payment_method",r["first_payment_method"] or "Net Banking"),
       sc,sr,sy,sa,si,second,ss,sdt,request.form.get("second_payment_method",r["second_payment_method"] or "Net Banking"),
       other,odt,request.form.get("other_payment_method",r["other_payment_method"] or ""),
       request.form.get("other_payment_bank",r["other_payment_bank"] or ""),
       total,ss,request.form.get("remarks",r["remarks"] or ""),now,case_id))
    _case_cols=_table_columns(con,"client_cases")
    if all(x in _case_cols for x in ("first_payment_bank_id","second_payment_bank_id","other_payment_bank_id")):
        def _oldv(name):
            try:return r[name]
            except Exception:return None
        def _bid_keep(name,old):
            v=(request.form.get(name) or "").strip()
            return int(v) if v.isdigit() else old
        con.execute("UPDATE client_cases SET first_payment_bank_id=?,second_payment_bank_id=?,other_payment_bank_id=? WHERE id=?",
                    (_bid_keep("first_payment_bank_id",_oldv("first_payment_bank_id")),
                     _bid_keep("second_payment_bank_id",_oldv("second_payment_bank_id")),
                     _bid_keep("other_payment_bank_id",_oldv("other_payment_bank_id")),case_id))
    if _table_exists(con,"client_payment_proofs"):
        for st,field in (("FIRST","first_payment_proof"),("SECOND","second_payment_proof"),("OTHER","other_payment_proof")):
            f=request.files.get(field)
            if f and f.filename:
                data=f.read()
                if len(data)<=10*1024*1024:
                    con.execute("""INSERT INTO client_payment_proofs(case_id,payment_stage,file_name,mime_type,file_data,uploaded_by_id,uploaded_by_name,uploaded_at)
                                   VALUES(?,?,?,?,?,?,?,?)""",(case_id,st,secure_filename(f.filename)[:180],f.mimetype or "application/octet-stream",data,u["id"],u["full_name"],now))
    con.commit()
    # AM ne submit kiya toh mark karo
    if u["role"]=="AM":
        try:
            if IS_POSTGRES:
                con.execute("UPDATE client_cases SET am_submitted=1, am_submitted_by=%s, am_submitted_at=%s WHERE id=%s",
                           (u["login_id"], now, case_id))
            else:
                con.execute("UPDATE client_cases SET am_submitted=1, am_submitted_by=?, am_submitted_at=? WHERE id=?",
                           (u["login_id"], now, case_id))
            con.commit()
        except Exception:
            try: con.rollback()
            except Exception: pass
    con.close();flash("✓ Enrollment details saved." + (" GM/MD will complete the payment details." if u['role']=='AM' else ""),"success");return redirect(url_for("cases"))

@app.route("/allocate", methods=["POST"])
@require_roles("MD","GM")
def allocate():
    u=current_user(); con=db()
    am_id=request.form.get("am_id")
    qty=int(request.form.get("qty","0") or 0)
    am=con.execute("SELECT * FROM users WHERE id=? AND role='AM' AND active=1",(am_id,)).fetchone()
    if not am:
        flash("Select a valid Assistant Manager","error")
    else:
        if u["role"]=="GM" and am["company_code"]!=u["company_code"]:
            flash("Cannot allocate outside your company","error")
        else:
            ids=[r["id"] for r in con.execute("SELECT id FROM leads WHERE company_code=? AND assigned_am IS NULL ORDER BY id LIMIT ?",(am["company_code"],qty)).fetchall()]
            now=datetime.datetime.now().isoformat(timespec="seconds")
            for lid in ids:
                con.execute("UPDATE leads SET assigned_am=?,assigned_at=?,assigned_by=? WHERE id=?",(am["id"],now,u["login_id"],lid))
            con.execute("INSERT INTO allocation_history(company_code,am_user_id,quantity,allocated_by,allocated_at) VALUES(?,?,?,?,?)",(am["company_code"],am["id"],len(ids),u["login_id"],now))
            con.commit()
            flash(f"{len(ids)} leads allocated to {am['full_name']}","success")
    con.close()
    return redirect(url_for("leads"))


def can_access_lead(u, lead):
    if not u or not lead: return False
    if u["role"]=="MD": return True
    if u["company_code"]!=lead["company_code"]: return False
    if u["role"] in ("GM","RECEPTION"): return True
    if u["role"]=="AM" and lead["assigned_am"]==u["id"]: return True
    return False


def ensure_enrollment_case_for_lead(con, lead_db_id, actor_login="SYSTEM"):
    """Create the GM/MD enrollment-account shell exactly once for an Enrolled lead."""
    lead=con.execute("""SELECT l.*,u.full_name am_name
                        FROM leads l LEFT JOIN users u ON u.id=l.assigned_am
                        WHERE l.id=?""",(lead_db_id,)).fetchone()
    if not lead or (lead["status"] or "").strip().lower()!="enrolled":
        return None, False
    existing=con.execute("SELECT * FROM client_cases WHERE lead_db_id=? ORDER BY id DESC LIMIT 1",(lead_db_id,)).fetchone()
    if existing:
        return existing, False

    company=lead["company_code"] or ""
    today=datetime.date.today().isoformat()
    now=datetime.datetime.now().isoformat(timespec="seconds")
    seq=con.execute("SELECT COUNT(*) c FROM client_cases WHERE company_code=?",(company,)).fetchone()["c"]+1
    case_id=f"{company}-CASE-{datetime.date.today().strftime('%y%m%d')}-{seq:04d}"
    while con.execute("SELECT id FROM client_cases WHERE case_id=?",(case_id,)).fetchone():
        seq+=1
        case_id=f"{company}-CASE-{datetime.date.today().strftime('%y%m%d')}-{seq:04d}"

    con.execute("""INSERT INTO client_cases(
        case_id,company_code,client_name,mobile,country,visa_type,enrollment_date,
        lead_db_id,counselor_name,am_name,interaction_date,package_amount,after_visa_payment,
        first_cash,first_rbl,first_yes_bank,first_au_bank,booking_amount,first_payment_status,first_payment_datetime,
        second_cash,second_rbl,second_yes_bank,second_au_bank,second_payment,second_payment_status,second_payment_datetime,
        other_payment_received,other_payment_datetime,total_received,payment_status,filing_status,remarks,created_by,created_at,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (case_id,company,lead["client_name"] or "",lead["mobile"] or "",lead["country"] or "",lead["visa_type"] or "",today,
         lead_db_id,"",lead["am_name"] or "","",0,0,
         0,0,0,0,0,"Pending","",
         0,0,0,0,0,"Pending","",
         0,"",0,"Pending","Documents Pending",lead["remarks"] or "",actor_login,now,now))
    return con.execute("SELECT * FROM client_cases WHERE lead_db_id=? ORDER BY id DESC LIMIT 1",(lead_db_id,)).fetchone(), True


@app.route("/lead/<int:lead_db_id>", methods=["GET","POST"])
@require_roles("MD","GM","AM")
def lead_profile(lead_db_id):
    u=current_user(); con=db()
    lead=con.execute("SELECT l.*,usr.full_name am_name,usr.login_id am_login FROM leads l LEFT JOIN users usr ON usr.id=l.assigned_am WHERE l.id=?",(lead_db_id,)).fetchone()
    if not can_access_lead(u,lead):
        con.close(); flash("Lead not found or access denied","error"); return redirect(url_for("dashboard"))
    if request.method=="POST":
        now=datetime.datetime.now().isoformat(timespec="seconds")
        vals=(request.form.get("client_name",""),request.form.get("mobile",""),request.form.get("alternate_mobile",""),request.form.get("email",""),
              request.form.get("date_of_birth",""),request.form.get("education",""),request.form.get("occupation",""),request.form.get("address",""),
              request.form.get("city",""),request.form.get("state",""),request.form.get("pin_code",""),request.form.get("passport_no",""),
              request.form.get("country",""),request.form.get("visa_type",""),request.form.get("travel_history",""),request.form.get("english_test",""),
              request.form.get("budget",""),request.form.get("preferred_call_time",""),request.form.get("status","New"),
              max(0,min(100,int(request.form.get("interest_score","0") or 0))),request.form.get("followup_date",""),
              request.form.get("remarks",""),lead_db_id)
        con.execute("""UPDATE leads SET client_name=?,mobile=?,alternate_mobile=?,email=?,date_of_birth=?,education=?,occupation=?,address=?,city=?,state=?,pin_code=?,passport_no=?,country=?,visa_type=?,travel_history=?,english_test=?,budget=?,preferred_call_time=?,status=?,interest_score=?,followup_date=?,remarks=? WHERE id=?""",vals)
        con.execute("INSERT INTO lead_activity(lead_id,status,followup_date,remarks,updated_by,updated_at) VALUES(?,?,?,?,?,?)",
                    (lead_db_id,request.form.get("status","New"),request.form.get("followup_date",""),request.form.get("remarks",""),u["login_id"],now))
        f=request.files.get("client_form")
        if f and f.filename:
            ext=f.filename.rsplit('.',1)[-1].lower() if '.' in f.filename else ''
            if ext not in ALLOWED_PHOTO_EXTENSIONS.union({'pdf'}):
                con.rollback(); con.close(); flash("Client form must be JPG, PNG, WEBP or PDF","error"); return redirect(url_for("lead_profile",lead_db_id=lead_db_id))
            data=f.read(); mime=f.mimetype or mimetypes.guess_type(f.filename)[0] or "application/octet-stream"
            con.execute("INSERT INTO lead_documents(lead_id,document_type,file_name,file_data,mime_type,uploaded_by,uploaded_at) VALUES(?,?,?,?,?,?,?)",
                        (lead_db_id,"Client Enquiry Form",secure_filename(f.filename),data,mime,u["login_id"],now))
        enrollment_created=False
        if (request.form.get("status","New") or "").strip().lower()=="enrolled":
            _case,enrollment_created=ensure_enrollment_case_for_lead(con,lead_db_id,u["login_id"])
            if enrollment_created:
                con.execute("INSERT INTO lead_activity(lead_id,status,followup_date,remarks,updated_by,updated_at) VALUES(?,?,?,?,?,?)",
                            (lead_db_id,"Enrolled","","Enrollment sent automatically to GM Accounts / Enrollment using same Lead ID.",u["login_id"],now))
        con.commit()
        if enrollment_created:
            flash("Client marked Enrolled. Same Lead ID has been sent to GM Enrollment / Accounts for payment completion.","success")
        else:
            flash("Client profile updated successfully","success")
        lead=con.execute("SELECT l.*,usr.full_name am_name,usr.login_id am_login FROM leads l LEFT JOIN users usr ON usr.id=l.assigned_am WHERE l.id=?",(lead_db_id,)).fetchone()
    docs=con.execute("SELECT id,document_type,file_name,mime_type,uploaded_by,uploaded_at FROM lead_documents WHERE lead_id=? ORDER BY id DESC",(lead_db_id,)).fetchall()
    activity=con.execute("SELECT * FROM lead_activity WHERE lead_id=? ORDER BY id DESC LIMIT 50",(lead_db_id,)).fetchall()
    con.close(); return render_template("lead_profile.html",u=u,lead=lead,docs=docs,activity=activity)

@app.route("/lead-document/<int:document_id>")
@require_roles("MD","GM","AM")
def lead_document(document_id):
    u=current_user(); con=db()
    row=con.execute("SELECT d.*,l.company_code,l.assigned_am FROM lead_documents d JOIN leads l ON l.id=d.lead_id WHERE d.id=?",(document_id,)).fetchone()
    if not row or (u["role"]!="MD" and (u["company_code"]!=row["company_code"] or (u["role"]=="AM" and row["assigned_am"]!=u["id"]))):
        con.close(); flash("Document not found or access denied","error"); return redirect(url_for("dashboard"))
    data=row["file_data"]; mime=row["mime_type"] or "application/octet-stream"; name=row["file_name"]
    con.close(); return send_file(io.BytesIO(data),mimetype=mime,download_name=name,as_attachment=False)

@app.route("/my-leads", methods=["GET","POST"])
@require_roles("AM")
def my_leads():
    u=current_user(); con=db()
    if request.method=="POST":
        lid=request.form.get("lead_id")
        con.execute("UPDATE leads SET status=?,followup_date=?,remarks=? WHERE id=? AND assigned_am=?",
                    (request.form.get("status"),request.form.get("followup_date"),request.form.get("remarks"),lid,u["id"]))
        created=False
        if str(lid).isdigit() and (request.form.get("status") or "").strip().lower()=="enrolled":
            _case,created=ensure_enrollment_case_for_lead(con,int(lid),u["login_id"])
        con.commit()
        flash("Lead updated"+(" • Enrollment sent to GM Accounts / Enrollment." if created else ""),"success")
    search_text=request.args.get("q","").strip()
    rank_filter=request.args.get("rank","").strip()
    status_filter=request.args.get("status","").strip()
    sort_filter=request.args.get("sort","interest_desc").strip()
    q="SELECT * FROM leads WHERE assigned_am=?"; params=[u["id"]]
    if search_text:
        q+=" AND (LOWER(COALESCE(client_name,'')) LIKE ? OR COALESCE(mobile,'') LIKE ? OR LOWER(COALESCE(lead_id,'')) LIKE ?)"
        like="%"+search_text.lower()+"%"; digits=re.sub(r"\D","",search_text); params.extend([like,"%"+digits+"%" if digits else like,like])
    if rank_filter=="hot": q+=" AND COALESCE(interest_score,0)>=71"
    elif rank_filter=="warm": q+=" AND COALESCE(interest_score,0) BETWEEN 50 AND 70"
    elif rank_filter=="cold": q+=" AND COALESCE(interest_score,0)<50"
    if status_filter: q+=" AND status=?"; params.append(status_filter)
    order_map={"interest_desc":"COALESCE(interest_score,0) DESC,id DESC","interest_asc":"COALESCE(interest_score,0) ASC,id DESC","newest":"id DESC","followup":"CASE WHEN COALESCE(followup_date,'')='' THEN 1 ELSE 0 END,followup_date ASC,id DESC"}
    q+=" ORDER BY "+order_map.get(sort_filter,order_map["interest_desc"])
    rows=con.execute(q,params).fetchall()
    con.close()
    filters={"q":search_text,"rank":rank_filter,"status":status_filter,"sort":sort_filter}
    return render_template("my_leads.html",u=u,rows=rows,filters=filters)


# ===== v3.98 UI PRIVACY + CASH BOOK =====

_GAUR_V337_TEMPLATES['finance_cash_book.html'] = r"""{% extends "base.html" %}{% block content %}
<h1 class="title">Daily Cash Book & Monthly Accounts</h1>
<p style="opacity:.78">MD CONFIDENTIAL • Daily Payment Inflow & Outflow • Monthly Account Register</p>
<div class="card no-print">
<form method="get" class="grid3">
<div><label>Date</label><input type="date" name="date" value="{{day}}"></div>
<div><label>Month</label><input type="month" name="month" value="{{month}}"></div>
<div><label>Company</label><select name="company"><option value="">Both Companies</option><option value="SCIC" {% if company=='SCIC' %}selected{% endif %}>ABC</option><option value="WWIC" {% if company=='WWIC' %}selected{% endif %}>XYZ</option></select></div>
<div><button class="btn">View Cash Book</button></div>
</form>
</div>

<div class="card" style="border-color:#e6b73f">
<h2 style="text-align:center;color:#e6b73f">{{day}} • Payment Inflow And Outflow</h2>
<div class="tablewrap"><table>
<thead><tr><th>Sr.</th><th>By</th><th>Inflow Particular</th><th>Cash / Bank</th><th>Amount</th><th>To</th><th>Outflow Particular</th><th>Cash / Bank</th><th>Amount</th></tr></thead>
<tbody>
{% set ins=daily|selectattr('direction','equalto','IN')|list %}
{% set outs=daily|rejectattr('direction','equalto','IN')|list %}
{% set n=[ins|length,outs|length]|max %}
{% for i in range(n) %}
<tr>
<td>{{i+1}}</td>
{% if i < ins|length %}{% set a=ins[i] %}<td>{{a.created_by_name}}</td><td>{{a.party_name}} / {{a.head}}</td><td>{{a.payment_mode}}{% if a.bank_name %} • {{a.bank_name}}{% endif %}</td><td><b style="color:#48d58b">₹{{"{:,.0f}".format(a.amount or 0)}}</b></td>{% else %}<td></td><td></td><td></td><td></td>{% endif %}
{% if i < outs|length %}{% set a=outs[i] %}<td>{{a.created_by_name}}</td><td>{{a.party_name}} / {{a.head}}</td><td>{{a.payment_mode}}{% if a.bank_name %} • {{a.bank_name}}{% endif %}</td><td><b style="color:#ff7b86">₹{{"{:,.0f}".format(a.amount or 0)}}</b></td>{% else %}<td></td><td></td><td></td><td></td>{% endif %}
</tr>{% endfor %}
{% if n==0 %}<tr><td colspan="9">No entry for this date.</td></tr>{% endif %}
</tbody>
<tfoot>
<tr><th colspan="4">Total Inflow</th><th>₹{{"{:,.0f}".format(daily_in)}}</th><th colspan="3">Total Outflow</th><th>₹{{"{:,.0f}".format(daily_out)}}</th></tr>
<tr><th colspan="8">Balance / Net Cash Flow</th><th>₹{{"{:,.0f}".format(daily_net)}}</th></tr>
</tfoot></table></div>
</div>

<div class="metrics">
<div class="metric">Total Monthly Receipt<b>₹{{"{:,.0f}".format(monthly_in)}}</b><small>{{month}}</small></div>
<div class="metric">Total Monthly Expenses<b>₹{{"{:,.0f}".format(monthly_out)}}</b><small>{{month}}</small></div>
<div class="metric">Net Profit / Cash Flow<b>₹{{"{:,.0f}".format(monthly_net)}}</b><small>{{month}}</small></div>
<div class="metric">Monthly Entries<b>{{monthly|length}}</b><small>Transactions</small></div>
</div>

<div class="card">
<h2 style="color:#e6b73f">Monthly Transaction Register</h2>
<div class="tablewrap"><table>
<thead><tr><th>Date</th><th>Voucher</th><th>Company</th><th>Type</th><th>Particular</th><th>Mode / Bank</th><th>Reference</th><th>Amount</th><th>Entered By</th><th>Date & Time Entered</th></tr></thead>
<tbody>{% for a in monthly %}<tr><td>{{a.txn_date}}</td><td>{{a.voucher_no}}</td><td>{{a.company_code}}</td><td>{{a.direction}}</td><td>{{a.party_name}} / {{a.head}}</td><td>{{a.payment_mode}}{% if a.bank_name %} • {{a.bank_name}}{% endif %}</td><td>{{a.reference_no or '-'}}</td><td>₹{{"{:,.0f}".format(a.amount or 0)}}</td><td>{{a.created_by_name}}</td><td>{{a.created_at}}</td></tr>{% else %}<tr><td colspan="10">No monthly entries.</td></tr>{% endfor %}</tbody>
</table></div></div>
<div class="no-print"><button class="toolbtn" onclick="window.print()">Print / Save PDF</button> <a class="toolbtn" href="{{url_for('finance_export',date_from=month+'-01',date_to=month+'-31',company=company)}}">Export Excel</a></div>
{% endblock %}"""

# GM gets a safe enrollment-only screen. MD keeps complete structured payment ledger.
_GAUR_V337_TEMPLATES['cases.html'] = r"""{% extends "base.html" %}{% block content %}
<h1 class="title">{% if u.role=='MD' %}Enrollment • Payment • Filing{% else %}Enrollment & Filing{% endif %}</h1>
<div class="card">
<h2 style="color:#e6b73f;margin-top:0">Create Client Case</h2>
<form method="post">
<div class="grid3">
{% if u.role=='MD' %}<div><label>Company</label><select name="company_code" required><option value="SCIC">ABC</option><option value="WWIC">XYZ</option></select></div>{% endif %}
<div><label>Client Name *</label><input name="client_name" required></div>
<div><label>Mobile</label><input name="mobile"></div>
<div><label>Country</label><input name="country"></div>
<div><label>Visa Type</label><input name="visa_type"></div>
<div><label>Enrollment Date</label><input type="date" name="enrollment_date"></div>
<div><label>Assigned Employee</label><select name="assigned_employee_id"><option value="">Unassigned</option>{% for e in staff %}<option value="{{e.id}}">{{e.full_name}} • {{e.company_code}}</option>{% endfor %}</select></div>
</div>

{% if u.role=='MD' %}
<h2 style="color:#e6b73f;margin-top:22px">Payment Details • MD Confidential</h2>
<div class="grid3">
<div><label>1. Package Amount</label><input type="number" step="0.01" min="0" name="package_amount" value="0"><label>Date & Time</label><input type="datetime-local" name="package_amount_datetime"></div>
<div><label>2. After Visa Payment</label><input type="number" step="0.01" min="0" name="after_visa_payment" value="0"><label>Date & Time</label><input type="datetime-local" name="after_visa_payment_datetime"></div>
<div><label>3. First Payment</label><input type="number" step="0.01" min="0" name="booking_amount" value="0"><label>Date & Time</label><input type="datetime-local" name="first_payment_datetime"></div>
<div><label>4. First Payment Status</label><select name="first_payment_status"><option>Received</option><option>Partially Received</option><option>Pending</option><option>Cancelled</option></select></div>
<div><label>5. 2nd Payment</label><input type="number" step="0.01" min="0" name="second_payment" value="0"><label>Date & Time</label><input type="datetime-local" name="second_payment_datetime"></div>
<div><label>6. 2nd Payment Status</label><select name="second_payment_status"><option>Pending</option><option>Received</option><option>Partially Received</option><option>Cancelled</option></select></div>
<div><label>7. Other Payment Received</label><input type="number" step="0.01" min="0" name="other_payment_received" value="0"><label>Date & Time</label><input type="datetime-local" name="other_payment_datetime"></div>
</div>
{% endif %}

<div class="grid3"><div><label>Filing Status</label><select name="filing_status"><option>Documents Pending</option><option>Documents Complete</option><option>Filed</option><option>Visa Granted</option><option>Visa Refused</option></select></div>
<div><label>Remarks</label><textarea name="remarks"></textarea></div></div>
<div style="margin-top:16px"><button class="btn">{% if u.role=='MD' %}Create Client Case{% else %}Create Enrollment{% endif %}</button></div>
</form></div>

{% if u.role=='MD' %}
<div class="card"><h2 style="color:#e6b73f">MD Payment & Enrollment Register</h2>
<div class="tablewrap"><table><thead><tr><th>Case / Client</th><th>Package</th><th>After Visa</th><th>First Payment</th><th>1st Status</th><th>1st Date/Time</th><th>2nd Payment</th><th>2nd Status</th><th>2nd Date/Time</th><th>Other Received</th><th>Other Date/Time</th><th>Actual Collection</th><th>Update</th></tr></thead>
<tbody>{% for r in rows %}<tr><td><b>{{r.case_id}}</b><br>{{r.client_name}}<br><small>{{r.employee_name or '-'}}</small></td>
<td>₹{{"{:,.0f}".format(r.package_amount or 0)}}<br><small>{{r.package_amount_datetime or '-'}}</small></td>
<td>₹{{"{:,.0f}".format(r.after_visa_payment or 0)}}<br><small>{{r.after_visa_payment_datetime or '-'}}</small></td>
<form method="post" action="{{url_for('update_case',case_id=r.id)}}">
<td><input type="number" step="0.01" name="booking_amount" value="{{r.booking_amount or 0}}"></td>
<td><select name="first_payment_status">{% for s in ['Received','Partially Received','Pending','Cancelled'] %}<option {% if (r.first_payment_status or 'Received')==s %}selected{% endif %}>{{s}}</option>{% endfor %}</select></td>
<td><input type="datetime-local" name="first_payment_datetime" value="{{(r.first_payment_datetime or '')[:16]}}"></td>
<td><input type="number" step="0.01" name="second_payment" value="{{r.second_payment or 0}}"></td>
<td><select name="second_payment_status">{% for s in ['Pending','Received','Partially Received','Cancelled'] %}<option {% if (r.second_payment_status or r.payment_status or 'Pending')==s %}selected{% endif %}>{{s}}</option>{% endfor %}</select></td>
<td><input type="datetime-local" name="second_payment_datetime" value="{{(r.second_payment_datetime or '')[:16]}}"></td>
<td><input type="number" step="0.01" name="other_payment_received" value="{{r.other_payment_received or 0}}"></td>
<td><input type="datetime-local" name="other_payment_datetime" value="{{(r.other_payment_datetime or '')[:16]}}"></td>
<td><b style="color:#48d58b">₹{{"{:,.0f}".format(r.total_received or 0)}}</b></td>
<td>
<input type="hidden" name="package_amount" value="{{r.package_amount or 0}}"><input type="hidden" name="package_amount_datetime" value="{{(r.package_amount_datetime or '')[:16]}}">
<input type="hidden" name="after_visa_payment" value="{{r.after_visa_payment or 0}}"><input type="hidden" name="after_visa_payment_datetime" value="{{(r.after_visa_payment_datetime or '')[:16]}}">
<input type="hidden" name="filing_status" value="{{r.filing_status or 'Documents Pending'}}"><input type="hidden" name="remarks" value="{{r.remarks or ''}}">
<button class="toolbtn">Save</button></td></form></tr>{% else %}<tr><td colspan="13">No records.</td></tr>{% endfor %}</tbody></table></div></div>
{% else %}
<div class="card"><h2 style="color:#e6b73f">Enrollment Records</h2>
<div class="tablewrap"><table><thead><tr><th>Case ID</th><th>Client</th><th>Country</th><th>Visa Type</th><th>Enrollment Date</th><th>Assigned Employee</th><th>Filing Status</th><th>Remarks</th></tr></thead>
<tbody>{% for r in rows %}<tr><td>{{r.case_id}}</td><td>{{r.client_name}}</td><td>{{r.country}}</td><td>{{r.visa_type}}</td><td>{{r.enrollment_date}}</td><td>{{r.employee_name or '-'}}</td><td>{{r.filing_status}}</td><td>{{r.remarks}}</td></tr>{% else %}<tr><td colspan="8">No enrollments.</td></tr>{% endfor %}</tbody></table></div></div>
{% endif %}
{% endblock %}"""

# Dashboard: revenue card exists only for MD.
_GAUR_V337_TEMPLATES['dashboard.html']=_GAUR_V337_TEMPLATES['dashboard.html'].replace(
    '<div class="metric">Revenue Received<b>₹{{"%.0f"|format(revenue or 0)}}</b></div>',
    '{% if u.role=="MD" %}<div class="metric">Revenue Received<b>₹{{"%.0f"|format(revenue or 0)}}</b></div>{% endif %}'
)

# Replace visible score wording to make the rule clear.
for _k in ('dashboard.html','performance_directory.html','employee_performance.html'):
    if _k in _GAUR_V337_TEMPLATES:
        _GAUR_V337_TEMPLATES[_k]=_GAUR_V337_TEMPLATES[_k].replace('Business Score','Enrollment Score').replace('Performance Score','Enrollment Score')
        _GAUR_V337_TEMPLATES[_k]=_GAUR_V337_TEMPLATES[_k].replace('AM PERFORMER OF THE MONTH','AM ENROLLMENT PERFORMER OF THE MONTH')

# Navigation: payment/finance options are MD confidential.
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
    '<a class="nav-main-link" href="{{ url_for(\'payment_accounts\') }}"><span class="nav-icon">🏦</span><span>Company Payment Accounts</span></a>',
    '{% if u.role=="MD" %}<a class="nav-main-link" href="{{ url_for(\'payment_accounts\') }}"><span class="nav-icon">🏦</span><span>Company Payment Accounts</span></a><a class="nav-main-link" href="{{url_for(\'finance_cash_book\')}}"><span class="nav-icon">📒</span><span>Daily Cash Book</span></a>{% endif %}'
)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
    '<a class="nav-main-link" href="{{ url_for(\'cases\') }}"><span class="nav-num">6</span><span>Enrolment Payment Report</span></a>\n<a class=\"nav-main-link\" href=\"{{ url_for("cases") }}\"><span class=\"nav-num\">{% set _pend=namespace(c=0) %}{% if pending_enrollments is defined %}{% set _pend.c=pending_enrollments %}{% endif %}{% if _pend.c > 0 %}<span style=\"background:#ff4444;color:#fff;border-radius:50%;padding:1px 5px;font-size:11px;font-weight:800\">{{_pend.c}}</span>{% endif %}</span><span>⚡ Pending Enrollments</span></a>',
    '<a class="nav-main-link" href="{{ url_for(\'cases\') }}"><span class="nav-num">6</span><span>{% if u.role=="MD" %}Enrolment Payment Report{% else %}Enrolment Report{% endif %}</span></a>'
)

# Competition top bar and popup: payment visible only when API says MD.
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
    "text('scicLiveScore',(s.ABC?.enrollments||0)+' • '+money(s.ABC?.revenue||0));text('wwicLiveScore',(s.XYZ?.enrollments||0)+' • '+money(s.XYZ?.revenue||0));",
    "text('scicLiveScore',(s.ABC?.enrollments||0)+(d.show_payment?' • '+money(s.ABC?.revenue||0):''));text('wwicLiveScore',(s.XYZ?.enrollments||0)+(d.show_payment?' • '+money(s.XYZ?.revenue||0):''));"
)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
    "text('competitionRevenue',money(e.revenue||0));",
    "text('competitionRevenue',e.payment_visible===false?'ENROLLMENT':money(e.revenue||0));"
)



_GAUR_V337_TEMPLATES['cases.html']=r"""{% extends "base.html" %}{% block content %}
<style>
.mod-head{display:flex;justify-content:space-between;cursor:pointer;padding:12px;background:#071d32;border:1px solid #315a7b;border-radius:10px;margin-top:10px}.mod-body{display:none;padding:12px;border:1px solid #315a7b}.mod-body.open{display:block}.paygrid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.due-row td{background:rgba(127,46,55,.35)!important}.clear-row td{background:rgba(23,79,52,.16)!important}.due-badge{background:#a52e3b;padding:5px 9px;border-radius:18px;font-weight:bold}.ok-badge{background:#174f34;padding:5px 9px;border-radius:18px;font-weight:bold}@media(max-width:900px){.paygrid{grid-template-columns:1fr 1fr}}
</style>
<h1 class="title">Enrollment • Client Accounts</h1><div class="card" style="padding:10px 14px;border-color:#315f85"><b style="color:#e6b73f">Lead ID Enrollment Ledger</b> • AM Enrolled status and GM payment completion use the same client record. {% if banks %}<span style="color:#55dca0">Bank Manager Connected: {{banks|length}} active account(s).</span>{% else %}<span style="color:#e6a34a">No active bank yet — ledger remains fully accessible.</span>{% endif %}</div>{% if not banks %}<div class="card" style="border-color:#d08b33;background:#2b2414"><b style="color:#f0bd43">No active Bank Manager accounts found.</b> Enrollment Report is still available. Add a bank from Accounts → Bank Manager to use bank selection.</div>{% endif %}<p style="opacity:.8">AM marks Enrolled → same Lead ID automatically appears here for GM payment completion. No duplicate client entry.</p>
<div class="card"><h2 style="color:#e6b73f">New Enrollment / Account Entry</h2><form method="post" enctype="multipart/form-data">
<div class="grid3">{% if u.role=='MD' %}<div><label>Company</label><select name="company_code" id="company"><option value="SCIC">ABC</option><option value="WWIC">XYZ</option></select></div>{% endif %}
<div><label>Link AM Lead Data</label><select name="lead_db_id" id="lead"><option value="">Manual Entry</option>{% for l in leads %}<option value="{{l.id}}" data-name="{{l.client_name}}" data-mobile="{{l.mobile}}" data-country="{{l.country}}" data-visa="{{l.visa_type}}" data-am="{{l.am_name or ''}}" data-company="{{l.company_code}}">{{l.lead_id}} • {{l.client_name}} • {{l.mobile}}</option>{% endfor %}</select></div>
<div><label>Enrollment Date</label><input type="date" name="enrollment_date"></div><div><label>Client Name</label><input name="client_name" id="nm" required></div><div><label>Contact No.</label><input name="mobile" id="mob"></div><div><label>Country / Program</label><input name="country" id="ctry"></div><div><label>Visa Type</label><input name="visa_type" id="visa"></div>
<div><label>Counselor</label><select name="counselor_name"><option value="">Select Counselor</option>{% for c in counselors %}<option>{{c.full_name}}</option>{% endfor %}</select></div><div><label>AM</label><input name="am_name" id="am" readonly></div><div><label>Date of Interaction Call</label><input type="date" name="interaction_date"></div><div><label>Package</label><input type="number" step=".01" name="package_amount" id="pkg" value="0"></div><div><label>After Visa</label><input type="number" step=".01" name="after_visa_payment" value="0"></div></div>

<div class="mod-head" onclick="tog('one')"><b>1st Payment — Cash + RBL + Yes Bank + AU Bank + IndusInd</b><span>Open ▾</span></div><div class="mod-body" id="one"><div class="paygrid">
{% for n,l in [('first_cash','Cash'),('first_rbl','RBL'),('first_yes_bank','Yes Bank'),('first_au_bank','AU Bank'),('first_indusind','IndusInd')] %}<div><label>{{l}}</label><input type="number" step=".01" name="{{n}}" class="f" value="0"></div>{% endfor %}
<div><label>1st Payment Total</label><input id="ft" readonly></div><div><label>Status</label><select name="first_payment_status"><option>Received</option><option>Partially Received</option><option>Pending</option></select></div><div><label>Date & Time</label><input type="datetime-local" name="first_payment_datetime"></div><div><label>Payment Screenshot / Proof</label><input type="file" name="first_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf"></div></div></div>
<div class="mod-head" onclick="tog('two')"><b>2nd Payment — Cash + RBL + Yes Bank + AU Bank + IndusInd</b><span>Open ▾</span></div><div class="mod-body" id="two"><div class="paygrid">
{% for n,l in [('second_cash','Cash'),('second_rbl','RBL'),('second_yes_bank','Yes Bank'),('second_au_bank','AU Bank'),('second_indusind','IndusInd')] %}<div><label>{{l}}</label><input type="number" step=".01" name="{{n}}" class="s" value="0"></div>{% endfor %}
<div><label>2nd Payment Total</label><input id="st" readonly></div><div><label>Status</label><select name="second_payment_status"><option>Pending</option><option>Received</option><option>Partially Received</option></select></div><div><label>Date & Time</label><input type="datetime-local" name="second_payment_datetime"></div><div><label>Payment Screenshot / Proof</label><input type="file" name="second_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf"></div></div></div>
<div class="mod-head" onclick="tog('final')"><b>Other Payment + Final Calculation</b><span>Open ▾</span></div><div class="mod-body" id="final"><div class="grid3"><div><label>Other Payment Received</label><input type="number" step=".01" name="other_payment_received" id="oth" value="0"></div><div><label>Other Date & Time</label><input type="datetime-local" name="other_payment_datetime"></div><div><label>Payment Screenshot / Proof</label><input type="file" name="other_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf"></div><div><label>Total Received</label><input id="tot" readonly></div><div><label>Balance Pending</label><input id="bal" readonly></div><div><label>Filing Status</label><select name="filing_status"><option>Documents Pending</option><option>Documents Complete</option><option>Filed</option></select></div><div><label>Remarks</label><textarea name="remarks"></textarea></div></div></div>
<button class="btn" style="margin-top:14px">Save Enrollment & Account</button></form></div>

<div class="card"><h2 style="color:#e6b73f">Client Ledger • Balance Pending Highlighted</h2><div class="tablewrap"><table><thead><tr><th>Client / Lead ID</th><th>GM Action</th><th>Counselor</th><th>AM</th><th>Package</th><th>1st Total</th><th>2nd Total</th><th>Other</th><th>Total Received</th><th>Balance Pending</th><th>Lead Data</th><th>Edit</th></tr></thead><tbody>
{% for r in rows %}{% set b=(r.package_amount or 0)-(r.total_received or 0) %}<tr class="{{'due-row' if b>0 else 'clear-row'}}" onclick="tog('r{{r.id}}')" style="cursor:pointer"><td><b>{{r.client_name}}</b><br><small>Lead: {{r.lead_code or 'Manual'}} • {{r.case_id}}</small></td><td>{% if (r.package_amount or 0)==0 and (r.total_received or 0)==0 %}<span class="due-badge">PENDING GM COMPLETION</span>{% else %}<span class="ok-badge">ACCOUNT UPDATED</span>{% endif %}</td><td><b>{{r.counselor_name or '-'}}</b></td><td><b>{{r.am_name or r.linked_am_name or '-'}}</b></td><td>₹{{"{:,.0f}".format(r.package_amount or 0)}}</td><td>₹{{"{:,.0f}".format(r.booking_amount or 0)}}</td><td>₹{{"{:,.0f}".format(r.second_payment or 0)}}</td><td>₹{{"{:,.0f}".format(r.other_payment_received or 0)}}</td><td><b>₹{{"{:,.0f}".format(r.total_received or 0)}}</b></td><td>{% if b>0 %}<span class="due-badge">₹{{"{:,.0f}".format(b)}} PENDING</span>{% else %}<span class="ok-badge">CLEAR</span>{% endif %}</td><td>{% if r.lead_db_id %}<a class="toolbtn" onclick="event.stopPropagation()" href="{{url_for('lead_profile',lead_db_id=r.lead_db_id)}}">Open Client ↗</a>{% else %}Not Linked{% endif %}</td>
<td>{% if r.lead_db_id %}<a class="toolbtn" onclick="event.stopPropagation()" href="{{url_for('edit_client_record',lead_db_id=r.lead_db_id)}}">✏️ Edit</a>{% else %}<span style="opacity:.55">No linked lead</span>{% endif %}</td></tr>
<tr><td colspan="12" style="padding:0"><div class="mod-body" id="r{{r.id}}"><form method="post" enctype="multipart/form-data" action="{{url_for('update_case',case_id=r.id)}}">
<div class="grid3"><div><b>Mobile:</b> {{r.mobile}}</div><div><b>Country:</b> {{r.country}}</div><div><b>Interaction:</b> {{r.interaction_date or '-'}}</div></div>
<div class="mod-head" onclick="event.stopPropagation();tog('f{{r.id}}')"><b>1st Payment Breakup</b><span>▾</span></div><div class="mod-body" id="f{{r.id}}"><div class="paygrid">{% for n,l,v in [('first_cash','Cash',r.first_cash),('first_rbl','RBL',r.first_rbl),('first_yes_bank','Yes Bank',r.first_yes_bank),('first_au_bank','AU Bank',r.first_au_bank),('first_indusind','IndusInd',r.first_indusind)] %}<div><label>{{l}}</label><input type="number" step=".01" name="{{n}}" value="{{v or 0}}"></div>{% endfor %}<div><label>Status</label><select name="first_payment_status"><option selected>{{r.first_payment_status}}</option><option>Received</option><option>Partially Received</option><option>Pending</option></select></div><div><label>Date & Time</label><input type="datetime-local" name="first_payment_datetime" value="{{(r.first_payment_datetime or '')[:16]}}"></div><div><label>New Screenshot / Proof</label><input type="file" name="first_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf">{% if r.first_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.first_proof_id)}}">View Existing Proof</a>{% endif %}</div></div></div>
<div class="mod-head" onclick="event.stopPropagation();tog('s{{r.id}}')"><b>2nd Payment Breakup</b><span>▾</span></div><div class="mod-body" id="s{{r.id}}"><div class="paygrid">{% for n,l,v in [('second_cash','Cash',r.second_cash),('second_rbl','RBL',r.second_rbl),('second_yes_bank','Yes Bank',r.second_yes_bank),('second_au_bank','AU Bank',r.second_au_bank),('second_indusind','IndusInd',r.second_indusind)] %}<div><label>{{l}}</label><input type="number" step=".01" name="{{n}}" value="{{v or 0}}"></div>{% endfor %}<div><label>Status</label><select name="second_payment_status"><option selected>{{r.second_payment_status}}</option><option>Received</option><option>Partially Received</option><option>Pending</option></select></div><div><label>Date & Time</label><input type="datetime-local" name="second_payment_datetime" value="{{(r.second_payment_datetime or '')[:16]}}"></div><div><label>New Screenshot / Proof</label><input type="file" name="second_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf">{% if r.second_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.second_proof_id)}}">View Existing Proof</a>{% endif %}</div></div></div>
<div class="grid3"><div><label>Package</label><input type="number" step=".01" name="package_amount" value="{{r.package_amount or 0}}"></div><div><label>After Visa</label><input type="number" step=".01" name="after_visa_payment" value="{{r.after_visa_payment or 0}}"></div><div><label>Other Received</label><input type="number" step=".01" name="other_payment_received" value="{{r.other_payment_received or 0}}"></div><div><label>Other Date/Time</label><input type="datetime-local" name="other_payment_datetime" value="{{(r.other_payment_datetime or '')[:16]}}"></div><div><label>New Screenshot / Proof</label><input type="file" name="other_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf">{% if r.other_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.other_proof_id)}}">View Existing Proof</a>{% endif %}</div><div><label>Counselor</label><input name="counselor_name" value="{{r.counselor_name or ''}}"></div><div><label>AM</label><input name="am_name" value="{{r.am_name or r.linked_am_name or ''}}"></div><div><label>Interaction</label><input type="date" name="interaction_date" value="{{r.interaction_date or ''}}"></div><div><label>Remarks</label><input name="remarks" value="{{r.remarks or ''}}"></div></div><button class="btn" style="margin-top:10px">Update Account</button></form></div></td></tr>{% else %}<tr><td colspan="12">No records.</td></tr>{% endfor %}</tbody></table></div></div>
<script>
function tog(id){let e=document.getElementById(id);if(e)e.classList.toggle('open')}function num(x){return Number(x||0)}
function calc(){let f=0,s=0;document.querySelectorAll('.f').forEach(x=>f+=num(x.value));document.querySelectorAll('.s').forEach(x=>s+=num(x.value));let o=num(document.getElementById('oth').value),p=num(document.getElementById('pkg').value);ft.value=f.toLocaleString('en-IN');st.value=s.toLocaleString('en-IN');tot.value=(f+s+o).toLocaleString('en-IN');bal.value=Math.max(p-f-s-o,0).toLocaleString('en-IN')}
document.querySelectorAll('.f,.s,#oth,#pkg').forEach(x=>x.addEventListener('input',calc));calc();
document.getElementById('lead').addEventListener('change',function(){let o=this.options[this.selectedIndex];if(!o.value)return;nm.value=o.dataset.name||'';mob.value=o.dataset.mobile||'';ctry.value=o.dataset.country||'';visa.value=o.dataset.visa||'';am.value=o.dataset.am||'';let c=document.getElementById('company');if(c)c.value=o.dataset.company||'SCIC'});
</script>{% endblock %}"""

# Put same Client Accounts option under Accounts for GM/MD.
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
'<a href="{{ url_for(\'finance_center\') }}">Daily Passbook</a>',
'<a href="{{ url_for(\'cases\') }}">Client Accounts / Enrollment</a><a href="{{ url_for(\'finance_center\') }}">Daily Passbook</a>'
)



_GAUR_V337_TEMPLATES['edit_client_record.html']=r"""{% extends "base.html" %}{% block content %}
<h1 class="title">Edit Client Details</h1>
<p style="opacity:.8">Editing the same AM-maintained client/lead record. No duplicate client is created.</p>

<div class="card">
<div class="grid3">
<div><b>Lead ID</b><br>{{lead.lead_id}}</div>
<div><b>Company</b><br>{{lead.company_code}}</div>
<div><b>Assistant Manager</b><br>{{lead.am_name or '-'}}</div>
</div>
</div>

<div class="card">
<form method="post">
<div class="grid3">
<div><label>Client Name *</label><input name="client_name" value="{{lead.client_name or ''}}" required></div>
<div><label>Contact No.</label><input name="mobile" value="{{lead.mobile or ''}}"></div>
<div><label>Country / Program</label><input name="country" value="{{lead.country or ''}}"></div>
<div><label>Visa Type</label><input name="visa_type" value="{{lead.visa_type or ''}}"></div>
<div><label>Client Status</label><select name="status">
{% for s in ['New','Follow-up','Interested','Hot','Warm','Cold','Enrolled','Not Interested','Closed'] %}
<option {% if (lead.status or 'New')==s %}selected{% endif %}>{{s}}</option>
{% endfor %}
</select></div>
<div><label>Next Follow-up Date</label><input type="date" name="followup_date" value="{{lead.followup_date or ''}}"></div>
<div style="grid-column:1/-1"><label>Remarks</label><textarea name="remarks">{{lead.remarks or ''}}</textarea></div>
</div>
<div style="margin-top:15px">
<button class="btn">Save Client Changes</button>
<a class="toolbtn" href="{{url_for('cases')}}">Cancel</a>
{% if case %}<span style="margin-left:12px;opacity:.75">Linked Enrollment: {{case.case_id}}</span>{% endif %}
</div>
</form>
</div>
{% endblock %}"""



_GAUR_V337_TEMPLATES['my_profile.html']=r'''{% extends "base.html" %}{% block content %}
<h1 class="title">My Profile</h1>
<div class="grid" style="grid-template-columns:1fr 1fr">
<div class="card"><h2 style="color:#e6b73f">Official Identity</h2>
<label>Login ID</label><div class="readonly-value">{{u['login_id']}}</div>
<label>Company</label><div class="readonly-value">{{u['company_code'] or 'THE IMMIGRATION MANAGEMENT SYSTEM'}}</div>
<label>Designation</label><div class="readonly-value">{{u['designation'] or u['role']}}</div>
<label>Reporting Manager</label><div class="readonly-value">{{reporting_manager['full_name'] if reporting_manager else '-'}}</div>
<p style="opacity:.72">Login ID, Company, Designation and Reporting Manager are Management-controlled.</p></div>
<div class="card"><h2 style="color:#e6b73f">Update My Professional Details</h2>
<form method="post" enctype="multipart/form-data">
<div class="grid">
<div><label>Official Mobile No.</label><input name="official_mobile" inputmode="numeric" maxlength="10" value="{{u['official_mobile'] or ''}}" placeholder="10-digit mobile number"></div>
<div><label>Official Email</label><input type="email" name="official_email" value="{{u['official_email'] or ''}}"></div>
<div><label>Department</label><input name="department" value="{{u['department'] or ''}}"></div>
<div><label>Branch / Office</label><input name="branch" value="{{u['branch'] or ''}}"></div>
</div>
<label>Profile Photo</label><input type="file" name="profile_photo" accept=".jpg,.jpeg,.png,.webp,image/*">
<label>Professional Headline</label><input name="profile_headline" maxlength="120" value="{{u['profile_headline'] or ''}}">
<label>About Me / Professional Bio</label><textarea name="profile_bio" rows="5" maxlength="600">{{u['profile_bio'] or ''}}</textarea>
<button class="btn">Save My Profile</button></form>
{% if u['photo_mime'] %}<form method="post" action="{{url_for('remove_my_profile_photo')}}" style="margin-top:10px"><button class="toolbtn">Remove Current Photo</button></form>{% endif %}
</div></div>{% endblock %}'''


# === v3.98 Responsive Mobile UI / PWA metadata ===
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
    '</head>',
    '''<link rel="manifest" href="/manifest.webmanifest">
<meta name="theme-color" content="#061629">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="THE IMMIGRATION MANAGEMENT SYSTEM">
<link rel="apple-touch-icon" href="/static/pwa/icon-192.png">
<style>
.mobile-nav-btn,.mobile-nav-overlay{display:none}
@media(max-width:900px){
 html,body{width:100%;max-width:100%;overflow-x:hidden}
 body{padding-top:env(safe-area-inset-top);padding-bottom:env(safe-area-inset-bottom)}
 .layout{display:block!important;min-height:100dvh}
 aside{position:fixed!important;left:0;top:0;bottom:0;z-index:1000002;width:min(86vw,320px);height:100dvh;overflow-y:auto;transform:translateX(-105%);transition:transform .25s ease;box-shadow:12px 0 40px rgba(0,0,0,.55);padding-top:calc(18px + env(safe-area-inset-top))}
 body.mobile-nav-open aside{transform:translateX(0)}
 .mobile-nav-overlay{display:block;position:fixed;inset:0;z-index:1000001;background:rgba(0,0,0,.62);opacity:0;visibility:hidden;transition:.25s ease}
 body.mobile-nav-open .mobile-nav-overlay{opacity:1;visibility:visible}
 main{padding:12px!important;padding-top:70px!important;width:100%;max-width:100vw}
 .mobile-nav-btn{display:flex;position:fixed;left:12px;top:calc(10px + env(safe-area-inset-top));z-index:1000003;align-items:center;gap:8px;min-height:46px;padding:8px 13px;border-radius:12px;border:1px solid #e6b73f;background:#071d32;color:#fff;font:700 16px "Times New Roman",serif;box-shadow:0 8px 24px rgba(0,0,0,.35)}
 .mobile-nav-btn b{color:#e6b73f;font-size:22px}
 .title{font-size:27px!important;line-height:1.08}
 .card{padding:13px!important;border-radius:12px}
 .grid,.grid3,.metrics,.kpi-grid,.performance-grid,.rank-grid,.donut-grid,.company-compare,.analytics-grid,.company-metrics,.filterbar,.paygrid,.summary-strip{grid-template-columns:1fr!important}
 .btn,.toolbtn,button,input,select,textarea{min-height:44px}
 input,select,textarea{font-size:16px!important}
 .btn,.toolbtn{padding:10px 12px!important}
 .tablewrap{overflow-x:auto!important;-webkit-overflow-scrolling:touch;border-radius:10px;max-width:100%}
 table{min-width:760px}
 th,td{padding:8px 7px;font-size:14px}
 .priority-strip{display:grid!important;grid-template-columns:auto 1fr;gap:8px!important}
 .priority-action{grid-column:1/-1;text-align:center}
 .competition-live-bar{left:8px!important;right:8px!important;top:8px!important}
 .celebrate-stage{max-height:92dvh;overflow-y:auto}
 .celebrate-details{grid-template-columns:1fr!important}
 .report-print-header,.report-print-footer{display:none!important}
}
@media(max-width:480px){
 main{padding-left:9px!important;padding-right:9px!important}
 .card{padding:11px!important}
 .title{font-size:24px!important}
 .metric b,.kpi-card .kpi-value{font-size:24px!important}
 .global-tools .toolbtn,.global-tools .btn{flex:1;text-align:center}
 table{min-width:700px}
}
@media(display-mode:standalone){
 body{overscroll-behavior-y:none}
 .mobile-nav-btn{top:calc(8px + env(safe-area-inset-top))}
}
</style>
</head>''',
    1
)


# Mobile nav button already embedded in base.html template above - no injection needed.

_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
    '</body>',
    '''<script>
(function(){
 const btn=document.getElementById('mobileNavBtn');
 const overlay=document.getElementById('mobileNavOverlay');
 function closeNav(){document.body.classList.remove('mobile-nav-open')}
 if(btn)btn.addEventListener('click',function(){document.body.classList.toggle('mobile-nav-open');});
 if(overlay)overlay.addEventListener('click',closeNav);
 document.querySelectorAll('aside a').forEach(function(a){a.addEventListener('click',closeNav);});
 window.addEventListener('resize',function(){if(window.innerWidth>900)closeNav();});
 if('serviceWorker' in navigator){
   window.addEventListener('load',function(){navigator.serviceWorker.register('/service-worker.js').catch(function(){});});
 }
})();
</script>
</body>''',
    1
)



# === v3.98 Login title branding only ===
for _login_key in ('login.html','md_login.html','gaur_login.html'):
    if _login_key in _GAUR_V337_TEMPLATES:
        _GAUR_V337_TEMPLATES[_login_key]=_GAUR_V337_TEMPLATES[_login_key].replace(
            '</head>',
            '''<style>
.login-card h1,.gaur-login-card h1,.login-title,#loginTitle{
  font-weight:900!important;
  letter-spacing:3px!important;
}
</style></head>''',
            1
        )



# === v3.98 Quick Multi Lead Allocation UI ===
_GAUR_V337_TEMPLATES['lead_control.html']=r"""{% extends "base.html" %}{% block content %}
<style>
.alloc-filter{position:sticky;top:0;z-index:40;background:#071a2d;border:1px solid #315a7b;border-radius:12px;padding:12px;margin-bottom:12px;box-shadow:0 8px 28px rgba(0,0,0,.26)}
.filter-grid{display:grid;grid-template-columns:repeat(6,minmax(130px,1fr));gap:8px;align-items:end}
.quick-row{display:flex;gap:7px;flex-wrap:wrap;margin-top:10px}
.quick-btn{border:1px solid #426f93;background:#0b2c49;color:white;padding:8px 11px;border-radius:9px;cursor:pointer;font-weight:700}
.quick-btn:hover{border-color:#e6b73f}.quick-btn.gold{background:#e6b73f;color:#071629;border-color:#e6b73f}.quick-btn.red{border-color:#9b3e4a}
.alloc-row{cursor:pointer;transition:.12s}.alloc-row:hover td{background:rgba(59,119,166,.12)}
.alloc-row.selected td{background:rgba(230,183,63,.18)!important;border-top-color:#e6b73f;border-bottom-color:#e6b73f}
.alloc-row .bigcheck{width:22px;height:22px;cursor:pointer}
.badge-mini{display:inline-block;padding:3px 7px;border-radius:11px;border:1px solid #315a7b;font-size:12px;white-space:nowrap}
.badge-unalloc{border-color:#b94150;background:rgba(185,65,80,.18)}
.sticky-allocate{position:fixed;left:290px;right:18px;bottom:14px;z-index:1000;background:rgba(5,24,42,.97);border:2px solid #e6b73f;border-radius:14px;padding:10px 12px;box-shadow:0 10px 35px rgba(0,0,0,.55);display:flex;gap:9px;align-items:center;flex-wrap:wrap}
.selected-count{font-size:17px;font-weight:900;color:#e6b73f;min-width:135px}.sticky-allocate select{min-width:210px}.sticky-space{height:92px}
.summary-cards{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;margin-bottom:12px}.sum-card{background:#0c2d49;border:1px solid #315a7b;border-radius:10px;padding:10px;text-align:center}.sum-card b{display:block;font-size:23px;color:#e6b73f}
.details-panel{margin-top:12px}.details-panel summary{cursor:pointer;color:#e6b73f;font-weight:800;padding:10px}
@media(max-width:1100px){.filter-grid{grid-template-columns:repeat(3,1fr)}}
@media(max-width:900px){.sticky-allocate{left:8px;right:8px;bottom:8px}.filter-grid{grid-template-columns:1fr 1fr}.summary-cards{grid-template-columns:repeat(3,1fr)}}
@media(max-width:560px){.filter-grid{grid-template-columns:1fr}.summary-cards{grid-template-columns:1fr 1fr}.sticky-allocate select,.sticky-allocate .btn{width:100%}}
</style>

<h1 class="title">Quick Lead Allocation</h1>
<p style="opacity:.78">Filter → multi-select → choose AM → Allocate. Row par click karke bhi lead select/deselect kar sakte hain.</p>

<div class="summary-cards">
<div class="sum-card">Filtered Leads<b>{{k.total}}</b></div>
<div class="sum-card">Unallocated<b>{{k.unallocated}}</b></div>
<div class="sum-card">Latest Batch<b>{{k.latest_batch}}</b></div>
<div class="sum-card">Old Data<b>{{k.old_data}}</b></div>
<div class="sum-card">Not Worked<b>{{k.not_worked}}</b></div>
<div class="sum-card">Overdue<b>{{k.overdue}}</b></div>
</div>

<div class="alloc-filter">
<form method="get" id="filterForm">
<div class="filter-grid">
{% if u.role=='MD' %}<div><label>Company</label><select name="company" onchange="this.form.submit()"><option value="">Both Companies</option><option value="SCIC" {% if f.company=='SCIC' %}selected{% endif %}>ABC</option><option value="WWIC" {% if f.company=='WWIC' %}selected{% endif %}>XYZ</option></select></div>{% endif %}
<div><label>AM</label><select name="am"><option value="">All AM / Allocation</option><option value="unallocated" {% if f.am=='unallocated' %}selected{% endif %}>⚠ Unallocated Only</option>{% for a in am_rows %}<option value="{{a.id}}" {% if f.am==a.id|string %}selected{% endif %}>{{a.full_name}} • {{a.company_code}}</option>{% endfor %}</select></div>
<div><label>Area</label><select name="area"><option value="">All Areas</option>{% for x in area_options %}<option value="{{x}}" {% if f.area==x %}selected{% endif %}>{{x}}</option>{% endfor %}</select></div>
<div><label>Location / City</label><select name="location"><option value="">All Locations</option>{% for x in location_options %}<option value="{{x}}" {% if f.location==x %}selected{% endif %}>{{x}}</option>{% endfor %}</select></div>
<div><label>Data Batch</label><select name="batch"><option value="">All Batches</option><option value="latest" {% if f.batch=='latest' %}selected{% endif %}>Latest Batch</option><option value="old" {% if f.batch=='old' %}selected{% endif %}>Old Data</option><option value="legacy" {% if f.batch=='legacy' %}selected{% endif %}>Legacy</option>{% for b in batches %}<option value="{{b.batch_code}}" {% if f.batch==b.batch_code %}selected{% endif %}>{{b.batch_code}} • {{b.company_code}}</option>{% endfor %}</select></div>
<div><label>Work Status</label><select name="work"><option value="">All</option><option value="worked" {% if f.work=='worked' %}selected{% endif %}>Worked</option><option value="not_worked" {% if f.work=='not_worked' %}selected{% endif %}>Not Worked</option><option value="overdue" {% if f.work=='overdue' %}selected{% endif %}>Overdue</option></select></div>
<div><label>Search Client / Mobile / Lead</label><input name="q" value="{{f.q}}" placeholder="Type name, mobile or Lead ID"></div>
<div><label>Upload From</label><input type="date" name="upload_from" value="{{f.upload_from}}"></div>
<div><label>Upload To</label><input type="date" name="upload_to" value="{{f.upload_to}}"></div>
<div><label>Allocated From</label><input type="date" name="alloc_from" value="{{f.alloc_from}}"></div>
<div><label>Allocated To</label><input type="date" name="alloc_to" value="{{f.alloc_to}}"></div>
<div><button class="btn" style="width:100%">Apply Filters</button></div>
</div>
<div class="quick-row">
<a class="toolbtn" href="{{url_for('lead_control_center')}}">Reset All Filters</a><a class="toolbtn" href="{{url_for('lead_daily_report')}}">📊 Daily Lead Report</a>
<button type="button" class="quick-btn gold" onclick="selectAllFiltered()">✓ Select All Filtered</button>
<button type="button" class="quick-btn" onclick="selectUnallocated()">Select Unallocated</button>
<button type="button" class="quick-btn" onclick="selectFirst(10)">First 10</button>
<button type="button" class="quick-btn" onclick="selectFirst(25)">First 25</button>
<button type="button" class="quick-btn" onclick="selectFirst(50)">First 50</button>
<button type="button" class="quick-btn red" onclick="clearSelection()">Clear Selection</button>
{% if u.role in ['MD','GM'] %}<button type="button" class="quick-btn bin397" onclick="moveSelectedToBin()">🗑 Move Selected to Bin</button>{% endif %}
{% if u.role=='MD' %}<a class="quick-btn recycle397" href="#mdRecycleBin">♻ MD Recycle Bin ({{archived|length}})</a>{% endif %}
</div>
</form>
</div>

<div style="margin:0 0 10px;padding:9px 12px;border:1px solid #315a7b;border-radius:9px;background:#0c2d49">
<b style="color:#e6b73f">Interest Priority:</b> 100% Interested leads automatically appear first → 0% leads last.
</div>
<form method="post" action="{{url_for('lead_bulk_action')}}" id="bulkForm">
<div class="card" style="padding:0">
<div class="tablewrap"><table id="allocationTable">
<thead><tr><th style="width:55px">Select</th><th>Lead ID</th><th>Client</th><th>Company</th><th>Area</th><th>Location</th><th>Country / Visa</th><th>Interest</th><th>Current AM</th><th>Batch</th><th>Work</th></tr></thead>
<tbody>
{% for r in rows %}
<tr class="alloc-row" data-unallocated="{{1 if not r.assigned_am else 0}}" data-company="{{r.company_code}}">
<td><input class="leadbox bigcheck" type="checkbox" name="lead_ids" value="{{r.id}}" aria-label="Select {{r.client_name}}"></td>
<td><b>{{r.lead_id}}</b></td>
<td><b>{{r.client_name or '-'}}</b><br><small>{{r.mobile or '-'}}</small></td>
<td>{{r.company_code}}</td>
<td>{{r.source or '-'}}</td>
<td>{{r.city or '-'}}</td>
<td>{{r.country or '-'}}<br><small>{{r.visa_type or '-'}}</small></td>
<td><b style="color:#e6b73f;font-size:16px">{{r.interest_score or 0}}%</b></td>
<td>{% if r.am_name %}<span class="badge-mini">{{r.am_name}}</span>{% else %}<span class="badge-mini badge-unalloc">UNALLOCATED</span>{% endif %}</td>
<td>{{r.upload_batch or 'Legacy'}}<br><small>{{r.data_age}}</small></td>
<td><span class="badge-mini">{{r.work_state}}</span></td>
</tr>
{% else %}
<tr><td colspan="11">No leads match the selected filters.</td></tr>
{% endfor %}
</tbody></table></div>
</div>

<div class="sticky-space"></div>
<div class="sticky-allocate">
<div class="selected-count"><span id="selectedCount">0</span> Leads Selected</div>
<select name="am_id" id="allocateAM">
<option value="">Choose Assistant Manager</option>
{% for a in am_rows %}<option value="{{a.id}}" data-company="{{a.company_code}}">{{a.full_name}} • {{a.company_code}}</option>{% endfor %}
</select>
<input name="reason" placeholder="Allocation / deletion reason (optional)" style="flex:1;min-width:180px">
<button class="btn" name="action" value="allocate" onclick="return confirmAllocation()">Allocate Now</button>
<button class="btn bin397" name="action" value="archive" onclick="return confirmArchive()">🗑 Move to Bin</button>
</div>
</form>

<details class="card details-panel">
<summary>AM Allocation Summary ▾</summary>
<div class="tablewrap"><table><thead><tr><th>AM</th><th>Company</th><th>Allocated</th><th>Worked</th><th>Not Worked</th><th>Hot</th><th>Positive</th><th>Overdue</th></tr></thead><tbody>
{% for a in summary %}<tr><td><b>{{a.full_name}}</b></td><td>{{a.company_code}}</td><td>{{a.allocated or 0}}</td><td>{{a.worked or 0}}</td><td>{{a.not_worked or 0}}</td><td>{{a.hot or 0}}</td><td>{{a.positive or 0}}</td><td>{{a.overdue or 0}}</td></tr>{% endfor %}
</tbody></table></div></details>

<details class="card details-panel">
<summary>Batch Summary ▾</summary>
<div class="tablewrap"><table><thead><tr><th>Batch</th><th>Company</th><th>Total</th><th>Allocated</th><th>Unallocated</th><th>Worked</th><th>Not Worked</th></tr></thead><tbody>
{% for b in batch_summary %}<tr><td>{{b.batch_code}}{% if b.is_latest %} • LATEST{% endif %}</td><td>{{b.company_code}}</td><td>{{b.total or 0}}</td><td>{{b.allocated or 0}}</td><td>{{b.unallocated or 0}}</td><td>{{b.worked or 0}}</td><td>{{b.not_worked or 0}}</td></tr>{% endfor %}
</tbody></table></div></details>


{% if u.role=='MD' %}
<div class="card md-recycle397" id="mdRecycleBin">
<h2>♻ MD Recycle Bin</h2>
<p style="opacity:.78"><b>MD CONTROL ONLY.</b> Leads deleted by MD or GM remain here and are hidden from all active lead screens. Restore returns them to active data. Permanent deletion cannot be undone.</p>
<form method="post" action="{{url_for('recycle_bin_bulk_action')}}">
<div class="quick-row" style="margin-bottom:10px">
<label class="toolbtn"><input type="checkbox" id="recycleAll397"> Select All</label>
<button class="btn" name="action" value="restore" onclick="return confirm('Restore selected lead(s) back to active data?')">♻ Restore Selected</button>
<button class="btn bin397" name="action" value="erase" onclick="return confirm('PERMANENT DELETE selected lead(s)? This cannot be undone.')">🗑 Permanently Delete Selected</button>
{% if archived %}<button class="btn bin397" name="action" value="empty" onclick="return confirm('EMPTY ENTIRE RECYCLE BIN? This permanently deletes all archived leads and cannot be undone.')">⚠ Empty Recycle Bin</button>{% endif %}
</div>
<div class="tablewrap"><table>
<thead><tr><th>Select</th><th>Lead ID</th><th>Client</th><th>Company</th><th>Batch</th><th>Deleted By</th><th>Deleted At</th><th>Reason</th></tr></thead>
<tbody>
{% for r in archived %}
<tr>
<td><input class="recyclebox397" type="checkbox" name="lead_ids" value="{{r.id}}"></td>
<td><b>{{r.lead_id}}</b></td><td>{{r.client_name or '-'}}</td><td>{{r.company_code}}</td><td>{{r.upload_batch or 'Legacy'}}</td>
<td>{{r.deleted_by or '-'}}</td><td>{{r.deleted_at or '-'}}</td><td>{{r.deletion_reason or '-'}}</td>
</tr>
{% else %}<tr><td colspan="8">Recycle Bin is empty.</td></tr>{% endfor %}
</tbody></table></div>
</form>
</div>
<script>
document.addEventListener('DOMContentLoaded',function(){
 const a=document.getElementById('recycleAll397');
 if(a)a.addEventListener('change',function(){document.querySelectorAll('.recyclebox397').forEach(x=>x.checked=this.checked);});
});
</script>
{% endif %}

<script>
const boxes=()=>Array.from(document.querySelectorAll('.leadbox'));
const rows=()=>Array.from(document.querySelectorAll('.alloc-row'));
function refreshSelected(){
 let n=0;
 rows().forEach(r=>{
   const cb=r.querySelector('.leadbox');
   if(cb&&cb.checked){n++;r.classList.add('selected')}else r.classList.remove('selected');
 });
 document.getElementById('selectedCount').textContent=n;
}
rows().forEach(r=>{
 r.addEventListener('click',e=>{
   if(e.target.closest('input,select,button,a,label'))return;
   const cb=r.querySelector('.leadbox');if(cb){cb.checked=!cb.checked;refreshSelected()}
 });
});
boxes().forEach(x=>x.addEventListener('change',refreshSelected));
function clearSelection(){boxes().forEach(x=>x.checked=false);refreshSelected()}
function selectAllFiltered(){boxes().forEach(x=>x.checked=true);refreshSelected()}
function selectUnallocated(){rows().forEach(r=>{const cb=r.querySelector('.leadbox');cb.checked=r.dataset.unallocated==='1'});refreshSelected()}
function selectFirst(n){clearSelection();boxes().slice(0,n).forEach(x=>x.checked=true);refreshSelected()}
function confirmAllocation(){
 const n=boxes().filter(x=>x.checked).length, am=document.getElementById('allocateAM');
 if(!n){alert('Please select at least one lead.');return false}
 if(!am.value){alert('Please choose an Assistant Manager.');return false}
 return confirm('Allocate '+n+' selected lead(s) to '+am.options[am.selectedIndex].text+'?');
}
function confirmArchive(){
 const n=boxes().filter(x=>x.checked).length;
 if(!n){alert('Please select at least one lead.');return false}
 return confirm('Move '+n+' selected lead(s) to Recycle Bin?\\n\\nOnly MD can restore or permanently delete them.');
}
function moveSelectedToBin(){
 const n=boxes().filter(x=>x.checked).length;
 if(!n){alert('Please select at least one lead first.');return false}
 if(!confirm('Move '+n+' selected lead(s) to Recycle Bin?\\n\\nGM will no longer see them. Only MD can restore them.')) return false;
 const f=document.getElementById('bulkForm');
 const h=document.createElement('input');h.type='hidden';h.name='action';h.value='archive';f.appendChild(h);
 f.submit();return false;
}
refreshSelected();
</script>
{% endblock %}"""



# === v3.98 THE GAUR Easy Chat ===


_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
    '<div class="chat-head"><strong>Chat Up</strong>',
    '<div class="chat-head"><strong>💬 THE GAUR Chat</strong>'
)


# v3.98 full chat viewport fix

# === v3.98 Compact Dashboard + Profile Settings ===
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</head>','<style>.dashboard-identity-card{display:none!important}.dashboard-heading-row{display:block!important}.dashboard-heading-row h1{margin-bottom:8px!important}</style></head>',1)
_GAUR_V337_TEMPLATES['my_profile.html']=_GAUR_V337_TEMPLATES['my_profile.html'].replace('{% endblock %}','\n<div class="card" style="margin-top:16px;border-color:#e6b73f">\n  <h2 style="color:#e6b73f;margin-top:0">🔐 Login ID & Password</h2>\n  <p style="opacity:.75">Login and password changes are available only here in My Profile. Current password verification is required.</p>\n  <form method="post" action="{{url_for(\'my_profile_account\')}}">\n    <div class="grid3">\n      <div><label>Login ID</label><input name="new_login_id" value="{{u[\'login_id\']}}" required autocomplete="username"></div>\n      <div><label>Current Password *</label><input type="password" name="current_password" required autocomplete="current-password" placeholder="Enter current password"></div>\n      <div><label>New Password</label><input type="password" name="new_password" minlength="8" autocomplete="new-password" placeholder="Leave blank if not changing"></div>\n      <div><label>Confirm New Password</label><input type="password" name="confirm_password" minlength="8" autocomplete="new-password" placeholder="Repeat new password"></div>\n    </div>\n    <button class="btn" style="margin-top:12px">Save Login & Security Changes</button>\n  </form>\n</div>\n{% endblock %}',1)

# === v3.98 Monthly Lead Overview at Dashboard Top ===
_dash=_GAUR_V337_TEMPLATES['dashboard.html']
_marker='<form method="get" class="month-toolbar no-print">'
_top='\n<div class="month-overview-top">\n  <div class="month-main-kpis">\n    <div class="month-kpi"><span>Total Leads of {{month_label}}</span><b>{{total_leads}}</b><small>Selected month</small></div>\n    <div class="month-kpi"><span>Allocated</span><b>{{allocated}}</b><small>Assigned to AMs</small></div>\n    <div class="month-kpi"><span>Unallocated</span><b>{{unallocated}}</b><small>Pending allocation</small></div>\n  </div>\n\n  {% if u[\'role\'] in [\'MD\',\'GM\'] %}\n  <div class="am-month-grid">\n    {% for a in allocation_summary %}\n    <div class="am-month-card">\n      <div class="am-card-head">\n        <div>\n          <div class="am-card-name">{{a[\'full_name\']}}</div>\n          <small>{{a[\'company_code\']}} • Assistant Manager</small>\n        </div>\n        <div class="am-conversion">{{a[\'positive_pct\'] or 0}}%</div>\n      </div>\n      <div class="am-card-stats">\n        <div><span>Interested / Allocated</span><b>{{a[\'positive_count\'] or 0}} / {{a[\'allocated_count\'] or 0}}</b></div>\n        <div><span>Hot Leads</span><b>{{a[\'hot_count\'] or 0}}</b></div>\n        <div><span>Enrolled</span><b>{{a[\'enrolled_count\'] or 0}}</b></div>\n      </div>\n      <div class="am-progress"><i style="width:{{a[\'positive_pct\'] or 0}}%"></i></div>\n      <small class="am-note">Conversion = Interested (50%+) ÷ Allocated</small>\n    </div>\n    {% else %}\n    <div class="card">No Assistant Managers created yet.</div>\n    {% endfor %}\n  </div>\n  {% endif %}\n</div>\n'
if _marker in _dash and 'month-overview-top' not in _dash:
    _dash=_dash.replace(_marker,_top+_marker,1)
# Hide/remove the old general metric row and old AM table to prevent duplicate information.
_old_start='<div class="metrics">\n<div class="metric">Total Leads<b>{{total_leads}}</b></div>'
_old_end='{% endif %}\n\n<div class="card"><h2>Lead Interest Snapshot</h2>'
if _old_start in _dash and _old_end in _dash:
    _s=_dash.index(_old_start); _e=_dash.index(_old_end,_s)
    _dash=_dash[:_s]+'<div class="card"><h2>Lead Interest Snapshot</h2>'+_dash[_e+len(_old_end):]
_GAUR_V337_TEMPLATES['dashboard.html']=_dash
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</head>','\n<style>\n.month-overview-top{margin:8px 0 14px}\n.month-main-kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:12px}\n.month-kpi{background:#0d3151;border:1px solid #315f85;border-radius:14px;padding:15px 18px;box-shadow:0 7px 20px rgba(0,0,0,.18)}\n.month-kpi span{font-weight:800}.month-kpi b{display:block;color:#f3bd2f;font-size:32px;line-height:1.05;margin-top:6px}.month-kpi small{opacity:.65}\n.am-month-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:10px}\n.am-month-card{background:linear-gradient(145deg,#0a2945,#0d3557);border:1px solid #315f85;border-radius:14px;padding:13px;box-shadow:0 6px 20px rgba(0,0,0,.18)}\n.am-card-head{display:flex;justify-content:space-between;gap:10px;align-items:flex-start}.am-card-name{font-size:18px;font-weight:900;color:#f3bd2f}.am-card-head small{opacity:.7}\n.am-conversion{font-size:25px;font-weight:900;color:#55dca0;white-space:nowrap}\n.am-card-stats{display:grid;grid-template-columns:1.5fr .7fr .7fr;gap:7px;margin-top:12px}\n.am-card-stats>div{background:#071d32;border:1px solid #244866;border-radius:9px;padding:8px}.am-card-stats span{display:block;font-size:11px;opacity:.7}.am-card-stats b{font-size:17px}\n.am-progress{height:7px;background:#071629;border-radius:8px;overflow:hidden;margin-top:10px}.am-progress i{display:block;height:100%;background:#55dca0;border-radius:8px}\n.am-note{display:block;margin-top:7px;opacity:.58}\n@media(max-width:900px){.month-main-kpis{grid-template-columns:1fr 1fr 1fr}.am-card-stats{grid-template-columns:1fr 1fr 1fr}}\n@media(max-width:600px){.month-main-kpis{grid-template-columns:1fr}.am-month-grid{grid-template-columns:1fr}}\n</style></head>\n',1)

# === v3.98 Gallery-style Lead Multi Selection ===
if 'lead_control.html' in _GAUR_V337_TEMPLATES:
    _lc=_GAUR_V337_TEMPLATES['lead_control.html']
    _lc=_lc.replace('<button type="button" class="quick-btn gold" onclick="selectAllFiltered()">✓ Select All Filtered</button>','<button type="button" class="quick-btn gold" id="selectionModeBtn" onclick="toggleSelectionMode()">Selection Mode</button>\\n<button type="button" class="quick-btn gold" onclick="selectAllFiltered()">✓ Select All Filtered</button>',1)
    _lc=_lc.replace('</style>','.selection-mode-banner{display:none;position:sticky;top:72px;z-index:55;background:#e6b73f;color:#071629;border-radius:10px;padding:9px 12px;margin-bottom:8px;font-weight:900;box-shadow:0 6px 18px rgba(0,0,0,.25)}\n.selection-mode-on .selection-mode-banner{display:flex;justify-content:space-between;align-items:center;gap:10px}\n.selection-mode-on .alloc-row{user-select:none;-webkit-user-select:none;touch-action:pan-y}\n.selection-mode-on .alloc-row td{cursor:crosshair}\n.selection-mode-on .alloc-row.selected td{background:rgba(230,183,63,.26)!important}\n.selection-live-count{background:#071629;color:#e6b73f;border-radius:18px;padding:5px 10px;white-space:nowrap}\n@media(max-width:900px){.selection-mode-banner{top:64px}.selection-mode-on .alloc-row td{cursor:pointer}}\n</style>',1)
    _lc=_lc.replace('<form method="post" action="{{url_for(\'lead_bulk_action\')}}" id="bulkForm">','<div class="selection-mode-banner" id="selectionModeBanner">\n<div>Swipe / drag over leads to select multiple</div>\n<div class="selection-live-count"><span id="selectionModeCount">0</span> selected</div>\n</div>\n<form method="post" action="{{url_for(\'lead_bulk_action\')}}" id="bulkForm">',1)
    _lc=_lc.replace("rows().forEach(r=>{\n r.addEventListener('click',e=>{\n   if(e.target.closest('input,select,button,a,label'))return;\n   const cb=r.querySelector('.leadbox');if(cb){cb.checked=!cb.checked;refreshSelected()}\n });\n});\nboxes().forEach(x=>x.addEventListener('change',refreshSelected));","let selectionMode=false,dragSelecting=false,dragValue=true,lastTouchedRow=null,longPressTimer=null,longPressStart=null;\nfunction setRowSelected(r,val){const cb=r&&r.querySelector('.leadbox');if(!cb)return;cb.checked=!!val;refreshSelected()}\nfunction enterSelectionMode(){selectionMode=true;document.body.classList.add('selection-mode-on');const b=document.getElementById('selectionModeBtn');if(b)b.textContent='Selection Mode ON'}\nfunction exitSelectionMode(){selectionMode=false;dragSelecting=false;lastTouchedRow=null;document.body.classList.remove('selection-mode-on');const b=document.getElementById('selectionModeBtn');if(b)b.textContent='Selection Mode'}\nfunction toggleSelectionMode(){selectionMode?exitSelectionMode():enterSelectionMode()}\nrows().forEach(r=>{\n r.addEventListener('click',e=>{\n   if(e.target.closest('input,select,button,a,label'))return;\n   if(!selectionMode){const cb=r.querySelector('.leadbox');if(cb){cb.checked=!cb.checked;refreshSelected()}return}\n   setRowSelected(r,!r.querySelector('.leadbox').checked);\n });\n r.addEventListener('mousedown',e=>{\n   if(!selectionMode||e.button!==0||e.target.closest('input,select,button,a,label'))return;\n   e.preventDefault();dragSelecting=true;dragValue=!r.querySelector('.leadbox').checked;setRowSelected(r,dragValue);\n });\n r.addEventListener('mouseenter',()=>{if(selectionMode&&dragSelecting)setRowSelected(r,dragValue)});\n r.addEventListener('touchstart',e=>{\n   if(e.touches.length!==1)return;\n   longPressStart={x:e.touches[0].clientX,y:e.touches[0].clientY,row:r};\n   clearTimeout(longPressTimer);\n   longPressTimer=setTimeout(()=>{enterSelectionMode();dragSelecting=true;dragValue=!r.querySelector('.leadbox').checked;setRowSelected(r,dragValue);if(navigator.vibrate)navigator.vibrate(35)},420);\n },{passive:true});\n r.addEventListener('touchmove',e=>{\n   if(!longPressStart)return;\n   const t=e.touches[0];\n   if(!selectionMode&&Math.hypot(t.clientX-longPressStart.x,t.clientY-longPressStart.y)>12)clearTimeout(longPressTimer);\n   if(selectionMode&&dragSelecting){\n      const el=document.elementFromPoint(t.clientX,t.clientY);\n      const rr=el&&el.closest('.alloc-row');\n      if(rr&&rr!==lastTouchedRow){setRowSelected(rr,dragValue);lastTouchedRow=rr}\n   }\n },{passive:true});\n r.addEventListener('touchend',()=>{clearTimeout(longPressTimer);dragSelecting=false;longPressStart=null;lastTouchedRow=null},{passive:true});\n});\ndocument.addEventListener('mouseup',()=>{dragSelecting=false;lastTouchedRow=null});\nboxes().forEach(x=>x.addEventListener('change',()=>{if(x.checked)enterSelectionMode();refreshSelected()}));",1)
    _lc=_lc.replace("document.getElementById('selectedCount').textContent=n;","document.getElementById('selectedCount').textContent=n;const smc=document.getElementById('selectionModeCount');if(smc)smc.textContent=n;",1)
    _lc=_lc.replace('function clearSelection(){boxes().forEach(x=>x.checked=false);refreshSelected()}','function clearSelection(){boxes().forEach(x=>x.checked=false);refreshSelected();exitSelectionMode()}',1)
    _GAUR_V337_TEMPLATES['lead_control.html']=_lc



# === v3.98 LinkedIn-style Personal Header integrated into Main Dashboard ===
if 'dashboard.html' in _GAUR_V337_TEMPLATES:
    _d=_GAUR_V337_TEMPLATES['dashboard.html']
    _old_start='<div class="dashboard-heading-row">'
    _old_end='</div>\\n\\n<form method="get" class="month-toolbar no-print">'
    if _old_start in _d and _old_end in _d:
        _s=_d.index(_old_start)
        _e=_d.index(_old_end,_s)
        _personal=r"""
<div class="pro-dash-nav">
  <a class="active" href="{{url_for('dashboard')}}">🏠 Home</a>
  <a href="{{url_for('dashboard')}}">📊 Dashboard</a>
  <a href="{{url_for('my_profile')}}">👤 Me</a>
</div>

<div class="pro-dashboard-hero">
  {% set theme=u['profile_theme'] or 'royal' %}
  <div class="pro-cover theme-{{theme}}" {% if u['profile_cover_mime'] %}style="background-image:url('{{url_for('user_cover',user_id=u['id'])}}')"{% endif %}>
    <a class="pro-edit-cover" href="{{url_for('my_profile')}}">✎ Edit Profile & Theme</a>
  </div>
  <div class="pro-profile-body">
    <div class="pro-profile-photo-wrap">
      {% if u['photo_mime'] %}<img class="pro-profile-photo" src="{{url_for('user_photo',user_id=u['id'],v='359')}}">{% else %}<div class="pro-profile-photo-ph">{{(u['full_name'] or '?')[0]|upper}}</div>{% endif %}
    </div>
    <div class="pro-profile-info">
      <div class="pro-profile-name">{{u['full_name']}}</div>
      <div class="pro-profile-headline">{{u['profile_headline'] or role_display(u['role'],u['designation'])}}</div>
      <div class="pro-profile-meta">
        <span>{{u['company_code'] or 'THE IMMIGRATION MANAGEMENT SYSTEM'}}</span>
        {% if u['department'] %}<span>{{u['department']}}</span>{% endif %}
        {% if u['branch'] %}<span>{{u['branch']}}</span>{% endif %}
        {% if u['official_mobile'] %}<span>{{u['official_mobile']}}</span>{% endif %}
      </div>
    </div>
    <div class="pro-profile-actions">
      <a class="toolbtn" href="{{url_for('my_profile')}}">Manage My Profile</a>
    </div>
  </div>
</div>

<div class="pro-dashboard-title">
  <h1 class="title">{% if u['role']=='MD' %}Executive Dashboard{% elif u['role']=='GM' %}General Manager Dashboard{% elif u['role']=='AM' %}Assistant Manager Dashboard{% else %}Reception Dashboard{% endif %}</h1>
  <div id="dueFollowupChip" class="due-followup-chip" style="display:none">🔔 <span id="dueFollowupCount">0</span> Follow-ups Due</div>
</div>

<form method="get" class="month-toolbar no-print">"""
        _d=_d[:_s]+_personal+_d[_e+len(_old_end):]

    # Add dashboard reminder popup JS before endblock.
    _d=_d.replace(
      '{% endblock %}',
      r"""
<div id="desktopReminderStack" class="desktop-reminder-stack"></div>
<script>
(function(){
  const appearance="{{u['appearance_mode'] or 'dark'}}";
  if(appearance==='light') document.body.classList.add('user-light');
  if(appearance==='system' && window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) document.body.classList.add('user-light');

  function escapeHtml(s){return String(s||'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
  function showReminder(item,index){
    const stack=document.getElementById('desktopReminderStack'); if(!stack)return;
    const box=document.createElement('div');
    box.className='desktop-reminder '+(item.overdue?'overdue':'');
    box.style.animationDelay=(index*120)+'ms';
    box.innerHTML='<button class="desktop-reminder-close">×</button>'+
      '<div class="desktop-reminder-app">🔔 IMS • Follow-up Reminder</div>'+
      '<div class="desktop-reminder-person">'+escapeHtml(item.client_name)+'</div>'+
      '<div class="desktop-reminder-text">'+(item.overdue?'OVERDUE • ':'Due Today • ')+escapeHtml(item.followup_date)+'<br>'+
      'Interest '+escapeHtml(item.interest)+'%'+(item.am_name?' • '+escapeHtml(item.am_name):'')+'</div>'+
      '<div class="desktop-reminder-actions"><a href="'+item.url+'">Open Client</a><button type="button">Dismiss</button></div>';
    box.querySelector('.desktop-reminder-close').onclick=()=>box.remove();
    box.querySelector('.desktop-reminder-actions button').onclick=()=>box.remove();
    stack.appendChild(box);
    if(index===0) setTimeout(()=>{if(box.isConnected)box.classList.add('visible')},50);
    else setTimeout(()=>{if(box.isConnected)box.classList.add('visible')},index*250+50);
  }

  fetch('{{url_for("dashboard_reminders_api")}}',{credentials:'same-origin'})
    .then(r=>r.ok?r.json():null).then(data=>{
      if(!data)return;
      const chip=document.getElementById('dueFollowupChip'), count=document.getElementById('dueFollowupCount');
      if(data.count>0){count.textContent=data.count;chip.style.display='inline-flex'}
      const key='gaur-reminder-popup-'+new Date().toISOString().slice(0,10)+'-{{u["id"]}}';
      if(data.items && data.items.length && sessionStorage.getItem(key)!=='1'){
        data.items.slice(0,3).forEach((x,i)=>showReminder(x,i));
        sessionStorage.setItem(key,'1');
      }
    }).catch(()=>{});
})();
</script>
{% endblock %}""",
      1
    )
    _GAUR_V337_TEMPLATES['dashboard.html']=_d

# Main dashboard styling
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
  '</head>',
  r"""<style>
.pro-dash-nav{display:flex;justify-content:flex-end;gap:8px;flex-wrap:wrap;margin:2px 0 12px}
.pro-dash-nav a{position:relative;padding:9px 13px;border:1px solid #315f85;border-radius:10px;background:#0b2946;color:#fff;text-decoration:none;font-weight:800}
.pro-dash-nav a:hover,.pro-dash-nav a.active{border-color:#e6b73f}.pro-badge{position:absolute;right:-6px;top:-8px;background:#d64755;color:#fff;border-radius:12px;min-width:20px;height:20px;padding:0 5px;display:flex;align-items:center;justify-content:center;font-size:11px}
.pro-dashboard-hero{border:1px solid #315f85;border-radius:18px;overflow:hidden;background:#0b2946;margin-bottom:14px;box-shadow:0 10px 28px rgba(0,0,0,.2)}
.pro-cover{height:150px;position:relative;background:linear-gradient(135deg,#0d3151,#163b5a);background-size:cover;background-position:center}.pro-cover.theme-flowers{background:radial-gradient(circle at 20% 30%,#ffd2e8 0 7%,transparent 8%),radial-gradient(circle at 75% 35%,#ffeaa6 0 7%,transparent 8%),linear-gradient(135deg,#7d375f,#233a63)}.pro-cover.theme-nature{background:linear-gradient(135deg,#416f54,#9bc78b)}.pro-cover.theme-minimal{background:linear-gradient(135deg,#dfe7ef,#f7f8fa)}.pro-cover.theme-ocean{background:linear-gradient(135deg,#0b4a6f,#20a6c9,#7ad8dc)}.pro-cover.theme-sunset{background:linear-gradient(135deg,#7f3765,#e3765b,#f4c06a)}
.pro-edit-cover{position:absolute;right:14px;top:14px;background:rgba(5,22,39,.82);border:1px solid #e6b73f;color:#fff!important;border-radius:9px;padding:7px 10px;text-decoration:none;font-weight:800}
.pro-profile-body{display:grid;grid-template-columns:130px 1fr auto;gap:16px;align-items:end;padding:0 20px 18px}.pro-profile-photo-wrap{margin-top:-62px}.pro-profile-photo,.pro-profile-photo-ph{width:122px;height:122px;border-radius:50%;border:5px solid #071629;background:#0b2946}.pro-profile-photo{object-fit:cover}.pro-profile-photo-ph{display:flex;align-items:center;justify-content:center;font-size:46px;font-weight:900;color:#e6b73f}.pro-profile-name{font-size:28px;font-weight:900;color:#e6b73f}.pro-profile-headline{font-size:17px;font-weight:700;margin-top:3px}.pro-profile-meta{display:flex;gap:7px;flex-wrap:wrap;margin-top:8px}.pro-profile-meta span{border:1px solid #315f85;background:#071d32;border-radius:16px;padding:4px 8px;font-size:12px}.pro-profile-actions{display:flex;gap:8px;flex-wrap:wrap}
.pro-dashboard-title{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.due-followup-chip{align-items:center;gap:5px;padding:8px 11px;border:1px solid #d64755;background:#351820;color:#fff;border-radius:18px;font-weight:900}
.desktop-reminder-stack{position:fixed;right:20px;bottom:20px;z-index:1000010;width:min(365px,calc(100vw - 24px));display:flex;flex-direction:column;gap:10px;pointer-events:none}
.desktop-reminder{pointer-events:auto;position:relative;background:#1d2228;color:#fff;border:1px solid #4b535c;border-radius:12px;padding:13px 14px;box-shadow:0 16px 40px rgba(0,0,0,.48);transform:translateX(115%);opacity:0;transition:.28s ease}.desktop-reminder.visible{transform:translateX(0);opacity:1}.desktop-reminder.overdue{border-left:4px solid #d64755}.desktop-reminder-close{position:absolute;right:8px;top:6px;background:transparent;border:0;color:#aaa;font-size:20px;cursor:pointer}.desktop-reminder-app{font-size:12px;color:#9dd0ff;font-weight:800;margin-bottom:8px}.desktop-reminder-person{font-size:18px;font-weight:900}.desktop-reminder-text{margin-top:5px;color:#d8dee5;line-height:1.4}.desktop-reminder-actions{display:flex;gap:8px;margin-top:10px}.desktop-reminder-actions a,.desktop-reminder-actions button{border:1px solid #66717d;background:#2b333c;color:#fff;border-radius:8px;padding:7px 10px;text-decoration:none;cursor:pointer;font-weight:700}
body.user-light{background:#eef2f6!important;color:#132238!important}body.user-light main{background:#eef2f6!important}body.user-light .pro-dashboard-hero,body.user-light .card,body.user-light .metric,body.user-light .am-month-card{background:#fff!important;color:#132238!important;border-color:#cbd6e2!important}body.user-light .pro-profile-meta span{background:#f3f6f8!important;color:#132238!important}
@media(max-width:900px){.pro-dash-nav{justify-content:flex-start;overflow-x:auto;flex-wrap:nowrap;padding-bottom:3px}.pro-dash-nav a{white-space:nowrap}.pro-profile-body{grid-template-columns:100px 1fr}.pro-profile-photo,.pro-profile-photo-ph{width:92px;height:92px}.pro-profile-photo-wrap{margin-top:-46px}.pro-profile-actions{grid-column:1/-1}.pro-cover{height:120px}.desktop-reminder-stack{right:8px;bottom:84px;width:calc(100vw - 16px)}}
</style></head>""",
  1
)



_GAUR_V337_TEMPLATES['accounts_reports.html']=r"""{% extends "base.html" %}{% block content %}
<style>
.acc-tabs{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px}.acc-tabs a{padding:9px 13px;border:1px solid #315f85;border-radius:10px;color:#fff;text-decoration:none;background:#0b2946;font-weight:800}.acc-tabs a.active{background:#e6b73f;color:#071629}
.acc-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.acc-kpi{background:#0d3151;border:1px solid #315f85;border-radius:13px;padding:13px}.acc-kpi b{display:block;color:#e6b73f;font-size:26px}.bank-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:9px}.bank-card{background:#081e34;border:1px solid #315f85;border-radius:11px;padding:11px}.bank-card b{display:block;color:#55dca0;font-size:21px}.report-filters{display:flex;gap:8px;flex-wrap:wrap;align-items:end}.inmoney{color:#55dca0;font-weight:900}.outmoney{color:#ff7b86;font-weight:900}@media(max-width:900px){.acc-kpis{grid-template-columns:1fr 1fr}}@media(max-width:560px){.acc-kpis{grid-template-columns:1fr}}
</style>
<h1 class="title">📒 Accounts • Inflow & Outflow</h1>
<p style="opacity:.75">Enrollment payments + manual accounts entries • Bank-wise • Method-wise • Daily / Weekly / Monthly / Yearly.</p>
<div class="acc-tabs">
<a class="{{'active' if mode=='daily' else ''}}" href="{{url_for('accounts_report_center',mode='daily',company=company)}}">Daily</a>
<a class="{{'active' if mode=='weekly' else ''}}" href="{{url_for('accounts_report_center',mode='weekly',company=company)}}">Weekly</a>
<a class="{{'active' if mode=='monthly' else ''}}" href="{{url_for('accounts_report_center',mode='monthly',company=company)}}">Monthly</a>
<a class="{{'active' if mode=='yearly' else ''}}" href="{{url_for('accounts_report_center',mode='yearly',company=company)}}">Yearly</a>
<a href="{{url_for('finance_banks')}}">🏦 Bank Manager</a><a href="{{url_for('finance_center')}}">➕ Accounts Entry</a>
</div>
<div class="card no-print"><form method="get" class="report-filters"><input type="hidden" name="mode" value="{{mode}}">
{% if mode in ['daily','weekly'] %}<div><label>Date</label><input type="date" name="date" value="{{start}}"></div>{% endif %}
{% if mode=='monthly' %}<div><label>Month</label><input type="month" name="month" value="{{start[:7]}}"></div>{% endif %}
{% if mode=='yearly' %}<div><label>Year</label><input type="number" name="year" value="{{start[:4]}}" min="2020" max="2100"></div>{% endif %}
{% if u.role=='MD' %}<div><label>Company</label><select name="company"><option value="">Both Companies</option>{% for c in companies %}<option value="{{c}}" {% if company==c %}selected{% endif %}>{{c}}</option>{% endfor %}</select></div>{% endif %}
<div><button class="btn">View Report</button></div></form></div>
<h2 style="color:#e6b73f">{{label}}</h2>
<div class="acc-kpis"><div class="acc-kpi">Total Receipt<b>₹{{"{:,.0f}".format(total_in)}}</b></div><div class="acc-kpi">Total Expenses / Refunds<b>₹{{"{:,.0f}".format(total_out)}}</b></div><div class="acc-kpi">Net Cash Flow<b>₹{{"{:,.0f}".format(net)}}</b></div><div class="acc-kpi">Receipt Entries<b>{{inflow|length}}</b></div></div>
<div class="card"><h2 style="color:#e6b73f">🏦 Bank / Cash-wise Receipt</h2><div class="bank-grid">{% for bank,total in bank_totals %}<div class="bank-card"><span>{{bank}}</span><b>₹{{"{:,.0f}".format(total)}}</b></div>{% else %}<div>No receipts.</div>{% endfor %}</div></div>
<div class="card"><h2 style="color:#e6b73f">💳 Payment Method-wise Receipt</h2><div class="bank-grid">{% for method,total in method_totals %}<div class="bank-card"><span>{{method}}</span><b>₹{{"{:,.0f}".format(total)}}</b></div>{% else %}<div>No methods recorded.</div>{% endfor %}</div></div>
<div class="card" style="border-color:#e6b73f"><h2 style="text-align:center;color:#e6b73f;font-size:20px;letter-spacing:1px">💰 Payment Inflow And Outflow • {{label}}</h2>
<div class="tablewrap"><table style="width:100%;border-collapse:collapse">
<thead>
<tr style="background:#0a2540">
  <th colspan="5" style="text-align:center;color:#48d58b;padding:8px;border-bottom:2px solid #1e4060;font-size:13px;letter-spacing:1px">📥 INFLOW — RECEIPTS</th>
  <th colspan="4" style="text-align:center;color:#ff7a7a;padding:8px;border-bottom:2px solid #1e4060;font-size:13px;letter-spacing:1px">📤 OUTFLOW — EXPENSES</th>
</tr>
<tr style="background:#071d32">
  <th style="padding:9px 10px;color:#e6b73f;font-size:12px;border-bottom:1px solid #1e4060">Date / Co.</th>
  <th style="padding:9px 10px;color:#e6b73f;font-size:12px;border-bottom:1px solid #1e4060">Client / Particular</th>
  <th style="padding:9px 10px;color:#e6b73f;font-size:12px;border-bottom:1px solid #1e4060">Bank</th>
  <th style="padding:9px 10px;color:#e6b73f;font-size:12px;border-bottom:1px solid #1e4060">Method</th>
  <th style="padding:9px 10px;color:#48d58b;font-size:12px;border-bottom:1px solid #1e4060;text-align:right">Amount</th>
  <th style="padding:9px 10px;color:#e6b73f;font-size:12px;border-bottom:1px solid #1e4060">Date / Co.</th>
  <th style="padding:9px 10px;color:#e6b73f;font-size:12px;border-bottom:1px solid #1e4060">Particular</th>
  <th style="padding:9px 10px;color:#e6b73f;font-size:12px;border-bottom:1px solid #1e4060">Bank / Method</th>
  <th style="padding:9px 10px;color:#ff7a7a;font-size:12px;border-bottom:1px solid #1e4060;text-align:right">Amount</th>
</tr>
</thead>
<tbody>{% for p in paired %}<tr style="border-bottom:1px solid #0d2e4a">
{% if p['in'] %}
  <td style="padding:9px 10px;font-size:13px;color:#ccc;vertical-align:top">
    <b style="color:#fff">{{p['in']['date']}}</b><br>
    <span style="background:{{'#1a3a5c' if p['in']['company']=='SCIC' else '#1a3a2a'}};color:{{'#8fc8ff' if p['in']['company']=='SCIC' else '#48d58b'}};padding:2px 7px;border-radius:10px;font-size:11px;font-weight:600">{{p['in']['company']}}</span>
  </td>
  <td style="padding:9px 10px;font-size:13px;vertical-align:top">
    <b style="color:#e6b73f">{{p['in']['party']}}</b><br>
    <small style="color:#8899aa">{{p['in']['stage']}} • {{p['in']['ref']}}</small>
  </td>
  <td style="padding:9px 10px;font-size:13px;color:#8fc8ff;vertical-align:top">{{p['in']['bank']}}</td>
  <td style="padding:9px 10px;vertical-align:top">
    <span style="background:#0d2e4a;color:#f0a030;padding:3px 8px;border-radius:10px;font-size:12px;font-weight:600">{{p['in']['method']}}</span>
  </td>
  <td style="padding:9px 10px;text-align:right;vertical-align:top">
    <b style="color:#48d58b;font-size:14px">₹{{"{:,.0f}".format(p['in']['amount'])}}</b>
  </td>
{% else %}<td></td><td></td><td></td><td></td><td></td>{% endif %}
{% if p['out'] %}
  <td style="padding:9px 10px;font-size:13px;color:#ccc;vertical-align:top">
    <b style="color:#fff">{{p['out']['date']}}</b><br>
    <span style="background:#2a1a1a;color:#ff9999;padding:2px 7px;border-radius:10px;font-size:11px;font-weight:600">{{p['out']['company']}}</span>
  </td>
  <td style="padding:9px 10px;font-size:13px;vertical-align:top">
    <b style="color:#ffb3b3">{{p['out']['party']}}</b><br>
    <small style="color:#8899aa">{{p['out']['stage']}} • {{p['out']['ref']}}</small>
  </td>
  <td style="padding:9px 10px;font-size:12px;color:#8fc8ff;vertical-align:top">{{p['out']['bank']}}<br><small style="color:#8899aa">{{p['out']['method']}}</small></td>
  <td style="padding:9px 10px;text-align:right;vertical-align:top">
    <b style="color:#ff7a7a;font-size:14px">₹{{"{:,.0f}".format(p['out']['amount'])}}</b>
  </td>
{% else %}<td></td><td></td><td></td><td></td>{% endif %}
</tr>{% else %}<tr><td colspan="9" style="text-align:center;padding:20px;color:#8899aa">No transactions.</td></tr>{% endfor %}</tbody>
<tfoot>
  <tr style="background:#0a2540;border-top:2px solid #1e4060">
    <th colspan="4" style="padding:10px;color:#48d58b;font-size:13px;text-align:right">✅ Total Receipt</th>
    <th style="padding:10px;color:#48d58b;font-size:15px;text-align:right">₹{{"{:,.0f}".format(total_in)}}</th>
    <th colspan="3" style="padding:10px;color:#ff7a7a;font-size:13px;text-align:right">❌ Total Expenses</th>
    <th style="padding:10px;color:#ff7a7a;font-size:15px;text-align:right">₹{{"{:,.0f}".format(total_out)}}</th>
  </tr>
  <tr style="background:#071d32;border-top:1px solid #1e4060">
    <th colspan="8" style="padding:10px;color:#e6b73f;font-size:14px;text-align:right">💰 Net Cash Flow</th>
    <th style="padding:10px;color:#e6b73f;font-size:16px;font-weight:800;text-align:right">₹{{"{:,.0f}".format(net)}}</th>
  </tr>
</tfoot>
</table></div></div><div class="no-print"><button class="toolbtn" onclick="window.print()">Print / Save PDF</button></div>
{% endblock %}"""

if 'cases.html' in _GAUR_V337_TEMPLATES:
    _c=_GAUR_V337_TEMPLATES['cases.html']
    _c=_c.replace('<div><label>1st Payment Total</label><input id="ft" readonly></div>',
                  '<div><label>1st Payment Total</label><input id="ft" readonly></div><div><label>How Received</label><select name="first_payment_method"><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option></select></div>',1)
    _c=_c.replace('<div><label>2nd Payment Total</label><input id="st" readonly></div>',
                  '<div><label>2nd Payment Total</label><input id="st" readonly></div><div><label>How Received</label><select name="second_payment_method"><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option></select></div>',1)
    _c=_c.replace('<div><label>Other Payment Received</label><input type="number" step=".01" name="other_payment_received" id="oth" value="0"></div>',
                  '<div><label>Other Payment Received</label><input type="number" step=".01" name="other_payment_received" id="oth" value="0"></div><div><label>Other Payment Bank / Cash</label><input name="other_payment_bank" placeholder="Cash / Bank name"></div><div><label>How Received</label><select name="other_payment_method"><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option><option>Other</option></select></div>',1)
    _c=_c.replace('<div><label>Status</label><select name="first_payment_status">',
                  '<div><label>How Received</label><select name="first_payment_method"><option selected>{{r.first_payment_method or "Net Banking"}}</option><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option></select></div><div><label>Status</label><select name="first_payment_status">',1)
    _c=_c.replace('<div><label>Status</label><select name="second_payment_status">',
                  '<div><label>How Received</label><select name="second_payment_method"><option selected>{{r.second_payment_method or "Net Banking"}}</option><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option></select></div><div><label>Status</label><select name="second_payment_status">',1)
    _c=_c.replace('<div><label>Other Received</label><input type="number" step=".01" name="other_payment_received" value="{{r.other_payment_received or 0}}"></div>',
                  '<div><label>Other Received</label><input type="number" step=".01" name="other_payment_received" value="{{r.other_payment_received or 0}}"></div><div><label>Other Bank / Cash</label><input name="other_payment_bank" value="{{r.other_payment_bank or \'\'}}"></div><div><label>Other Method</label><select name="other_payment_method"><option selected>{{r.other_payment_method or "Other"}}</option><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option><option>Other</option></select></div>',1)
    _GAUR_V337_TEMPLATES['cases.html']=_c

_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace(
    '<a href="{{ url_for(\'cases\') }}">Client Accounts / Enrollment</a>',
    '<a href="{{url_for(\'accounts_report_center\')}}">📒 Accounts Reports</a><a href="{{url_for(\'finance_banks\')}}">🏦 Bank Manager</a><a href="{{ url_for(\'cases\') }}">Client Accounts / Enrollment</a>',1)



# === v3.98 Bank Manager Add/Delete for MD + GM ===
_GAUR_V337_TEMPLATES['finance_banks.html']=r"""{% extends "base.html" %}{% block content %}
<style>
.bank-actions{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0 16px}
.bank-form-grid{display:grid;grid-template-columns:repeat(4,minmax(180px,1fr));gap:10px}
.bank-card-actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}
.danger-btn{background:#7d2632!important;border-color:#b94353!important;color:#fff!important}
.bank-manager-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(310px,1fr));gap:12px}
.payment-bank-card{background:#0b2946;border:1px solid #315f85;border-radius:14px;padding:15px}
.payment-bank-card h3{color:#e6b73f;margin:0 0 10px}
.payment-details{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.payment-details div{background:#071d32;border:1px solid #294f6d;border-radius:9px;padding:8px}
.payment-details small{display:block;opacity:.65}.payment-details b{display:block;margin-top:2px}
.verified-account{display:inline-block;border:1px solid #3c9c70;color:#70e1a6;border-radius:16px;padding:4px 8px;font-size:12px;margin-bottom:10px}
@media(max-width:1000px){.bank-form-grid{grid-template-columns:1fr 1fr}}
@media(max-width:600px){.bank-form-grid,.payment-details{grid-template-columns:1fr}}
</style>

<h1 class="title">🏦 Bank Manager</h1>
<div class="bank-actions">
<a class="toolbtn" href="{{url_for('accounts_report_center')}}">← Accounts Reports</a>
<a class="toolbtn" href="{{url_for('finance_center')}}">Daily Passbook / Accounts Entry</a>
<a class="toolbtn" href="{{url_for('payment_accounts')}}">Employee Payment View</a>
{% if u.role=='MD' %}<a class="toolbtn" href="{{url_for('bank_share_history')}}">Share History</a>{% endif %}
</div>

<div class="card">
<h2 style="color:#e6b73f;margin-top:0">➕ Add Bank / Cash Account</h2>
<p style="opacity:.72">
{% if u.role=='GM' %}You can add and manage accounts for {{u.company_code}} only.{% else %}MD can manage accounts for both companies.{% endif %}
</p>
<form method="post" class="bank-form-grid">
{% if u.role=='MD' %}
<div><label>Company</label><select name="company_code" required><option value="SCIC">ABC</option><option value="WWIC">XYZ</option></select></div>
{% else %}
<input type="hidden" name="company_code" value="{{u.company_code}}">
<div><label>Company</label><input value="{{u.company_code}}" readonly></div>
{% endif %}
<div><label>Bank Name *</label><input name="bank_name" required placeholder="RBL / Yes Bank / AU Bank / ICICI..."></div>
<div><label>Account Nickname</label><input name="account_nickname" placeholder="Main Collection / Client Payments"></div>
<div><label>Account Name *</label><input name="account_name" required></div>
<div><label>Account Number *</label><input name="account_number" required></div>
<div><label>IFSC Code</label><input name="ifsc_code"></div>
<div><label>Branch</label><input name="branch"></div>
<div><label>City</label><input name="city"></div>
<div><label>UPI ID / QR ID</label><input name="upi_id" placeholder="company@bank"></div>
<div><label>Account Type</label><select name="account_type"><option>Current</option><option>Savings</option><option>Cash</option></select></div>
<div><label>Opening Balance</label><input type="number" step="0.01" name="opening_balance" value="0"></div>
<div><label>Bank RM Name</label><input name="rm_name"></div>
<div><label>RM Mobile</label><input name="rm_mobile"></div>
<div style="display:flex;align-items:end"><button class="btn" style="width:100%">Add Account</button></div>
</form>
</div>

<div class="card">
<h2 style="color:#e6b73f">Approved / Active Accounts</h2>
<div class="bank-manager-grid">
{% for b in banks %}
<div class="payment-bank-card">
<h3>{{b.bank_name}} • {{b.account_nickname or b.account_name}}</h3>
<span class="verified-account">✓ Active Company Account</span>
<div class="payment-details">
<div><small>Company</small><b>{{b.company_code}}</b></div>
<div><small>Account Name</small><b>{{b.account_name}}</b></div>
<div><small>Account Number</small><b>{% if u.role=='MD' %}{{b.account_number}}{% else %}••••••••{{b.account_last4}}{% endif %}</b></div>
<div><small>IFSC</small><b>{{b.ifsc_code or '-'}}</b></div>
<div><small>Branch / City</small><b>{{b.branch or '-'}} {{b.city or ''}}</b></div>
<div><small>UPI / QR ID</small><b>{{b.upi_id or '-'}}</b></div>
<div><small>Account Type</small><b>{{b.account_type or '-'}}</b></div>
<div><small>Opening Balance</small><b>₹{{"{:,.0f}".format(b.opening_balance or 0)}}</b></div>
</div>
<div class="bank-card-actions">
<form method="post" action="{{url_for('finance_bank_delete',bank_id=b.id)}}" onsubmit="return confirm('Remove '+{{b.bank_name|tojson}}+' account ending {{b.account_last4}}? Historical transactions will NOT be deleted.')">
<button class="toolbtn danger-btn">🗑 Remove Account</button>
</form>
</div>
</div>
{% else %}
<div style="opacity:.7">No approved bank accounts added. Use the form above to add your first account.</div>
{% endfor %}
</div>
</div>
{% endblock %}"""



_GAUR_V337_TEMPLATES['enrollment_report_safe.html']=r"""{% extends "base.html" %}{% block content %}
<style>
.er-head{display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap}
.er-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:12px 0}
.er-kpi{background:#0d3151;border:1px solid #315f85;border-radius:13px;padding:12px}.er-kpi b{display:block;color:#e6b73f;font-size:25px}
.er-row{border:1px solid #315f85;border-radius:13px;margin:10px 0;overflow:hidden;background:#0a243d}
.er-summary{display:grid;grid-template-columns:1.5fr .8fr .8fr .8fr .8fr auto;gap:8px;align-items:center;padding:12px}
.er-payment{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;padding:12px;border-top:1px solid #315f85;background:#071d32}
.er-pay-card{border:1px solid #294f6d;border-radius:11px;padding:10px}.er-pay-card h3{color:#e6b73f;margin:0 0 7px}.pending{color:#ffbd67}.received{color:#55dca0}.er-note{padding:10px 12px;background:#2a2211;border:1px solid #8b6923;border-radius:10px;margin:10px 0}.er-filters{display:grid;grid-template-columns:1fr 1fr 2fr auto auto;gap:9px;align-items:end;margin:12px 0}.er-filters label{display:block;font-size:12px;opacity:.75;margin-bottom:4px}.er-filter-card{background:#0b2946;border:1px solid #315f85;border-radius:13px;padding:12px}.search-help{font-size:11px;opacity:.62;margin-top:4px}@media(max-width:900px){.er-filters{grid-template-columns:1fr 1fr}.er-search-wide{grid-column:1/-1}}@media(max-width:560px){.er-filters{grid-template-columns:1fr}.er-search-wide{grid-column:auto}}
@media(max-width:1000px){.er-summary{grid-template-columns:1fr 1fr 1fr}.er-payment{grid-template-columns:1fr}}
@media(max-width:600px){.er-kpis{grid-template-columns:1fr 1fr}.er-summary{grid-template-columns:1fr}}
</style>

<div class="er-head">
<div><h1 class="title">Enrollment Payment Report</h1><div style="display:inline-block;margin-bottom:8px;padding:4px 8px;border:1px solid #55dca0;border-radius:12px;color:#55dca0;font-size:12px">SAFE ROUTE • /accounts/client-ledger</div>
<p style="opacity:.75;margin-top:-8px">AM Enrolled → same Lead ID → GM payment completion.</p></div>
<div><a class="toolbtn" href="{{url_for('finance_banks')}}">🏦 Bank Manager</a> <a class="toolbtn" href="{{url_for('accounts_report_center')}}">📒 Accounts Reports</a>{% if u.role=='MD' %} <a class="toolbtn" href="{{url_for('enrollment_recycle_bin')}}">🗑 Enrollment Recycle Bin</a>{% endif %}</div>
</div>

<div class="er-filter-card no-print">
<form method="get" action="{{url_for('enrollment_report')}}" class="er-filters">
  <div>
    <label>Company</label>
    <select name="company" {% if u.role=='GM' %}disabled{% endif %}>
      {% if u.role=='MD' %}<option value="">All Companies</option>{% endif %}
      {% for c in company_options %}
      <option value="{{c}}" {% if company_filter==c %}selected{% endif %}>{{'ABC' if c=='SCIC' else ('XYZ' if c=='WWIC' else c)}}</option>
      {% endfor %}
    </select>
    {% if u.role=='GM' %}<input type="hidden" name="company" value="{{u.company_code}}">{% endif %}
  </div>
  <div>
    <label>Assistant Manager</label>
    <select name="am">
      <option value="">All AMs</option>
      {% for a in am_options %}
      <option value="{{a.id}}" {% if am_filter|string==a.id|string %}selected{% endif %}>{{a.full_name}}{% if u.role=='MD' %} • {{a.company_code}}{% endif %}</option>
      {% endfor %}
    </select>
  </div>
  <div class="er-search-wide">
    <label>Search Client</label>
    <input type="search" name="q" value="{{search_q}}" placeholder="Client name / Mobile no. / Passport no. / Lead ID">
    <div class="search-help">Search works with partial name, mobile number, passport number, Lead ID or case ID.</div>
  </div>
  <div><button class="btn" type="submit">🔎 Search / Filter</button></div>
  <div><a class="toolbtn" href="{{url_for('enrollment_report')}}">Clear</a></div>
</form>
</div>

{% if company_filter or am_filter or search_q %}
<div style="margin:8px 0;opacity:.78">
Showing <b>{{rows|length}}</b> matching enrollment case(s)
{% if search_q %} • Search: <b>{{search_q}}</b>{% endif %}
</div>
{% endif %}

{% if not banks %}
<div class="er-note"><b>No active bank account added yet.</b> Enrollment Report is still working. Add an account in Bank Manager to enable receiving-bank selection.</div>
{% endif %}

{% set total_received = rows|sum(attribute='total_received') %}
<div class="er-kpis">
<div class="er-kpi">Enrollment Cases<b>{{rows|length}}</b></div>
<div class="er-kpi">Total Received<b>₹{{"{:,.0f}".format(total_received or 0)}}</b></div>
<div class="er-kpi">Pending GM Completion<b>{{rows|selectattr('total_received','equalto',0)|list|length}}</b></div>
<div class="er-kpi">Active Banks<b>{{banks|length}}</b></div>
</div>

{% for r in rows %}
<div class="er-row">
<div class="er-summary">
<div><b style="color:#e6b73f;font-size:17px">{{r.client_name or 'Unnamed Client'}}</b><br><small>Lead ID: {{r.lead_code or r.lead_db_id or '-'}} • {{r.case_id}}{% if r.passport_no %} • Passport: {{r.passport_no}}{% endif %}</small></div>
<div><small>AM</small><br><b>{{r.linked_am_name or r.am_name or '-'}}</b></div>
<div><small>Package</small><br><b>₹{{"{:,.0f}".format(r.package_amount or 0)}}</b></div>
<div><small>Received</small><br><b class="received">₹{{"{:,.0f}".format(r.total_received or 0)}}</b></div>
<div><small>Balance</small><br><b class="pending">₹{{"{:,.0f}".format((r.package_amount or 0)-(r.total_received or 0))}}</b></div>
<div style="display:flex;gap:6px;flex-wrap:wrap">
<a class="toolbtn" href="{{url_for('cases')}}">Open Full Payment Form</a>
<form method="post" action="{{url_for('enrollment_soft_delete',case_id=r.id)}}" onsubmit="return confirm('Move this enrollment to Recycle Bin?')" style="margin:0">
<input type="hidden" name="reason" value="Deleted from Enrollment Report">
<button class="toolbtn danger" type="submit">🗑 Delete</button>
</form>
</div>
</div>
<div class="er-payment">
<div class="er-pay-card"><h3>1st Payment</h3>
<div>Total: ₹{{"{:,.0f}".format(r.booking_amount or 0)}}</div>
<div>Status: {{r.first_payment_status or 'Pending'}}</div>
<div>Bank: {{r.first_bank_name or '-'}}</div>
<div>Mode: {{r.first_payment_method or '-'}}</div>
<div>Date: {{r.first_payment_datetime or '-'}}</div>
{% if r.first_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.first_proof_id)}}">View Proof</a>{% endif %}
</div>
<div class="er-pay-card"><h3>2nd Payment</h3>
<div>Total: ₹{{"{:,.0f}".format(r.second_payment or 0)}}</div>
<div>Status: {{r.second_payment_status or 'Pending'}}</div>
<div>Bank: {{r.second_bank_name or '-'}}</div>
<div>Mode: {{r.second_payment_method or '-'}}</div>
<div>Date: {{r.second_payment_datetime or '-'}}</div>
{% if r.second_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.second_proof_id)}}">View Proof</a>{% endif %}
</div>
<div class="er-pay-card"><h3>Other Payment</h3>
<div>Total: ₹{{"{:,.0f}".format(r.other_payment_received or 0)}}</div>
<div>Bank: {{r.other_bank_name or r.other_payment_bank or '-'}}</div>
<div>Mode: {{r.other_payment_method or '-'}}</div>
<div>Date: {{r.other_payment_datetime or '-'}}</div>
{% if r.other_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.other_proof_id)}}">View Proof</a>{% endif %}
</div>
</div>
</div>
{% else %}
<div class="card"><h2 style="color:#e6b73f">No Enrollment Cases Yet</h2><p>When an AM changes a client to Enrolled, that client will appear here using the same Lead ID.</p></div>
{% endfor %}
{% endblock %}"""



# === v3.98 Global Enrollment Link Route Fix ===
for _name,_tpl in list(_GAUR_V337_TEMPLATES.items()):
    if isinstance(_tpl,str):
        _tpl=_tpl.replace('/accounts/client-ledger-legacy','/accounts/client-ledger')
        _tpl=_tpl.replace("url_for('cases') }}\">Enrolment Report","url_for('enrollment_report') }}\">Enrolment Report")
        _tpl=_tpl.replace("url_for('cases') }}\">Enrollment Report","url_for('enrollment_report') }}\">Enrollment Report")
        _tpl=_tpl.replace("url_for('cases') }}\">Client Accounts / Enrollment","url_for('enrollment_report') }}\">Client Accounts / Enrollment")
        _GAUR_V337_TEMPLATES[_name]=_tpl



_GAUR_V337_TEMPLATES['cases_safe.html']=r"""{% extends "base.html" %}{% block content %}
<style>
.cs-head{display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap}
.cs-row{border:1px solid #315f85;border-radius:14px;margin:12px 0;overflow:hidden;background:#0a243d}
.cs-summary{display:grid;grid-template-columns:1.5fr .8fr .8fr .8fr auto;gap:9px;align-items:center;padding:13px}
.cs-edit{display:none;padding:14px;border-top:1px solid #315f85;background:#071d32}.cs-edit.open{display:block}
.cs-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}.cs-section{border:1px solid #294f6d;border-radius:11px;padding:11px;margin-top:12px}.cs-section h3{color:#e6b73f;margin:0 0 9px}
.cs-balance{color:#ffbd67;font-weight:900}.cs-received{color:#55dca0;font-weight:900}.cs-note{background:#2b2414;border:1px solid #8b6923;border-radius:10px;padding:10px 12px;margin:10px 0}
@media(max-width:1000px){.cs-summary{grid-template-columns:1fr 1fr 1fr}.cs-grid{grid-template-columns:1fr 1fr}}
@media(max-width:600px){.cs-summary,.cs-grid{grid-template-columns:1fr}}
</style>

<div class="cs-head">
<div><h1 class="title">Enrollment • Payment Completion</h1>
<div style="display:inline-block;padding:4px 8px;border:1px solid #55dca0;border-radius:12px;color:#55dca0;font-size:12px">SAFE PAYMENT ROUTE • /cases</div>
<p style="opacity:.72">Same Lead ID • GM updates payment/bank/proof details here.</p></div>
<div><a class="toolbtn" href="{{url_for('enrollment_report')}}">← Enrollment Report</a> <a class="toolbtn" href="{{url_for('finance_banks')}}">🏦 Bank Manager</a>{% if u.role=='MD' %} <a class="toolbtn" href="{{url_for('enrollment_recycle_bin')}}">🗑 Recycle Bin</a>{% endif %}</div>
</div>

{% if not banks %}
<div class="cs-note"><b>No active bank account added yet.</b> This page will still open. Add a bank in Bank Manager to enable bank selection.</div>
{% endif %}

{% for r in rows %}
{% set balance=(r.package_amount or 0)-(r.total_received or 0) %}
<div class="cs-row">
  <div class="cs-summary">
    <div><b style="color:#e6b73f;font-size:18px">{{r.client_name or 'Unnamed Client'}}</b><br>
      <small>Lead ID: {{r.lead_code or r.lead_db_id or '-'}} • {{r.case_id}}{% if r.passport_no %} • Passport: {{r.passport_no}}{% endif %}</small>
    </div>
    <div><small>AM</small><br><b>{{r.linked_am_name or r.am_name or '-'}}</b></div>
    <div><small>Package</small><br><b>₹{{"{:,.0f}".format(r.package_amount or 0)}}</b></div>
    <div><small>Balance</small><br><b class="cs-balance">₹{{"{:,.0f}".format(balance if balance>0 else 0)}}</b></div>
    <div style="display:flex;gap:6px;flex-wrap:wrap">
{% set am_done = r.get('am_submitted',0) if r is mapping else (r.am_submitted or 0) %}
{% if u.role in ('MD','GM') %}
<button class="toolbtn" type="button" onclick="document.getElementById('edit{{r.id}}').classList.toggle('open')">✏ Update Payment</button>
<form method="post" action="{{url_for('enrollment_soft_delete',case_id=r.id)}}" onsubmit="return confirm('Delete this enrollment? It will move to Recycle Bin.')" style="margin:0">
<input type="hidden" name="reason" value="Deleted from Payment Completion">
<button class="toolbtn danger" type="submit">🗑 Delete</button>
</form>
{% elif u.role=='AM' and not am_done %}
<button class="toolbtn" type="button" onclick="document.getElementById('edit{{r.id}}').classList.toggle('open')" style="background:#e6b73f;color:#07192b;font-weight:700">✏ Submit Enrollment Details</button>
<span style="background:#2a1a00;color:#f0a030;padding:4px 10px;border-radius:10px;font-size:12px">⚠ First & only time — GM/MD will take over after this</span>
{% else %}
<span style="display:inline-block;padding:5px 12px;border:1px solid #315f85;border-radius:8px;color:#48d58b;font-size:13px">✅ Submitted — View Only</span>
{% endif %}
</div>
  </div>

  <div class="cs-edit" id="edit{{r.id}}">
    <form method="post" enctype="multipart/form-data" action="{{url_for('update_case',case_id=r.id)}}">
      <div class="cs-grid">
        <div><label>Package Amount</label><input type="number" step=".01" name="package_amount" value="{{r.package_amount or 0}}"></div>
        <div><label>After Visa Payment</label><input type="number" step=".01" name="after_visa_payment" value="{{r.after_visa_payment or 0}}"></div>
        <div><label>Counselor</label><input name="counselor_name" value="{{r.counselor_name or ''}}"></div>
        <div><label>AM</label><input name="am_name" value="{{r.am_name or r.linked_am_name or ''}}"></div>
        <div><label>Interaction Date</label><input type="date" name="interaction_date" value="{{r.interaction_date or ''}}"></div>
      </div>

      <div class="cs-section"><h3>1st Payment</h3><div class="cs-grid">
        <div><label>Cash</label><input type="number" step=".01" name="first_cash" value="{{r.first_cash or 0}}"></div>
        <div><label>RBL</label><input type="number" step=".01" name="first_rbl" value="{{r.first_rbl or 0}}"></div>
        <div><label>Yes Bank</label><input type="number" step=".01" name="first_yes_bank" value="{{r.first_yes_bank or 0}}"></div>
        <div><label>AU Bank</label><input type="number" step=".01" name="first_au_bank" value="{{r.first_au_bank or 0}}"></div>
        <div><label>Receiving Bank</label><select name="first_payment_bank_id"><option value="">Select Active Bank</option>{% for b in banks %}<option value="{{b.id}}" {% if r.first_payment_bank_id==b.id %}selected{% endif %}>{{b.bank_name}} • {{b.account_nickname or b.account_name}} • ****{{b.account_last4}}</option>{% endfor %}</select></div>
        <div><label>Payment Mode</label><select name="first_payment_method"><option selected>{{r.first_payment_method or 'Net Banking'}}</option><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option><option>Other</option></select></div>
        <div><label>Status</label><select name="first_payment_status"><option selected>{{r.first_payment_status or 'Pending'}}</option><option>Received</option><option>Partially Received</option><option>Pending</option></select></div>
        <div><label>Date & Time</label><input type="datetime-local" name="first_payment_datetime" value="{{(r.first_payment_datetime or '')[:16]}}"></div>
        <div><label>Screenshot / Proof</label><input type="file" name="first_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf">{% if r.first_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.first_proof_id)}}">View Existing</a>{% endif %}</div>
      </div></div>

      <div class="cs-section"><h3>2nd Payment</h3><div class="cs-grid">
        <div><label>Cash</label><input type="number" step=".01" name="second_cash" value="{{r.second_cash or 0}}"></div>
        <div><label>RBL</label><input type="number" step=".01" name="second_rbl" value="{{r.second_rbl or 0}}"></div>
        <div><label>Yes Bank</label><input type="number" step=".01" name="second_yes_bank" value="{{r.second_yes_bank or 0}}"></div>
        <div><label>AU Bank</label><input type="number" step=".01" name="second_au_bank" value="{{r.second_au_bank or 0}}"></div>
        <div><label>Receiving Bank</label><select name="second_payment_bank_id"><option value="">Select Active Bank</option>{% for b in banks %}<option value="{{b.id}}" {% if r.second_payment_bank_id==b.id %}selected{% endif %}>{{b.bank_name}} • {{b.account_nickname or b.account_name}} • ****{{b.account_last4}}</option>{% endfor %}</select></div>
        <div><label>Payment Mode</label><select name="second_payment_method"><option selected>{{r.second_payment_method or 'Net Banking'}}</option><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option><option>Other</option></select></div>
        <div><label>Status</label><select name="second_payment_status"><option selected>{{r.second_payment_status or 'Pending'}}</option><option>Received</option><option>Partially Received</option><option>Pending</option></select></div>
        <div><label>Date & Time</label><input type="datetime-local" name="second_payment_datetime" value="{{(r.second_payment_datetime or '')[:16]}}"></div>
        <div><label>Screenshot / Proof</label><input type="file" name="second_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf">{% if r.second_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.second_proof_id)}}">View Existing</a>{% endif %}</div>
      </div></div>

      <div class="cs-section"><h3>Other Payment / Final</h3><div class="cs-grid">
        <div><label>Other Payment Received</label><input type="number" step=".01" name="other_payment_received" value="{{r.other_payment_received or 0}}"></div>
        <div><label>Receiving Bank</label><select name="other_payment_bank_id"><option value="">Select Active Bank</option>{% for b in banks %}<option value="{{b.id}}" {% if r.other_payment_bank_id==b.id %}selected{% endif %}>{{b.bank_name}} • {{b.account_nickname or b.account_name}} • ****{{b.account_last4}}</option>{% endfor %}</select></div>
        <div><label>Payment Mode</label><select name="other_payment_method"><option selected>{{r.other_payment_method or 'Other'}}</option><option>QR / UPI</option><option>Net Banking</option><option>Cash Deposit</option><option>Cheque</option><option>Card / POS</option><option>Cash</option><option>Other</option></select></div>
        <div><label>Bank/Cash Note</label><input name="other_payment_bank" value="{{r.other_payment_bank or ''}}"></div>
        <div><label>Date & Time</label><input type="datetime-local" name="other_payment_datetime" value="{{(r.other_payment_datetime or '')[:16]}}"></div>
        <div><label>Screenshot / Proof</label><input type="file" name="other_payment_proof" accept=".jpg,.jpeg,.png,.webp,.pdf,image/*,application/pdf">{% if r.other_proof_id %}<a class="toolbtn" target="_blank" href="{{url_for('client_payment_proof',proof_id=r.other_proof_id)}}">View Existing</a>{% endif %}</div>
        <div><label>Remarks</label><textarea name="remarks">{{r.remarks or ''}}</textarea></div>
      </div></div>

      <button class="btn" style="margin-top:12px">Save Payment Details</button>
      {% if r.lead_db_id %}<a class="toolbtn" href="{{url_for('lead_profile',lead_db_id=r.lead_db_id)}}">Open Client</a>{% endif %}
    </form>
  </div>
</div>
{% else %}
<div class="card"><h2 style="color:#e6b73f">No enrollment cases found.</h2><p>AM-marked Enrolled clients will appear here through the same Lead ID.</p></div>
{% endfor %}
{% endblock %}"""



# === v3.98 Global Safe /cases Link Fix ===
for _n,_t in list(_GAUR_V337_TEMPLATES.items()):
    if isinstance(_t,str):
        _GAUR_V337_TEMPLATES[_n]=_t.replace('/cases-legacy','/cases')


# === v3.98 TODAY REPORT Dashboard Top ===
if 'dashboard.html' in _GAUR_V337_TEMPLATES:
    _d=_GAUR_V337_TEMPLATES['dashboard.html']
    _today='\n<div class="today-report-wrap luxury-today-report">\n  <div class="today-report-head">\n    <div><span class="today-live-dot"></span><b>TODAY\'S REPORT</b><small>{{today_date}}</small></div>\n    <div class="today-report-note">Auto-calculated day-wise from Enrollment Date & Payment Received Date/Time</div>\n  </div>\n  {% if u.role==\'MD\' %}\n  <div class="today-company-grid">\n    <div class="today-company-card today-scic-card">\n      <div class="today-brand-row">\n        <div class="today-logo today-logo-scic"><img src="" alt="ABC"></div>\n        <div class="today-brand-title"><i>♛</i><div class="today-company-name">ABC</div><small>★ ★ ★ ★ ★</small></div>\n      </div>\n      <div class="today-stat-grid">\n        <div><span>Today\'s Enrollments</span><b>{{today_report[\'ABC\'][\'enrollments\']}}</b></div>\n        <div><span>Today\'s Revenue</span><b>₹{{"{:,.0f}".format(today_report[\'ABC\'][\'revenue\'] or 0)}}</b></div>\n      </div>\n    </div>\n    <div class="today-vs"><span>VS</span></div>\n    <div class="today-company-card today-wwic-card">\n      <div class="today-brand-row today-brand-row-right">\n        <div class="today-brand-title"><i>♛</i><div class="today-company-name">XYZ</div><small>★ ★ ★ ★ ★</small></div>\n        <div class="today-logo today-logo-wwic"><img src="" alt="XYZ"></div>\n      </div>\n      <div class="today-stat-grid">\n        <div><span>Today\'s Enrollments</span><b>{{today_report[\'XYZ\'][\'enrollments\']}}</b></div>\n        <div><span>Today\'s Revenue</span><b>₹{{"{:,.0f}".format(today_report[\'XYZ\'][\'revenue\'] or 0)}}</b></div>\n      </div>\n    </div>\n  </div>\n  <div class="today-total-strip">\n    <span>Total Today</span>\n    <b>{{today_report[\'ABC\'][\'enrollments\'] + today_report[\'XYZ\'][\'enrollments\']}} Enrollments</b>\n    <b>₹{{"{:,.0f}".format((today_report[\'ABC\'][\'revenue\'] or 0)+(today_report[\'XYZ\'][\'revenue\'] or 0))}} Revenue</b>\n  </div>\n  {% elif u.role==\'GM\' %}\n  {% set c=u.company_code %}\n  <div class="today-company-grid single">\n    <div class="today-company-card {% if c==\'ABC\' %}today-scic-card{% else %}today-wwic-card{% endif %}">\n      <div class="today-brand-row">\n        {% if c==\'ABC\' %}<div class="today-logo today-logo-scic"><img src="" alt="ABC"></div>{% endif %}\n        <div class="today-brand-title"><i>♛</i><div class="today-company-name">{{\'ABC\' if c==\'ABC\' else \'XYZ\'}}</div><small>★ ★ ★ ★ ★</small></div>\n        {% if c!=\'ABC\' %}<div class="today-logo today-logo-wwic"><img src="" alt="XYZ"></div>{% endif %}\n      </div>\n      <div class="today-stat-grid">\n        <div><span>Today\'s Enrollments</span><b>{{today_report[c][\'enrollments\']}}</b></div>\n        <div><span>Today\'s Revenue</span><b>₹{{"{:,.0f}".format(today_report[c][\'revenue\'] or 0)}}</b></div>\n      </div>\n    </div>\n  </div>\n  {% endif %}\n</div>\n'
    _marker='<h1 class="title">'
    if _marker in _d and 'today-report-wrap' not in _d:
        _d=_d.replace(_marker,_today+_marker,1)
    _GAUR_V337_TEMPLATES['dashboard.html']=_d
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</head>','\n<style>\n.today-report-wrap.luxury-today-report{border:2px solid #d7aa35;border-radius:24px;background:radial-gradient(circle at 50% 58%,rgba(224,171,42,.12),transparent 22%),linear-gradient(115deg,#020b14,#082039 35%,#020910 50%,#082039 65%,#020b14);padding:14px 18px 18px;margin:8px 0 18px;box-shadow:0 12px 30px rgba(0,0,0,.38),inset 0 0 30px rgba(35,88,125,.16);overflow:hidden}\n.today-report-head{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;border-bottom:1px solid rgba(230,183,63,.35);padding:0 0 10px;margin-bottom:12px}.today-report-head>div:first-child{display:flex;align-items:center;gap:8px;color:#f0bd43;font-size:16px}.today-report-head small{color:#dce8f2;opacity:.72}.today-report-note{font-size:12px;opacity:.68}.today-live-dot{width:9px;height:9px;border-radius:50%;background:#55dca0;box-shadow:0 0 0 4px rgba(85,220,160,.12),0 0 9px rgba(85,220,160,.7);display:inline-block}\n.today-company-grid{display:grid;grid-template-columns:1fr 90px 1fr;gap:14px;align-items:stretch}.today-company-grid.single{grid-template-columns:1fr;max-width:760px;margin:auto}.today-company-card{position:relative;background:linear-gradient(145deg,#08213a,#061626);border:1px solid #9a7626;border-radius:18px;padding:13px 14px 14px;box-shadow:inset 0 0 20px rgba(34,93,133,.12),0 6px 18px rgba(0,0,0,.18)}\n.today-brand-row{display:grid;grid-template-columns:105px 1fr;gap:12px;align-items:center;margin-bottom:10px}.today-brand-row-right{grid-template-columns:1fr 105px}.today-logo{display:flex;align-items:center;justify-content:center;width:96px;height:96px;margin:auto}.today-logo img{max-width:100%;max-height:100%;object-fit:contain;display:block}.today-logo-scic{background:transparent;filter:drop-shadow(0 0 7px rgba(241,192,60,.25))}.today-logo-wwic{background:#fff;border:3px solid #d9aa35;border-radius:50%;padding:6px;box-shadow:0 0 8px #e0b43d,0 0 18px rgba(232,186,66,.22)}.today-logo-wwic img{background:#fff;border-radius:50%}\n.today-brand-title{text-align:center}.today-brand-title i{display:block;color:#f4c13b;font-style:normal;font-size:24px;line-height:1}.today-company-name{text-align:center;color:#f0bd43;font-weight:900;letter-spacing:.5px;font-size:20px;line-height:1.05}.today-brand-title small{display:block;color:#f5c544;letter-spacing:3px;font-size:9px;margin-top:4px}\n.today-stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px}.today-stat-grid>div{background:linear-gradient(180deg,#06182a,#04121f);border:1px solid #315a78;border-radius:12px;padding:10px;text-align:center;box-shadow:inset 0 0 12px rgba(22,74,109,.14)}.today-stat-grid span{display:block;font-size:12px;opacity:.74}.today-stat-grid b{display:block;color:#fff4d2;font-size:25px;margin-top:4px;text-shadow:0 0 7px rgba(255,215,93,.35)}\n.today-vs{display:flex;align-items:center;justify-content:center;width:72px;height:72px;border-radius:50%;align-self:center;justify-self:center;background:radial-gradient(circle,#10202e 35%,#02080d 72%);border:4px solid #d8a932;box-shadow:0 0 0 4px #07111a,0 0 18px #d4a42f;color:#f4c64d;font-size:24px;font-weight:900;text-shadow:0 0 7px rgba(244,198,77,.45)}\n.today-total-strip{display:flex;justify-content:center;gap:20px;flex-wrap:wrap;margin-top:12px;padding-top:10px;border-top:1px solid rgba(230,183,63,.25)}.today-total-strip span{opacity:.72}.today-total-strip b{color:#f0bd43;text-shadow:0 0 5px rgba(240,189,67,.18)}\n@media(max-width:900px){.today-company-grid{grid-template-columns:1fr 58px 1fr;gap:8px}.today-brand-row,.today-brand-row-right{display:flex;flex-direction:column;gap:4px}.today-logo{width:62px;height:62px}.today-logo-wwic{padding:3px;border-width:2px}.today-brand-title i,.today-brand-title small{display:none}.today-company-name{font-size:13px}.today-stat-grid b{font-size:20px}.today-vs{width:48px;height:48px;font-size:16px;border-width:2px}}\n@media(max-width:700px){.today-company-grid{grid-template-columns:1fr}.today-vs{margin:auto}.today-stat-grid{grid-template-columns:1fr 1fr}.today-brand-row-right{flex-direction:column-reverse}}\n@media(max-width:450px){.today-stat-grid{grid-template-columns:1fr}.today-report-note{display:none}}\n</style>\n</head>',1)



# v3.98 MD-only merged live performance + today's report
try:
    _d381=_GAUR_V337_TEMPLATES.get("dashboard.html","")
    _w381=_d381.find('<div class="today-report-wrap luxury-today-report">')
    if _w381>=0:
        _m381=_d381.find("{% if u.role=='MD' %}",_w381)
        _g381=_d381.find("{% elif u.role=='GM' %}",_m381)
        if _m381>=0 and _g381>_m381:
            _md381="""{% if u.role=='MD' %}
<style>
/* v3.98 safe MD-only single-panel display.
   No template cutting: simply suppress duplicate shells visually. */
.competition-livebar.lux-score,
.competition-livebar.lux-score.show{display:none!important}
.today-report-wrap.luxury-today-report{
  border:0!important;
  background:transparent!important;
  box-shadow:none!important;
  padding:0!important;
  margin:8px 0 18px!important;
  overflow:visible!important
}
.today-report-wrap.luxury-today-report>.today-report-head{display:none!important}
</style>
<div id="mdMergedPerformance" class="md-merged-performance">
  <div class="md-merge-head">
    <div><span class="today-live-dot"></span><b>LIVE PERFORMANCE + TODAY'S REPORT</b><small>{{today_date}}</small></div>
    <div class="md-merge-note">Overall live score + today's enrollments & received payments</div>
  </div>
  <div class="md-merge-grid">
    <section class="md-team md-team-scic">
      <div class="md-team-logo md-team-logo-scic"><img class="brand-logo sc-logo" src=""></div>
      <div class="md-team-content">
        <div class="md-crown">♛</div><h3>ABC</h3><div class="md-stars">★ ★ ★ ★ ★</div>
        <div id="mdScicLiveScore" class="md-live-score">0 • ₹0</div>
        <div class="md-live-label">OVERALL LIVE PERFORMANCE</div>
        <div class="md-meter"><span id="mdScicMeter" style="width:0%"></span></div>
        <div class="md-today-stats">
          <div><span>Today's Enrollments</span><b>{{today_report['SCIC']['enrollments']}}</b></div>
          <div><span>Today's Revenue</span><b>₹{{"{:,.0f}".format(today_report['SCIC']['revenue'] or 0)}}</b></div>
        </div>
      </div>
    </section>
    <div class="md-merge-vs">VS</div>
    <section class="md-team md-team-wwic">
      <div class="md-team-content">
        <div class="md-crown">♛</div><h3>XYZ</h3><div class="md-stars">★ ★ ★ ★ ★</div>
        <div id="mdWwicLiveScore" class="md-live-score">0 • ₹0</div>
        <div class="md-live-label">OVERALL LIVE PERFORMANCE</div>
        <div class="md-meter md-meter-blue"><span id="mdWwicMeter" style="width:0%"></span></div>
        <div class="md-today-stats">
          <div><span>Today's Enrollments</span><b>{{today_report['WWIC']['enrollments']}}</b></div>
          <div><span>Today's Revenue</span><b>₹{{"{:,.0f}".format(today_report['WWIC']['revenue'] or 0)}}</b></div>
        </div>
      </div>
      <div class="md-team-logo md-team-logo-wwic"><img class="brand-logo ww-logo" src=""></div>
    </section>
  </div>
  <div class="md-merge-footer">
    <span>● TODAY</span>
    <b>{{today_report['SCIC']['enrollments']+today_report['WWIC']['enrollments']}} Enrollments</b>
    <b>₹{{"{:,.0f}".format((today_report['SCIC']['revenue'] or 0)+(today_report['WWIC']['revenue'] or 0))}} Revenue</b>
  </div>
</div>
<script>
(function(){
 function sync(){
  const pairs=[['scicLiveScore','mdScicLiveScore'],['wwicLiveScore','mdWwicLiveScore']];
  pairs.forEach(p=>{const a=document.getElementById(p[0]),b=document.getElementById(p[1]);if(a&&b)b.textContent=a.textContent;});
  [['scicMeter','mdScicMeter'],['wwicMeter','mdWwicMeter']].forEach(p=>{const a=document.getElementById(p[0]),b=document.getElementById(p[1]);if(a&&b)b.style.width=a.style.width||'0%';});
 }
 sync(); setInterval(sync,750);
})();
</script>
"""
            _d381=_d381[:_m381]+_md381+_d381[_g381:]
            _GAUR_V337_TEMPLATES["dashboard.html"]=_d381
except Exception as _e381:
    print("v3.98 merged dashboard patch skipped",_e381,flush=True)

_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</head>','\\n<style>\\n\n.md-standalone-score-hidden{display:none!important}\n.md-merged-performance{border:2px solid #d7aa35;border-radius:26px;padding:14px 18px 18px;background:radial-gradient(circle at 50% 45%,rgba(226,173,42,.14),transparent 18%),linear-gradient(115deg,#020b14,#082039 35%,#020910 50%,#082039 65%,#020b14);box-shadow:0 12px 32px rgba(0,0,0,.42),inset 0 0 32px rgba(35,88,125,.16)}\n.md-merge-head{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;padding-bottom:10px;border-bottom:1px solid rgba(224,177,57,.35)}.md-merge-head>div:first-child{display:flex;align-items:center;gap:8px;color:#f1c048}.md-merge-head small{color:#c8d6e4}.md-merge-note{font-size:12px;opacity:.7}\n.md-merge-grid{display:grid;grid-template-columns:1fr 100px 1fr;align-items:center;gap:14px;padding:13px 0 8px}.md-team{display:grid;align-items:center;gap:16px}.md-team-scic{grid-template-columns:145px 1fr}.md-team-wwic{grid-template-columns:1fr 145px}\n.md-team-logo{width:136px;height:136px;display:flex;align-items:center;justify-content:center;margin:auto}.md-team-logo img{max-width:100%;max-height:100%;object-fit:contain}.md-team-logo-scic{background:transparent}.md-team-logo-wwic{background:#fff;border:4px solid #d9aa35;border-radius:50%;padding:8px;box-shadow:0 0 8px #e8ba42,0 0 22px rgba(232,186,66,.3)}\n.md-team-content{text-align:center}.md-crown{font-size:25px;color:#f4c13b}.md-team h3{margin:1px 0!important;color:#f2c34a!important;font-size:22px!important}.md-stars{font-size:10px;color:#f4c543;letter-spacing:4px}.md-live-score{font-size:39px;font-weight:900;color:#fff2cf;text-shadow:0 0 9px rgba(255,215,93,.55)}.md-live-label{font-size:9px;color:#eabb3c;letter-spacing:1.6px;font-weight:800;margin:1px 0 5px}\n.md-meter{height:10px;background:#020b14;border:1px solid #a77e22;border-radius:18px;overflow:hidden}.md-meter span{display:block;height:100%;background:linear-gradient(90deg,#bf8610,#ffe679,#dba221)}.md-meter-blue span{background:linear-gradient(90deg,#1165d8,#32b9ff,#84ddff)}\n.md-today-stats{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:9px}.md-today-stats>div{background:#061829;border:1px solid #315a78;border-radius:11px;padding:8px}.md-today-stats span{display:block;font-size:10px;opacity:.78}.md-today-stats b{display:block;font-size:20px;color:#fff4d2}\n.md-merge-vs{width:76px;height:76px;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:auto;background:#07111a;border:4px solid #d8a932;box-shadow:0 0 18px #d4a42f;color:#f4c64d;font-size:25px;font-weight:900}\n.md-merge-footer{display:flex;justify-content:center;gap:18px;flex-wrap:wrap;padding-top:9px;border-top:1px solid rgba(230,183,63,.26)}.md-merge-footer b{color:#f1bd3c}\n@media(max-width:900px){.md-merge-grid{grid-template-columns:1fr 60px 1fr}.md-team-scic{grid-template-columns:90px 1fr}.md-team-wwic{grid-template-columns:1fr 90px}.md-team-logo{width:82px;height:82px}.md-live-score{font-size:28px}.md-team h3{font-size:16px!important}.md-merge-vs{width:50px;height:50px;font-size:17px}}\n@media(max-width:700px){.md-merge-grid{grid-template-columns:1fr}.md-team-scic,.md-team-wwic{display:flex;flex-direction:column;gap:5px}.md-team-wwic .md-team-content{order:2}.md-team-wwic .md-team-logo{order:1}.md-merge-note{display:none}}\n\\n</style>\\n</head>',1)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</head>','\n<style>\n\n/* v3.98 Smart unread update toast */\n.update-toast{position:fixed;right:18px;bottom:92px;z-index:1002;width:min(390px,calc(100vw - 28px));display:grid;grid-template-columns:42px 1fr;gap:10px;padding:13px 38px 13px 13px;border:1px solid #c79a2d;border-radius:16px;background:linear-gradient(135deg,#24141d 0%,#0b2943 100%);box-shadow:0 14px 34px rgba(0,0,0,.42),0 0 18px rgba(225,178,52,.14);animation:updateToastIn .22s ease-out}\n.update-toast.urgent{border-color:#e0ad2e;box-shadow:0 14px 34px rgba(0,0,0,.44),0 0 22px rgba(225,173,46,.22)}\n.update-toast.closing{opacity:0;transform:translateY(8px);transition:.18s ease}\n.update-toast-icon{width:38px;height:38px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#0c3150;border:1px solid #315f80;font-size:20px}\n.update-toast-kicker{font-size:9px;font-weight:900;letter-spacing:1.8px;color:#edbd43;margin-bottom:2px}\n.update-toast-title{font-size:14px;font-weight:900;color:#ffd46a;line-height:1.15;padding-right:4px}\n.update-toast-text{font-size:12px;color:#eef4fa;opacity:.9;line-height:1.25;margin-top:3px;max-height:46px;overflow:hidden}\n.update-toast-open{display:inline-block;margin-top:8px;padding:6px 10px;border-radius:9px;border:1px solid #39739a;background:#08243c;color:#fff!important;font-size:12px;font-weight:800;text-decoration:none!important}\n.update-toast-close{position:absolute;right:8px;top:7px;width:25px;height:25px;border:0!important;border-radius:50%!important;background:transparent!important;color:#d9e5ee!important;font-size:20px;line-height:20px;padding:0!important;cursor:pointer}\n.update-toast-close:hover{background:rgba(255,255,255,.08)!important}\n@keyframes updateToastIn{from{opacity:0;transform:translateY(12px) scale(.98)}to{opacity:1;transform:none}}\n@media(max-width:600px){.update-toast{right:10px;bottom:82px;width:calc(100vw - 20px)}}\n\n</style>\n</head>',1)

_GAUR_V337_TEMPLATES['management_reporting.html']='{% extends "base.html" %}{% block content %}\n<h1 class="title">📊 Management Reporting</h1>\n<p style="opacity:.76">AM-wise Daily Report + Month Till Date • MD / GM Management View</p>\n\n<div class="card no-print">\n<form method="get" class="grid3">\n  <div><label>Report Date</label><input type="date" name="date" value="{{report_date}}"></div>\n  {% if u.role==\'MD\' %}<div><label>Company</label><select name="company">\n    <option value="">Both Companies</option>\n    <option value="SCIC" {% if company==\'ABC\' %}selected{% endif %}>ABC</option>\n    <option value="WWIC" {% if company==\'XYZ\' %}selected{% endif %}>XYZ</option>\n  </select></div>{% endif %}\n  <div style="display:flex;align-items:end;gap:8px;flex-wrap:wrap">\n    <button class="btn">View Report</button>\n    <a class="toolbtn" href="{{url_for(\'management_reporting_export\',date=report_date,company=company)}}">⬇ Download Excel</a>\n    <button class="toolbtn" type="button" onclick="shareManagementReport()">🟢 Share on WhatsApp</button>\n  </div>\n</form>\n</div>\n\n<div class="metrics">\n  <div class="metric">Today\'s Allocated<b>{{totals.today_allocated}}</b><small>{{report_date}}</small></div>\n  <div class="metric">Today\'s Enrolled<b>{{totals.today_enrolled}}</b><small>AM-wise</small></div>\n  <div class="metric">Today\'s Revenue<b>₹{{"{:,.0f}".format(totals.today_revenue)}}</b><small>Actual payment date</small></div>\n  <div class="metric">MTD Enrolled<b>{{totals.mtd_enrolled}}</b><small>{{month_start}} to {{month_end}}</small></div>\n  <div class="metric">MTD Revenue<b>₹{{"{:,.0f}".format(totals.mtd_revenue)}}</b><small>Till selected date</small></div>\n</div>\n\n<div class="card">\n<h2 style="color:#e6b73f;margin-top:0">AM-wise Daily + Month Till Date Report</h2>\n<div class="tablewrap"><table id="managementReportTable">\n<thead><tr>\n<th>Company</th><th>AM Name</th>\n<th>Today Allocated</th><th>Today Interested</th><th>Today Enrolled</th><th>Today Revenue</th>\n<th>MTD Allocated</th><th>MTD Interested</th><th>MTD Enrolled</th><th>MTD Revenue</th><th>MTD Conversion</th>\n</tr></thead>\n<tbody>\n{% for x in rows %}\n<tr>\n<td>{{x.company_code}}</td><td><b style="color:#f1c34b">{{x.am_name}}</b></td>\n<td>{{x.today_allocated}}</td><td>{{x.today_interested}}</td><td>{{x.today_enrolled}}</td><td>₹{{"{:,.0f}".format(x.today_revenue)}}</td>\n<td>{{x.mtd_allocated}}</td><td>{{x.mtd_interested}}</td><td>{{x.mtd_enrolled}}</td><td>₹{{"{:,.0f}".format(x.mtd_revenue)}}</td><td>{{"%.1f"|format(x.conversion)}}%</td>\n</tr>\n{% else %}<tr><td colspan="11">No AM records found for this scope.</td></tr>{% endfor %}\n</tbody>\n<tfoot><tr>\n<th colspan="2">TOTAL</th><th>{{totals.today_allocated}}</th><th>{{totals.today_interested}}</th><th>{{totals.today_enrolled}}</th><th>₹{{"{:,.0f}".format(totals.today_revenue)}}</th>\n<th>{{totals.mtd_allocated}}</th><th>{{totals.mtd_interested}}</th><th>{{totals.mtd_enrolled}}</th><th>₹{{"{:,.0f}".format(totals.mtd_revenue)}}</th><th>-</th>\n</tr></tfoot>\n</table></div>\n</div>\n\n<script>\nconst managementWhatsAppText={{wa_text|tojson}};\nfunction shareManagementReport(){\n  window.open(\'https://wa.me/?text=\'+encodeURIComponent(managementWhatsAppText),\'_blank\');\n}\n</script>\n{% endblock %}'
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</head>','\n<style>\n\n/* v3.98 global report controls */\n.global-tools{display:flex!important;gap:8px!important;flex-wrap:wrap!important;align-items:center!important}\n.global-tools .toolbtn{white-space:nowrap}\n@media(max-width:700px){.global-tools{gap:6px!important}.global-tools .toolbtn{font-size:12px!important;padding:8px 9px!important}}\n\n</style>\n</head>',1)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</body>','\n<script>\n(function(){\n  function pageTitle(){\n    var t=document.querySelector(\'h1.title\')||document.querySelector(\'main h1\')||document.querySelector(\'main h2\');\n    return (t?t.textContent:document.title).replace(/\\s+/g,\' \').trim();\n  }\n\n  var wa=document.getElementById(\'safeWhatsAppBtn\');\n  if(wa){\n    wa.addEventListener(\'click\',function(){\n      var msg=\'THE IMMIGRATION MANAGEMENT SYSTEM\\nReport/Page: \'+pageTitle()+\'\\nLink: \'+window.location.href;\n      window.open(\'https://wa.me/?text=\'+encodeURIComponent(msg),\'_blank\',\'noopener\');\n    });\n  }\n\n  var pr=document.getElementById(\'safePrintBtn\');\n  if(pr){\n    pr.addEventListener(\'click\',function(){ window.print(); });\n  }\n\n  var ex=document.getElementById(\'safeExcelBtn\');\n  if(ex){\n    ex.addEventListener(\'click\',function(){\n      try{\n        var main=document.querySelector(\'main\');\n        if(!main){ alert(\'No report content found on this page.\'); return; }\n\n        var tables=[].slice.call(main.querySelectorAll(\'table\')).filter(function(t){\n          return t.offsetParent!==null;\n        });\n\n        if(!tables.length){\n          alert(\'No tabular data is available on this page for Excel export.\');\n          return;\n        }\n\n        var name=pageTitle();\n        var html=\'<html><head><meta charset="UTF-8"><style>table{border-collapse:collapse}td,th{border:1px solid #777;padding:5px}</style></head><body>\';\n        html+=\'<h2>\'+name.replace(/[<>&]/g,\'\')+\'</h2>\';\n        tables.forEach(function(t,i){\n          html+=\'<h3>Section \'+(i+1)+\'</h3>\'+t.outerHTML;\n        });\n        html+=\'</body></html>\';\n\n        var blob=new Blob([\'\\ufeff\',html],{type:\'application/vnd.ms-excel\'});\n        var a=document.createElement(\'a\');\n        a.href=URL.createObjectURL(blob);\n        a.download=name.replace(/[^a-z0-9]+/gi,\'_\').replace(/^_|_$/g,\'\')+\'_\'+new Date().toISOString().slice(0,10)+\'.xls\';\n        document.body.appendChild(a);\n        a.click();\n        setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},700);\n      }catch(err){\n        alert(\'Excel export could not start. Please refresh once and try again.\');\n      }\n    });\n  }\n})();\n</script>\n\n</body>',1)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</head>',"\n<style>\n\n/* =========================================================\n   v3.98 GM Dashboard — SAFE CSS-ONLY MERGED DISPLAY\n   Uses the two already-working blocks and visually combines them.\n   ========================================================= */\n.competition-livebar.gm-css-merged-score{\n  margin-bottom:0!important;\n  border-radius:26px 26px 0 0!important;\n  border-bottom-color:rgba(215,170,53,.55)!important;\n  padding-bottom:22px!important;\n  box-shadow:0 12px 26px rgba(0,0,0,.28),inset 0 0 28px rgba(35,88,125,.14)!important;\n}\n\n/* The GM Today's Report following the live scoreboard becomes the lower half\n   of the SAME luxury module. The ~ selector is deliberate so it still works\n   even if another harmless element exists between the two blocks. */\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report{\n  margin-top:0!important;\n  margin-bottom:20px!important;\n  border-top:0!important;\n  border-radius:0 0 26px 26px!important;\n  background:\n    radial-gradient(circle at 50% 16%,rgba(224,171,42,.09),transparent 24%),\n    linear-gradient(115deg,#020b14,#082039 35%,#020910 50%,#082039 65%,#020b14)!important;\n  box-shadow:0 14px 28px rgba(0,0,0,.30),inset 0 0 28px rgba(35,88,125,.12)!important;\n  padding:18px 22px 20px!important;\n  overflow:visible!important;\n}\n\n/* Compact integrated TODAY header */\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-report-head{\n  max-width:980px!important;\n  margin:0 auto 12px!important;\n  padding:2px 4px 10px!important;\n  border-bottom:1px solid rgba(230,183,63,.25)!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-report-head>div:first-child{\n  margin:auto!important;\n  font-size:15px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-report-note{\n  display:none!important;\n}\n\n/* Make the GM company card modular, centered and wider */\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-company-grid.single{\n  width:100%!important;\n  max-width:980px!important;\n  margin:0 auto!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-company-card{\n  display:grid!important;\n  grid-template-columns:minmax(220px,34%) minmax(0,1fr)!important;\n  gap:18px!important;\n  align-items:center!important;\n  min-height:170px!important;\n  padding:16px 18px!important;\n  border-radius:18px!important;\n  border:1px solid #8f7028!important;\n  background:linear-gradient(145deg,#08213a,#061626)!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-brand-row{\n  display:flex!important;\n  flex-direction:column!important;\n  justify-content:center!important;\n  gap:6px!important;\n  margin:0!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-logo{\n  width:90px!important;\n  height:90px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-company-name{\n  font-size:19px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid{\n  grid-template-columns:1fr 1fr!important;\n  gap:12px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid>div{\n  min-height:92px!important;\n  display:flex!important;\n  flex-direction:column!important;\n  justify-content:center!important;\n  border-radius:14px!important;\n  padding:12px!important;\n}\n.competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid b{\n  font-size:28px!important;\n}\n\n/* LIVE PERFORMANCE pill sits exactly on the join and makes the two halves read as one module */\n.competition-livebar.gm-css-merged-score .lux-live-badge{\n  bottom:-18px!important;\n  z-index:4!important;\n}\n\n@media(max-width:850px){\n  .competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-company-card{\n    grid-template-columns:1fr!important;\n  }\n  .competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid{\n    grid-template-columns:1fr 1fr!important;\n  }\n}\n@media(max-width:560px){\n  .competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report{\n    padding:16px 9px 14px!important;\n  }\n  .competition-livebar.gm-css-merged-score ~ .today-report-wrap.luxury-today-report .today-stat-grid{\n    grid-template-columns:1fr!important;\n  }\n}\n\n</style>\n</head>",1)
_GAUR_V337_TEMPLATES['base.html']=_GAUR_V337_TEMPLATES['base.html'].replace('</head>','<style>\n.bin397{background:linear-gradient(180deg,#b92e3d,#791a27)!important;border:1px solid #e45a69!important;color:#fff!important;font-weight:900!important}\n.recycle397{display:inline-flex!important;align-items:center!important;justify-content:center!important;background:#0b3328!important;border:1px solid #42c985!important;color:#a9f2cb!important;text-decoration:none!important;font-weight:900!important}\n.md-recycle397{border:2px solid #42c985!important;box-shadow:0 0 22px rgba(66,201,133,.10)!important;margin-top:18px!important}.md-recycle397 h2{color:#72e6a9!important}\n</style></head>',1)
_GAUR_V337_TEMPLATES['enrollment_recycle_bin.html']=r"""{% extends "base.html" %}{% block content %}
<style>
.erb-head{display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap}
.erb-actions{display:flex;gap:8px;flex-wrap:wrap}.danger-btn{background:#7d2632!important;border-color:#b94353!important;color:#fff!important}
</style>
<div class="erb-head">
<div><h1 class="title">🗑 Enrollment Recycle Bin</h1>
<p style="opacity:.75">MD-only control • Restore or permanently delete archived enrollment/payment records.</p></div>
<div><a class="toolbtn" href="{{url_for('enrollment_report')}}">← Enrollment Report</a></div>
</div>

<div class="card" style="border-color:#b27a2d">
<b style="color:#e6b73f">Safety Rule:</b> MD and GM can move an enrollment to Recycle Bin, but <b>only MD</b> can restore or permanently delete it.
</div>

<div class="tablewrap"><table>
<thead><tr><th>Client / Case</th><th>Company</th><th>Enrollment Date</th><th>Total Received</th><th>Deleted By</th><th>Deleted At</th><th>Reason</th><th>MD Action</th></tr></thead>
<tbody>
{% for r in rows %}
<tr>
<td><b>{{r.client_name}}</b><br><small>{{r.case_id}}</small></td>
<td>{{r.company_code}}</td>
<td>{{r.enrollment_date or '-'}}</td>
<td>₹{{"{:,.0f}".format(r.total_received or 0)}}</td>
<td>{{r.deleted_by or '-'}}</td>
<td>{{r.deleted_at or '-'}}</td>
<td>{{r.deletion_reason or '-'}}</td>
<td><div class="erb-actions">
<form method="post" action="{{url_for('enrollment_restore',case_id=r.id)}}"><button class="btn">↩ Restore</button></form>
<form method="post" action="{{url_for('enrollment_permanent_delete',case_id=r.id)}}" onsubmit="return confirm('FINAL DELETE: This enrollment and its payment proofs will be permanently erased. This cannot be undone. Continue?')"><button class="toolbtn danger-btn">🔥 Final Delete</button></form>
</div></td>
</tr>
{% else %}
<tr><td colspan="8">Enrollment Recycle Bin is empty.</td></tr>
{% endfor %}
</tbody></table></div>
{% endblock %}"""


@app.context_processor
def inject_chat_context():
    return {"chat_unread_count":0,"notification_unread_count":0,"priority_notification":None}


@app.context_processor
def inject_user():
    return dict(current_user=current_user())


# v3.21: when served by Gunicorn (`app:app`), __main__ does not run.
# Attempt schema initialization at import time, but never crash the HTTP process.
if __name__!="__main__":
    # Railway: bind HTTP immediately; never wait for PostgreSQL DDL during Gunicorn import.
    pass

if __name__=="__main__":
    import traceback
    port=int(os.environ.get("PORT","5050"))
    try:
        init_db()
        print("GAUR CRM database initialization: OK",flush=True)
    except Exception:
        # Critical Railway resilience:
        # log the database/migration error but still bind the HTTP server,
        # so /healthz remains reachable and deployment diagnostics can be opened.
        print("\nGAUR CRM DATABASE INITIALIZATION WARNING:\n",flush=True)
        traceback.print_exc()
    print(f"GAUR CRM v3.38 ready | DB={'PostgreSQL' if IS_POSTGRES else 'SQLite'} | port={port}",flush=True)
    from waitress import serve
    serve(app,host="0.0.0.0",port=port,threads=8)
