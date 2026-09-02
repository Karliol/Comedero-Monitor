
import os
from datetime import datetime, timedelta
from io import BytesIO
from zoneinfo import ZoneInfo

from flask import Flask, request, jsonify, send_file, Response
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, select
from sqlalchemy.orm import declarative_base, sessionmaker

APP_TZ = ZoneInfo("America/Lima")
API_KEY = os.environ.get("ESP32_API_KEY", "CAMBIA_ESTA_CLAVE")

database_url = os.environ.get("DATABASE_URL", "sqlite:///comedero_cloud.db")

# Neon entrega normalmente postgresql://...
if database_url.startswith("postgresql://"):
    database_url = database_url.replace("postgresql://", "postgresql+psycopg://", 1)
elif database_url.startswith("postgres://"):
    database_url = database_url.replace("postgres://", "postgresql+psycopg://", 1)

connect_args = {}
if database_url.startswith("sqlite"):
    connect_args = {"check_same_thread": False}

engine = create_engine(
    database_url,
    pool_pre_ping=True,
    connect_args=connect_args
)

SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

class Medicion(Base):
    __tablename__ = "mediciones"

    id = Column(Integer, primary_key=True, autoincrement=True)
    mascota_id = Column(String(60), nullable=False, index=True)
    nombre = Column(String(100), nullable=False)
    fecha_hora = Column(DateTime, nullable=False, index=True)
    peso = Column(Float, nullable=False)

Base.metadata.create_all(engine)

app = Flask(__name__)

def now_lima():
    return datetime.now(APP_TZ).replace(tzinfo=None)

def start_for_scale(scale):
    now = now_lima()
    options = {
        "1h": now - timedelta(hours=1),
        "6h": now - timedelta(hours=6),
        "12h": now - timedelta(hours=12),
        "24h": now - timedelta(hours=24),
        "3d": now - timedelta(days=3),
        "7d": now - timedelta(days=7),
    }
    return options.get(scale, now - timedelta(hours=24))

def fmt_dt(dt):
    return dt.strftime("%Y-%m-%d %H:%M:%S")

def valid_api_key():
    return request.headers.get("X-API-Key", "") == API_KEY

HTML = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Monitoreo de alimentación</title>

<style>
:root{
  --bg:#f3f5f7;
  --card:#ffffff;
  --text:#1f2937;
  --muted:#6b7280;
  --border:#e5e7eb;
  --accent:#2563eb;
}
*{box-sizing:border-box}
html,body{
  margin:0;
  padding:0;
  width:100%;
  max-width:100%;
  overflow-x:hidden;
}
body{
  font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;
  background:var(--bg);
  color:var(--text);
  -webkit-text-size-adjust:100%;
  text-size-adjust:100%;
}
header{
  width:100%;
  background:#fff;
  border-bottom:1px solid var(--border);
  padding:16px;
  display:flex;
  justify-content:space-between;
  gap:14px;
  align-items:center;
  flex-wrap:wrap;
}
h1{font-size:22px;line-height:1.2;margin:0}
main{width:100%;max-width:1200px;margin:0 auto;padding:16px}
.toolbar{
  display:flex;
  gap:8px;
  align-items:center;
  flex-wrap:wrap;
  min-width:0;
}
select,button{font:inherit;font-size:16px;max-width:100%}
select{
  min-height:44px;
  padding:10px 12px;
  border:1px solid var(--border);
  border-radius:10px;
  background:#fff;
}
.cards{
  width:100%;
  display:grid;
  grid-template-columns:repeat(4,minmax(0,1fr));
  gap:12px;
  margin:16px 0;
}
.card,.panel{
  min-width:0;
  width:100%;
  background:#fff;
  border:1px solid var(--border);
  border-radius:16px;
  padding:16px;
}
.label{font-size:13px;color:var(--muted);margin-bottom:8px}
.value{font-size:30px;line-height:1.15;font-weight:750;word-break:break-word}
.small{font-size:14px}
.dot{
  display:inline-block;
  width:10px;height:10px;
  border-radius:50%;
  background:#9ca3af;
  margin-right:6px;
}
.panel-head{
  display:flex;
  justify-content:space-between;
  gap:12px;
  align-items:center;
  flex-wrap:wrap;
  margin-bottom:12px;
}
.panel-head h2{font-size:18px;line-height:1.3;margin:0}
.ranges{display:flex;gap:6px;flex-wrap:wrap}
.range,.action{
  border:1px solid var(--border);
  background:#fff;
  border-radius:9px;
  padding:9px 12px;
  cursor:pointer;
  min-height:44px;
  touch-action:manipulation;
}
.range.active{background:#1f2937;color:#fff;border-color:#1f2937}
.action.primary{background:var(--accent);color:#fff;border-color:var(--accent)}
.chart-wrap{
  position:relative;
  width:100%;
  min-width:0;
  height:420px;
  border-top:1px solid var(--border);
  padding-top:12px;
  overflow:hidden;
}
canvas{
  display:block;
  width:100%!important;
  max-width:100%;
  height:100%!important;
  touch-action:pan-y;
}
.tip{
  position:absolute;
  display:none;
  pointer-events:none;
  background:#111827;
  color:#fff;
  padding:8px 10px;
  border-radius:8px;
  font-size:12px;
  white-space:nowrap;
  transform:translate(-50%,-115%);
}
.actions{
  display:flex;
  justify-content:flex-end;
  gap:8px;
  flex-wrap:wrap;
  margin-top:14px;
}
.empty{color:var(--muted);font-size:14px}
.cloud{font-size:12px;color:var(--muted);margin-top:4px}

@media(max-width:850px){
  .cards{grid-template-columns:repeat(2,minmax(0,1fr))}
  .chart-wrap{height:360px}
}

@media(max-width:600px){
  header{padding:12px;align-items:stretch}
  header>div{min-width:0}
  h1{font-size:20px}
  .toolbar{width:100%;display:grid;grid-template-columns:1fr}
  .toolbar label{font-size:13px;color:var(--muted)}
  .toolbar select,.toolbar button{width:100%}
  main{padding:10px}
  .cards{grid-template-columns:1fr;gap:10px;margin:10px 0}
  .card,.panel{padding:14px;border-radius:13px}
  .value{font-size:28px}
  .panel-head{align-items:stretch}
  .ranges{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));width:100%}
  .range{width:100%;padding:8px 5px}
  .chart-wrap{height:310px}
  .actions{display:grid;grid-template-columns:1fr}
  .actions .action{width:100%}
}

@media(max-width:360px){
  h1{font-size:18px}
  .ranges{grid-template-columns:repeat(2,minmax(0,1fr))}
  .value{font-size:25px}
  .chart-wrap{height:280px}
}
</style>

</head>
<body>
<header>
<div>
<h1>🐱 Monitoreo de alimentación</h1>
<div class="cloud">Plataforma en Internet · hora de Perú</div>
</div>
<div class="toolbar">
<label for="pet">Mascota</label>
<select id="pet"><option value="">Esperando datos...</option></select>
<button id="refresh" class="action">Actualizar</button>
</div>
</header>

<main>
<div class="cards">
<div class="card"><div class="label">Peso actual</div><div id="weight" class="value">0000 g</div><div class="small" style="margin-top:8px"><span id="dot" class="dot"></span><span id="status">Sin datos</span></div></div>
<div class="card"><div class="label">Consumo estimado hoy</div><div id="today" class="value">0.0 g</div></div>
<div class="card"><div class="label">Último consumo detectado</div><div id="lastc" class="value">0.0 g</div></div>
<div class="card"><div class="label">Última actualización</div><div id="lastt" class="value small">--</div></div>
</div>

<section class="panel">
<div class="panel-head">
<h2>Peso de comida en el plato vs tiempo</h2>
<div class="ranges">
<button class="range" data-range="1h">1 h</button>
<button class="range" data-range="6h">6 h</button>
<button class="range" data-range="12h">12 h</button>
<button class="range active" data-range="24h">24 h</button>
<button class="range" data-range="3d">3 días</button>
<button class="range" data-range="7d">7 días</button>
</div>
</div>

<div class="chart-wrap">
<canvas id="chart"></canvas>
<div id="tip" class="tip"></div>
</div>
<div id="empty" class="empty" hidden>No hay mediciones en esta escala de tiempo.</div>

<div class="actions">
<button id="excelRange" class="action">📥 Excel de la escala visible</button>
<button id="excelAll" class="action primary">📥 Excel completo de la mascota</button>
</div>
</section>
</main>

<script>
(()=>{
const pet=document.getElementById('pet');
const weight=document.getElementById('weight');
const today=document.getElementById('today');
const lastc=document.getElementById('lastc');
const lastt=document.getElementById('lastt');
const status=document.getElementById('status');
const dot=document.getElementById('dot');
const canvas=document.getElementById('chart');
const ctx=canvas.getContext('2d');
const tip=document.getElementById('tip');
const empty=document.getElementById('empty');

let selected='',scale='24h',points=[];

async function getj(url){
  const r=await fetch(url,{cache:'no-store'});
  if(!r.ok) throw new Error('HTTP '+r.status);
  return r.json();
}
function parseLocal(s){return s?new Date(s.replace(' ','T')):null;}
function fmt(v){return Math.round(Number(v)||0).toString().padStart(4,'0')+' g';}

async function loadPets(){
  const arr=await getj('/api/mascotas');
  const old=pet.value;
  pet.innerHTML='';
  if(!arr.length){
    pet.innerHTML='<option value="">Esperando datos...</option>';
    selected='';draw([]);return;
  }
  arr.forEach(x=>{
    const o=document.createElement('option');
    o.value=x.id;o.textContent=x.nombre;pet.appendChild(o);
  });
  pet.value=arr.some(x=>x.id===old)?old:arr[0].id;
  selected=pet.value;
  await updateAll();
}

async function updateCurrent(){
  if(!selected)return;
  const d=await getj('/api/ultimo/'+encodeURIComponent(selected));
  weight.textContent=fmt(d.peso);
  lastt.textContent=d.fecha_hora||'--';
  const t=parseLocal(d.fecha_hora);
  if(!t){status.textContent='Sin datos';dot.style.background='#9ca3af';return;}
  const sec=(Date.now()-t.getTime())/1000;
  if(sec<90){status.textContent='En línea';dot.style.background='#16a34a';}
  else if(sec<300){status.textContent='Datos recientes';dot.style.background='#d97706';}
  else{status.textContent='Sin conexión';dot.style.background='#dc2626';}
}

async function updateStats(){
  if(!selected)return;
  const d=await getj('/api/estadisticas/'+encodeURIComponent(selected));
  today.textContent=Number(d.consumo_hoy||0).toFixed(1)+' g';
  lastc.textContent=Number(d.ultimo_consumo||0).toFixed(1)+' g';
}

async function updateChart(){
  if(!selected){draw([]);return;}
  const d=await getj('/api/historico/'+encodeURIComponent(selected)+'?escala='+scale);
  points=d.fechas.map((f,i)=>({t:parseLocal(f),raw:f,y:Number(d.pesos[i])}));
  draw(points);
}
async function updateAll(){await Promise.all([updateCurrent(),updateStats(),updateChart()]);}

function resize(){
  const dpr=Math.min(window.devicePixelRatio||1,2);
  const r=canvas.getBoundingClientRect();
  const cssW=Math.max(280,Math.floor(r.width||canvas.parentElement.clientWidth||320));
  const cssH=Math.max(240,Math.floor(r.height||310));
  canvas.width=Math.round(cssW*dpr);
  canvas.height=Math.round(cssH*dpr);
  ctx.setTransform(dpr,0,0,dpr,0,0);
}

function draw(data){
  resize();
  const W=canvas.clientWidth,H=canvas.clientHeight;
  ctx.clearRect(0,0,W,H);
  const p={l:58,r:18,t:18,b:42},pw=W-p.l-p.r,ph=H-p.t-p.b;
  ctx.font='12px system-ui';ctx.lineWidth=1;
  for(let g=0;g<=1000;g+=100){
    const y=p.t+ph-(g/1000)*ph;
    ctx.strokeStyle='#e5e7eb';ctx.beginPath();ctx.moveTo(p.l,y);ctx.lineTo(W-p.r,y);ctx.stroke();
    ctx.fillStyle='#6b7280';ctx.textAlign='right';ctx.textBaseline='middle';ctx.fillText(g+' g',p.l-8,y);
  }
  ctx.fillStyle='#6b7280';ctx.textAlign='center';ctx.textBaseline='top';ctx.fillText('Tiempo',p.l+pw/2,H-20);
  if(!data.length){empty.hidden=false;return;}
  empty.hidden=true;
  const times=data.map(x=>x.t.getTime());
  let t0=Math.min(...times),t1=Math.max(...times);
  if(t0===t1){t0-=30000;t1+=30000;}
  const X=t=>p.l+((t.getTime()-t0)/(t1-t0))*pw;
  const Y=v=>p.t+ph-(Math.max(0,Math.min(1000,v))/1000)*ph;
  for(let i=0;i<=6;i++){
    const d=new Date(t0+(t1-t0)*(i/6)),x=p.l+pw*(i/6);
    const lab=(scale==='3d'||scale==='7d')
      ?d.toLocaleString([],{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'})
      :d.toLocaleTimeString([],{hour:'2-digit',minute:'2-digit'});
    ctx.fillStyle='#6b7280';ctx.textAlign='center';ctx.fillText(lab,x,H-p.b+9);
  }
  ctx.strokeStyle='#2563eb';ctx.lineWidth=2;ctx.beginPath();
  data.forEach((q,i)=>{const x=X(q.t),y=Y(q.y);if(i===0)ctx.moveTo(x,y);else ctx.lineTo(x,y);});
  ctx.stroke();
  if(data.length<=250){
    ctx.fillStyle='#2563eb';
    data.forEach(q=>{ctx.beginPath();ctx.arc(X(q.t),Y(q.y),2.3,0,Math.PI*2);ctx.fill();});
  }
  canvas._geom={X,Y};
}

canvas.addEventListener('mousemove',ev=>{
  if(!points.length||!canvas._geom){tip.style.display='none';return;}
  const r=canvas.getBoundingClientRect(),mx=ev.clientX-r.left,G=canvas._geom;
  let best=null,bd=Infinity;
  points.forEach(q=>{const x=G.X(q.t),d=Math.abs(x-mx);if(d<bd){bd=d;best={q,x,y:G.Y(q.y)};}});
  if(!best||bd>30){tip.style.display='none';return;}
  tip.style.display='block';tip.style.left=best.x+'px';tip.style.top=best.y+'px';
  tip.textContent=best.q.raw+' · '+best.q.y.toFixed(1)+' g';
});
canvas.addEventListener('mouseleave',()=>tip.style.display='none');

pet.addEventListener('change',()=>{selected=pet.value;updateAll();});
document.querySelectorAll('.range').forEach(b=>b.addEventListener('click',()=>{
  document.querySelectorAll('.range').forEach(x=>x.classList.remove('active'));
  b.classList.add('active');scale=b.dataset.range;updateChart();
}));
document.getElementById('refresh').addEventListener('click',updateAll);
document.getElementById('excelRange').addEventListener('click',()=>{if(selected)location.href='/api/excel/'+encodeURIComponent(selected)+'?escala='+scale;});
document.getElementById('excelAll').addEventListener('click',()=>{if(selected)location.href='/api/excel/'+encodeURIComponent(selected)+'?escala=todo;'});
window.addEventListener('resize',()=>draw(points));
window.addEventListener('orientationchange',()=>setTimeout(()=>draw(points),250));
if(window.visualViewport){
  window.visualViewport.addEventListener('resize',()=>draw(points));
}
setInterval(()=>{if(selected){updateCurrent();updateStats();updateChart();}else loadPets();},10000);
loadPets().catch(console.error);
})();
</script>
</body>
</html>"""

@app.get("/")
def home():
    return Response(HTML, mimetype="text/html")

@app.get("/health")
def health():
    return jsonify(status="ok")

@app.post("/api/peso")
def receive_weight():
    if not valid_api_key():
        return jsonify(error="API key inválida"), 401

    d = request.get_json(silent=True) or {}
    pet_id = str(d.get("mascota_id", "")).strip()
    name = str(d.get("nombre", "")).strip()

    try:
        weight = float(d.get("peso"))
    except Exception:
        return jsonify(error="Peso inválido"), 400

    if not pet_id or not name:
        return jsonify(error="Faltan mascota_id o nombre"), 400

    weight = max(0.0, min(1000.0, weight))
    timestamp = now_lima()

    with SessionLocal() as db:
        db.add(Medicion(
            mascota_id=pet_id,
            nombre=name,
            fecha_hora=timestamp,
            peso=weight
        ))
        db.commit()

    return jsonify(
        estado="ok",
        peso=weight,
        fecha_hora=fmt_dt(timestamp)
    )

@app.get("/api/mascotas")
def pets():
    with SessionLocal() as db:
        rows = db.execute(
            select(Medicion).order_by(Medicion.id.desc())
        ).scalars().all()

    seen = set()
    result = []
    for r in rows:
        if r.mascota_id not in seen:
            seen.add(r.mascota_id)
            result.append({"id": r.mascota_id, "nombre": r.nombre})

    result.sort(key=lambda x: x["nombre"].lower())
    return jsonify(result)

@app.get("/api/ultimo/<pet_id>")
def latest(pet_id):
    with SessionLocal() as db:
        r = db.execute(
            select(Medicion)
            .where(Medicion.mascota_id == pet_id)
            .order_by(Medicion.id.desc())
            .limit(1)
        ).scalar_one_or_none()

    if not r:
        return jsonify(nombre="", fecha_hora="", peso=0)

    return jsonify(
        nombre=r.nombre,
        fecha_hora=fmt_dt(r.fecha_hora),
        peso=r.peso
    )

@app.get("/api/historico/<pet_id>")
def history(pet_id):
    scale = request.args.get("escala", "24h")
    start = start_for_scale(scale)

    with SessionLocal() as db:
        rows = db.execute(
            select(Medicion)
            .where(
                Medicion.mascota_id == pet_id,
                Medicion.fecha_hora >= start
            )
            .order_by(Medicion.fecha_hora.asc())
        ).scalars().all()

    if len(rows) > 1200:
        step = max(1, len(rows)//1200)
        sampled = rows[::step]
        if sampled[-1].id != rows[-1].id:
            sampled.append(rows[-1])
        rows = sampled

    return jsonify(
        fechas=[fmt_dt(r.fecha_hora) for r in rows],
        pesos=[r.peso for r in rows]
    )

def calc_consumption(rows, threshold=2.0):
    if len(rows) < 2:
        return 0.0, 0.0

    total = 0.0
    last = 0.0
    prev = float(rows[0].peso)

    for r in rows[1:]:
        cur = float(r.peso)
        drop = prev - cur
        if drop >= threshold:
            total += drop
            last = drop
        prev = cur

    return round(total, 1), round(last, 1)

@app.get("/api/estadisticas/<pet_id>")
def stats(pet_id):
    start = now_lima().replace(hour=0, minute=0, second=0, microsecond=0)

    with SessionLocal() as db:
        rows = db.execute(
            select(Medicion)
            .where(
                Medicion.mascota_id == pet_id,
                Medicion.fecha_hora >= start
            )
            .order_by(Medicion.fecha_hora.asc())
        ).scalars().all()

    total, last = calc_consumption(rows)
    return jsonify(consumo_hoy=total, ultimo_consumo=last)

@app.get("/api/excel/<pet_id>")
def excel(pet_id):
    scale = request.args.get("escala", "todo")

    with SessionLocal() as db:
        stmt = select(Medicion).where(Medicion.mascota_id == pet_id)
        if scale != "todo":
            stmt = stmt.where(Medicion.fecha_hora >= start_for_scale(scale))
        rows = db.execute(
            stmt.order_by(Medicion.fecha_hora.asc())
        ).scalars().all()

    if not rows:
        return Response("No hay datos para exportar.", status=404, mimetype="text/plain")

    name = rows[0].nombre
    wb = Workbook()
    ws = wb.active
    ws.title = "Registro"

    headers = [
        "Mascota", "Fecha", "Hora", "Peso alimento (g)",
        "Cambio de peso (g)", "Consumo estimado (g)"
    ]
    ws.append(headers)

    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9EAF7")
        c.alignment = Alignment(horizontal="center")

    prev = None
    for r in rows:
        cur = float(r.peso)
        change = 0.0 if prev is None else cur - prev
        consumption = abs(change) if change <= -2.0 else 0.0

        ws.append([
            name,
            r.fecha_hora.strftime("%d/%m/%Y"),
            r.fecha_hora.strftime("%H:%M:%S"),
            round(cur, 1),
            round(change, 1),
            round(consumption, 1)
        ])

        prev = cur

    for i, width in enumerate([18, 14, 12, 20, 20, 24], 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    f = BytesIO()
    wb.save(f)
    f.seek(0)

    stamp = now_lima().strftime("%Y%m%d_%H%M")
    suffix = "completo" if scale == "todo" else scale
    filename = f"{name}_{suffix}_{stamp}.xlsx"

    return send_file(
        f,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)

