"""Genera index.html: catálogo navegable con filtros y compra por WhatsApp.

Uso:
    python3 generar_web.py [ruta_al_excel]

El index.html resultante se sube a la MISMA carpeta del repositorio donde están
los PNG, así las imágenes se referencian por ruta relativa y no hay nada que configurar.
"""
import sys, json, html
import pandas as pd

# ─── Configura esto ──────────────────────────────────────────────────────────
TIENDA = "Sun Sale Club"                  # nombre que aparece arriba
WHATSAPP = "573217514936"          # tu número: 57 + celular, sin + ni espacios
SUBTITULO = "Monturas originales · Cali"
# ─────────────────────────────────────────────────────────────────────────────

EXCEL = sys.argv[1] if len(sys.argv) > 1 else 'Catalogo_gafas.xlsx'
SALIDA = 'index.html'

TINTES = {
    'Negro': '#1b1d21', 'Dorado': '#c8963e', 'Plateado': '#b9bec6',
    'Carey / Havana': '#7a4a24', 'Bicolor': 'linear', 'Gris': '#6d7278',
    'Azul': '#2f5d8a', 'Marrón': '#5b3a22', 'Otro': '#7d7a74',
}

SIGLA = {'Aviador': 'aviador', 'Wayfarer': 'wayfarer', 'Clubmaster': 'clubmaster',
         'Redonda': 'redonda', 'Hexagonal': 'hexagonal', 'Gatuna': 'gatuna',
         'Cuadrada': 'cuadrada'}


def numero(v):
    t = str(v).strip()
    if t.endswith('.0'):
        t = t[:-2]
    t = t.replace(',', '').replace('$', '').replace('.', '').replace(' ', '')
    return int(t) if t.isdigit() else None


def cargar(ruta):
    df = pd.read_excel(ruta, sheet_name='Catálogo', dtype=str).fillna('')
    df = df[(df['SKU'] != '') & (df['SKU'] != 'EJEMPLO')]
    productos = []
    for _, r in df.iterrows():
        forma = str(r['Forma']).strip() or 'Por definir'
        color = str(r['Familia color']).strip() or 'Por definir'
        productos.append({
            'sku': r['SKU'],
            'marca': r['Marca'],
            'modelo': str(r['Modelo']),
            'nombre': str(r['Nombre']),
            'tono': str(r['Color']),
            'forma': forma,
            'color': color,
            'ref': str(r['Referencia']),
            'precio': numero(r['Precio (COP)']),
            'stock': numero(r['Cantidad']) or 0,
            'link': str(r['Link limpio']),
            'probador': str(r['Probador virtual']).strip().lower().startswith('s'),
            'img': f"{r['SKU']}.png",
        })
    return productos


PLANTILLA = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TIENDA__ · Catálogo</title>
<meta name="description" content="Catálogo de gafas de sol originales. Filtra por forma, color y marca.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,600;12..96,800&family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@400;500;600&display=swap" rel="stylesheet">
<style>
:root{
  --noche:#131519; --humo:#1b1e24; --borde:#2b2f37;
  --niebla:#8b919b; --hueso:#f7f5f1; --ambar:#c8963e;
  --display:'Bricolage Grotesque',system-ui,sans-serif;
  --texto:'IBM Plex Sans',system-ui,sans-serif;
  --spec:'IBM Plex Mono',ui-monospace,monospace;
}
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{
  background:var(--noche); color:var(--hueso); font-family:var(--texto);
  -webkit-font-smoothing:antialiased; line-height:1.5;
}
a{color:inherit}
button{font:inherit;cursor:pointer}
:focus-visible{outline:2px solid var(--ambar); outline-offset:3px; border-radius:2px}

/* ── Encabezado ───────────────────────────────────── */
.tope{
  position:sticky; top:0; z-index:40; background:rgba(19,21,25,.94);
  backdrop-filter:blur(12px); border-bottom:1px solid var(--borde);
}
.tope-in{max-width:1180px;margin:0 auto;padding:14px 18px 0}
.marca{
  font-family:var(--display); font-weight:800; font-size:clamp(22px,5.5vw,30px);
  letter-spacing:-.03em; margin:0; display:flex; align-items:baseline; gap:10px;
}
.marca span{font-family:var(--spec);font-size:11px;font-weight:400;color:var(--niebla);letter-spacing:.08em}
.sub{margin:2px 0 12px;font-size:13px;color:var(--niebla)}

/* ── Controles ────────────────────────────────────── */
.buscar{position:relative;margin-bottom:12px}
.buscar input{
  width:100%; padding:11px 14px 11px 38px; background:var(--humo);
  border:1px solid var(--borde); border-radius:10px; color:var(--hueso);
  font-family:var(--texto); font-size:15px;
}
.buscar input::placeholder{color:#666c76}
.buscar svg{position:absolute;left:13px;top:50%;transform:translateY(-50%);color:var(--niebla)}

.riel{display:flex;gap:8px;overflow-x:auto;padding-bottom:12px;scrollbar-width:none;-webkit-overflow-scrolling:touch}
.riel::-webkit-scrollbar{display:none}

.forma{
  flex:0 0 auto; background:transparent; border:1px solid var(--borde);
  border-radius:12px; padding:8px 12px 6px; color:var(--niebla);
  display:flex; flex-direction:column; align-items:center; gap:3px;
  min-width:70px; transition:border-color .15s,color .15s,background .15s;
}
.forma svg{display:block}
.forma small{font-family:var(--spec);font-size:9.5px;letter-spacing:.06em;text-transform:uppercase}
.forma[aria-pressed="true"]{background:var(--hueso);color:var(--noche);border-color:var(--hueso)}
.forma:hover:not([aria-pressed="true"]){color:var(--hueso);border-color:#3d434d}

.tinte{
  flex:0 0 auto; width:34px; height:34px; border-radius:50%; padding:0;
  border:1px solid var(--borde); position:relative; transition:transform .12s;
}
.tinte[aria-pressed="true"]{box-shadow:0 0 0 2px var(--noche),0 0 0 4px var(--ambar)}
.tinte:hover{transform:scale(1.08)}
.tinte b{position:absolute;inset:0;border-radius:50%;display:block}

.fila-marcas{display:flex;gap:8px;overflow-x:auto;padding-bottom:14px;scrollbar-width:none}
.fila-marcas::-webkit-scrollbar{display:none}
.chip{
  flex:0 0 auto; background:transparent; border:1px solid var(--borde);
  border-radius:999px; padding:6px 13px; color:var(--niebla); font-size:13px;
  transition:.15s;
}
.chip[aria-pressed="true"]{background:var(--hueso);color:var(--noche);border-color:var(--hueso)}
.chip:hover:not([aria-pressed="true"]){color:var(--hueso)}

.estado{
  display:flex;align-items:center;justify-content:space-between;gap:12px;
  padding:0 0 12px; font-family:var(--spec); font-size:11.5px; color:var(--niebla);
  letter-spacing:.05em;
}
.limpiar{background:none;border:none;color:var(--ambar);font-family:var(--spec);font-size:11.5px;padding:0;text-decoration:underline;text-underline-offset:3px}

/* ── Rejilla ──────────────────────────────────────── */
main{max-width:1180px;margin:0 auto;padding:20px 18px 90px}
.rejilla{display:grid;grid-template-columns:repeat(2,1fr);gap:12px}
@media(min-width:640px){.rejilla{grid-template-columns:repeat(3,1fr);gap:16px}}
@media(min-width:940px){.rejilla{grid-template-columns:repeat(4,1fr);gap:18px}}

.tarjeta{
  background:var(--humo); border:1px solid var(--borde); border-radius:14px;
  overflow:hidden; text-align:left; padding:0; color:inherit; width:100%;
  display:flex; flex-direction:column; transition:border-color .18s,transform .18s;
}
.tarjeta:hover{border-color:#454c57;transform:translateY(-2px)}
.tarjeta img{width:100%;aspect-ratio:1;object-fit:cover;display:block;background:var(--hueso)}
.cuerpo{padding:11px 12px 13px;display:flex;flex-direction:column;gap:3px;flex:1}
.eyebrow{font-family:var(--spec);font-size:9.5px;letter-spacing:.12em;text-transform:uppercase;color:var(--niebla)}
.mod{font-family:var(--display);font-weight:600;font-size:16px;letter-spacing:-.02em;margin:0}
.nom{font-size:12.5px;color:var(--niebla);margin:0;min-height:1.2em;
  overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical}
.precio{font-family:var(--spec);font-size:14px;font-weight:500;color:var(--ambar);margin-top:6px}
.precio.consultar{color:var(--niebla);font-size:12px}
.agotado{position:absolute;inset:0;background:rgba(19,21,25,.62);display:grid;place-items:center;font-family:var(--spec);font-size:11px;letter-spacing:.1em}
.marco-img{position:relative}

.vacio{text-align:center;padding:70px 20px;color:var(--niebla)}
.vacio p{margin:0 0 6px}

/* ── Panel de detalle ─────────────────────────────── */
.velo{
  position:fixed;inset:0;background:rgba(8,9,11,.72);z-index:60;
  display:flex;align-items:flex-end;justify-content:center;
  opacity:0;pointer-events:none;transition:opacity .2s;
}
.velo.abierto{opacity:1;pointer-events:auto}
@media(min-width:700px){.velo{align-items:center}}
.panel{
  background:var(--humo);border:1px solid var(--borde);width:100%;max-width:440px;
  border-radius:18px 18px 0 0;padding:18px;max-height:92vh;overflow-y:auto;
  transform:translateY(14px);transition:transform .22s;
}
.velo.abierto .panel{transform:none}
@media(min-width:700px){.panel{border-radius:18px}}
.panel img{width:100%;aspect-ratio:1;object-fit:cover;border-radius:12px;background:var(--hueso)}
.panel h2{font-family:var(--display);font-weight:800;font-size:24px;letter-spacing:-.03em;margin:14px 0 2px}
.datos{width:100%;border-collapse:collapse;margin:14px 0 16px;font-size:13.5px}
.datos th{text-align:left;font-weight:400;color:var(--niebla);padding:7px 0;border-bottom:1px solid var(--borde);width:42%}
.datos td{text-align:right;padding:7px 0;border-bottom:1px solid var(--borde);font-family:var(--spec);font-size:12.5px}
.acciones{display:flex;flex-direction:column;gap:9px}
.wa{
  display:flex;align-items:center;justify-content:center;gap:9px;
  background:var(--ambar);color:#15171b;border:none;border-radius:11px;
  padding:14px;font-weight:600;font-size:15px;text-decoration:none;
}
.wa:hover{filter:brightness(1.07)}
.ver{
  display:flex;align-items:center;justify-content:center;gap:8px;
  background:transparent;color:var(--hueso);border:1px solid var(--borde);
  border-radius:11px;padding:12px;font-size:14px;text-decoration:none;
}
.ver:hover{border-color:var(--hueso)}
.cerrar{position:absolute;top:26px;right:26px;background:rgba(19,21,25,.82);border:none;color:var(--hueso);
  font-size:22px;line-height:1;width:34px;height:34px;border-radius:50%;display:grid;place-items:center;padding:0}
.cerrar:hover{background:rgba(19,21,25,.95)}
.panel{position:relative}
.pie{text-align:center;font-family:var(--spec);font-size:10.5px;color:#5b616b;padding:26px 18px 34px;letter-spacing:.05em}

@media(prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
</style>
</head>
<body>

<header class="tope">
  <div class="tope-in">
    <h1 class="marca">__TIENDA__ <span id="conteo"></span></h1>
    <p class="sub">__SUBTITULO__</p>

    <div class="buscar">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>
      <input id="q" type="search" placeholder="Buscar modelo, marca o color" autocomplete="off">
    </div>

    <div class="riel" id="formas" role="group" aria-label="Filtrar por forma"></div>
    <div class="riel" id="tintes" role="group" aria-label="Filtrar por color"></div>
    <div class="fila-marcas" id="marcas" role="group" aria-label="Filtrar por marca"></div>

    <div class="estado">
      <span id="resumen"></span>
      <button class="limpiar" id="limpiar" hidden>Quitar filtros</button>
    </div>
  </div>
</header>

<main>
  <div class="rejilla" id="rejilla"></div>
  <div class="vacio" id="vacio" hidden>
    <p>No hay monturas con esos filtros.</p>
    <button class="limpiar" onclick="limpiarTodo()">Ver todas</button>
  </div>
</main>

<p class="pie">__TIENDA__ · Escríbenos por WhatsApp para disponibilidad y envíos</p>

<div class="velo" id="velo" role="dialog" aria-modal="true" aria-labelledby="pt">
  <div class="panel" id="panel"></div>
</div>

<script>
const PRODUCTOS = __DATOS__;
const WHATSAPP = "__WHATSAPP__";
const TIENDA = "__TIENDA__";

const FORMAS = __FORMAS__;
const TINTES = __TINTES__;
const MARCAS = __MARCAS__;

/* siluetas: las mismas formas de las fichas */
const SVG = {
  aviador:'M2 4h20l-2 7-8 8-8-8z',
  wayfarer:'M2 4h20l-2 10H4z',
  clubmaster:'M2 4h20l-2 10H4z',
  redonda:'',
  hexagonal:'M6 4h12l4 6-4 6H6l-4-6z',
  gatuna:'M2 6l4-2h12l4 2-3 8H5z',
  cuadrada:'M3 4h18v10H3z'
};

function iconoForma(f){
  const k = FORMAS[f] || 'cuadrada';
  const t = 'stroke="currentColor" fill="none" stroke-width="1.6" stroke-linejoin="round"';
  if(k === 'redonda')
    return `<svg width="30" height="16" viewBox="0 0 30 16"><circle cx="8" cy="8" r="6" ${t}/><circle cx="22" cy="8" r="6" ${t}/><path d="M14 7h2M2 6L0 5M28 6l2-1" ${t}/></svg>`;
  const d = {
    aviador:'M1 2h11l-1.2 7L6.5 13 2.2 9z',
    wayfarer:'M1 2h11l-1.4 9H2.4z',
    clubmaster:'M1 2h11l-1.4 9H2.4z',
    hexagonal:'M3.5 2h7l2.5 4.5L10.5 11h-7L1 6.5z',
    gatuna:'M1 3.5L4 2h8l.6 7L7 12 1.8 9.6z',
    cuadrada:'M1 2h11v9H1z'
  }[k];
  const ceja = k === 'clubmaster'
    ? `<path d="M0.6 2.2h11.8M16.6 2.2h11.8M12.4 2.2h4.2" ${t} stroke-width="3.4" stroke-linecap="square"/>` : '';
  return `<svg width="30" height="16" viewBox="0 0 30 16"><path d="${d}" ${t}/>`
       + `<g transform="translate(29,0) scale(-1,1)"><path d="${d}" ${t}/></g>`
       + ceja + `<path d="M12.4 4.2h5.2" ${t}/></svg>`;
}

let filtros = { forma:null, color:null, marca:null, q:'' };

const $ = s => document.querySelector(s);
const pesos = n => '$' + n.toLocaleString('es-CO');

/* ── construir filtros ── */
Object.keys(FORMAS).forEach(f => {
  const b = document.createElement('button');
  b.className = 'forma'; b.setAttribute('aria-pressed','false');
  b.innerHTML = iconoForma(f) + `<small>${f}</small>`;
  b.onclick = () => { filtros.forma = filtros.forma === f ? null : f; pintar(); };
  b.dataset.forma = f;
  $('#formas').appendChild(b);
});

TINTES.forEach(([nombre, css]) => {
  const b = document.createElement('button');
  b.className = 'tinte'; b.setAttribute('aria-pressed','false');
  b.title = nombre; b.setAttribute('aria-label','Color ' + nombre);
  const fondo = css === 'linear'
    ? 'linear-gradient(135deg,#1b1d21 0 50%,#c8963e 50% 100%)' : css;
  b.innerHTML = `<b style="background:${fondo}"></b>`;
  b.onclick = () => { filtros.color = filtros.color === nombre ? null : nombre; pintar(); };
  b.dataset.color = nombre;
  $('#tintes').appendChild(b);
});

MARCAS.forEach(m => {
  const b = document.createElement('button');
  b.className = 'chip'; b.setAttribute('aria-pressed','false'); b.textContent = m;
  b.onclick = () => { filtros.marca = filtros.marca === m ? null : m; pintar(); };
  b.dataset.marca = m;
  $('#marcas').appendChild(b);
});

$('#q').addEventListener('input', e => { filtros.q = e.target.value.toLowerCase().trim(); pintar(); });
$('#limpiar').onclick = limpiarTodo;

function limpiarTodo(){
  filtros = { forma:null, color:null, marca:null, q:'' };
  $('#q').value = '';
  pintar();
}

function coincide(p){
  if(filtros.forma && p.forma !== filtros.forma) return false;
  if(filtros.color && p.color !== filtros.color) return false;
  if(filtros.marca && p.marca !== filtros.marca) return false;
  if(filtros.q){
    const heno = [p.marca,p.modelo,p.nombre,p.tono,p.forma,p.color,p.ref].join(' ').toLowerCase();
    if(!heno.includes(filtros.q)) return false;
  }
  return true;
}

function pintar(){
  const lista = PRODUCTOS.filter(coincide);
  const r = $('#rejilla');
  r.innerHTML = '';

  lista.forEach(p => {
    const b = document.createElement('button');
    b.className = 'tarjeta';
    b.onclick = () => abrir(p);
    b.innerHTML = `
      <div class="marco-img">
        <img src="${p.img}" alt="${p.marca} ${p.modelo}" loading="lazy">
        ${p.stock === 0 ? '<span class="agotado">AGOTADA</span>' : ''}
      </div>
      <div class="cuerpo">
        <span class="eyebrow">${p.marca}</span>
        <p class="mod">${p.modelo || p.ref}</p>
        <p class="nom">${[p.nombre, p.tono].filter(Boolean).join(' · ')}</p>
        <span class="precio ${p.precio ? '' : 'consultar'}">${p.precio ? pesos(p.precio) : 'Consultar precio'}</span>
      </div>`;
    r.appendChild(b);
  });

  $('#vacio').hidden = lista.length > 0;
  $('#conteo').textContent = PRODUCTOS.length + ' monturas';
  $('#resumen').textContent = lista.length === PRODUCTOS.length
    ? 'Mostrando todas'
    : `${lista.length} de ${PRODUCTOS.length}`;
  const activo = filtros.forma || filtros.color || filtros.marca || filtros.q;
  $('#limpiar').hidden = !activo;

  document.querySelectorAll('[data-forma]').forEach(b =>
    b.setAttribute('aria-pressed', String(b.dataset.forma === filtros.forma)));
  document.querySelectorAll('[data-color]').forEach(b =>
    b.setAttribute('aria-pressed', String(b.dataset.color === filtros.color)));
  document.querySelectorAll('[data-marca]').forEach(b =>
    b.setAttribute('aria-pressed', String(b.dataset.marca === filtros.marca)));
}

function abrir(p){
  const titulo = [p.marca, p.modelo, p.nombre].filter(Boolean).join(' ');
  const msg = encodeURIComponent(
    `Hola! Me interesan las ${titulo}` +
    (p.tono ? ` en ${p.tono.toLowerCase()}` : '') +
    ` (ref. ${p.ref}). ¿Están disponibles?`);

  const fila = (k,v) => v ? `<tr><th>${k}</th><td>${v}</td></tr>` : '';
  $('#panel').innerHTML = `
    <button class="cerrar" aria-label="Cerrar" onclick="cerrar()">&times;</button>
    <img src="${p.img}" alt="${titulo}">
    <span class="eyebrow">${p.marca}</span>
    <h2 id="pt">${p.modelo || p.ref}</h2>
    <p class="nom">${p.nombre || ''}</p>
    <table class="datos">
      ${fila('Color', p.tono)}
      ${fila('Forma', p.forma !== 'Por definir' ? p.forma : '')}
      ${fila('Referencia', p.ref)}
      ${fila('Precio', p.precio ? pesos(p.precio) : 'Consultar')}
      ${fila('Disponibilidad', p.stock > 0 ? 'En stock' : 'Agotada')}
    </table>
    <div class="acciones">
      <a class="wa" href="https://wa.me/${WHATSAPP}?text=${msg}" target="_blank" rel="noopener">
        <svg width="19" height="19" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2a10 10 0 00-8.6 15L2 22l5.2-1.4A10 10 0 1012 2zm5.8 14.2c-.2.7-1.4 1.3-2 1.4-.5.1-1.1.1-1.8-.1-.4-.1-1-.3-1.7-.6-3-1.3-4.9-4.3-5-4.5-.2-.2-1.2-1.6-1.2-3s.7-2.1 1-2.4c.2-.3.5-.4.7-.4h.5c.2 0 .4 0 .6.5l.8 2c.1.2.1.3 0 .5l-.4.5-.3.3c-.1.1-.2.3 0 .5.2.3.8 1.3 1.7 2.1 1.1 1 2 1.3 2.3 1.4.2.1.4.1.5-.1l.8-.9c.2-.2.3-.2.5-.1l2 1c.2.1.4.2.4.3.1.1.1.6-.1 1.2z"/></svg>
        Pedir por WhatsApp
      </a>
      ${p.probador ? `<a class="ver" href="${p.link}" target="_blank" rel="noopener nofollow">Ver fotos y probártelas →</a>` : ''}
    </div>`;
  $('#velo').classList.add('abierto');
  document.body.style.overflow = 'hidden';
  $('#panel').querySelector('.cerrar').focus();
}

function cerrar(){
  $('#velo').classList.remove('abierto');
  document.body.style.overflow = '';
}
$('#velo').addEventListener('click', e => { if(e.target.id === 'velo') cerrar(); });
document.addEventListener('keydown', e => { if(e.key === 'Escape') cerrar(); });

pintar();
</script>
</body>
</html>"""


def construir(productos):
    formas = {f: SIGLA[f] for f in
              ['Aviador', 'Wayfarer', 'Clubmaster', 'Redonda', 'Hexagonal', 'Gatuna', 'Cuadrada']
              if any(p['forma'] == f for p in productos)}
    tintes = [[n, TINTES[n]] for n in TINTES
              if any(p['color'] == n for p in productos)]
    marcas = sorted({p['marca'] for p in productos})

    salida = PLANTILLA
    for clave, valor in [
        ('__DATOS__', json.dumps(productos, ensure_ascii=False, separators=(',', ':'))),
        ('__FORMAS__', json.dumps(formas, ensure_ascii=False)),
        ('__TINTES__', json.dumps(tintes, ensure_ascii=False)),
        ('__MARCAS__', json.dumps(marcas, ensure_ascii=False)),
        ('__TIENDA__', html.escape(TIENDA)),
        ('__SUBTITULO__', html.escape(SUBTITULO)),
        ('__WHATSAPP__', WHATSAPP),
    ]:
        salida = salida.replace(clave, valor)
    return salida


if __name__ == '__main__':
    productos = cargar(EXCEL)
    open(SALIDA, 'w', encoding='utf-8').write(construir(productos))
    con_precio = sum(1 for p in productos if p['precio'])
    print(f'{SALIDA} generado con {len(productos)} monturas ({con_precio} con precio).')
    if WHATSAPP == '573001234567':
        print('¡Falta tu número! Cambia WHATSAPP arriba en este archivo.')
