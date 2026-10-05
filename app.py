#!/usr/bin/env python3
"""
app.py — Interfaz web local para gendoc.

Levanta un servidor en http://localhost:5000 donde podés subir una plantilla
.docx, pegar un JSON y ver el PDF y/o Word generado (o el diagnóstico de error)
al instante. Reutiliza toda la lógica de gendoc.py.

Correr:
    pip install flask docxtpl        # + LibreOffice instalado en el sistema
    python app.py
"""

import json
import tempfile
import uuid
from pathlib import Path

from flask import Flask, request, jsonify, send_file, Response

import gendoc  # reutiliza cargar_plantilla / analizar_variables / renderizar / a_pdf / GenError

app = Flask(__name__)

# Archivos generados en esta sesión: id -> (ruta, nombre_descarga, mimetype)
_TMP = Path(tempfile.mkdtemp(prefix="gendoc_ui_"))
_ARCHIVOS: dict[str, tuple[Path, str, str]] = {}
_MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _parse_json(texto: str):
    """Parsea el JSON pegado y devuelve (datos, avisos). En error levanta un
    GenError con lineno/colno para que el front resalte la línea."""
    return gendoc.parsear_json(texto)


def _err_dict(e: gendoc.GenError) -> dict:
    return {
        "ok": False,
        "error": {
            "etapa": e.etapa,
            "detalle": e.detalle,
            "ubicacion": e.ubicacion,
            "sugerencia": e.sugerencia,
            "lineno": getattr(e, "lineno", None),
        },
    }


def _guardar_entradas(req):
    """Guarda la plantilla subida y el texto JSON en archivos temporales.
    Devuelve (docx_path, datos_dict, nombre_base, avisos) o levanta GenError."""
    f = req.files.get("plantilla")
    if not f or not f.filename:
        raise gendoc.GenError("entrada", "Falta la plantilla .docx.")
    if not f.filename.lower().endswith(".docx"):
        raise gendoc.GenError("entrada", "La plantilla tiene que ser un archivo .docx.")
    texto = req.form.get("datos", "")
    if not texto.strip():
        raise gendoc.GenError("entrada", "Falta el JSON de datos.")

    datos, avisos = _parse_json(texto)
    docx_path = _TMP / f"in_{uuid.uuid4().hex}.docx"
    f.save(str(docx_path))
    return docx_path, datos, Path(f.filename).stem, avisos


# --------------------------------------------------------------------------- #
# API
# --------------------------------------------------------------------------- #
@app.post("/api/check")
def api_check():
    try:
        docx_path, datos, _, avisos = _guardar_entradas(request)
        doc = gendoc.cargar_plantilla(docx_path)
        an = gendoc.analizar_variables(doc, datos)
        return jsonify({
            "ok": True,
            "pistas": gendoc.diagnosticar_faltantes(doc, an.faltantes),
            "esperadas": an.esperadas,
            "faltantes": an.faltantes,
            "sin_usar": an.sin_usar,
            "opcionales": an.opcionales,
            "atributos": an.atributos,
            "envoltorios": an.envoltorios,
            "avisos": avisos,
        })
    except gendoc.GenError as e:
        return jsonify(_err_dict(e)), 200


@app.post("/api/generate")
def api_generate():
    try:
        docx_path, datos, nombre, avisos = _guardar_entradas(request)
        doc = gendoc.cargar_plantilla(docx_path)

        # Pre-chequeo: variables de nivel superior faltantes (lista completa).
        # analizar_variables desanida el JSON si viene envuelto en una clave.
        an = gendoc.analizar_variables(doc, datos)
        if an.faltantes:
            pistas = gendoc.diagnosticar_faltantes(doc, an.faltantes)
            raise gendoc.GenError(
                "datos incompletos",
                "Faltan variables que la plantilla usa: " + ", ".join(an.faltantes)
                + ("\n\n" + "\n".join(pistas.values()) if pistas else ""),
                sugerencia="Agregalas al JSON. Con «Validar» ves todas de una."
                if not pistas else
                "Esa variable no va en el JSON: hay que arreglar la etiqueta en el Word.",
            )
        if an.atributos:
            raise gendoc.GenError(
                "datos incompletos",
                "La plantilla usa atributos que el JSON no trae:\n"
                + "\n".join("• " + a for a in an.atributos),
                sugerencia="Agregalos al JSON, o corregí/sacá esa etiqueta en el Word.",
            )

        formato = request.form.get("formato", "pdf")
        docx_path = _TMP / f"o_{uuid.uuid4().hex}.docx"
        pdf_path = _TMP / f"o_{uuid.uuid4().hex}.pdf"
        generados = gendoc.generar(doc, an.contexto, formato, pdf_path, docx_path)

        nombre = nombre or "documento"
        resp = {"ok": True, "avisos": avisos}
        for path in generados:
            ext = path.suffix.lstrip(".")
            fid = uuid.uuid4().hex
            _ARCHIVOS[fid] = (path, f"{nombre}.{ext}", _MIME[ext])
            resp[ext] = {"url": f"/archivo/{fid}", "nombre": f"{nombre}.{ext}"}
        return jsonify(resp)
    except gendoc.GenError as e:
        return jsonify(_err_dict(e)), 200


@app.get("/archivo/<fid>")
def archivo(fid):
    entry = _ARCHIVOS.get(fid)
    if not entry:
        return "Archivo no encontrado", 404
    path, nombre, mime = entry
    descarga = request.args.get("download") == "1"
    return send_file(path, mimetype=mime,
                     as_attachment=descarga, download_name=nombre)


@app.get("/")
def index():
    return Response(PAGE, mimetype="text/html")


# --------------------------------------------------------------------------- #
# Frontend (una sola página)
# --------------------------------------------------------------------------- #
PAGE = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>gendoc · probador de plantillas</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Hanken+Grotesk:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  :root{
    --desk:#E8E9E4; --surface:#fff; --ink:#171A2B; --ink-soft:#5B5F72;
    --line:#DBDCD6; --line-strong:#C7C8C1;
    --accent:#2E37B8; --accent-dark:#242CA0; --accent-tint:#EDEEFB;
    --err:#B42318; --err-bg:#FDF3F2; --err-line:#F3C9C4;
    --ok:#12703A; --ok-bg:#EDF8F0;
    --radius:10px;
  }
  *{box-sizing:border-box}
  html,body{margin:0;height:100%}
  body{
    background:var(--desk); color:var(--ink);
    font-family:"Hanken Grotesk",system-ui,sans-serif;
    font-size:15px; line-height:1.5; -webkit-font-smoothing:antialiased;
    display:flex; flex-direction:column; min-height:100%;
  }
  code,.mono{font-family:"JetBrains Mono",ui-monospace,monospace}

  header{
    padding:18px 26px; border-bottom:1px solid var(--line);
    background:var(--surface); display:flex; align-items:baseline; gap:16px; flex-wrap:wrap;
  }
  .brand{font-weight:700; font-size:21px; letter-spacing:-.02em; display:flex; align-items:center; gap:2px}
  .brand .caret{width:9px;height:19px;background:var(--accent);display:inline-block;
    border-radius:1px;transform:translateY(2px);animation:blink 1.15s steps(1) infinite}
  @keyframes blink{50%{opacity:.15}}
  .tagline{color:var(--ink-soft); font-size:14px}
  .tagline b{color:var(--ink);font-weight:600}

  main{
    flex:1; display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1.05fr);
    gap:18px; padding:18px 26px; align-items:stretch;
  }
  .panel{background:var(--surface); border:1px solid var(--line); border-radius:var(--radius);
    display:flex; flex-direction:column; min-height:0}
  .panel > h2{
    margin:0; padding:13px 16px; font-size:13px; font-weight:600; color:var(--ink-soft);
    border-bottom:1px solid var(--line); letter-spacing:.01em;
    display:flex; justify-content:space-between; align-items:center; gap:10px;
  }
  .panel-body{padding:16px; display:flex; flex-direction:column; gap:14px; flex:1; min-height:0}

  /* Dropzone plantilla */
  .drop{
    border:1.5px dashed var(--line-strong); border-radius:8px; padding:16px;
    display:flex; align-items:center; gap:12px; cursor:pointer; transition:border-color .12s, background .12s;
  }
  .drop:hover{border-color:var(--accent); background:var(--accent-tint)}
  .drop.armed{border-color:var(--accent); background:var(--accent-tint)}
  .drop.has{border-style:solid; border-color:var(--ok); background:var(--ok-bg)}
  .drop .ic{width:34px;height:34px;flex:none;border-radius:7px;background:var(--accent-tint);
    color:var(--accent);display:grid;place-items:center}
  .drop.has .ic{background:#fff;color:var(--ok)}
  .drop .t{font-weight:600}
  .drop .s{color:var(--ink-soft); font-size:13px}
  .drop.has .fname{font-family:"JetBrains Mono",monospace;font-size:13px}

  /* Editor JSON con gutter */
  .editor-label{font-size:13px;color:var(--ink-soft);display:flex;justify-content:space-between}
  .editor{flex:1; min-height:220px; display:flex; border:1px solid var(--line-strong);
    border-radius:8px; overflow:hidden; background:#FCFCFB}
  .editor:focus-within{border-color:var(--accent); box-shadow:0 0 0 3px var(--accent-tint)}
  .gutter{
    padding:12px 8px 12px 0; text-align:right; color:#A9AAA2; user-select:none;
    font-family:"JetBrains Mono",monospace; font-size:13px; line-height:1.55;
    background:#F4F4F1; border-right:1px solid var(--line); min-width:44px; overflow:hidden;
  }
  .gutter .ln.err{color:#fff;background:var(--err);border-radius:3px;padding:0 5px;margin:0 -5px 0 0;font-weight:600}
  textarea#json{
    flex:1; border:0; outline:0; resize:none; background:transparent; color:var(--ink);
    padding:12px 12px; font-family:"JetBrains Mono",monospace; font-size:13px; line-height:1.55;
    white-space:pre; overflow:auto; tab-size:2;
  }

  .actions{display:flex; gap:10px}
  button{font-family:inherit; font-size:14px; font-weight:600; border-radius:8px;
    padding:10px 18px; cursor:pointer; border:1px solid transparent; transition:background .12s,border-color .12s}
  button:disabled{opacity:.55; cursor:not-allowed}
  .btn-primary{background:var(--accent); color:#fff}
  .btn-primary:hover:not(:disabled){background:var(--accent-dark)}
  .btn-ghost{background:#fff; color:var(--ink); border-color:var(--line-strong)}
  .btn-ghost:hover:not(:disabled){border-color:var(--accent); color:var(--accent)}

  .seg{display:inline-flex; border:1px solid var(--line-strong); border-radius:8px; overflow:hidden; margin-left:auto}
  .seg label{position:relative}
  .seg input{position:absolute; opacity:0; pointer-events:none}
  .seg span{display:block; padding:10px 13px; font-size:13px; font-weight:600; color:var(--ink-soft); cursor:pointer;
    border-left:1px solid var(--line-strong); user-select:none}
  .seg label:first-child span{border-left:0}
  .seg input:checked + span{background:var(--accent-tint); color:var(--accent)}
  .seg input:focus-visible + span{outline:2px solid var(--accent); outline-offset:-2px}

  .wordcard{flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center;
    gap:10px; padding:30px; text-align:center; color:var(--ink-soft)}
  .wordcard .big{font-size:15px; color:var(--ink); font-weight:600}
  .wordcard .ic{width:52px;height:52px;border-radius:12px;background:var(--ok-bg);color:var(--ok);display:grid;place-items:center}

  /* Panel resultado */
  #result{flex:1; min-height:0; display:flex; flex-direction:column}
  .empty{flex:1; display:flex; flex-direction:column; align-items:center; justify-content:center;
    text-align:center; color:var(--ink-soft); gap:8px; padding:30px}
  .empty .big{font-size:15px; color:var(--ink); font-weight:600}
  .empty svg{opacity:.5}

  iframe#pdf{flex:1; width:100%; border:0; background:#f2f2f0; min-height:340px}
  .result-bar{display:flex; align-items:center; gap:10px; padding:10px 16px; border-top:1px solid var(--line)}
  .result-bar .fname{font-family:"JetBrains Mono",monospace; font-size:13px; color:var(--ink-soft); margin-right:auto}

  /* Diagnóstico */
  .diag{margin:16px; border:1px solid var(--err-line); background:var(--err-bg); border-radius:9px; overflow:hidden}
  .diag .head{display:flex; align-items:center; gap:9px; padding:12px 15px; border-bottom:1px solid var(--err-line);
    color:var(--err); font-weight:700}
  .diag .rows{padding:6px 15px 14px}
  .diag .row{padding:9px 0; border-bottom:1px solid #f2ddda}
  .diag .row:last-child{border-bottom:0}
  .diag .k{font-size:12px; color:#93544d; margin-bottom:2px}
  .diag .v{color:var(--ink)}
  .diag .v.mono{font-family:"JetBrains Mono",monospace; font-size:13px; white-space:pre-wrap; word-break:break-word}

  /* Variables (validación) */
  .vars{margin:16px; display:flex; flex-direction:column; gap:16px}
  .vars .ok-banner{display:flex;align-items:center;gap:9px;color:var(--ok);background:var(--ok-bg);
    border:1px solid #C6E7CF;border-radius:9px;padding:11px 14px;font-weight:600}
  .vargroup h3{margin:0 0 8px; font-size:13px; color:var(--ink-soft); font-weight:600}
  .chips{display:flex; flex-wrap:wrap; gap:7px}
  .chip{font-family:"JetBrains Mono",monospace; font-size:13px; padding:4px 10px; border-radius:6px;
    border:1px solid var(--line-strong); background:#fff}
  .chip.miss{border-color:var(--err-line); background:var(--err-bg); color:var(--err); font-weight:600}
  .chip.ok{border-color:#C6E7CF; background:var(--ok-bg); color:var(--ok)}
  .chip.unused{color:var(--ink-soft)}
  .note{border:1px solid var(--line-strong);background:#FCFCFB;border-radius:9px;
    padding:11px 14px;color:var(--ink-soft);font-size:14px}
  .note code{background:var(--accent-tint);color:var(--accent);border-radius:4px;padding:1px 5px}
  .note.warn{margin:16px;border-color:#E8D6A8;background:#FDF8EC;color:#7A5B12}

  .spin{width:15px;height:15px;border:2px solid rgba(255,255,255,.45);border-top-color:#fff;
    border-radius:50%;display:inline-block;animation:sp .7s linear infinite;vertical-align:-2px;margin-right:7px}
  @keyframes sp{to{transform:rotate(360deg)}}

  @media (max-width:860px){
    main{grid-template-columns:1fr}
    iframe#pdf{min-height:420px}
  }
  @media (prefers-reduced-motion:reduce){*{animation:none!important}}
</style>
</head>
<body>
<header>
  <div class="brand">gendoc<span class="caret"></span></div>
  <div class="tagline">plantilla <b>.docx</b> con Jinja2 + <b>JSON</b> → <b>PDF</b> o <b>Word</b>. Subí, pegá y probá.</div>
</header>

<main>
  <!-- ENTRADAS -->
  <section class="panel">
    <h2>Entradas</h2>
    <div class="panel-body">
      <div class="drop" id="drop">
        <div class="ic" id="dropic">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7l5 5v11a2 2 0 0 1-2 2z"/></svg>
        </div>
        <div>
          <div class="t" id="droptitle">Plantilla .docx</div>
          <div class="s" id="dropsub">Arrastrala acá o hacé clic para elegir</div>
        </div>
        <input type="file" id="file" accept=".docx" hidden>
      </div>

      <div class="editor-label">
        <span>Datos (JSON)</span>
        <span id="jsonstatus"></span>
      </div>
      <div class="editor">
        <div class="gutter" id="gutter"><span class="ln">1</span></div>
        <textarea id="json" spellcheck="false" placeholder='Pegá acá el JSON de datos…'></textarea>
      </div>

      <div class="actions">
        <button class="btn-primary" id="btn-gen">Generar PDF</button>
        <button class="btn-ghost" id="btn-check">Validar</button>
        <div class="seg" role="radiogroup" aria-label="Formato de salida">
          <label><input type="radio" name="formato" value="pdf" checked><span>PDF</span></label>
          <label><input type="radio" name="formato" value="docx"><span>Word</span></label>
          <label><input type="radio" name="formato" value="ambos"><span>Ambos</span></label>
        </div>
      </div>
    </div>
  </section>

  <!-- RESULTADO -->
  <section class="panel">
    <h2>Resultado <span id="resmeta" style="color:var(--ink-soft);font-weight:400"></span></h2>
    <div id="result">
      <div class="empty">
        <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h5"/></svg>
        <div class="big">Todavía no generaste nada</div>
        <div>Subí una plantilla, pegá el JSON y tocá <b>Generar</b>. Elegí si querés PDF, Word o ambos.<br>Con <b>Validar</b> ves qué variables espera la plantilla sin generar nada.</div>
      </div>
    </div>
  </section>
</main>

<script>
const $ = s => document.querySelector(s);
const fileInput = $("#file"), drop = $("#drop"), ta = $("#json"), gutter = $("#gutter");
const result = $("#result"), resmeta = $("#resmeta"), jsonstatus = $("#jsonstatus");
let errLine = null;

/* --- Dropzone --- */
drop.onclick = () => fileInput.click();
drop.ondragover = e => { e.preventDefault(); drop.classList.add("armed"); };
drop.ondragleave = () => drop.classList.remove("armed");
drop.ondrop = e => {
  e.preventDefault(); drop.classList.remove("armed");
  const f = e.dataTransfer.files[0];
  if (f) { fileInput.files = e.dataTransfer.files; setFile(f); }
};
fileInput.onchange = () => { if (fileInput.files[0]) setFile(fileInput.files[0]); };
function setFile(f){
  drop.classList.add("has");
  $("#droptitle").innerHTML = 'Plantilla cargada';
  $("#dropsub").innerHTML = '<span class="fname">' + escapeHtml(f.name) + '</span>';
  $("#dropic").innerHTML = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M20 6 9 17l-5-5"/></svg>';
}

/* --- Editor: gutter de líneas + resaltado de error --- */
function renderGutter(){
  const n = ta.value.split("\n").length || 1;
  let h = "";
  for (let i = 1; i <= n; i++)
    h += '<span class="ln' + (i === errLine ? ' err' : '') + '">' + i + '</span>\n';
  gutter.innerHTML = h;
  gutter.scrollTop = ta.scrollTop;
}
ta.addEventListener("input", () => { errLine = null; renderGutter(); });
ta.addEventListener("scroll", () => { gutter.scrollTop = ta.scrollTop; });
renderGutter();

/* --- Llamadas --- */
function payload(){
  const fd = new FormData();
  if (fileInput.files[0]) fd.append("plantilla", fileInput.files[0]);
  fd.append("datos", ta.value);
  fd.append("formato", formato());
  return fd;
}
const ETIQUETAS = { pdf: "Generar PDF", docx: "Generar Word", ambos: "Generar PDF + Word" };
function formato(){ return document.querySelector('input[name="formato"]:checked').value; }
function syncBtn(){ const b = $("#btn-gen"); if (!b.disabled) b.textContent = ETIQUETAS[formato()]; }
document.querySelectorAll('input[name="formato"]').forEach(r => r.onchange = syncBtn);
try{ const f = localStorage.getItem("gendoc.formato");
     const r = f && document.querySelector('input[name="formato"][value="' + f + '"]');
     if (r){ r.checked = true; syncBtn(); } }catch(e){}
document.querySelectorAll('input[name="formato"]').forEach(r => r.addEventListener("change", () => {
  try{ localStorage.setItem("gendoc.formato", formato()); }catch(e){} }));
function busy(btn, on, label){
  btn.disabled = on;
  if (on){ btn.dataset.txt = btn.innerHTML; btn.innerHTML = '<span class="spin"></span>' + label; }
  else if (btn.dataset.txt){ btn.innerHTML = btn.dataset.txt; }
}

$("#btn-gen").onclick = async () => {
  const btn = $("#btn-gen"); busy(btn, true, "Generando…"); resmeta.textContent = "";
  try{
    const r = await fetch("/api/generate", { method:"POST", body: payload() });
    const d = await r.json();
    if (d.ok) showResult(d); else showError(d.error);
  }catch(e){ showError({ etapa:"conexión", detalle:String(e) }); }
  finally{ busy(btn, false); syncBtn(); }
};
$("#btn-check").onclick = async () => {
  const btn = $("#btn-check"); busy(btn, true, "Validando…"); resmeta.textContent = "";
  try{
    const r = await fetch("/api/check", { method:"POST", body: payload() });
    const d = await r.json();
    if (d.ok) showVars(d); else showError(d.error);
  }catch(e){ showError({ etapa:"conexión", detalle:String(e) }); }
  finally{ busy(btn, false); }
};

/* --- Vistas de resultado --- */
function showResult(d){
  errLine = null; renderGutter();
  resmeta.textContent = "· " + [d.pdf && "PDF", d.docx && "Word"].filter(Boolean).join(" + ") + " listo";
  const boton = (f, txt, cls) => '<a href="' + f.url + '?download=1"><button class="' + cls + '">' + txt + '</button></a>';
  // El navegador puede mostrar el PDF; el Word sólo se descarga.
  const vista = d.pdf
    ? '<iframe id="pdf" src="' + d.pdf.url + '#toolbar=1"></iframe>'
    : '<div class="wordcard"><div class="ic"><svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 3v4a1 1 0 0 0 1 1h4"/><path d="M17 21H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h7l5 5v11a2 2 0 0 1-2 2z"/><path d="m8 12 1.5 5 2.5-4 2.5 4 1.5-5"/></svg></div>' +
      '<div class="big">Word generado</div><div>El navegador no puede previsualizar .docx: descargalo para abrirlo en Word.</div></div>';
  const principal = d.pdf || d.docx;
  result.innerHTML =
    avisosHtml(d.avisos) + vista +
    '<div class="result-bar"><span class="fname">' + escapeHtml(principal.nombre) + (d.pdf && d.docx ? ' · .docx' : '') + '</span>' +
    (d.docx ? boton(d.docx, d.pdf ? "Descargar Word" : "Descargar", d.pdf ? "btn-ghost" : "btn-primary") : '') +
    (d.pdf ? boton(d.pdf, d.docx ? "Descargar PDF" : "Descargar", "btn-primary") : '') +
    '</div>';
}

function showVars(d){
  errLine = null; renderGutter();
  resmeta.textContent = "· validación";
  let html = '<div class="vars">';
  html += avisosHtml(d.avisos);
  if (d.envoltorios && d.envoltorios.length)
    html += '<div class="note">El JSON venía envuelto en ' +
      d.envoltorios.map(k => '<code>' + escapeHtml(k) + '</code>').join(' → ') +
      '; se abrió ese nivel para resolver las variables.</div>';
  const attrs = d.atributos || [];
  if (d.faltantes.length === 0 && attrs.length === 0)
    html += '<div class="ok-banner"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M20 6 9 17l-5-5"/></svg>La plantilla y los datos son compatibles.</div>';
  const opc = d.opcionales || [];
  html += '<div class="vargroup"><h3>Variables que la plantilla espera (' + d.esperadas.length + ')</h3><div class="chips">' +
    (d.esperadas.map(v => '<span class="chip ' + (d.faltantes.includes(v) ? 'miss' : opc.includes(v) ? 'unused' : 'ok') + '">' + escapeHtml(v) + '</span>').join('') || '<span class="chip">ninguna</span>') +
    '</div></div>';
  if (d.faltantes.length){
    html += '<div class="vargroup"><h3 style="color:var(--err)">Faltan en el JSON (' + d.faltantes.length + ')</h3><div class="chips">' +
      d.faltantes.map(v => '<span class="chip miss">' + escapeHtml(v) + '</span>').join('') + '</div>';
    const pistas = Object.values(d.pistas || {});
    if (pistas.length)
      html += pistas.map(p => '<div class="note warn" style="margin:10px 0 0">' + escapeHtml(p) + '</div>').join('');
    html += '</div>';
  }
  if (attrs.length)
    html += '<div class="vargroup"><h3 style="color:var(--err)">Atributos que faltan en el JSON (' + attrs.length + ')</h3>' +
      attrs.map(a => '<div class="note warn" style="margin:0 0 8px">' + escapeHtml(a) + '</div>').join('') + '</div>';
  if (opc.length)
    html += '<div class="vargroup"><h3>Opcionales que no vienen en el JSON (' + opc.length + ')</h3><div class="chips">' +
      opc.map(v => '<span class="chip unused">' + escapeHtml(v) + '</span>').join('') + '</div>' +
      '<div class="note" style="margin-top:10px">Sólo se usan dentro de su propio <code>{% if … %}</code>: como no vienen, esos bloques no se imprimen.</div></div>';
  if (d.sin_usar.length)
    html += '<div class="vargroup"><h3>En el JSON pero sin usar</h3><div class="chips">' +
      d.sin_usar.map(v => '<span class="chip unused">' + escapeHtml(v) + '</span>').join('') + '</div></div>';
  html += '</div>';
  result.innerHTML = html;
}

function showError(err){
  resmeta.textContent = "· error";
  errLine = err && err.lineno ? err.lineno : null;
  renderGutter();
  if (errLine){
    const lh = 1.55 * 13;
    ta.scrollTop = Math.max(0, (errLine - 4) * lh);
    gutter.scrollTop = ta.scrollTop;
  }
  const rows = [
    ["Etapa", err.etapa, false],
    ["Ubicación", err.ubicacion, true],
    ["Detalle", err.detalle, true],
    ["Sugerencia", err.sugerencia, false],
  ].filter(r => r[1]);
  result.innerHTML =
    '<div class="diag"><div class="head">' +
    '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/></svg>' +
    'No se pudo generar</div><div class="rows">' +
    rows.map(([k, v, m]) => '<div class="row"><div class="k">' + k + '</div><div class="v' + (m ? ' mono' : '') + '">' + escapeHtml(String(v)) + '</div></div>').join('') +
    '</div></div>';
}

function avisosHtml(avisos){
  if (!avisos || !avisos.length) return '';
  return avisos.map(a => '<div class="note warn">' + escapeHtml(a) + '</div>').join('');
}

function escapeHtml(s){ return s.replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
</script>
</body>
</html>"""


if __name__ == "__main__":
    print("gendoc UI → http://localhost:5000")
    app.run(host="127.0.0.1", port=5000, debug=False)
