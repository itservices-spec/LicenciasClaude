#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_google_site.py
====================
Genera 'dashboard_google_site.html': una versión LIGERA y AUTÓNOMA del
dashboard con los datos ya incrustados, pensada para incrustar en una página
de Google Sites (o cualquier web).

Qué hace respecto a 'dashboard_licencias.html':
  - Incrusta los datos (pestañas «Asignación» y «Usabilidad») como JSON, así
    el tablero se renderiza solo (sin cargar archivo).
  - Elimina la librería SheetJS (solo servía para leer/exportar .xlsx),
    reduciendo el tamaño de ~950 KB a ~140 KB.
  - Quita la zona de carga de archivo; la exportación pasa a CSV (sin
    dependencias).

Uso:
    python3 build_google_site.py [archivo_entrada.xlsx] [dashboard_base.html]
"""

import sys
import re
import json
import math
import pandas as pd

ENTRADA = sys.argv[1] if len(sys.argv) > 1 else "Usuarios de licencias Claude con CeCo.xlsx"
BASE = sys.argv[2] if len(sys.argv) > 2 else "dashboard_licencias.html"
SALIDA = "dashboard_google_site.html"


def hoja_a_json(sheet):
    df = pd.read_excel(ENTRADA, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    out = []
    for _, r in df.iterrows():
        o = {}
        for c in df.columns:
            v = r[c]
            if isinstance(v, float) and math.isnan(v):
                o[c] = None
            elif pd.api.types.is_datetime64_any_dtype(df[c].dtype) or hasattr(v, "isoformat"):
                try:
                    o[c] = pd.Timestamp(v).strftime("%Y-%m-%d") if pd.notna(v) else None
                except Exception:
                    o[c] = None
            elif hasattr(v, "item"):
                o[c] = v.item()
            else:
                o[c] = v
        out.append(o)
    return out


def main():
    asig = hoja_a_json("Asignación")
    usab = hoja_a_json("Usabilidad")
    baked = ("const ASIG_DATA=" + json.dumps(asig, ensure_ascii=False) + ";\n"
             "const USAB_DATA=" + json.dumps(usab, ensure_ascii=False) + ";\n")

    html = open(BASE, encoding="utf-8").read()

    # 1) Quitar SheetJS incrustado (o CDN).
    i = html.index("<script>/* SheetJS")
    j = html.index("</script>", i) + len("</script>")
    html = html[:i] + "<!-- SheetJS removido: versión ligera para incrustar -->" + html[j:]
    html = re.sub(r'<script src="https://cdn\.sheetjs\.com[^"]*"></script>', "", html)

    # 2) Cabecera: quitar botones de carga; dejar Tema + Exportar CSV.
    for s in [
        '<label class="btn ghost" for="fileInput" title="Cargar un archivo .xlsx (alternativa)">⬆ Archivo</label>',
        '<label class="btn primary" for="fileInput">⬆ Cargar / actualizar datos</label>',
        '<input id="fileInput" type="file" accept=".xlsx,.xls" hidden>',
        '<button class="btn primary" id="updateBtn" title="Leer los datos desde Google Drive">⟳ Actualizar datos</button>',
    ]:
        html = html.replace(s, "")
    html = html.replace(
        '<div class="sub">Integración Asignación × Usabilidad — carga y actualización recurrente</div>',
        '<div class="sub">Análisis de usabilidad · vista embebida</div>')
    html = html.replace(
        '<div class="sub">Datos desde Google Drive · Asignación × Usabilidad</div>',
        '<div class="sub">Análisis de usabilidad · vista embebida</div>')

    # 3) Quitar el dropzone (empty state).
    di = html.index('<div id="dropzone"')
    de = html.index('<div id="errBox">', di)
    html = html[:di] + html[de:]

    # 4) App visible de inicio; neutralizar ocultamiento de dropzone.
    html = html.replace('#app{display:none}', '#app{display:block}')
    html = html.replace('$("#dropzone").style.display="none";', '')

    # 5) Neutralizar wiring de archivo/dropzone y arranque por localStorage.
    html = html.replace(
        '$("#fileInput").addEventListener("change",e=>{ if(e.target.files[0]) manejarArchivo(e.target.files[0]); e.target.value=""; });', "")
    html = html.replace('$("#exportBtn").addEventListener("click", exportar);', "")
    _dz = html.index('const dz=$("#dropzone");')
    _tema = html.index('// Tema', _dz)
    html = html[:_dz] + html[_tema:]
    html = html.replace('window.addEventListener("DOMContentLoaded", cargarLocal);', "")
    html = html.replace('if(document.readyState!=="loading") cargarLocal();', "")

    # 6) Bootstrap: datos incrustados + render + exportación CSV.
    bootstrap = r"""
/* ====== Bootstrap para incrustar (Google Sites): datos + render directo ====== */
__BAKED__
function revive(rows){ return rows.map(r=>{ const o={...r};
  for(const k of ["Fecha","Fecha Asig","Last Active"]) if(o[k]!=null) o[k]=String(o[k]);
  return o; }); }
function exportarCSV(){
  if(!window.STATE) return;
  const cols=["Nombre","Correo electrónico","User Team","Licencia","ID licencia","Movimiento",
    "Fecha","Días calendario","Días hábiles","Días sin uso","Días desde últ. conexión",
    "Days Active","% de uso","Nivel de uso","Amplitud de adopción (0-5)","Acción sugerida",
    "Costo anual (USD)","Facturado (MXN)","Presupuesto Ops","Cantidad lic"];
  const esc=v=>{ if(v==null) return ""; if(v instanceof Date) v=v.toISOString().slice(0,10);
    v=String(v); return /[",\n;]/.test(v)? '"'+v.replace(/"/g,'""')+'"' : v; };
  const lines=[cols.join(",")];
  for(const r of window.STATE.rowsFull) lines.push(cols.map(c=>esc(r[c])).join(","));
  const blob=new Blob(["﻿"+lines.join("\r\n")],{type:"text/csv;charset=utf-8;"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob);
  a.download="Analisis_Usabilidad.csv"; document.body.appendChild(a); a.click(); a.remove();
}
(function(){
  const old=document.getElementById("exportBtn");
  if(old){ const b=old.cloneNode(true); b.disabled=false; b.textContent="⬇ Exportar CSV";
    old.replaceWith(b); b.addEventListener("click", exportarCSV); }
  try{ renderAll(procesar(revive(ASIG_DATA), revive(USAB_DATA))); }
  catch(e){ document.getElementById("errBox").innerHTML=
    '<div class="err">Error al renderizar: '+e.message+'</div>'; console.error(e); }
})();
"""
    bootstrap = bootstrap.replace("__BAKED__", baked)
    last = html.rindex("</script>")
    html = html[:last] + bootstrap + "\n" + html[last:]

    open(SALIDA, "w", encoding="utf-8").write(html)
    print(f"Generado {SALIDA} -> {round(len(html)/1024,1)} KB "
          f"({len(asig)} filas Asignación, {len(usab)} filas Usabilidad)")


if __name__ == "__main__":
    main()
