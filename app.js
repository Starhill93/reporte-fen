/* Reporte FEN — INABIF
 * Lee el Excel del SIAF (hoja SheetGasto) en el navegador y arma una lámina 16:9
 * con el mismo estilo que REPORTE_PRESUPUESTAL_INABIF.pptx (scraper_inabif.py).
 * Nada se envía a ningún servidor.
 */
(() => {
  "use strict";

  // ── Configuración ─────────────────────────────────────────────────────
  const ACT_FEN = "5006412";
  const NOMBRES_DEFECTO = {          // nombre de cada meta (sec_func) del FEN, como en el reporte del SIAF
    "213": "MANTENIMIENTO E INFRAESTRUCTURA",
    "214": "FORTALECIMIENTO DE CAPACIDADES",
    "215": "BIENES Y SERVICIOS",
  };
  const MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", "NOV", "DIC"];
  const MESES_LARGOS = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
    "setiembre", "octubre", "noviembre", "diciembre"];
  const NIVEL1 = [
    { id: "sec_func", t: "Meta" },
    { id: "Genérica (det)", t: "Genérica" },
    { id: "UNIDAD FUNCIONAL", t: "Unidad funcional" },
    { id: "DEPENDENCIA", t: "Dependencia" },
    { id: "departamento_meta", t: "Departamento" },
    { id: "fuente_financ", t: "Fuente de financiamiento" },
  ];
  const NIVEL2 = [
    { id: "CONDICION 2", t: "Condición" },
    { id: "CONDICION", t: "Condición (agrupada)" },
    { id: "Sub Genérica (det)", t: "Subgenérica" },
    { id: "EEGG", t: "Específica de gasto" },
    { id: "departamento_meta", t: "Departamento" },
    { id: "DEPENDENCIA", t: "Dependencia" },
  ];
  const FILTROS = [
    { el: "fFuente", col: "fuente_financ" },
    { el: "fUnidad", col: "UNIDAD FUNCIONAL" },
    { el: "fDepto", col: "departamento_meta" },
  ];
  const OBLIGATORIAS = ["activ_obra_accinv", "sec_func", "mto_pim", "mto_certificado", "mto_compro_anual"];
  // columnas de texto que se guardan de cada fila (también viajan en el enlace)
  const CAMPOS = [...new Set([...NIVEL1, ...NIVEL2].map(n => n.id).concat(FILTROS.map(f => f.col),
    ["activ_obra_accinv", "unidad_ejecutora", "sec_ejec", "DEPENDENCIA", "finalidad", "EEGG"]))];

  // Tokens y medidas del PPT (scraper_inabif.py)
  const C = { rojo: "#B51F35", navy: "#0A3A70", coral: "#E05B55", ocre: "#CE985C", verde: "#2E7D6E", gestion: "#5B6B85", linea: "#E6EAF0" };
  const IN = 128, PT = 128 / 72;                       // 1 in y 1 pt en px de la lámina
  const MIN_FILA = { 6: 0.145, 5.5: 0.135, 5: 0.128 }; // alto mínimo de fila (in) según la letra
  const semaforo = p => (p < 0.6 ? C.rojo : p < 0.9 ? C.ocre : C.verde);

  // ── Estado ────────────────────────────────────────────────────────────
  const $ = id => document.getElementById(id);
  let REG = [], ARCHIVO = null, VISTA = null;
  const nombres = leerLS("reporteFen.nombres", {});

  function leerLS(k, def) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch { return def; } }
  function guardarLS(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* sin almacenamiento */ } }

  // ── Formatos (iguales al PPT) ─────────────────────────────────────────
  const num = v => { if (v == null || v === "") return 0; const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, "")); return isFinite(n) ? n : 0; };
  const pct = (a, b) => (b ? a / b : 0);
  const nf = {};
  const fmt = (v, d = 0) => (nf[d] ||= new Intl.NumberFormat("en-US", { minimumFractionDigits: d, maximumFractionDigits: d })).format(Math.abs(v) < 0.005 ? 0 : v);
  const miles = v => fmt(v, 0);
  const dec = v => fmt(v, 2);
  const fpct = v => `${(v * 100).toFixed(1)}%`;
  const soles = v => {
    const a = Math.abs(v);
    if (a >= 1e6) return `S/ ${fmt(v / 1e6, 2)} mill`;
    if (a >= 1e5) return `S/ ${fmt(v / 1e3, 0)} mil`;
    if (a >= 1e3) return `S/ ${fmt(v / 1e3, 1)} mil`;
    return `S/ ${fmt(v, 0)}`;
  };
  const txt = v => (v == null ? "" : String(v).trim());
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const natural = (a, b) => a.localeCompare(b, "es", { numeric: true, sensitivity: "base" });
  const sinCodigo = s => s.replace(/^[\d.]+\s*\.?\s*/, "");
  const ddmmyyyy = iso => { if (!iso) return ""; const [y, m, d] = iso.split("-"); return `${d}/${m}/${y}`; };
  const isoDe = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  const vacio = () => ({ pto: 0, pim: 0, cert: 0, comp: 0, dev: 0, compM: Array(12).fill(0), devM: Array(12).fill(0) });
  const acumular = (a, r) => {
    a.pto += r.pto; a.pim += r.pim; a.cert += r.cert; a.comp += r.comp; a.dev += r.dev;
    for (let i = 0; i < 12; i++) { a.compM[i] += r.compM[i]; a.devM[i] += r.devM[i]; }
    return a;
  };
  const esCero = r => [r.pto, r.pim, r.cert, r.comp, r.dev].every(v => Math.abs(v) < 0.005);
  const estado = (m, error = false) => { const e = $("estado"); e.textContent = m; e.classList.toggle("error", error); };

  // Columnas de la tabla (las del reporte original del SIAF, sin PTO inicial)
  const COLS = [
    { t: "DETALLE / CONDICIÓN", w: 2.6, a: "l" },
    { t: "PIM S/", w: 0.85, f: t => miles(t.pim) },
    { t: "CERT. S/", w: 0.95, f: t => dec(t.cert) },
    { t: "CERT. %", w: 0.55, f: t => fpct(pct(t.cert, t.pim)) },
    { t: "SALDxCERT S/", w: 0.92, f: t => dec(t.pim - t.cert) },
    { t: "COMP. ANUAL S/", w: 0.98, f: t => dec(t.comp) },
    { t: "COMP. ANUAL %", w: 0.72, f: t => fpct(pct(t.comp, t.pim)) },
    { t: "SALDxCOMP S/", w: 0.92, f: t => dec(t.pim - t.comp) },
    { t: "EJEC. S/", w: 0.82, f: t => dec(t.dev), ejec: true },
    { t: "EJEC. %", w: 0.55, f: t => fpct(pct(t.dev, t.pim)) },
    { t: "SALDxDEVENG S/", w: 0.98, f: t => dec(t.pim - t.dev) },
  ];

  // ── Lectura del Excel ─────────────────────────────────────────────────
  async function cargar(file) {
    estado("Leyendo " + file.name + "…");
    await new Promise(r => setTimeout(r, 30));
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array", dense: true });
      const norm = c => txt(c).toLowerCase();
      let filas = null;
      for (const n of ["SheetGasto", ...wb.SheetNames.filter(n => n !== "SheetGasto")]) {
        if (!wb.Sheets[n]) continue;
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: null, raw: true });
        const h = aoa.slice(0, 30).findIndex(row => row && row.some(c => norm(c) === "activ_obra_accinv") && row.some(c => norm(c) === "mto_pim"));
        if (h >= 0) { filas = { aoa, h }; break; }
      }
      if (!filas) throw new Error("Este archivo no tiene la base del SIAF: no encontré una hoja con las columnas activ_obra_accinv y mto_pim.");

      const ix = {};
      filas.aoa[filas.h].forEach((c, i) => { const k = norm(c); if (k && !(k in ix)) ix[k] = i; });
      const tiene = c => c.toLowerCase() in ix;
      const col = (row, c) => (tiene(c) ? row[ix[c.toLowerCase()]] : null);
      const faltan = OBLIGATORIAS.filter(c => !tiene(c));
      if (faltan.length) throw new Error("Faltan columnas: " + faltan.join(", "));
      const avisos = [];
      const cPto = tiene("PRESUPUESTO INICIAL") ? "PRESUPUESTO INICIAL" : "mto_pia";
      if (cPto === "mto_pia") avisos.push("sin columna PRESUPUESTO INICIAL (se usa mto_pia)");
      if (!tiene("CONDICION 2")) avisos.push("sin columna CONDICION 2");

      REG = [];
      for (let i = filas.h + 1; i < filas.aoa.length; i++) {
        const row = filas.aoa[i];
        if (!row || col(row, "activ_obra_accinv") == null) continue;
        const r = vacio();
        r.pto = num(col(row, cPto)); r.pim = num(col(row, "mto_pim"));
        r.cert = num(col(row, "mto_certificado")); r.comp = num(col(row, "mto_compro_anual"));
        for (let m = 0; m < 12; m++) {
          const mm = String(m + 1).padStart(2, "0");
          r.compM[m] = num(col(row, "mto_at_comp_" + mm));
          r.devM[m] = num(col(row, "mto_devenga_" + mm));
        }
        r.dev = r.devM.reduce((a, b) => a + b, 0);
        r.c = {};
        for (const c of CAMPOS) r.c[c] = txt(col(row, c));
        REG.push(r);
      }
      if (!REG.length) throw new Error("La hoja no tiene filas de datos.");

      ARCHIVO = { nombre: file.name, fecha: new Date(file.lastModified), filas: REG.length, avisos };
      if (location.hash) history.replaceState(null, "", location.pathname + location.search);
      document.body.classList.remove("solo-lectura");
      prepararControles();
      estado("");
      mostrar(true);
    } catch (e) {
      console.error(e);
      estado(e.message || "No se pudo leer el archivo.", true);
    }
  }

  function fechaDelNombre(nombre) {
    const n = nombre.toLowerCase();
    let m = n.match(/(\d{1,2})\s*[-_ ]?\s*(enero|febrero|marzo|abril|mayo|junio|julio|agosto|setiembre|septiembre|octubre|noviembre|diciembre)\s*[-_ ]?\s*(\d{4})/);
    if (m) return new Date(+m[3], m[2] === "septiembre" ? 8 : MESES_LARGOS.indexOf(m[2]), +m[1]);
    m = n.match(/(\d{1,2})[.\-_/](\d{1,2})[.\-_/](\d{2,4})/);
    if (m) return new Date(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2] - 1, +m[1]);
    return null;
  }

  // ── Controles ─────────────────────────────────────────────────────────
  function opciones(sel, items, valor) {
    sel.innerHTML = items.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join("");
    if (valor != null && items.some(([v]) => v === valor)) sel.value = valor;
  }

  function prepararControles() {
    // siempre arranca con la vista del reporte original: actividad FEN, por meta y condición, sin filtros
    const acts = [...new Set(REG.map(r => r.c.activ_obra_accinv))].sort(natural);
    const def = acts.find(a => a.startsWith(ACT_FEN)) || acts[0];
    opciones($("selAct"), acts.map(a => [a, a]), def);
    opciones($("selN1"), NIVEL1.map(n => [n.id, n.t]), "sec_func");
    opciones($("selN2"), NIVEL2.map(n => [n.id, n.t]), "CONDICION 2");
    $("chkCeros").checked = false;
    prepararFiltros();
    // el Excel del SIAF descargado el día X trae datos con corte al día X-1
    if (ARCHIVO.corte) { $("fDatos").value = ARCHIVO.corte; return; }
    const f = fechaDelNombre(ARCHIVO.nombre) || ARCHIVO.fecha;
    const d = new Date(f); d.setDate(d.getDate() - 1);
    $("fDatos").value = isoDe(d);
  }

  function prepararFiltros() {
    const act = $("selAct").value;
    const sub = REG.filter(r => r.c.activ_obra_accinv === act);
    for (const { el, col } of FILTROS) {
      const vals = [...new Set(sub.map(r => r.c[col]).filter(Boolean))].sort(natural);
      opciones($(el), [["", "(Todas)"], ...vals.map(v => [v, v])], "");
      $(el).closest("[data-filtro]").hidden = vals.length < 2;
    }
  }

  function mostrar(reporte) {
    $("carga").hidden = reporte;
    $("escenario").hidden = $("acciones").hidden = $("archivoInfo").hidden = !reporte;
    if (!reporte) { $("opciones").hidden = true; $("btnOpciones").setAttribute("aria-expanded", "false"); }
    if (reporte) { escalar(); render(); }
  }

  // ── Cálculo ───────────────────────────────────────────────────────────
  function nombreGrupo(n1, clave, ejemplo) {
    if (n1 !== "sec_func") return clave || "(sin dato)";
    const nom = nombres[clave] || NOMBRES_DEFECTO[clave] || ejemplo.c.DEPENDENCIA || sinCodigo(ejemplo.c.finalidad) || "";
    return `${clave}.- ${nom.toUpperCase()}`;
  }

  function calcular() {
    const act = $("selAct").value, n1 = $("selN1").value, n2 = $("selN2").value;
    const filtros = FILTROS.map(f => [f.col, $(f.el).value]).filter(([, v]) => v);
    const verCeros = $("chkCeros").checked;
    const regs = REG.filter(r => r.c.activ_obra_accinv === act && filtros.every(([c, v]) => r.c[c] === v));

    const grupos = new Map();
    for (const r of regs) {
      const k1 = r.c[n1] || "(sin dato)", k2 = r.c[n2] || "(sin dato)";
      if (!grupos.has(k1)) grupos.set(k1, { clave: k1, ejemplo: r, tot: vacio(), hijos: new Map(), regs: [] });
      const g = grupos.get(k1);
      acumular(g.tot, r); g.regs.push(r);
      if (!g.hijos.has(k2)) g.hijos.set(k2, { clave: k2, nombre: k2, tot: vacio(), regs: [] });
      const h = g.hijos.get(k2);
      acumular(h.tot, r); h.regs.push(r);
    }
    const lista = [...grupos.values()]
      .filter(g => verCeros || !esCero(g.tot))
      .sort((a, b) => natural(a.clave, b.clave))
      .map(g => ({
        ...g,
        nombre: nombreGrupo(n1, g.clave, g.ejemplo),
        hijos: [...g.hijos.values()].filter(h => verCeros || !esCero(h.tot)).sort((a, b) => natural(a.clave, b.clave)),
      }));
    return { act, n1, n2, regs, grupos: lista, total: regs.reduce(acumular, vacio()) };
  }

  // ── Componentes de la lámina ──────────────────────────────────────────
  function dona(p, color) {
    // doughnut de PowerPoint: área 94 % del cuadro, agujero 58 %
    const R = 47, r = R * 0.58, rm = (R + r) / 2, ancho = R - r, L = 2 * Math.PI * rm;
    const v = Math.max(0, Math.min(1, p));
    const arco = v >= 0.9999
      ? `<circle cx="50" cy="50" r="${rm}" fill="none" stroke="${color}" stroke-width="${ancho}"/>`
      : v > 0 ? `<circle cx="50" cy="50" r="${rm}" fill="none" stroke="${color}" stroke-width="${ancho}" stroke-dasharray="${v * L} ${L}" transform="rotate(-90 50 50)"/>` : "";
    return `<svg viewBox="0 0 100 100" aria-hidden="true"><circle cx="50" cy="50" r="${rm}" fill="none" stroke="${C.linea}" stroke-width="${ancho}"/>${arco}</svg>`;
  }

  function kpi(tipo, color, val, lbl, sub, p) {
    const izq = tipo === "dona" ? dona(p, color)
      : `<div class="insignia" style="background:${color}"><img src="${(window.LOGOS && window.LOGOS.monedas) || "img/monedas.png"}" alt=""></div>`;
    return `<div class="kpi ${tipo === "dona" ? "con-dona" : ""}">${izq}<div class="kpi-txt">` +
      `<div class="kpi-val" style="color:${color}">${val}</div><div class="kpi-lbl">${lbl}</div><div class="kpi-sub">${sub}</div></div></div>`;
  }

  // achica la letra hasta que el texto quepa en una línea (máx. y mín. en px)
  function ajustar(el, max, min) {
    let s = max;
    el.style.fontSize = s + "px";
    while (el.scrollWidth > el.clientWidth && s > min) { s -= 0.5; el.style.fontSize = s + "px"; }
  }

  function celdas(t, etiqueta) {
    return COLS.map(c => c.a === "l"
      ? `<td class="l">${esc(etiqueta)}</td>`
      : `<td${c.ejec ? ' class="ejec"' : ""}>${c.f(t)}</td>`).join("");
  }

  // ── Render ────────────────────────────────────────────────────────────
  function render() {
    const V = VISTA = calcular(), T = V.total;
    const ej = REG.find(r => r.c.activ_obra_accinv === V.act) || REG[0];
    const esFen = V.act.startsWith(ACT_FEN);
    const codAct = V.act.split(".")[0];
    const ueNom = ej.c.unidad_ejecutora.replace(/^\s*\d+\.\s*/, "");
    const ueCod = (ej.c.unidad_ejecutora.match(/^\s*(\d+)/) || [, ""])[1];
    const corte = ddmmyyyy($("fDatos").value);
    const n1t = NIVEL1.find(n => n.id === V.n1).t.toLowerCase(), n2t = NIVEL2.find(n => n.id === V.n2).t.toLowerCase();
    const marco = esFen ? "EN EL MARCO DEL FENÓMENO EL NIÑO" : `EN ${sinCodigo(V.act).toUpperCase()}`;
    const marcoSinTilde = esFen ? "EN EL MARCO DEL FENOMENO EL NIÑO" : marco;

    // títulos tal cual del reporte original
    $("lTitulo").textContent = `PRESUPUESTO ASIGNADO ${marco} POR GENÉRICA DE GASTO`;
    ajustar($("lTitulo"), 24, 16);
    $("lSub").textContent = `${ueCod}-${ej.c.sec_ejec}: ${ueNom}`;
    $("lNota").textContent = `Fuente: SIAF - MEF · Corte al ${corte}`;
    $("archivoInfo").textContent = ARCHIVO.enlace
      ? `Reporte compartido por enlace · ${ARCHIVO.nombre}`
      : `${ARCHIVO.nombre} · ${ARCHIVO.filas.toLocaleString("es-PE")} filas` + (ARCHIVO.avisos.length ? " · " + ARCHIVO.avisos.join(" · ") : "");

    // franja de indicadores
    $("kpis").innerHTML =
      kpi("insignia", C.navy, soles(T.pim), "PIM", `S/ ${miles(T.pim)}`) +
      kpi("dona", C.ocre, fpct(pct(T.cert, T.pim)), "Certificado", soles(T.cert), pct(T.cert, T.pim)) +
      kpi("dona", C.coral, fpct(pct(T.comp, T.pim)), "Compromiso anual", soles(T.comp), pct(T.comp, T.pim)) +
      kpi("dona", C.rojo, fpct(pct(T.dev, T.pim)), "Devengado", soles(T.dev), pct(T.dev, T.pim));

    // tabla (misma lógica de tamaño de letra y alto de fila que el PPT)
    $("lTablaTit").textContent = esFen && V.n1 === "sec_func" && V.n2 === "CONDICION 2"
      ? `EJECUCION PRESUPUESTAL ${marcoSinTilde} POR ACTIVIDAD Y GENERICA DE GASTO`
      : `EJECUCION PRESUPUESTAL ${marcoSinTilde} POR ${n1t.toUpperCase()} Y ${n2t.toUpperCase()}`;
    ajustar($("lTablaTit"), 17, 12);
    const tot = COLS.reduce((s, c) => s + c.w, 0);
    $("cols").innerHTML = COLS.map(c => `<col style="width:${(c.w / tot * 100).toFixed(3)}%">`).join("");
    $("thead").innerHTML = `<tr>${COLS.map(c => `<th class="${c.a || ""}">${c.t}</th>`).join("")}</tr>`;

    let html = "", nFilas = 0;
    V.grupos.forEach((g, gi) => {
      html += `<tr class="n0" data-g="${gi}">${celdas(g.tot, g.nombre)}</tr>`; nFilas++;
      g.hijos.forEach((h, hi) => { html += `<tr class="n1" data-g="${gi}" data-h="${hi}">${celdas(h.tot, h.nombre)}</tr>`; nFilas++; });
    });
    html += `<tr class="total">${celdas(T, "Total")}</tr>`; nFilas++;
    $("tbody").innerHTML = html;

    // alto de fila y letra según cuántas filas hay (la cabecera cuenta 1.4 filas)
    const env = $("tablaEnv");
    const disp = env.clientHeight || 352;
    const fila = Math.max(11, Math.min(26, disp / (nFilas + 1.4)));
    env.style.setProperty("--fila", fila + "px");
    env.style.setProperty("--letra", Math.max(9, Math.min(14, fila * 0.64)) + "px");
    env.style.setProperty("--letra-cab", Math.max(9, Math.min(12.5, fila * 0.6)) + "px");
    $("thead").querySelector("tr").style.height = fila * 1.4 + "px";
  }

  // ── Escala de la lámina (diseñada a 1280×720) ─────────────────────────
  function escalar() {
    const m = $("marco");
    if (m.clientWidth) m.style.setProperty("--s", m.clientWidth / 1280);
  }

  // ── Detalle de una fila ───────────────────────────────────────────────
  function abrirDetalle(gi, hi) {
    const g = VISTA.grupos[gi], hj = hi == null ? null : g.hijos[hi];
    const obj = hj || g;
    $("mTitulo").textContent = hj ? `${g.nombre} · ${hj.nombre}` : g.nombre;
    const esp = new Map();
    for (const r of obj.regs) {
      const k = r.c.EEGG || "(sin específica)";
      if (!esp.has(k)) esp.set(k, vacio());
      acumular(esp.get(k), r);
    }
    const filas = [...esp.entries()].filter(([, v]) => !esCero(v)).sort((a, b) => natural(a[0], b[0]));
    const cel = t => `<td>${miles(t.pim)}</td><td>${dec(t.cert)}</td><td>${fpct(pct(t.cert, t.pim))}</td><td>${dec(t.comp)}</td><td>${fpct(pct(t.comp, t.pim))}</td><td>${dec(t.dev)}</td><td>${fpct(pct(t.dev, t.pim))}</td>`;
    let html = `<h4>Por específica de gasto</h4><div class="m-env"><table class="m-tabla"><thead><tr><th>ESPECÍFICA</th><th>PIM S/</th><th>CERT. S/</th><th>CERT. %</th><th>COMP. ANUAL S/</th><th>COMP. %</th><th>EJEC. S/</th><th>EJEC. %</th></tr></thead><tbody>` +
      filas.map(([k, v]) => `<tr><td>${esc(k)}</td>${cel(v)}</tr>`).join("") +
      `</tbody><tfoot><tr><td>Total</td>${cel(obj.tot)}</tr></tfoot></table></div>`;
    const t = obj.tot;
    html += `<h4>Compromiso y devengado por mes (S/)</h4>` + (Math.max(...t.compM, ...t.devM) > 0.5
      ? `<div class="m-env"><table class="m-tabla"><thead><tr><th></th>${MESES.map(m => `<th>${m}</th>`).join("")}</tr></thead><tbody>` +
        `<tr><td>Compromiso</td>${t.compM.map(v => `<td>${miles(v)}</td>`).join("")}</tr>` +
        `<tr><td>Devengado</td>${t.devM.map(v => `<td>${miles(v)}</td>`).join("")}</tr></tbody></table></div>`
      : `<p class="m-nota">Todavía no hay compromisos mensuales ni devengados registrados.</p>`);
    $("mCuerpo").innerHTML = html;
    $("modal").hidden = false;
  }

  // ── Exportar ──────────────────────────────────────────────────────────
  function nombreSalida(ext) {
    const act = VISTA.act.split(".")[0];
    return `Reporte_${act === ACT_FEN ? "FEN" : act}_corte_${$("fDatos").value || "sin_fecha"}.${ext}`;
  }

  function exportarExcel() {
    const V = VISTA;
    const X = [["PIM S/", t => t.pim, "#,##0"], ["CERT. S/", t => t.cert, "#,##0.00"],
      ["CERT. %", t => pct(t.cert, t.pim), "0.0%"], ["SALDO x CERT", t => t.pim - t.cert, "#,##0.00"],
      ["COMP. ANUAL S/", t => t.comp, "#,##0.00"], ["COMP. %", t => pct(t.comp, t.pim), "0.0%"],
      ["SALDO x COMP", t => t.pim - t.comp, "#,##0.00"], ["EJEC. S/", t => t.dev, "#,##0.00"],
      ["EJEC. %", t => pct(t.dev, t.pim), "0.0%"], ["SALDO x DEVENG", t => t.pim - t.dev, "#,##0.00"]];
    const fila = (etq, t) => [etq, ...X.map(c => c[1](t))];
    const aoa = [[$("lSub").textContent], [$("lTablaTit").textContent], [$("lNota").textContent], [], ["DETALLE / CONDICIÓN", ...X.map(c => c[0])]];
    for (const g of V.grupos) {
      aoa.push(fila(g.nombre, g.tot));
      for (const h of g.hijos) aoa.push(fila("    " + h.nombre, h.tot));
    }
    aoa.push(fila("Total", V.total));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (let r = 5; r < aoa.length; r++) X.forEach((c, j) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c: j + 1 })];
      if (cell) cell.z = c[2];
    });
    ws["!cols"] = [{ wch: 48 }, ...X.map(() => ({ wch: 15 }))];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reporte");
    XLSX.writeFile(wb, nombreSalida("xlsx"));
  }

  async function exportarPng() {
    document.body.classList.add("capturando");
    try {
      const canvas = await html2canvas($("lamina"), { scale: 2, width: 1280, height: 720, backgroundColor: "#F0F3F7" });
      const a = document.createElement("a");
      a.download = nombreSalida("png");
      a.href = canvas.toDataURL("image/png");
      a.click();
    } catch (e) {
      console.error(e);
      $("archivoInfo").textContent = "No se pudo generar la imagen: " + e.message;
    } finally {
      document.body.classList.remove("capturando");
    }
  }

  // ── Enlace para compartir ─────────────────────────────────────────────
  // Las filas de la actividad elegida viajan comprimidas en el "#" del enlace.
  // Esa parte no se envía a ningún servidor: solo la lee el navegador de quien abre el enlace.
  const aB64 = bytes => { let s = ""; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""); };
  const deB64 = s => { const b = atob(s.replace(/-/g, "+").replace(/_/g, "/")); const u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; };
  async function comprimir(texto) {
    const st = new Blob([texto]).stream().pipeThrough(new CompressionStream("deflate-raw"));
    return aB64(new Uint8Array(await new Response(st).arrayBuffer()));
  }
  async function descomprimir(s) {
    const st = new Blob([deB64(s)]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Response(st).text();
  }
  const r2 = v => Math.round(v * 100) / 100;

  async function copiarEnlace() {
    // solo viajan las filas del reporte que se está viendo (actividad y filtros actuales)
    const datos = {
      v: 1, a: ARCHIVO.nombre, corte: $("fDatos").value, n1: $("selN1").value, n2: $("selN2").value,
      z: $("chkCeros").checked, k: CAMPOS,
      r: VISTA.regs.map(r => [...CAMPOS.map(k => r.c[k]), r2(r.pto), r2(r.pim), r2(r.cert), r2(r.comp), ...r.compM.map(r2), ...r.devM.map(r2)]),
    };
    const url = location.href.split("#")[0] + "#r=" + await comprimir(JSON.stringify(datos));
    let ok = false;
    try { await navigator.clipboard.writeText(url); ok = true; } catch {
      const ta = Object.assign(document.createElement("textarea"), { value: url });
      document.body.appendChild(ta); ta.select();
      try { ok = document.execCommand("copy"); } catch { /* sin portapapeles */ }
      ta.remove();
    }
    const b = $("btnEnlace");
    b.textContent = ok ? "Enlace copiado ✓" : "No se pudo copiar";
    setTimeout(() => { b.textContent = "Copiar enlace"; }, 2500);
    if (location.protocol === "file:")
      $("archivoInfo").textContent = "Ojo: abriste la página desde tu PC; el enlace solo servirá a otros cuando la página esté publicada.";
  }

  async function abrirEnlace() {
    const m = location.hash.match(/^#r=([\w-]+)$/);
    if (!m) return;
    try {
      const d = JSON.parse(await descomprimir(m[1]));
      const nk = d.k.length;
      REG = d.r.map(row => {
        const r = vacio();
        r.c = {};
        d.k.forEach((k, i) => { r.c[k] = row[i]; });
        [r.pto, r.pim, r.cert, r.comp] = row.slice(nk, nk + 4);
        r.compM = row.slice(nk + 4, nk + 16);
        r.devM = row.slice(nk + 16, nk + 28);
        r.dev = r.devM.reduce((a, b) => a + b, 0);
        return r;
      });
      ARCHIVO = { nombre: d.a, fecha: new Date(), filas: REG.length, avisos: [], corte: d.corte, enlace: true };
      prepararControles();
      $("selN1").value = d.n1; $("selN2").value = d.n2;
      $("chkCeros").checked = !!d.z;
      document.body.classList.add("solo-lectura");     // el enlace muestra solo el reporte, sin opciones
      mostrar(true);
    } catch (e) {
      console.error(e);
      estado("El enlace está incompleto o dañado. Pide que te lo vuelvan a enviar o sube el Excel.", true);
    }
  }

  // ── Eventos ───────────────────────────────────────────────────────────
  const zona = $("zona");
  $("archivo").addEventListener("change", e => { const f = e.target.files[0]; if (f) cargar(f); e.target.value = ""; });
  ["dragenter", "dragover"].forEach(ev => zona.addEventListener(ev, e => { e.preventDefault(); zona.classList.add("sobre"); }));
  ["dragleave", "drop"].forEach(ev => zona.addEventListener(ev, () => zona.classList.remove("sobre")));
  window.addEventListener("dragover", e => e.preventDefault());
  window.addEventListener("drop", e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) cargar(f); });

  $("selAct").addEventListener("change", () => { prepararFiltros(); render(); });
  ["selN1", "selN2", "fFuente", "fUnidad", "fDepto", "fDatos", "chkCeros"].forEach(id => $(id).addEventListener("change", render));

  $("tbody").addEventListener("click", e => {
    const tr = e.target.closest("tr");
    if (tr && tr.dataset.g != null) abrirDetalle(+tr.dataset.g, tr.dataset.h == null ? null : +tr.dataset.h);
  });
  $("mCerrar").addEventListener("click", () => { $("modal").hidden = true; });
  $("modal").addEventListener("click", e => { if (e.target === $("modal")) $("modal").hidden = true; });
  document.addEventListener("keydown", e => { if (e.key === "Escape") $("modal").hidden = true; });

  $("btnOpciones").addEventListener("click", () => {
    const abrir = $("opciones").hidden;
    $("opciones").hidden = !abrir;
    $("btnOpciones").setAttribute("aria-expanded", String(abrir));
    escalar();
  });
  $("btnOtro").addEventListener("click", () => mostrar(false));
  $("btnXlsx").addEventListener("click", exportarExcel);
  $("btnPng").addEventListener("click", exportarPng);
  $("btnPdf").addEventListener("click", () => window.print());
  $("btnEnlace").addEventListener("click", copiarEnlace);
  $("btnReset").addEventListener("click", () => { prepararControles(); render(); });
  // borra las preferencias que guardaban versiones anteriores (actividad / agrupación)
  ["reporteFen.act", "reporteFen.n1", "reporteFen.n2"].forEach(k => { try { localStorage.removeItem(k); } catch { /* nada */ } });
  window.addEventListener("hashchange", abrirEnlace);
  $("btnPresentar").addEventListener("click", () => {
    const el = $("escenario");
    (el.requestFullscreen || el.webkitRequestFullscreen).call(el);
  });
  const reajustar = () => { if (!$("escenario").hidden) escalar(); };
  window.addEventListener("resize", reajustar);
  document.addEventListener("fullscreenchange", reajustar);

  // Imágenes embebidas: permiten generar la imagen PNG también al abrir el archivo local
  if (window.LOGOS) document.querySelectorAll("img").forEach(img => {
    const k = (img.getAttribute("src") || "").replace(/^img\/|\.png$/g, "");
    if (window.LOGOS[k]) img.src = window.LOGOS[k];
  });

  abrirEnlace();
})();
