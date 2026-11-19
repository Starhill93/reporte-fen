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
    if (a >= 1e6) return `S/ ${fmt(v / 1e6, 2).replace(/\.?0+$/, "")} mill`;   // 10.90 → 10.9
    if (a >= 1e5) return `S/ ${fmt(v / 1e3, 0)} mil`;
    if (a >= 1e3) return `S/ ${fmt(v / 1e3, 1)} mil`;
    return `S/ ${fmt(v, 0)}`;
  };
  const txt = v => (v == null ? "" : String(v).trim());
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const natural = (a, b) => a.localeCompare(b, "es", { numeric: true, sensitivity: "base" });
  const sinCodigo = s => s.replace(/^[\d.]+\s*\.?\s*/, "");
  const p2 = n => String(n).padStart(2, "0");
  // fecha y hora de proceso como "AAAA-MM-DDTHH:MM" (valor del control datetime-local)
  const marcaDe = d => `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}T${p2(d.getHours())}:${p2(d.getMinutes())}`;
  const partes = v => {
    const [f, h = "00:00"] = String(v || "").split("T");
    const [y, m, d] = f.split("-");
    const [hh, mm] = h.split(":");
    return { y, m, d, hh, mm };
  };
  // "18/11/2026 - 19:30 p.m."
  const fechaHora = v => { if (!v) return ""; const p = partes(v); return `${p.d}/${p.m}/${p.y} - ${p.hh}:${p.mm} ${+p.hh < 12 ? "a.m." : "p.m."}`; };
  // "18.11.2026-19.30"
  const fechaArchivo = v => { if (!v) return "sin_fecha"; const p = partes(v); return `${p.d}.${p.m}.${p.y}-${p.hh}.${p.mm}`; };
  const oracion = s => (s ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : s);
  const vacio = () => ({ pto: 0, pim: 0, cert: 0, comp: 0, dev: 0, compM: Array(12).fill(0), devM: Array(12).fill(0) });
  const acumular = (a, r) => {
    a.pto += r.pto; a.pim += r.pim; a.cert += r.cert; a.comp += r.comp; a.dev += r.dev;
    for (let i = 0; i < 12; i++) { a.compM[i] += r.compM[i]; a.devM[i] += r.devM[i]; }
    return a;
  };
  const esCero = r => [r.pto, r.pim, r.cert, r.comp, r.dev].every(v => Math.abs(v) < 0.005);
  const estado = (m, error = false) => { const e = $("estado"); e.textContent = m; e.classList.toggle("error", error); };

  // Columnas de la tabla: concepto, PIM, certificado, compromiso y devengado con sus %
  const COLS = [
    { t: "CONCEPTO", w: 2.65, a: "l" },
    { t: "PIM", w: 1.3, f: t => miles(t.pim) },
    { t: "CERTIFICADO", w: 1.7, f: t => miles(t.cert) },
    { t: "%", w: 1.2, p: t => pct(t.cert, t.pim) },
    { t: "COMPROMISO", w: 1.7, f: t => miles(t.comp) },
    { t: "%", w: 1.1, p: t => pct(t.comp, t.pim) },
    { t: "DEVENGADO", w: 1.4, f: t => miles(t.dev) },
    { t: "%", w: 0.85, p: t => pct(t.dev, t.pim) },
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

      // la fecha y hora del reporte es el momento en que se lee y procesa el archivo
      ARCHIVO = { nombre: file.name, procesado: marcaDe(new Date()), filas: REG.length, avisos };
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
    // fecha y hora = momento en que se procesó el archivo (o la que trae el enlace compartido)
    $("fDatos").value = ARCHIVO.procesado;
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
    if (ejemplo.c._nombre) return ejemplo.c._nombre;          // viene de un enlace compartido
    if (n1 !== "sec_func") return clave || "(sin dato)";
    const nom = nombres[clave] || NOMBRES_DEFECTO[clave] || ejemplo.c.DEPENDENCIA || sinCodigo(ejemplo.c.finalidad) || clave;
    return oracion(nom);                                       // "Mantenimiento e infraestructura", sin número de meta
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
  function kpi(val, lbl, sub) {
    return `<div class="kpi"><div class="kpi-val">${val}</div><div class="kpi-lbl">${lbl}</div><div class="kpi-sub">${sub}</div></div>`;
  }

  // achica la letra hasta que el texto quepa en una línea (máx. y mín. en px)
  function ajustar(el, max, min) {
    let s = max;
    el.style.fontSize = s + "px";
    while (el.scrollWidth > el.clientWidth && s > min) { s -= 0.5; el.style.fontSize = s + "px"; }
  }

  function celdas(t, etiqueta) {
    return COLS.map(c => {
      if (c.a === "l") return `<div class="c l">${esc(etiqueta)}</div>`;
      if (c.p) return `<div class="c p">${(c.p(t) * 100).toFixed(1)}</div>`;
      return `<div class="c">${c.f(t)}</div>`;
    }).join("");
  }

  // ── Render ────────────────────────────────────────────────────────────
  function render() {
    const V = VISTA = calcular(), T = V.total;
    const esFen = V.act.startsWith(ACT_FEN);
    const cuando = fechaHora($("fDatos").value);

    const base = esFen ? "EJECUCIÓN DE RECURSOS FEN" : `EJECUCIÓN DE RECURSOS · ${sinCodigo(V.act).toUpperCase()}`;
    $("lTitulo").textContent = cuando ? `${base} – ${cuando}` : base;
    ajustar($("lTitulo"), 30, 16);
    $("lNota").textContent = "Fuente: SIAF - MEF";
    $("archivoInfo").textContent = ARCHIVO.enlace
      ? `Reporte compartido por enlace · ${cuando}`
      : `${ARCHIVO.nombre} · ${ARCHIVO.filas.toLocaleString("es-PE")} filas · procesado ${cuando}` + (ARCHIVO.avisos.length ? " · " + ARCHIVO.avisos.join(" · ") : "");

    // franja de indicadores
    $("kpis").innerHTML =
      kpi(soles(T.pim), "PIM", `S/ ${miles(T.pim)}`) +
      kpi(fpct(pct(T.cert, T.pim)), "Certificado", soles(T.cert)) +
      kpi(fpct(pct(T.comp, T.pim)), "Compromiso", soles(T.comp)) +
      kpi(fpct(pct(T.dev, T.pim)), "Devengado", soles(T.dev));

    // tabla: solo los conceptos (sin detalle) y el total
    $("tablaEnv").style.setProperty("--cols", COLS.map(c => c.w + "fr").join(" "));
    $("thead").innerHTML = `<div class="trow cab">${COLS.map(c => `<div>${c.t}</div>`).join("")}</div>`;
    $("tbody").innerHTML = V.grupos.map((g, gi) => `<div class="trow" data-g="${gi}">${celdas(g.tot, g.nombre)}</div>`).join("") +
      `<div class="trow total">${celdas(T, "TOTAL")}</div>`;

    // si hay muchas filas se achica todo (letra, alto y márgenes) hasta que la tabla y la fuente quepan en la lámina
    const card = $("tablaCard"), LIMITE = 705;               // px de la lámina
    let k = 1;
    card.style.setProperty("--k", k);
    while (card.offsetTop + card.offsetHeight > LIMITE && k > 0.45) { k -= 0.05; card.style.setProperty("--k", k.toFixed(2)); }

    prepararEnlace();
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
    if (ARCHIVO.enlace) {   // el enlace solo trae la tabla; el detalle sale del Excel
      $("mCuerpo").innerHTML = `<p class="m-nota">El detalle por específica de gasto y por mes está disponible al cargar el Excel del SIAF en la página.</p>`;
      $("modal").hidden = false;
      return;
    }
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
    return `Reporte_financiero_${act === ACT_FEN ? "FEN" : act}_${fechaArchivo($("fDatos").value)}.${ext}`;
  }

  function exportarExcel() {
    const V = VISTA;
    const X = [["PIM", t => t.pim, "#,##0"], ["CERTIFICADO", t => t.cert, "#,##0"],
      ["%", t => pct(t.cert, t.pim), "0.0%"], ["COMPROMISO", t => t.comp, "#,##0"],
      ["%", t => pct(t.comp, t.pim), "0.0%"], ["DEVENGADO", t => t.dev, "#,##0"],
      ["%", t => pct(t.dev, t.pim), "0.0%"]];
    const fila = (etq, t) => [etq, ...X.map(c => c[1](t))];
    const aoa = [[$("lTitulo").textContent], [$("lNota").textContent], [], [], ["CONCEPTO", ...X.map(c => c[0])]];
    for (const g of V.grupos) aoa.push(fila(g.nombre, g.tot));
    aoa.push(fila("TOTAL", V.total));
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
      const canvas = await html2canvas($("lamina"), { scale: 2, width: 1280, height: 720, backgroundColor: "#FFFFFF" });
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
  // En el "#" del enlace viaja solo la tabla que se ve (grupos, filas y sus montos en céntimos),
  // comprimida con lz-string (funciona en cualquier navegador, también el de WhatsApp).
  // Esa parte del enlace no se envía a ningún servidor: solo la lee el navegador de quien lo abre.
  const comprimir = t => LZString.compressToBase64(t).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  const descomprimir = s => {
    s = s.replace(/[^\w-]/g, "").replace(/-/g, "+").replace(/_/g, "/");
    while (s.length % 4) s += "=";
    return LZString.decompressFromBase64(s);
  };
  const cent = v => Math.round(v * 100);

  // El enlace se arma apenas cambia el reporte, para que al tocar el botón se copie/comparta
  // en el mismo instante (los celulares solo lo permiten durante el toque).
  let ENLACE = null, enlaceVer = 0;
  const esMovil = () => window.matchMedia("(pointer: coarse)").matches;
  const puedeCompartir = () => !!navigator.share && esMovil();
  const textoBoton = () => (puedeCompartir() ? "Compartir enlace" : "Copiar enlace");

  function prepararEnlace() {
    if (!VISTA || document.body.classList.contains("solo-lectura")) return;
    const ej = VISTA.regs[0];
    if (!ej) { ENLACE = null; return; }
    // [act, ue, sec_ejec, corte, n1, [[clave, nombre, pim, cert, comp, dev], …]]  (montos en céntimos)
    const datos = [VISTA.act, ej.c.unidad_ejecutora, ej.c.sec_ejec, $("fDatos").value, VISTA.n1,
      VISTA.grupos.map(g => [g.clave, g.nombre, cent(g.tot.pim), cent(g.tot.cert), cent(g.tot.comp), cent(g.tot.dev)])];
    ENLACE = location.href.split("#")[0].split("?")[0] + "#e=" + comprimir(JSON.stringify(datos));
  }

  function avisoBoton(t) {
    const b = $("btnEnlace");
    b.textContent = t;
    setTimeout(() => { b.textContent = textoBoton(); }, 2500);
  }

  function mostrarEnlaceManual(url) {
    $("mTitulo").textContent = "Enlace del reporte";
    $("mCuerpo").innerHTML = `<p class="m-nota">Copia este enlace y envíalo (mantén presionado si estás en el celular):</p>
      <textarea class="enlace-txt" readonly rows="4">${esc(url)}</textarea>`;
    $("modal").hidden = false;
    const ta = $("mCuerpo").querySelector("textarea");
    ta.focus(); ta.select();
  }

  function copiarEnlace() {
    if (!ENLACE) prepararEnlace();
    const url = ENLACE;
    if (!url) { avisoBoton("Sin datos"); return; }
    if (location.protocol === "file:")
      $("archivoInfo").textContent = "Ojo: abriste la página desde tu PC; el enlace solo servirá a otros cuando la página esté publicada.";
    if (puedeCompartir()) {
      navigator.share({ title: "Reporte FEN · INABIF", text: $("lTitulo").textContent, url })
        .catch(e => { if (e && e.name !== "AbortError") mostrarEnlaceManual(url); });
      return;
    }
    const manual = () => {
      const ta = Object.assign(document.createElement("textarea"), { value: url });
      ta.setAttribute("readonly", ""); ta.style.position = "fixed"; ta.style.opacity = "0";
      document.body.appendChild(ta); ta.select(); ta.setSelectionRange(0, url.length);
      let ok = false;
      try { ok = document.execCommand("copy"); } catch { /* sin portapapeles */ }
      ta.remove();
      if (ok) avisoBoton("Enlace copiado ✓"); else mostrarEnlaceManual(url);
    };
    if (navigator.clipboard && window.isSecureContext)
      navigator.clipboard.writeText(url).then(() => avisoBoton("Enlace copiado ✓"), manual);
    else manual();
  }

  function abrirEnlace() {
    const h = decodeURIComponent(location.hash || "");
    if (/^#[rd]=/.test(h)) { estado("Este enlace es de una versión anterior de la página. Pide uno nuevo o sube el Excel.", true); return; }
    const m = h.match(/^#e=(.+)$/);
    if (!m) return;
    try {
      const [act, ue, sec, corte, n1, grupos] = JSON.parse(descomprimir(m[1]));
      REG = grupos.map(([k1, nombre, pim, cert, comp, dev]) => {
        const r = vacio();
        r.pim = pim / 100; r.cert = cert / 100; r.comp = comp / 100; r.dev = dev / 100;
        r.c = { activ_obra_accinv: act, unidad_ejecutora: ue, sec_ejec: sec, [n1]: k1, _nombre: nombre };
        return r;
      });
      // enlaces anteriores traían solo la fecha (AAAA-MM-DD): se muestran a las 00:00
      const procesado = corte && !corte.includes("T") ? corte + "T00:00" : corte;
      ARCHIVO = { nombre: "", filas: REG.length, avisos: [], procesado, enlace: true };
      prepararControles();
      $("selN1").value = n1;
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
    const tr = e.target.closest(".trow");
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
  $("btnEnlace").textContent = textoBoton();
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
