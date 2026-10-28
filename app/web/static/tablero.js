// Tablero de conciliación SRI vs Odoo (Compras / Ventas). JS simple, sin dependencias.
//
// Flujo por pestaña: el primer TXT crea la sesión (POST /sesion); cada TXT se sube una vez
// (POST /sesion/{id}/txt, en lotes); tras subir o quitar se espera 700 ms sin cambios y se pide
// UN recálculo (POST /sesion/{id}/conciliar -> generacion). Luego se sondea /status cada 2 s y,
// al terminar, se lee /tablero: si su `generacion` es anterior a la última pedida se ignora.
// Mientras tanto se conserva el tablero anterior en pantalla.
// Todo texto dinámico se escapa con esc() antes de entrar al HTML.
"use strict";
(() => {
  const LIMITE = { compras: 5, ventas: 400 };
  const MAX_BYTES_TXT = 20 * 1024 * 1024;          // igual que el servidor
  const LOTE_ARCHIVOS = 20, LOTE_BYTES = 40 * 1024 * 1024;
  const DEBOUNCE_MS = 700, SONDEO_MS = 2000, MAX_FILAS = 500;
  const COLUMNAS = [
    ["tipo", "Tipo"], ["estado", "Estado"], ["documento", "Documento"], ["fecha", "Fecha"],
    ["contraparte", "Contraparte"], ["ruc", "RUC"], ["valor_sri", "Valor SRI", "num"],
    ["valor_odoo", "Valor Odoo", "num"], ["diferencia", "Diferencia", "num"],
    ["detalle", "Detalle"], ["que_corregir", "Qué corregir en Odoo"],
  ];

  function esc(t) {
    return String(t ?? "").replace(/[&<>"']/g,
      c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  const $ = id => document.getElementById(id);
  const esInfo = estado => String(estado || "").startsWith("Informativo");

  // ---------------------------------------------------------------- API
  class NoAutenticado extends Error {}

  async function api(url, opciones = {}) {
    let r;
    try {
      r = await fetch(url, opciones);
    } catch (_) {
      throw new Error("No se pudo contactar al servidor. Revisa la conexión e inténtalo de nuevo.");
    }
    if (r.status === 401) { location = "/login"; throw new NoAutenticado(); }
    if (!r.ok) {
      let detalle = "";
      try { const j = await r.json(); if (typeof j.detail === "string") detalle = j.detail; } catch (_) { /* sin JSON */ }
      const err = new Error(detalle || `Error del servidor (HTTP ${r.status})`);
      err.status = r.status;
      throw err;
    }
    return r;
  }

  // ---------------------------------------------------------------- estado por pestaña
  function nuevoEstado(modo) {
    return {
      modo, sid: null, creando: null, archivos: [], cola: [], subiendo: false, cambios: false,
      tablero: null, genPedida: 0, actualizando: false, mensaje: "", progreso: null,
      errorCalculo: null, errores: [], tDebounce: null, tSondeo: null,
      f: { tipo: "", estado: "", q: "", info: false, orden: "", desc: false },
    };
  }
  const estados = { compras: nuevoEstado("compras"), ventas: nuevoEstado("ventas") };
  let modo = "compras";
  const actual = () => estados[modo];
  const vigente = st => estados[st.modo] === st;  // "Nueva conciliación" reemplaza el objeto

  function avisar(st, msg) {
    st.errores.push(msg);
    if (st.errores.length > 6) st.errores.shift();
    render();
  }

  // Igual que el servidor: sin carpeta y extensión .txt en minúscula.
  function normalizar(nombre) {
    const base = String(nombre).split(/[\\/]/).pop();
    return /\.txt$/i.test(base) ? base.slice(0, -4) + ".txt" : base;
  }
  const buscar = (st, nombre) => st.archivos.find(a => a.nombre === nombre);

  function poner(st, entrada) {
    const i = st.archivos.findIndex(a => a.nombre === entrada.nombre);
    if (i < 0) st.archivos.push(entrada); else st.archivos[i] = { ...st.archivos[i], ...entrada };
  }

  // ---------------------------------------------------------------- carga de TXT
  function agregar(lista) {
    const st = actual();
    const ocupados = () => st.archivos.filter(a => a.enServidor || a.estado === "cola" || a.estado === "subiendo").length;
    for (const f of lista) {
      if (!/\.txt$/i.test(f.name)) { avisar(st, `«${f.name}» no se cargó: solo se aceptan archivos .txt.`); continue; }
      const nombre = normalizar(f.name);
      const previo = buscar(st, nombre);
      if (previo && (previo.estado === "subiendo" || previo.estado === "quitando")) {
        avisar(st, `«${nombre}» se está procesando; vuelve a cargarlo en unos segundos.`);
        continue;
      }
      st.cola = st.cola.filter(x => x.nombre !== nombre);
      const base = { nombre, filas: 0, tipos: {}, enServidor: previo ? previo.enServidor : false };
      if (f.size > MAX_BYTES_TXT) {
        poner(st, { ...base, estado: "error", error: "Archivo demasiado grande (máximo 20 MB)." });
      } else if (!previo?.enServidor && previo?.estado !== "cola" && ocupados() >= LIMITE[st.modo]) {
        poner(st, { ...base, estado: "error", error: `Se admiten como máximo ${LIMITE[st.modo]} TXT en ${st.modo}.` });
      } else {
        poner(st, { ...base, estado: "cola", error: null });
        st.cola.push({ nombre, file: f });
      }
    }
    render();
    procesarCola(st);
  }

  async function asegurarSesion(st) {
    if (st.sid) return true;
    st.creando = st.creando || api("/sesion", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ modo: st.modo }),
    }).then(r => r.json());
    try {
      const j = await st.creando;
      if (vigente(st)) st.sid = j.sesion_id;
      return vigente(st);
    } catch (e) {
      st.creando = null;
      if (e instanceof NoAutenticado) return false;
      for (const x of st.cola) poner(st, { nombre: x.nombre, estado: "error", error: e.message });
      st.cola = [];
      avisar(st, `No se pudo iniciar la conciliación: ${e.message}`);
      return false;
    }
  }

  function tomarLote(st) {
    const lote = [];
    let bytes = 0;
    while (st.cola.length && lote.length < LOTE_ARCHIVOS
           && (!lote.length || bytes + st.cola[0].file.size <= LOTE_BYTES)) {
      const x = st.cola.shift();
      bytes += x.file.size;
      lote.push(x);
    }
    return lote;
  }

  async function procesarCola(st) {
    if (st.subiendo) return;  // la cola en curso recogerá lo nuevo
    st.subiendo = true;
    try {
      while (st.cola.length && vigente(st)) {
        if (!(await asegurarSesion(st))) return;
        const lote = tomarLote(st);
        lote.forEach(x => poner(st, { nombre: x.nombre, estado: "subiendo" }));
        render();
        const fd = new FormData();
        lote.forEach(x => fd.append("archivos", x.file, x.nombre));
        try {
          const j = await (await api(`/sesion/${st.sid}/txt`, { method: "POST", body: fd })).json();
          if (!vigente(st)) return;
          for (const a of j.archivos) poner(st, { ...a, enServidor: true });
          st.cambios = true;
        } catch (e) {
          if (e instanceof NoAutenticado || !vigente(st)) return;
          for (const x of lote) poner(st, { nombre: x.nombre, estado: "error", error: e.message, filas: 0, tipos: {} });
          if (e.status === 404) avisar(st, "La sesión expiró. Pulsa «Nueva conciliación» y vuelve a subir los TXT.");
        }
        render();
      }
    } finally {
      st.subiendo = false;
    }
    if (vigente(st) && st.cambios) programarConciliar(st);
    render();
  }

  async function quitar(st, nombre) {
    const e = buscar(st, nombre);
    if (!e || e.estado === "subiendo" || e.estado === "quitando") return;
    st.cola = st.cola.filter(x => x.nombre !== nombre);
    if (!e.enServidor || !st.sid) {  // nunca llegó al servidor: solo se quita de la lista
      st.archivos = st.archivos.filter(a => a !== e);
      render();
      return;
    }
    const antes = e.estado;
    e.estado = "quitando";
    render();
    try {
      await api(`/sesion/${st.sid}/txt/${encodeURIComponent(nombre)}`, { method: "DELETE" });
      if (!vigente(st)) return;
      st.archivos = st.archivos.filter(a => a.nombre !== nombre);
      st.cambios = true;
      programarConciliar(st);
    } catch (err) {
      if (err instanceof NoAutenticado || !vigente(st)) return;
      e.estado = antes;
      avisar(st, `No se pudo quitar «${nombre}»: ${err.message}`);
    }
    render();
  }

  // ---------------------------------------------------------------- recálculo y sondeo
  function programarConciliar(st) {
    clearTimeout(st.tDebounce);
    st.tDebounce = setTimeout(() => conciliar(st), DEBOUNCE_MS);
  }

  async function conciliar(st) {
    st.tDebounce = null;
    if (!vigente(st) || !st.sid) return;
    if (st.subiendo || st.cola.length) return;  // al vaciarse, la cola vuelve a programarlo
    st.cambios = false;
    try {
      const j = await (await api(`/sesion/${st.sid}/conciliar`, { method: "POST" })).json();
      if (!vigente(st)) return;
      st.genPedida = Math.max(st.genPedida, j.generacion);
      st.actualizando = true;
      st.errorCalculo = null;
      st.mensaje = "";
      st.progreso = null;
      sondear(st, 0);
    } catch (e) {
      if (e instanceof NoAutenticado || !vigente(st)) return;
      if (e.status === 404) {
        st.cambios = false;  // la sesión ya no existe: reintentar no sirve
        avisar(st, "La sesión expiró. Pulsa «Nueva conciliación» y vuelve a subir los TXT.");
      } else {
        st.cambios = true;   // el banner «Resultado desactualizado» ofrece reintentar
        avisar(st, `No se pudo recalcular: ${e.message}`);
      }
    }
    render();
  }

  function sondear(st, espera = SONDEO_MS) {
    clearTimeout(st.tSondeo);
    st.tSondeo = setTimeout(() => tick(st), espera);
  }

  async function tick(st) {
    st.tSondeo = null;
    if (!vigente(st) || !st.actualizando) return;
    const pedida = st.genPedida;
    try {
      const s = await (await api(`/status/${st.sid}`)).json();
      if (!vigente(st)) return;
      if (s.estado === "error" && pedida === st.genPedida) {
        st.actualizando = false;
        st.errorCalculo = s.mensaje || "Error desconocido";
      } else if (s.estado === "done") {
        if (await cargarTablero(st)) st.actualizando = false;
      } else {
        st.mensaje = s.mensaje || "";
        st.progreso = typeof s.progreso === "number" ? s.progreso : null;
      }
    } catch (e) {
      if (e instanceof NoAutenticado || !vigente(st)) return;
      if (e.status === 404) {
        st.actualizando = false;
        st.errorCalculo = "La sesión expiró o ya no existe. Pulsa «Nueva conciliación» y vuelve a subir los TXT.";
      }  // otros fallos (red): se reintenta en el siguiente sondeo
    }
    render();
    if (vigente(st) && st.actualizando) sondear(st);
  }

  async function cargarTablero(st) {
    let t;
    try {
      t = await (await api(`/tablero/${st.sid}`)).json();
    } catch (e) {
      if (e.status === 404) return false;  // aún no hay un primer resultado
      throw e;
    }
    if (!vigente(st) || (t.generacion || 0) < st.genPedida) return false;  // generación superada: se ignora
    st.tablero = t;
    st.errorCalculo = null;
    return true;
  }

  // ---------------------------------------------------------------- descargas
  async function descargar(tipo) {
    const st = actual();
    if (!st.sid) return;
    try {
      const r = await api(`/download/${st.sid}?tipo=${tipo}`);
      const blob = await r.blob();
      const m = (r.headers.get("content-disposition") || "").match(/filename="([^"]+)"/);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = m ? m[1] : `${tipo}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 30000);
    } catch (e) {
      if (e instanceof NoAutenticado) return;
      avisar(st, e.status === 404 ? "El Excel aún no está disponible para esta conciliación." : `No se pudo descargar: ${e.message}`);
    }
  }

  // ---------------------------------------------------------------- formato
  const fmtNum = v => v == null ? "" : Number(v).toLocaleString("es-EC", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtFecha = v => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(v || ""); return m ? `${m[3]}/${m[2]}/${m[1]}` : (v || ""); };
  const plural = (n, uno, varios) => `${n} ${n === 1 ? uno : varios}`;

  function claseEstado(e) {
    if (e === "Coincide") return "ok";
    if (esInfo(e) || e === "XML no disponible") return "";
    if (e === "Solo en SRI") return "naranja";
    if (e === "Solo en Odoo" || e === "Odoo anulado, SRI autorizado") return "rojo";
    if (e.startsWith("Diferencia")) return "amar";
    if (e === "Factura incorrecta en Odoo" || e === "Duplicado en Odoo") return "viol";
    return "azul";
  }
  const insignia = e => `<span class="est ${claseEstado(String(e))}">${esc(e)}</span>`;

  // ---------------------------------------------------------------- pintado
  function render() {
    const st = actual();
    for (const m of ["compras", "ventas"]) {
      $("tab-" + m).classList.toggle("active", m === modo);
      $("tab-" + m).setAttribute("aria-selected", String(m === modo));
      $("hint-" + m).hidden = m !== modo;
    }
    pintarArchivos(st);
    pintarAvisos(st);
    pintarResultado(st);
    pintarExcepciones(st);
  }

  function pintarArchivos(st) {
    const etiquetas = { cola: ["En cola", ""], subiendo: ["Subiendo…", "azul"], quitando: ["Quitando…", ""] };
    const ok = st.archivos.filter(a => a.estado === "ok").length;
    const err = st.archivos.filter(a => a.estado === "error").length;
    $("archivos-cuenta").textContent = st.archivos.length
      ? `${plural(st.archivos.length, "archivo", "archivos")} · ${plural(ok, "listo", "listos")}${err ? ` · ${err} con error` : ""} (máx. ${LIMITE[st.modo]})`
      : `Máximo ${LIMITE[st.modo]} TXT`;
    if (!st.archivos.length) {
      $("lista-archivos").innerHTML = '<li class="vacia">Aún no hay archivos. Sube uno o varios TXT para empezar.</li>';
      return;
    }
    $("lista-archivos").innerHTML = st.archivos.map(a => {
      let est;
      if (a.estado === "ok") {
        const tipos = Object.entries(a.tipos || {}).filter(([, n]) => n > 0).map(([t, n]) => `${t} ${n}`).join(" · ");
        est = `<span class="est ok" title="${esc(tipos)}">Listo · ${esc(plural(a.filas, "fila", "filas"))}</span>`;
      } else if (a.estado === "error") {
        est = '<span class="est rojo">Error</span>';
      } else {
        const [txt, cls] = etiquetas[a.estado] || [a.estado, ""];
        est = `<span class="est ${cls}">${esc(txt)}</span>`;
      }
      const ocupado = a.estado === "subiendo" || a.estado === "quitando";
      return `<li><span class="nombre" title="${esc(a.nombre)}">${esc(a.nombre)}</span>${est}`
        + `<button class="quitar" data-quitar="${esc(a.nombre)}" title="Quitar ${esc(a.nombre)}" aria-label="Quitar ${esc(a.nombre)}"${ocupado ? " disabled" : ""}>&times;</button>`
        + (a.estado === "error" && a.error ? `<span class="err-msg">${esc(a.error)}</span>` : "") + "</li>";
    }).join("");
  }

  function pintarAvisos(st) {
    let html = st.errores.map((m, i) =>
      `<div class="aviso error" role="alert"><span>${esc(m)}</span><button class="cerrar" data-cerrar="${i}" aria-label="Cerrar">&times;</button></div>`).join("");
    if (st.errorCalculo) {
      html += `<div class="aviso error" role="alert"><span>No se pudo actualizar el resultado: ${esc(st.errorCalculo)}`
        + `${st.tablero ? " Se muestra el último resultado disponible." : ""}</span></div>`;
    }
    if (st.cambios && !st.actualizando && !st.tDebounce && !st.subiendo && !st.cola.length && st.sid) {
      html += '<div class="aviso error" role="alert"><span>Resultado desactualizado.</span>'
        + '<button class="reintentar" data-reintentar="1">Reintentar</button></div>';
    }
    $("avisos").innerHTML = html;
  }

  function pintarResultado(st) {
    const t = st.tablero;
    const [d, h] = (t && t.periodo) || [];
    $("periodo").textContent = d ? `Período: ${fmtFecha(d)} – ${fmtFecha(h)}` : "";
    const act = Boolean(st.tDebounce || st.actualizando);
    $("actualizando").hidden = !act;
    const pct = st.progreso != null && st.progreso > 0 ? ` (${Math.round(st.progreso * 100)} %)` : "";
    $("act-msg").textContent = `Actualizando…${st.mensaje && st.mensaje !== "Actualizando..." ? " " + st.mensaje : ""}${pct}`;

    const adv = (t && t.advertencias) || [];
    $("advertencias").innerHTML = adv.length
      ? `<div class="advert" role="note"><strong>Advertencias</strong><ul>${adv.map(a => `<li>${esc(a)}</li>`).join("")}</ul></div>`
      : "";

    if (!t) {
      const msg = act || st.subiendo ? "Conciliando contra Odoo…"
        : st.archivos.length ? "Sin resultado todavía." : "Sube los TXT del portal SRI para ver qué coincide con Odoo.";
      $("tarjetas").innerHTML = `<div class="vacio">${esc(msg)}</div>`;
      return;
    }
    if (!t.tarjetas.length) {
      const hayError = st.archivos.some(a => a.estado === "error");
      $("tarjetas").innerHTML = `<div class="vacio">${hayError
        ? "No hay comprobantes válidos para conciliar. Revisa los archivos con error."
        : "Agrega TXT para comenzar."}</div>`;
      return;
    }
    $("tarjetas").innerHTML = t.tarjetas.map(c => {
      const pctOk = c.total ? Math.min(100, Math.round((c.coinciden / c.total) * 1000) / 10) : 0;
      const otros = Object.entries(c.por_estado || {}).filter(([e]) => e !== "Coincide")
        .sort((a, b) => b[1] - a[1]).map(([e, n]) => `${esc(e)} ${esc(n)}`).join(" · ");
      const soloOdoo = (c.por_estado || {})["Solo en Odoo"] || 0;
      const sel = st.f.tipo === c.tipo;
      return `<button class="tarjeta${sel ? " sel" : ""}${c.total && c.coinciden === c.total ? " completa" : ""}" data-tipo="${esc(c.tipo)}" aria-pressed="${sel}" title="Filtrar la tabla por ${esc(c.etiqueta)}">`
        + `<span class="etq">${esc(c.etiqueta)}</span>`
        + `<span class="cifra"><b>${esc(c.coinciden)}</b> de ${esc(c.total)} comprobantes del SRI coinciden</span>`
        + `<span class="barra-av"><i style="width:${pctOk}%"></i></span>`
        + (soloOdoo ? `<span class="desglose">+${esc(soloOdoo)} solo en Odoo</span>` : "")
        + `<span class="desglose">${otros || "Sin observaciones"}</span></button>`;
    }).join("");
  }

  function comparar(col, desc) {
    const dir = desc ? -1 : 1;
    return (a, b) => {
      const x = a[col], y = b[col];
      if (x == null || x === "") return (y == null || y === "") ? 0 : 1;  // vacíos al final
      if (y == null || y === "") return -1;
      const r = typeof x === "number" && typeof y === "number" ? x - y
        : String(x).localeCompare(String(y), "es", { numeric: true, sensitivity: "base" });
      return r * dir;
    };
  }

  function pintarExcepciones(st) {
    const t = st.tablero, f = st.f;
    const exc = (t && t.excepciones) || [];
    const pendiente = Boolean(st.tDebounce || st.actualizando || st.cambios);
    const espera = "Espera a que termine la actualización";
    $("btn-obs").disabled = !st.sid || !exc.some(e => e.accionable) || pendiente;
    $("btn-completo").disabled = !st.sid || !t || !t.tarjetas.length || pendiente;
    $("btn-obs").title = $("btn-completo").title = pendiente ? espera : "";

    const tipos = (t && t.tarjetas) || [];
    if (f.tipo && !tipos.some(c => c.tipo === f.tipo)) f.tipo = "";
    $("f-tipo").innerHTML = '<option value="">Todos los tipos</option>' + tipos.map(c =>
      `<option value="${esc(c.tipo)}"${c.tipo === f.tipo ? " selected" : ""}>${esc(c.etiqueta)}</option>`).join("");

    const q = f.q.trim().toLowerCase();
    const delTipo = exc.filter(e => !f.tipo || e.tipo === f.tipo);
    const nInfo = delTipo.filter(e => esInfo(e.estado)).length;
    $("info-cuenta").textContent = `(${nInfo})`;
    const base = delTipo.filter(e => (f.info || !esInfo(e.estado))
      && (!q || [e.documento, e.contraparte, e.ruc].some(v => String(v ?? "").toLowerCase().includes(q))));

    const cuenta = {};
    base.forEach(e => { cuenta[e.estado] = (cuenta[e.estado] || 0) + 1; });
    if (f.estado && !cuenta[f.estado]) f.estado = "";
    $("chips-estado").hidden = !exc.length;
    $("chips-estado").innerHTML = `<button class="chip${f.estado ? "" : " sel"}" data-estado="">Todos<b>${base.length}</b></button>`
      + Object.entries(cuenta).sort((a, b) => b[1] - a[1]).map(([e, n]) =>
        `<button class="chip${f.estado === e ? " sel" : ""}" data-estado="${esc(e)}">${esc(e)}<b>${n}</b></button>`).join("");

    let filas = base.filter(e => !f.estado || e.estado === f.estado);
    if (f.orden) filas = filas.slice().sort(comparar(f.orden, f.desc));

    let vacio = null;
    if (!t) vacio = '<div class="vacio">El detalle aparecerá aquí cuando termine la primera conciliación.</div>';
    else if (!t.tarjetas.length) vacio = '<div class="vacio">Sin comprobantes conciliados.</div>';
    else if (!exc.length) vacio = '<div class="vacio bien"><strong>Todo coincide</strong>Todos los comprobantes cargados están conciliados con Odoo.</div>';
    else if (!exc.some(e => !esInfo(e.estado)) && !f.info) {
      vacio = `<div class="vacio bien"><strong>Nada que corregir</strong>Solo hay ${esc(plural(exc.length, "registro informativo", "registros informativos"))}; actívalos con «Informativos».</div>`;
    } else if (!filas.length) vacio = '<div class="vacio">Ningún registro coincide con los filtros.</div>';
    if (vacio) { $("tabla").innerHTML = vacio; return; }

    const cab = COLUMNAS.map(([k, titulo]) => {
      const flecha = f.orden === k ? (f.desc ? "▼" : "▲") : "";
      return `<th data-orden="${k}" aria-sort="${f.orden === k ? (f.desc ? "descending" : "ascending") : "none"}">${esc(titulo)}<span class="ord">${flecha}</span></th>`;
    }).join("");
    const cuerpo = filas.slice(0, MAX_FILAS).map(e => {
      const celdas = COLUMNAS.map(([k, , tipo]) => {
        const v = e[k];
        if (k === "estado") return `<td>${insignia(v)}</td>`;
        if (k === "fecha") return `<td>${esc(fmtFecha(v))}</td>`;
        if (k === "documento") return `<td class="doc">${esc(v)}</td>`;
        if (k === "detalle") return `<td class="det">${esc(v)}</td>`;
        if (k === "que_corregir") return `<td class="corr">${esc(v)}</td>`;
        if (tipo === "num") return `<td class="num${k === "diferencia" && v ? " neg" : ""}">${esc(fmtNum(v))}</td>`;
        return `<td>${esc(v)}</td>`;
      }).join("");
      return `<tr${esInfo(e.estado) ? ' class="info"' : ""}>${celdas}</tr>`;
    }).join("");
    const pie = filas.length > MAX_FILAS
      ? `<div class="tabla-pie">Se muestran ${MAX_FILAS} de ${filas.length} registros; usa los filtros o descarga el Excel para verlos todos.</div>`
      : `<div class="tabla-pie">${esc(plural(filas.length, "registro", "registros"))}</div>`;
    $("tabla").innerHTML = `<table><thead><tr>${cab}</tr></thead><tbody>${cuerpo}</tbody></table>${pie}`;
  }

  // ---------------------------------------------------------------- eventos
  function cambiarModo(m) {
    if (!estados[m] || m === modo) return;
    modo = m;
    $("f-buscar").value = actual().f.q;
    $("f-info").checked = actual().f.info;
    render();
  }

  function nuevaConciliacion() {
    const st = actual();
    if (st.archivos.length && !confirm(`¿Descartar los TXT y el resultado de ${st.modo === "compras" ? "Compras" : "Ventas"}?`)) return;
    clearTimeout(st.tDebounce);
    clearTimeout(st.tSondeo);
    estados[st.modo] = nuevoEstado(st.modo);  // las respuestas pendientes de la sesión anterior se ignoran
    $("f-buscar").value = "";
    $("f-info").checked = false;
    render();
  }

  document.addEventListener("DOMContentLoaded", () => {
    const drop = $("drop"), input = $("input-txt");
    document.querySelector(".tabs").addEventListener("click", e => {
      const b = e.target.closest("[data-modo]");
      if (b) cambiarModo(b.dataset.modo);
    });
    $("btn-nueva").addEventListener("click", nuevaConciliacion);
    input.addEventListener("change", () => { agregar([...input.files]); input.value = ""; });
    drop.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); }
    });
    ["dragenter", "dragover"].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add("sobre"); }));
    ["dragleave", "drop"].forEach(ev => drop.addEventListener(ev, () => drop.classList.remove("sobre")));
    drop.addEventListener("drop", e => { e.preventDefault(); agregar([...e.dataTransfer.files]); });
    // Soltar fuera de la zona no debe abrir el archivo en el navegador.
    ["dragover", "drop"].forEach(ev => window.addEventListener(ev, e => e.preventDefault()));

    $("lista-archivos").addEventListener("click", e => {
      const b = e.target.closest("[data-quitar]");
      if (b) quitar(actual(), b.dataset.quitar);
    });
    $("avisos").addEventListener("click", e => {
      const b = e.target.closest("[data-cerrar]");
      if (b) { actual().errores.splice(Number(b.dataset.cerrar), 1); render(); }
      if (e.target.closest("[data-reintentar]")) conciliar(actual());
    });
    $("tarjetas").addEventListener("click", e => {
      const b = e.target.closest("[data-tipo]");
      if (!b) return;
      const f = actual().f;
      f.tipo = f.tipo === b.dataset.tipo ? "" : b.dataset.tipo;
      f.estado = "";
      render();
    });
    $("chips-estado").addEventListener("click", e => {
      const b = e.target.closest("[data-estado]");
      if (b) { actual().f.estado = b.dataset.estado; render(); }
    });
    $("f-tipo").addEventListener("change", e => { const f = actual().f; f.tipo = e.target.value; f.estado = ""; render(); });
    $("f-buscar").addEventListener("input", e => { actual().f.q = e.target.value; render(); });
    $("f-info").addEventListener("change", e => { actual().f.info = e.target.checked; render(); });
    $("tabla").addEventListener("click", e => {
      const th = e.target.closest("[data-orden]");
      if (!th) return;
      const f = actual().f;
      if (f.orden === th.dataset.orden) f.desc = !f.desc; else { f.orden = th.dataset.orden; f.desc = false; }
      render();
    });
    $("btn-obs").addEventListener("click", () => descargar("observaciones"));
    $("btn-completo").addEventListener("click", () => descargar("completo"));
    render();
  });
})();
