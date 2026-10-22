// Activos ESCINF: pantalla para consultar el material y equipo del Excel del S-4.
(function () {
  'use strict';

  const CLAVE = 'activos_escinf_datos';
  const REVISTA = {
    revistado: { texto: 'Revistado', punto: '🟠' },
    no_revistado: { texto: 'No revistado', punto: '🟡' },
    por_revistar: { texto: 'Por revistar', punto: '🟢' }
  };

  const $ = (id) => document.getElementById(id);
  const estado = { datos: null, tipo: '', vista: 'buscar' };

  const normal = (t) => String(t || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const esc = (t) => String(t == null ? '' : t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const miles = (n) => Number(n || 0).toLocaleString('es-CL');

  // ---------- Guardar / recuperar ----------
  function guardar(datos) {
    try { localStorage.setItem(CLAVE, JSON.stringify(datos)); } catch (e) { /* sin memoria: igual se muestra */ }
  }
  function recuperar() {
    try { return JSON.parse(localStorage.getItem(CLAVE) || 'null'); } catch (e) { return null; }
  }

  // ---------- Cargar el Excel ----------
  function cargarArchivo(archivo) {
    if (!archivo) return;
    const lector = new FileReader();
    lector.onload = () => {
      try {
        const libro = XLSX.read(new Uint8Array(lector.result), { type: 'array', cellStyles: true });
        const datos = LeerExcel.leer(XLSX, libro, archivo.name);
        guardar(datos);
        mostrarDatos(datos);
      } catch (e) {
        alert('No pude leer el Excel.\n\n' + e.message);
      }
    };
    lector.readAsArrayBuffer(archivo);
  }

  // ---------- Mostrar ----------
  function mostrarDatos(datos) {
    estado.datos = datos;
    estado.tipo = '';
    $('bienvenida').hidden = true;
    $('pestanas').hidden = false;
    $('imprimir').hidden = false;
    const fecha = new Date(datos.cargado).toLocaleDateString('es-CL');
    $('fuente').textContent = (datos.archivo || 'Excel cargado') + ' · cargado el ' + fecha;

    const sel = $('filtroUnidad');
    sel.innerHTML = '<option value="">Todas las unidades</option>' +
      datos.unidades.map((u) => `<option value="${esc(u)}">${esc(u)}</option>`).join('');
    // Filtros que solo se muestran si el Excel trae esa información
    $('filtroRevista').hidden = !datos.activos.some((a) => a.revista);
    $('filtroNota').hidden = !datos.activos.some((a) => a.nota);
    if ($('filtroRevista').hidden) $('filtroRevista').value = '';
    if ($('filtroNota').hidden) $('filtroNota').value = '';
    // Pestañas de hojas que este Excel no trae
    const pestana = (v) => document.querySelector(`#pestanas [data-vista="${v}"]`);
    pestana('movimientos').hidden = !datos.movimientos.filas.length;
    pestana('combustible').hidden = !datos.combustible.filas.length;
    if (pestana(estado.vista).hidden) estado.vista = 'buscar';

    pintarResumen();
    pintarTipos();
    pintarLista();
    pintarUnidades();
    pintarTabla($('tablaMovimientos'), datos.movimientos, ['CANTIDAD']);
    pintarTabla($('tablaCombustible'), datos.combustible, ['CANTIDAD']);
    pintarAlertas();
    irA(estado.vista);
  }

  function irA(vista) {
    estado.vista = vista;
    document.querySelectorAll('.vista').forEach((s) => { s.hidden = s.id !== 'vista-' + vista; });
    document.querySelectorAll('#pestanas button').forEach((b) => b.classList.toggle('activa', b.dataset.vista === vista));
  }

  // Cantidad que se considera según la unidad elegida
  function cantidad(a, unidad) { return unidad ? (a.reparto[unidad] || 0) : a.total; }

  function filtrados() {
    const unidad = $('filtroUnidad').value;
    const revista = $('filtroRevista').value;
    const conNota = $('filtroNota').value === 'si';
    const palabras = normal($('busqueda').value).split(/\s+/).filter(Boolean);
    let lista = estado.datos.activos.filter((a) => {
      if (unidad && !a.reparto[unidad]) return false;
      if (estado.tipo && a.tipo !== estado.tipo) return false;
      if (revista && a.revista !== revista) return false;
      if (conNota && !a.nota) return false;
      const texto = normal(a.nombre + ' ' + a.tipo + ' ' + a.cimage + ' ' + (a.nota || ''));
      return palabras.every((p) => texto.includes(p));
    });
    const orden = $('orden').value;
    if (orden === 'cantidad') lista = lista.slice().sort((x, y) => cantidad(y, unidad) - cantidad(x, unidad));
    if (orden === 'nombre') lista = lista.slice().sort((x, y) => x.nombre.localeCompare(y.nombre, 'es'));
    return lista;
  }

  function pintarResumen() {
    const d = estado.datos;
    const total = d.activos.reduce((s, a) => s + a.total, 0);
    const tipos = new Set(d.activos.map((a) => a.tipo)).size;
    const armas = d.activos.filter((a) => normal(a.tipo) === 'MATERIAL DE GUERRA').reduce((s, a) => s + a.total, 0);
    const vehiculos = d.activos.filter((a) => normal(a.tipo) === 'VEHICULO').reduce((s, a) => s + a.total, 0);
    $('resumen').innerHTML = [
      [d.activos.length, 'clases de material distintas'],
      [miles(total), 'unidades en total'],
      [miles(armas), 'de material de guerra'],
      [miles(vehiculos), 'vehículos y remolques'],
      [d.unidades.length, 'unidades / almacenes'],
      [tipos, 'tipos de activo']
    ].map(([n, t]) => `<div class="cifra"><b>${n}</b><span>${t}</span></div>`).join('');
  }

  function pintarTipos() {
    const cuenta = {};
    estado.datos.activos.forEach((a) => { cuenta[a.tipo] = (cuenta[a.tipo] || 0) + 1; });
    const tipos = Object.keys(cuenta).sort((a, b) => (a === LeerExcel.TIPO_SIN) - (b === LeerExcel.TIPO_SIN) || a.localeCompare(b, 'es'));
    $('tipos').innerHTML = `<button class="chip ${estado.tipo ? '' : 'activo'}" data-tipo="">Todo <small>${estado.datos.activos.length}</small></button>` +
      tipos.map((t) => `<button class="chip ${estado.tipo === t ? 'activo' : ''}" data-tipo="${esc(t)}">${esc(t)} <small>${cuenta[t]}</small></button>`).join('');
  }

  function resaltar(texto, palabras) {
    let html = esc(texto);
    if (!palabras.length) return html;
    // Resalta respetando tildes: compara sobre el texto normalizado letra por letra
    const plano = normal(texto);
    const marcas = new Array(texto.length).fill(false);
    palabras.forEach((p) => {
      let i = plano.indexOf(p);
      while (i >= 0) { for (let k = i; k < i + p.length; k++) marcas[k] = true; i = plano.indexOf(p, i + 1); }
    });
    if (plano.length !== texto.length) return html;
    html = '';
    for (let k = 0; k < texto.length; k++) {
      const c = esc(texto[k]);
      if (marcas[k] && !marcas[k - 1]) html += '<mark>';
      html += c;
      if (marcas[k] && !marcas[k + 1]) html += '</mark>';
    }
    return html;
  }

  function pintarLista() {
    const unidad = $('filtroUnidad').value;
    const palabras = normal($('busqueda').value).split(/\s+/).filter(Boolean);
    const lista = filtrados();
    const suma = lista.reduce((s, a) => s + cantidad(a, unidad), 0);
    $('contador').textContent = lista.length
      ? `${lista.length} resultado${lista.length === 1 ? '' : 's'} · ${miles(suma)} unidades${unidad ? ' a cargo de ' + unidad : ''}`
      : 'No hay material con esos filtros.';
    $('lista').innerHTML = lista.map((a) => {
      const reparto = Object.entries(a.reparto)
        .map(([u, n]) => `<span class="etiqueta ${u === unidad ? 'marcada' : ''}">${esc(u)}: <b>${miles(n)}</b></span>`).join('');
      const rev = a.revista ? `${REVISTA[a.revista].punto} ${REVISTA[a.revista].texto} · ` : '';
      const aviso = a.diferencia ? `<div class="alerta-txt">⚠ El cargo dice ${miles(a.cargo)} y lo distribuido suma ${miles(a.total)}</div>` : '';
      return `<button class="item ${a.revista || ''}" data-fila="${a.fila}">
        <div><div class="tipo">${rev}${esc(a.tipo)}</div><div class="nombre">${resaltar(a.nombre, palabras)}</div></div>
        <div class="total"><b>${miles(cantidad(a, unidad))}</b><small>${unidad ? 'en ' + esc(unidad) : 'en total'}</small></div>
        <div class="reparto">${reparto || '<span class="nota">Sin distribución</span>'}</div>
        ${aviso}
        ${a.nota ? `<div class="nota-txt">📝 ${resaltar(a.nota, palabras)}</div>` : ''}
      </button>`;
    }).join('');
  }

  function abrirFicha(fila) {
    const a = estado.datos.activos.find((x) => x.fila === fila);
    if (!a) return;
    const max = Math.max(1, ...Object.values(a.reparto));
    const barras = estado.datos.unidades.filter((u) => a.reparto[u]).map((u) => `
      <div class="barra-unidad"><span>${esc(u)}</span>
        <div class="fondo-barra"><div class="relleno" style="width:${(a.reparto[u] / max) * 100}%"></div></div>
        <span class="num">${miles(a.reparto[u])}</span></div>`).join('');
    // Movimientos OAL que nombran este material (por la palabra más distintiva)
    const mov = estado.datos.movimientos;
    const iNom = mov.titulos.findIndex((t) => normal(t) === 'NOMENCLATURA');
    const claves = normal(a.nombre).split(/[^A-Z0-9]+/).filter((p) => p.length >= 4 && !/^(FUSIL|CAL|MOD|PARA|CON)$/.test(p));
    const relacionados = iNom < 0 ? [] : mov.filas.filter((f) => claves.some((c) => normal(f[iNom]).includes(c)));
    $('fichaContenido').innerHTML = `
      <div class="tipo">${esc(a.tipo)}</div>
      <h2>${esc(a.nombre)}</h2>
      <dl class="datos-ficha">
        <dt>Total distribuido</dt><dd><b>${miles(a.total)}</b></dd>
        <dt>Cargo</dt><dd>${miles(a.cargo)}${a.diferencia ? ` <span class="alerta-txt">(diferencia ${miles(a.diferencia)})</span>` : ''}</dd>
        ${a.nota ? `<dt>Observación</dt><dd class="nota-txt">📝 ${esc(a.nota)}</dd>` : ''}
        <dt>CIMAGE</dt><dd>${esc(a.cimage) || '—'}</dd>
        <dt>Revista</dt><dd>${a.revista ? REVISTA[a.revista].punto + ' ' + REVISTA[a.revista].texto : 'Sin marca en el Excel'}</dd>
        <dt>En el Excel</dt><dd>N° ${esc(a.n) || '—'}, fila ${a.fila}</dd>
      </dl>
      <h3>¿Dónde está?</h3>
      ${barras || '<p class="nota">No está distribuido en ninguna unidad.</p>'}
      ${relacionados.length ? `<h3>Movimientos OAL relacionados</h3>${tablaHTML({ titulos: mov.titulos, filas: relacionados }, ['CANTIDAD'])}` : ''}`;
    $('ficha').showModal();
    $('ficha').scrollTop = 0;
  }

  function pintarUnidades() {
    const d = estado.datos;
    $('tarjetasUnidades').innerHTML = d.unidades.map((u) => {
      const suyos = d.activos.filter((a) => a.reparto[u]);
      const porTipo = {};
      suyos.forEach((a) => { porTipo[a.tipo] = (porTipo[a.tipo] || 0) + a.reparto[u]; });
      const lineas = Object.entries(porTipo).sort((x, y) => y[1] - x[1])
        .map(([t, n]) => `<div class="linea-tipo"><span>${esc(t)}</span><b>${miles(n)}</b></div>`).join('');
      const total = suyos.reduce((s, a) => s + a.reparto[u], 0);
      return `<button class="unidad ${suyos.length ? '' : 'vacia'}" data-unidad="${esc(u)}">
        <h3>${esc(u)}</h3>
        <div class="nota">${suyos.length} clases de material · ${miles(total)} unidades</div>
        ${lineas}
      </button>`;
    }).join('');
  }

  function tablaHTML(tabla, columnasNumero) {
    if (!tabla || !tabla.filas.length) return '<p class="nota">Esta hoja no tiene datos.</p>';
    const esNum = tabla.titulos.map((t) => columnasNumero.some((c) => normal(t).startsWith(c)));
    return `<div class="tabla-scroll"><table><thead><tr>${tabla.titulos.map((t) => `<th>${esc(t)}</th>`).join('')}</tr></thead>
      <tbody>${tabla.filas.map((f) => `<tr>${tabla.titulos.map((_, i) => `<td class="${esNum[i] ? 'num' : ''}">${esc(f[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  }
  function pintarTabla(destino, tabla, columnasNumero) { destino.innerHTML = tablaHTML(tabla, columnasNumero); }

  function pintarAlertas() {
    const d = estado.datos;
    const bloques = [];
    const conNota = d.activos.filter((a) => a.nota);
    if (conNota.length) {
      bloques.push(`<div class="bloque-alerta"><h3>📝 Observaciones del S-4 (${conNota.length})</h3>
        ${tablaHTML({ titulos: ['N°', 'Nomenclatura', 'Cantidad', 'Observación'],
          filas: conNota.map((a) => [a.n, a.nombre, miles(a.total), a.nota]) }, ['CANTIDAD'])}</div>`);
    }
    const conDiferencia = d.activos.filter((a) => a.diferencia);
    if (conDiferencia.length) {
      bloques.push(`<div class="bloque-alerta grave"><h3>⚠️ Cargo distinto a lo distribuido (${conDiferencia.length})</h3>
        <p>La columna CARGO no coincide con la suma de las unidades. Revisar en el Excel.</p>
        ${tablaHTML({ titulos: ['N°', 'Nomenclatura', 'Cargo', 'Distribuido', 'Diferencia'],
          filas: conDiferencia.map((a) => [a.n, a.nombre, miles(a.cargo), miles(a.total), miles(a.diferencia)]) }, ['CARGO', 'DISTRIBUIDO', 'DIFERENCIA'])}</div>`);
    }
    const sinReparto = d.activos.filter((a) => !a.total);
    if (sinReparto.length) {
      bloques.push(`<div class="bloque-alerta"><h3>📭 Material sin unidad asignada (${sinReparto.length})</h3>
        ${tablaHTML({ titulos: ['N°', 'Nomenclatura', 'Cargo'], filas: sinReparto.map((a) => [a.n, a.nombre, miles(a.cargo)]) }, ['CARGO'])}</div>`);
    }
    const noRev = d.activos.filter((a) => a.revista === 'no_revistado' || a.revista === 'por_revistar');
    if (noRev.length) {
      bloques.push(`<div class="bloque-alerta"><h3>🟡 Material de guerra pendiente de revista (${noRev.length})</h3>
        ${tablaHTML({ titulos: ['N°', 'Nomenclatura', 'Estado', 'Cantidad'],
          filas: noRev.map((a) => [a.n, a.nombre, REVISTA[a.revista].texto, miles(a.total)]) }, ['CANTIDAD'])}</div>`);
    }
    const sinTipo = d.activos.filter((a) => a.tipo === LeerExcel.TIPO_SIN).length;
    if (sinTipo) {
      bloques.push(`<div class="bloque-alerta"><h3>❔ Material sin «tipo de activo» (${sinTipo})</h3>
        <p>En el Excel la columna TIPO DE ACTIVO dice «?». Aparecen en el filtro «${LeerExcel.TIPO_SIN}». Si se completa esa columna, la búsqueda por tipo queda más ordenada.</p></div>`);
    }
    $('listaAlertas').innerHTML = bloques.join('') || '<p>Sin alertas. 👍</p>';
    $('nAlertas').textContent = conNota.length + conDiferencia.length + sinReparto.length || '';
  }

  // ---------- Eventos ----------
  $('archivo').addEventListener('change', (e) => { cargarArchivo(e.target.files[0]); e.target.value = ''; });
  $('imprimir').addEventListener('click', () => window.print());
  ['busqueda', 'filtroUnidad', 'filtroRevista', 'filtroNota', 'orden'].forEach((id) => $(id).addEventListener('input', pintarLista));
  $('tipos').addEventListener('click', (e) => {
    const b = e.target.closest('.chip');
    if (!b) return;
    estado.tipo = b.dataset.tipo;
    pintarTipos();
    pintarLista();
  });
  $('lista').addEventListener('click', (e) => {
    const b = e.target.closest('.item');
    if (b) abrirFicha(Number(b.dataset.fila));
  });
  $('tarjetasUnidades').addEventListener('click', (e) => {
    const b = e.target.closest('.unidad');
    if (!b) return;
    $('filtroUnidad').value = b.dataset.unidad;
    $('busqueda').value = '';
    estado.tipo = '';
    pintarTipos();
    pintarLista();
    irA('buscar');
  });
  $('pestanas').addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (b) irA(b.dataset.vista);
  });
  $('ficha').addEventListener('click', (e) => { if (e.target === $('ficha')) $('ficha').close(); });

  // Arrastrar el archivo sobre la página
  const zona = $('zonaCarga');
  document.addEventListener('dragover', (e) => { e.preventDefault(); zona.classList.add('encima'); });
  document.addEventListener('dragleave', () => zona.classList.remove('encima'));
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    zona.classList.remove('encima');
    cargarArchivo(e.dataTransfer.files[0]);
  });

  const previos = recuperar();
  if (previos && previos.activos) mostrarDatos(previos);
  window.ActivosESCINF = { mostrarDatos };

  // Instalable y sin internet (solo cuando se abre desde una dirección web, no como archivo)
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(() => {});
  }
})();
