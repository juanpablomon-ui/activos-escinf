// Activos ESCINF: pantalla para consultar el material y equipo del Excel del S-4.
(function () {
  'use strict';

  // Cada Excel se guarda por separado (la clave de activos es la de siempre, para no perder lo ya cargado)
  const CLAVES = { activos: 'activos_escinf_datos', vestuario: 'activos_escinf_vestuario', municion: 'activos_escinf_municion' };
  const LIBROS = {
    activos: { nombre: 'activos bélicos', ejemplo: 'DISTRIBUCIÓN ACTIVOS BÉLICOS ESCINF S-4 22OCT2026.xlsx' },
    vestuario: { nombre: 'vestuario y equipo', ejemplo: 'CUADRO DISTRIBUCIÓN VEST Y EQUIPO ESCINF 2026.xlsx' },
    municion: { nombre: 'munición', ejemplo: 'CUADROS MUNICIÓN AL 30OCT.xlsx' }
  };
  const CLAVE_LIBRO = 'activos_escinf_libro';
  const DIAS_LIMITE = 15; // pasado este plazo, la información se considera vieja
  const REVISTA = {
    revistado: { texto: 'Revistado', punto: '🟠' },
    no_revistado: { texto: 'No revistado', punto: '🟡' },
    por_revistar: { texto: 'Por revistar', punto: '🟢' }
  };

  const $ = (id) => document.getElementById(id);
  const estado = { libros: { activos: null, vestuario: null, municion: null }, libro: 'activos', datos: null, tipo: '', vista: 'buscar' };

  const normal = (t) => String(t || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const esc = (t) => String(t == null ? '' : t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const miles = (n) => Number(n || 0).toLocaleString('es-CL');
  const dia = (iso) => new Date(iso).toLocaleDateString('es-CL');

  // ---------- Antigüedad de la información ----------
  // Días desde la fecha de los datos (última modificación del Excel); los datos cargados antes de esta versión usan la fecha de carga
  function antiguedad(datos) {
    const fecha = datos.fechaDatos || datos.cargado;
    const dias = Math.max(0, Math.floor((Date.now() - new Date(fecha)) / 86400000));
    return { fecha, dias, vieja: dias > DIAS_LIMITE, origen: datos.origenFecha || 'día en que se cargó' };
  }
  const textoDias = (n) => (n === 0 ? 'de hoy' : n === 1 ? 'de ayer' : `de hace ${n} días`);

  // ---------- Guardar / recuperar ----------
  function guardar(clase, datos) {
    try { localStorage.setItem(CLAVES[clase], JSON.stringify(datos)); } catch (e) { /* sin memoria: igual se muestra */ }
  }
  function recuperar(clase) {
    try { return JSON.parse(localStorage.getItem(CLAVES[clase]) || 'null'); } catch (e) { return null; }
  }
  function recordarLibro() {
    try { localStorage.setItem(CLAVE_LIBRO, estado.libro); } catch (e) { /* no importa */ }
  }

  // ---------- Cargar el Excel ----------
  function cargarArchivo(archivo) {
    if (!archivo) return;
    const lector = new FileReader();
    lector.onload = () => {
      try {
        const libro = XLSX.read(new Uint8Array(lector.result), { type: 'array', cellStyles: true });
        const datos = LeerExcel.leer(XLSX, libro, archivo.name, archivo.lastModified);
        const clase = datos.clase || 'activos';
        guardar(clase, datos);
        estado.libros[clase] = datos;
        cambiarLibro(clase);
      } catch (e) {
        alert('No pude leer el Excel.\n\n' + e.message);
      }
    };
    lector.readAsArrayBuffer(archivo);
  }

  // ---------- Mostrar ----------
  function pintarLibros() {
    document.querySelectorAll('#libros button').forEach((b) => {
      const d = estado.libros[b.dataset.libro];
      const a = d && antiguedad(d);
      b.classList.toggle('activo', b.dataset.libro === estado.libro);
      b.classList.toggle('vieja', !!(a && a.vieja));
      b.querySelector('small').textContent = !d ? '(sin cargar)' : (a.vieja ? '⚠️ ' : '🟢 ') + a.dias + ' d';
      b.title = !d ? 'Aún no se carga este Excel' : `Información ${textoDias(a.dias)} (${dia(a.fecha)})`;
    });
  }

  // Aviso rojo sobre la pantalla cuando la información de la sección tiene más de 15 días
  function pintarAvisoViejo() {
    const d = estado.datos;
    const a = d && antiguedad(d);
    $('avisoViejo').hidden = !(a && a.vieja);
    if (!a || !a.vieja) return;
    $('avisoViejo').innerHTML = `<b>⚠️ Información de hace ${a.dias} días.</b>
      El Excel es del ${dia(a.fecha)} (${esc(a.origen)}) y pasaron más de ${DIAS_LIMITE} días.
      Pida al S-4 el Excel actualizado y cárguelo con <b>📂 Cargar Excel</b>.`;
  }

  // Cuadro con cuándo se cargó cada Excel y de qué fecha es su información
  function cuadroFechas() {
    const mayuscula = (t) => t.charAt(0).toUpperCase() + t.slice(1);
    const filas = Object.keys(LIBROS).map((clase) => {
      const d = estado.libros[clase];
      const nombre = mayuscula(LIBROS[clase].nombre);
      if (!d) return `<tr><td>${esc(nombre)}</td><td colspan="4" class="nota">Sin cargar</td></tr>`;
      const a = antiguedad(d);
      return `<tr class="${a.vieja ? 'fila-vieja' : ''}"><td>${esc(nombre)}<br><small class="nota">${esc(d.archivo)}</small></td>
        <td>${dia(a.fecha)}<br><small class="nota">${esc(a.origen)}</small></td>
        <td>${dia(d.cargado)}</td>
        <td class="num">${a.dias}</td>
        <td>${a.vieja ? '🔴 Vieja: pedir Excel nuevo' : '🟢 Al día'}</td></tr>`;
    }).join('');
    return `<div class="bloque-alerta"><h3>📅 Excel cargados</h3>
      <p class="nota">La información se marca en rojo si tiene más de ${DIAS_LIMITE} días.</p>
      <div class="tabla-scroll"><table><thead><tr><th>Sección</th><th>Información del</th><th>Cargado el</th><th>Días</th><th>Estado</th></tr></thead>
      <tbody>${filas}</tbody></table></div></div>`;
  }

  // Pasa a la sección elegida: muestra sus datos o, si aún no se carga su Excel, la bienvenida
  function cambiarLibro(clase) {
    estado.libro = clase;
    recordarLibro();
    pintarLibros();
    // al cambiar de Excel, los filtros vuelven a cero
    $('busqueda').value = '';
    $('filtroRevista').value = '';
    $('filtroNota').value = '';
    $('busquedaMun').value = '';
    $('filtroMun').value = '';
    const datos = estado.libros[clase];
    if (datos) { mostrarDatos(datos); return; }
    estado.datos = null;
    $('avisoViejo').hidden = true;
    $('bienvenida').hidden = false;
    $('pestanas').hidden = true;
    $('imprimir').hidden = true;
    document.querySelectorAll('.vista').forEach((s) => { s.hidden = true; });
    $('bienvenidaTitulo').textContent = 'Cargue el Excel de ' + LIBROS[clase].nombre;
    $('bienvenidaEjemplo').textContent = LIBROS[clase].ejemplo;
    $('fuente').textContent = 'Material y equipo de la Escuela de Infantería';
  }

  function mostrarDatos(datos) {
    estado.datos = datos;
    estado.tipo = '';
    $('bienvenida').hidden = true;
    $('pestanas').hidden = false;
    $('imprimir').hidden = false;
    const a = antiguedad(datos);
    $('fuente').textContent = `${datos.archivo || 'Excel cargado'} · información del ${dia(a.fecha)} · cargado el ${dia(datos.cargado)}`;
    pintarAvisoViejo();
    // Pestañas: la munición tiene su propia vista; las demás, buscar y por unidad
    const pestana = (v) => document.querySelector(`#pestanas [data-vista="${v}"]`);
    const esMunicion = datos.clase === 'municion';
    pestana('municion').hidden = !esMunicion;
    pestana('buscar').hidden = esMunicion;
    pestana('unidades').hidden = esMunicion;
    if (esMunicion) {
      pestana('movimientos').hidden = true;
      pestana('combustible').hidden = true;
      pintarMunicion();
      pintarAlertas();
      irA(estado.vista === 'alertas' ? 'alertas' : 'municion');
      return;
    }
    $('tipos').hidden = new Set(datos.activos.map((a) => a.tipo)).size < 2;
    $('busqueda').placeholder = 'Escriba lo que busca: ' +
      (datos.clase === 'vestuario' ? 'casco, mochila, carpa, botas…' : 'fusil, galil, radio, casco, camioneta…');

    const sel = $('filtroUnidad');
    sel.innerHTML = '<option value="">Todas las unidades</option>' +
      datos.unidades.map((u) => `<option value="${esc(u)}">${esc(u)}</option>`).join('');
    // Filtros que solo se muestran si el Excel trae esa información
    $('filtroRevista').hidden = !datos.activos.some((a) => a.revista);
    $('filtroNota').hidden = !datos.activos.some((a) => a.nota);
    if ($('filtroRevista').hidden) $('filtroRevista').value = '';
    if ($('filtroNota').hidden) $('filtroNota').value = '';
    // Pestañas de hojas que este Excel no trae
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

  // ---------- Munición ----------
  // Parte del año transcurrida a la fecha de corte del cuadro: lo que «debería» llevarse consumido si se consume parejo
  function avanceDelAnio(d) {
    const corte = new Date(d.corte || d.fechaDatos || d.cargado);
    const inicio = new Date(corte.getFullYear(), 0, 1);
    const fin = new Date(corte.getFullYear() + 1, 0, 1);
    return { corte, esperado: Math.min(1, (corte - inicio) / (fin - inicio)) };
  }

  // Cuentas de cada munición: % consumido, lo que falta consumir y si alcanza lo disponible
  function cuentasMunicion(m, esperado) {
    const pct = m.asignado ? m.consumido / m.asignado : null;
    const falta = Math.max(0, m.asignado - m.consumido);
    return {
      pct,
      falta,
      atrasada: m.asignado > 0 && pct < esperado,
      noAlcanza: m.asignado > 0 && falta > m.disponible,
      descuadre: m.sumaActividades !== m.consumido
    };
  }
  const porc = (x) => (x == null ? '—' : Math.round(x * 100) + '%');

  function pintarMunicion() {
    const d = estado.datos;
    const { corte, esperado } = avanceDelAnio(d);
    $('corteMun').innerHTML = `${esc(d.titulo || 'Situación de munición')}.<br>
      Al ${corte.toLocaleDateString('es-CL')} ya pasó el <b>${porc(esperado)}</b> del año: la munición asignada
      que lleve menos que eso consumido se marca <b style="color:var(--rojo)">en rojo (consumo atrasado)</b>.`;
    const asignadas = d.municiones.filter((m) => m.asignado > 0);
    const asignado = asignadas.reduce((s, m) => s + m.asignado, 0);
    const consumido = asignadas.reduce((s, m) => s + Math.min(m.consumido, m.asignado), 0);
    const atrasadas = asignadas.filter((m) => cuentasMunicion(m, esperado).atrasada).length;
    const noAlcanza = asignadas.filter((m) => cuentasMunicion(m, esperado).noAlcanza).length;
    $('resumenMun').innerHTML = [
      [d.municiones.length, 'tipos de munición'],
      [asignadas.length, 'con asignación de consumo 2026'],
      [porc(asignado ? consumido / asignado : 0), 'consumido de lo asignado (total)'],
      [atrasadas, 'con consumo atrasado'],
      [noAlcanza, 'no alcanza lo disponible'],
      [d.municiones.filter((m) => m.disponible > 0).length, 'con munición disponible']
    ].map(([n, t]) => `<div class="cifra"><b>${n}</b><span>${t}</span></div>`).join('');
    pintarListaMunicion();
  }

  function pintarListaMunicion() {
    const d = estado.datos;
    const { esperado } = avanceDelAnio(d);
    const palabras = normal($('busquedaMun').value).split(/\s+/).filter(Boolean);
    const filtro = $('filtroMun').value;
    let lista = d.municiones.map((m) => ({ m, c: cuentasMunicion(m, esperado) })).filter(({ m, c }) => {
      if (filtro === 'asignada' && !m.asignado) return false;
      if (filtro === 'atrasada' && !c.atrasada) return false;
      if (filtro === 'no_alcanza' && !c.noAlcanza) return false;
      if (filtro === 'disponible' && !m.disponible) return false;
      if (filtro === 'nota' && !m.nota) return false;
      const texto = normal(m.nombre + ' ' + m.nota);
      return palabras.every((p) => texto.includes(p));
    });
    const orden = $('ordenMun').value;
    if (orden === 'falta') lista.sort((x, y) => y.c.falta - x.c.falta);
    if (orden === 'porcentaje') lista.sort((x, y) => (x.c.pct ?? 9) - (y.c.pct ?? 9));
    if (orden === 'disponible') lista.sort((x, y) => y.m.disponible - x.m.disponible);
    $('contadorMun').textContent = lista.length ? `${lista.length} tipo${lista.length === 1 ? '' : 's'} de munición` : 'No hay munición con esos filtros.';
    $('listaMun').innerHTML = lista.map(({ m, c }) => {
      const actividades = Object.entries(m.actividades)
        .map(([a, n]) => `<span class="etiqueta">${esc(a)}: <b>${miles(n)}</b></span>`).join('');
      const barra = m.asignado ? `<div class="consumo">
          <div class="fondo-barra"><div class="relleno" style="width:${Math.min(100, c.pct * 100)}%"></div>
            <div class="meta" style="left:${esperado * 100}%" title="Lo que debería llevarse a la fecha de corte"></div></div>
          <div class="linea"><span>Asignado <b>${miles(m.asignado)}</b></span><span>Consumido <b>${miles(m.consumido)}</b> (${porc(c.pct)})</span>
            <span>Falta consumir <b>${miles(c.falta)}</b></span></div></div>`
        : `<div class="consumo"><div class="linea"><span>Sin asignación de consumo 2026</span>${m.consumido ? `<span>Consumido <b>${miles(m.consumido)}</b></span>` : ''}</div></div>`;
      const avisos = [
        c.atrasada ? `🔴 Consumo atrasado: lleva ${porc(c.pct)} y ya pasó el ${porc(esperado)} del año` : '',
        c.noAlcanza ? `⚠️ No alcanza: faltan consumir ${miles(c.falta)} y hay ${miles(m.disponible)} disponibles` : '',
        c.descuadre ? `⚠️ Las actividades suman ${miles(m.sumaActividades)} pero el total consumido dice ${miles(m.consumido)}` : ''
      ].filter(Boolean).map((t) => `<div class="alerta-txt">${t}</div>`).join('');
      return `<div class="item ${m.asignado ? (c.atrasada ? 'atrasada' : 'al-dia') : ''}">
        <div><div class="tipo">Munición</div><div class="nombre">${resaltar(m.nombre, palabras)}</div></div>
        <div class="total"><b>${miles(m.disponible)}</b><small>disponible</small></div>
        ${actividades ? `<div class="reparto">${actividades}</div>` : ''}
        ${barra}
        ${avisos}
        ${m.nota ? `<div class="nota-txt">📝 ${resaltar(m.nota, palabras)}</div>` : ''}
      </div>`;
    }).join('');
  }

  function alertasMunicion() {
    const d = estado.datos;
    const { esperado } = avanceDelAnio(d);
    const con = d.municiones.map((m) => ({ m, c: cuentasMunicion(m, esperado) }));
    const bloques = [];
    const noAlcanza = con.filter((x) => x.c.noAlcanza);
    if (noAlcanza.length) {
      bloques.push(`<div class="bloque-alerta grave"><h3>⚠️ No alcanza lo disponible para cumplir lo asignado (${noAlcanza.length})</h3>
        <p>Lo que falta consumir es más que la munición disponible: hay que pedirla o reprogramar.</p>
        ${tablaHTML({ titulos: ['Munición', 'Asignado', 'Falta consumir', 'Disponible'],
          filas: noAlcanza.map(({ m, c }) => [m.nombre, miles(m.asignado), miles(c.falta), miles(m.disponible)]) }, ['ASIGNADO', 'FALTA', 'DISPONIBLE'])}</div>`);
    }
    const atrasadas = con.filter((x) => x.c.atrasada);
    if (atrasadas.length) {
      bloques.push(`<div class="bloque-alerta"><h3>🔴 Consumo atrasado (${atrasadas.length})</h3>
        <p>Ya pasó el ${porc(esperado)} del año y esta munición lleva menos consumido que eso.</p>
        ${tablaHTML({ titulos: ['Munición', 'Asignado', 'Consumido', '% consumo', 'Falta consumir'],
          filas: atrasadas.sort((x, y) => x.c.pct - y.c.pct).map(({ m, c }) => [m.nombre, miles(m.asignado), miles(m.consumido), porc(c.pct), miles(c.falta)]) },
        ['ASIGNADO', 'CONSUMIDO', '% CONSUMO', 'FALTA'])}</div>`);
    }
    const descuadre = con.filter((x) => x.c.descuadre);
    if (descuadre.length) {
      bloques.push(`<div class="bloque-alerta grave"><h3>⚠️ Total consumido no cuadra con las actividades (${descuadre.length})</h3>
        <p>La suma de FOCO, FECOME, PEC, HBC y otros no es igual a la columna «Total consumido». Revisar en el Excel.</p>
        ${tablaHTML({ titulos: ['Munición', 'Suma de actividades', 'Total consumido'],
          filas: descuadre.map(({ m }) => [m.nombre, miles(m.sumaActividades), miles(m.consumido)]) }, ['SUMA', 'TOTAL'])}</div>`);
    }
    const conNota = d.municiones.filter((m) => m.nota);
    if (conNota.length) {
      bloques.push(`<div class="bloque-alerta"><h3>📝 Observaciones del S-4 (${conNota.length})</h3>
        ${tablaHTML({ titulos: ['Munición', 'Disponible', 'Observación'], filas: conNota.map((m) => [m.nombre, miles(m.disponible), m.nota]) }, ['DISPONIBLE'])}</div>`);
    }
    return { html: bloques.join(''), cuenta: noAlcanza.length + descuadre.length };
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
    const cifra = ([n, t]) => `<div class="cifra"><b>${n}</b><span>${t}</span></div>`;
    if (d.clase === 'vestuario') {
      // Cuánto hay en cada situación especial (columnas ISA, III CAT, BAJO RECIBO… si el Excel las trae)
      const enColumna = (u) => d.activos.reduce((s, a) => s + (a.reparto[u] || 0), 0);
      const especiales = d.unidades.filter((u) => /BAJO RECIBO|III CAT|ISA|GUARDIA/.test(normal(u)))
        .map((u) => [miles(enColumna(u)), 'en ' + u]);
      $('resumen').innerHTML = [
        [d.activos.length, 'artículos distintos'],
        [miles(total), 'unidades en total'],
        [d.unidades.length, 'unidades / destinos'],
        ...especiales
      ].map(cifra).join('');
      return;
    }
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
    ].map(cifra).join('');
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
        ${estado.datos.clase === 'vestuario' ? '' : `<dt>CIMAGE</dt><dd>${esc(a.cimage) || '—'}</dd>
        <dt>Revista</dt><dd>${a.revista ? REVISTA[a.revista].punto + ' ' + REVISTA[a.revista].texto : 'Sin marca en el Excel'}</dd>`}
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
      // Por tipo de activo; si el Excel tiene un solo tipo (vestuario), los artículos principales
      const porTipo = {};
      const unSoloTipo = new Set(d.activos.map((a) => a.tipo)).size < 2;
      suyos.forEach((a) => { const k = unSoloTipo ? a.nombre : a.tipo; porTipo[k] = (porTipo[k] || 0) + a.reparto[u]; });
      const lineas = Object.entries(porTipo).sort((x, y) => y[1] - x[1]).slice(0, unSoloTipo ? 6 : undefined)
        .map(([t, n]) => `<div class="linea-tipo"><span>${esc(t)}</span><b>${miles(n)}</b></div>`).join('');
      const total = suyos.reduce((s, a) => s + a.reparto[u], 0);
      const grupo = (d.grupos || {})[u];
      return `<button class="unidad ${suyos.length ? '' : 'vacia'}" data-unidad="${esc(u)}">
        ${grupo ? `<div class="grupo">${esc(grupo)}</div>` : ''}
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
    if (d.clase === 'municion') {
      const m = alertasMunicion();
      const viejaM = antiguedad(d).vieja ? 1 : 0;
      $('listaAlertas').innerHTML = cuadroFechas() + (m.html || '<p>Sin otras alertas. 👍</p>');
      $('nAlertas').textContent = viejaM + m.cuenta || '';
      return;
    }
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
    const vieja = antiguedad(d).vieja ? 1 : 0;
    $('listaAlertas').innerHTML = cuadroFechas() + (bloques.join('') || '<p>Sin otras alertas. 👍</p>');
    $('nAlertas').textContent = vieja + conNota.length + conDiferencia.length + sinReparto.length || '';
  }

  // ---------- Eventos ----------
  $('archivo').addEventListener('change', (e) => { cargarArchivo(e.target.files[0]); e.target.value = ''; });
  $('imprimir').addEventListener('click', () => window.print());
  ['busqueda', 'filtroUnidad', 'filtroRevista', 'filtroNota', 'orden'].forEach((id) => $(id).addEventListener('input', pintarLista));
  ['busquedaMun', 'filtroMun', 'ordenMun'].forEach((id) => $(id).addEventListener('input', pintarListaMunicion));
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
  $('libros').addEventListener('click', (e) => {
    const b = e.target.closest('button');
    if (b && b.dataset.libro !== estado.libro) cambiarLibro(b.dataset.libro);
  });

  // Arrastrar el archivo sobre la página
  const zona = $('zonaCarga');
  document.addEventListener('dragover', (e) => { e.preventDefault(); zona.classList.add('encima'); });
  document.addEventListener('dragleave', () => zona.classList.remove('encima'));
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    zona.classList.remove('encima');
    cargarArchivo(e.dataTransfer.files[0]);
  });

  for (const clase of Object.keys(CLAVES)) {
    const previos = recuperar(clase);
    if (previos && (previos.activos || previos.municiones)) estado.libros[clase] = previos;
  }
  let ultimo = 'activos';
  try { ultimo = localStorage.getItem(CLAVE_LIBRO) || 'activos'; } catch (e) { /* no importa */ }
  cambiarLibro(CLAVES[ultimo] ? ultimo : 'activos');
  window.ActivosESCINF = { mostrarDatos, cambiarLibro };

  // Instalable y sin internet (solo cuando se abre desde una dirección web, no como archivo)
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(() => {});
  }
})();
