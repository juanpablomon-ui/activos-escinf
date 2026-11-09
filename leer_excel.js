// Lee el Excel de distribución de activos bélicos (S-4) y lo deja como datos simples.
// Funciona en el navegador (window.LeerExcel) y en node (require) para las pruebas.
(function (raiz) {
  'use strict';

  // Colores de la leyenda del Excel (por si no se encuentra la leyenda en la hoja)
  const COLORES_REVISTA = { '92D050': 'por_revistar', 'FFC000': 'revistado', 'FFFF00': 'no_revistado' };
  const TIPO_SIN = 'SIN TIPO EN EL EXCEL';

  const limpiar = (t) => String(t == null ? '' : t).replace(/\s+/g, ' ').trim();
  const normal = (t) => limpiar(t).toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

  function numero(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v).replace(/\./g, '').replace(',', '.'));
    return isNaN(n) ? 0 : n;
  }

  function colorDe(celda) {
    const s = celda && celda.s;
    if (!s || s.patternType !== 'solid' || !s.fgColor || !s.fgColor.rgb) return null;
    return String(s.fgColor.rgb).slice(-6).toUpperCase();
  }

  // Busca la fila de títulos (la que dice NOMENCLATURA)
  function buscarTitulos(XLSX, hoja) {
    const rango = XLSX.utils.decode_range(hoja['!ref'] || 'A1');
    for (let f = rango.s.r; f <= Math.min(rango.e.r, rango.s.r + 40); f++) {
      for (let c = rango.s.c; c <= rango.e.c; c++) {
        const celda = hoja[XLSX.utils.encode_cell({ r: f, c })];
        if (celda && normal(celda.v) === 'NOMENCLATURA') return { fila: f, rango };
      }
    }
    return null;
  }

  // Lee la leyenda (POR REVISTAR / REVISTADO / NO REVISTADO) que está sobre la tabla
  function leerLeyenda(XLSX, hoja, filaTitulos, rango) {
    const leyenda = {};
    for (let f = rango.s.r; f < filaTitulos; f++) {
      for (let c = rango.s.c + 1; c <= Math.min(rango.e.c, 6); c++) {
        const t = normal((hoja[XLSX.utils.encode_cell({ r: f, c })] || {}).v);
        const color = colorDe(hoja[XLSX.utils.encode_cell({ r: f, c: c - 1 })]);
        if (!color) continue;
        if (t === 'POR REVISTAR') leyenda[color] = 'por_revistar';
        else if (t === 'REVISTADO') leyenda[color] = 'revistado';
        else if (t === 'NO REVISTADO') leyenda[color] = 'no_revistado';
      }
    }
    return Object.keys(leyenda).length ? leyenda : COLORES_REVISTA;
  }

  function leerDistribucion(XLSX, hoja) {
    const t = buscarTitulos(XLSX, hoja);
    if (!t) return null;
    const { fila, rango } = t;
    const celda = (f, c) => hoja[XLSX.utils.encode_cell({ r: f, c })];
    const leyenda = leerLeyenda(XLSX, hoja, fila, rango);

    // Ubicar columnas por su título
    const col = {};
    const unidades = [];
    // Observaciones: columna con título OBSERV… o columnas sin título después de las unidades
    const notas = [];
    // Las unidades van después de DIFERENCIA (activos bélicos) o, si no hay, después de CARGO (vestuario y equipo)
    let hayDiferencia = false;
    for (let c = rango.s.c; c <= rango.e.c; c++) if (normal((celda(fila, c) || {}).v).startsWith('DIFERENCIA')) hayDiferencia = true;
    let enUnidades = false;
    for (let c = rango.s.c; c <= rango.e.c; c++) {
      const titulo = limpiar((celda(fila, c) || {}).v);
      const n = normal(titulo);
      if (/OBSERV|NOTA|COMENTARIO/.test(n)) { notas.push(c); continue; }
      if (!n) { if (enUnidades) notas.push(c); continue; }
      if (/^IGUAL AL CARGO|^TOTAL/.test(n)) continue; // columna de control del Excel, no es unidad
      if (enUnidades) { unidades.push({ col: c, nombre: titulo }); continue; }
      if (n === 'N°' || n === 'Nº' || n === 'N.°' || n === 'NO' || n === 'N') col.n = c;
      else if (n === 'CIMAGE') col.cimage = c;
      else if (n.startsWith('TIPO')) col.tipo = c;
      else if (n === 'NOMENCLATURA') col.nombre = c;
      else if (n.startsWith('CARGO DISTRIBUIDO')) col.distribuido = c;
      else if (n === 'CARGO' || n === 'GARGO') { col.cargo = c; if (!hayDiferencia) enUnidades = true; }
      else if (n.startsWith('DIFERENCIA')) { col.diferencia = c; enUnidades = true; }
    }
    const clase = col.cimage != null || col.tipo != null ? 'activos' : 'vestuario';

    // Grupo de cada unidad (fila sobre los títulos, p. ej. «SECRET» sobre CBOA…CASA, «BTN» sobre COMP FUS)
    const grupos = {};
    const combinadas = hoja['!merges'] || [];
    for (const u of unidades) {
      let g = limpiar((celda(fila - 1, u.col) || {}).v);
      if (!g) {
        const m = combinadas.find((x) => x.s.r === fila - 1 && x.s.c <= u.col && u.col <= x.e.c);
        if (m) g = limpiar((celda(m.s.r, m.s.c) || {}).v);
      }
      if (g) grupos[u.nombre] = g;
    }
    // Título de la planilla (primer texto largo sobre la tabla)
    let titulo = '';
    for (let f = rango.s.r; f < fila && !titulo; f++) {
      for (let c = rango.s.c; c <= rango.e.c; c++) {
        const v = limpiar((celda(f, c) || {}).v);
        if (v.length > 25) { titulo = v; break; }
      }
    }

    const activos = [];
    for (let f = fila + 1; f <= rango.e.r; f++) {
      const nombre = limpiar((celda(f, col.nombre) || {}).v);
      if (!nombre) continue;
      const reparto = {};
      let suma = 0;
      for (const u of unidades) {
        const cant = numero((celda(f, u.col) || {}).v);
        if (cant) { reparto[u.nombre] = cant; suma += cant; }
      }
      let tipo = limpiar((celda(f, col.tipo) || {}).v);
      if (clase === 'vestuario') tipo = 'VESTUARIO Y EQUIPO';
      else if (!tipo || /^\?+$/.test(tipo)) tipo = TIPO_SIN;
      // El estado de revista va en el color de la fila (se mira primero la nomenclatura)
      let revista = null;
      for (const c of [col.nombre, col.n, col.cimage, col.tipo]) {
        const e = leyenda[colorDe(celda(f, c))];
        if (e) { revista = e; break; }
      }
      const nota = notas.map((c) => (celda(f, c) || {}).v)
        .filter((v) => typeof v === 'string' && limpiar(v)).map(limpiar).join(' · ');
      const cimage = (celda(f, col.cimage) || {});
      const cargo = col.cargo != null ? numero((celda(f, col.cargo) || {}).v) : suma;
      activos.push({
        n: limpiar((celda(f, col.n) || {}).v),
        cimage: limpiar(cimage.w != null && typeof cimage.v === 'number' ? cimage.v.toFixed(0) : cimage.v),
        tipo,
        nombre,
        total: suma,
        cargo,
        diferencia: suma - cargo,
        reparto,
        revista,
        nota,
        fila: f + 1
      });
    }
    return { clase, titulo, grupos, unidades: unidades.map((u) => u.nombre), activos };
  }

  // ---------- Cuadro de situación de munición (asignada por obligación de consumo) ----------
  // Títulos: «Municiones por calibre y efecto | Asignación de consumo 2026 | Actividades de consumo (FOCO, FECOME, PEC,
  // HBC, Otros en la fila de abajo) | Total consumido | % Consumo | MUNICIÓN DISPONIBLE | Observaciones»
  function leerMunicion(XLSX, hoja) {
    const rango = XLSX.utils.decode_range(hoja['!ref'] || 'A1');
    const celda = (f, c) => hoja[XLSX.utils.encode_cell({ r: f, c })];
    const texto = (f, c) => limpiar((celda(f, c) || {}).v);
    let fila = -1;
    for (let f = rango.s.r; f <= Math.min(rango.e.r, rango.s.r + 40) && fila < 0; f++) {
      for (let c = rango.s.c; c <= rango.e.c; c++) if (/DISPONIBLE/.test(normal(texto(f, c))) && /MUNICION/.test(normal(texto(f, c)))) { fila = f; break; }
    }
    if (fila < 0) return null;

    const col = { actividades: [] };
    for (let c = rango.s.c; c <= rango.e.c; c++) {
      const n = normal(texto(fila, c));
      if (/^MUNICION.*CALIBRE|^MUNICIONES/.test(n)) col.nombre = c;
      else if (n.startsWith('ASIGNACION')) col.asignado = c;
      else if (n.startsWith('TOTAL CONSUMIDO')) col.consumido = c;
      else if (n.includes('CONSUMO') && n.includes('%')) col.porcentaje = c;
      else if (n.includes('DISPONIBLE')) col.disponible = c;
      else if (n.startsWith('OBSERV')) col.obs = c;
    }
    if (col.nombre == null) col.nombre = rango.s.c;
    // Actividades: los subtítulos de la fila de abajo, entre la asignación y el total consumido
    const hasta = col.consumido != null ? col.consumido : rango.e.c;
    for (let c = (col.asignado != null ? col.asignado + 1 : col.nombre + 1); c < hasta; c++) {
      const t = texto(fila + 1, c);
      if (t) col.actividades.push({ col: c, nombre: t });
    }
    const inicio = col.actividades.length ? fila + 2 : fila + 1;
    const valor = (f, c) => (c == null ? 0 : numero(String((celda(f, c) || {}).v ?? '').replace(/^-$/, '')));

    const municiones = [];
    for (let f = inicio; f <= rango.e.r; f++) {
      const nombre = texto(f, col.nombre);
      if (!nombre) continue;
      // las filas de la firma (nombre, grado, cargo) no traen ningún número
      const conNumeros = [col.asignado, col.consumido, col.disponible].some((c) => c != null && typeof (celda(f, c) || {}).v === 'number');
      if (!conNumeros) continue;
      const actividades = {};
      let suma = 0;
      for (const a of col.actividades) { const n = valor(f, a.col); if (n) { actividades[a.nombre] = n; suma += n; } }
      municiones.push({
        nombre,
        asignado: valor(f, col.asignado),
        actividades,
        sumaActividades: suma,
        consumido: valor(f, col.consumido),
        disponible: valor(f, col.disponible),
        nota: texto(f, col.obs),
        fila: f + 1
      });
    }
    // Título, fecha de corte («… AL 30OCT») y año
    let titulo = '';
    let anio = '';
    for (let f = rango.s.r; f < fila; f++) {
      for (let c = rango.s.c; c <= rango.e.c; c++) {
        const t = texto(f, c);
        if (!titulo && /CUADRO/.test(normal(t))) titulo = t;
        if (/^ANO$/.test(normal(t).replace(/[^A-Z]/g, ''))) {
          for (let k = c + 1; k <= rango.e.c; k++) { const v = texto(f, k); if (/^\d{4}$/.test(v)) { anio = v; break; } }
        }
      }
    }
    const m = normal(titulo).match(/AL\s*(\d{1,2})\s*(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|SET|OCT|NOV|DIC)\s*(\d{4})?/);
    const corte = m ? new Date(Number(m[3] || anio || new Date().getFullYear()), MESES[m[2]], Number(m[1]), 12).toISOString() : null;
    return { titulo, corte, actividades: col.actividades.map((a) => a.nombre), municiones };
  }

  // Hojas simples (OAL y combustible): primera fila = títulos
  function leerTabla(XLSX, hoja) {
    const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: '', raw: false });
    const conDatos = filas.filter((f) => f.some((v) => limpiar(v)));
    if (!conDatos.length) return { titulos: [], filas: [] };
    // Montos con formato chileno ($1,000,000 → $1.000.000)
    const valor = (v) => (/^\$?\d{1,3}(,\d{3})+$/.test(limpiar(v)) ? limpiar(v).replace(/,/g, '.') : limpiar(v));
    return { titulos: conDatos[0].map(limpiar), filas: conDatos.slice(1).map((f) => f.map(valor)) };
  }

  // Fecha en el nombre del archivo, p. ej. «… 22OCT2026.xlsx» o «15 ENE 2026»
  const MESES = { ENE: 0, FEB: 1, MAR: 2, ABR: 3, MAY: 4, JUN: 5, JUL: 6, AGO: 7, SEP: 8, SET: 8, OCT: 9, NOV: 10, DIC: 11 };
  function fechaEnNombre(nombre) {
    const m = normal(nombre).match(/(\d{1,2})\s*(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|SET|OCT|NOV|DIC)\s*(\d{4})/);
    return m ? new Date(Number(m[3]), MESES[m[2]], Number(m[1]), 12) : null;
  }

  // ¿De cuándo es la información? 1) última vez que se guardó el Excel (dato interno del archivo),
  // 2) fecha escrita en el nombre, 3) fecha del archivo en el equipo, 4) el momento de la carga
  function fechaDeLosDatos(wb, archivo, modificado) {
    const props = wb.Props || {};
    const interna = props.ModifiedDate ? new Date(props.ModifiedDate) : null;
    if (interna && !isNaN(interna)) return { fecha: interna.toISOString(), origen: 'última modificación del Excel' };
    const enNombre = fechaEnNombre(archivo || '');
    if (enNombre) return { fecha: enNombre.toISOString(), origen: 'fecha en el nombre del archivo' };
    if (modificado) return { fecha: new Date(modificado).toISOString(), origen: 'fecha del archivo' };
    return { fecha: new Date().toISOString(), origen: 'día en que se cargó' };
  }

  function leer(XLSX, wb, archivo, modificado) {
    const fecha = () => { const f = fechaDeLosDatos(wb, archivo, modificado); return { fechaDatos: f.fecha, origenFecha: f.origen }; };
    // ¿Es el cuadro de munición?
    for (const n of wb.SheetNames) {
      const mun = leerMunicion(XLSX, wb.Sheets[n]);
      if (mun && mun.municiones.length) {
        // la información es de la fecha de corte del cuadro («… AL 30OCT»), aunque el Excel se haya guardado después
        const f = mun.corte ? { fechaDatos: mun.corte, origenFecha: 'fecha de corte del cuadro' } : fecha();
        return { clase: 'municion', archivo: archivo || '', cargado: new Date().toISOString(), ...f, ...mun };
      }
    }
    let dist = null;
    const nombres = wb.SheetNames.slice().sort((a, b) => (normal(b).includes('DISTRIB') ? 1 : 0) - (normal(a).includes('DISTRIB') ? 1 : 0));
    for (const n of nombres) {
      dist = leerDistribucion(XLSX, wb.Sheets[n]);
      if (dist) break;
    }
    if (!dist) throw new Error('No encontré la tabla de distribución (falta la columna NOMENCLATURA).');
    const hojaCon = (palabra) => wb.SheetNames.find((n) => normal(n).includes(palabra));
    const oal = hojaCon('OAL');
    const comb = hojaCon('COMBUSTIBLE');
    return {
      clase: dist.clase, // 'activos' (distribución de activos bélicos) o 'vestuario' (cuadro de vestuario y equipo)
      titulo: dist.titulo,
      grupos: dist.grupos,
      archivo: archivo || '',
      cargado: new Date().toISOString(),
      // fechaDatos / origenFecha: de cuándo es la información (para avisar si está vieja)
      ...(() => { const f = fechaDeLosDatos(wb, archivo, modificado); return { fechaDatos: f.fecha, origenFecha: f.origen }; })(),
      unidades: dist.unidades,
      activos: dist.activos,
      movimientos: oal ? leerTabla(XLSX, wb.Sheets[oal]) : { titulos: [], filas: [] },
      combustible: comb ? leerTabla(XLSX, wb.Sheets[comb]) : { titulos: [], filas: [] }
    };
  }

  const api = { leer, fechaEnNombre, TIPO_SIN };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else raiz.LeerExcel = api;
})(this);
