// Portal de Tecnicos Tigo: cada tecnico entra con su ID (ultimos 6 caracteres de su
// RUT, K -> 0) y ve SOLO sus propios indicadores: Repetido Reparado, Averias de
// Infancia, Productividad y NPS. Igual al portal de VTR desde el 14-10-2026 (ver el
// bloque "Desde el 14-10-2026" mas abajo): la pagina publicada no trae los datos;
// cada tecnico pide los suyos a Supabase con su ID.
//
// El RUT se usa UNICAMENTE en este script, en este computador, para cruzar los
// reportes y calcular el ID. El RUT nunca se guarda en ninguna pagina ni archivo.
//
// Lee los mismos CSV que el proyecto "Supervisor" desde ..\bbdd\ (no duplica datos).


const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const carpeta = __dirname;
const carpetaBbdd = path.join(carpeta, '..', 'bbdd');
const carpetaBaremos = path.join(carpeta, '..', 'baremosTigo');
const META_REINCIDENCIA = 0.04;
const META_INFANCIA = 0.025;
const META_PRODUCTIVIDAD = 5; // productos por dia trabajado (mas alto es mejor)

// Tecnicos desvinculados: no se les asigna login (no pueden entrar al
// Portal), pero su trabajo pasado se sigue contando en los promedios y
// rankings del equipo -- se excluyen recien al armar el login, no antes.
// Agregar el nombre completo exacto tal como aparece en los reportes (la
// comparacion es insensible a mayusculas/acentos).
const DESVINCULADOS = [
  'JOAN JOEL HERNANDEZ AMAYA',
  'LUCAS RENATO SERON VELASQUEZ',
];

function encontrarCsv(prefijoRegex) {
  if (!fs.existsSync(carpetaBbdd)) {
    throw new Error('No existe la carpeta bbdd en ' + path.join(carpeta, '..'));
  }
  const candidatos = fs
    .readdirSync(carpetaBbdd)
    .filter((f) => prefijoRegex.test(f))
    .map((f) => ({ nombre: f, mtime: fs.statSync(path.join(carpetaBbdd, f)).mtimeMs }))
    .sort((a, b) => b.mtime - a.mtime);
  if (candidatos.length === 0) {
    throw new Error('No se encontro ningun CSV que coincida con ' + prefijoRegex + ' en ' + carpetaBbdd);
  }
  return path.join(carpetaBbdd, candidatos[0].nombre);
}

function parseCsvLine(line) {
  const result = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (line[i + 1] === '"') { cur += '"'; i++; }
        else inQuotes = false;
      } else cur += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ';') { result.push(cur); cur = ''; }
      else cur += c;
    }
  }
  result.push(cur);
  return result;
}

function leerCsv(csvPath) {
  const raw = fs.readFileSync(csvPath, 'utf8');
  const lines = raw.split(/\r?\n/).filter((l) => l.length > 0);
  const headers = parseCsvLine(lines[0]);
  return lines.slice(1).map((l) => {
    const vals = parseCsvLine(l);
    const obj = {};
    headers.forEach((h, i) => (obj[h] = vals[i] !== undefined ? vals[i] : ''));
    return obj;
  });
}

const NOMBRES_MES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
function nombreMes(slug) {
  const [y, m] = slug.split('-');
  return NOMBRES_MES[parseInt(m, 10) - 1] + ' ' + y;
}

// El "mes" del informe se toma de partition_date del CSV (la fecha del snapshot
// de datos), no de la fecha del sistema, igual que en Supervisor.
function obtenerSlugDesdeCsv(rows) {
  const pd = (rows[0] && rows[0]['partition_date']) || '';
  if (/^\d{4}-\d{2}/.test(pd)) return pd.slice(0, 7);
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

function listarMesesExistentes(carpeta) {
  const re = /^index_(\d{4}-\d{2})\.html$/;
  return fs.readdirSync(carpeta)
    .map((f) => f.match(re))
    .filter(Boolean)
    .map((m) => m[1]);
}

// Solo se usa para calcular la clave de 4 digitos; el resultado (soloDigitos)
// nunca identifica por si solo a una persona. Si el digito verificador es
// "K" se reemplaza por "0" para que la clave sea siempre numerica (el campo
// de clave del portal usa teclado numerico en el celular, que no tiene K).
function ultimos4DigitosRut(r) {
  const limpio = (r || '').toString().trim().toUpperCase().replace(/\./g, '').replace(/-/g, '');
  return limpio.slice(-4).replace(/K$/, '0');
}

function normalizarTexto(s) {
  return (s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
}

// Nombre chileno tipico: Nombre1 [Nombre2] Apellido1 [Apellido2]
function primerNombreApellido(nombreCompleto) {
  const partes = (nombreCompleto || '').trim().split(/\s+/).filter(Boolean);
  const nombre = partes[0] || '';
  const apellido = partes.length >= 3 ? partes[2] : (partes[1] || '');
  return { nombre, apellido };
}

function loginKey(nombreCompleto) {
  const { nombre, apellido } = primerNombreApellido(nombreCompleto);
  return normalizarTexto(nombre + ' ' + apellido);
}

// Clave interna de agrupacion (solo en memoria, nunca se publica): el RUT
// completo normalizado. Se usa para que dos filas de la MISMA persona nunca
// se separen, y para poder detectar cuando dos personas DISTINTAS compartirian
// el mismo login (mismo nombre corto).
function rutInterno(r) {
  return (r || '').toString().trim().toUpperCase().replace(/\./g, '').replace(/-/g, '');
}

// Lee el DATA embebido de un mes ya archivado (index.html o index_YYYY-MM.html)
// para poder comparar el mes actual contra el anterior. Si el archivo no
// existe (primer mes del historial) o no se puede parsear, retorna null y la
// comparacion simplemente no se muestra -- no rompe la generacion normal.
function cargarDatosDeArchivo(nombreArchivo) {
  const p = path.join(carpeta, nombreArchivo);
  if (!fs.existsSync(p)) return null;
  try {
    const html = fs.readFileSync(p, 'utf8');
    const m = html.match(/const DATA = (\{[\s\S]*?\});\r?\n\r?\nfunction normalizarTexto/);
    if (!m) return null;
    return JSON.parse(m[1]);
  } catch (err) {
    console.log('AVISO: no se pudo leer el mes anterior desde ' + nombreArchivo + ' (' + err.message + ').');
    return null;
  }
}

// Lee el NPS por tecnico que exporta el informe NPS (repo/proyecto aparte,
// ver actualizar_informe.ps1). Esa BBDD no trae RUT, asi que el cruce con
// cada persona se hace mas abajo por nombre corto (mismo criterio del login).
function cargarNpsTecnicos() {
  const npsPath = path.join(carpetaBbdd, 'nps-tecnicos.json');
  if (!fs.existsSync(npsPath)) {
    console.log('AVISO: no se encontro nps-tecnicos.json en bbdd -- el portal se genera sin datos de NPS.');
    return null;
  }
  try {
    return JSON.parse(fs.readFileSync(npsPath, 'utf8'));
  } catch (err) {
    console.log('AVISO: no se pudo leer nps-tecnicos.json (' + err.message + ') -- el portal se genera sin datos de NPS.');
    return null;
  }
}

// Lee la productividad de cada tecnico desde el INF del mes mas reciente
// (INF-09 = Septiembre, INF-10 = Octubre...; produccion de Punta Arenas +
// Coyhaique) que copiamos a ../baremosTigo con copiar_baremos.bat.
// Cruce por RUT. Suma productos de Instala y de Repara por separado y cuenta
// los dias distintos trabajados. Devuelve { mapa por RUT normalizado, periodo }.
async function cargarProductividad() {
  if (!fs.existsSync(carpetaBaremos)) {
    console.log('AVISO: no existe la carpeta baremosTigo -- el portal se genera sin productividad.');
    return null;
  }
  const MESES_INF = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const claveInf = (f) => {
    const m = f.match(/(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s+(\d{4})/i);
    return m ? Number(m[2]) * 100 + MESES_INF.indexOf(m[1].toLowerCase()) + 1 : 0;
  };
  const cand = fs.readdirSync(carpetaBaremos)
    .filter((f) => /^INF-\d+.*\.xlsx$/i.test(f) && !f.startsWith('~$'))
    .map((f) => ({ f, k: claveInf(f), m: fs.statSync(path.join(carpetaBaremos, f)).mtimeMs }))
    .sort((a, b) => b.k - a.k || b.m - a.m)[0];
  if (!cand) {
    console.log('AVISO: no se encontro ningun INF-*.xlsx en baremosTigo -- el portal se genera sin productividad.');
    return null;
  }
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path.join(carpetaBaremos, cand.f));
    const ws = wb.getWorksheet('BASE_MES_ PUNTA_ARENAS');
    if (!ws) throw new Error('no se encontro la hoja "BASE_MES_ PUNTA_ARENAS"');

    // Columnas: 1 Folio | 5 Nro Productos | 6 RUT | 10 Dia | 14 Tipo (Instala/Repara)
    const porRut = {};
    for (let i = 2; i <= ws.rowCount; i++) {
      const row = ws.getRow(i);
      if (!row.getCell(1).value) continue;
      const rut = rutInterno(row.getCell(6).value);
      if (!rut) continue;
      const prod = Number(row.getCell(5).value) || 0;
      const tipo = (row.getCell(14).value || '').toString().trim();
      const dia = Number(row.getCell(10).value);
      if (!porRut[rut]) porRut[rut] = { instala: 0, repara: 0, dias: new Set() };
      if (tipo === 'Instala') porRut[rut].instala += prod;
      else if (tipo === 'Repara') porRut[rut].repara += prod;
      if (Number.isFinite(dia)) porRut[rut].dias.add(dia);
    }

    const mapa = {};
    Object.entries(porRut).forEach(([rut, x]) => {
      const total = +(x.instala + x.repara).toFixed(2);
      const dias = x.dias.size;
      mapa[rut] = {
        instala: +x.instala.toFixed(2),
        repara: +x.repara.toFixed(2),
        total,
        dias,
        prodDia: dias ? +(total / dias).toFixed(2) : null,
      };
    });

    // Periodo: se saca del nombre del archivo ("... Septiembre 2026.xlsx").
    const m = cand.f.match(/(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\s+(\d{4})/i);
    const periodo = m ? (m[1] + ' ' + m[2]) : null;
    console.log('Productividad cargada desde', cand.f, '(' + Object.keys(mapa).length + ' tecnicos)');
    return { mapa, periodo };
  } catch (err) {
    console.log('AVISO: no se pudo leer la productividad (' + err.message + ') -- el portal se genera sin productividad.');
    return null;
  }
}

// Cruza el NPS (agregado por tecnico+zona) con cada persona (identificada por
// RUT) usando el mismo nombre corto del login. Si el tecnico trabajo en mas
// de una zona ese mes (aparece en mas de un grupo tecnico+zona), se combinan
// todas sus encuestas en un solo NPS en vez de elegir una sola zona.
function asignarNps(personas, npsData) {
  if (!npsData || !Array.isArray(npsData.tecnicos)) return;
  const porLogin = {};
  npsData.tecnicos.forEach((t) => {
    const key = loginKey(t.tecnico);
    if (!porLogin[key]) porLogin[key] = [];
    porLogin[key].push(t);
  });

  personas.forEach((p) => {
    const candidatos = porLogin[loginKey(p.nombre)];
    if (!candidatos || candidatos.length === 0) { p.nps = null; return; }
    const total = candidatos.reduce((a, c) => a + c.total, 0);
    const promotores = candidatos.reduce((a, c) => a + c.P, 0);
    const neutros = candidatos.reduce((a, c) => a + c.N, 0);
    const detractores = candidatos.reduce((a, c) => a + c.D, 0);
    const zonas = [...new Set(candidatos.map((c) => c.zona))];
    p.nps = {
      nps: total ? +(((promotores - detractores) / total) * 100).toFixed(1) : null,
      total, promotores, neutros, detractores,
      zona: zonas.join(' y '),
    };
  });
}

function mediana(arr) {
  if (arr.length === 0) return null;
  const s = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 === 0 ? (s[mid - 1] + s[mid]) / 2 : s[mid];
}
function promedio(arr) {
  if (arr.length === 0) return null;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

// ---------------- Repetido Reparado ----------------
function analizarReincidencias(csvPath) {
  // Mismos criterios que Supervisor: agencia SUR, tecnologia FO, empresa
  // COBRA, sin tecnicos "PTA", y sin ciertas claves de cierre no validas
  // (derivadas de rdy_prd_cat_nivel_2_sol).
  const AGENCIAS_VALIDAS = ['COYHAIQUE', 'PUNTA ARENAS'];
  const CLAVES_CIERRE_EXCLUIDAS = ['913', 'F20', 'F23', '561', 'A12', 'B01', 'B36', 'I50'];
  function tecnologiaFila(r) {
    return (r['toa_piv_tecnologia'] === 'FO' || r['toa_piv_descripcion'] === 'FO') ? 'FO' : r['toa_piv_tecnologia'];
  }
  function claveCierreFila(r) {
    return (r['rdy_prd_cat_nivel_2_sol'] || '').split('_')[0];
  }
  const rows = leerCsv(csvPath).filter((r) =>
    AGENCIAS_VALIDAS.includes((r['toa_piv_agencia'] || '').trim()) &&
    !CLAVES_CIERRE_EXCLUIDAS.includes(claveCierreFila(r)) &&
    !(r['toa_piv_nombre_tecnico'] || '').includes('PTA') &&
    tecnologiaFila(r) === 'FO' &&
    (r['toa_piv_empresa'] || '').trim() === 'COBRA'
  );
  const reit = rows.filter((r) => (r['rdy_prd_tiene_reitero_30d'] || '').trim() === '1');
  const reitFolios = new Set(reit.map((r) => r['toa_piv_folio_toa']));

  const porRut = {};
  rows.forEach((r) => {
    const nombre = (r['toa_piv_nombre_tecnico'] || '').trim();
    const rut = rutInterno(r['toa_piv_rut_tecnico']);
    if (!nombre || !rut) return;
    if (!porRut[rut]) {
      porRut[rut] = {
        rut, nombre, agencia: (r['toa_piv_agencia'] || '').trim(), clave: ultimos4DigitosRut(r['toa_piv_rut_tecnico']),
        total: 0, reincidencias: 0, dias: [], causas: {}, casos: [],
      };
    }
    porRut[rut].total += 1;
    if (reitFolios.has(r['toa_piv_folio_toa'])) {
      porRut[rut].reincidencias += 1;
      const dias = Number(r['rdy_prd_q_dias_reitero']);
      if (!Number.isNaN(dias)) porRut[rut].dias.push(dias);
      const causa = (r['toa_piv_causa'] || '').trim() || '(sin dato)';
      porRut[rut].causas[causa] = (porRut[rut].causas[causa] || 0) + 1;
      // Las dos peticiones (pedido del usuario 14-10-2026): la reparacion del tecnico y la
      // reclamacion que vino despues. Solo folio/fecha/causa/dias -- SIN direccion ni
      // telefono del cliente (existen en el CSV pero no se publican, es info de terceros).
      porRut[rut].casos.push({
        folio: (r['toa_piv_folio_toa'] || '').trim(),
        fecha: (r['toa_piv_fecha_ingreso'] || '').trim().slice(0, 10),
        causa,
        folio2: (r['rdy_prd_reiterado'] || '').trim(),
        fecha2: (r['rdy_prd_fecha_ingreso'] || '').trim().slice(0, 10),
        causa2: (r['rdy_prd_sintoma'] || '').split(' - ')[0].trim(),
        dias: !Number.isNaN(dias) ? dias : null,
      });
    }
  });

  const byFolio = {};
  rows.forEach((r) => { byFolio[r['toa_piv_folio_toa']] = r; });
  let mismoTotal = {}, mismoSi = {};
  reit.forEach((r) => {
    const rut = rutInterno(r['toa_piv_rut_tecnico']);
    if (!rut || !porRut[rut]) return;
    const r2 = byFolio[r['rdy_prd_reiterado']];
    if (!r2) return;
    mismoTotal[rut] = (mismoTotal[rut] || 0) + 1;
    if (r2['toa_piv_nombre_tecnico'] === r['toa_piv_nombre_tecnico']) mismoSi[rut] = (mismoSi[rut] || 0) + 1;
  });
  Object.keys(porRut).forEach((k) => {
    porRut[k].mismoTotal = mismoTotal[k] || 0;
    porRut[k].mismoSi = mismoSi[k] || 0;
  });

  return terminarAnalisis(porRut);
}

// ---------------- Averias de Infancia ----------------
function analizarInfancia(csvPath) {
  // Mismos criterios que Supervisor: solo agencias de la zona SUR, solo filas
  // de la empresa COBRA, y se descartan tecnicos "PTA" (placeholders que han
  // aparecido sueltos en el archivo origen).
  const AGENCIAS_VALIDAS = ['COYHAIQUE', 'PUNTA ARENAS'];
  const rows = leerCsv(csvPath).filter((r) =>
    AGENCIAS_VALIDAS.includes((r['toa_xa_original_agency'] || '').trim()) &&
    (r['toa_xr_company_name'] || '').trim().toUpperCase() === 'COBRA' &&
    !(r['toa_provider_name'] || '').includes('PTA')
  );
  // Instalaciones (Alta), traslados (T) y tipo B; se excluyen ademas ciertas
  // claves de cierre de la reparacion de infancia que no representan un caso valido.
  const CLAVES_CIERRE_EXCLUIDAS = ['F20', 'F23', '913', '561', 'A12', 'B01', 'B36', 'I50'];
  const instalaciones = rows.filter((r) =>
    ['A', 'T', 'B'].includes((r['vpi_tipo_trabajo_producto'] || '').trim()) &&
    !CLAVES_CIERRE_EXCLUIDAS.includes((r['rmdy_clave_cierre'] || '').trim())
  );

  const porRut = {};
  instalaciones.forEach((r) => {
    const nombre = (r['toa_provider_name'] || '').trim();
    const rut = rutInterno(r['toa_provider_external_id']);
    if (!nombre || !rut) return;
    if (!porRut[rut]) {
      porRut[rut] = {
        rut, nombre, agencia: (r['toa_xa_original_agency'] || '').trim(), clave: ultimos4DigitosRut(r['toa_provider_external_id']),
        total: 0, reincidencias: 0, dias: [], causas: {}, casos: [], mismoTotal: 0, mismoSi: 0,
      };
    }
    porRut[rut].total += 1;
    if ((r['infancia'] || '').trim() === '1') {
      porRut[rut].reincidencias += 1;
      const dias = Number(r['q_dias_infancia']);
      if (!Number.isNaN(dias)) porRut[rut].dias.push(dias);
      const causa = (r['rmdy_causa'] || '').trim() || '(sin dato)';
      porRut[rut].causas[causa] = (porRut[rut].causas[causa] || 0) + 1;
      // Las dos peticiones (pedido del usuario 14-10-2026): la instalacion del tecnico y la
      // reparacion que vino despues. Solo folio/fecha/causa/dias -- SIN direccion ni
      // telefono del cliente (existen en el CSV pero no se publican, es info de terceros).
      porRut[rut].casos.push({
        folio: (r['toa_appt_number'] || '').trim(),
        fecha: (r['fin_actividad'] || r['toa_eta'] || '').trim().slice(0, 10),
        causa,
        folio2: (r['rmdy_folio_repara'] || '').trim(),
        fecha2: (r['rmdy_fecha_creacion_repara'] || '').trim().slice(0, 10),
        causa2: causa.split(' - ')[0].trim(),
        dias: !Number.isNaN(dias) ? dias : null,
      });
      porRut[rut].mismoTotal += 1;
      if (r['rmdy_nombre_tecnico'] === r['toa_provider_name']) porRut[rut].mismoSi += 1;
    }
  });

  return terminarAnalisis(porRut);
}

function terminarAnalisis(porRut) {
  const lista = Object.values(porRut).map((t) => {
    const tasa = t.total ? t.reincidencias / t.total : 0;
    const causaFrecuente = Object.entries(t.causas).sort((a, b) => b[1] - a[1])[0];
    return {
      rut: t.rut, nombre: t.nombre, agencia: t.agencia, clave: t.clave,
      total: t.total, reincidencias: t.reincidencias, tasa,
      diasPromedio: t.dias.length ? +promedio(t.dias).toFixed(1) : null,
      diasMediana: t.dias.length ? mediana(t.dias) : null,
      causaFrecuente: causaFrecuente ? causaFrecuente[0] : null,
      causaFrecuenteCasos: causaFrecuente ? causaFrecuente[1] : 0,
      casos: (t.casos || []).slice().sort((a, b) => (b.fecha || '').localeCompare(a.fecha || '')),
      mismoTotal: t.mismoTotal, mismoSi: t.mismoSi,
      mismoPct: t.mismoTotal ? +((t.mismoSi / t.mismoTotal) * 100).toFixed(1) : null,
    };
  });

  const elegibles = lista.filter((t) => t.total >= 10).sort((a, b) => a.tasa - b.tasa);
  elegibles.forEach((t, i) => { t.ranking = i + 1; t.rankingTotal = elegibles.length; });

  const promedioEquipo = elegibles.length ? promedio(elegibles.map((t) => t.tasa)) : null;

  const mapa = {};
  lista.forEach((t) => { mapa[t.rut] = t; });
  return { mapa, promedioEquipo, totalTecnicos: lista.length };
}


// =====================================================================
// Desde el 14-10-2026 el portal es igual al de VTR (pedido del usuario; ver
// C:\Bases_Vtr\portal-tecnicos):
// - Cada tecnico entra solo con su ID: ultimos 6 caracteres del RUT (K -> 0).
// - index.html (publico) NO trae los datos: cada tecnico pide los suyos a Supabase
//   con su ID (rpc datos_tecnico, proyecto del Portal Supervisor). La copia completa
//   queda en privado/index.html (no se sube a GitHub): la lee el Portal Supervisor
//   y la carga a Supabase subir_datos_tecnicos.js.
// - El historial mes a mes vive en privado/meses/AAAA-MM.json. Ya no se publican
//   las paginas index_AAAA-MM.html ni meses.json (traian los datos de todos).
//
//   node generar_portal.js          -> index.html + privado/ (lo publica la tarea de las 11:00)
//   node generar_portal.js prueba   -> todo en prueba/ (no se publica ni guarda historial)
// =====================================================================
const MODO = process.argv[2] === 'prueba' ? 'prueba' : 'publicar';
const SALIDA = MODO === 'prueba' ? path.join(carpeta, 'prueba') : carpeta;
const PRIVADO = path.join(SALIDA, 'privado');
const HISTORIAL = path.join(carpeta, 'privado', 'meses');
const DIR_SUPERVISOR = 'C:\\Bases_Tigo\\portal-supervisor';
const ASSETS = ['logo-cobra.png', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png', 'manifest.json', 'sw.js', 'fondo-tecnicos.jpg', 'logo-tigo.png'];

// ID del tecnico: ultimos 6 caracteres del RUT, con K -> 0 (igual que VTR).
function idDesdeRut(r) {
  let s = rutInterno(r).slice(-6);
  if (s.endsWith('K')) s = s.slice(0, -1) + '0';
  return s;
}

// Direccion y llave PUBLICA de Supabase (proyecto del Portal Supervisor).
function leerEnvSupervisor() {
  const env = {};
  const p = path.join(DIR_SUPERVISOR, '.env.local');
  if (!fs.existsSync(p)) return env;
  for (const linea of fs.readFileSync(p, 'utf8').split(/\r?\n/)) {
    const m = linea.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*?)\s*$/);
    if (m) env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
  return env;
}

// Supervisor de cada tecnico (Maestros\Certificacion Cobra.xlsx, hoja NEW_BBDD), con
// las mismas funciones del Portal Supervisor. Si falla, el portal sale sin supervisor.
async function cargarSupervisores() {
  try {
    const comun = require(path.join(DIR_SUPERVISOR, 'lib', 'comun.js'));
    const sups = await comun.leerSupervisores();
    return (nombre) => comun.buscar(sups, nombre) || null;
  } catch (err) {
    console.log('AVISO: no se pudieron leer los supervisores (' + err.message + ') -- el portal sale sin supervisor.');
    return () => null;
  }
}

// ---------------- Historial mes a mes (privado) ----------------
function leerHistorial() {
  const hist = {};
  if (!fs.existsSync(HISTORIAL)) return hist;
  for (const f of fs.readdirSync(HISTORIAL)) {
    const m = f.match(/^(\d{4}-\d{2})\.json$/);
    if (!m) continue;
    try { hist[m[1]] = JSON.parse(fs.readFileSync(path.join(HISTORIAL, f), 'utf8')); } catch (e) { /* archivo danado: se ignora */ }
  }
  return hist;
}
function guardarFoto(foto) {
  fs.mkdirSync(HISTORIAL, { recursive: true });
  fs.writeFileSync(path.join(HISTORIAL, foto.slug + '.json'), JSON.stringify(foto), 'utf8');
}
// Las paginas publicas de meses anteriores (index_AAAA-MM.html) traian los datos por
// nombre corto. Se pasan una sola vez al historial privado, por ID.
function migrarMesesAntiguos(idPorLogin, hist) {
  const origenes = [carpeta, path.join(carpeta, 'privado', 'antiguos')];
  for (const dir of origenes) {
    if (!fs.existsSync(dir)) continue;
    for (const f of fs.readdirSync(dir)) {
      const m = f.match(/^index_(\d{4}-\d{2})\.html$/);
      if (!m || hist[m[1]]) continue;
      const p = path.join(dir, f);
      let d = null;
      try {
        const html = fs.readFileSync(p, 'utf8');
        const x = html.match(/const DATA = (\{[\s\S]*?\});\r?\n\r?\nfunction normalizarTexto/);
        d = x ? JSON.parse(x[1]) : null;
      } catch (e) { d = null; }
      if (!d || !d.tecnicos) continue;
      const foto = { slug: m[1], label: nombreMes(m[1]), promedioRep: d.promedioEquipoReincidencias, promedioInf: d.promedioEquipoInfancia, tecnicos: {} };
      for (const [login, t] of Object.entries(d.tecnicos)) {
        const id = idPorLogin[normalizarTexto(login)];
        if (!id) continue;
        foto.tecnicos[id] = {
          rep: t.reincidencias && t.reincidencias.total ? t.reincidencias.tasa : null,
          inf: t.infancia && t.infancia.total ? t.infancia.tasa : null,
        };
      }
      hist[m[1]] = foto;
      guardarFoto(foto);
      console.log('Historial: ' + f + ' pasado a privado/meses/' + m[1] + '.json (' + Object.keys(foto.tecnicos).length + ' tecnicos)');
    }
  }
}

// Posicion de cada tecnico en el equipo por indicador (igual que VTR):
// t.ranking.rep = { puesto, de, promedio }. Empates comparten puesto.
// Repetido e Infancia: solo quienes tienen al menos 10 casos (como el promedio del equipo).
function agregarRanking(lista) {
  const IND = {
    rep: { valor: (t) => (t.reincidencias && t.reincidencias.total >= 10 ? t.reincidencias.tasa : null), menor: true, dec: 10 },
    inf: { valor: (t) => (t.infancia && t.infancia.total >= 10 ? t.infancia.tasa : null), menor: true, dec: 10 },
    prod: { valor: (t) => (t.productividad && t.productividad.prodDia != null ? t.productividad.prodDia : null), menor: false, dec: 100 },
    nps: { valor: (t) => (t.nps && t.nps.nps != null ? t.nps.nps : null), menor: false, dec: 10 },
  };
  for (const t of lista) t.ranking = {};
  for (const [clave, ind] of Object.entries(IND)) {
    const valores = lista.map(ind.valor).filter((v) => v != null);
    if (!valores.length) continue;
    const promedio = Math.round((valores.reduce((a, b) => a + b, 0) / valores.length) * ind.dec) / ind.dec;
    for (const t of lista) {
      const v = ind.valor(t);
      if (v == null) continue;
      const mejores = valores.filter((x) => (ind.menor ? x < v : x > v)).length;
      t.ranking[clave] = { puesto: mejores + 1, de: valores.length, promedio };
    }
  }
}

function avisoVisible(texto) {
  try { require('child_process').execSync('msg "%USERNAME%" "' + texto.replace(/"/g, "'") + '"', { shell: 'cmd.exe', stdio: 'ignore' }); } catch (e) { /* sin msg */ }
}

async function main() {
  const csvReincidencias = encontrarCsv(/^p23-averias-reiteradas.*COBRA.*\.csv$/i);
  const csvInfancia = encontrarCsv(/^p22-Averias-infancia.*COBRA.*\.csv$/i);
  console.log('Modo:', MODO);
  console.log('CSV Repetido Reparado:', csvReincidencias);
  console.log('CSV Averias de Infancia:', csvInfancia);

  const periodoSlug = obtenerSlugDesdeCsv(leerCsv(csvReincidencias));
  const reincidencias = analizarReincidencias(csvReincidencias);
  const infancia = analizarInfancia(csvInfancia);

  // Fusion por RUT interno (nunca por nombre): dos personas distintas jamas se mezclan.
  const rutsTodos = new Set([...Object.keys(reincidencias.mapa), ...Object.keys(infancia.mapa)]);
  console.log('Tecnicos con datos:', rutsTodos.size);
  const prodData = await cargarProductividad();
  const prodMapa = prodData ? prodData.mapa : {};

  const personas = [...rutsTodos].map((rut) => {
    const r = reincidencias.mapa[rut];
    const i = infancia.mapa[rut];
    return {
      rut,
      nombre: (r && r.nombre) || (i && i.nombre) || '',
      agencia: (r && r.agencia) || (i && i.agencia) || '',
      reincidencias: r ? {
        total: r.total, reincidencias: r.reincidencias, tasa: +(r.tasa * 100).toFixed(1),
        diasPromedio: r.diasPromedio, causaFrecuente: r.causaFrecuente, causaFrecuenteCasos: r.causaFrecuenteCasos,
        mismoPct: r.mismoPct, casos: r.casos || [], tasaAnterior: null,
      } : null,
      infancia: i ? {
        total: i.total, reincidencias: i.reincidencias, tasa: +(i.tasa * 100).toFixed(1),
        diasPromedio: i.diasPromedio, causaFrecuente: i.causaFrecuente, causaFrecuenteCasos: i.causaFrecuenteCasos,
        mismoPct: i.mismoPct, casos: i.casos || [], tasaAnterior: null,
      } : null,
      productividad: prodMapa[rut] || null,
    };
  });
  const npsData = cargarNpsTecnicos();
  asignarNps(personas, npsData);

  // Desvinculados: no entran al portal (su trabajo pasado sigue contando en los promedios).
  const vigentes = personas.filter((p) => !DESVINCULADOS.includes(normalizarTexto(p.nombre)));
  const fuera = personas.length - vigentes.length;
  if (fuera) console.log('Tecnicos desvinculados excluidos del portal:', fuera);

  // ID de cada tecnico. Si dos personas DISTINTAS quedan con el mismo ID, ninguna de
  // las dos queda habilitada (nadie debe ver los datos de otro) y se avisa.
  const porId = {};
  vigentes.forEach((p) => { p.id = idDesdeRut(p.rut); (porId[p.id] = porId[p.id] || []).push(p); });
  const colisiones = Object.values(porId).filter((g) => g.length > 1);
  if (colisiones.length) {
    const nombres = colisiones.map((g) => g.map((p) => p.nombre).join(' / ')).join('; ');
    console.log('\n==============================================================');
    console.log('  ALERTA: tecnicos DISTINTOS con el mismo ID (ultimos 6 del RUT).');
    console.log('  No quedan habilitados hasta asignarles un ID distinto: ' + nombres);
    console.log('==============================================================\n');
    avisoVisible('Portal Tecnicos: dos tecnicos quedaron con el mismo ID y no pueden entrar: ' + nombres + '. Revisa log_reportes.txt');
  }
  const habilitados = vigentes.filter((p) => porId[p.id].length === 1);

  // Historial: meses anteriores (privado/meses) + paso unico de las paginas antiguas.
  const idPorLogin = {};
  habilitados.forEach((p) => { idPorLogin[loginKey(p.nombre)] = p.id; });
  const hist = leerHistorial();
  migrarMesesAntiguos(idPorLogin, hist);
  const mesesPrevios = Object.keys(hist).filter((s) => s < periodoSlug).sort().slice(-11);
  const mesAnterior = mesesPrevios.length ? mesesPrevios[mesesPrevios.length - 1] : null;
  const labelMesAnterior = mesAnterior ? nombreMes(mesAnterior) : null;
  const meses = [...mesesPrevios, periodoSlug];
  const promRep = reincidencias.promedioEquipo != null ? +(reincidencias.promedioEquipo * 100).toFixed(1) : null;
  const promInf = infancia.promedioEquipo != null ? +(infancia.promedioEquipo * 100).toFixed(1) : null;

  const supervisorDe = await cargarSupervisores();
  agregarRanking(habilitados);

  const tecnicos = {};
  for (const p of habilitados) {
    const previo = mesAnterior && hist[mesAnterior].tecnicos[p.id];
    if (previo && p.reincidencias) p.reincidencias.tasaAnterior = previo.rep;
    if (previo && p.infancia) p.infancia.tasaAnterior = previo.inf;
    const enMes = (s, k) => (hist[s] && hist[s].tecnicos[p.id] ? hist[s].tecnicos[p.id][k] : null);
    p.evolutivo = {
      meses,
      rep: meses.map((s) => (s === periodoSlug ? (p.reincidencias && p.reincidencias.total ? p.reincidencias.tasa : null) : enMes(s, 'rep'))),
      inf: meses.map((s) => (s === periodoSlug ? (p.infancia && p.infancia.total ? p.infancia.tasa : null) : enMes(s, 'inf'))),
      equipo_rep: meses.map((s) => (s === periodoSlug ? promRep : (hist[s].promedioRep != null ? hist[s].promedioRep : null))),
      equipo_inf: meses.map((s) => (s === periodoSlug ? promInf : (hist[s].promedioInf != null ? hist[s].promedioInf : null))),
    };
    // IMPORTANTE: el RUT y la clave antigua de 4 digitos nunca salen de este script.
    const { rut, id, clave, ...publicable } = p;
    tecnicos[p.id] = { ...publicable, supervisor: supervisorDe(p.nombre) };
  }
  if (!Object.keys(tecnicos).length) throw new Error('no quedo ningun tecnico habilitado; no se genera el portal');

  const DATA = {
    generadoEl: new Date().toLocaleString('es-CL'),
    periodoSlug,
    periodoLabel: nombreMes(periodoSlug),
    // archivos: lo usa el Portal Supervisor (etiqueta del mes)
    archivos: meses.slice().reverse().map((s) => ({ slug: s, label: nombreMes(s) })),
    metaReincidencias: +(META_REINCIDENCIA * 100).toFixed(0),
    metaInfancia: +(META_INFANCIA * 100).toFixed(1),
    promedioEquipoReincidencias: promRep,
    promedioEquipoInfancia: promInf,
    metaNps: npsData ? npsData.meta : null,
    periodoNps: npsData ? npsData.periodo : null,
    metaProductividad: META_PRODUCTIVIDAD,
    periodoProductividad: prodData ? prodData.periodo : null,
    labelMesAnterior,
    totalTecnicos: Object.keys(tecnicos).length,
    tecnicos,
  };

  const env = leerEnvSupervisor();
  if (!env.SUPABASE_URL || !env.SUPABASE_ANON_KEY) throw new Error('falta SUPABASE_URL / SUPABASE_ANON_KEY en ' + path.join(DIR_SUPERVISOR, '.env.local'));
  const plantilla = fs.readFileSync(path.join(carpeta, 'template.html'), 'utf8');
  const estilo = fs.readFileSync(path.join(DIR_SUPERVISOR, 'plantillas', 'estilo-academia.css'), 'utf8');
  const aviso = MODO === 'prueba'
    ? '<div style="background:#fdf1d8;border-left:3px solid #f0a500;color:#5c4300;font:13px/1.4 Helvetica,Arial,sans-serif;padding:10px 16px;margin:0 0 22px;"><b>Versi&oacute;n de prueba.</b> No est&aacute; publicada.</div>'
    : '';
  const armar = (data) => plantilla
    .replace('/*__ESTILO_ACADEMIA__*/', () => estilo)
    .replace('__DATA_JSON__', () => JSON.stringify(data).replace(/</g, '\\u003c'))
    .replace('__SUPABASE_URL__', () => env.SUPABASE_URL.replace(/\/+$/, ''))
    .replace('__SUPABASE_KEY__', () => env.SUPABASE_ANON_KEY)
    .replace('<!--__AVISO_PRUEBA__-->', () => aviso);

  // Pagina publica SIN datos de tecnicos; copia completa en privado/.
  const publica = armar({ ...DATA, tecnicos: {} });
  if (Object.values(tecnicos).some((t) => publica.includes(JSON.stringify(t.nombre)))) throw new Error('la pagina publica quedo con nombres de tecnicos; no se escribe nada');
  fs.mkdirSync(PRIVADO, { recursive: true });
  fs.writeFileSync(path.join(PRIVADO, 'index.html'), armar(DATA), 'utf8');
  fs.writeFileSync(path.join(SALIDA, 'index.html'), publica, 'utf8');
  console.log('Portal generado:', path.join(SALIDA, 'index.html'), '(publico, sin datos) y', path.join(PRIVADO, 'index.html'), '(' + DATA.totalTecnicos + ' tecnicos)');

  if (MODO === 'prueba') {
    for (const a of ASSETS) fs.copyFileSync(path.join(carpeta, a), path.join(SALIDA, a));
  } else {
    // Mes actual al historial privado.
    const tecFoto = {};
    for (const [id, t] of Object.entries(tecnicos)) {
      tecFoto[id] = { rep: t.reincidencias && t.reincidencias.total ? t.reincidencias.tasa : null, inf: t.infancia && t.infancia.total ? t.infancia.tasa : null };
    }
    guardarFoto({ slug: periodoSlug, label: nombreMes(periodoSlug), promedioRep: promRep, promedioInf: promInf, tecnicos: tecFoto });
    // Las paginas publicas antiguas (con los datos de todos) se retiran de la carpeta publicada.
    const antiguos = path.join(carpeta, 'privado', 'antiguos');
    for (const f of fs.readdirSync(carpeta)) {
      if (!/^index_\d{4}-\d{2}\.html$/.test(f) && f !== 'meses.json') continue;
      fs.mkdirSync(antiguos, { recursive: true });
      fs.renameSync(path.join(carpeta, f), path.join(antiguos, f));
      console.log('Retirado de la carpeta publicada:', f, '-> privado/antiguos/');
    }
  }

  // El Excel es solo de apoyo: si esta abierto no se puede reescribir, y eso no debe frenar el portal.
  try {
    await generarExcelCredenciales(tecnicos, path.join(SALIDA, 'Credenciales_Tecnicos_NO_SUBIR.xlsx'));
  } catch (err) {
    console.log('AVISO: no se actualizo el Excel de accesos (' + err.code + '; ¿esta abierto?). El portal se genero igual.');
  }
}

// Excel SOLO local (no se sube a GitHub) con el ID de cada tecnico, para repartirlo.
async function generarExcelCredenciales(tecnicos, outPath) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Portal de Tecnicos';
  wb.created = new Date();
  const ws = wb.addWorksheet('Accesos', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = [
    { header: 'Nombre completo', key: 'nombre', width: 34 },
    { header: 'Agencia', key: 'agencia', width: 16 },
    { header: 'Supervisor', key: 'supervisor', width: 30 },
    { header: 'ID (dato de acceso)', key: 'id', width: 20 },
  ];
  Object.entries(tecnicos)
    .sort((a, b) => a[1].nombre.localeCompare(b[1].nombre))
    .forEach(([id, t]) => ws.addRow({ nombre: t.nombre, agencia: t.agencia, supervisor: t.supervisor || '', id }));
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF003575' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  headerRow.height = 26;
  ws.eachRow({ includeEmpty: false }, (row, n) => {
    if (n === 1) return;
    row.eachCell((cell) => { cell.font = { name: 'Arial', size: 10.5 }; });
    row.getCell('id').font = { name: 'Arial', size: 11, bold: true };
    row.getCell('id').numFmt = '@';
  });
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: Object.keys(tecnicos).length + 1, column: 4 } };
  const notas = wb.addWorksheet('Notas');
  notas.columns = [{ width: 100 }];
  [
    'ACCESOS DEL PORTAL DE TECNICOS',
    '',
    'Este archivo es SOLO PARA USO LOCAL. No se sube a GitHub ni se publica en ningun lado.',
    '',
    'Cada tecnico entra en https://supervision360.github.io/portal-tecnicos/ escribiendo solo su ID.',
    'El ID son los ultimos 6 caracteres de su RUT, incluido el digito verificador y sin guion. Si el RUT termina en K, se usa 0.',
    'Desde el 14-10-2026 ya no se usan el nombre ni la clave de 4 digitos.',
    '',
    'Se regenera cada dia con la publicacion de las 11:00.',
  ].forEach((t, i) => {
    const row = notas.addRow([t]);
    row.getCell(1).font = i === 0 ? { name: 'Arial', bold: true, size: 12 } : { name: 'Arial', size: 10 };
    row.getCell(1).alignment = { wrapText: true, vertical: 'top' };
  });
  await wb.xlsx.writeFile(outPath);
  console.log('Excel de accesos generado (local, no se sube):', outPath);
}

main().catch((err) => {
  console.error('ERROR:', err.message);
  process.exitCode = 1;
});
