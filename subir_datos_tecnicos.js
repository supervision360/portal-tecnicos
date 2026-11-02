// Sube a Supabase (tabla datos_tecnicos del proyecto del Portal Supervisor) los datos de
// cada tecnico, desde la copia completa privado/index.html que deja generar_portal.js.
// El index.html publico va sin esos datos: el portal se los pide a Supabase con el ID.
// Igual que C:\Bases_Vtr\portal-tecnicos\reportes\subir_datos_tecnicos.js.
//
//   node subir_datos_tecnicos.js          (despues de node generar_portal.js)
//   node subir_datos_tecnicos.js prueba   (despues de node generar_portal.js prueba)

const fs = require('fs');
const path = require('path');
const { configSupabase } = require('C:\\Bases_Tigo\\portal-supervisor\\lib\\proteger.js');

async function main() {
  const base = process.argv[2] === 'prueba' ? path.join(__dirname, 'prueba') : __dirname;
  const html = fs.readFileSync(path.join(base, 'privado', 'index.html'), 'utf8');
  const m = html.match(/^const DATA = (.+);\r?$/m);
  if (!m) throw new Error('privado/index.html no trae el bloque de datos');
  const tecnicos = JSON.parse(m[1]).tecnicos || {};
  const n = Object.keys(tecnicos).length;
  if (!n) throw new Error('privado/index.html no trae tecnicos; no se sube nada');
  const cfg = configSupabase();
  const r = await fetch(cfg.url + '/rest/v1/rpc/cargar_datos_tecnicos', {
    method: 'POST',
    headers: { apikey: cfg.serviceKey, Authorization: 'Bearer ' + cfg.serviceKey, 'Content-Type': 'application/json' },
    body: JSON.stringify({ p_datos: tecnicos }),
  });
  if (!r.ok) throw new Error('Supabase no guardo los datos de los tecnicos (' + r.status + '): ' + (await r.text()).slice(0, 200));
  const cargados = await r.json();
  if (cargados !== n) throw new Error('Supabase cargo ' + cargados + ' de ' + n + ' tecnicos');
  console.log('==> Supabase: datos de ' + n + ' tecnicos cargados');
}

main().catch((err) => { console.error('ERROR:', err.message); process.exit(1); });
