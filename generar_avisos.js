// Avisos de la campanita (portal de tecnicos Tigo y portal supervisor Tigo): arma avisos.json
// desde avisos_manuales.json (se publica junto a index.html). Solo manuales: para avisar de algo,
// agregar una linea en avisos_manuales.json y correr este script, luego publicar avisos.json.
//   { "id": "texto-unico", "fecha": "AAAA-MM-DD", "tipo": "comunicado|video|otro", "titulo": "Texto", "enlace": "https://..." }
// Solo se guardan los avisos de los ultimos 30 dias. Uso: node generar_avisos.js
const fs = require("fs");
const path = require("path");
const DIAS = 30;
const manuales = JSON.parse(fs.readFileSync(path.join(__dirname, "avisos_manuales.json"), "utf-8"));
const limite = new Date(Date.now() - DIAS * 86400000).toISOString().slice(0, 10);
const ids = new Set();
const avisos = manuales
  .filter((a) => a.id && a.fecha && a.titulo && a.enlace && a.fecha >= limite && !ids.has(a.id) && ids.add(a.id))
  .sort((a, b) => b.fecha.localeCompare(a.fecha));
fs.writeFileSync(path.join(__dirname, "avisos.json"), JSON.stringify({ avisos }, null, 1) + "\n");
console.log("avisos.json: " + avisos.length + " avisos");
