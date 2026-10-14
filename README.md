# Portal de Técnicos Tigo — COBRA

https://supervision360.github.io/portal-tecnicos/

Igual al portal de técnicos de VTR (`C:\Bases_Vtr\portal-tecnicos`) desde el 14-10-2026:

- Cada técnico entra solo con su **ID**: últimos 6 caracteres de su RUT, con el dígito
  verificador y sin guion (si termina en K, se usa 0).
- La página publicada (`index.html`) **no trae los datos de nadie**. Al escribir su ID, el
  portal le pide a Supabase solo sus datos (función `datos_tecnico` del proyecto del Portal
  Supervisor), que además anota el ingreso y frena a quien prueba muchos ID seguidos.
- Ve: Tu resumen (Repetido Reparado, Averías de Infancia, Productividad y NPS), Tu posición
  en el equipo, Tu evolución mes a mes y el detalle de cada indicador. Instalable como app
  (botón "Instalar app").

El RUT se usa solo al generar, en este computador; nunca se guarda en ninguna página.

## Archivos

```
portal-tecnicos/
├── template.html                 ← plantilla del portal (igual a VTR, indicadores de Tigo)
├── generar_portal.js             ← genera index.html (público, sin datos) y privado/index.html
├── subir_datos_tecnicos.js       ← carga a Supabase los datos de cada técnico (desde privado/)
├── privado/                      ← NO se sube: datos completos, historial por mes (meses/), antiguos/
├── prueba/                       ← NO se sube: versión de prueba (node generar_portal.js prueba)
├── Credenciales_Tecnicos_NO_SUBIR.xlsx ← ID de cada técnico para repartir (no se sube)
└── manifest.json, sw.js, icon-*.png, logo-cobra.png, fondo-portada.jpg
```

## Actualizar

```
node generar_portal.js            (o "prueba": todo queda en prueba/)
node subir_datos_tecnicos.js      (o "prueba")
```

La publicación diaria (11:00, `..\Generar_Reporte_Reincidencias_programado.bat`) hace lo
mismo y después `git push`. El Portal Supervisor lee `privado/index.html`.

## ID repetidos

Si dos técnicos distintos quedan con el mismo ID, ninguno de los dos queda habilitado (nadie
debe ver los datos de otro) y aparece un aviso. Hay que darle a uno un ID distinto.

## Técnicos desvinculados

Lista `DESVINCULADOS` en `generar_portal.js`: no entran al portal, pero su trabajo sigue
contando en los promedios del equipo.
