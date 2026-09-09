# Plataforma Pausas Activas SST

Aplicacion web para registrar evidencia fotografica de pausas activas por fecha, turno A/B/C y centro de trabajo. La plantilla original queda intacta en `01_REGISTRO PAUSAS ACTIVAS AGOSTO  2026--.xlsx`.

## Archivos

- `index.html`: dashboard y formulario.
- `styles.css`: diseno visual con Bootstrap.
- `app.js`: adjunto de fotos, carga de historial desde Excel en linea, metricas, descarga Excel y sincronizacion.
- `apps-script-pausas-activas.gs`: backend para Google Sheets y subida de fotos a Drive.

## Conexion con Excel en linea

La forma mas estable para guardar historial fotografico es usar Google Sheets como hoja en linea:

1. Sube o crea una hoja en Google Drive para el historial.
2. Abre `Extensiones > Apps Script`.
3. Pega el contenido de `apps-script-pausas-activas.gs`.
4. Guarda y publica como `Implementar > Nueva implementacion > Aplicacion web`.
5. Ejecutar como: `Yo`.
6. Acceso: usuarios autorizados segun tu necesidad.
7. Publica la implementacion. La URL ya quedo interna en el sistema:
`https://script.google.com/macros/s/AKfycbzDIVViPUDu4KNQph9GfiaVgaBnA6fVW9AyW5-beqik554U4sGkJ3R1J_G1wefLs_LHeQ/exec`

Al abrir esa URL en el navegador debe responder `ok:true` y mostrar `spreadsheetUrl`. Si el Apps Script no esta ligado a una hoja, el mismo script crea una hoja llamada `REGISTRO PAUSAS ACTIVAS SST EN LINEA`.

Cada sincronizacion sube las fotos a la carpeta `Evidencias Pausas Activas SST` y crea o actualiza estas hojas:

- `Historial Pausas Activas`
- `Resumen Mensual`

El dashboard muestra las fotos por turno y alerta cuando falta evidencia en A, B o C para la fecha seleccionada.
El logo y colores de ZIMA se usan en el sistema y en el Excel descargable.

## Uso local

Abre `index.html` en el navegador. Mientras no sincronices, los registros quedan guardados en el navegador. El boton `Descargar Excel`, ubicado en el header superior, genera `REGISTRO-PAUSAS-ACTIVAS-HISTORIAL.xlsx` con toda la informacion ingresada agrupada por mes, fecha y turnos A/B/C.
