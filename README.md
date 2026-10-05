# Agenda de Salud · ATA Consultores

Calendario de efemérides del sector salud (Global y Venezuela) con la identidad de ATA Consultores.

- **Sitio:** https://agenda-de-salud-ata.vercel.app
- **Modo incrustado:** `https://agenda-de-salud-ata.vercel.app/?embed=1` (sin barra de logo, avisa su altura al sitio que lo contiene)

## Archivos

| Archivo | Qué es |
|---|---|
| `index.html` | El calendario completo (vistas Mes / Semana / Lista, filtros, búsqueda, suscripción). Página estática, sin build. |
| `ata-logo.webp` / `ata-logo.png` | Logo de ATA (webp para la web, png para los correos). |
| `vercel.json` | Cabeceras de caché y permiso para incrustarse en iframe. |
| `embed-wordpress.html` | Código para pegar en un widget HTML de Elementor en ataconsultores.com. |
| `AgendaSaludATA.gs` | Script de Google Sheets: correos semanal y mensual, Google Calendar suscribible y API JSON para la web. |
| `efemerides.csv` | Lista maestra de efemérides (142 fechas) tal como se importó a la hoja. |

## Configuración del calendario web

Al inicio del `<script>` de `index.html`:

```js
var CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbwgFwuT3UWt2mYw9XbEIHCujuEb1q4avd-5e6tZMkKUT7vr6tBFSsQW_ZHjqhUB9xFS/exec",
  CALENDARIO_ID: "c_4b81c60910c6e9cbb473b71d9c72468f8397a0cb14332daabf45909c262333f5@group.calendar.google.com"
};
```

- `API_URL`: aplicación web de Apps Script (cuenta karla.oviedo@ataconsultores.com, hoja "Agenda de Salud ATA – Efemérides"). La web lee la hoja en vivo; si la API falla, usa los datos incluidos en el archivo.
- `CALENDARIO_ID`: ID del Google Calendar público (termina en `@group.calendar.google.com`). Activa el botón "Suscribirme al calendario".

## Reglas de fechas móviles (columna "Regla" de la hoja)

| Regla | Significado |
|---|---|
| vacía | Fecha fija (Mes + Día) |
| `JUE#2` | 2.º jueves del mes |
| `DOM#U` | Último domingo del mes |
| `VIE#U` + Día 19 | Último viernes en o antes del 19 |
| `ULTDIA` | Último día del mes |

## Despliegue

Vercel (proyecto `agenda-de-salud-ata`, cuenta personal). El repositorio está conectado al proyecto: cada commit a `main` publica la nueva versión.
