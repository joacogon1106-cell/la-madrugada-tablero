# La Madrugada — Tablero Agrícola

Tablero interactivo de gestión agrícola que se actualiza automáticamente desde planillas en Dropbox.

🌐 **App en vivo**: [lamadrugada.vercel.app](https://lamadrugada.vercel.app)

## Cómo funciona

```
[Planillas Excel en Dropbox]
            ↓
[GitHub Actions corre todos los días a las 7 AM Argentina]
            ↓
[Lee planillas vía Dropbox API]
            ↓
[Procesa datos con Python]
            ↓
[Genera index.html actualizado]
            ↓
[Push automático al repo]
            ↓
[Vercel detecta el cambio y redeploya]
            ↓
[lamadrugada.vercel.app actualizado]
```

## Archivos

- `actualizar.py` — Script principal: descarga planillas, procesa y genera HTML
- `template.html` — Plantilla HTML con un `__DATA_PLACEHOLDER__` donde se inyectan los datos
- `requirements.txt` — Dependencias de Python
- `.github/workflows/actualizar.yml` — Configuración de GitHub Actions
- `index.html` — Archivo generado automáticamente (no editar a mano)

## Secrets de GitHub que necesita

Configurados en: Settings → Secrets and variables → Actions

- `DROPBOX_APP_KEY` — App key de la app de Dropbox
- `DROPBOX_APP_SECRET` — App secret
- `DROPBOX_REFRESH_TOKEN` — Refresh token de larga duración

## Forzar actualización manual

Andá a la pestaña **Actions** → "Actualizar tablero diario" → "Run workflow".

## Carpeta de Dropbox

Las planillas viven en:
```
/JG/TRABAJO/Clientes/La Madrugada/Agricultura La Madrugada/
```

Y deben llamarse exactamente:
- `Planilla de Cultivos La Madrugada.xlsx`
- `Planilla Agricultura La Madrugada 26-27.xlsx`
