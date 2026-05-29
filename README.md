# Genoma PDF Studio

App web para cargar informes PDF de microbiota, validar datos extraidos y generar un PDF final imprimible.

## Despliegue en EasyPanel

1. Crear una app desde GitHub.
2. Seleccionar este repositorio.
3. Elegir build con `Dockerfile`.
4. Usar `Dockerfile` como ruta del Dockerfile.
5. Exponer el puerto `5055`.
6. Configurar variables:
   - `HOST=0.0.0.0`
   - `PORT=5055`
   - `GENOMA_DATA_DIR=/data`
   - `CHROME_PATH=/usr/bin/chromium`
7. Agregar volumen persistente:
   - `/data`

## Actualizaciones

Cada cambio enviado a GitHub puede redesplegarse desde EasyPanel. Si EasyPanel tiene auto deploy activo, el despliegue se dispara al hacer push.
