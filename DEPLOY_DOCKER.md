# Despliegue Docker - Genoma PDF Studio

Esta app esta preparada para correr en un VPS con Docker, Docker Compose o Easypanel.

## Opcion recomendada: Docker Compose

Desde la carpeta del proyecto:

```bash
docker compose up -d --build
```

Luego abre:

```text
http://TU_SERVIDOR:5055
```

## Variables importantes

- `PORT`: puerto interno de la app. Por defecto `5055`.
- `HOST`: debe ser `0.0.0.0` dentro de Docker.
- `GENOMA_DATA_DIR`: carpeta persistente de datos. Por defecto `/data`.
- `CHROME_PATH`: ruta de Chromium dentro del contenedor. Por defecto `/usr/bin/chromium`.

## Datos persistentes

El contenedor usa `/data` para guardar:

- PDFs subidos.
- Valores del sistema.
- Micobioma.
- Consideraciones.

En `docker-compose.yml` esto queda en el volumen `genoma_data`.

## Easypanel

En Easypanel puedes crear una app Docker desde este repositorio/carpeta:

1. Tipo: Dockerfile.
2. Puerto expuesto: `5055`.
3. Variables:
   - `HOST=0.0.0.0`
   - `PORT=5055`
   - `GENOMA_DATA_DIR=/data`
   - `CHROME_PATH=/usr/bin/chromium`
4. Agrega un volumen persistente montado en `/data`.
5. Configura el dominio apuntando al puerto `5055`.

## Seguridad

La app maneja datos clinicos. Antes de abrirla publicamente conviene agregar:

- Usuario y contrasena.
- HTTPS en el proxy/dominio.
- Backups del volumen `/data`.
