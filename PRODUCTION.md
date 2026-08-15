# Produccion en Railway

## Variables de entorno

Configura estas variables en Railway:

- `DATABASE_URL`: URL del PostgreSQL de Railway. Si existe, la app usa PostgreSQL y no SQLite.
- `SECRET_KEY`: obligatoria en Railway. Debe ser una cadena larga, fija y secreta.
- `ADMIN_PASSWORD`: solo se usa para crear el primer administrador si no existe ningun usuario con `is_admin=True`. Nunca reinicia ni cambia contrasenas existentes.
- `ADMIN_USERNAME`: opcional. Nombre del primer administrador; por defecto `admin`.
- `ADMIN_FORCE_RESET=1`: opcion temporal para recuperar acceso. Si esta activa, fuerza que `ADMIN_USERNAME` sea administrador activo y le asigna `ADMIN_PASSWORD`. Quitala despues de iniciar sesion.
- `FLASK_DEBUG=0`: recomendado en produccion.

## Admin inicial

El sistema solo crea un administrador inicial cuando no existe ningun usuario administrador en la base de datos.

Una vez creado el primer administrador, `ADMIN_PASSWORD` y `ADMIN_USERNAME` no modifican usuarios existentes. Para cambiar contrasenas o crear usuarios, entra con un administrador y usa el modulo de Usuarios.

Si pierdes acceso en Railway, configura temporalmente `ADMIN_FORCE_RESET=1` junto con `ADMIN_USERNAME` y `ADMIN_PASSWORD`, redeploya, entra con esas credenciales y luego elimina `ADMIN_FORCE_RESET`.

## Despliegue

1. Sube los cambios a GitHub.
2. En Railway, conecta el repositorio.
3. Agrega PostgreSQL al proyecto.
4. Verifica que Railway tenga `DATABASE_URL`, `SECRET_KEY` y `ADMIN_PASSWORD`.
5. Railway ejecutara `web: gunicorn app:app` desde el `Procfile`.
