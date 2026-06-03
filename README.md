# Mensajería Masiva — v1.0

Aplicación web en Flask para preparar envíos masivos desde archivos Excel.

## Requisitos

- Python 3.10 o superior

## Instalación

```bash
# 1. Crear entorno virtual
python -m venv venv

# 2. Activar entorno (Windows)
venv\Scripts\activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar la aplicación
python app.py
```

Luego abre tu navegador en: `http://127.0.0.1:5050`

## Despliegue en Railway

1. Sube este repositorio a GitHub.
2. Crea un nuevo proyecto en Railway y conecta el repositorio.
3. Railway detectará Python y usará `gunicorn` con el `Procfile`.
4. No necesitas configurar un comando de inicio manual.
5. Agrega un servicio de PostgreSQL en Railway y enlaza su `DATABASE_URL` al servicio de la app.

Variables recomendadas:

- `DATABASE_URL`: la URL del Postgres de Railway.
- `SECRET_KEY`: una cadena larga y fija para mantener las sesiones estables.
- `FLASK_DEBUG=0`: desactiva debug en producción.

Con esto la app deja de depender de SQLite local y los datos quedan persistidos en la base externa.

## Flujo de uso

| Paso | Acción |
|------|--------|
| 1 | Sube un archivo `.xlsx` o `.xls` |
| 2 | Selecciona la hoja visible/activa a procesar |
| 3 | Revisa y confirma columnas (teléfono, nombre, apellido, documento y edad si aplica) |
| 4 | Revisa la vista previa con estadísticas y filtros |
| 5 | Exporta los contactos válidos en Excel o CSV |

## Estructura del proyecto

```
/
├── app.py                  ← Aplicación Flask principal
├── requirements.txt
├── README.md
├── templates/
│   ├── base.html
│   ├── index.html          ← Pantalla de carga
│   ├── select_sheet.html   ← Selección de hoja
│   ├── map_columns.html    ← Asignación manual de columnas
│   └── preview.html        ← Vista previa y exportación
├── static/
│   └── css/style.css
├── uploads/                ← Archivos subidos (temporales)
├── exports/                ← Archivos exportados
└── database/
    └── mensajeria.db       ← SQLite (auto-generado)
```

## Limpieza de datos aplicada

- Elimina filas sin número de teléfono.
- Elimina duplicados por número de teléfono.
- Normaliza teléfonos (quita espacios, guiones, paréntesis).
- Limpia espacios en blanco en nombres y apellidos.

## Próxima etapa (v2.0)

- Composición y envío de mensajes personalizados.
