import os
import re
import uuid
import pandas as pd
from datetime import datetime
from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file, jsonify
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from werkzeug.utils import secure_filename

# ─────────────────────────── Configuración ───────────────────────────
BASE_DIR    = os.path.abspath(os.path.dirname(__file__))
UPLOAD_DIR  = os.path.join(BASE_DIR, "uploads")
EXPORT_DIR  = os.path.join(BASE_DIR, "exports")
DB_PATH     = os.path.join(BASE_DIR, "database", "mensajeria.db")

ALLOWED_EXTENSIONS = {"xlsx", "xls"}

app = Flask(__name__)
app.secret_key = os.urandom(32)

app.config["SQLALCHEMY_DATABASE_URI"]        = f"sqlite:///{DB_PATH}"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["UPLOAD_FOLDER"]                  = UPLOAD_DIR
app.config["MAX_CONTENT_LENGTH"]             = 16 * 1024 * 1024  # 16 MB

db = SQLAlchemy(app)

for d in [UPLOAD_DIR, EXPORT_DIR, os.path.dirname(DB_PATH)]:
    os.makedirs(d, exist_ok=True)


# ─────────────────────────── Modelo ──────────────────────────────────
class ContactoTemporal(db.Model):
    __tablename__ = "contactos_temporales"

    id          = db.Column(db.Integer, primary_key=True)
    sesion_id   = db.Column(db.String(36), nullable=False, index=True)
    nombre      = db.Column(db.String(200))
    apellido    = db.Column(db.String(200))
    documento_madre = db.Column(db.String(100))
    telefono    = db.Column(db.String(50))
    estado      = db.Column(db.String(20), default="valido")   # valido | duplicado | sin_telefono
    creado_en   = db.Column(db.DateTime, default=datetime.utcnow)

    def to_dict(self):
        return {
            "id":       self.id,
            "nombre":   self.nombre   or "",
            "apellido": self.apellido or "",
            "documento_madre": self.documento_madre or "",
            "telefono": self.telefono or "",
            "estado":   self.estado,
        }


with app.app_context():
    db.create_all()
    # Migración ligera para bases SQLite existentes sin la nueva columna.
    columnas = db.session.execute(text("PRAGMA table_info(contactos_temporales)")).fetchall()
    nombres_columnas = {c[1] for c in columnas}
    if "documento_madre" not in nombres_columnas:
        db.session.execute(text("ALTER TABLE contactos_temporales ADD COLUMN documento_madre VARCHAR(100)"))
        db.session.commit()


# ─────────────────────────── Utilidades ──────────────────────────────
def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# Palabras clave para detectar columnas automáticamente.
# Se prioriza coincidencia exacta (strip+lower) antes de coincidencia parcial.
KEYWORDS = {
    # Prioridad: columnas exactas del formato CVS → luego genéricas
    "nombre": [
        "madre nombre1", "nombre madre", "nombre de la madre",
        "madre nombre", "nombre1", "nombre", "nombres", "primer nombre",
    ],
    "apellido": [
        "madre apellido1", "apellido madre", "apellido de la madre",
        "madre apellido", "apellido1", "apellido", "apellidos", "primer apellido",
    ],
    "documento_madre": [
        "madre documento", "documento madre", "doc madre",
        "documento de la madre", "madre doc", "cedula madre", "cédula madre",
    ],
    "telefono": [
        "telefono", "teléfono", "celular", "cel", "movil", "móvil",
        "numero", "número", "whatsapp", "contacto", "tel",
        "telefono madre", "teléfono madre", "celular madre",
        "numero de contacto", "número de contacto",
    ],
}


def detectar_columna(columnas: list[str], tipo: str) -> str | None:
    """
    Retorna la primera columna que coincida con las palabras clave del tipo dado.
    Busca primero por coincidencia exacta, luego por coincidencia parcial,
    respetando el orden de prioridad definido en KEYWORDS.
    """
    keywords = KEYWORDS.get(tipo, [])
    col_map = {c.lower().strip(): c for c in columnas}

    # 1) Coincidencia exacta (prioridad máxima, respeta orden de keywords)
    for kw in keywords:
        if kw in col_map:
            return col_map[kw]

    # 2) Coincidencia parcial
    for kw in keywords:
        for col_lower, col_orig in col_map.items():
            if kw in col_lower or col_lower in kw:
                return col_orig

    return None


def normalizar_telefono(valor) -> str:
    """Limpia y normaliza un número telefónico."""
    if pd.isna(valor):
        return ""
    tel = str(valor).strip()
    # Quitar caracteres que no sean dígitos ni el + inicial
    tel = re.sub(r"[^\d+]", "", tel)
    # Si queda vacío o solo tiene + devolver cadena vacía
    if not tel or tel == "+":
        return ""
    # Mínimo 7 dígitos para ser un número válido (rechaza ej. solo "595")
    solo_digitos = re.sub(r"[^\d]", "", tel)
    if len(solo_digitos) < 7:
        return ""
    return tel


def limpiar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip()


def _leer_hoja_detectando_header(ruta: str, hoja: str) -> pd.DataFrame:
    """
    Lee una hoja detectando automáticamente la fila de encabezados.
    Soporta archivos donde la primera fila es un título (ej. formato CVS).
    """
    # Leer sin encabezado para escanear las primeras filas
    preview = pd.read_excel(ruta, sheet_name=hoja, engine="openpyxl",
                            header=None, nrows=10, dtype=str)

    header_row = 0  # valor por defecto
    for idx, row in preview.iterrows():
        # Buscar la fila donde la mayoría de celdas son texto no-vacío
        valores = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
        if len(valores) >= 3:
            # Si al menos una celda contiene palabra clave de encabezado real
            row_lower = " ".join(v.lower() for v in valores)
            if any(kw in row_lower for kw in
                   ["nombre", "apellido", "telefono", "id persona", "documento", "celular"]):
                header_row = idx
                break

    df = pd.read_excel(ruta, sheet_name=hoja, engine="openpyxl",
                       header=header_row, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def procesar_dataframe(
    df: pd.DataFrame,
    col_nombre: str,
    col_apellido: str,
    col_telefono: str,
    col_doc_madre: str = "",
):
    """
    Filtra, limpia y devuelve una lista de dicts con estadísticas.
    """
    registros   = []
    total       = len(df)
    sin_tel     = 0
    duplicados  = 0
    telefonos_vistos = set()

    for _, row in df.iterrows():
        nombre   = limpiar_texto(row.get(col_nombre,   ""))
        apellido = limpiar_texto(row.get(col_apellido, ""))
        documento_madre = limpiar_texto(row.get(col_doc_madre, "")) if col_doc_madre else ""
        telefono = normalizar_telefono(row.get(col_telefono, ""))

        if not telefono:
            sin_tel += 1
            registros.append({
                "nombre":   nombre,
                "apellido": apellido,
                "documento_madre": documento_madre,
                "telefono": telefono,
                "estado":   "sin_telefono",
            })
            continue

        if telefono in telefonos_vistos:
            duplicados += 1
            registros.append({
                "nombre":   nombre,
                "apellido": apellido,
                "documento_madre": documento_madre,
                "telefono": telefono,
                "estado":   "duplicado",
            })
            continue

        telefonos_vistos.add(telefono)
        registros.append({
            "nombre":   nombre,
            "apellido": apellido,
            "documento_madre": documento_madre,
            "telefono": telefono,
            "estado":   "valido",
        })

    validos = total - sin_tel - duplicados
    stats = {
        "total":      total,
        "sin_tel":    sin_tel,
        "duplicados": duplicados,
        "validos":    validos,
    }
    return registros, stats


# ─────────────────────────── Rutas ───────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "excel_file" not in request.files:
        flash("No se seleccionó ningún archivo.", "danger")
        return redirect(url_for("index"))

    archivo = request.files["excel_file"]

    if archivo.filename == "":
        flash("El archivo no tiene nombre.", "danger")
        return redirect(url_for("index"))

    if not allowed_file(archivo.filename):
        flash("Solo se permiten archivos .xlsx o .xls.", "danger")
        return redirect(url_for("index"))

    nombre_seguro = secure_filename(archivo.filename)
    ruta          = os.path.join(app.config["UPLOAD_FOLDER"], nombre_seguro)
    archivo.save(ruta)

    # Leer nombres de hojas
    try:
        xls   = pd.ExcelFile(ruta, engine="openpyxl")
        hojas = xls.sheet_names
    except Exception as e:
        flash(f"Error al leer el archivo: {e}", "danger")
        return redirect(url_for("index"))

    if not hojas:
        flash("El archivo no contiene hojas.", "danger")
        return redirect(url_for("index"))

    session["archivo"]    = nombre_seguro
    session["ruta"]       = ruta
    session["hojas"]      = hojas

    return redirect(url_for("select_sheet"))


@app.route("/select-sheet")
def select_sheet():
    if "hojas" not in session:
        flash("Primero debes cargar un archivo.", "warning")
        return redirect(url_for("index"))

    return render_template(
        "select_sheet.html",
        archivo=session.get("archivo"),
        hojas=session.get("hojas"),
    )


@app.route("/process", methods=["POST"])
def process():
    hoja = request.form.get("hoja", "").strip()
    if not hoja:
        flash("Debes seleccionar una hoja.", "warning")
        return redirect(url_for("select_sheet"))

    ruta = session.get("ruta")
    if not ruta or not os.path.exists(ruta):
        flash("El archivo ya no está disponible. Cárgalo nuevamente.", "danger")
        return redirect(url_for("index"))

    try:
        df = _leer_hoja_detectando_header(ruta, hoja)
    except Exception as e:
        flash(f"Error al leer la hoja '{hoja}': {e}", "danger")
        return redirect(url_for("select_sheet"))

    if df.empty:
        flash("La hoja seleccionada está vacía.", "warning")
        return redirect(url_for("select_sheet"))

    columnas = list(df.columns)

    # Detección automática de columnas
    col_nombre   = detectar_columna(columnas, "nombre")
    col_apellido = detectar_columna(columnas, "apellido")
    col_doc_madre = detectar_columna(columnas, "documento_madre")
    col_telefono = detectar_columna(columnas, "telefono")

    # Columnas enviadas manualmente (si el usuario las ajusta)
    col_nombre   = request.form.get("col_nombre",   col_nombre   or "")
    col_apellido = request.form.get("col_apellido", col_apellido or "")
    col_doc_madre = request.form.get("col_doc_madre", col_doc_madre or "")
    col_telefono = request.form.get("col_telefono", col_telefono or "")

    # Si faltan columnas, redirigir con columnas detectadas para que el usuario ajuste
    if not col_telefono:
        flash(
            "No se detectó la columna de teléfono automáticamente. "
            "Selecciónala manualmente.",
            "warning",
        )
        session["hoja"]      = hoja
        session["columnas"]  = columnas
        session["col_nombre"]   = col_nombre
        session["col_apellido"] = col_apellido
        session["col_doc_madre"] = col_doc_madre
        session["col_telefono"] = col_telefono
        return redirect(url_for("map_columns"))

    # Procesar
    registros, stats = procesar_dataframe(
        df,
        col_nombre,
        col_apellido,
        col_telefono,
        col_doc_madre,
    )

    # Guardar en BD con sesión única
    sesion_id = str(uuid.uuid4())
    session["sesion_id"] = sesion_id
    session["hoja"]      = hoja
    session["stats"]     = stats

    # Borrar registros anteriores de esta sesión (por si recargó)
    ContactoTemporal.query.filter_by(sesion_id=sesion_id).delete()

    for r in registros:
        db.session.add(ContactoTemporal(
            sesion_id=sesion_id,
            nombre=r["nombre"],
            apellido=r["apellido"],
            documento_madre=r["documento_madre"],
            telefono=r["telefono"],
            estado=r["estado"],
        ))
    db.session.commit()

    return redirect(url_for("preview"))


@app.route("/map-columns", methods=["GET", "POST"])
def map_columns():
    """Permite al usuario seleccionar manualmente las columnas."""
    if "hojas" not in session:
        return redirect(url_for("index"))

    if request.method == "POST":
        return redirect(url_for("process"), code=307)  # reenviar con POST

    return render_template(
        "map_columns.html",
        archivo=session.get("archivo"),
        hoja=session.get("hoja"),
        columnas=session.get("columnas", []),
        col_nombre=session.get("col_nombre", ""),
        col_apellido=session.get("col_apellido", ""),
        col_doc_madre=session.get("col_doc_madre", ""),
        col_telefono=session.get("col_telefono", ""),
    )


@app.route("/preview")
def preview():
    sesion_id = session.get("sesion_id")
    if not sesion_id:
        flash("No hay datos procesados. Carga un archivo primero.", "warning")
        return redirect(url_for("index"))

    contactos = ContactoTemporal.query.filter_by(sesion_id=sesion_id).all()
    stats     = session.get("stats", {})

    return render_template(
        "preview.html",
        contactos=contactos,
        stats=stats,
        archivo=session.get("archivo"),
        hoja=session.get("hoja"),
    )


@app.route("/export/<fmt>")
def export(fmt: str):
    sesion_id = session.get("sesion_id")
    if not sesion_id:
        flash("No hay datos para exportar.", "warning")
        return redirect(url_for("index"))

    # Solo exportar válidos
    contactos = ContactoTemporal.query.filter_by(
        sesion_id=sesion_id, estado="valido"
    ).all()

    if not contactos:
        flash("No hay contactos válidos para exportar.", "warning")
        return redirect(url_for("preview"))

    datos = [c.to_dict() for c in contactos]
    df    = pd.DataFrame(datos, columns=["nombre", "apellido", "documento_madre", "telefono", "estado"])
    df    = df.drop(columns=["id"], errors="ignore")

    timestamp    = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_base  = f"contactos_{timestamp}"

    if fmt == "excel":
        ruta_export = os.path.join(EXPORT_DIR, f"{nombre_base}.xlsx")
        df.to_excel(ruta_export, index=False, engine="openpyxl")
        return send_file(
            ruta_export,
            as_attachment=True,
            download_name=f"{nombre_base}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif fmt == "csv":
        ruta_export = os.path.join(EXPORT_DIR, f"{nombre_base}.csv")
        df.to_csv(ruta_export, index=False, encoding="utf-8-sig")
        return send_file(
            ruta_export,
            as_attachment=True,
            download_name=f"{nombre_base}.csv",
            mimetype="text/csv",
        )
    else:
        flash("Formato de exportación no soportado.", "danger")
        return redirect(url_for("preview"))


@app.route("/reset")
def reset():
    """Limpia la sesión actual y borra los contactos temporales."""
    sesion_id = session.get("sesion_id")
    if sesion_id:
        ContactoTemporal.query.filter_by(sesion_id=sesion_id).delete()
        db.session.commit()

    # Borrar archivo subido
    ruta = session.get("ruta")
    if ruta and os.path.exists(ruta):
        try:
            os.remove(ruta)
        except OSError:
            pass

    session.clear()
    flash("Sesión reiniciada correctamente.", "success")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.getenv("PORT", "5050")),
        debug=os.getenv("FLASK_DEBUG", "0") == "1",
    )
