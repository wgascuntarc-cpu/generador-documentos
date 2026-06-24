import os
import re
import smtplib
from flask import Flask, render_template, request, send_file
from docx import Document
from email.message import EmailMessage

app = Flask(__name__)

# =========================
# RUTAS BASE
# =========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PLANTILLA = os.path.join(BASE_DIR, "plantilla.docx")

CARPETA = os.path.join(BASE_DIR, "documentos_generados")
os.makedirs(CARPETA, exist_ok=True)

# =========================
# CREDENCIALES RENDER
# =========================

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

# =========================
# OBTENER VARIABLES
# =========================

def obtener_variables():
    doc = Document(PLANTILLA)

    texto = ""

    for p in doc.paragraphs:
        texto += p.text + "\n"

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                texto += cell.text + "\n"

    variables = re.findall(r"\{\{(.*?)\}\}", texto)
    return sorted(set(variables))

# =========================
# REEMPLAZAR VARIABLES
# =========================

def reemplazar(doc, valores):
    for p in doc.paragraphs:
        for clave, valor in valores.items():
            p.text = p.text.replace("{{" + clave + "}}", valor)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for clave, valor in valores.items():
                        p.text = p.text.replace("{{" + clave + "}}", valor)

# =========================
# ENVIAR CORREO
# =========================

# ENVIAR CORREO
if correo:
    enviar_correo(
        correo,
        archivo_salida
    )

# =========================
# RUTA PRINCIPAL
# =========================

@app.route("/", methods=["GET", "POST"])
def index():

    try:
        variables = obtener_variables()

        if request.method == "POST":

            valores = {}

            for v in variables:
                valores[v] = request.form.get(v, "")

            correo = request.form.get("correo", "").strip()

            doc = Document(PLANTILLA)
            reemplazar(doc, valores)

            archivo_salida = os.path.join(CARPETA, "documento.docx")
            doc.save(archivo_salida)

            # ENVIAR CORREO EN SEGUNDO PLANO (IMPORTANTE PARA RENDER)
           # ENVIAR CORREO
if correo:
    enviar_correo(
        correo,
        archivo_salida
    )
            return send_file(
                archivo_salida,
                as_attachment=True,
                download_name="documento.docx"
            )

        return render_template("index.html", variables=variables)

    except Exception as e:
        import traceback
        return f"<pre>{str(e)}\n\n{traceback.format_exc()}</pre>", 500

# =========================
# RUN LOCAL
# =========================

if __name__ == "__main__":
    app.run(debug=True)
