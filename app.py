import os
import re
import smtplib
import traceback

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
# CREDENCIALES
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

            p.text = p.text.replace(
                "{{" + clave + "}}",
                valor
            )

    for table in doc.tables:

        for row in table.rows:

            for cell in row.cells:

                for p in cell.paragraphs:

                    for clave, valor in valores.items():

                        p.text = p.text.replace(
                            "{{" + clave + "}}",
                            valor
                        )

# =========================
# ENVIAR CORREO
# =========================

def enviar_correo(destino, archivo):

    try:

        print("========== INICIO ENVIO ==========")

        print("EMAIL_USER:", EMAIL_USER)
        print("EMAIL_PASSWORD EXISTE:", bool(EMAIL_PASSWORD))
        print("DESTINO:", destino)

        if not EMAIL_USER:
            print("ERROR: EMAIL_USER NO CONFIGURADO")
            return

        if not EMAIL_PASSWORD:
            print("ERROR: EMAIL_PASSWORD NO CONFIGURADO")
            return

        mensaje = EmailMessage()

        mensaje["Subject"] = "Documento generado automáticamente"
        mensaje["From"] = EMAIL_USER
        mensaje["To"] = destino

        mensaje.set_content(
            "Adjunto encontrarás el documento generado."
        )

        with open(archivo, "rb") as f:

            mensaje.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename="documento.docx"
            )

        print("Conectando a Gmail...")

        with smtplib.SMTP(
            "smtp.gmail.com",
            587,
            timeout=60
        ) as smtp:

            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()

            print("Iniciando sesión Gmail...")

            smtp.login(
                EMAIL_USER,
                EMAIL_PASSWORD
            )

            print("Enviando correo...")

            smtp.send_message(
                mensaje
            )

        print("CORREO ENVIADO CORRECTAMENTE")
        print("========== FIN ENVIO ==========")

    except Exception:

        print("========== ERROR CORREO ==========")
        print(traceback.format_exc())
        print("=================================")

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

            correo = request.form.get(
                "correo",
                ""
            ).strip()

            doc = Document(PLANTILLA)

            reemplazar(
                doc,
                valores
            )

            archivo_salida = os.path.join(
                CARPETA,
                "documento.docx"
            )

            doc.save(
                archivo_salida
            )

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

        return render_template(
            "index.html",
            variables=variables
        )

    except Exception as e:

        return (
            "<pre>"
            + str(e)
            + "\n\n"
            + traceback.format_exc()
            + "</pre>",
            500
        )

# =========================
# RUN LOCAL
# =========================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=10000,
        debug=True
    )
