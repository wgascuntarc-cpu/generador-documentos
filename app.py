import os
import re
import smtplib

from email.message import EmailMessage
from flask import Flask, render_template, request, send_file
from docx import Document

app = Flask(__name__)

# =========================
# CONFIGURACIÓN
# =========================

PLANTILLA = "plantilla.docx"
CARPETA = "documentos_generados"

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASSWORD")

os.makedirs(CARPETA, exist_ok=True)

# =========================
# OBTENER VARIABLES
# =========================

def obtener_variables():
    doc = Document(PLANTILLA)
    variables = set()

    for p in doc.paragraphs:
        variables.update(re.findall(r"\{\{(.*?)\}\}", p.text))
        variables.update(re.findall(r"\[(.*?)\]", p.text))

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                variables.update(re.findall(r"\{\{(.*?)\}\}", cell.text))
                variables.update(re.findall(r"\[(.*?)\]", cell.text))

    return sorted(list(variables))

# =========================
# REEMPLAZAR VARIABLES
# =========================

def reemplazar(doc, valores):

    for p in doc.paragraphs:
        for run in p.runs:
            texto = run.text
            for k, v in valores.items():
                texto = texto.replace(f"{{{{{k}}}}}", v)
                texto = texto.replace(f"[{k}]", v)
            run.text = texto

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                texto = cell.text
                for k, v in valores.items():
                    texto = texto.replace(f"{{{{{k}}}}}", v)
                    texto = texto.replace(f"[{k}]", v)
                cell.text = texto

# =========================
# ENVIAR CORREO
# =========================

def enviar_correo(destino, docx_file):

    if not EMAIL_USER or not EMAIL_PASS:
        print("No se configuró EMAIL_USER o EMAIL_PASSWORD.")
        return

    msg = EmailMessage()
    msg["Subject"] = "Documento generado"
    msg["From"] = EMAIL_USER
    msg["To"] = destino

    msg.set_content("Adjunto encontrarás el documento generado.")

    with open(docx_file, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="documento.docx"
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(msg)

# =========================
# RUTA PRINCIPAL
# =========================

@app.route("/", methods=["GET", "POST"])
def index():

    variables = obtener_variables()

    if request.method == "POST":

        valores = {}

        for variable in variables:
            valores[variable] = request.form.get(variable, "")

        correo = request.form.get("correo", "").strip()

        doc = Document(PLANTILLA)

        reemplazar(doc, valores)

        docx_path = os.path.join(CARPETA, "documento.docx")

        doc.save(docx_path)

        if correo:
            try:
                enviar_correo(correo, docx_path)
            except Exception as e:
                print("Error enviando correo:", e)

        return send_file(
            docx_path,
            as_attachment=True,
            download_name="documento.docx"
        )

    return render_template(
        "index.html",
        variables=variables
    )

# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
