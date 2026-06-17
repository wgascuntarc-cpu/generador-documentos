import os
import re
import subprocess
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
            for k, v in valores.items():
                run.text = run.text.replace(f"{{{{{k}}}}}", v)
                run.text = run.text.replace(f"[{k}]", v)

    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for k, v in valores.items():
                    cell.text = cell.text.replace(f"{{{{{k}}}}}", v)
                    cell.text = cell.text.replace(f"[{k}]", v)


# =========================
# CONVERTIR A PDF (RENDER)
# =========================

def convertir_pdf(docx_path):

    subprocess.run([
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        docx_path,
        "--outdir",
        CARPETA
    ], check=True)

    return docx_path.replace(".docx", ".pdf")


# =========================
# ENVIAR CORREO
# =========================

def enviar_correo(destino, docx_file, pdf_file):

    msg = EmailMessage()
    msg["Subject"] = "Documento generado"
    msg["From"] = EMAIL_USER
    msg["To"] = destino

    msg.set_content("Adjunto encontrarás el documento en Word y PDF.")

    # DOCX
    with open(docx_file, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="documento.docx"
        )

    # PDF
    with open(pdf_file, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename="documento.pdf"
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

        valores = {
            v: request.form.get(v, "")
            for v in variables
        }

        correo = request.form.get("correo")

        doc = Document(PLANTILLA)
        reemplazar(doc, valores)

        docx_path = os.path.join(CARPETA, "documento.docx")
        doc.save(docx_path)

        pdf_path = convertir_pdf(docx_path)

        if correo:
            enviar_correo(correo, docx_path, pdf_path)

        return send_file(pdf_path, as_attachment=True)

    


# =========================
# RUN
# =========================

if __name__ == "__main__":
    app.run(debug=True)
