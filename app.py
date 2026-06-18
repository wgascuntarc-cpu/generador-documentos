import smtplib
from email.message import EmailMessage
import os
import re
import smtplib
def enviar_correo(archivo):

    remitente = "TU_CORREO@gmail.com"
    contraseña = "TU_CONTRASEÑA_DE_APLICACION"

    destinatario = "geovanyasc@gmail.com"

    mensaje = EmailMessage()

    mensaje["Subject"] = "Documento generado automáticamente"
    mensaje["From"] = remitente
    mensaje["To"] = destinatario

    mensaje.set_content(
        "Se ha generado un nuevo documento desde el sistema."
    )


    with open(archivo, "rb") as f:
        contenido = f.read()

    mensaje.add_attachment(
        contenido,
        maintype="application",
        subtype="docx",
        filename="documento_generado.docx"
    )


    servidor = smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465
    )

    servidor.login(
        remitente,
        contraseña
    )

    servidor.send_message(mensaje)

    servidor.quit()


from flask import Flask, render_template, request, send_file
from docx import Document
from email.message import EmailMessage

app = Flask(__name__)

# =====================================
# RUTAS
# =====================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PLANTILLA = os.path.join(BASE_DIR, "plantilla.docx")

CARPETA = os.path.join(BASE_DIR, "documentos_generados")

os.makedirs(CARPETA, exist_ok=True)

# =====================================
# DATOS DEL CORREO
# =====================================

EMAIL = os.environ.get("EMAIL_USER")
PASSWORD = os.environ.get("EMAIL_PASSWORD")

# =====================================
# OBTENER VARIABLES DE LA PLANTILLA
# =====================================

def obtener_variables():

    if not os.path.exists(PLANTILLA):
        raise FileNotFoundError(
            f"No existe la plantilla:\n{PLANTILLA}"
        )

    doc = Document(PLANTILLA)

    texto = ""

    for p in doc.paragraphs:
        texto += p.text + "\n"

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                texto += celda.text + "\n"

    variables = sorted(
        list(
            set(
                re.findall(r"\{\{(.*?)\}\}", texto)
            )
        )
    )

    return variables

# =====================================
# REEMPLAZAR VARIABLES
# =====================================

def reemplazar(doc, valores):

    for p in doc.paragraphs:

        for clave, valor in valores.items():

            marcador = "{{" + clave + "}}"

            if marcador in p.text:

                p.text = p.text.replace(
                    marcador,
                    valor
                )

    for tabla in doc.tables:

        for fila in tabla.rows:

            for celda in fila.cells:

                for p in celda.paragraphs:

                    for clave, valor in valores.items():

                        marcador = "{{" + clave + "}}"

                        if marcador in p.text:

                            p.text = p.text.replace(
                                marcador,
                                valor
                            )

# =====================================
# ENVIAR CORREO
# =====================================

def enviar_correo(destino, archivo):

    if not EMAIL or not PASSWORD:
        raise Exception(
            "No existen EMAIL_USER o EMAIL_PASSWORD en Render."
        )

    mensaje = EmailMessage()

    mensaje["Subject"] = "Documento generado"

    mensaje["From"] = EMAIL

    mensaje["To"] = destino

    mensaje.set_content(
        "Adjunto encontrará el documento generado."
    )

    with open(archivo, "rb") as f:
        datos = f.read()

    mensaje.add_attachment(
        datos,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="documento.docx"
    )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:

        smtp.login(EMAIL, PASSWORD)

        smtp.send_message(mensaje)

# =====================================
# RUTA PRINCIPAL
# =====================================

@app.route("/", methods=["GET", "POST"])
def index():

    try:

        variables = obtener_variables()

        if request.method == "POST":

            valores = {}

            for variable in variables:
                valores[variable] = request.form.get(variable, "")

            correo = request.form.get("correo", "").strip()

            if not os.path.exists(PLANTILLA):
                raise FileNotFoundError(
                    f"No existe la plantilla:\n{PLANTILLA}"
                )

            doc = Document(PLANTILLA)

            reemplazar(doc, valores)

            docx_path = os.path.join(
                CARPETA,
                "documento.docx"
            )

            doc.save(docx_path)

            if correo != "":

                try:
                    enviar_correo(correo, docx_path)

                except Exception as e:
                    print("ERROR ENVIANDO CORREO:", e)

            enviar_correo(nombre_archivo)
            return send_file(
                docx_path,
                as_attachment=True,
                download_name="documento.docx"
            )

        return render_template(
            "index.html",
            variables=variables
        )

    except Exception as e:

        import traceback

        return f"""
        <h2>Error encontrado</h2>

        <pre>{e}</pre>

        <pre>{traceback.format_exc()}</pre>

        """, 500

# =====================================
# INICIO
# =====================================

if __name__ == "__main__":
    app.run(debug=True)
