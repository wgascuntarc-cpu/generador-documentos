import os
import re
import smtplib

from flask import Flask, render_template, request, send_file
from docx import Document
from email.message import EmailMessage
from dotenv import load_dotenv


# Cargar variables .env
load_dotenv()


app = Flask(__name__)


# =====================================
# CONFIGURACIÓN
# =====================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PLANTILLA = os.path.join(
    BASE_DIR,
    "plantilla.docx"
)

CARPETA = os.path.join(
    BASE_DIR,
    "documentos_generados"
)

os.makedirs(
    CARPETA,
    exist_ok=True
)


# =====================================
# DATOS CORREO
# =====================================

EMAIL = os.environ.get("EMAIL_USER")

PASSWORD = os.environ.get("EMAIL_PASSWORD")


DESTINO = "geovanyasc@gmail.com"


# =====================================
# OBTENER VARIABLES WORD
# =====================================

def obtener_variables():

    doc = Document(PLANTILLA)

    texto = ""


    for p in doc.paragraphs:
        texto += p.text + "\n"


    for tabla in doc.tables:

        for fila in tabla.rows:

            for celda in fila.cells:

                texto += celda.text + "\n"


    variables = sorted(
        set(
            re.findall(
                r"\{\{(.*?)\}\}",
                texto
            )
        )
    )


    return variables



# =====================================
# REEMPLAZAR CAMPOS
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
# ENVIAR DOCUMENTO POR CORREO
# =====================================

def enviar_correo(archivo):


    if not EMAIL or not PASSWORD:

        raise Exception(
            "Faltan EMAIL_USER o EMAIL_PASSWORD"
        )


    mensaje = EmailMessage()


    mensaje["Subject"] = (
        "Documento generado automáticamente"
    )


    mensaje["From"] = EMAIL


    mensaje["To"] = DESTINO


    mensaje.set_content(
        """
        Se ha generado un nuevo documento.

        Adjunto encontrará el archivo.
        """
    )


    with open(archivo, "rb") as f:

        contenido = f.read()


    mensaje.add_attachment(
        contenido,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="documento.docx"
    )


    with smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465
    ) as smtp:


        smtp.login(
            EMAIL,
            PASSWORD
        )


        smtp.send_message(
            mensaje
        )



# =====================================
# PÁGINA PRINCIPAL
# =====================================

@app.route(
    "/",
    methods=["GET", "POST"]
)
def index():


    try:

        variables = obtener_variables()


        if request.method == "POST":


            valores = {}


            for variable in variables:

                valores[variable] = request.form.get(
                    variable,
                    ""
                )



            doc = Document(
                PLANTILLA
            )


            reemplazar(
                doc,
                valores
            )



            docx_path = os.path.join(
                CARPETA,
                "documento.docx"
            )



            doc.save(
                docx_path
            )



            # ENVIAR AUTOMÁTICAMENTE
            try:

                enviar_correo(
                    docx_path
                )

                print(
                    "Correo enviado correctamente"
                )


            except Exception as e:

                print(
                    "ERROR ENVIANDO CORREO:",
                    e
                )



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

        """,500




# =====================================
# INICIO
# =====================================

if __name__ == "__main__":

    app.run(
        debug=True
    )
