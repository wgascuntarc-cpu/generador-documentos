import os
import re
import smtplib

from flask import Flask, render_template, request, send_file
from docx import Document
from email.message import EmailMessage
from dotenv import load_dotenv


# Cargar archivo .env
load_dotenv()


app = Flask(__name__)


# ==============================
# CONFIGURACIÓN ARCHIVOS
# ==============================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PLANTILLA = os.path.join(
    BASE_DIR,
    "plantilla.docx"
)

CARPETA_SALIDA = os.path.join(
    BASE_DIR,
    "documentos_generados"
)

os.makedirs(
    CARPETA_SALIDA,
    exist_ok=True
)


# ==============================
# CONFIGURACIÓN CORREO
# ==============================

EMAIL_USER = os.getenv(
    "EMAIL_USER"
)

EMAIL_PASSWORD = os.getenv(
    "EMAIL_PASSWORD"
)


CORREO_DESTINO = "geovanyasc@gmail.com"


print("================================")
print("Correo configurado:")
print(EMAIL_USER)
print("================================")



# ==============================
# LEER CAMPOS DEL WORD
# ==============================

def obtener_variables():

    doc = Document(
        PLANTILLA
    )

    texto = ""


    for p in doc.paragraphs:

        texto += p.text + "\n"



    for tabla in doc.tables:

        for fila in tabla.rows:

            for celda in fila.cells:

                texto += celda.text + "\n"



    variables = re.findall(
        r"\{\{(.*?)\}\}",
        texto
    )


    return sorted(
        set(variables)
    )



# ==============================
# CAMBIAR CAMPOS
# ==============================

def reemplazar_variables(doc, datos):


    for p in doc.paragraphs:

        for clave, valor in datos.items():

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


                    for clave, valor in datos.items():

                        marcador = "{{" + clave + "}}"


                        if marcador in p.text:

                            p.text = p.text.replace(
                                marcador,
                                valor
                            )



# ==============================
# ENVIAR CORREO
# ==============================

def enviar_correo(archivo):

    print("ENTRO A ENVIAR CORREO")


    if not EMAIL_USER or not EMAIL_PASSWORD:

        print("Faltan datos del correo")

        return



    mensaje = EmailMessage()


    mensaje["Subject"] = (
        "Documento generado automáticamente"
    )


    mensaje["From"] = EMAIL_USER


    mensaje["To"] = CORREO_DESTINO



    mensaje.set_content(
        "Se adjunta el documento generado."
    )



    with open(archivo, "rb") as f:

        contenido = f.read()



    mensaje.add_attachment(
        contenido,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="documento.docx"
    )



    print("Conectando con Gmail...")



    try:

        with smtplib.SMTP_SSL(
            "smtp.gmail.com",
            465
        ) as servidor:


            servidor.login(
                EMAIL_USER,
                EMAIL_PASSWORD
            )


            servidor.send_message(
                mensaje
            )


        print("CORREO ENVIADO CORRECTAMENTE")


    except Exception as error:

        print(
            "ERROR ENVIANDO CORREO:",
            error
        )



# ==============================
# PAGINA PRINCIPAL
# ==============================

@app.route(
    "/",
    methods=["GET","POST"]
)

def index():


    variables = obtener_variables()



    if request.method == "POST":


        datos = {}



        for variable in variables:

            datos[variable] = request.form.get(
                variable,
                ""
            )



        documento = Document(
            PLANTILLA
        )



        reemplazar_variables(
            documento,
            datos
        )



        archivo = os.path.join(
            CARPETA_SALIDA,
            "documento.docx"
        )



        documento.save(
            archivo
        )


        print("DOCUMENTO CREADO:", archivo)



        # ENVIO AUTOMATICO

        enviar_correo(
            archivo
        )



        return send_file(
            archivo,
            as_attachment=True,
            download_name="documento.docx"
        )



    return render_template(
        "index.html",
        variables=variables
    )



# ==============================
# INICIAR
# ==============================

if __name__ == "__main__":

    app.run(
        debug=True
    )
