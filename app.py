import os
import re
import smtplib

from flask import Flask, render_template, request, send_file
from docx import Document
from email.message import EmailMessage
from dotenv import load_dotenv


# Cargar variables del archivo .env
load_dotenv()


app = Flask(__name__)


# =====================================
# CONFIGURACIÓN DE ARCHIVOS
# =====================================

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



# =====================================
# CONFIGURACIÓN CORREO
# =====================================

EMAIL = os.environ.get("EMAIL_USER")

PASSWORD = os.environ.get("EMAIL_PASSWORD")


DESTINO = "geovanyasc@gmail.com"



print("Correo configurado:", EMAIL)



# =====================================
# LEER VARIABLES DEL WORD
# =====================================

def obtener_variables():

    doc = Document(PLANTILLA)

    texto = ""


    for parrafo in doc.paragraphs:

        texto += parrafo.text + "\n"



    for tabla in doc.tables:

        for fila in tabla.rows:

            for celda in fila.cells:

                texto += celda.text + "\n"



    variables = re.findall(
        r"\{\{(.*?)\}\}",
        texto
    )


    return sorted(set(variables))



# =====================================
# CAMBIAR DATOS EN WORD
# =====================================

def reemplazar_variables(doc, datos):


    for parrafo in doc.paragraphs:

        for clave, valor in datos.items():

            marcador = "{{" + clave + "}}"


            if marcador in parrafo.text:

                parrafo.text = parrafo.text.replace(
                    marcador,
                    valor
                )



    for tabla in doc.tables:

        for fila in tabla.rows:

            for celda in fila.cells:


                for parrafo in celda.paragraphs:


                    for clave, valor in datos.items():

                        marcador = "{{" + clave + "}}"


                        if marcador in parrafo.text:

                            parrafo.text = parrafo.text.replace(
                                marcador,
                                valor
                            )



# =====================================
# ENVIAR CORREO
# =====================================

def enviar_correo(archivo):


    print("=== INICIANDO ENVIO ===")


    print("Desde:", EMAIL)

    print("Destino:", DESTINO)



    if not EMAIL or not PASSWORD:

        raise Exception(
            "No existe EMAIL_USER o EMAIL_PASSWORD"
        )



    mensaje = EmailMessage()


    mensaje["Subject"] = (
        "Documento generado automáticamente"
    )


    mensaje["From"] = EMAIL


    mensaje["To"] = DESTINO



    mensaje.set_content(
        "Se adjunta documento generado por el sistema."
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



    with smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465
    ) as servidor:


        servidor.login(
            EMAIL,
            PASSWORD
        )


        servidor.send_message(
            mensaje
        )



    print("=== CORREO ENVIADO CORRECTAMENTE ===")



# =====================================
# PAGINA PRINCIPAL
# =====================================

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



        # ENVIAR CORREO AUTOMÁTICO

        try:

            enviar_correo(
                archivo
            )


        except Exception as error:

            print(
                "ERROR DEL CORREO:",
                error
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



# =====================================
# EJECUTAR
# =====================================

if __name__ == "__main__":

    app.run(
        debug=True
    )
