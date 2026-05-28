
import os
import re
from flask import Flask, render_template, request, send_file
from docx import Document

# =========================================
# CREAR APP FLASK
# =========================================

app = Flask(__name__)

# =========================================
# CONFIGURACIÓN
# =========================================

PLANTILLA = "plantilla.docx"

CARPETA_SALIDA = "documentos_generados"

# Crear carpeta si no existe
if not os.path.exists(CARPETA_SALIDA):

    os.makedirs(CARPETA_SALIDA)

# =========================================
# LEER VARIABLES DEL WORD
# =========================================

def obtener_variables():

    doc = Document(PLANTILLA)

    variables = set()

    # Buscar en párrafos
    for p in doc.paragraphs:

        texto = p.text

        vars1 = re.findall(
            r"\{\{(.*?)\}\}",
            texto
        )

        vars2 = re.findall(
            r"\[(.*?)\]",
            texto
        )

        variables.update(vars1)
        variables.update(vars2)

    # Buscar en tablas
    for tabla in doc.tables:

        for fila in tabla.rows:

            for celda in fila.cells:

                texto = celda.text

                vars1 = re.findall(
                    r"\{\{(.*?)\}\}",
                    texto
                )

                vars2 = re.findall(
                    r"\[(.*?)\]",
                    texto
                )

                variables.update(vars1)
                variables.update(vars2)

    return list(variables)

# =========================================
# PÁGINA PRINCIPAL
# =========================================

@app.route("/", methods=["GET", "POST"])

def index():

    variables = obtener_variables()

    # =====================================
    # SI ENVÍAN EL FORMULARIO
    # =====================================

    if request.method == "POST":

        nuevo_doc = Document(PLANTILLA)

        valores = {}

        # Obtener datos del formulario
        for variable in variables:

            valores[variable] = request.form.get(variable)

        # =================================
        # REEMPLAZAR EN PÁRRAFOS
        # =================================

        for p in nuevo_doc.paragraphs:

            for variable, valor in valores.items():

                p.text = p.text.replace(
                    f"{{{{{variable}}}}}",
                    valor
                )

                p.text = p.text.replace(
                    f"[{variable}]",
                    valor
                )

        # =================================
        # REEMPLAZAR EN TABLAS
        # =================================

        for tabla in nuevo_doc.tables:

            for fila in tabla.rows:

                for celda in fila.cells:

                    for variable, valor in valores.items():

                        celda.text = celda.text.replace(
                            f"{{{{{variable}}}}}",
                            valor
                        )

                        celda.text = celda.text.replace(
                            f"[{variable}]",
                            valor
                        )

        # =================================
        # GUARDAR DOCUMENTO
        # =================================

        nombre_archivo = "documento_generado.docx"

        ruta_guardado = os.path.join(
            CARPETA_SALIDA,
            nombre_archivo
        )

        nuevo_doc.save(ruta_guardado)

        # =================================
        # DESCARGAR ARCHIVO
        # =================================

        return send_file(
            ruta_guardado,
            as_attachment=True
        )

    # =====================================
    # MOSTRAR FORMULARIO
    # =====================================

    return render_template(
        "index.html",
        variables=variables
    )

# =========================================
# EJECUTAR SERVIDOR
# =========================================

if __name__ == "__main__":

    app.run(
        debug=True
    )

