from flask import Flask, request, send_file
import os
from excel_generator import generate_excel

app = Flask(__name__)

@app.route('/generate_excel', methods=['POST'])
def generate():
    data = request.json  # Recibe datos en formato JSON desde la web
    if not data:
        return {"error": "No se enviaron datos"}, 400

    filename = "generated_report.xlsx"
    file_path = os.path.join(os.getcwd(), filename)

    # Generar el archivo Excel
    generate_excel(data, filename=filename)

    # Enviar el archivo al usuario
    return send_file(file_path, as_attachment=True)

# No ejecutamos Flask con app.run() en producción

@app.route('/routes', methods=['GET'])
def list_routes():
    import urllib
    output = []
    for rule in app.url_map.iter_rules():
        methods = ','.join(rule.methods)
        output.append(f"{rule.endpoint}: {rule.rule} [{methods}]")
    return '<br>'.join(output)
