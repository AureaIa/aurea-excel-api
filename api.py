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

if __name__ == '__main__':
    app.run(debug=True, port=5000)
