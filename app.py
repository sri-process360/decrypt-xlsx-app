from flask import Flask, request, jsonify
import msoffcrypto
import openpyxl
from io import BytesIO
import os

app = Flask(__name__)

@app.route('/decrypt', methods=['POST'])
def decrypt_xlsx():
    data = request.get_json()
    base64_data = data.get('fileBase64', '')
    password = data.get('password', '')

    if not base64_data or not password:
        return jsonify({'error': 'Missing base64 data or password'}), 400

    try:
        file_bytes = BytesIO(base64.b64decode(base64_data))
        decrypted_file = BytesIO()
        with msoffcrypto.OfficeFile(file_bytes) as encrypted:
            encrypted.load_key(password=password)
            encrypted.decrypt(decrypted_file)

        workbook = openpyxl.load_workbook(decrypted_file)
        worksheet = workbook.active
        data = [[cell.value for cell in row] for row in worksheet.rows]

        return jsonify({'data': data}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # Use Render's PORT or default to 5000
    app.run(host='0.0.0.0', port=port)
