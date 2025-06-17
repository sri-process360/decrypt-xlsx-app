from flask import Flask, request, jsonify
import base64
import io
import msoffcrypto
import openpyxl
import logging

app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

@app.route('/decrypt', methods=['POST'])
def decrypt():
    try:
        data = request.get_json()
        if not data or 'fileBase64' not in data or 'password' not in data:
            logger.error("Missing fileBase64 or password in request")
            return jsonify({'error': 'Missing base64 data or password'}), 400

        file_data = base64.b64decode(data['fileBase64'])
        password = data['password']

        # Decrypt the file
        file = io.BytesIO(file_data)
        decrypted = io.BytesIO()
        with msoffcrypto.OfficeFile(file) as office_file:
            if not office_file.is_encrypted():
                logger.error("File is not encrypted")
                return jsonify({'error': 'File is not encrypted'}), 400
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)

        # Read the decrypted Excel file
        decrypted.seek(0)
        workbook = openpyxl.load_workbook(decrypted)
        sheet = workbook.active
        data = [[cell.value for cell in row] for row in sheet.iter_rows()]

        logger.info("Decryption and data extraction successful")
        return jsonify({'data': data})

    except base64.binascii.Error as e:
        logger.error(f"Base64 decoding error: {str(e)}")
        return jsonify({'error': 'Invalid base64 data'}), 400
    except msoffcrypto.exceptions.InvalidFileError as e:
        logger.error(f"Invalid file format: {str(e)}")
        return jsonify({'error': 'Invalid or corrupt Excel file'}), 400
    except msoffcrypto.exceptions.DecryptionError as e:
        logger.error(f"Decryption failed: {str(e)}")
        return jsonify({'error': 'Incorrect password or file format'}), 400
    except MemoryError as e:
        logger.error(f"Memory error: {str(e)}")
        return jsonify({'error': 'File too large for server'}), 400
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        return jsonify({'error': 'Internal server error'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
