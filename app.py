from flask import Flask, send_file, request, jsonify
from flask_cors import CORS
import pandas as pd
import openpyxl
import io

app = Flask(__name__)
CORS(app)  # GitHub Pages se request allow karne ke liye

MASTER_SHEET_ID = "1J9WXMfNwVLTQIp0avtglFzpzHmHNqcLvNz0mxscrh2U"

@app.route('/download-excel', methods=['POST'])
def download_excel():
    try:
        data = request.get_json()
        selected_tabs = data.get('items', [])
        
        if not selected_tabs:
            return jsonify({"status": "error", "message": "No tabs selected"}), 400

        # Ek nayi Excel file banayein jisme multiple sheets hongi
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for tab_name in selected_tabs:
                # Google Sheet ke public export URL se specific tab ka data CSV format me uthayein
                csv_url = f"https://docs.google.com/spreadsheets/d/{MASTER_SHEET_ID}/gviz/tq?tqx=out:csv&sheet={tab_name}"
                try:
                    df = pd.read_csv(csv_url)
                    # Data ko naye Excel sheet me daalein
                    df.to_excel(writer, sheet_name=tab_name, index=False)
                except Exception as e:
                    # Agar koi tab na mile ya error aaye toh empty sheet daal dein
                    df = pd.DataFrame({"Error": [f"Could not load tab: {str(e)}"]})
                    df.to_excel(writer, sheet_name=tab_name, index=False)
        
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Era_Inspection_Report.xlsx'
        )

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
