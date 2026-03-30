"""
Liora — Fiche Fournisseur Auto-Complete
Serveur Flask local (packagé avec PyInstaller)
"""

import io
import json
import os
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ---------------------------------------------------------------------------
# Paths — works both in dev and inside PyInstaller bundle
# ---------------------------------------------------------------------------

def get_base_path():
    """Return the base path for bundled resources."""
    if getattr(sys, 'frozen', False):
        # Running inside PyInstaller bundle
        return Path(sys._MEIPASS)
    # Running in dev — files are next to server.py
    return Path(__file__).resolve().parent


BASE_PATH = get_base_path()
print(f"[startup] BASE_PATH = {BASE_PATH}")
print(f"[startup] index.html exists = {(BASE_PATH / 'index.html').exists()}")

app = Flask(__name__,
            static_folder=str(BASE_PATH),
            static_url_path='',
            template_folder=str(BASE_PATH))
CORS(app)


# ---------------------------------------------------------------------------
# Routes — Static files (serve ALL files from BASE_PATH)
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return send_from_directory(str(BASE_PATH), 'index.html')


@app.route('/<path:filename>')
def static_files(filename):
    filepath = BASE_PATH / filename
    if filepath.exists() and filepath.is_file():
        return send_from_directory(str(BASE_PATH), filename)
    return 'Not Found', 404


# ---------------------------------------------------------------------------
# Route — Fill Excel (.xlsx / .xls)
# ---------------------------------------------------------------------------

@app.route('/api/fill-excel', methods=['POST'])
def fill_excel():
    """
    Receives:
      - file: the original Excel file (multipart)
      - updates: JSON string — array of {sheetName, cellRef, value}
    Returns:
      - The modified Excel file (download)
    """
    try:
        file = request.files.get('file')
        updates_json = request.form.get('updates', '[]')
        updates = json.loads(updates_json)

        if not file:
            return jsonify({'error': 'No file provided'}), 400

        # Load workbook preserving everything
        wb = load_workbook(file, keep_vba=True, data_only=False)

        filled_count = 0
        for u in updates:
            sheet_name = u.get('sheetName', '')
            cell_ref = u.get('cellRef', '')
            value = u.get('value', '')

            if not cell_ref or not value:
                continue

            # Find the sheet
            target_sheets = []
            if sheet_name:
                # Try exact match
                if sheet_name in wb.sheetnames:
                    target_sheets.append(sheet_name)
                else:
                    # Case-insensitive match
                    for sn in wb.sheetnames:
                        if sn.lower() == sheet_name.lower():
                            target_sheets.append(sn)
                            break

            if not target_sheets:
                # Write to all non-menu sheets
                skip = {'menus', 'menu', 'listes', 'lists', 'config'}
                target_sheets = [sn for sn in wb.sheetnames if sn.lower() not in skip]

            for sn in target_sheets:
                ws = wb[sn]
                ws[cell_ref] = value
                filled_count += 1

        # Save to buffer
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # Determine filename
        original_name = file.filename or 'document.xlsx'
        base, ext = os.path.splitext(original_name)
        if ext.lower() not in ('.xlsx', '.xlsm'):
            ext = '.xlsx'
        out_name = f"{base}_complété{ext}"

        print(f"[fill-excel] {filled_count} cells written → {out_name}")
        return send_file(output, as_attachment=True, download_name=out_name,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        print(f"[fill-excel] ERROR: {e}")
        return jsonify({'error': str(e)}), 500


# ---------------------------------------------------------------------------
# Route — Fill PDF form
# ---------------------------------------------------------------------------

@app.route('/api/fill-pdf', methods=['POST'])
def fill_pdf():
    """
    Receives:
      - file: the original PDF file
      - updates: JSON string — array of {fieldName, value}
    Returns:
      - The modified PDF file
    """
    try:
        file = request.files.get('file')
        updates_json = request.form.get('updates', '[]')
        updates = json.loads(updates_json)

        if not file:
            return jsonify({'error': 'No file provided'}), 400

        reader = PdfReader(file)
        writer = PdfWriter()
        writer.append_pages_from_reader(reader)

        # Fill form fields
        filled = 0
        for u in updates:
            field_name = u.get('fieldName', '')
            value = u.get('value', '')
            if field_name and value:
                try:
                    writer.update_page_form_field_values(None, {field_name: value})
                    filled += 1
                except Exception:
                    pass

        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        original_name = file.filename or 'document.pdf'
        base, _ = os.path.splitext(original_name)
        out_name = f"{base}_complété.pdf"

        print(f"[fill-pdf] {filled} fields written → {out_name}")
        return send_file(output, as_attachment=True, download_name=out_name,
                         mimetype='application/pdf')

    except Exception as e:
        print(f"[fill-pdf] ERROR: {e}")
        return jsonify({'error': str(e)}), 500


# ---------------------------------------------------------------------------
# Route — Generate recap PDF
# ---------------------------------------------------------------------------

@app.route('/api/recap-pdf', methods=['POST'])
def recap_pdf():
    """
    Receives JSON body:
      { fileName, companyName, fields: [{label, value}], contact: {name, role, email, phone} }
    Returns: a styled recap PDF
    """
    try:
        data = request.get_json()
        fields = data.get('fields', [])
        company = data.get('companyName', '')
        file_name = data.get('fileName', 'document')
        contact = data.get('contact', {})

        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4,
                                leftMargin=25 * mm, rightMargin=25 * mm,
                                topMargin=20 * mm, bottomMargin=20 * mm)

        styles = getSampleStyleSheet()
        orange = HexColor('#F47458')
        dark = HexColor('#1a1a2e')

        title_style = ParagraphStyle('Title', parent=styles['Heading1'],
                                     textColor=orange, fontSize=16, spaceAfter=6)
        sub_style = ParagraphStyle('Sub', parent=styles['Normal'],
                                   textColor=HexColor('#666666'), fontSize=9)
        label_style = ParagraphStyle('Label', parent=styles['Normal'],
                                     fontSize=9, textColor=HexColor('#333333'),
                                     fontName='Helvetica-Bold')
        value_style = ParagraphStyle('Value', parent=styles['Normal'],
                                     fontSize=9, textColor=dark)

        elements = []
        elements.append(Paragraph(f"FICHE FOURNISSEUR — {company}", title_style))
        elements.append(Paragraph(
            f"Fichier : {file_name} — {time.strftime('%d/%m/%Y')}", sub_style))
        elements.append(Spacer(1, 10))

        # Build table
        table_data = []
        for f in fields:
            label = f.get('label', '')
            value = f.get('value', '')
            table_data.append([
                Paragraph(label, label_style),
                Paragraph(value or '—', value_style)
            ])

        if table_data:
            t = Table(table_data, colWidths=[200, 300])
            t.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('LINEBELOW', (0, 0), (-1, -1), 0.5, HexColor('#eeeeee')),
            ]))
            elements.append(t)

        # Signature
        elements.append(Spacer(1, 20))
        if contact:
            sig = f"<b>{contact.get('name', '')}</b> — {contact.get('role', '')}<br/>"
            sig += f"{contact.get('email', '')} | {contact.get('phone', '')}"
            elements.append(Paragraph(sig, sub_style))

        doc.build(elements)
        output.seek(0)

        base, _ = os.path.splitext(file_name)
        return send_file(output, as_attachment=True,
                         download_name=f"{base}_recap.pdf",
                         mimetype='application/pdf')

    except Exception as e:
        print(f"[recap-pdf] ERROR: {e}")
        return jsonify({'error': str(e)}), 500


# ---------------------------------------------------------------------------
# Route — Health check
# ---------------------------------------------------------------------------

@app.route('/api/health')
def health():
    return jsonify({'status': 'ok', 'version': '1.0.0'})


# ---------------------------------------------------------------------------
# Server startup
# ---------------------------------------------------------------------------

def find_free_port(start=5000):
    """Find the first available port starting from `start`."""
    for port in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('127.0.0.1', port))
                return port
            except OSError:
                continue
    return start


def open_browser(port):
    """Open the app in the default browser after a short delay."""
    time.sleep(1.2)
    webbrowser.open(f'http://127.0.0.1:{port}')


def main():
    port = find_free_port(5000)
    print(f"""
    ╔══════════════════════════════════════════╗
    ║   Liora — Fiche Fournisseur             ║
    ║   Auto-Complete v1.0                    ║
    ║                                          ║
    ║   Serveur démarré sur le port {port}       ║
    ║   http://127.0.0.1:{port}                  ║
    ╚══════════════════════════════════════════╝
    """)

    # Open browser in a thread so it doesn't block Flask
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()

    # Run Flask (use_reloader=False is critical for PyInstaller)
    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)


if __name__ == '__main__':
    main()
