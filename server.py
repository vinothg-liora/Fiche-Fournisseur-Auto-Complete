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
from openpyxl.utils import column_index_from_string
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


def parse_cell_ref(cell_ref):
    """Parse 'B11' into (col_index, row_index)."""
    import re
    m = re.match(r'^([A-Z]+)(\d+)$', cell_ref)
    if not m:
        return None, None
    return column_index_from_string(m.group(1)), int(m.group(2))


def get_writable_cell(ws, cell_ref):
    """Return the writable cell, handling merged cells.
    If cell_ref is inside a merged range, return the top-left cell."""
    col, row = parse_cell_ref(cell_ref)
    if col is None:
        return ws[cell_ref]
    for merge_range in ws.merged_cells.ranges:
        if (merge_range.min_row <= row <= merge_range.max_row and
                merge_range.min_col <= col <= merge_range.max_col):
            return ws.cell(merge_range.min_row, merge_range.min_col)
    return ws.cell(row, col)


def find_empty_cell(ws, cell_ref):
    """Find the best empty cell to write into.
    If the target cell already has content, search adjacent cells.
    Returns (cell, actual_ref) or (None, None) if no empty cell found."""
    from openpyxl.utils import get_column_letter

    col, row = parse_cell_ref(cell_ref)
    if col is None:
        return None, None

    # First: check the target cell itself
    target = get_writable_cell(ws, cell_ref)
    if target.value is None or str(target.value).strip() == '':
        return target, cell_ref

    # Search order: right, below, left
    search = [
        (row, col + 1),  # right
        (row + 1, col),  # below
        (row, col - 1),  # left (if col > 1)
    ]

    for r, c in search:
        if c < 1 or r < 1:
            continue
        ref = f"{get_column_letter(c)}{r}"
        candidate = get_writable_cell(ws, ref)
        if candidate.value is None or str(candidate.value).strip() == '':
            return candidate, ref

    return None, None

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
# Route — Extract dropdown options from Excel
# ---------------------------------------------------------------------------

@app.route('/api/extract-dropdowns', methods=['POST'])
def extract_dropdowns():
    """
    Receives an Excel file, returns all data validation dropdown options.
    Returns JSON: { "B25": ["Option1", "Option2"], "B30": ["A", "B"] }
    """
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({}), 200

        file_bytes = file.read()
        if len(file_bytes) == 0:
            return jsonify({}), 200

        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
        all_options = {}

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if not hasattr(ws, 'data_validations') or not ws.data_validations:
                continue

            for dv in ws.data_validations.dataValidation:
                if dv.type != 'list' or not dv.formula1:
                    continue

                # Extract options from formula
                options = []
                formula = dv.formula1

                if formula.startswith('"'):
                    # Inline list: "Option1,Option2,Option3"
                    options = [o.strip() for o in formula.strip('"').split(',')]
                elif '!' in formula:
                    # Range reference: Menus!$A$1:$A$10
                    try:
                        ref_sheet, ref_range = formula.split('!')
                        ref_sheet = ref_sheet.strip("'")
                        if ref_sheet in wb.sheetnames:
                            ref_ws = wb[ref_sheet]
                            for row in ref_ws[ref_range.replace('$', '')]:
                                for cell in row:
                                    if cell.value is not None:
                                        options.append(str(cell.value))
                    except Exception as e:
                        print(f"  [dropdown] Error reading range {formula}: {e}")
                else:
                    # Simple range in same sheet: $A$1:$A$10
                    try:
                        for row in ws[formula.replace('$', '')]:
                            for cell in row:
                                if cell.value is not None:
                                    options.append(str(cell.value))
                    except Exception as e:
                        print(f"  [dropdown] Error reading local range {formula}: {e}")

                if not options:
                    continue

                # Map options to all cells in the sqref
                for cell_range in str(dv.sqref).split():
                    if ':' in cell_range:
                        try:
                            for row in ws[cell_range]:
                                for cell in row:
                                    ref = cell.coordinate
                                    all_options[ref] = options
                        except Exception:
                            pass
                    else:
                        all_options[cell_range.replace('$', '')] = options

        wb.close()
        print(f"[extract-dropdowns] Found {len(all_options)} cells with dropdown options")
        return jsonify(all_options)

    except Exception as e:
        print(f"[extract-dropdowns] ERROR: {e}")
        return jsonify({}), 200


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

        # Read the entire file into memory first (Flask stream can be incomplete)
        file_bytes = file.read()
        file_size = len(file_bytes)
        original_name = file.filename or 'document.xlsx'
        print(f"[fill-excel] Received: {original_name} ({file_size} bytes), {len(updates)} updates")

        if file_size == 0:
            return jsonify({'error': 'Empty file received'}), 400

        # Load workbook from BytesIO (not directly from Flask stream)
        file_buffer = io.BytesIO(file_bytes)
        # Only keep_vba for macro-enabled files (.xlsm), not regular .xlsx
        is_macro = original_name.lower().endswith('.xlsm')
        wb = load_workbook(file_buffer, keep_vba=is_macro, data_only=False)

        filled_count = 0
        for u in updates:
            sheet_name = u.get('sheetName', '')
            cell_ref = u.get('cellRef', '')
            value = u.get('value', '')
            valid_options = u.get('validOptions')  # list of valid dropdown values

            if not cell_ref or not value:
                continue

            # Find the target sheet(s)
            target_sheets = []
            if sheet_name:
                if sheet_name in wb.sheetnames:
                    target_sheets.append(sheet_name)
                else:
                    for sn in wb.sheetnames:
                        if sn.lower() == sheet_name.lower():
                            target_sheets.append(sn)
                            break

            if not target_sheets:
                skip = {'menus', 'menu', 'listes', 'lists', 'config'}
                target_sheets = [sn for sn in wb.sheetnames if sn.lower() not in skip]

            for sn in target_sheets:
                ws = wb[sn]

                # Validate dropdown value if options provided
                if valid_options and isinstance(valid_options, list) and len(valid_options) > 0:
                    if value not in valid_options:
                        print(f"  [WARN] {cell_ref} in {sn}: '{value}' not in valid options, writing anyway")

                # Get the writable cell (handles merged cells)
                cell = get_writable_cell(ws, cell_ref)

                # Log what was there before
                old_val = cell.value
                if old_val is not None and str(old_val).strip() != '':
                    print(f"  [OVERWRITE] {sn}!{cell_ref}: '{str(old_val)[:30]}' -> '{value[:50]}'")
                else:
                    print(f"  [OK] {sn}!{cell_ref} = '{value[:50]}'")

                # Write the value — trust Claude's cell identification
                cell.value = value
                filled_count += 1

        # Save to BytesIO buffer
        output = io.BytesIO()
        wb.save(output)
        wb.close()
        output_size = output.tell()
        output.seek(0)

        if output_size == 0:
            return jsonify({'error': 'Generated file is empty'}), 500

        # Determine filename and mimetype
        base, ext = os.path.splitext(original_name)
        if is_macro:
            ext = '.xlsm'
            mime = 'application/vnd.ms-excel.sheet.macroEnabled.12'
        else:
            ext = '.xlsx'
            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        out_name = f"{base}_complete{ext}"

        print(f"[fill-excel] {filled_count} cells written, output={output_size} bytes -> {out_name}")

        return send_file(
            output,
            mimetype=mime,
            as_attachment=True,
            download_name=out_name
        )

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
