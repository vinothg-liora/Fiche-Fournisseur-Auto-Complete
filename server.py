"""
Liora — Fiche Fournisseur Auto-Complete
Serveur Flask local (packagé avec PyInstaller)
"""

import io
import json
import os
import re
import socket
import sys
import threading
import time
import webbrowser
from pathlib import Path

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

def get_base_path():
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
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
# Excel helpers
# ---------------------------------------------------------------------------

def parse_cell_ref(cell_ref):
    """Parse 'B11' into (col_index, row_index). Returns (None, None) on failure."""
    m = re.match(r'^([A-Z]+)(\d+)$', cell_ref.upper())
    if not m:
        return None, None
    return column_index_from_string(m.group(1)), int(m.group(2))


def get_writable_cell(ws, row, col):
    """Return the writable cell, redirecting to merge master if needed."""
    for merge_range in ws.merged_cells.ranges:
        if (merge_range.min_row <= row <= merge_range.max_row and
                merge_range.min_col <= col <= merge_range.max_col):
            return ws.cell(merge_range.min_row, merge_range.min_col)
    return ws.cell(row, col)


def find_target_cell(ws, row, col):
    """Find the best empty cell to write into.
    Rule 1: target cell itself if empty
    Rule 2: search right (up to 10 cols)
    Rule 3: search below (up to 5 rows)
    Rule 4: no empty cell found -> return None
    """
    # Rule 1: target cell
    cell = get_writable_cell(ws, row, col)
    if cell.value is None or str(cell.value).strip() == '':
        return cell

    # Rule 2: search right
    for c in range(col + 1, col + 10):
        candidate = get_writable_cell(ws, row, c)
        if candidate.value is None or str(candidate.value).strip() == '':
            return candidate

    # Rule 3: search below
    for r in range(row + 1, row + 5):
        candidate = get_writable_cell(ws, r, col)
        if candidate.value is None or str(candidate.value).strip() == '':
            return candidate

    # Rule 4: no empty cell
    ref = f"{get_column_letter(col)}{row}"
    print(f"  [WARNING] No empty cell found near {ref}")
    return None


def get_dropdown_options(ws, cell_ref):
    """Read data validation dropdown options for a specific cell."""
    options = []
    if not hasattr(ws, 'data_validations') or not ws.data_validations:
        return options

    for dv in ws.data_validations.dataValidation:
        if cell_ref not in str(dv.sqref):
            continue
        if dv.type != 'list' or not dv.formula1:
            continue

        formula = dv.formula1

        if formula.startswith('"'):
            # Inline list: "Option1,Option2,Option3"
            options = [o.strip() for o in formula.strip('"').split(',')]
        elif '!' in formula:
            # Range in another sheet: Menus!$A$1:$A$10
            try:
                sheet_name, range_ref = formula.split('!')
                sheet_name = sheet_name.strip("'")
                ref_ws = ws.parent[sheet_name]
                clean_range = range_ref.replace('$', '')
                for row in ref_ws[clean_range]:
                    for c in row:
                        if c.value is not None:
                            options.append(str(c.value))
            except Exception as e:
                print(f"  [dropdown] Error reading {formula}: {e}")
        else:
            # Range in same sheet: $A$1:$A$10
            try:
                clean_range = formula.replace('$', '')
                for row in ws[clean_range]:
                    for c in row:
                        if c.value is not None:
                            options.append(str(c.value))
            except Exception as e:
                print(f"  [dropdown] Error reading {formula}: {e}")
        break

    return options


def extract_all_dropdowns(wb):
    """Extract all dropdown options from all sheets. Returns {cellRef: [options]}."""
    all_options = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if not hasattr(ws, 'data_validations') or not ws.data_validations:
            print(f"[dropdown] Sheet '{sheet_name}': no data validations")
            continue

        dvs = ws.data_validations.dataValidation
        print(f"[dropdown] Sheet '{sheet_name}': {len(dvs)} validation(s) found")

        for dv in dvs:
            print(f"[dropdown]   type={dv.type} formula1={dv.formula1} sqref={dv.sqref}")
            if dv.type != 'list' or not dv.formula1:
                continue

            formula = dv.formula1
            options = []

            if formula.startswith('"'):
                options = [o.strip() for o in formula.strip('"').split(',')]
            elif '!' in formula:
                try:
                    ref_sheet, ref_range = formula.split('!')
                    ref_sheet = ref_sheet.strip("'")
                    if ref_sheet in wb.sheetnames:
                        ref_ws = wb[ref_sheet]
                        for row in ref_ws[ref_range.replace('$', '')]:
                            for c in row:
                                if c.value is not None:
                                    options.append(str(c.value))
                except Exception as e:
                    print(f"  [dropdown] Error: {e}")
            else:
                try:
                    for row in ws[formula.replace('$', '')]:
                        for c in row:
                            if c.value is not None:
                                options.append(str(c.value))
                except Exception as e:
                    print(f"  [dropdown] Error: {e}")

            if not options:
                continue

            for cell_range in str(dv.sqref).split():
                if ':' in cell_range:
                    try:
                        for row in ws[cell_range.replace('$', '')]:
                            for c in row:
                                all_options[c.coordinate] = options
                    except Exception:
                        pass
                else:
                    all_options[cell_range.replace('$', '')] = options

    return all_options


# ---------------------------------------------------------------------------
# Routes — Static files
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
# Route — Extract dropdown options
# ---------------------------------------------------------------------------

@app.route('/api/extract-dropdowns', methods=['POST'])
def api_extract_dropdowns():
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({}), 200

        file_bytes = file.read()
        print(f"[extract-dropdowns] {file.filename} ({len(file_bytes)} bytes)")

        if len(file_bytes) == 0:
            return jsonify({}), 200

        wb = load_workbook(io.BytesIO(file_bytes), data_only=False)
        result = extract_all_dropdowns(wb)
        wb.close()

        print(f"[extract-dropdowns] Found {len(result)} cells with dropdown options")
        for ref, opts in list(result.items())[:5]:
            print(f"  {ref}: {opts[:3]}{'...' if len(opts) > 3 else ''}")

        return jsonify(result)

    except Exception as e:
        print(f"[extract-dropdowns] ERROR: {e}")
        return jsonify({}), 200


# ---------------------------------------------------------------------------
# Route — Fill Excel
# ---------------------------------------------------------------------------

@app.route('/api/fill-excel', methods=['POST'])
def fill_excel():
    try:
        file = request.files.get('file')
        updates_json = request.form.get('updates', '[]')
        updates = json.loads(updates_json)

        if not file:
            return jsonify({'error': 'No file provided'}), 400

        file_bytes = file.read()
        original_name = file.filename or 'document.xlsx'
        print(f"[fill-excel] Received: {original_name} ({len(file_bytes)} bytes), {len(updates)} updates")

        if len(file_bytes) == 0:
            return jsonify({'error': 'Empty file received'}), 400

        buf = io.BytesIO(file_bytes)
        is_macro = original_name.lower().endswith('.xlsm')
        wb = load_workbook(buf, keep_vba=is_macro, data_only=False)

        filled_count = 0
        for u in updates:
            sheet_name = u.get('sheetName', '')
            cell_ref = u.get('cellRef', '')
            value = u.get('value', '')

            if not cell_ref or not value:
                continue

            col, row = parse_cell_ref(cell_ref)
            if col is None:
                print(f"  [SKIP] Invalid cell ref: {cell_ref}")
                continue

            # Find target sheet(s)
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
                target = find_target_cell(ws, row, col)

                if target is None:
                    continue

                actual_ref = target.coordinate
                if actual_ref != cell_ref:
                    print(f"  [REDIRECT] {cell_ref} -> {actual_ref} = '{value[:50]}'")
                else:
                    print(f"  [OK] {sn}!{actual_ref} = '{value[:50]}'")

                target.value = value
                filled_count += 1

        # Save
        output = io.BytesIO()
        wb.save(output)
        wb.close()
        output_size = output.tell()
        output.seek(0)

        if output_size == 0:
            return jsonify({'error': 'Generated file is empty'}), 500

        base, ext = os.path.splitext(original_name)
        if is_macro:
            ext = '.xlsm'
            mime = 'application/vnd.ms-excel.sheet.macroEnabled.12'
        else:
            ext = '.xlsx'
            mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        out_name = f"{base}_complete{ext}"

        print(f"[fill-excel] {filled_count} cells written, output={output_size} bytes -> {out_name}")

        return send_file(output, mimetype=mime, as_attachment=True, download_name=out_name)

    except Exception as e:
        print(f"[fill-excel] ERROR: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# ---------------------------------------------------------------------------
# Route — Fill PDF form
# ---------------------------------------------------------------------------

@app.route('/api/fill-pdf', methods=['POST'])
def fill_pdf():
    try:
        file = request.files.get('file')
        updates_json = request.form.get('updates', '[]')
        updates = json.loads(updates_json)

        if not file:
            return jsonify({'error': 'No file provided'}), 400

        file_bytes = file.read()
        print(f"[fill-pdf] Received: {file.filename} ({len(file_bytes)} bytes), {len(updates)} updates")

        reader = PdfReader(io.BytesIO(file_bytes))
        writer = PdfWriter()
        writer.append_pages_from_reader(reader)

        filled = 0
        for u in updates:
            field_name = u.get('fieldName', '')
            value = u.get('value', '')
            if field_name and value:
                try:
                    writer.update_page_form_field_values(None, {field_name: value})
                    filled += 1
                    print(f"  [OK] {field_name} = '{value[:50]}'")
                except Exception as e:
                    print(f"  [SKIP] {field_name}: {e}")

        output = io.BytesIO()
        writer.write(output)
        output.seek(0)

        base, _ = os.path.splitext(file.filename or 'document.pdf')
        out_name = f"{base}_complete.pdf"

        print(f"[fill-pdf] {filled} fields written -> {out_name}")
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

        title_style = ParagraphStyle('LioraTitle', parent=styles['Heading1'],
                                     textColor=orange, fontSize=16, spaceAfter=6)
        sub_style = ParagraphStyle('LioraSub', parent=styles['Normal'],
                                   textColor=HexColor('#666666'), fontSize=9)
        note_style = ParagraphStyle('LioraNote', parent=styles['Normal'],
                                    textColor=HexColor('#cc0000'), fontSize=8,
                                    fontName='Helvetica-Bold', spaceAfter=12)
        label_style = ParagraphStyle('LioraLabel', parent=styles['Normal'],
                                     fontSize=9, textColor=HexColor('#333333'),
                                     fontName='Helvetica-Bold')
        value_style = ParagraphStyle('LioraValue', parent=styles['Normal'],
                                     fontSize=9, textColor=dark)

        elements = []
        elements.append(Paragraph(f"FICHE FOURNISSEUR - {company}", title_style))
        elements.append(Paragraph(
            f"Fichier : {file_name} - {time.strftime('%d/%m/%Y')}", sub_style))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(
            "Donnees a reporter dans le document original", note_style))

        table_data = []
        for f in fields:
            label = f.get('label', '')
            value = f.get('value', '')
            table_data.append([
                Paragraph(label, label_style),
                Paragraph(value or '-', value_style)
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

        elements.append(Spacer(1, 20))
        if contact:
            sig = f"<b>{contact.get('name', '')}</b> - {contact.get('role', '')}<br/>"
            sig += f"{contact.get('email', '')} | {contact.get('phone', '')}"
            elements.append(Paragraph(sig, sub_style))

        doc.build(elements)
        output.seek(0)

        base, _ = os.path.splitext(file_name)
        print(f"[recap-pdf] Generated for {file_name}, {len(fields)} fields")
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
    for port in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('127.0.0.1', port))
                return port
            except OSError:
                continue
    return start


def open_browser(port):
    time.sleep(1.2)
    webbrowser.open(f'http://127.0.0.1:{port}')


def main():
    port = find_free_port(5000)
    print(f"""
    =============================================
      Liora - Fiche Fournisseur Auto-Complete
      Serveur demarre sur le port {port}
      http://127.0.0.1:{port}
    =============================================
    """)
    threading.Thread(target=open_browser, args=(port,), daemon=True).start()
    app.run(host='127.0.0.1', port=port, debug=False, use_reloader=False)


if __name__ == '__main__':
    main()
