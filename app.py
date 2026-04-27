from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.styles import PatternFill
import io
import zipfile
import os
import re
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app, expose_headers=['X-Found-Optic','X-Found-Sun','X-Not-Found','X-Red-Refs'])

RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

def serial_to_year_month(serial):
    epoch = datetime(1899, 12, 30)
    d = epoch + timedelta(days=int(serial))
    return d.year, d.month

def build_lookup(ws):
    lookup = {}
    for row in ws.iter_rows(min_row=2):
        name_cell = row[4]
        if name_cell.value:
            key = str(name_cell.value).strip().upper()
            lookup[key] = row[13]
    return lookup

def get_candidates(style, month, is_optic=True):
    m_str = str(month).zfill(2)
    if is_optic:
        return [f"CL{style} OPT {m_str}"]
    else:
        return [f"CL{style} SG OPT {m_str}", f"CL{style} SG {m_str}", f"CL{style} SG Z OPT {m_str}"]

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/process', methods=['POST'])
def process():
    try:
        f1 = request.files.get('commande')
        f2 = request.files.get('optique')
        f3 = request.files.get('solaire')
        if not f1 or not f2 or not f3: return jsonify({"error": "3 fichiers requis"}), 400

        wb1 = openpyxl.load_workbook(f1)
        wb2 = openpyxl.load_workbook(f2, keep_vba=True)
        wb3 = openpyxl.load_workbook(f3, keep_vba=True)

        ws1 = wb1.active
        sn2 = 'NuORDER Order Data' if 'NuORDER Order Data' in wb2.sheetnames else wb2.sheetnames[0]
        sn3 = 'NuORDER Order Data' if 'NuORDER Order Data' in wb3.sheetnames else wb3.sheetnames[0]
        lookup2 = build_lookup(wb2[sn2])
        lookup3 = build_lookup(wb3[sn3])

        data_start = 11
        for i, row in enumerate(ws1.iter_rows(min_row=1, max_row=25), start=1):
            if row[0].value and 'REFERENCE' in str(row[0].value).upper():
                data_start = i + 1; break

        found_optic = found_sun = not_found = 0
        red_refs = []
        for row in ws1.iter_rows(min_row=data_start):
            cell_a, cell_b = row[0], row[1]
            if cell_a.value is None: continue
            style = month = None
            if hasattr(cell_a.value, 'year'):
                style, month = str(cell_a.value.year), cell_a.value.month
            elif isinstance(cell_a.value, (int, float)) and cell_a.value > 1000:
                style, month = serial_to_year_month(cell_a.value)
            elif isinstance(cell_a.value, str):
                m = re.match(r'^(\d{3,5})[^0-9](\d{1,2})$', cell_a.value.strip())
                if m: style, month = m.group(1), int(m.group(2))
            if not style: continue
            qty = cell_b.value if cell_b.value not in (None, '', 0) else 1
            found = False
            for c in get_candidates(style, month, True):
                if c.upper() in lookup2:
                    lookup2[c.upper()].value = qty
                    found_optic += 1; found = True; break
            if not found:
                for c in get_candidates(style, month, False):
                    if c.upper() in lookup3:
                        lookup3[c.upper()].value = qty
                        found_sun += 1; found = True; break
            if not found:
                not_found += 1; red_refs.append(f"{style}-{month}")
                for cell in row: cell.fill = RED_FILL
                row[2].value = '⚠ INTROUVABLE'

        buf1, buf2, buf3 = io.BytesIO(), io.BytesIO(), io.BytesIO()
        wb1.save(buf1); wb2.save(buf2); wb3.save(buf3)
        buf1.seek(0); buf2.seek(0); buf3.seek(0)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            zf.writestr(f1.filename.replace('.xlsx','') + '_traite.xlsx', buf1.read())
            zf.writestr(f2.filename.replace('.xlsx','') + '_MAJ.xlsx', buf2.read())
            zf.writestr(f3.filename.replace('.xlsx','') + '_MAJ.xlsx', buf3.read())
        zip_buf.seek(0)
        res = send_file(zip_buf, mimetype='application/zip', as_attachment=True, download_name='resultats.zip')
        res.headers['X-Found-Optic'], res.headers['X-Found-Sun'], res.headers['X-Not-Found'] = str(found_optic), str(found_sun), str(not_found)
        return res
    except Exception as e: return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
