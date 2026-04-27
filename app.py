from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.styles import PatternFill, Font
import io
import zipfile
import os
import re
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app, expose_headers=['X-Found-Optic','X-Found-Sun','X-Not-Found','X-Red-Refs'])

RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
RED_FONT = Font(color="CC0000", bold=True)
HEADER_FONT = Font(bold=True)

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

def optic_candidates(style, month):
    return [f"CL{style} OPT {str(month).zfill(2)}"]

def sun_candidates(style, month):
    c = str(month).zfill(2)
    return [
        f"CL{style} SG OPT {c}",
        f"CL{style} SG {c}",
        f"CL{style} SG Z OPT {c}",
    ]

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/process', methods=['POST'])
def process():
    try:
        f1 = request.files.get('commande')
        f2 = request.files.get('optique')
        f3 = request.files.get('solaire')

        if not f1 or not f2 or not f3:
            return jsonify({"error": "3 fichiers requis"}), 400

        wb1 = openpyxl.load_workbook(f1, keep_vba=False)
        wb2 = openpyxl.load_workbook(f2, keep_vba=False)
        wb3 = openpyxl.load_workbook(f3, keep_vba=False)

        ws1 = wb1.active
        sn2 = 'NuORDER Order Data' if 'NuORDER Order Data' in wb2.sheetnames else wb2.sheetnames[0]
        sn3 = 'NuORDER Order Data' if 'NuORDER Order Data' in wb3.sheetnames else wb3.sheetnames[0]
        ws2 = wb2[sn2]
        ws3 = wb3[sn3]

        lookup2 = build_lookup(ws2)
        lookup3 = build_lookup(ws3)

        data_start = 11
        for i, row in enumerate(ws1.iter_rows(min_row=1, max_row=20), start=1):
            if row[0].value and 'REFERENCE' in str(row[0].value).upper():
                data_start = i + 1
                break

        found_optic = found_sun = not_found = 0
        red_refs = []

        for row in ws1.iter_rows(min_row=data_start):
            cell_a = row[0]
            cell_b = row[1]
            if cell_a.value is None:
                continue

            style = month = None

            if hasattr(cell_a.value, 'year'):
                style = str(cell_a.value.year)
                month = cell_a.value.month
            elif isinstance(cell_a.value, (int, float)) and cell_a.value > 1000:
                y, m = serial_to_year_month(cell_a.value)
                style = str(y)
                month = m
            elif isinstance(cell_a.value, str):
                m = re.match(r'^(\d{3,5})[^0-9](\d{1,2})$', cell_a.value.strip())
                if m:
                    style = m.group(1)
                    month = int(m.group(2))

            if not style:
                continue

            qty = cell_b.value if cell_b.value not in (None, '', 0) else 1
            found = False

            for candidate in optic_candidates(style, month):
                if candidate.upper() in lookup2:
                    lookup2[candidate.upper()].value = qty
                    found_optic += 1
                    found = True
                    break

            if not found:
                for candidate in sun_candidates(style, month):
                    if candidate.upper() in lookup3:
                        lookup3[candidate.upper()].value = qty
                        found_sun += 1
                        found = True
                        break

            if not found:
                ref_str = f"{style}-{month}"
                red_refs.append(ref_str)
                not_found += 1
                for cell in row:
                    cell.fill = RED_FILL
                    cell.font = Font(
                        color="CC0000",
                        bold=cell.font.bold if cell.font else False,
                        size=cell.font.size if cell.font else None,
                        name=cell.font.name if cell.font else None
                    )
                row[2].value = '⚠ INTROUVABLE'
                row[2].font = RED_FONT

        if not_found > 0:
            if '⚠ Refs introuvables' in wb1.sheetnames:
                del wb1['⚠ Refs introuvables']
            ws_err = wb1.create_sheet('⚠ Refs introuvables')
            ws_err.column_dimensions['A'].width = 18
            ws_err.column_dimensions['B'].width = 40
            ws_err['A1'] = 'Référence'
            ws_err['B1'] = 'Statut'
            ws_err['A1'].font = HEADER_FONT
            ws_err['B1'].font = HEADER_FONT
            for i, ref in enumerate(red_refs, start=2):
                ws_err[f'A{i}'] = ref
                ws_err[f'A{i}'].font = Font(color="CC0000", bold=True)
                ws_err[f'B{i}'] = '⚠ Introuvable dans les deux catalogues'
                ws_err[f'B{i}'].font = Font(color="CC0000")

        buf1 = io.BytesIO()
        buf2 = io.BytesIO()
        buf3 = io.BytesIO()
        wb1.save(buf1); buf1.seek(0)
        wb2.save(buf2); buf2.seek(0)
        wb3.save(buf3); buf3.seek(0)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            name1 = f1.filename.replace('.xlsx','') + '_traite.xlsx'
            name2 = f2.filename.replace('.xlsx','') + '_MAJ.xlsx'
            name3 = f3.filename.replace('.xlsx','') + '_MAJ.xlsx'
            zf.writestr(name1, buf1.read())
            zf.writestr(name2, buf2.read())
            zf.writestr(name3, buf3.read())
        zip_buf.seek(0)

        response = send_file(
            zip_buf,
            mimetype='application/zip',
            as_attachment=True,
            download_name='commande_traitee.zip'
        )
        response.headers['X-Found-Optic'] = str(found_optic)
        response.headers['X-Found-Sun'] = str(found_sun)
        response.headers['X-Not-Found'] = str(not_found)
        response.headers['X-Red-Refs'] = ','.join(red_refs)
        return response

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
