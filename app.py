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
# On expose les headers pour le front-end
CORS(app, expose_headers=['X-Found-Optic','X-Found-Sun','X-Not-Found','X-Red-Refs'])

RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
RED_FONT = Font(color="CC0000", bold=True)
HEADER_FONT = Font(bold=True)

def serial_to_year_month(serial):
    epoch = datetime(1899, 12, 30)
    d = epoch + timedelta(days=int(serial))
    return d.year, d.month

def build_lookup(ws):
    """
    Crée un dictionnaire qui pointe vers l'OBJET cellule de la colonne Qty (N).
    On utilise le nom du produit (colonne E) comme clé.
    """
    lookup = {}
    for row in ws.iter_rows(min_row=2):
        name_cell = row[4] # Colonne E (Index 4)
        if name_cell.value:
            key = str(name_cell.value).strip().upper()
            # On cible la cellule de la colonne N (Index 13) pour la quantité
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

        # --- CONFIGURATION CRUCIALE POUR LES IMAGES NUORDER ---
        # keep_vba=True est souvent nécessaire pour les fichiers NuORDER car ils 
        # contiennent des structures de données complexes que le mode standard ignore.
        # On ne met surtout pas data_only=True car cela détruit les images.
        wb1 = openpyxl.load_workbook(f1)
        wb2 = openpyxl.load_workbook(f2, keep_vba=True)
        wb3 = openpyxl.load_workbook(f3, keep_vba=True)

        ws1 = wb1.active
        # Identification des feuilles de données NuORDER
        sn2 = 'NuORDER Order Data' if 'NuORDER Order Data' in wb2.sheetnames else wb2.sheetnames[0]
        sn3 = 'NuORDER Order Data' if 'NuORDER Order Data' in wb3.sheetnames else wb3.sheetnames[0]
        ws2 = wb2[sn2]
        ws3 = wb3[sn3]

        lookup2 = build_lookup(ws2)
        lookup3 = build_lookup(ws3)

        # Détection de la ligne de début des données sur le bon de commande
        data_start = 11
        for i, row in enumerate(ws1.iter_rows(min_row=1, max_row=25), start=1):
            if row[0].value and 'REFERENCE' in str(row[0].value).upper():
                data_start = i + 1
                break

        found_optic = found_sun = not_found = 0
        red_refs = []

        for row in ws1.iter_rows(min_row=data_start):
            cell_a = row[0] # Référence
            cell_b = row[1] # Quantité
            if cell_a.value is None:
                continue

            style = month = None
            # Conversion de la date ou du texte en Style/Mois
            if hasattr(cell_a.value, 'year'):
                style, month = str(cell_a.value.year), cell_a.value.month
            elif isinstance(cell_a.value, (int, float)) and cell_a.value > 1000:
                style, month = serial_to_year_month(cell_a.value)
            elif isinstance(cell_a.value, str):
                m = re.match(r'^(\d{3,5})[^0-9](\d{1,2})$', cell_a.value.strip())
                if m: style, month = m.group(1), int(m.group(2))

            if not style:
                continue

            qty = cell_b.value if cell_b.value not in (None, '', 0) else 1
            found = False

            # Recherche dans Optique
            for cand in optic_candidates(style, month):
                if cand.upper() in lookup2:
                    lookup2[cand.upper()].value = qty
                    found_optic += 1
                    found = True
                    break

            # Recherche dans Solaire
            if not found:
                for cand in sun_candidates(style, month):
                    if cand.upper() in lookup3:
                        lookup3[cand.upper()].value = qty
                        found_sun += 1
                        found = True
                        break

            # Marquage si non trouvé
            if not found:
                not_found += 1
                red_refs.append(f"{style}-{month}")
                for cell in row:
                    cell.fill = RED_FILL
                row[2].value = '⚠ INTROUVABLE'

        # Sauvegarde dans les buffers
        buf1 = io.BytesIO(); wb1.save(buf1); buf1.seek(0)
        buf2 = io.BytesIO(); wb2.save(buf2); buf2.seek(0)
        buf3 = io.BytesIO(); wb3.save(buf3); buf3.seek(0)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, 'w') as zf:
            name1 = f1.filename.replace('.xlsx','') + '_traite.xlsx'
            name2 = f2.filename.replace('.xlsx','') + '_MAJ.xlsx'
            name3 = f3.filename.replace('.xlsx','') + '_MAJ.xlsx'
            zf.writestr(name1, buf1.read())
            zf.writestr(name2, buf2.read())
            zf.writestr(name3, buf3.read())
        zip_buf.seek(0)

        response = send_file(zip_buf, mimetype='application/zip', as_attachment=True, download_name='resultats.zip')
        response.headers['X-Found-Optic'] = str(found_optic)
        response.headers['X-Found-Sun'] = str(found_sun)
        response.headers['X-Not-Found'] = str(not_found)
        response.headers['X-Red-Refs'] = ','.join(red_refs)
        return response

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
