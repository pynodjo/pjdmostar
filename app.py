from flask import Flask, request, jsonify, render_template, send_from_directory, abort
import json
import pandas as pd
import numpy as np
import logging
from datetime import datetime
import re
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)

def sanitize_for_json(value):
    """Convert NaN, None, and numpy types to JSON-safe values"""
    if pd.isna(value) or value is None:
        return 'N/A'
    if isinstance(value, (np.integer, np.int64, np.int32)):
        return int(value)
    if isinstance(value, (np.floating, np.float64, np.float32)):
        if np.isnan(value):
            return 'N/A'
        return float(value)
    if isinstance(value, np.bool_):
        return bool(value)
    return value

def sanitize_dict_for_json(data):
    """Recursively sanitize a dictionary for JSON serialization"""
    if isinstance(data, dict):
        return {k: sanitize_dict_for_json(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [sanitize_dict_for_json(item) for item in data]
    else:
        return sanitize_for_json(data)

# Setup logging
logging.basicConfig(level=logging.DEBUG)

# ============ DATA LOADING AND PREPROCESSING ============

# Load JSON files once at startup
with open('data.json', 'r', encoding='utf-8') as file:
    data = json.load(file)

with open('trafostanica_data.json', 'r', encoding='utf-8') as file:
    trafostanica_data = json.load(file)

with open('rastavljac_data.json', 'r', encoding='utf-8') as f:
    rastavljac_data = json.load(f)

# Create dictionaries for quick lookup
sifra_to_coordinates = {
    feature['properties']['SIFRA']: feature['geometry']['coordinates']
    for feature in data['features']
    if 'geometry' in feature and 'coordinates' in feature['geometry']
}

# Read and process the Excel file
df = pd.read_excel('EP_Eksport_Uredjaja.xlsx', sheet_name='Eksport_uredjaja', skiprows=6)
df = df.rename(columns=lambda x: x.strip())
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
df = df[~(df.duplicated(subset='Šifra') & df['Naziv TS'].isnull())]
df['Serijski'] = df['Serijski'].astype(int)
df['Šifra'] = df['Šifra'].astype(int)
df = df.drop_duplicates(subset=['Serijski', 'Šifra'])

# Create optimized mappings
serijski_broj_to_sifra = dict(zip(df['Serijski'], df['Šifra']))
sifra_to_row = {row['Šifra']: row for _, row in df.iterrows()}

# Customer search mapping
kupac_to_info = (
    df.groupby('Kupac')
    .apply(lambda x: x[['Serijski', 'Adresa', 'Šifra']].to_dict('records'))
    .reset_index(name='info')
    .set_index('Kupac')['info']
    .to_dict()
)

# Trafostanica mapping
trafostanica_to_info = {
    feature['properties']['NAZIV']: {
        "coordinates": feature['geometry']['coordinates'],
        "naziv": feature['properties']['NAZIV'],
        "snaga": feature['properties']['SNAGA'],
        "broj_transformatora": feature['properties']['BR_TRANSFORMATORA'],
        "konfiguracija_SN_postrojenja": feature['properties']['CONF_SN_POST'],
        "napojna_trafostanica": feature['properties']['NAPOJNA_TS'],
        "naziv_SN_odlaza": feature['properties']['ODLAZ_SN_NAZIV'],
        "tip_trafostanice": feature['properties']['TIP_TS'],
        "poslovnica": feature['properties']['POSLOVNICA'],
        "tip_kucista": feature['properties']['TS_GD'],
        "tip_izolacije": feature['properties']['TS_SN_POST'],
        "vlasnik": feature['properties']['VLASNIK'],
        "godina_izgradnje": feature['properties']['GODINA_IZGRADNJE'],
    }
    for feature in trafostanica_data['features']
    if 'geometry' in feature and 'coordinates' in feature['geometry']
}

# ============ UTILITY FUNCTIONS ============

def find_coordinates_by_sifra(sifra):
    return sifra_to_coordinates.get(sifra)

def find_sifra_by_serijski_broj(serijski_broj):
    return serijski_broj_to_sifra.get(serijski_broj)

def create_google_maps_url(coordinates):
    lon, lat = coordinates
    return f"https://www.google.com/maps?q={lat},{lon}"

def format_date(date):
    if pd.isna(date):
        return None
    return date.strftime('%d.%m.%Y')

def convert_to_native_types(value):
    """Convert numpy/pandas types to native Python types."""
    if isinstance(value, (np.integer, np.int64)):
        return int(value)
    elif isinstance(value, (np.floating, np.float64)):
        return float(value)
    elif pd.isna(value):
        return None
    return value

def get_additional_info(sifra):
    """Get additional info from DataFrame by sifra."""
    row = sifra_to_row.get(sifra)
    if row is None:
        return None
    
    info = {
        "Tip brojila": convert_to_native_types(row['Tip']),
        "Godina proizvodnje": format_date(row['Proizvodnj']),
        "Godina baždarenja": format_date(row['Baždarenje']),
        "Datum montaže": format_date(row['Datum žc']),
        "Serijski broj brojila": convert_to_native_types(row['Serijski']),
        "Kupac": convert_to_native_types(row['Kupac']),
        "Adresa": convert_to_native_types(row['Adresa']),
        "Šifra mjernog mjesta": convert_to_native_types(row['Šifra']),
        "Tarifna grupa": convert_to_native_types(row['T']),
        "Angažovana snaga": convert_to_native_types(row['A.sn']),
        "Naziv trafostanice": convert_to_native_types(row['Naziv TS'])
    }
    
    return info

# ============ ROUTES ============

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_coordinates_by_sifra', methods=['POST'])
def get_coordinates_by_sifra():
    sifra = request.form.get('sifra')
    app.logger.debug(f"Received SIFRA: {sifra}")
    
    if not sifra:
        return jsonify({"error": "Šifra nije pronađena, provjerite tačnost unesenih podataka."}), 404
    
    if not sifra.isdigit():
        return jsonify({"error": "Uneseni podatak nije tipa integer (mora sadržavati samo brojeve)."}), 400
    
    try:
        sifra = int(sifra)
        coordinates = find_coordinates_by_sifra(sifra)
        additional_info = get_additional_info(sifra)
        
        if not additional_info:
            return jsonify({"error": "Šifra mjernog mjesta nije pronađena u bazi podataka."}), 404
        
        result = {"additional_info": additional_info}
        if coordinates:
            result["url"] = create_google_maps_url(coordinates)
        else:
            result["message"] = "Lokacija nije dostupna."
        
        return jsonify(result)
    except ValueError:
        return jsonify({"error": "Invalid SIFRA format."}), 400

@app.route('/get_coordinates_by_serijski_broj', methods=['POST'])
def get_coordinates_by_serijski_broj():
    serijski_broj = request.form.get('serijski_broj')
    app.logger.debug(f"Received SERIJSKI_BROJ: {serijski_broj}")
    
    if not serijski_broj:
        return jsonify({"error": "Serijski broj nije pronađen, provjerite tačnost unesenih podataka."}), 404
    
    if not serijski_broj.isdigit():
        return jsonify({"error": "Uneseni podatak nije tipa integer (mora sadržavati samo brojeve)."}), 400
    
    serijski_broj = int(serijski_broj)
    sifra = find_sifra_by_serijski_broj(serijski_broj)
    
    if not sifra:
        return jsonify({"error": "Serijski broj nije pronađen u bazi podataka."}), 404
    
    coordinates = find_coordinates_by_sifra(sifra)
    additional_info = get_additional_info(sifra)
    
    if not additional_info:
        return jsonify({"error": "Šifra mjernog mjesta nije pronađena u bazi podataka."}), 404
    
    result = {"additional_info": additional_info}
    if coordinates:
        result["url"] = create_google_maps_url(coordinates)
    else:
        result["message"] = "Lokacija nije dostupna."
    
    return jsonify(result)

@app.route('/get_customer_suggestions', methods=['POST'])
def get_customer_suggestions():
    kupac_input = request.form.get('input', '').strip()
    
    if len(kupac_input) < 3:
        return jsonify([])

    pattern = re.compile(re.escape(kupac_input), re.IGNORECASE)
    suggestions = []

    for kupac, info_list in kupac_to_info.items():
        if pattern.search(kupac):
            for record in info_list:
                suggestions.append(f"{kupac} ({record['Serijski']}, {record['Adresa']})")
    
    return jsonify(suggestions[:10])




@app.route('/get_coordinates_by_kupac', methods=['POST'])
def get_coordinates_by_kupac():
    kupac_input = request.form.get('kupac')
    serijski_input = request.form.get('serijski')
    
    if not (kupac_input and serijski_input):
        return jsonify({"error": "Kupac ili serijski broj nije pronađen u bazi podataka."}), 404
    
    if kupac_input not in kupac_to_info:
        return jsonify({"error": "Kupac nije pronađen u bazi podataka."}), 404
    
    try:
        serijski = int(serijski_input)
        sifra = find_sifra_by_serijski_broj(serijski)
        
        if not sifra:
            return jsonify({"error": "Kupac ili serijski broj nije pronađen u bazi podataka."}), 404
        
        coordinates = find_coordinates_by_sifra(sifra)
        additional_info = get_additional_info(sifra)
        
        result = {"additional_info": additional_info}
        if coordinates:
            result["url"] = create_google_maps_url(coordinates)
        else:
            result["message"] = "Lokacija nije dostupna."
        
        return jsonify(result)
    except (ValueError, TypeError):
        return jsonify({"error": "Neispravan serijski broj."}), 400

@app.route('/get_oh_values_by_oj/<oj_value>', methods=['GET'])
def get_oh_values_by_oj(oj_value):
    app.logger.debug(f"get_oh_values_by_oj called with oj_value: {oj_value}")
    
    try:
        if oj_value == "303":
            filtered_df = df[df['OJ'].isin([3031, 3032])]
        else:
            filtered_df = df[df['OJ'] == int(oj_value)]
        
        unique_oh_values = sorted(filtered_df['OH'].unique().tolist())
        return jsonify(unique_oh_values)
    except Exception as e:
        app.logger.error(f"Error in get_oh_values_by_oj: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/search_by_oj_oh', methods=['GET'])
def search_by_oj_oh():
    oj_value = request.args.get('oj')
    oh_value = request.args.get('oh')
    
    if not oj_value or not oh_value:
        return jsonify({"error": "Missing oj or oh parameter"}), 400
    
    app.logger.debug(f"search_by_oj_oh: oj={oj_value}, oh={oh_value}")
    
    try:
        if oj_value == "303":
            filtered_df = df[(df['OJ'].isin([3031, 3032])) & (df['OH'] == oh_value)]
        else:
            filtered_df = df[(df['OJ'] == int(oj_value)) & (df['OH'] == oh_value)]
        
        features = []
        for _, row in filtered_df.iterrows():
            sifra = row['Šifra']
            coordinates = find_coordinates_by_sifra(sifra)
            
            if coordinates:
                # Sanitize EACH value as you create the dictionary
                features.append({
                    'geometry': {
                        'coordinates': coordinates,
                        'type': 'Point'
                    },
                    'properties': {
                        'IME_PREZIME': sanitize_for_json(row['Kupac']),
                        'ADRESA_MM': sanitize_for_json(row['Adresa']),
                        'SIFRA': sanitize_for_json(sifra),
                        'SERIJSKI': sanitize_for_json(row['Serijski']),
                        'TIP': sanitize_for_json(row['Tip']),
                        'ROH': sanitize_for_json(row['ROH']),
                        'ANG_SNAGA': sanitize_for_json(row['A.sn'])  # THIS was the problem
                    },
                    'type': 'Feature'
                })
        
        if not features:
            return jsonify({"error": "No features found for this combination"}), 404
        
        center = [features[0]['geometry']['coordinates'][1], features[0]['geometry']['coordinates'][0]]
        
        # Now you can just return directly, no need to sanitize again
        return jsonify({
            "features": features,
            "center": center,
            "total": len(features)
        })
    
    except Exception as e:
        app.logger.error(f"Error in search_by_oj_oh: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/get_trafostanica_data', methods=['POST'])
def get_trafostanica_data():
    trafostanica = request.json.get('trafostanica', '').strip()
    
    if trafostanica not in trafostanica_to_info:
        return jsonify({"error": "Trafostanica nije pronađena u bazi podataka."}), 404
    
    info = trafostanica_to_info[trafostanica]
    coordinates = info["coordinates"]
    
    return jsonify({
        **info,
        "google_maps_url": create_google_maps_url(coordinates)
    })

@app.route('/get_trafostanica_suggestions', methods=['POST'])
def get_trafostanica_suggestions():
    try:
        user_input = request.form.get('input', '').strip().lower()
        
        if len(user_input) < 3:
            return jsonify({'suggestions': []})
        
        pattern = re.compile(f'.*{re.escape(user_input)}.*', re.IGNORECASE)
        matches = [
            feature['properties']['NAZIV']
            for feature in trafostanica_data['features']
            if 'properties' in feature 
            and 'NAZIV' in feature['properties']
            and pattern.search(feature['properties']['NAZIV'])
        ]
        
        return jsonify({'suggestions': matches[:10]})
    except Exception as e:
        app.logger.error(f"Error in get_trafostanica_suggestions: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/view_all_trafostanice', methods=['GET'])
def view_all_trafostanice():
    try:
        features = []
        for feature in trafostanica_data['features']:
            if ('geometry' in feature and 
                'coordinates' in feature['geometry'] and 
                'properties' in feature):
                
                coords = feature['geometry']['coordinates']
                properties = feature['properties']
                
                # Sanitize as you build
                features.append({
                    'type': 'Feature',
                    'geometry': {
                        'type': 'Point',
                        'coordinates': coords
                    },
                    'properties': {
                        'naziv': sanitize_for_json(properties.get('NAZIV', 'N/A')),
                        'snaga': sanitize_for_json(properties.get('SNAGA', 'N/A'))
                    }
                })
        
        if features:
            lats = [f['geometry']['coordinates'][1] for f in features]
            lons = [f['geometry']['coordinates'][0] for f in features]
            center = [sum(lats)/len(lats), sum(lons)/len(lons)]
        else:
            center = [43.343, 17.807]
        
        return jsonify({
            "features": features,
            "center": center,
            "total": len(features)
        })
    
    except Exception as e:
        app.logger.error(f"Error generating trafostanica data: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/get_ts_naziv_values', methods=['GET'])
def get_ts_naziv_values():
    try:
        search_term = request.args.get('search', '').lower()
        
        # Get unique TS names from DataFrame, not JSON
        ts_naziv_values = df['Naziv TS'].dropna().unique().tolist()
        
        if search_term:
            ts_naziv_values = [
                ts for ts in ts_naziv_values 
                if search_term in ts.lower()
            ]
        
        return jsonify(sorted(ts_naziv_values))
    except Exception as e:
        app.logger.error(f"Error in get_ts_naziv_values: {str(e)}")
        return jsonify({"error": "An error occurred while processing TS_NAZIV values."}), 500

@app.route('/filter_data_by_ts_naziv', methods=['GET'])
def filter_data_by_ts_naziv():
    ts_naziv = request.args.get('ts_naziv')
    
    if not ts_naziv:
        return jsonify({"error": "Missing ts_naziv parameter"}), 400
    
    app.logger.debug(f"filter_data_by_ts_naziv called with ts_naziv: {ts_naziv}")
    
    try:
        # Filter DataFrame by exact match of Naziv TS
        filtered_df = df[df['Naziv TS'] == ts_naziv]
        app.logger.debug(f"Filtered DataFrame shape: {filtered_df.shape}")
        
        if filtered_df.empty:
            # Try case-insensitive match as fallback
            filtered_df = df[df['Naziv TS'].str.lower() == ts_naziv.lower()]
            app.logger.debug(f"Case-insensitive match shape: {filtered_df.shape}")
        
        if filtered_df.empty:
            return jsonify({"error": "No features found for this TS_NAZIV"}), 404
        
        features = []
        for _, row in filtered_df.iterrows():
            sifra = row['Šifra']
            coordinates = find_coordinates_by_sifra(sifra)
            
            if coordinates:
                # Sanitize EACH value as you create the dictionary
                features.append({
                    'geometry': {
                        'coordinates': coordinates,
                        'type': 'Point'
                    },
                    'properties': {
                        'IME_PREZIME': sanitize_for_json(row['Kupac']),
                        'ADRESA_MM': sanitize_for_json(row['Adresa']),
                        'SIFRA': sanitize_for_json(sifra),
                        'SERIJSKI': sanitize_for_json(row['Serijski']),
                        'TIP': sanitize_for_json(row['Tip']),
                        'ANG_SNAGA': sanitize_for_json(row['A.sn'])  # THIS was the problem
                    },
                    'type': 'Feature'
                })
        
        if not features:
            return jsonify({"error": "No coordinates found for meters in this TS"}), 404
        
        center = [
            features[0]['geometry']['coordinates'][1],
            features[0]['geometry']['coordinates'][0]
        ]
        
        return jsonify({
            "features": features,
            "center": center,
            "total": len(features)
        })
        
    except Exception as e:
        app.logger.error(f"Error in filter_data_by_ts_naziv: {str(e)}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route('/get_rastavljac_suggestions', methods=['POST'])
def get_rastavljac_suggestions():
    try:
        input_value = request.form.get('input', '').strip().lower()
        
        if len(input_value) < 3:
            return jsonify({"suggestions": []})
        
        suggestions = set()
        for feature in rastavljac_data.get('features', []):
            props = feature.get('properties', {})
            candidates = [
                props.get('NTS_NAZIV'),
                props.get('SIFRA'),
                props.get('SNO_NAZIV'),
                props.get('DSN_NAZIV')
            ]
            
            for c in candidates:
                if c and input_value in str(c).lower():
                    suggestions.add(str(c))
        
        return jsonify({"suggestions": sorted(list(suggestions))[:10]})
    except Exception as e:
        app.logger.error(f"Error in get_rastavljac_suggestions: {e}")
        return jsonify({"error": "Error fetching suggestions"}), 500

@app.route('/get_rastavljac_data', methods=['POST'])
def get_rastavljac_data():
    try:
        body = request.get_json(silent=True) or {}
        q = body.get('rastavljac', '').strip()
        
        if not q:
            return jsonify({"error": "Invalid rastavljač value"}), 400
        
        q_lower = q.lower()
        
        # Exact matches
        features = [
            f for f in rastavljac_data.get('features', [])
            if (str(f.get('properties', {}).get('SIFRA', '')) == q)
            or (str(f.get('properties', {}).get('NTS_NAZIV', '')) == q)
        ]
        
        # Fallback: substring match
        if not features:
            features = []
            for f in rastavljac_data.get('features', []):
                props = f.get('properties', {})
                joined = " ".join([
                    str(props.get('SIFRA', '')),
                    str(props.get('NTS_NAZIV', '')),
                    str(props.get('SNO_NAZIV', '')),
                    str(props.get('DSN', '')),
                    str(props.get('DSN_NAZIV', ''))
                ]).lower()
                
                if q_lower in joined:
                    features.append(f)
        
        if not features:
            return jsonify({"error": "No features found for this rastavljač"}), 404
        
        # Get first valid coordinates
        first_coords = None
        for f in features:
            geom = f.get('geometry', {})
            coords = geom.get('coordinates')
            if coords and len(coords) >= 2:
                first_coords = coords
                break
        
        if first_coords:
            lon, lat = first_coords[0], first_coords[1]
            center = [lat, lon]
            google_maps_url = f"https://www.google.com/maps/search/?api=1&query={lat},{lon}"
        else:
            center = None
            google_maps_url = None
        
        first_props = features[0].get('properties', {})
        
        return jsonify({
            "features": features,
            "center": center,
            "google_maps_url": google_maps_url,
            "naziv": first_props.get('NTS_NAZIV') or first_props.get('SNO_NAZIV') or '',
            "sifra": first_props.get('SIFRA'),
            "pj": first_props.get('PJ'),
            "vrsta_upravljanja": first_props.get('VRSTA_UPRAVLJANJA'),
            "sno_naziv": first_props.get('SNO_NAZIV'),
            "dsn": first_props.get('DSN'),
            "dsn_naziv": first_props.get('DSN_NAZIV'),
        })
    except Exception as e:
        app.logger.error(f"Error in get_rastavljac_data: {e}")
        return jsonify({"error": "An error occurred while filtering rastavljač data."}), 500

# PDF Downloads
PDF_FOLDER = os.path.join(app.static_folder, "pdfs")
ALLOWED_PDFS = {
    "Jednopolna_shema_PJD_Mostar.pdf": "Jednopolna shema PJD Mostar",
    "Jednopolna_shema_PJD_Jablanica.pdf": "Jednopolna shema PJD Jablanica",
    "Jednopolna_shema_PJD_Konjic.pdf": "Jednopolna shema PJD Konjic",
}

@app.route("/download_pdf/<path:filename>")
def download_pdf(filename):
    filename_safe = secure_filename(filename)
    if filename_safe not in ALLOWED_PDFS:
        app.logger.warning(f"Attempt to access unauthorized file: {filename_safe}")
        abort(404)
    return send_from_directory(directory=PDF_FOLDER, filename=filename_safe, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)