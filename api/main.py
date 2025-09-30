import requests
import json
import pandas as pd
from base64 import b64decode
from io import BytesIO, StringIO

URL_LS = "https://limesurvey.midap.cl/index.php?r=admin/remotecontrol"
USERNAME_LS = 'admin'
PASSWORD_LS = 'L1m3#5uRv3_'
survey_id = '717227'
format_export = 'xls'  # o 'csv'

def get_session_key(url, username, password):
    payload = json.dumps({
        "method": "get_session_key",
        "params": [username, password],
        "id": 1
    })
    headers = {'Content-Type': 'application/json'}
    response = requests.post(url, headers=headers, data=payload)
    print("Status code:", response.status_code)
    print("Response text:", response.text)
    response.raise_for_status()  # Lanza error si falla la request
    return json.loads(response.text)['result']

def export_responses(url, session_key, survey_id, format='xls'):
    payload = json.dumps({
        "method": "export_responses",
        "params": [
            session_key,
            survey_id,
            format,
            None,
            "all",
            "code"
        ],
        "id": 1
    })
    headers = {'Content-Type': 'application/json'}
    response = requests.post(url, headers=headers, data=payload)
    response.raise_for_status()

    result = json.loads(response.text)['result']
    
    if isinstance(result, dict):  # <-- significa que es un error, no un string base64
        raise ValueError(f"Error al exportar respuestas: {result}")
    
    return result

def transform_to_dataframe(data, format='xls'):
    decoded = b64decode(data)
    if format == 'csv':
        return pd.read_csv(StringIO(decoded.decode('utf-8')), sep=';')
    elif format == 'xls':
        file = BytesIO(decoded)
        return pd.read_excel(file)

def get_full_survey():
    session_key = get_session_key(URL_LS, USERNAME_LS, PASSWORD_LS)
    data = export_responses(URL_LS, session_key, survey_id, format_export)
    df = transform_to_dataframe(data, format_export)

    if "EMAIL" in df.columns and "PERMA23" in df.columns:
        df = df[df["EMAIL"].notna() & df["PERMA23"].notna()]
        df = df[df["EMAIL"].astype(str).str.strip() != ""]
        df = df[df["PERMA23"].astype(str).str.strip() != ""]
    else:
        print("Warning: 'EMAIL' or 'PERMA23' column not found in the data.")
    
    return df
