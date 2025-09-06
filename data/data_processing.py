import pandas as pd
import os


def load_encuesta(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    if 'EMAIL' in df.columns:
        df['EMAIL'] = df['EMAIL'].str.replace(r'@.*', '', regex=True)
    return df

def load_alumnos(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    if 'EMAIL' in df.columns:
        df['EMAIL'] = df['EMAIL'].str.replace(r'@.*', '', regex=True)
    return df

def marcar_respuestas(df_alumnos: pd.DataFrame, df_encuesta: pd.DataFrame) -> pd.DataFrame:
    df_alumnos = df_alumnos.copy()
    df_alumnos['RESPONDIO'] = df_alumnos['EMAIL'].isin(df_encuesta['EMAIL']).astype(int)
    return df_alumnos

def resumen_por_sede(df_alumnos: pd.DataFrame) -> pd.DataFrame:
    tabla = df_alumnos.groupby('SEDE').agg(
        Alumnos=('EMAIL', 'count'),
        Respuestas=('RESPONDIO', 'sum')
    ).reset_index()
    tabla['%'] = (tabla['Respuestas'] / tabla['Alumnos'] * 100).round(2)
    return tabla

def resumen_por_escuela(df_alumnos: pd.DataFrame) -> pd.DataFrame:
    tabla = df_alumnos.groupby('ESCUELA').agg(
        Alumnos=('EMAIL', 'count'),
        Respuestas=('RESPONDIO', 'sum')
    ).reset_index()
    tabla['%'] = (tabla['Respuestas'] / tabla['Alumnos'] * 100).round(2)
    return tabla

def guardar_resumen(df_resumen: pd.DataFrame, file_path: str):
    df_resumen.to_excel(file_path, index=False)
    print(f"Archivo guardado en '{file_path}'")


def procesar_encuesta(encuesta_file: str, alumnos_file: str, output_dir: str):
    os.makedirs(output_dir, exist_ok=True)

    df_encuesta = load_encuesta(encuesta_file)
    df_alumnos = load_alumnos(alumnos_file)

    df_alumnos = marcar_respuestas(df_alumnos, df_encuesta)

    resumen_sede = resumen_por_sede(df_alumnos)
    resumen_escuela = resumen_por_escuela(df_alumnos)

    guardar_resumen(resumen_sede, os.path.join(output_dir, "resumen_sedes.xlsx"))
    guardar_resumen(resumen_escuela, os.path.join(output_dir, "resumen_escuelas.xlsx"))

    return resumen_sede, resumen_escuela
