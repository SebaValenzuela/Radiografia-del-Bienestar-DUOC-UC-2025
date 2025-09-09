import pandas as pd
import os


def load_encuesta(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    if 'EMAIL' in df.columns:
        df['EMAIL'] = df['EMAIL'].str.replace(r'@.*', '', regex=True)
    return df

def load_estudiantes(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    if 'EMAIL' in df.columns:
        df['EMAIL'] = df['EMAIL'].str.replace(r'@.*', '', regex=True)
    return df

def marcar_respuestas(df_estudiantes: pd.DataFrame, df_encuesta: pd.DataFrame) -> pd.DataFrame:
    df_estudiantes = df_estudiantes.copy()
    df_estudiantes['RESPONDIO'] = df_estudiantes['EMAIL'].isin(df_encuesta['EMAIL']).astype(int)
    return df_estudiantes

def resumen_por_sede(df_estudiantes: pd.DataFrame) -> pd.DataFrame:
    tabla = df_estudiantes.groupby('SEDE').agg(
        Estudiantes=('EMAIL', 'count'),
        Respuestas=('RESPONDIO', 'sum')
    ).reset_index()
    tabla['%'] = (tabla['Respuestas'] / tabla['Estudiantes'] * 100).round(2)
    return tabla

def resumen_por_escuela(df_estudiantes: pd.DataFrame) -> pd.DataFrame:
    tabla = df_estudiantes.groupby('ESCUELA').agg(
        Estudiantes=('EMAIL', 'count'),
        Respuestas=('RESPONDIO', 'sum')
    ).reset_index()
    tabla['%'] = (tabla['Respuestas'] / tabla['Estudiantes'] * 100).round(2)
    return tabla

def guardar_resumen(df_resumen: pd.DataFrame, file_path: str):
    df_resumen.to_excel(file_path, index=False)
    print(f"Archivo guardado en '{file_path}'")


def procesar_encuesta(encuesta_file: str, estudiantes_file: str, output_dir: str):
    os.makedirs(output_dir, exist_ok=True)

    df_encuesta = load_encuesta(encuesta_file)
    df_estudiantes = load_estudiantes(estudiantes_file)

    df_estudiantes = marcar_respuestas(df_estudiantes, df_encuesta)

    resumen_sede = resumen_por_sede(df_estudiantes)
    resumen_escuela = resumen_por_escuela(df_estudiantes)

    guardar_resumen(resumen_sede, os.path.join(output_dir, "resumen_sedes.xlsx"))
    guardar_resumen(resumen_escuela, os.path.join(output_dir, "resumen_escuelas.xlsx"))

    return resumen_sede, resumen_escuela
