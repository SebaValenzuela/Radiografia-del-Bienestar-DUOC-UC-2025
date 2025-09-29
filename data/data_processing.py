import pandas as pd
import os


def load_encuesta(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    if 'EMAIL' in df.columns:
        df['EMAIL'] = df['EMAIL'].str.replace(r'@.*', '', regex=True)
    return df

def load_estudiantes(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name="Universo Final")
    if 'EMAIL' in df.columns:
        df['EMAIL'] = df['EMAIL'].str.replace(r'@.*', '', regex=True)
    return df

def generar_sedes_matriculas(input_file: str, output_file: str):
    # Lista de sedes que nos interesa
    sedes = [
        "Alameda",
        "Antonio Varas",
        "Campus Arauco",
        "Campus Nacimiento",
        "Campus Villarrica",
        "Concepción",
        "Maipú",
        "Melipilla",
        "Padre Alonso de Ovalle",
        "Plaza Norte",
        "Plaza Oeste",
        "Plaza Vespucio",
        "Puente Alto",
        "Puerto Montt",
        "San Bernardo",
        "San Carlos de Apoquindo",
        "San Joaquín",
        "Valparaíso",
        "Viña del Mar",
        "Online"
    ]

    # --- Leer archivo ---
    df = pd.read_excel(input_file)
    df = df.rename(columns={
        "Etiquetas de fila": "SEDE",
        "Suma de Matrícula": "Cantidad de estudiantes"
    })

    # --- Agrupar por sede sumando las escuelas ---
    sedes_totales = []
    for sede in sedes:
        mask = df['SEDE'] == sede
        total = df.loc[mask, 'Cantidad de estudiantes'].sum()
        sedes_totales.append({'SEDE': sede, 'Cantidad de estudiantes': total})

    df_sedes = pd.DataFrame(sedes_totales)

    # --- Guardar ---
    df_sedes.to_excel(output_file, index=False)
    print(f"Archivo '{output_file}' generado con éxito")
    return df_sedes

def marcar_respuestas(df_estudiantes: pd.DataFrame, df_encuesta: pd.DataFrame) -> pd.DataFrame:
    df_estudiantes = df_estudiantes.copy()
    df_estudiantes['RESPONDIO'] = df_estudiantes['EMAIL'].isin(df_encuesta['EMAIL']).astype(int)
    return df_estudiantes

def resumen_por_sede(cantidad_estudiantes_file: str, df_estudiantes: pd.DataFrame = None, n_referencial_file: str = None) -> pd.DataFrame:

    sedes = [
        "Alameda",
        "Antonio Varas",
        "Campus Arauco",
        "Campus Nacimiento",
        "Campus Villarrica",
        "Concepción",
        "Maipú",
        "Melipilla",
        "Padre Alonso de Ovalle",
        "Plaza Norte",
        "Plaza Oeste",
        "Plaza Vespucio",
        "Puente Alto",
        "Puerto Montt",
        "San Bernardo",
        "San Carlos de Apoquindo",
        "San Joaquín",
        "Valparaíso",
        "Viña del Mar",
        "Online"
    ]

    # --- Leer archivo de matrículas ---
    df = pd.read_excel(cantidad_estudiantes_file)
    df = df.rename(columns={
        "Etiquetas de fila": "SEDE",
        "Suma de Matrícula": "Cantidad de estudiantes"
    })

    # --- Agrupar por sede sumando escuelas ---
    sedes_totales = []
    for sede in sedes:
        mask = df['SEDE'] == sede
        total = df.loc[mask, 'Cantidad de estudiantes'].sum()
        sedes_totales.append({'SEDE': sede, 'Cantidad de estudiantes': int(total)})

    df_sedes = pd.DataFrame(sedes_totales)

    if 'Online' in df_sedes['SEDE'].values:
        df_sedes.loc[df_sedes['SEDE'] == 'Online', 'Cantidad de estudiantes'] = 1893

    # --- Si se entrega df_estudiantes, calcular respuestas y % de avance ---
    if df_estudiantes is not None:
        # Marcar respuestas si aún no está hecho
        if 'RESPONDIO' not in df_estudiantes.columns:
            df_estudiantes = marcar_respuestas(df_estudiantes, df_estudiantes)  # aquí df_estudiantes debe tener emails correctos
        
        df_sedes['SEDE'] = df_sedes['SEDE'].str.upper()
        df_estudiantes['SEDE'] = df_estudiantes['SEDE'].str.upper()
        print(set(df_sedes['SEDE']) - set(df_estudiantes['SEDE']))
        # Columna de respuestas por sede
        df_sedes['Cantidad de respuestas'] = df_sedes['SEDE'].map(
            df_estudiantes.groupby('SEDE')['RESPONDIO'].sum()
        )

        # Columna de % de avance
        df_sedes['% de avance respecto a total'] = (
            (df_sedes['Cantidad de respuestas'] / df_sedes['Cantidad de estudiantes'] * 100)
            .round(2)
            .astype(str) + '%'
        )

    if n_referencial_file:
        df_n = pd.read_excel(n_referencial_file)
        # Sumar N referencial por sede
        df_n['SEDE'] = df_n['SEDE'].str.upper().fillna(method='ffill')  # llenar nombres repetidos por sede
        n_por_sede = df_n.groupby('SEDE')['N referencial'].sum()
        df_sedes['N referencial'] = (
            df_sedes['SEDE'].map(n_por_sede)
            .fillna(0)
            .round()
            .astype(int)
        )

        df_sedes['Encuestas restantes para N referencial'] = (
            df_sedes['N referencial'] - df_sedes['Cantidad de respuestas']
        ).astype(int)

    return df_sedes



def resumen_por_escuela(df_estudiantes: pd.DataFrame, df_cantidad_estudiantes: pd.DataFrame) -> pd.DataFrame:
    # Lista de escuelas válidas (tal como quieres que aparezcan)
    ESCUELAS_VALIDAS = [
        "Administración y Negocios",
        "Comunicación",
        "Construcción",
        "Diseño",
        "Informática y Telecomunicaciones",
        "Ingeniería, Medio Ambiente y Recursos Naturales",
        "Salud y Bienestar",
        "Turismo y Hospitalidad",
        "Gastronomía"
    ]


    # Filtrar solo las filas con escuelas válidas
    df_cantidad_estudiantes = df_cantidad_estudiantes[
        df_cantidad_estudiantes['ESCUELA'].isin(ESCUELAS_VALIDAS)
    ].copy()

    df_cantidad_estudiantes['Cantidad de estudiantes'] = df_cantidad_estudiantes['Cantidad de estudiantes'].fillna(0).astype(int)

    # Agrupar por escuela
    resumen = (
        df_cantidad_estudiantes
        .groupby('ESCUELA')['Cantidad de estudiantes']
        .sum()
        .reset_index()
    )

    # Agregar cantidad de respuestas
    respuestas_por_escuela = df_estudiantes.groupby('ESCUELA')['RESPONDIO'].sum()
    resumen['Cantidad de respuestas'] = resumen['ESCUELA'].map(respuestas_por_escuela).fillna(0).astype(int)

    # Porcentaje de avance
    resumen['% de avance respecto a total'] = (
        (resumen['Cantidad de respuestas'] / resumen['Cantidad de estudiantes'] * 100)
        .round(2)
        .astype(str) + '%'
    )

    return resumen[['ESCUELA', 'Cantidad de estudiantes', 'Cantidad de respuestas', '% de avance respecto a total']]

def resumen_escuela_por_sede(df_cantidad_estudiantes: pd.DataFrame, df_estudiantes: pd.DataFrame, n_referencial_file: str = None) -> pd.DataFrame:
    """
    Genera un resumen combinando SEDE y ESCUELA con la cantidad de estudiantes,
    cantidad de respuestas, porcentaje de avance, N referencial (opcional) y
    encuestas restantes para N referencial.
    Detecta sedes y escuelas en la misma columna y agrupa las escuelas por sede.

    Devuelve un DataFrame con columnas:
        - SEDE
        - ESCUELA
        - Cantidad de estudiantes
        - Cantidad de respuestas
        - % de avance respecto a total
        - N referencial (si se entrega n_referencial_file)
        - Encuestas restantes para N referencial (si N referencial existe)
    """

    sedes = [
        "Alameda",
        "Antonio Varas",
        "Campus Arauco",
        "Campus Nacimiento",
        "Campus Villarrica",
        "Concepción",
        "Maipú",
        "Melipilla",
        "Padre Alonso de Ovalle",
        "Plaza Norte",
        "Plaza Oeste",
        "Plaza Vespucio",
        "Puente Alto",
        "Puerto Montt",
        "San Bernardo",
        "San Carlos de Apoquindo",
        "San Joaquín",
        "Valparaíso",
        "Viña del Mar",
        "Online"
    ]

    ESCUELAS_VALIDAS = [
        "Administración y Negocios",
        "Comunicación",
        "Construcción",
        "Diseño",
        "Informática y Telecomunicaciones",
        "Ingeniería, Medio Ambiente y Recursos Naturales",
        "Salud y Bienestar",
        "Turismo y Hospitalidad",
        "Gastronomía"
    ]

    sedes_asignadas = []
    escuelas_asignadas = []
    estudiantes_asignados = []
    respuestas_asignadas = []

    current_sede = None

    for idx, row in df_cantidad_estudiantes.iterrows():
        nombre = str(row['ESCUELA']).strip()

        if nombre in sedes:
            current_sede = nombre
        elif nombre in ESCUELAS_VALIDAS and current_sede:
            sedes_asignadas.append(current_sede)
            escuelas_asignadas.append(nombre)

            # Cantidad de estudiantes desde df_cantidad_estudiantes
            cantidad_estudiantes = row.get('Cantidad de estudiantes', 0)
            estudiantes_asignados.append(cantidad_estudiantes)

            # Calcular cantidad de respuestas desde df_estudiantes
            mask = (
                (df_estudiantes['SEDE'].str.upper() == current_sede.upper()) &
                (df_estudiantes['ESCUELA'] == nombre)
            )
            cantidad_respuestas = df_estudiantes.loc[mask, 'RESPONDIO'].sum()
            respuestas_asignadas.append(cantidad_respuestas)

    # Crear DataFrame final
    df_resumen = pd.DataFrame({
        'SEDE': sedes_asignadas,
        'ESCUELA': escuelas_asignadas,
        'Cantidad de estudiantes': estudiantes_asignados,
        'Cantidad de respuestas': respuestas_asignadas
    })

    # Agregar porcentaje de avance
    df_resumen['% de avance respecto a total'] = (
        (df_resumen['Cantidad de respuestas'] / df_resumen['Cantidad de estudiantes'] * 100)
        .round(2)
        .astype(str) + '%'
    )

    # Agregar N referencial si se entrega archivo
    if n_referencial_file:
        df_n = pd.read_excel(n_referencial_file)
        df_n['SEDE'] = df_n['SEDE'].str.upper().fillna(method='ffill')
        df_n['ESCUELA'] = df_n['ESCUELA'].fillna(method='ffill')
        n_map = df_n.set_index(['SEDE', 'ESCUELA'])['N referencial'].to_dict()

        df_resumen['N referencial'] = df_resumen.apply(
            lambda x: int(round(n_map.get((x['SEDE'].upper(), x['ESCUELA']), 0))),
            axis=1
        )

        # Agregar columna de encuestas restantes
        df_resumen['Encuestas restantes para N referencial'] = (
            df_resumen['N referencial'] - df_resumen['Cantidad de respuestas']
        ).astype(int)

    return df_resumen





def guardar_resumen(df_resumen: pd.DataFrame, file_path: str):
    df_resumen.to_excel(file_path, index=False)
    print(f"Archivo guardado en '{file_path}'")


def procesar_encuesta(encuesta_file: str, estudiantes_file: str, output_dir: str, cantidad_estudiantes_file: str, n_referencial_file: str):
    os.makedirs(output_dir, exist_ok=True)

    df_encuesta = load_encuesta(encuesta_file)
    df_estudiantes = load_estudiantes(estudiantes_file)

    # --- Resumen de sedes usando tabla dinámica ---

    # --- Marcar respuestas ---
    df_estudiantes = marcar_respuestas(df_estudiantes, df_encuesta)

    total_respuestas_sede = df_estudiantes.groupby("SEDE")["RESPONDIO"].sum().sum()
    total_respuestas_escuela = df_estudiantes[df_estudiantes["ESCUELA"].isin([
        "Administración y Negocios",
        "Comunicación",
        "Construcción",
        "Diseño",
        "Informática y Telecomunicaciones",
        "Ingeniería, Medio Ambiente y Recursos Naturales",
        "Salud y Bienestar",
        "Turismo y Hospitalidad",
        "Gastronomía"
    ])].groupby("ESCUELA")["RESPONDIO"].sum().sum()

    print("Total respuestas por sede:", total_respuestas_sede)
    print("Total respuestas por escuela:", total_respuestas_escuela)
    print("Diferencia (respuestas con escuelas no válidas):", total_respuestas_sede - total_respuestas_escuela)

    print(df_estudiantes.loc[~df_estudiantes["ESCUELA"].isin([
        "Administración y Negocios",
        "Comunicación",
        "Construcción",
        "Diseño",
        "Informática y Telecomunicaciones",
        "Ingeniería, Medio Ambiente y Recursos Naturales",
        "Salud y Bienestar",
        "Turismo y Hospitalidad",
        "Gastronomía"
    ]), "ESCUELA"].unique())


    df_cantidad_estudiantes = pd.read_excel(cantidad_estudiantes_file)
    df_cantidad_estudiantes = df_cantidad_estudiantes.rename(columns={
        "Etiquetas de fila": "ESCUELA",
        "Suma de Matrícula": "Cantidad de estudiantes"
    })

    # --- Resumen por escuela ---
    resumen_escuela = resumen_por_escuela(df_estudiantes, df_cantidad_estudiantes)
    resumen_sede = resumen_por_sede(cantidad_estudiantes_file, df_estudiantes, n_referencial_file)
    resumen_escuela_y_sede = resumen_escuela_por_sede(df_cantidad_estudiantes, df_estudiantes, n_referencial_file)

    # --- Guardar archivos ---
    guardar_resumen(resumen_sede, os.path.join(output_dir, "resumen_sedes.xlsx"))
    guardar_resumen(resumen_escuela, os.path.join(output_dir, "resumen_escuelas.xlsx"))
    guardar_resumen(resumen_escuela_y_sede, os.path.join(output_dir, "resumen_escuelas_por_sede.xlsx"))

    return resumen_sede, resumen_escuela, resumen_escuela_y_sede
