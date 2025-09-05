import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches
import os
import time

# Estilo global para los gráficos
plt.style.use('seaborn-whitegrid')
plt.rcParams.update({
    'font.size': 12,
    'figure.dpi': 150
})

def crear_grafico_anillo(df, col_respuestas, col_total, output_path):
    responded = df[col_respuestas].sum()
    not_responded = df[col_total].sum() - responded
    sizes = [responded, not_responded]
    labels = ['Encuestas respondidas', 'Encuestas no respondidas']
    colors = ['#F8B416', '#000000']  # amarillo y negro
    explode = (0.05, 0.05)

    fig, ax = plt.subplots(figsize=(5,5))
    wedges, texts, autotexts = ax.pie(
        sizes,
        labels=labels,
        autopct='%1.1f%%',
        startangle=90,
        colors=colors,
        explode=explode,
        wedgeprops={'edgecolor':'white', 'linewidth':1.5},
        textprops={'fontsize':12}
    )

    # --- Convertir en anillo ---
    centre_circle = plt.Circle((0,0),0.60,fc='white')
    fig.gca().add_artist(centre_circle)

    ax.axis('equal')  # círculo perfecto
    plt.tight_layout()
    plt.savefig(output_path, transparent=True, bbox_inches='tight')
    plt.close()

def crear_grafico_barras(df, col_categoria, col_valor, output_path):
    fig, ax = plt.subplots(figsize=(6,4))
    bars = ax.bar(df[col_categoria], df[col_valor], color='#2196F3', edgecolor='black')
    
    ax.set_ylabel('Porcentaje de respuestas', fontsize=12)
    ax.set_xlabel('')
    ax.set_ylim(0, 100)
    ax.set_xticks(range(len(df[col_categoria])))
    ax.set_xticklabels(df[col_categoria], rotation=45, ha='right', fontsize=11)
    
    # Etiquetas encima de cada barra
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height:.1f}%',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0,3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=10)
    
    plt.tight_layout()
    plt.savefig(output_path, transparent=True, bbox_inches='tight')
    plt.close()

def rellenar_tabla(slide, placeholder_name, df):
    for shape in slide.shapes:
        if shape.has_text_frame and placeholder_name in shape.text:
            rows, cols = df.shape[0]+1, df.shape[1]
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            table = slide.shapes.add_table(rows, cols, left, top, width, height).table

            # Cabeceras
            for j, col_name in enumerate(df.columns):
                table.cell(0, j).text = str(col_name)

            # Contenido
            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    table.cell(i+1, j).text = str(df.iloc[i,j])

            shape.text = ""  # eliminar placeholder
            break

def rellenar_grafico(slide, placeholder_name, image_path):
    for shape in slide.shapes:
        if shape.has_text_frame and placeholder_name in shape.text:
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            slide.shapes.add_picture(image_path, left, top, width, height)
            shape.text = ""  # eliminar placeholder
            break

def generar_presentacion(template_path, output_path, resumen_sede, resumen_escuela):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs = Presentation(template_path)

    # --- Crear nombres únicos de archivo ---
    timestamp = int(time.time())
    grafico_anillo_file = f"grafico_anillo_{timestamp}.png"
    grafico_barras_file = f"grafico_barras_{timestamp}.png"

    # --- Crear gráficos ---
    crear_grafico_anillo(resumen_sede, 'RESPUESTAS', 'N_ALUMNOS', grafico_anillo_file)
    crear_grafico_barras(resumen_escuela, 'ESCUELA', '%', grafico_barras_file)

    # --- Reemplazar placeholders ---
    for slide in prs.slides:
        rellenar_tabla(slide, "TABLE_PLACEHOLDER_1", resumen_sede)
        rellenar_tabla(slide, "TABLE_PLACEHOLDER_2", resumen_escuela)
        rellenar_grafico(slide, "CHART_PLACEHOLDER_1", grafico_anillo_file)
        rellenar_grafico(slide, "CHART_PLACEHOLDER_2", grafico_barras_file)

    prs.save(output_path)
    print(f"Presentación generada en '{output_path}'")
