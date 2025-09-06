import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
import os
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Estilo global para los gráficos
plt.style.use('ggplot')
plt.rcParams.update({
    'font.size': 12,
    'figure.dpi': 150
})

def crear_grafico_pie(df, col_respuestas, col_total, output_path):
    responded = df[col_respuestas].sum()
    not_responded = df[col_total].sum() - responded
    sizes = [responded, not_responded]
    labels_base = ['Encuestas respondidas', 'Encuestas no respondidas']
    
    # --- Crear labels con porcentaje y valor ---
    total = sum(sizes)
    labels = [f"{label}\n{size} ({size/total*100:.2f}%)"
              for label, size in zip(labels_base, sizes)]
    
    colors = ['#000000', '#F8B416']  # negro y amarillo
    explode = (0.05, 0.05)

    fig, ax = plt.subplots(figsize=(5,5))

    wedges, texts = ax.pie(
        sizes,
        labels=labels,
        startangle=90,
        colors=colors,
        explode=explode,
        wedgeprops={'edgecolor':'white', 'linewidth':1.5},
        textprops={'fontsize':12},
    )

    # --- Anillo blanco en el centro ---
    centre_circle = plt.Circle((0,-0.05), 0.60, fc='white')
    ax.add_artist(centre_circle)

    ax.axis('equal')  # círculo perfecto
    plt.tight_layout()
    plt.savefig(output_path, transparent=True, bbox_inches='tight')
    print("Gráfico de pastel guardado en:", output_path)
    plt.close()

def crear_grafico_barras(df, col_categoria, col_valor, output_path):
    fig, ax = plt.subplots(figsize=(6,4))
    colors = ["#000000", "#F78DA7", "#CF2E2E", "#FF6900", "#FCB900", "#7BDCB5", "#00D084", "#8ED1FC", "#156082"]
    bars = ax.bar(df[col_categoria], df[col_valor], color=colors, edgecolor='black')

    ax.set_ylim(0, 0.1)
    ax.set_xticks([])
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('black')   
    ax.spines['bottom'].set_color('black')

    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width()/2,  # posición x centrada
            height + 0.002,                   # un poco arriba de la barra
            f"{height:.2f}%",             # formato porcentaje
            ha='center', va='bottom', fontsize=9, fontweight='bold'
        )
    
    plt.tight_layout()
    plt.savefig(output_path, transparent=True, bbox_inches='tight')
    plt.close()

def rellenar_tabla(slide, placeholder_name, df):
    for shape in slide.shapes:
        if shape.has_text_frame and placeholder_name in shape.text:
            rows, cols = df.shape[0] + 1, df.shape[1]
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            table = slide.shapes.add_table(
                rows, cols, left, top, int(width), int(height)
            ).table

            # --- Ajustar anchos de columnas uniformemente ---
            for col in table.columns:
                col.width = int((width*0.8) // cols)

            # --- Ajustar altos de filas uniformemente ---
            for row in table.rows:
                row.height = int((height*0.6) // rows)

            # --- Cabeceras ---
            for j, col_name in enumerate(df.columns):
                cell = table.cell(0, j)
                cell.text = str(col_name)

                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(142, 209, 252)

                # Estilo de cabecera
                text_frame = cell.text_frame
                p = text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run = p.runs[0]
                run.font.size = Pt(9)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)

            # --- Contenido ---
            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    cell = table.cell(i + 1, j)
                    cell.text = str(df.iloc[i, j])

                    # Estilo de contenido
                    text_frame = cell.text_frame
                    p = text_frame.paragraphs[0]
                    p.alignment = PP_ALIGN.CENTER
                    run = p.runs[0]
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(50, 50, 50)

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

    # --- Crear gráficos ---
    crear_grafico_pie(resumen_sede, 'Respuestas', 'Alumnos', "grafico_anillo.png")
    crear_grafico_barras(resumen_escuela, 'ESCUELA', '%', "grafico_barras.png")

    # --- Reemplazar placeholders ---
    for slide in prs.slides:
        rellenar_tabla(slide, "TABLE_PLACEHOLDER_1", resumen_sede)
        rellenar_tabla(slide, "TABLE_PLACEHOLDER_2", resumen_escuela)
        rellenar_grafico(slide, "CHART_PLACEHOLDER_1", "grafico_anillo.png")
        rellenar_grafico(slide, "CHART_PLACEHOLDER_2", "grafico_barras.png")

    prs.save(output_path)
    print(f"Presentación generada en '{output_path}'")
