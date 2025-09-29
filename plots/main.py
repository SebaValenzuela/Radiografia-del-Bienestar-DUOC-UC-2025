from io import BytesIO
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
import os
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import numpy as np
import matplotlib.ticker as mticker
from pptx.util import Inches
from pptx.util import Cm
from copy import deepcopy

# Estilo global para los gráficos
plt.style.use('ggplot')
plt.rcParams.update({
    'font.size': 12,
    'figure.dpi': 150
})

def crear_grafico_pie(df, col_respuestas, col_total, output_path):
    responded = df[col_respuestas].sum()
    not_responded = max(df[col_total].sum() - responded, 0)
    sizes = [responded, not_responded]
    labels_base = ['Encuestas respondidas', 'Encuestas no respondidas']
    
    # --- Crear labels con porcentaje y valor ---
    total = sum(sizes)
    if total == 0:
        labels = [f"{label}\n0 (0%)" for label in labels_base]
    else:
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

def crear_grafico_barras(df, col_categoria, col_valor, output_path, fontsize=12):
    fig, ax = plt.subplots(figsize=(6,4))
    x = np.arange(len(df[col_categoria]))
    bars = ax.bar(x, df[col_valor], edgecolor='black', color='royalblue',width=0.4)
    grid_color = '#CCCCCC'
    grid_linestyle = '-'

    ax.set_ylim(0, 100)
    ax.set_yticks(np.arange(0, 101, 10))
    ax.yaxis.set_major_formatter(mticker.PercentFormatter())
    ax.tick_params(left=False, bottom=False)
    ax.set_xticks(x)
    ax.set_xticklabels(df[col_categoria], rotation=45, ha='right', fontsize=fontsize)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.yaxis.grid(True, color=grid_color, linestyle=grid_linestyle)
    ax.xaxis.grid(False)
    
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

            for col in table.columns:
                col.width = int((width*0.8) // cols)

            for row in table.rows:
                row.height = int((height*0.6) // rows)

            for j, col_name in enumerate(df.columns):
                cell = table.cell(0, j)
                cell.text = str(col_name)

                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(142, 209, 252)

                text_frame = cell.text_frame
                p = text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run = p.runs[0]
                run.font.size = Pt(9)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
            
            if "%" in df.columns:
                valores = df["%"]
                top3_idx = valores.nlargest(3).index
                bottom3_idx = valores.nsmallest(3).index
            else:
                top3_idx, bottom3_idx = [], []


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
                
                    if df.columns[j] == "% de avance respecto a total":
                        if i in top3_idx:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(11, 241, 30)  # verde
                        elif i in bottom3_idx:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(229, 239, 29)  # amarillo

            shape.text = ""  # eliminar placeholder
            break

def rellenar_grafico(slide, placeholder_name, image_path):
    for shape in slide.shapes:
        if shape.has_text_frame and placeholder_name in shape.text:
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            slide.shapes.add_picture(image_path, left, top, width, height)
            shape.text = ""  # eliminar placeholder
            break

def rellenar_tabla_parcial(prs, slide_template, placeholder_name, df, start_row=0, max_rows_por_slide=10):
    total_rows = df.shape[0]
    end_row = min(start_row + max_rows_por_slide, total_rows)
    
    df_parcial = df.iloc[start_row:end_row]
    
    # Copiar slide
    slide = copiar_slide(prs, slide_template)
    
    for shape in slide.shapes:
        if shape.has_text_frame and placeholder_name in shape.text:
            rows, cols = df_parcial.shape[0] + 1, df_parcial.shape[1]
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            table = slide.shapes.add_table(
                rows, cols, left, top, width, height
            ).table

            # Ajustar tamaño de columnas y filas
            for col in table.columns:
                col.width = int((width*0.8) // cols)
            for row in table.rows:
                row.height = int((height*0.6) // rows)

            # Encabezado
            for j, col_name in enumerate(df_parcial.columns):
                cell = table.cell(0, j)
                cell.text = str(col_name)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(142, 209, 252)
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                run = p.runs[0]
                run.font.size = Pt(9)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)

            # Resaltar top3/bottom3 si hay columna "%"
            if "%" in df_parcial.columns:
                valores = df_parcial["%"]
                top3_idx = valores.nlargest(3).index
                bottom3_idx = valores.nsmallest(3).index
            else:
                top3_idx, bottom3_idx = [], []

            # Contenido
            for i in range(df_parcial.shape[0]):
                for j in range(df_parcial.shape[1]):
                    cell = table.cell(i + 1, j)
                    cell.text = str(df_parcial.iloc[i, j])
                    p = cell.text_frame.paragraphs[0]
                    p.alignment = PP_ALIGN.CENTER
                    run = p.runs[0]
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(50, 50, 50)

                    # Colorear top3/bottom3
                    if df_parcial.columns[j] == "% de avance respecto a total":
                        if df_parcial.index[i] in top3_idx:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(11, 241, 30)  # verde
                        elif df_parcial.index[i] in bottom3_idx:
                            cell.fill.solid()
                            cell.fill.fore_color.rgb = RGBColor(229, 239, 29)  # amarillo

            shape.text = ""  # eliminar placeholder
            break

    # Si quedan más filas, llamar recursivamente
    if end_row < total_rows:
        rellenar_tabla_parcial(prs, slide_template, placeholder_name, df, start_row=end_row, max_rows_por_slide=max_rows_por_slide)



def copiar_slide(prs, slide):
    """Copia un slide completo y devuelve el nuevo slide"""
    # Agregar nueva slide vacía con el mismo layout que la original
    new_slide = prs.slides.add_slide(slide.slide_layout)
    
    # Copiar todos los elementos (formas, tablas, gráficos, etc.)
    for shape in slide.shapes:
        new_el = deepcopy(shape._element)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    
    return new_slide

def generar_presentacion(template_path, output_path, resumen_sede, resumen_escuela, resumen_escuela_y_sede):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs = Presentation(template_path)

    # --- Crear gráficos ---
    crear_grafico_pie(resumen_sede, 'Cantidad de respuestas', 'Cantidad de estudiantes', "grafico_avance_global.png")
    crear_grafico_barras(resumen_sede, 'SEDE', '% de avance respecto a total', "grafico_avance_sedes.png", fontsize=10)
    crear_grafico_barras(resumen_escuela, 'ESCUELA', '% de avance respecto a total', "grafico_avance_escuelas.png", fontsize=8)

    # --- Buscar slide template para la tabla larga ---
    template_slide = None
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and "tabla_resumen_escuela_en_cada_sede" in shape.text:
                template_slide = slide
                break
        if template_slide:
            break

    # --- Reemplazar placeholders de gráficos y tablas ---
    for slide in prs.slides:
        rellenar_grafico(slide, "grafico_avance_global", "grafico_avance_global.png")
        rellenar_tabla(slide, "tabla_avance_sedes", resumen_sede)
        rellenar_grafico(slide, "grafico_avance_sedes", "grafico_avance_sedes.png")
        rellenar_tabla(slide, "tabla_avance_escuelas", resumen_escuela)
        rellenar_grafico(slide, "grafico_avance_escuelas", "grafico_avance_escuelas.png")

    # --- Rellenar tabla larga en varias diapositivas ---
    if template_slide:
        rellenar_tabla_parcial(prs, template_slide, "tabla_resumen_escuela_en_cada_sede", resumen_escuela_y_sede, max_rows_por_slide=9)
        # opcional: eliminar la diapositiva template original
        # prs.slides._sldIdLst.remove(template_slide._element)

    prs.save(output_path)
    print(f"Presentación generada en '{output_path}'")

