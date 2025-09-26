import os
import io
import re
from flask import Flask, render_template, request, send_file
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pystrich.datamatrix import DataMatrixEncoder
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import matplotlib
matplotlib.use('Agg') # IMPORTANTE: Evita que matplotlib intente abrir una ventana gráfica
import matplotlib.pyplot as plt
import matplotlib.patches as patches

app = Flask(__name__)

# --- LÓGICA COPIADA DEL NOTEBOOK (ahora como funciones) ---

# Función para crear las flechas (exactamente como en tu notebook)
def create_arrow_image(direction, size_mm):
    size_inches = size_mm / 25.4
    fig, ax = plt.subplots(figsize=(size_inches, size_inches), dpi=300)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis('off')
    if direction == "down":
        arrow = patches.FancyArrow(0.5, 0.9, 0, -0.6, width=0.2, head_width=0.5, head_length=0.2, fc='black', ec='black')
    elif direction == "up":
        arrow = patches.FancyArrow(0.5, 0.1, 0, 0.6, width=0.2, head_width=0.5, head_length=0.2, fc='black', ec='black')
    else:
        plt.close(fig)
        return None
    ax.add_patch(arrow)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', transparent=True, bbox_inches='tight', pad_inches=0)
    buf.seek(0)
    plt.close(fig)
    return ImageReader(buf)

# Función principal que genera el PDF (adaptada para no depender de Colab)
def generate_label_pdf_from_dataframe(dataframe):
    pdf_buffer = io.BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=landscape(A4))
    width, height = landscape(A4)
    
    # Intentamos registrar una fuente. Es mejor poner el archivo de la fuente en la carpeta 'static'
    try:
        font_path = os.path.join('static', 'arial-black.ttf') # Suponiendo que tienes arial-black.ttf en 'static'
        pdfmetrics.registerFont(TTFont('Arial-Black', font_path))
        font_name = "Arial-Black"
    except:
        font_name = "Helvetica-Bold"
        print("Advertencia: No se pudo registrar Arial-Black. Usando fuente por defecto.")

    label_width_pt = 260 * mm
    label_height_pt = 80 * mm
    margin_x = (width - label_width_pt) / 2
    y_positions = [height - (label_height_pt * 1) - 20*mm, height - (label_height_pt * 2) - 40*mm]
    
    label_on_page_count = 0
    for index, row in dataframe.iterrows():
        ubicacion = str(row['Ubicaciones'])
        
        if label_on_page_count == 2:
            c.showPage()
            label_on_page_count = 0

        nivel = 0
        match = re.match(r'^.{3}(\d).*', ubicacion)
        if match:
            nivel = int(match.group(1))

        current_x = margin_x
        current_y = y_positions[label_on_page_count]

        # 1. Código Data Matrix
        encoder = DataMatrixEncoder(ubicacion)
        datamatrix_img_data = encoder.get_imagedata()
        datamatrix_image = ImageReader(io.BytesIO(datamatrix_img_data))
        dm_size_pt = 60 * mm
        dm_x = current_x
        dm_y = current_y + (label_height_pt / 2) - (dm_size_pt / 2)
        c.drawImage(datamatrix_image, dm_x, dm_y, width=dm_size_pt, height=dm_size_pt)

        # 2. Texto de Ubicación
        font_size = 80
        c.setFont(font_name, font_size)
        text_ubicacion_x = current_x + (label_width_pt / 2)
        text_height = font_size * 0.8
        text_ubicacion_y = current_y + (label_height_pt / 2) - (text_height / 2)
        c.drawCentredString(text_ubicacion_x, text_ubicacion_y, ubicacion)

        # 3. Flecha
        arrow_image = None
        arrow_size_mm = 50
        arrow_size_pt = arrow_size_mm * mm
        if nivel == 1:
            arrow_image = create_arrow_image("down", arrow_size_mm)
        elif nivel == 2:
            arrow_image = create_arrow_image("up", arrow_size_mm)
        if arrow_image:
            arrow_x = current_x + label_width_pt - arrow_size_pt - 10 * mm
            arrow_y = current_y + (label_height_pt / 2) - (arrow_size_pt / 2)
            c.drawImage(arrow_image, arrow_x, arrow_y, width=arrow_size_pt, height=arrow_size_pt, mask='auto')

        label_on_page_count += 1
    
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer

# --- RUTAS DE FLASK ---

@app.route('/')
def index():
    # Esta ruta solo muestra la página de subida de archivos
    return render_template('index.html')

@app.route('/generar-etiquetas', methods=['POST'])
def generar_etiquetas():
    # Verificamos si se subió un archivo
    if 'archivo_excel' not in request.files:
        return render_template('index.html', error="No se seleccionó ningún archivo.")
    
    file = request.files['archivo_excel']
    
    # Verificamos si el archivo tiene nombre (no es un envío vacío)
    if file.filename == '':
        return render_template('index.html', error="No se seleccionó ningún archivo.")

    try:
        # Leemos el archivo Excel directamente desde el objeto de archivo subido
        df = pd.read_excel(file)
        
        # Verificamos que la columna 'Ubicaciones' exista
        if 'Ubicaciones' not in df.columns:
            columnas = ', '.join(df.columns)
            return render_template('index.html', error=f"Error: La columna 'Ubicaciones' no existe. Columnas encontradas: {columnas}")
            
        # Llamamos a nuestra función de procesamiento
        pdf_file = generate_label_pdf_from_dataframe(df)
        
        # Enviamos el PDF generado al usuario para que lo descargue
        return send_file(
            pdf_file,
            mimetype='application/pdf',
            as_attachment=True,
            download_name='etiquetas_generadas.pdf'
        )

    except Exception as e:
        # Si algo falla al leer el Excel o al generar el PDF, mostramos un error
        return render_template('index.html', error=f"Ocurrió un error: {e}")
# --- NUEVA RUTA: Para descargar la plantilla de ejemplo ---

@app.route('/descargar-plantilla')
def descargar_plantilla():
    try:
        # La ruta al archivo que pusimos en la carpeta 'static'
        path_to_file = os.path.join('static', 'plantilla_ubicaciones.xlsx')
        
        return send_file(
            path_to_file,
            as_attachment=True
        )
    except FileNotFoundError:
        return "Error: El archivo 'plantilla_ubicaciones.xlsx' no se encontró en el servidor.", 404

if __name__ == '__main__':
    app.run(debug=True)