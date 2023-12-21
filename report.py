from docx import Document
import pandas as pd
from datetime import datetime
from docx.shared import Inches, Pt, RGBColor  
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import sys
import folium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import tempfile

# Función para agregar una cabecera al documento
def agregar_cabecera(doc):
    cabecera = doc.sections[0].header  # Obtener la cabecera de la primera sección del documento
    cabecera.paragraphs[0].clear()  # Limpiar el contenido existente en la cabecera

    # Agregar el texto de la cabecera
    parrafo_cabecera = cabecera.paragraphs[0]
    texto_cabecera = "CONTRACTACIÓ PER LA PRESENTACIÓ DEL SERVEI D'UNITAT D'INTERVENCIÓ RÀPIDA, PEL MANTENIMENT DE LA VIA PÚBLICA I DELS EDIFICIS MUNICIPALS"
    parrafo_cabecera.text = texto_cabecera.upper()  # Texto de la cabecera en mayúsculas
    parrafo_cabecera.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Alinear el texto al centro
    parrafo_cabecera.runs[0].font.bold = True  # Aplicar negrita al texto
    parrafo_cabecera.runs[0].font.size = Pt(10)  # Tamaño de fuente
    parrafo_cabecera.runs[0].font.color.rgb = RGBColor(0x00, 0x70, 0xC0)  # Color azul ##0070C0

def agregar_pie_de_pagina(doc):
    pie_de_pagina = doc.sections[0].footer  # Obtener el pie de página de la primera sección del documento
    pie_de_pagina.paragraphs[0].clear()  # Limpiar el contenido existente en el pie de página

    # Crear un nuevo párrafo en el pie de página
    parrafo = pie_de_pagina.add_paragraph()

    # Agregar el texto "SOBRE A. CRITERIS AVALUABLES A JUDICI DE VALOR" al párrafo
    run = parrafo.add_run("SOBRE A. CRITERIS AVALUABLES A JUDICI DE VALOR")
    run.font.bold = True  # Aplicar negrita al texto
    run.font.size = Pt(9)  # Tamaño de fuente
    run.font.color.rgb = RGBColor(0x00, 0x70, 0xC0)  # Color azul #0088FF
    parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Alinear el texto al centro

def set_background_color(cell, color):
    # Obtener el elemento "tcPr" que representa las propiedades de la celda
    tc_pr = cell._element.get_or_add_tcPr()

    # Crear un elemento "shd" para establecer el color de fondo
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)
    tc_pr.append(shd)

def set_font_style(element, font_name):
    element.rPr.rFonts.set(qn('w:ascii'), font_name)
    element.rPr.rFonts.set(qn('w:hAnsi'), font_name)
    element.rPr.rFonts.set(qn('w:cs'), font_name)

def apply_font_format(cell, bold=False, font_color=None):
    # Obtener el párrafo de la celda
    paragraph = cell.paragraphs[0]

    # Obtener el elemento "r" que representa el texto del párrafo
    run = paragraph.runs[0]

    # Aplicar formato de fuente

    run.font.size = Pt(11)

    if bold==True:  
        run.bold = bold

    # Configurar fuente Arial
    run.font.name = "Arial"
    if font_color:
        run.font.color.rgb = font_color

def leer_datos_desde_excel(ruta_excel):
    # Leer el archivo Excel y cargarlo en un DataFrame de pandas
    df = pd.read_excel(ruta_excel)

    # Lista para almacenar los datos de cada fila
    datos_filas = []

    # Leer los datos de cada fila y agregarlos a la lista
    for _, fila in df.iterrows():
        titulo = fila["_title"]
        
        # Recogiendo los tipos de desperfectos, solo si existen (no son NaN)
        desperfectos = []
        for col in ["tipus_de_desperfecte", "2_tipus_de_desperfecte", "3_tipus_de_desperfecte"]:
            if pd.notna(fila[col]):
                desperfectos.append(fila[col])

        # Otros valores de interés
        lloc = str(fila["lloc"])
        lloc_thoroughfare = fila["lloc_thoroughfare"]  # Asegúrate de que este nombre de columna es correcto
        amidaments = [fila[f"{i}_amidament"] for i in range(1, 4)]
        unitats = [fila[f"{i}_unitats"] for i in range(1, 4)]
        fecha = fila["data"]
        imatges = fila["imatges"]
        latitud = fila["_latitude"]
        longitud = fila["_longitude"]

        # Agregar la fila de datos a la lista
        datos_filas.append((titulo, desperfectos, lloc_thoroughfare, lloc, amidaments, unitats, fecha, imatges, latitud, longitud))
        
    return datos_filas


def combinar_strings(lista1, lista2):
    # Combinar las listas en el formato deseado
    combinacion = []
    for i, item in enumerate(lista1):
        valor = lista2[i] if i < len(lista2) else ""
        combinacion.append(f"{item} {valor}")
    return ", ".join(combinacion)

def generar_mapa_folium(latitud, longitud):
    mapa = folium.Map(location=[latitud, longitud], max_zoom=19, zoom_start=19)
    
    # Agregar un marcador en la ubicación indicada por las coordenadas de latitud y longitud
    folium.Marker([latitud, longitud]).add_to(mapa)

    # Guardar el mapa en un archivo HTML temporal
    mapa_html = tempfile.NamedTemporaryFile(suffix=".html", delete=False)
    mapa.save(mapa_html.name)
    mapa_html.close()

    return mapa_html.name

def crear_driver_web():
    # Determinar la ruta del directorio actual del script o ejecutable
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

    # Agregar la ruta del directorio al PATH
    os.environ['PATH'] += os.pathsep + base_path

    # Configurar opciones para Chrome
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')

    # Crear y retornar el driver de Chrome sin especificar la ruta
    return webdriver.Chrome(options=options)


def crear_tablas_informes(datos_filas,carpeta_imagenes):
    # Crear el objeto Document
    doc = Document()

    # Crear una tabla para cada fila de datos
    for datos_fila in datos_filas:
        titulo, desperfectos,lloc_thoroughfare, lloc, amidaments, unitats, fecha, imatges, latitud, longitud = datos_fila

        # Dividir el título en elementos
        elementos = titulo.split(", ")

        # Crear la tabla
        tabla = doc.add_table(rows=1, cols=2)
        tabla.style = 'Table Grid'

        # Configurar la primera fila con la descripción de la incidencia
        tabla.cell(0, 0).text = "Incidència"
        descripcion_incidencia = ", ".join([f"{elem} {desp}" for elem, desp in zip(elementos, desperfectos)]) + f" a {lloc_thoroughfare}"
        tabla.cell(0, 1).text = descripcion_incidencia
        set_background_color(tabla.cell(0, 0), "002060")
        apply_font_format(tabla.cell(0, 0), bold=True, font_color=RGBColor(255, 255, 255))

        # Añadir filas para los desperfectos
        for i in range(len(desperfectos)):
            if pd.notna(amidaments[i]):  # Solo añadir si hay amidament
                row_cells = tabla.add_row().cells
                row_cells[0].text = f"Desperfecte {i + 1}"
                descripcion = f"{elementos[i]} {desperfectos[i]} - {amidaments[i]} {unitats[i]}"
                row_cells[1].text = descripcion
                set_background_color(row_cells[0], "002060")
                apply_font_format(row_cells[0], bold=True, font_color=RGBColor(255, 255, 255))

        # Añadir filas para fecha y lugar
        for i, valor in enumerate(["Data", "Lloc"], start=1):
            row_cells = tabla.add_row().cells
            row_cells[0].text = valor
            row_cells[1].text = fecha if valor == "Data" else lloc
            set_background_color(row_cells[0], "002060")
            apply_font_format(row_cells[0], bold=True, font_color=RGBColor(255, 255, 255))

        # Agregar imágenes dentro de la tabla
        identificadores_imagenes = imatges.split(",")[:2]  # Tomar los dos primeros identificadores
        if len(identificadores_imagenes) > 0:
            row_imagenes = tabla.add_row().cells
            for i, identificador in enumerate(identificadores_imagenes):
                ruta_imagen = os.path.join(carpeta_imagenes, f"{identificador}.jpg")
                if os.path.exists(ruta_imagen):
                    run = row_imagenes[i].paragraphs[0].add_run()
                    run.add_picture(ruta_imagen, height=Inches(2.0))
                    row_imagenes[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Generar y agregar el mapa
        print("Generando mapa con Folium")
        mapa_html_file = generar_mapa_folium(latitud, longitud)
        print("Mapa generado: ", mapa_html_file)
        options = Options()
        options.add_argument('--headless')
        driver = crear_driver_web()
        driver.get("file:///" + mapa_html_file)
        driver.set_window_size(800, 600)
        ruta_imagen_geolocalizacion = tempfile.mktemp(suffix=".png")
        driver.save_screenshot(ruta_imagen_geolocalizacion)
        driver.quit()
        os.remove(mapa_html_file)
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        run.add_picture(ruta_imagen_geolocalizacion, width=Inches(4.0))
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        os.remove(ruta_imagen_geolocalizacion)

        # Agregar un salto de página después de cada tabla
        doc.add_page_break()

    # Agregar cabecera y pie de página
    agregar_cabecera(doc)
    agregar_pie_de_pagina(doc)

    # Guardar el documento en un archivo
    doc.save("informes_word.docx")

import tkinter as tk
from tkinter import filedialog

def generar_informes(ruta_excel, carpeta_imagenes):
    print("Inicio de generar_informes")
    datos_filas = leer_datos_desde_excel(ruta_excel)
    crear_tablas_informes(datos_filas, carpeta_imagenes)
    print("Fin de generar_informes")


def main():
    def browse_file():
        filename = filedialog.askopenfilename()
        entry_excel_path.delete(0, tk.END)
        entry_excel_path.insert(0, filename)

    def browse_folder():
        foldername = filedialog.askdirectory()
        entry_photos_path.delete(0, tk.END)
        entry_photos_path.insert(0, foldername)

    def execute_script():
        ruta_excel = entry_excel_path.get()
        carpeta_imagenes = entry_photos_path.get()
        # Aquí llamas a tu función principal del script con ruta_excel y carpeta_imagenes
        generar_informes(ruta_excel, carpeta_imagenes)

    root = tk.Tk()
    root.title("Generador de Informes")

    tk.Label(root, text="Ruta al archivo Excel:").pack()
    entry_excel_path = tk.Entry(root, width=50)
    entry_excel_path.pack()
    tk.Button(root, text="Buscar", command=browse_file).pack()

    tk.Label(root, text="Ruta a la carpeta de fotos:").pack()
    entry_photos_path = tk.Entry(root, width=50)
    entry_photos_path.pack()
    tk.Button(root, text="Buscar", command=browse_folder).pack()

    tk.Button(root, text="Generar Informes", command=execute_script).pack()

    root.mainloop()

if __name__ == "__main__":
    main()

