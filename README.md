from pptx import Presentation
from pptx.util import Inches

# Create a presentation object
prs = Presentation()

# Add a slide with title and content layout
slide_layout = prs.slide_layouts[5]  # Using the Title and Content layout
slide = prs.slides.add_slide(slide_layout)

# Set the title
title = slide.shapes.title
title.text = "Cuadro Comparativo de Estilos o Periodos Musicales"

# Define the table dimensions
rows = 8
cols = 7
left = Inches(0.5)
top = Inches(1.5)
width = Inches(9)
height = Inches(4)

# Add a table to the slide
table = slide.shapes.add_table(rows, cols, left, top, width, height).table

# Set column widths
for i in range(cols):
    table.columns[i].width = Inches(1.2)

# Define the data for the table
data = [
    ["Criterio", "Edad Media", "Renacimiento", "Barroco", "Clasicismo", "Romanticismo", "Impresionismo"],
    ["Temporalidad", "Siglos V - XV", "Siglos XV - XVI", "Siglo XVII - primera mitad del XVIII", 
     "Segunda mitad del siglo XVIII", "Siglo XIX", "Finales del siglo XIX - principios del XX"],
    ["Ideología", "Religiosa, Feudalismo", "Humanismo, Renacimiento cultural", "Absolutismo, Contrarreforma",
     "Ilustración, Racionalismo", "Nacionalismo, Individualismo", "Evocación de sensaciones, subjetivismo"],
    ["Ritmo", "Libre, no mensural", "Regular, surgimiento del compás", "Compás definido, uso del bajo continuo",
     "Regular, simétrico", "Flexibilidad rítmica, rubato", "Sutil, enfocado en atmósferas"],
    ["Melodía", "Monofónica, luego polifónica", "Lírica, polifónica, modal", "Ornamentada, contrapuntística",
     "Clara, cantabile", "Expresiva, amplia, con grandes arcos", "Lírica, fragmentada, sin líneas definidas"],
    ["Armonía", "Modal", "Modal con tendencia al mayor-menor", "Tonal, uso del bajo continuo",
     "Tonal, uso de la cadencia perfecta", "Riqueza armónica, cromatismo", "Ambigüedad tonal, uso de escalas exóticas"],
    ["Instrumentos", "Voces, órgano, laúd", "Voces, laúd, violas, clavecín", "Clavecín, órgano, cuerdas, metales, maderas",
     "Piano, cuarteto de cuerdas, orquesta", "Piano, orquesta sinfónica, instrumentos nacionales", "Piano, arpa, flauta, clarinete, cuerdas"],
    ["Formas Musicales", "Canto gregoriano, motete", "Motete, madrigal, misa", "Suite, sonata, concierto grosso, ópera",
     "Sinfonía, sonata, concierto, ópera", "Sinfonía, lied, poema sinfónico, ópera", "Preludio, nocturno, piezas breves, sinfonía de carácter programático"],
    ["Compositores", "Guido d'Arezzo, Leonín, Perotín", "Palestrina, Josquin des Prez", "Bach, Händel, Vivaldi",
     "Haydn, Mozart, Beethoven", "Chopin, Schumann, Wagner", "Debussy, Ravel"]
]

# Populate the table with data
for row_idx, row_data in enumerate(data):
    for col_idx, cell_value in enumerate(row_data):
        table.cell(row_idx, col_idx).text = cell_value

# Save the presentation
ppt_output_path = "/mnt/data/Cuadro_Comparativo_Estilos_Musicales.pptx"
prs.save(ppt_output_path)

ppt_output_path
