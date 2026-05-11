# Matem-tica-import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image, ImageDraw

def create_fidelic_images():
    # 1. Heptadecágono
    img1 = Image.new('RGB', (400, 400), color='white')
    draw1 = ImageDraw.Draw(img1)
    draw1.ellipse(, outline='#1f77b4', width=3)
    draw1.line([(200, 200), (200, 50)], fill='red', width=2)
    draw1.text((210, 200), "Origen", fill='black')
    img1.save('q1_heptadecagono.png')

    # 2. Rotación 3x3
    img2 = Image.new('RGB', (400, 200), color='white')
    draw2 = ImageDraw.Draw(img2)
    draw2.rectangle(, outline='black', width=3)
    draw2.line([(200, 20), (200, 120)], fill='black', width=1)
    draw2.line([(150, 70), (250, 70)], fill='black', width=1)
    draw2.text((160, 140), "Secuencia: XY'X'Y", fill='red')
    img2.save('q2_3x3_rotation.png')

    # 3. Acarreo 11
    img3 = Image.new('RGB', (400, 150), color='white')
    draw3 = ImageDraw.Draw(img3)
    draw3.text((140, 50), "6      8      5", fill='black')
    draw3.text((150, 80), " +1     +1", fill='red')
    draw3.text((140, 110), "7      5      3      5", fill='blue')
    img3.save('q3_acarreo_11.png')

    # 5. Criptoaritmética Resta
    img5 = Image.new('RGB', (400, 150), color='white')
    draw5 = ImageDraw.Draw(img5)
    draw5.text((150, 30), "  A B C 0 0 0", fill='black')
    draw5.text((150, 50), "-         A B C", fill='black')
    draw5.line([(140, 70), (250, 70)], fill='black', width=2)
    draw5.text((150, 80), "  D E F 1 3 2", fill='red')
    img5.save('q5_cripto.png')

    # 7. Cerillos
    img7 = Image.new('RGB', (400, 150), color='white')
    draw7 = ImageDraw.Draw(img7)
    draw7.text((150, 60), "XI  -  V  =  IV", fill='black')
    img7.save('q7_cerillos.png')

    # 8. Dados Apilados
    img8 = Image.new('RGB', (200, 300), color='white')
    draw8 = ImageDraw.Draw(img8)
    draw8.rectangle(, outline='black', width=3)
    draw8.rectangle(, outline='black', width=3)
    draw8.rectangle(, outline='black', width=3)
    draw8.text((95, 75), "4", fill='black')
    img8.save('q8_dados.png')

    # 9. Gráfico SPC
    img9 = Image.new('RGB', (400, 200), color='white')
    draw9 = ImageDraw.Draw(img9)
    draw9.line([(50, 50), (350, 50)], fill='red', width=2)
    draw9.line([(50, 100), (350, 100)], fill='green', width=2)
    draw9.line([(50, 150), (350, 150)], fill='red', width=2)
    img9.save('q9_spc.png')

    # 10. Muestra vs Población
    img10 = Image.new('RGB', (400, 200), color='white')
    draw10 = ImageDraw.Draw(img10)
    draw10.ellipse(, outline='blue', width=3)
    draw10.ellipse(, outline='green', width=3)
    draw10.text((100, 95), "Población", fill='blue')
    draw10.text((285, 95), "Muestra", fill='green')
    img10.save('q10_muestra.png')

def create_advanced_docx(filename):
    create_fidelic_images()
    doc = Document()
    
    title = doc.add_heading('Cuestionario Avanzado de Matemáticas y Lógica - 5to Grado', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph('Institución Educativa J.N. Andrews')
    doc.add_paragraph('Profesor: Antonio Ticona\n')

    questions =
        },
        {
            "num": "2",
            "text": "En un rompecabezas numérico de 3x3, las reglas topológicas solo permiten rotar bloques completos de 2x2. Si un jugador aplica exactamente la secuencia algorítmica XY'X'Y (donde X e Y son rotaciones de distintos cuadrantes), ¿qué efecto logra en la cuadrícula?",
            "img": "q2_3x3_rotation.png",
            "options":
        },
        {
            "num": "3",
            "text": "Calcula la suma de las cifras del resultado de multiplicar 685 x 11 usando la regla mental de la propiedad distributiva (acarreo visual).",
            "img": "q3_acarreo_11.png",
            "options": ["a) 20", "b) 19", "c) 21", "d) 18", "e) 22"]
        },
        {
            "num": "4",
            "text": "Un teatro vende entradas a S/ 60 (adultos) y S/ 25 (niños). Si asisten 280 personas y la recaudación es S/ 14000, ¿cuántos adultos asistieron aplicando el diagrama tabular de sustitución?",
            "img": None,
            "options": ["a) 200 adultos", "b) 180 adultos", "c) 250 adultos", "d) 150 adultos", "e) 120 adultos"]
        },
        {
            "num": "5",
            "text": "Al multiplicar un número de tres cifras ABC por 999, las tres últimas cifras del resultado son 132. Si estructuramos la operación como ABC000 - ABC, ¿cuál es el valor de A + B + C?",
            "img": "q5_cripto.png",
            "options": ["a) 20", "b) 18", "c) 19", "d) 21", "e) 22"]
        },
        {
            "num": "6",
            "text": "Si al multiplicar un número desconocido por 5 obtienes 430, ¿qué obtienes si al mismo número original lo multiplicas por 25 (iterando la regla del 5)?",
            "img": None,
            "options": ["a) 2150", "b) 2100", "c) 2250", "d) 2000", "e) 2300"]
        },
        {
            "num": "7",
            "text": "Sobre una mesa se ha formado la ecuación falsa con cerillos 'XI - V = IV'. ¿Qué movimiento lícito de un solo cerillo debes hacer para que el balance matemático sea verdadero sin alterar los operadores?",
            "img": "q7_cerillos.png",
            "options":
        },
        {
            "num": "8",
            "text": "Se apilan 3 dados comunes formando una torre vertical. Si la única cara superior visible de la torre es un 4, ¿cuál es la suma exacta de las 5 caras ocultas que reposan sobre el eje vertical?",
            "img": "q8_dados.png",
            "options": ["a) 17", "b) 20", "c) 21", "d) 18", "e) 14"]
        },
        {
            "num": "9",
            "text": "Para vigilar la variabilidad en la fabricación, un ingeniero diagrama las mediciones históricas entre límites de control superior e inferior respecto a una media. ¿Qué esquema está utilizando?",
            "img": "q9_spc.png",
            "options":
        },
        {
            "num": "10",
            "text": "Una fábrica ensambla 4000 teléfonos diarios. El equipo de calidad extrae aleatoriamente 3 unidades de cada caja para evitar sesgos evaluativos. ¿Qué concepto define esta extracción aislada de la población?",
            "img": "q10_muestra.png",
            "options": ["a) Muestreo aleatorio simple", "b) Parámetro descriptivo", "c) Frecuencia absoluta", "d) Censo poblacional", "e) Variable cualitativa nominal"]
        }
    ]

    for q in questions:
        p = doc.add_paragraph()
        run = p.add_run(f"Pregunta {q['num']}: {q['text']}")
        run.bold = True
        
        if q['img'] and os.path.exists(q['img']):
            doc.add_picture(q['img'], width=Inches(3.0))
            last_p = doc.paragraphs[-1]
            last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for opt in q['options']:
            doc.add_paragraph(opt, style='List Bullet')
        doc.add_paragraph()

    doc.add_page_break()
    doc.add_heading('Clave de Respuestas', level=1)
    
    answers =
    
    for ans in answers:
        doc.add_paragraph(ans)

    doc.save(filename)
    print(f"Cuestionario creado exitosamente: {filename}")

if __name__ == '__main__':
    create_advanced_docx('Cuestionario_Avanzado_5to_Primaria_Fidedigno.docx')
