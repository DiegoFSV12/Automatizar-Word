import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

doc = DocxTemplate("plantilla.docx")

nombre = 'Sebastian Eduardo'
telefono = '948 987 074'
correo = 'sebastian@gmail.com'
fecha = datetime.today().strftime("%d/%m/%Y")
curso = 'Ingenieria se software'
periodo = '2024-02'

constantes = {
    'nombre':nombre,
    'telefono':telefono,
    'correo':correo,
    'fecha':fecha,
    'curso':curso,
    'periodo':periodo
}

df = pd.read_excel('Alumnos.xlsx')
for indice, fila in df.iterrows():
    contenido = {
        'nombre_alumno':fila['Nombre'],
        'nota_u1':fila['U1'],
        'nota_u2':fila['U2'],
        'nota_u3':fila['U3'],
        'nota_final':(float(fila['U1'])+float(fila['U2'])+float(fila['U3']))/3
    }
    contenido.update(constantes)
    doc.render(contenido)
    doc.save(f"notas/notas_de_{fila['Nombre']}.docx")