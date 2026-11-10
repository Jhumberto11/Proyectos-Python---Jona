from tkinter import Tk, Label, Entry, Button, filedialog, messagebox
from docx import Document
from datetime import datetime
import os
from num2words import num2words
from interface import *
from excel import cargar_excel






def generar_recibo(monto, num_words, day_1, day_2, mes, final_day, year, name_person, document_person, plantilla, salida):
    doc = Document(plantilla)
    for parrafos in doc.paragraphs:
        if "{monto}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{monto}", monto)
        if "{num_words}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{num_words}", num_words)
        if "{day_1}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{day_1}", day_1)
        if "{day_2}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{day_2}", day_2)
        if "{mes}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{mes}", mes)
        if "{final_day}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{final_day}", final_day)
        if "{year}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{year}", year)
        if "{name_person}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{name_person}", name_person)
        if "{document_person}" in parrafos.text:
            parrafos.text = parrafos.text.replace("{document_person}", document_person)

    archivo_salida = os.path.join(salida, f"Recibo_{name_person.replace(' ', '_')}.docx")
    doc.save(archivo_salida)
    return archivo_salida

# Función para manejar la generación desde la interfaz gráfica
def generar_desde_interfaz():
    try:
        # Recoger los datos de los campos
        monto = monto_entry.get()
        num_words = num2words(float(monto), lang='es').capitalize() + ""


        indice = combo.current()

        if indice == -1:
            messagebox.showerror("Error")
            return
        name_person,document_person = datos_trabajadores[indice]

        plantilla = plantilla_entry.get()
        salida = salida_entry.get()

        # Validar que todos los campos estén llenos
        if not monto or not name_person or not document_person or not plantilla or not salida:
            messagebox.showerror("Error", "Todos los campos son obligatorios.")
            return

        # Obtener la fecha actual
        fecha_actual = datetime.now()
        day_1 = fecha_actual.strftime("%d")
        day_2 = fecha_actual.strftime("%d")
        mes = fecha_actual.strftime("%B")
        final_day = fecha_actual.strftime("%d")
        year = fecha_actual.strftime("%Y")

        # Generar el recibo
        archivo_generado = generar_recibo(monto, num_words, day_1, day_2, mes, final_day, year, name_person, document_person, plantilla, salida)
        messagebox.showinfo("Éxito", f"Recibo generado: {archivo_generado}")

    except Exception as e:
        messagebox.showerror("Error", f"Ha ocurrido un error: {str(e)}")

# Función para seleccionar plantilla
def seleccionar_plantilla():
    archivo = filedialog.askopenfilename(filetypes=[("Documentos de Word", "*.docx")])
    if archivo:
        plantilla_entry.delete(0, 'end')
        plantilla_entry.insert(0, archivo)  
 
# Función para seleccionar carpeta de salida
def seleccionar_carpeta_salida():
    carpeta = filedialog.askdirectory()
    if carpeta:
        salida_entry.delete(0, 'end')
        salida_entry.insert(0, carpeta)
