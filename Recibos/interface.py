from tkinter import Tk, Label, Entry, Button, filedialog, messagebox
from tkcalendar import Calendar
from excel import *
from num2words import num2words
from datetime import datetime
from docx import Document
import os
import locale



excel_path = r"C:\Users\imosa\OneDrive\Escritorio\Trabajadores.xlsx"
datos_trabajadores = cargar_excel(excel_path)
locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')


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
        day_1 = calendar.get_date().split('/')[0]

        day_2 = calendar2.get_date().split('/')[0]
        mes = fecha_actual.strftime("%B")
        final_day = num2words(int(day_2), lang="es")
        year = num2words(int(fecha_actual.strftime("%Y")), lang="es")

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














# Configuración de la ventana principal
from tkinter import Tk, Label, Entry, Button
from tkinter.ttk import Combobox
from tkcalendar import Calendar

# Crear ventana principal
ventana = Tk()
ventana.title("Generador de Recibos")
ventana.geometry("1000x1000")  # Ajustar el tamaño para mejor distribución

# Configurar fuente global
font_labels = ("Arial", 10)
font_entries = ("Arial", 10)

# Etiqueta y campo para el monto
Label(ventana, text="Monto:", font=font_labels).grid(row=0, column=0, padx=10, pady=5, sticky="e")
monto_entry = Entry(ventana, font=font_entries)
monto_entry.grid(row=0, column=1, padx=10, pady=5)

# ComboBox para seleccionar trabajadores
Label(ventana, text="Trabajadores:", font=font_labels).grid(row=1, column=0, padx=10, pady=5, sticky="e")
combo = Combo_box_config(ventana=ventana, datos=datos_trabajadores)

# Calendario para seleccionar el primer día
Label(ventana, text="Primer día:", font=font_labels).grid(row=2, column=0, padx=10, pady=5, sticky="e")
calendar = Calendar(ventana, selectmode="day", date_pattern='dd/mm/yyyy')
calendar.grid(row=2, column=1, padx=10, pady=5, sticky="w")


Label(ventana, text="Segundo día:", font=font_labels).grid(row=2, column=3, padx=10, pady=5, sticky="e")
calendar2 = Calendar(ventana, selectmode="day", date_pattern='dd/mm/yyyy')
calendar2.grid(row=2, column=4, padx=10, pady=5, sticky="w")

# Campo para seleccionar plantilla
Label(ventana, text="Plantilla:", font=font_labels).grid(row=3, column=0, padx=10, pady=5, sticky="e")
plantilla_entry = Entry(ventana, font=font_entries, width=30)
plantilla_entry.grid(row=3, column=1, padx=10, pady=5)
Button(ventana, text="Seleccionar", command=seleccionar_plantilla).grid(row=3, column=2, padx=10, pady=5)

# Campo para seleccionar carpeta de salida
Label(ventana, text="Carpeta de salida:", font=font_labels).grid(row=4, column=0, padx=10, pady=5, sticky="e")
salida_entry = Entry(ventana, font=font_entries, width=30)
salida_entry.grid(row=4, column=1, padx=10, pady=5)
Button(ventana, text="Seleccionar", command=seleccionar_carpeta_salida).grid(row=4, column=2, padx=10, pady=5)

# Botón para generar el recibo
Button(
    ventana,
    text="Generar Recibo",
    command=generar_desde_interfaz,
    bg="green",
    fg="white",
    font=("Arial", 12, "bold")
).grid(row=5, column=0, columnspan=3, pady=20, ipadx=10, ipady=5)

# Ajustar columnas para mejorar la distribución
ventana.grid_columnconfigure(0, weight=1)
ventana.grid_columnconfigure(1, weight=3)
ventana.grid_columnconfigure(2, weight=1)

# Ejecutar la ventana
ventana.mainloop()







