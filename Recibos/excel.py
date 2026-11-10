import pandas as pd
from tkinter import StringVar
from tkinter.ttk import Combobox


def cargar_excel(ruta):

    try:
        df = pd.read_excel(ruta)

        nombres = df["Nombre"].tolist()
        documentos = df["Documento"].tolist()

        return list(zip(nombres, documentos))
    

    except Exception as e:
        print("No fue posible cargar el excel : ", e)
        return [] #Lista Vacia
   
def Combo_box_config(ventana , datos):


    trabajador_var = StringVar()
    combo = Combobox(ventana, textvariable=trabajador_var, state="readonly", font=("Arial",10))

    nombres = [item[0] for item in datos]
    combo['values'] = nombres
    combo.grid(column=1, row= 1 ,pady=5, padx=10)
    combo.set("Seleccionar Trabajador")

    def seleccion(event):
        indice = combo.current()

        if indice != -1:
            trabajador , documento = datos[indice]
            print(trabajador)
            print(documento)

        combo.bind("<<ComboboxSelected>>", seleccion)
    return combo


