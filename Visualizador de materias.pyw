import tkinter
import  tkinter.filedialog as fd
from tkinter import ttk, messagebox
from os import path

def elegirArchivos():
    #ListaArchivos
    global Archivos
    global filename
    #ListaArchivos = []
    if menu.get() == "Mod1" or menu.get() == "Mod2" or menu.get() == "Mod4" or menu.get() == "Mod5":
        Archivos = fd.askopenfilename(parent=ventana, title= "Elija los archivos")
        name = path.basename(Archivos)
        mostrarArchivos["text"] = name
        


opciones = [
    "Mod1",
    "Mod2",
    "Mod4",
    "Mod5"
]

ventana = tkinter.Tk()
ventana.title("Facultad de Ciencias Fisico Matematico")
ventana.config(background="#f5f6fa")
ventana.resizable(False,False)
altura = 350
anchura = 600

alt_pantalla = ventana.winfo_screenheight()
anch_pantalla = ventana.winfo_screenwidth()
x = int((anch_pantalla/2) - (anchura/2)) 
y = int((alt_pantalla/2) - (altura/2))
ventana.geometry("{}x{}+{}+{}".format(anchura, altura,x,y-100))

etiqueta = tkinter.Label(ventana, text= " Visualizador de Materias de planes de Estudio ",bg= "light grey", bd=2, relief="groove", font="Arial 10")
etiqueta.pack(fill=tkinter.X)

mostrarArchivos = tkinter.Message(ventana, bg="white", relief="groove", justify="center", width=390)
mostrarArchivos.place(height=60, width=400, x=50, y=80)

botonArchivos = tkinter.Button(ventana, text = "Elija los\narchivos", command = elegirArchivos, relief="groove", font="Arial 11", bg="#D4D4D4")
botonArchivos.place(height=60, width=100 , x=450, y=80)

etiquetaNombre = tkinter.Label(ventana,bg="#f8f9fa", text="Ingrese el nombre del Excel a crear", font= "arial 12")
etiquetaNombre.place(height=30, width=300, x=150, y=170)

botonExcel = tkinter.Button(ventana, text="Crear Excel", relief="groove", font="Arial 12", justify="center", bg="#D4D4D4")
botonExcel.place(height=50, width=110 , x=245, y=260)

nombreExcel = tkinter.Entry(ventana, bg="white",  relief="groove", font="arial 10", justify="center")
nombreExcel.place(height=40, width=350, x=100, y=205)

botonRuta = tkinter.Button(ventana, text="Ruta a\nguardar", relief="groove", font="arial 8", justify="center", bg="#D4D4D4")
botonRuta.place(height=40, width=50, x=450, y=205)


menu = ttk.Combobox(ventana,values=opciones,height=30,width=30,state="readonly")
menu.place(x=200,y=40)
menu.set("Elija una opción")

ventana.mainloop()