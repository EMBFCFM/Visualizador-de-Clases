from funciones import *
import tkinter
import tkinter.filedialog as fd
from tkinter import ttk, messagebox
from os import path



def elegirArchivos():
    #global listaArchivos 
    global archivos
    global filename
    #listaArchivos = []
    if menu.get() == "Alumnos con escuela de procedencia" or menu.get() == "Listas de Alumnos" or menu.get() == "Listas con promedios" or menu.get() == "Directorio de Alumnos con promedio" or menu.get() == "Horarios de Alumnos con promedio por Plan/Modalidad":
        archivos = fd.askopenfilename(parent= ventana, title = "Elija los archivos")
        name = path.basename(archivos)
        #name = name.removesuffix('.txt')
        mostrarArchivos["text"] = name
        if menu.get() == "Horarios de Alumnos con promedio por Plan/Modalidad":
                    changestatetxt()
        return archivos
    else:
        textoArchivos = ""
        archivos = fd.askopenfilenames(parent= ventana, title = "Elija los archivos")
        cont=1
        for archivo in archivos:
            nombreArch = archivo.split("/")
            #listaArchivos.append(str(nombreArch[-1]))
            textoArchivos = textoArchivos + str(nombreArch[-1])
            if cont < len(archivos):
                textoArchivos = textoArchivos + " / "
                cont+=1
        mostrarArchivos["text"] = str(textoArchivos)
        return archivos

def rutaGuardar():
    global rutaGuardado
    rutaGuardado = fd.askdirectory()
    print (rutaGuardado)

def takeinput():
    op = ingresarOportunidad.get(1.0, "end-1c")
    return op

def takeinput2():
    clave = ingresarClave.get(1.0, "end-1c")
    return clave

def changestatetxt():
    ingresarOportunidad['state'] = 'normal'
    ingresarClave['state'] = 'normal'

def parametrosExcel():
    try:
        print("archivos:", archivos)
        if archivos == "":
            messagebox.showinfo(message = "Por favor, elija los archivos.", title = "Aviso")
            return
    except:
        messagebox.showinfo(message = "Por favor, elija los archivos.", title = "Aviso")
        return
    try:
        print("ruta", rutaGuardado)
        if rutaGuardado == "":
            messagebox.showinfo(message = "Por favor, elija la ruta a guardar.", title = "Aviso")
            return
    except:
        messagebox.showinfo(message = "Por favor, elija la ruta a guardar.", title = "Aviso")
        return
    
    if str(nombreExcel.get()) == "":
        messagebox.showinfo(message = "Por favor, ingrese un nombre para el archivo.", title = "Aviso")
        return
    else:
        nombreArchivo= str(rutaGuardado)+"/"+str(nombreExcel.get()) + ".xlsx"
        if path.exists(nombreArchivo) == True:
            messagebox.showinfo(message = "Ya existe un archivo con ese nombre, ingrese uno distinto.", title = "Aviso")
            return
        else:
            if menu.get() == "Elija una opción":
                messagebox.showinfo(message="Por favor, elija una opción válida.")
            elif menu.get() == "Revisar Minutas":
                llamarFuncionesMinutas(archivos, rutaGuardado, nombreExcel.get())
            elif menu.get() == "Correos de alumnos":
                llamarFuncionesCorreo(archivos, rutaGuardado, nombreExcel.get())
            elif menu.get() == "Alumnos Reprobados":
                parametrosExcelReprobados(archivos, rutaGuardado, nombreExcel.get())
            elif menu.get() == "Alumnos con escuela de procedencia" or menu.get() == "Listas de Alumnos" or menu.get() == "Listas con promedios" or menu.get() == "Directorio de Alumnos con promedio" or menu.get() == "Horarios de Alumnos con promedio por Plan/Modalidad":
                reportesAlumnos(archivos,menu.get(),ingresarOportunidad.get(),ingresarClave.get())

ventana = tkinter.Tk()
ventana.title("Universidad Autónoma de Nuevo León")
ventana.config(background="#f8f9fa")
ventana.resizable(False,False)
altura = 350
anchura = 600
alt_pantalla = ventana.winfo_screenheight()
anch_pantalla = ventana.winfo_screenwidth()
x = int((anch_pantalla/2) - (anchura/2))
y = int((alt_pantalla/2) - (altura/2))
ventana.geometry("{}x{}+{}+{}".format(anchura, altura,x,y-100))


etiqueta = tkinter.Label(ventana, text = "Automatizador de archivos Excel", bg = "light grey", bd=2, relief="groove", font="Arial 10")
etiqueta.pack(fill=tkinter.X)

# nomenclatura = tkinter.Label(ventana, bg="#f8f9fa", text="Nomenclatura recomendada: Cuartas 401.txt, Segundas 420.txt, etc.", font="arial 8")
# nomenclatura.place(height=20, width=350, x=125, y=60)

mostrarArchivos = tkinter.Message(ventana, bg="white", relief="groove", justify="center", width=390)
mostrarArchivos.place(height=60, width=400, x=50, y=80)

ingresarOportunidad = tkinter.Entry(ventana, bg="white",  relief="groove", font="arial 10", justify="center",state="disabled")
ingresarOportunidad.place(height=20, width=30,x=250,y=150)

ingresarClave = tkinter.Entry(ventana, bg="white",  relief="groove", font="arial 10", justify="center", state="disabled")
ingresarClave.place(height=20, width=30,x=300,y=150)

botonArchivos = tkinter.Button(ventana, text = "Elija los\narchivos", command = elegirArchivos, relief="groove", font="Arial 11", bg="#D4D4D4")
botonArchivos.place(height=60, width=100 , x=450, y=80)

etiquetaNombre = tkinter.Label(ventana,bg="#f8f9fa", text="Ingrese el nombre del Excel a crear", font= "arial 12")
etiquetaNombre.place(height=30, width=300, x=150, y=170)

nombreExcel = tkinter.Entry(ventana, bg="white",  relief="groove", font="arial 10", justify="center")
nombreExcel.place(height=40, width=350, x=100, y=205)

botonExcel = tkinter.Button(ventana, text="Crear Excel", relief="groove", font="Arial 12", justify="center", bg="#D4D4D4", command=parametrosExcel)
botonExcel.place(height=50, width=110 , x=245, y=260)

botonRuta = tkinter.Button(ventana, text="Ruta a\nguardar", relief="groove", font="arial 8", justify="center", bg="#D4D4D4", command=rutaGuardar)
botonRuta.place(height=40, width=50, x=450, y=205)

options = [
    "Correos de alumnos",
    "Revisar Minutas",
    "Alumnos Reprobados",
    "Alumnos con escuela de procedencia",
    "Listas de Alumnos",
    "Listas con promedios",
    "Directorio de Alumnos con promedio",
    "Horarios de Alumnos con promedio por Plan/Modalidad"
]

menu = ttk.Combobox(ventana, values=options, height=30, width=30, state="readonly")
menu.place(x=200, y=40)
menu.set("Elija una opción")

# botonAyuda = tkinter.Button(ventana, text="Ayuda", relief="groove", font="arial 8", justify="center", bg="#D4D4D4", command=ayuda)
# botonAyuda.place(height=40, width=50, x=530, y=290)
ventana.mainloop()