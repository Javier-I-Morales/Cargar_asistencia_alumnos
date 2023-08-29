import tkinter as tk
from tkinter import filedialog as fd
from tkinter import messagebox as mg
from Cargar_asistencia_alumnos.Cargador_Asistencia import cargador
from threading import Thread


raiz = tk.Tk()
raiz.title("Cargador Asistencia")
raiz.resizable(False,False)

cargador = cargador()

def proceso():
    cargador.cargar_asistencias()

def archivo_origen():
    cargador.ruta_archivo = fd.askopenfilename(title='Seleccione archivo')
    ruta = str(cargador.ruta_archivo).split('/')
    cargador.ruta_save = "".join(ruta[i]+'/' for i in range(len(ruta)-1))


def acerca():
    mg.showinfo(title='Cargador de asistencias.', message='Aplicacion realizada por alumnos asistentes para la UNaHur.')

def ayuda():
    mg.showinfo(title='Ayuda', message='En la carpeta de la aplicación se encuentra un manual.')

def tomar_datos_ejecutar():
    hilo = Thread(target=proceso)
    cargador.desde = dato_desde_comision.get()
    cargador.hasta = dato_hasta_comision.get()
    cargador.usuario = dato_usuario.get()
    cargador.passw = dato_password.get()
    cargador.fecha = entryfecha.get()
    raiz.update_idletasks()
    hilo.start()
    hilo.join()

def delete_window():
    close = mg.askyesno(
        message="¿Está seguro de que quiere cerrar la aplicación?",
        title="Confirmar cierre"
    )
    if close:
        raiz.destroy()




################# DESDE AQUÍ LA GRAFICA


raiz.iconbitmap('F:\\Python\\Cargar_asistencia_alumnos\\image\\icono.ico')

frame = tk.Frame(raiz, bg="#424242", width=600, height=400)
frame.pack()

barra_menu = tk.Menu(raiz)
raiz.config(menu=barra_menu)

boton_menu=tk.Menu(barra_menu, tearoff=0, bg='#E9EAEE')
boton_ayuda=tk.Menu(barra_menu, tearoff=0, bg='#E9EAEE')
boton_menu.add_command(label='Seleccione archivo', command=archivo_origen)
boton_ayuda.add_command(label='Ayuda', command=ayuda)
boton_ayuda.add_command(label='Acerca', command=acerca)
barra_menu.add_cascade(label='Menu', menu=boton_menu)
barra_menu.add_cascade(label='Acerca', menu=boton_ayuda)

frame_imagen = tk.Frame(frame, bg="#ADC98F",width=220,height=395).place(x=0,y=5)
imagen = tk.PhotoImage(file="F:\\Python\\Cargar_asistencia_alumnos\\image\\logo.png")
label_imagen = tk.Label(frame_imagen, image=imagen)
label_imagen.place(x=40,y=100)



frame_elementos = tk.Frame(frame, bg="green",width=380,height=395).place(x=220,y=5)
# frame_elementos.grid_propagate(False)


label_usuario = tk.Label(frame_elementos, text='Usuario', font=(20), bg="green", fg='#FFF')
label_usuario.place(x=250, y=50)
# label_usuario.grid(row=1, column=4, padx=0, sticky='e')
#
dato_usuario = tk.Entry(frame_elementos)
dato_usuario.place(x=400, y=50)
# dato_usuario.grid(row=1, column=5)
#
label_password = tk.Label(frame_elementos, text='Contraseña', font=(20), bg="green", fg='#FFF')
label_password.place(x=250, y=100)
# label_password.grid(row=2, column=4, padx=0, sticky='e' )
#
dato_password = tk.Entry(frame_elementos)
dato_password.place(x=400,y=100)
# dato_password.grid(row=2, column=5)
#
# ##############
#
label_desde_comision = tk.Label(frame_elementos, text='Comisión inicial', font=(20), bg="green", fg='#FFF')
label_desde_comision.place(x=250, y=150)
# label_desde_comision.grid(row=3, column=4, padx=0, sticky='e')
#
dato_desde_comision = tk.Entry(frame_elementos)
dato_desde_comision.place(x= 400,y=150)
# dato_desde_comision.grid(row=3, column=5)
#
label_hasta_comision = tk.Label(frame_elementos, text='Comisión final', font=(20), bg="green", fg='#FFF')
label_hasta_comision.place(x=250,y=200)
# label_hasta_comision.grid(row=4, column=4, padx=0, sticky='e')
#
dato_hasta_comision  = tk.Entry(frame_elementos)
dato_hasta_comision.place(x=400,y=200)
# dato_hasta_comision .grid(row=4, column=5)
#
label_fecha = tk.Label(frame_elementos, text='Fecha dd/mm/aaaa', font=(20), bg="green", fg='#FFF')
label_fecha.place(x=250, y=250)
# label_cantidad_archivos_titulo.grid(row=5, column=4, padx=0, sticky='e')
#
entryfecha = tk.Entry(frame_elementos)
entryfecha.place(x=400,y=250)
# entryfecha.grid(row=5, column=5)
#


boton = tk.Button
boton(frame_elementos, text = "Ejecutar", command = tomar_datos_ejecutar, borderwidth=0 ,bg='#424242', fg='#FFF', width=15).place(x=360, y=320)

# cerrar = tk.Button
# cerrar(frame_elementos, text = "Cerrar", command = cerrar_proceso , borderwidth=0 ,bg='#303030', fg='#FFF', width=15).place(x=360, y=380)


raiz.protocol("WM_DELETE_WINDOW", delete_window)



raiz.mainloop()