#LIBRERIAS NECESARIAS
from tkinter import *
from tkinter import messagebox
import pyodbc

#CUERPO DE LA APP
root=Tk()
root.title("Envio Correo Cotizacion")
root.geometry("400x350")
root.resizable(False,False)
root.iconbitmap("Email.ico")

miFrame=Frame(root)
miFrame.pack()

#VARIABLES
varost=StringVar()
varano=StringVar()
varuasig=StringVar()
varestcot=StringVar()

varuser=StringVar()
varusername=StringVar()
varusermail=StringVar()

#FUNCIONES
def crearconexion():
    global conexion
    direccion_servidor = '10.120.25.80'
    nombre_bd = 'AnovoASR'
    nombre_usuario = 'ENVIRONMENT_PRD'
    password = '@env-PRD-2015$#'

    try:
        conexion = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' +
                              direccion_servidor+';DATABASE='+nombre_bd+';UID='+nombre_usuario+';PWD=' + password)
        print("Conectado")
    
    except Exception as e:
        print("Ocurrió un error al conectar a SQL Server: ", e)
    
def saliraplicacion():
    global conexion
    valor=messagebox.askquestion("Salir", "Desea salir de la aplicacion?")

    if valor=="yes":
        root.destroy()

def borrarcampos():
    varost.set("")
    varano.set("")
    varuasig.set("")
    varestcot.set("")
    varuser.set("")
    varusername.set("")
    varusermail.set("")


def enviarmsj():
    global conexion
    crearconexion()
    try:
        with conexion.cursor() as cursor:
            consulta = "exec usp_EnviarEmailCotizacion ?,?"
            cursor.execute(consulta, (varano.get(), varost.get()))
            conexion.commit()

            if varost.get() == '':
                messagebox.showwarning("Error","Ingrese una OST")

            elif cursor.rowcount > 0:
                messagebox.showinfo("Informacion","Se envio el correo OST: " + str(varost.get()) + " a el usuario" + str(varuasig.get()))
            else:
                messagebox.showwarning("Error","Error al enviar, la OST no existe o no esta En Espera")

    except Exception as e:
        print(e)
        messagebox.showwarning("Error","Error al enviar")

    finally:
        conexion.close()
        print("Conexion Cerrada")

def buscarost():
    global conexion
    crearconexion()
    try:
        with conexion.cursor() as cursor:
            consulta = "exec usp_ObtenerOSTMailCotizado ?, ?"
            cursor.execute(consulta, varost.get(), varano.get())
            datosost = cursor.fetchone()

            datosost2 = ''.join(str(e) for e in datosost)
            print(datosost2)

            varano.set(datosost[0])
            varuasig.set(datosost[1])
            varestcot.set(datosost[2])

    except Exception as e:
        print(e)
        if varano.get() == "":
            messagebox.showwarning("Error","Ingrese AÑO")
        elif varost.get() == "":
            messagebox.showwarning("Error","Ingrese una OST")
        else:
            messagebox.showwarning("Error","La OST ingresada no existe")

    finally:
        conexion.close()
        print("Conexion Cerrada")


def buscarasesor():
    global conexion
    
    crearconexion()
    try:
        with conexion.cursor() as cursor:
            consulta = "SELECT NOMBRE,CORREO_USUARIO,USUARIO FROM OPERADOR.ASESOR WHERE USUARIO = ?"
            cursor.execute(consulta,varuser.get())
            datasesor = cursor.fetchone()

            datasesor2 = ''.join(str(e) for e in datasesor)
            print(datasesor2)

            varusername.set(datasesor[0])
            varusermail.set(datasesor[1])
            varuasig.set(datasesor[2])
               
    except Exception as e:
        print(e)
        if varuser.get() == "":
            messagebox.showwarning("Error","Ingrese un Usuario")
        else:
            messagebox.showwarning("Error","El usuario ingresado no existe")

    finally:
        conexion.close()
        print("Conexion Cerrada")

def actualizarasesor():
    global conexion
    crearconexion()
    try:
        with conexion.cursor() as cursor:
            consulta = "UPDATE ORD_SERVICIO_USUARIODEV SET USUARIO_ASIGNADO = ? WHERE NUM_OST = ? AND COD_ANO = ?"
            cursor.execute(consulta, (varuasig.get(), varost.get(), varano.get()))
            conexion.commit()

            if varost.get() == "":
                messagebox.showwarning("Error","Ingrese una OST")

            elif cursor.rowcount > 0:
                messagebox.showinfo("Informacion","Se actualizo el usuario de la OST " + str(varost.get()) + " con " + str(varuasig.get()))

            else:
                messagebox.showwarning("Error","Error al actualizar, la OST no existe")

    except Exception as e:
        print(e)
        messagebox.showwarning("Error","Error al actualizar, faltan datos")

    finally:
        conexion.close()
        print("Conexion Cerrada")

def crearost():
    global conexion
    crearconexion()
    try:
        with conexion.cursor() as cursor:
            consulta = "exec usp_IngresarOSTenOS_USUARIODEV ?,?,?"
            cursor.execute(consulta, (varano.get(), varost.get(), varuasig.get()))
            conexion.commit()

            if varost.get() == '':
                messagebox.showwarning("Error","Ingrese una OST")

            elif cursor.rowcount > 0:
                messagebox.showinfo("Informacion","Se agrego la OST " + str(varost.get()) + " con el usuario" + str(varuasig.get()))

            elif varestcot != "":
                messagebox.showwarning("Error","La OST ya existe")

            else:
                messagebox.showwarning("Error","Error al crear, la OST no esta en ASR o no esta En Espera")

    except Exception as e:
        print(e)
        messagebox.showwarning("Error","Error al crear")

    finally:
        conexion.close()
        print("Conexion Cerrada")

#------------LABELS-----------#
mesagelabel=Label(miFrame, text="VER DETALLE DE ORDEN")
mesagelabel.grid(row=0, column=1, padx=5, pady=5, sticky="nswe", columnspan=3)

ostlabel=Label(miFrame, text="OST:")
ostlabel.grid(row=1, column=0, padx=5, pady=5, sticky="w")

anolabel=Label(miFrame, text="AÑO:")
anolabel.grid(row=1, column=3, padx=5, pady=5, sticky="w")

userasiglabel=Label(miFrame, text="USUARIO:")
userasiglabel.grid(row=2, column=0, padx=5, pady=5, sticky="w")

estcotlabel=Label(miFrame, text="COTIZACION:")
estcotlabel.grid(row=3, column=0, padx=5, pady=5, sticky="w")

#---------ENTRY-------
ostentry=Entry(miFrame, width=20, textvariable=varost)
ostentry.grid(row=1, column=1, padx=5, pady=5)

anoentry=Entry(miFrame, width=10, textvariable=varano)
anoentry.grid(row=1, column=4, padx=5, pady=5)

userasigtentry=Entry(miFrame, width=40, textvariable=varuasig, state='disabled')
userasigtentry.grid(row=2, column=1, padx=5, pady=5, columnspan=4)

estcottentry=Entry(miFrame, width=40, textvariable=varestcot, state='disabled')
estcottentry.grid(row=3, column=1, padx=5, pady=5, columnspan=4)

#---------BOTONES 1----------#
miFrame2=Frame(root)
miFrame2.pack()

searchbutton=Button(miFrame2, width=10, text="BUSCAR",command=buscarost)
searchbutton.grid(row=0,column=0, sticky="e", padx=5, pady=5)

createbutton=Button(miFrame2, width=10, text="CREAR", command=crearost)
createbutton.grid(row=0,column=1, sticky="e", padx=5, pady=5)

updatebutton=Button(miFrame2, width=10, text="ACTUALIZAR", command=actualizarasesor)
updatebutton.grid(row=0,column=2, sticky="e", padx=5, pady=5)

deletebutton=Button(miFrame2, width=10, text="NUEVO", command=borrarcampos)
deletebutton.grid(row=0,column=3, sticky="e", padx=5, pady=5)

sendbutton=Button(miFrame2, width=36,height=2, text="ENVIAR CORREO COTIZACION", command=enviarmsj)
sendbutton.grid(row=1,column=0, sticky="e", padx=5, pady=5, columnspan=3)

exitbutton=Button(miFrame2, width=10,height=2, text="SALIR", command=saliraplicacion)
exitbutton.grid(row=1,column=3, sticky="e", padx=5, pady=5)

#------------VER DATOS DE ASESOR-----------#
miFrame3 = Frame(root)
miFrame3.pack()

mesagelabel2=Label(miFrame3, text="CAMBIO DE USUARIO ASESOR")
mesagelabel2.grid(row=0, column=0, padx=5, pady=5, sticky="nswe", columnspan=4)

userlabel=Label(miFrame3, text="USUARIO:")
userlabel.grid(row=1, column=0, padx=5, pady=5, sticky="w")

usernamelabel=Label(miFrame3, text="NOMBRE:")
usernamelabel.grid(row=2, column=0, padx=5, pady=5, sticky="w")

useremaillabel=Label(miFrame3, text="CORREO:")
useremaillabel.grid(row=3, column=0, padx=5, pady=5, sticky="w")

#---------ENTRY ASESOR-------
searchuserbutton=Button(miFrame3, width=15, text="ASIGNAR", command=buscarasesor)
searchuserbutton.grid(row=1,column=2, sticky="e", padx=5, pady=5, columnspan=3)

userentry=Entry(miFrame3, width=30, textvariable=varuser)
userentry.grid(row=1, column=1, padx=5, pady=5)

usernameentry=Entry(miFrame3, width=50, textvariable=varusername, state='disabled')
usernameentry.grid(row=2, column=1, padx=5, pady=5, columnspan=3)

useremailentry=Entry(miFrame3, width=50, textvariable=varusermail, state='disabled')
useremailentry.grid(row=3, column=1, padx=5, pady=5, columnspan=3)



root.mainloop()