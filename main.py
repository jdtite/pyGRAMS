import ttkbootstrap as tkb
from ttkbootstrap.constants import *
from tkinter import Label, PhotoImage
from ttkbootstrap.dialogs import Messagebox
import importlib
import matplotlib.pyplot as plt
import os
from connection import connection
from connection import verifyLicense
import ventas
import seguimiento
import ingreso
import reposiciones
import templado
import entrega
import etiquetado

# ---> Rutas
#Carpeta principal
carpeta_principal = os.path.dirname(__file__)
#.\igrams\gui
carpeta_imagenes = os.path.join(carpeta_principal, "images")
#.\igrams\gui\images

class MainWindow(tkb.Window):
    def __init__(self, *args, **kwargs):
        super().__init__(iconphoto=None, themename='flatly', *args, **kwargs)
        global carpeta_imagenes

        # Iniciar ventana principal maximizada y vacía
        self.title("IGRAMS: Integrated Glass Registration And Management System")
        self.iconbitmap(default=os.path.join(carpeta_imagenes, "igrams.ico"))

        self.state('zoomed')

        ancho=self.winfo_screenwidth()
        alto=self.winfo_screenheight()
        
        #Crear el Notebook para contener los módulos
        self.notebookModules=tkb.Notebook(self, width=ancho, height=alto)

        self.login()
        
        # Bucle de ejecución
        self.protocol("WM_DELETE_WINDOW", self.CloseSession)
        self.mainloop()
    
    def login(self):
        self.usuario=tkb.StringVar()
        self.contrasenha=tkb.StringVar()

        # frame superior para centrar el login
        self.frmSuperior = tkb.Frame(self)
        self.frmSuperior.pack(anchor='center', expand=1)

        # Crear y ubicar el frame para contener todo este bloque.
        self.frmContenedor = tkb.Frame(self)
        self.frmContenedor.pack(padx=10, pady=10, anchor='center')

        # frame inferior para centrar el login
        self.frmInferior = tkb.Frame(self)
        self.frmInferior.pack(anchor='center', expand=1)

        # Bienvenida y Logo
        self.logo = PhotoImage(file=os.path.join(carpeta_imagenes, "igrams.png"))
        tkb.Label(self.frmContenedor,text="Welcome to IGRAMS",font=("Times 14")).pack(pady=20)
        Label(self.frmContenedor, image=self.logo, height=150).pack()

        # Campos de texto
        # Usuario
        tkb.Label(self.frmContenedor, text="Username",font=("Times 12 bold")).pack(pady=5)
        self.txtUsuario = tkb.Entry(self.frmContenedor, justify='center',textvariable=self.usuario)
        self.txtUsuario.pack()

        # Contraseña
        tkb.Label(self.frmContenedor, text="Password",font=("Times 12 bold")).pack(pady=5)
        self.txtContrasenha = tkb.Entry(self.frmContenedor, justify='center', show="*",textvariable=self.contrasenha)
        self.txtContrasenha.bind("<Return>", self.onEnter)
        self.txtContrasenha.pack()

        # Botón de envío
        self.btnLogin=tkb.Button(self.frmContenedor,text="Login",command=self.VerifySession)
        self.btnLogin.pack(pady=15)

        # Mensaje de aviso
        self.lblAviso=tkb.Label(self.frmContenedor, text="",font=("Times 10 bold"), foreground='red')
        self.lblAviso.pack(pady=5)

        #Poner el foco en la entrada de usuario
        self.txtUsuario.focus()

    def onEnter(self, event):
        self.VerifySession()

    def VerifySession(self):
        #---------------Consultar los datos del usuario--------------------------------------------
        if verifyLicense(self, self.usuario.get()):
            conn=connection(self)
            cursor=conn.cursor(dictionary=True)
            cursor.execute('SELECT id, contrasenha, modulos, nombre1, apellido1 FROM empleados WHERE usuario = %s;', (self.usuario.get(),))
            dataUser=cursor.fetchone()
            conn.commit()
            conn.close()
        
            if dataUser is None:
                self.lblAviso.config(text='Usuario no registrado.')
                self.usuario.set("")
                self.contrasenha.set("")
            else:
                if dataUser['contrasenha']==self.contrasenha.get():
                    usuario=dataUser['nombre1'] + ' ' + dataUser['apellido1']
                    id_empleado=dataUser['id']
                    modulos=dataUser['modulos']
                    self.cargar_modulos(id_empleado, usuario, modulos.split('-'))
                else:
                    self.lblAviso.config(text='Contraseña incorrecta.')
                    self.contrasenha.set("")
        else:
            Messagebox.show_error(title='Licencia no encontrada.', message=f'No se pudo encontrar una licencia válida para el usuario {self.usuario.get()}, contacte con su administrador.', parent=self)

    def cargar_modulos(self, id_empleado: int, usuario: str, modulos: list[str]):
        self.usuario=usuario
        self.id_empleado=id_empleado
        #Quitar contenedor del login y ubicar el Notebook
        self.frmSuperior.pack_forget()
        self.frmContenedor.pack_forget()
        self.frmInferior.pack_forget()
        self.notebookModules.pack(padx=10, pady=10,side='left', fill='both')

        self.title("IGRAMS: Integrated Glass Registration And Management System - " + self.usuario)
        if modulos:
            print(f'Estoy cargando los módulos:')
            for modulo in modulos:
                print(modulo)
                self.module=importlib.import_module(modulo).Main(self.notebookModules, ventana=self, register=self.register, id_empleado=self.id_empleado, usuario=self.usuario)
                self.notebookModules.add(self.module,text=modulo.capitalize())

    def CloseSession(self):
        close = Messagebox.yesno(message="¿Está seguro que desea salir de IGRAMS?", title="Confirmar cierre", parent=self)
        if close=='Yes':
            plt.close('all')
            self.quit()

#------------------------------------------------------------------------------------------
mainWindow = MainWindow()