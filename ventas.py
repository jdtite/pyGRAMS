import ttkbootstrap as tkb
from tkinter import  messagebox
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
from datetime import datetime, timedelta
import os
import numpy as np
import create_document
from decimal import Decimal
from connection import connection

#Carpeta principal
carpeta_principal = os.path.dirname(__file__)
#.\igams\gui\documentos
carpeta_documentos = os.path.join(carpeta_principal, "documentos")
#.\igrams\gui\images
carpeta_imagenes = os.path.join(carpeta_principal, "images")

class Main(tkb.Frame):
    def __init__(self, *args, ventana:tkb.Window, register:tkb.Window.register, id_empleado:int, usuario:str, **kwargs):
        super().__init__(*args, **kwargs)

        #instanciar cursor
        self.ventana=ventana
        self.registro=register

        self.usuario=usuario
        self.id_empleado=id_empleado

        #Define el Frame como contenedor total que se asocia al notebook
        self.contenedor=ScrolledFrame(self)
        self.contenedor.pack(padx=10,pady=10, fill='both', expand=True)
        #Define la primera columna del contenedor
        self.colum1=tkb.Frame(self.contenedor)
        self.colum1.pack(padx=10, pady=10, side='left', fill='y')

        #-------------------Crear y ubicar el Frame para Ver la lista de planos creados---------------------------
        self.frmPlanosCreados=tkb.LabelFrame(self.colum1,text='Mis Planos de Producción', bootstyle="primary")
        self.frmPlanosCreados.pack(padx=5, pady=5, ipadx=10, ipady=10)
        # Crear y ubicar la entrada de fecha de inicio para la búsqueda
        tkb.Label(self.frmPlanosCreados,text="Desde:",font=("Arial 10 bold")).grid(row=0,column=0, padx=5,pady=3,sticky='e')
        self.dtDesde=tkb.DateEntry(self.frmPlanosCreados, dateformat="%Y-%m-%d", width=12)
        self.dtDesde.grid(row=0,column=1,pady=3,sticky='w')
        # Crear y ubicar la entrada de fecha de fin para la búsqueda
        tkb.Label(self.frmPlanosCreados,text="Hasta:",font=("Arial 10 bold")).grid(row=0,column=2, padx=5,pady=3,sticky='e')
        self.dtHasta=tkb.DateEntry(self.frmPlanosCreados, dateformat="%Y-%m-%d", width=12)
        self.dtHasta.grid(row=0,column=3,pady=3,sticky='w')
        #Crear y ubicar el botón para actualizar los valores
        self.btnActualizar=tkb.Button(self.frmPlanosCreados,text="Actualizar",command=self.actualizar)
        self.btnActualizar.grid(row=0, column=4, padx=5, pady=5)

        #Crear y ubicar el marco para contener la tabla y la barra de desplazamiento
        self.frmListaPlanos = tkb.Frame(self.frmPlanosCreados, relief="flat")
        self.frmListaPlanos.grid(row=1, column=0, columnspan=5, sticky='we')
        self.frmListaPlanos.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmListaPlanos.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmListaPlanos, orient=VERTICAL)
        scrollbarV.grid(row=0,column=1, pady=5, sticky='ns')

        fuente=tkb.font.Font(family="times", size=11)
        style = tkb.Style()
        style.configure("mystyle.Treeview", font=('times', 11))

        #Crear y ubicar la tabla de los planos creados
        self.tablePlanosCreados=tkb.Treeview(self.frmListaPlanos, show='headings', height=10, yscrollcommand=scrollbarV.set, style="mystyle.Treeview")
        self.tablePlanosCreados.configure(columns=('plano', 'fecha-hora', 'materia-prima', 'cliente'))
        self.tablePlanosCreados.column('plano', stretch=False, anchor='center', width=fuente.measure(' 0000000 '))
        self.tablePlanosCreados.column('fecha-hora', stretch=False, anchor='w', width=fuente.measure('0 0000-00-00 00:00:00 0'))
        self.tablePlanosCreados.column('materia-prima', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxx'))
        self.tablePlanosCreados.column('cliente', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxxxxxxxx'))

        for col in self.tablePlanosCreados['columns']:
            self.tablePlanosCreados.heading(col, text=col.title(), anchor=W)

        self.tablePlanosCreados.grid(row=0,column=0, padx=5, pady=5)
        self.tablePlanosCreados.tag_configure('listo', background='plum')
        self.tablePlanosCreados.bind("<<TreeviewSelect>>", self.lboxPlanosCreados_change)
        self.tablePlanosCreados.bind("<Double-Button-1>", self.verPortada)

        scrollbarV.config(command=self.tablePlanosCreados.yview)

        #Crear y ubicar el marco para contener el resumen de la información
        self.frmResumenInfo = tkb.Frame(self.frmPlanosCreados, relief="flat")
        self.frmResumenInfo.grid(row=2, column=0, columnspan=3, sticky='we')
        #Crear y ubicar la etiqueta para la cantidad de vidrios repuestos
        tkb.Label(self.frmResumenInfo,text="Planos creados:",font=("Arial 10 bold"), justify="left").grid(row=0,column=0,padx=5,sticky='w')
        self.lblPlanosCreados=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblPlanosCreados.grid(row=0,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de área original
        tkb.Label(self.frmResumenInfo,text="Planos ingresados:",font=("Arial 10 bold"), justify="left").grid(row=1,column=0,padx=5,sticky='w')
        self.lblPlanosIngresados=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblPlanosIngresados.grid(row=1,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de área recuperada
        tkb.Label(self.frmResumenInfo,text="Área ingresada:",font=("Arial 10 bold"), justify="left").grid(row=2,column=0,padx=5,sticky='w')
        self.lblAreaIngresada=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblAreaIngresada.grid(row=2,column=1,sticky='w')

        #Crear y ubicar el marco para contener los botones.
        self.frmBotonesLista = tkb.Frame(self.frmPlanosCreados)
        self.frmBotonesLista.grid(row=2, column=3, columnspan=2, padx=5, pady=5)
        #Crear y ubicar el botón para editar el plano.
        self.btnEditarPlano=tkb.Button(self.frmBotonesLista,text="Editar",command=self.editarPlano, state='disabled')
        self.btnEditarPlano.pack(padx=5, pady=5, side='left')
        #Crear el botón para Crear un nuevo plano modificaciones en un plano.
        self.btnNuevo=tkb.Button(self.frmBotonesLista,text="Nuevo", width=10,command=self.nuevo_plano, bootstyle="success")
        self.btnNuevo.pack(padx=5, pady=5, side='left')

        #-------------------Crear y ubicar el Frame para ver la lista de entrega por día---------------------------
        #Define la primera columna del contenedor
        self.colum2=tkb.Frame(self.contenedor)
        self.colum2.pack(padx=10, pady=10, side='left', fill='y')

        self.frmPlanosDiarios=tkb.LabelFrame(self.colum2,text='Lista de entrega', bootstyle="primary")
        self.frmPlanosDiarios.pack(padx=5, pady=5, ipadx=10, ipady=10)
        # Crear y ubicar la entrada de fecha de inicio para la búsqueda
        tkb.Label(self.frmPlanosDiarios,text="Fecha:",font=("Arial 10 bold")).grid(row=0,column=0, padx=5,pady=3,sticky='e')
        self.dtFechaEntrega=tkb.DateEntry(self.frmPlanosDiarios, dateformat="%Y-%m-%d", width=12)
        self.dtFechaEntrega.grid(row=0,column=1,pady=3,sticky='w')
        #Crear y ubicar el botón para actualizar los valores
        self.btnActualizarPDiarios=tkb.Button(self.frmPlanosDiarios,text="Actualizar",command=self.consultar_planos_diarios)
        self.btnActualizarPDiarios.grid(row=0, column=2, padx=5, pady=5, sticky='w')

        #Crear y ubicar el marco para contener el listbox y la barra de desplazamiento
        self.frmListaPDiarios = tkb.Frame(self.frmPlanosDiarios, relief="flat")
        self.frmListaPDiarios.grid(row=1, column=0, columnspan=3, sticky='we')
        self.frmListaPDiarios.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmListaPDiarios.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmListaPDiarios, orient=VERTICAL)
        scrollbarV.grid(row=0,column=1, pady=5, sticky='ns')
        #Crear y ubicar la tabla de los planos a entregar
        self.tablePlanosEntrega=tkb.Treeview(self.frmListaPDiarios, show='headings', height=10, yscrollcommand=scrollbarV.set, style="mystyle.Treeview")
        self.tablePlanosEntrega.configure(columns=('plano', 'estado', 'materia-prima', 'cliente'))
        self.tablePlanosEntrega.column('plano', stretch=False, anchor='center', width=fuente.measure(' 0000000 '))
        self.tablePlanosEntrega.column('estado', stretch=False, anchor='center', width=fuente.measure(' 000/000 '))
        self.tablePlanosEntrega.column('materia-prima', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxx'))
        self.tablePlanosEntrega.column('cliente', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'))

        for col in self.tablePlanosEntrega['columns']:
            self.tablePlanosEntrega.heading(col, text=col.title(), anchor=W)

        self.tablePlanosEntrega.grid(row=0,column=0, padx=5, pady=5)
        self.tablePlanosEntrega.tag_configure('listo', background='sky blue')
        self.tablePlanosEntrega.bind("<<TreeviewSelect>>", self.tablePlanosEntrega_change)

        scrollbarV.config(command=self.tablePlanosEntrega.yview)

        #Crear y ubicar el marco para contener el resumen de la información
        self.frmResumenInfoPDiarios = tkb.Frame(self.frmPlanosDiarios, relief="flat")
        self.frmResumenInfoPDiarios.grid(row=4, column=0, columnspan=2, sticky='we')
        #Crear y ubicar la etiqueta para la cantidad de área templada
        tkb.Label(self.frmResumenInfoPDiarios,text="Planos a entregar:",font=("Arial 10 bold"), justify="left").grid(row=0,column=0,padx=5,sticky='w')
        self.lblPlanosDiarios=tkb.Label(self.frmResumenInfoPDiarios,text="",font=("Arial 10"))
        self.lblPlanosDiarios.grid(row=0,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de área templada.
        tkb.Label(self.frmResumenInfoPDiarios,text="Área total:",font=("Arial 10 bold"), justify="left").grid(row=1,column=0,padx=5,sticky='w')
        self.lblAreaDiaria=tkb.Label(self.frmResumenInfoPDiarios,text="",font=("Arial 10"))
        self.lblAreaDiaria.grid(row=1,column=1,sticky='w')

    def lboxPlanosCreados_change(self, Event):
        if self.tablePlanosCreados.selection():
            self.btnEditarPlano.config(state='active',bootstyle="info")
        else:
            self.btnEditarPlano.config(state='disabled',bootstyle="disabled")

    def actualizar(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT id, marca_creacion, cliente, id_materia_prima, color, espesor, marca_etiquetado FROM planos WHERE id_empleado=%s AND marca_creacion BETWEEN %s AND %s;",(self.id_empleado, self.dtDesde.entry.get() + " 00:00:00",self.dtHasta.entry.get() + " 23:59:59"))
        datosPlanos=cursor.fetchall()
        conn.commit()

        fila=[]
        self.tablePlanosCreados.delete(*self.tablePlanosCreados.get_children())
        ingresados=0
        area=Decimal(0.00)
        for plano in datosPlanos:
            fila.clear()
            fila.append(str(plano['id']))
            fila.append(str(plano['marca_creacion']))

            if plano['cliente'] is not None:
                if plano['id_materia_prima']:
                    cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(plano['id_materia_prima'],))
                    materia_prima=cursor.fetchone()
                    conn.commit()
                    fila.append(materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')            
                else:
                    fila.append(str(plano['color']) + ' ' + str(plano['espesor']) + ' mm')
                fila.append(str(plano['cliente']))

            if plano['marca_etiquetado']:
                self.tablePlanosCreados.insert('', 0, values=fila,tags=('listo',))
                ingresados=ingresados+1
                #Consultar todos los vidrios ingresados para este plano
                cursor.execute("SELECT SUM(area) FROM vidrios WHERE id_plano=%s;",(plano['id'],))
                aux=cursor.fetchone()
                conn.commit()

                area=aux['SUM(area)']
            
            else:
                self.tablePlanosCreados.insert('', 0, values=fila)

        self.lblPlanosCreados.config(text=str(len(datosPlanos)))
        self.lblPlanosIngresados.config(text=str(ingresados))
        self.lblAreaIngresada.config(text=str(area) + ' m²')
        conn.close()

    def nuevo_plano(self):
        crear = messagebox.askyesno(message="¿Está seguro que desea crear un nuevo plano?", title="Confirmar creación")
        if crear:
            #----------------Crear la instancia en la base de datos---------------
            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            cursor.execute("START TRANSACTION;")
            cursor.execute("INSERT INTO planos (id_empleado, marca_creacion) VALUES (%s, %s);", (self.id_empleado, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            cursor.execute("SELECT LAST_INSERT_ID();")
            aux=cursor.fetchone()
            self.id_plano=aux['LAST_INSERT_ID()']
            cursor.execute("COMMIT;")
            conn.close()
            #Abrir la ventana secundaria con el formulario------------------------
            FormPortada(ventana=self.ventana,register=self.registro, id_plano=self.id_plano, id_empleado=self.id_empleado, usuario=self.usuario, callback=self.actualizar)

    def editarPlano(self):
        indice=self.tablePlanosCreados.selection()[0]
        self.id_plano=self.tablePlanosCreados.item(indice,'values')[0]

        #Consultar si el plano ha empezado el proceso
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT marca_etiquetado FROM planos WHERE id = %s;",(self.id_plano,))
        plano=cursor.fetchone()
        conn.commit()
        conn.close()

        if plano['marca_etiquetado'] is None:
            FormPortada(ventana=self.ventana, register=self.registro, id_plano=self.id_plano, id_empleado=self.id_empleado, usuario=self.usuario, callback=self.actualizar)
        else:
            messagebox.showerror(title="Edición no permitida", message="Este plano ya fue entregado a producción, póngase en contacto con un supervisor para evaluar si es posible la edición del plano.")

    def verPortada(self, Event):
        indice=self.tablePlanosCreados.selection()[0]
        id_plano=self.tablePlanosCreados.item(indice,'values')[0]

        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT planos.*, empleados.nombre1, empleados.apellido1 FROM planos INNER JOIN empleados ON planos.id_empleado = empleados.id WHERE planos.id = %s;",(id_plano,))
        datosPlano=cursor.fetchone()
        conn.commit()

        if datosPlano['id_materia_prima']:
            cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(datosPlano['id_materia_prima'],))
            materia_prima=cursor.fetchone()
            conn.commit()
            materia=materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm'

        else:
            materia=datosPlano['color'] + ' ' + str(datosPlano['espesor']) + ' mm'
        conn.close()
        create_document.Cover(datosPlano=datosPlano,materia=materia)

    def consultar_planos_diarios(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT id, cantidad_vidrios, cliente, id_materia_prima, color, espesor, marca_etiquetado FROM planos WHERE marca_etiquetado IS NOT NULL AND id_empleado = %s AND marca_plazo BETWEEN %s AND %s;",(self.id_empleado, self.dtFechaEntrega.entry.get() + " 00:00:00",self.dtFechaEntrega.entry.get() + " 23:59:59"))
        self.PlanosDiarios=cursor.fetchall()
        conn.commit()

        fila=[]
        i=0
        area=Decimal(0.00)
        
        self.tablePlanosEntrega.delete(*self.tablePlanosEntrega.get_children())

        for plano in self.PlanosDiarios:
            fila.clear()
            fila.append(str(plano['id']))
            i+=1

            cursor.execute("SELECT SUM(vidrios.area), COUNT(IF(vidrios.marca_almacenado IS NOT NULL, 1, NULL)) FROM vidrios INNER JOIN planos WHERE vidrios.id_plano = planos.id AND planos.marca_plazo BETWEEN %s AND %s;",(self.dtDesdePEntrega.entry.get() + " 00:00:00", self.dtHastaPEntrega.entry.get() + " 23:59:59"))
            aux=cursor.fetchone()
            conn.commit()

            area_total=aux['SUM(vidrios.area)']
            vidrios_templados=aux['COUNT(IF(vidrios.marca_almacenado IS NOT NULL, 1, NULL))']

            fila.append(str(vidrios_templados) + '/' + str(plano['cantidad_vidrios']))
            
            #Consultar materia prima
            if plano['id_materia_prima']:
                cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(plano['id_materia_prima'],))
                materia_prima=cursor.fetchone()
                conn.commit()
                fila.append(materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')
            else:
                fila.append(str(plano['color']) + ' ' + str(plano['espesor']) + ' mm')
            
            fila.append(str(plano['cliente']))

            if vidrios_templados==plano['cantidad_vidrios']:
                self.tablePlanosEntrega.insert('', END, values=fila,tags=('listo',))
            else:
                self.tablePlanosEntrega.insert('', END, values=fila)
        
        self.lblPlanosDiarios.config(text=str(i))
        self.lblAreaDiaria.config(text=str(area_total) + ' m²')

    def tablePlanosEntrega_change(self, Event):
        print(self.tablePlanosEntrega.item(self.tablePlanosEntrega.selection()[0],'values'))
        print(self.tablePlanosEntrega.selection()[0])

    def graficar_planos_diarios(self):
        print('estoy graficando la entrega diaria')

class FormPortada(tkb.Toplevel):   
    def __init__(self, *args, ventana:tkb.Window, register:tkb.Window.register, id_empleado:int, usuario:str, id_plano:int, callback, **kwargs):
        super().__init__(*args, **kwargs)
        #instanciar cursor
        self.ventana=ventana
        self.registro=register

        self.usuario=usuario
        self.id_empleado=id_empleado
        self.id_plano=id_plano
        self.callback=callback

        self.title('Formulario de registro de Planos')
        self.resizable(False,False)
        self.attributes('-topmost', 1)
        #self.attributes('-toolwindow', True)
        self.iconbitmap(os.path.join(carpeta_imagenes, "igrams.ico"))

        # Crear y ubicar el frame para contener todo este bloque.
        self.frmContenedor = tkb.Frame(self)
        self.frmContenedor.pack()

        self.frmCrearPlano=tkb.LabelFrame(self.frmContenedor,text=f'Plano No. {self.id_plano}',bootstyle="primary")
        self.frmCrearPlano.pack(padx=5, pady=5, ipadx=10, ipady=10, fill='x')
        #Crear y ubicar la etiqueta para el Cliente
        tkb.Label(self.frmCrearPlano,text="Cliente:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=0,column=0,padx=5,pady=3,sticky='we')
        self.txtCliente=tkb.Entry(self.frmCrearPlano, font=("Arial 10"))
        self.txtCliente.grid(row=0,column=1,columnspan=3,pady=5,sticky='we')
        #Crear y ubicar la entrada para el Destino
        tkb.Label(self.frmCrearPlano,text="Destino:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=1,column=0,padx=5,pady=3,sticky='we')
        self.txtDestino=tkb.Entry(self.frmCrearPlano, font=("Arial 10"))
        self.txtDestino.grid(row=1,column=1,columnspan=3,pady=3,sticky='we')
        self.txtDestino.insert(1,'Guayaquil')

        #Crear y ubicar la etiqueta para la materia prima, y llenar sus campos
        tkb.Label(self.frmCrearPlano,text="Materia:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=2,column=0,padx=5,pady=3,sticky='we')
        self.cboxMateriaPrima=tkb.Combobox(self.frmCrearPlano,font=("Arial 10"))
        self.cboxMateriaPrima.grid(row=2,column=1,columnspan=2,pady=3,sticky='we')

        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute('SELECT * FROM materia_prima;')
        MateriaPrima=cursor.fetchall()
        conn.commit()
        conn.close()

        self.dictMateriaPrima={}
        for i in MateriaPrima:
            self.dictMateriaPrima[str(i['id'])]=i['id_color'] + ' ' + str(i['id_espesor']) + ' mm'
        self.dictMateriaPrimaInvert={}
        for i in MateriaPrima:
            self.dictMateriaPrimaInvert[i['id_color'] + ' ' + str(i['id_espesor']) + ' mm']=i['id']
        listAux=list(self.dictMateriaPrimaInvert.keys())
        listAux.sort()
        self.cboxMateriaPrima.config(values=listAux)
        #Crear y ubicar el botón para la otra materia prima
        self.bandera=True
        self.btnOtraMateria=tkb.Button(self.frmCrearPlano,text="Otra", command=self.otra_materia_prima)
        self.btnOtraMateria.grid(row=2,column=3,padx=5,pady=3,sticky='we')
        #Crear la etiqueta y la entrada para el color personalizado-----------------------------
        self.lblColor=tkb.Label(self.frmCrearPlano,text="Color:",font=("Arial 10 bold"),anchor="w", justify="left")
        self.txtColor=tkb.Entry(self.frmCrearPlano, width=15, font=("Arial 10"))
        #Crear la etiqueta y la entrada para el espesor personalizado-----------------------------
        self.lblEspesor=tkb.Label(self.frmCrearPlano,text="Espesor:",font=("Arial 10 bold"),anchor="w", justify="left")
        self.txtEspesor=tkb.Spinbox(self.frmCrearPlano, width=4, font=("Arial 10"),from_=1,to=1000,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"))

        #Crear y ubicar la etiqueta para la cantidad de vidrios
        tkb.Label(self.frmCrearPlano,text="Cantidad:",font=("Arial 10 bold"), justify="left").grid(row=4,column=0, padx=5,pady=3,sticky='w')
        self.spboxCantidad=tkb.Spinbox(self.frmCrearPlano, width=4, font=("Arial 10"),from_=1,to=1000,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"))
        self.spboxCantidad.grid(row=4,column=1,pady=3,sticky='w')
        self.spboxCantidad.set(1)
        #Crear y ubicar la etiqueta para el área de los vidrios
        tkb.Label(self.frmCrearPlano,text="Área:",font=("Arial 10 bold"),anchor='e',width=6).grid(row=4,column=2, padx=5,pady=3,sticky='e')
        self.txtArea=tkb.Entry(self.frmCrearPlano, width=10, font=("Arial 10"),validate='key')
        self.txtArea.grid(row=4,column=3,pady=3,sticky='w')
        
        #Crear y ubicar la entrada para el Tipo de Templado
        tkb.Label(self.frmCrearPlano,text="T. Temp.:",font=("Arial 10 bold"),anchor='w').grid(row=5,column=0,padx=3,pady=3,sticky='w')
        self.cboxTT=tkb.Combobox(self.frmCrearPlano,font=("Arial 10"),width=5,values=['Plano','Curvo','NO'])
        self.cboxTT.grid(row=5,column=1,pady=3,sticky='w')
        self.cboxTT.current(0)
        #Crear y ubicar la entrada para la ubicación del sello
        tkb.Label(self.frmCrearPlano,text="Sello:",font=("Arial 10 bold"),anchor='e').grid(row=5,column=2,padx=3,pady=3,sticky='e')
        self.cboxSello=tkb.Combobox(self.frmCrearPlano,font=("Arial 10"),width=8,values=['Visto','Canto','Ver Plano'])
        self.cboxSello.grid(row=5,column=3,pady=3,sticky='w')
        self.cboxSello.current(0)

        #Crear y ubicar la etiqueta para el Comentario 1
        tkb.Label(self.frmCrearPlano,text="Nota 1:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=6,column=0,padx=5,pady=3,sticky='we')
        self.txtNota1=tkb.Entry(self.frmCrearPlano, font=("Arial 10"))
        self.txtNota1.grid(row=6,column=1,columnspan=3,pady=5,sticky='we')
        #Crear y ubicar la etiqueta para el Comentario 2
        tkb.Label(self.frmCrearPlano,text="Nota 2:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=7,column=0,padx=5,pady=3,sticky='we')
        self.txtNota2=tkb.Entry(self.frmCrearPlano, font=("Arial 10"))
        self.txtNota2.grid(row=7,column=1,columnspan=3,pady=5,sticky='we')
        #Crear y ubicar la etiqueta para el Comentario 3
        tkb.Label(self.frmCrearPlano,text="Nota 3:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=8,column=0,padx=5,pady=3,sticky='we')
        self.txtNota3=tkb.Entry(self.frmCrearPlano, font=("Arial 10"))
        self.txtNota3.grid(row=8,column=1,columnspan=3,pady=5,sticky='we')

        #-------------------------------------------------------------------------------------------------------------
        #Crear un separador
        tkb.Separator(self.frmCrearPlano,orient='vertical').grid(row=0,column=4,padx=15,pady=5,rowspan=9,sticky='ns')
        
        self.requerimientos={'Moldes':tkb.BooleanVar(value=False),'Arenados':tkb.BooleanVar(value=False)}
        #Crear y ubicar el checkbutton para los Moldes
        self.cbtnMoldes=tkb.Checkbutton(self.frmCrearPlano, text='Moldes', variable=self.requerimientos['Moldes'], bootstyle="round-toggle")
        self.cbtnMoldes.grid(row=0,column=5,columnspan=2,pady=5)
        #Crear y ubicar el checkbutton para los Arenados
        self.cbtnArenados=tkb.Checkbutton(self.frmCrearPlano, text='Arenados', variable=self.requerimientos['Arenados'], bootstyle="round-toggle")
        self.cbtnArenados.grid(row=0,column=7,columnspan=2,pady=5)

        self.brocas=dict.fromkeys(['8','10','12','16','18','20','26','30','40','45','50','60','70','80','100'])
        for broca in self.brocas:
            self.brocas[broca]=tkb.BooleanVar(value=False)
        #Crear y ubicar el Frame para seleccionar los diámetros de brocas
        self.frmBrocas=tkb.LabelFrame(self.frmCrearPlano,text='Diámetros de brocas requeridas')
        self.frmBrocas.grid(row=1,column=5,columnspan=5, rowspan=5, padx=5,pady=5,sticky='we')
        #Crear y ubicar el checkbutton para el diámetro 8
        self.cbtn8mm=tkb.Checkbutton(self.frmBrocas, text='Ø 8mm', variable=self.brocas['8'])
        self.cbtn8mm.grid(padx=10,pady=10,row=0,column=0)
        #Crear y ubicar el checkbutton para el diámetro 10
        self.cbtn10mm=tkb.Checkbutton(self.frmBrocas, text='Ø 10mm', variable=self.brocas['10'])
        self.cbtn10mm.grid(padx=10,pady=10,row=0,column=1)
        #Crear y ubicar el checkbutton para el diámetro 12
        self.cbtn12mm=tkb.Checkbutton(self.frmBrocas, text='Ø 12mm', variable=self.brocas['12'])
        self.cbtn12mm.grid(padx=10,pady=10,row=0,column=2)
        #Crear y ubicar el checkbutton para el diámetro 16
        self.cbtn16mm=tkb.Checkbutton(self.frmBrocas, text='Ø 16mm', variable=self.brocas['16'])
        self.cbtn16mm.grid(padx=10,pady=10,row=0,column=3)

        #Crear y ubicar el checkbutton para el diámetro 18
        self.cbtn18mm=tkb.Checkbutton(self.frmBrocas, text='Ø 18mm', variable=self.brocas['18'])
        self.cbtn18mm.grid(padx=10,pady=10,row=1,column=0)
        #Crear y ubicar el checkbutton para el diámetro 20
        self.cbtn20mm=tkb.Checkbutton(self.frmBrocas, text='Ø 20mm', variable=self.brocas['20'])
        self.cbtn20mm.grid(padx=10,pady=10,row=1,column=1)
        #Crear y ubicar el checkbutton para el diámetro 26
        self.cbtn26mm=tkb.Checkbutton(self.frmBrocas, text='Ø 26mm', variable=self.brocas['26'])
        self.cbtn26mm.grid(padx=10,pady=10,row=1,column=2)
        #Crear y ubicar el checkbutton para el diámetro 30
        self.cbtn30mm=tkb.Checkbutton(self.frmBrocas, text='Ø 30mm', variable=self.brocas['30'])
        self.cbtn30mm.grid(padx=10,pady=10,row=1,column=3)

        #Crear y ubicar el checkbutton para el diámetro 40
        self.cbtn40mm=tkb.Checkbutton(self.frmBrocas, text='Ø 40mm', variable=self.brocas['40'])
        self.cbtn40mm.grid(padx=10,pady=10,row=2,column=0)
        #Crear y ubicar el checkbutton para el diámetro 45
        self.cbtn45mm=tkb.Checkbutton(self.frmBrocas, text='Ø 45mm', variable=self.brocas['45'])
        self.cbtn45mm.grid(padx=10,pady=10,row=2,column=1)
        #Crear y ubicar el checkbutton para el diámetro 50
        self.cbtn50mm=tkb.Checkbutton(self.frmBrocas, text='Ø 50mm', variable=self.brocas['50'])
        self.cbtn50mm.grid(padx=10,pady=10,row=2,column=2)
        #Crear y ubicar el checkbutton para el diámetro 60
        self.cbtn60mm=tkb.Checkbutton(self.frmBrocas, text='Ø 60mm', variable=self.brocas['60'])
        self.cbtn60mm.grid(padx=10,pady=10,row=2,column=3)

        #Crear y ubicar el checkbutton para el diámetro 70
        self.cbtn70mm=tkb.Checkbutton(self.frmBrocas, text='Ø 70mm', variable=self.brocas['70'])
        self.cbtn70mm.grid(padx=10,pady=10,row=3,column=0)
        #Crear y ubicar el checkbutton para el diámetro 80
        self.cbtn80mm=tkb.Checkbutton(self.frmBrocas, text='Ø 80mm', variable=self.brocas['80'])
        self.cbtn80mm.grid(padx=10,pady=10,row=3,column=1)
        #Crear y ubicar el checkbutton para el diámetro 100
        self.cbtn100mm=tkb.Checkbutton(self.frmBrocas, text='Ø 100mm', variable=self.brocas['100'])
        self.cbtn100mm.grid(padx=10,pady=10,row=3,column=2)

        #Crear y ubicar la entrada para el total a pagar
        tkb.Label(self.frmCrearPlano,text="A Pagar:",font=("Arial 10 bold"),anchor='w').grid(row=6,column=5, padx=5,pady=3,sticky='w')
        self.txtTotalPago=tkb.Entry(self.frmCrearPlano, width=10, font=("Arial 10"),validate='key')
        self.txtTotalPago.grid(row=6,column=6,pady=3,sticky='w')
        #Crear y ubicar la entrada para el abono
        tkb.Label(self.frmCrearPlano,text="Abono:",font=("Arial 10 bold"),anchor='e').grid(row=6,column=7, padx=5,pady=3,sticky='e')
        self.txtAbono=tkb.Entry(self.frmCrearPlano, width=10, font=("Arial 10"),validate='key')
        self.txtAbono.grid(row=6,column=8,padx=5,pady=3,sticky='w')

        #Crear y ubicar la entrada para el método de pago
        tkb.Label(self.frmCrearPlano,text="Método:",font=("Arial 10 bold"),anchor='w').grid(row=7,column=5, padx=5,pady=3,sticky='w')
        self.cboxMetodo=tkb.Combobox(self.frmCrearPlano,font=("Arial 10"),values=['Efectivo','Transferencia', 'Cheque'])
        self.cboxMetodo.grid(row=7,column=6,pady=3, columnspan=3, sticky='we')
        self.cboxMetodo.current(0)

        #Crear y ubicar la entrada para el No. de documento
        tkb.Label(self.frmCrearPlano,text="No. Doc.:",font=("Arial 10 bold"),anchor='w').grid(row=8,column=5, padx=5,pady=3,sticky='w')
        self.txtDocumento=tkb.Entry(self.frmCrearPlano, width=10, font=("Arial 10"))
        self.txtDocumento.grid(row=8,column=6,columnspan=2,pady=3,sticky='we')
        self.cboxDocumento=tkb.Combobox(self.frmCrearPlano,font=("Arial 10"), width=10, values=['Factura','Ticket'])
        self.cboxDocumento.grid(row=8,column=8,padx=5,pady=3,columnspan=2,sticky='w')
        self.cboxDocumento.current(0)
        #Crear y ubicar la entrada para el plazo de entrega
        self.diasPlazo=tkb.IntVar(value=2)
        self.fecha_plazo=datetime.now()+timedelta(days=2)
        tkb.Label(self.frmCrearPlano,text="Plazo:",font=("Arial 10 bold"),anchor='w').grid(row=9,column=5, padx=5,pady=3,sticky='w')
        self.spboxPlazo=tkb.Spinbox(self.frmCrearPlano, width=3, font=("Arial 10"),from_=1,to=1000,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"),textvariable=self.diasPlazo)
        self.spboxPlazo.grid(row=9,column=6,pady=3,sticky='w')
        self.diasPlazo.trace_add('write',self.actualizar_fecha)
        self.lblFechaPlazo=tkb.Label(self.frmCrearPlano,text="---",font=("Arial 10 bold"),anchor='w')
        self.lblFechaPlazo.grid(row=9,column=7,columnspan=3,padx=5,pady=3,sticky='w')
        self.actualizar_fecha(a=None,b=None,c=None)

        #Crear y ubicar el marco para contener los botones.
        self.frmBotonesPlanos = tkb.Frame(self.frmCrearPlano)
        self.frmBotonesPlanos.grid(row=13, column=0, columnspan=9, padx=10, pady=10)
        #Crear el botón para Guardar modificaciones en un plano
        self.btnGuardar=tkb.Button(self.frmBotonesPlanos,text="Guardar", width=10,command=self.guardar_plano, bootstyle="success")
        self.btnGuardar.pack(padx=5,pady=5,side='left')
        #Crear el botón para Cancelar la modificación de un plano
        self.btnCancelar=tkb.Button(self.frmBotonesPlanos,text="Cancelar", width=10,command=self.cancelar)
        self.btnCancelar.pack(padx=5,pady=5,side='left')

        self.recuperarPlano()

    def recuperarPlano(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM planos WHERE id = %s;",(self.id_plano,))
        datosPlano=cursor.fetchone()
        conn.commit()

        if datosPlano['cliente'] is not None:
            self.txtCliente.insert(1, datosPlano['cliente'])
            self.txtDestino.delete(0, END)
            self.txtDestino.insert(1, datosPlano['destino'])
            if datosPlano['id_materia_prima'] is None:
                self.txtColor.insert(1,datosPlano['color'])
                self.txtEspesor.insert(1,datosPlano['espesor'])
                self.bandera=True
                
                self.lblColor.grid(row=3,column=0, padx=5,pady=3,sticky='w')
                self.txtColor.grid(row=3,column=1,pady=3,sticky='w')
                self.lblEspesor.grid(row=3,column=2, padx=5,pady=3,sticky='e')
                self.txtEspesor.grid(row=3,column=3,pady=3,sticky='w')

                self.cboxMateriaPrima.set('')
                self.cboxMateriaPrima.config(state='disabled',bootstyle="readonly")
                self.btnOtraMateria.config(text='Stock')
                self.bandera=False

            else:
                cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(datosPlano['id_materia_prima'],))
                materia_prima=cursor.fetchone()
                conn.commit()
                self.cboxMateriaPrima.config(state='enabled')
                self.cboxMateriaPrima.set(materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')

                self.lblColor.grid_forget()
                self.txtColor.grid_forget()
                self.lblEspesor.grid_forget()
                self.txtEspesor.grid_forget()

                self.btnOtraMateria.config(text='Otra')
                self.bandera=True
            self.spboxCantidad.set(int(datosPlano['cantidad_vidrios']))
            self.txtArea.insert(1, str(datosPlano['area']))
            self.cboxTT.delete(0, END)
            self.cboxTT.insert(1,datosPlano['tipo_templado'])
            self.cboxSello.delete(0, END)
            self.cboxSello.insert(1,datosPlano['sello'])
            self.txtNota1.insert(1,datosPlano['comentario1'])
            self.txtNota2.insert(1,datosPlano['comentario2'])
            self.txtNota3.insert(1,datosPlano['comentario3'])

            self.requerimientos['Moldes'].set(value=bool(datosPlano['moldes']))
            self.requerimientos['Arenados'].set(value=bool(datosPlano['arenados']))

            brocas=datosPlano['brocas'].split('/')
            for broca in brocas:
                if bool(broca):
                    self.brocas[broca].set(value=True)

            self.txtTotalPago.insert(1,str(datosPlano['total']))
            self.txtAbono.insert(1,str(datosPlano['abono']))
            self.cboxMetodo.delete(0, END)
            self.cboxMetodo.insert(1,datosPlano['metodo'])
            self.txtDocumento.insert(1,datosPlano['no_doc'])
            self.cboxDocumento.delete(0, END)
            self.cboxDocumento.insert(1,datosPlano['tipo_doc'])

            dias=np.busday_count(datosPlano['marca_creacion'].date(), datosPlano['marca_plazo'].date())
            self.diasPlazo.set(value=dias)

            self.btnGuardar.pack(padx=5,pady=5,side='left')
            self.btnCancelar.pack(padx=5,pady=5,side='left')

            #Mostrar el formulario
            self.frmCrearPlano.pack(padx=5, pady=5, ipadx=10, ipady=10, fill='x')
        conn.close()

    def otra_materia_prima(self):
        if self.bandera:
            self.lblColor.grid(row=3,column=0, padx=5,pady=3,sticky='w')
            self.txtColor.grid(row=3,column=1,pady=3,sticky='w')
            self.lblEspesor.grid(row=3,column=2, padx=5,pady=3,sticky='e')
            self.txtEspesor.grid(row=3,column=3,pady=3,sticky='w')

            self.cboxMateriaPrima.set('')
            self.cboxMateriaPrima.config(state='disabled',bootstyle="readonly")
            self.btnOtraMateria.config(text='Stock')
            self.bandera=False

        else:
            self.cboxMateriaPrima.config(state='enabled')

            self.txtColor.delete(0,END)
            self.txtEspesor.set('')

            self.lblColor.grid_forget()
            self.txtColor.grid_forget()
            self.lblEspesor.grid_forget()
            self.txtEspesor.grid_forget()

            self.btnOtraMateria.config(text='Otra')
            self.bandera=True

    def verificar_campos(self)->bool:
        if not bool(self.txtCliente.get()):
            messagebox.showerror(title='Entrada no válida', message='La casilla del nombre del cliente no puede quedar vacía.')
            return False
        if not bool(self.txtDestino.get()):
            messagebox.showerror(title='Entrada no válida', message='La casilla de la ciudad de destino no puede quedar vacía.')
            return False
        
        if self.bandera:
            if self.cboxMateriaPrima.current() == -1:
                messagebox.showerror(title='Entrada no válida', message='Debe seleccionar una de las opciones para la materia prima.')
                return False
        else:
            if not bool(self.txtColor.get()):
                messagebox.showerror(title='Entrada no válida', message='La casilla de color de la materia prima no puede quedar vacía.')
                return False
            if not bool(self.txtEspesor.get()):
                messagebox.showerror(title='Entrada no válida', message='La casilla del espesor de la materia prima no puede quedar vacía.')
                return False
            
        if not self.spboxCantidad.get().isdigit():
            messagebox.showerror(title='Entrada no válida', message='La casilla de la cantidad de vidrios debe ser un número entero.')
            return False
        if not bool(self.txtArea.get()):
            partes = str(self.txtArea.get()).split(".")
            if not(len(partes) == 2 and len(partes[1]) == 2):
                messagebox.showerror(title='Entrada no válida', message='La casilla del área total del plano debe ser un número con dos dígitos decimales.')
                return False
            messagebox.showerror(title='Entrada no válida', message='La casilla del área total del plano no puede quedar vacía.')
            return False
        if self.cboxTT.current() == -1:
            messagebox.showerror(title='Entrada no válida', message='Debe seleccionar una de las opciones para el tipo de templado.')
            return False
        if self.cboxSello.current() == -1:
            messagebox.showerror(title='Entrada no válida', message='Debe seleccionar una de las opciones para la ubicación del sello.')
            return False

        if not bool(self.txtTotalPago.get()):
            partes = str(self.txtTotalPago.get()).split(".")
            if not(len(partes) == 2 and len(partes[1]) == 2):
                messagebox.showerror(title='Entrada no válida', message='La casilla del valor total a pagar debe ser un número con dos dígitos decimales.')
                return False
            messagebox.showerror(title='Entrada no válida', message='La casilla del valor total a pagar no puede quedar vacía.')
            return False
        if not bool(self.txtAbono.get()):
            partes = str(self.txtAbono.get()).split(".")
            if not(len(partes) == 2 and len(partes[1]) == 2):
                messagebox.showerror(title='Entrada no válida', message='La casilla del abono debe ser un número con dos dígitos decimales.')
                return False
            messagebox.showerror(title='Entrada no válida', message='La casilla del abono no puede quedar vacía.')
            return False
        if not bool(self.txtDocumento.get()):
            messagebox.showerror(title='Entrada no válida', message='La casilla del número del documento no puede quedar vacía.')
            return False
        if self.cboxDocumento.current() == -1:
            messagebox.showerror(title='Entrada no válida', message='Debe seleccionar una de las opciones para el tipo de documento de pago.')
            return False
        if self.cboxMetodo.current() == -1:
            messagebox.showerror(title='Entrada no válida', message='Debe seleccionar una de las opciones para el método de pago.')
            return False
        return True

    def actualizar_fecha(self, a, b, c):
        marca_temporal=datetime.now()
        self.fecha_plazo=marca_temporal
        try:
            delta=int(self.diasPlazo.get())
            while delta>0:
                self.fecha_plazo=self.fecha_plazo+timedelta(days=1)
                if self.fecha_plazo.isoweekday()<6:
                    delta=delta-1
            self.lblFechaPlazo.config(text=str(self.fecha_plazo.strftime("%d-%m-%Y (%A)")))
        except:
            self.lblFechaPlazo.config(text='----')

    def limpiarFormulario(self):
        self.txtCliente.delete(0, END)
        self.txtDestino.delete(0, END)
        self.txtDestino.insert(1,'Guayaquil')
        self.cboxMateriaPrima.set('')
        self.spboxCantidad.set(1)
        self.txtArea.delete(0, END)
        self.cboxTT.current(0)
        self.cboxSello.current(0)
        self.txtNota1.delete(0, END)
        self.txtNota2.delete(0, END)
        self.txtNota3.delete(0, END)

        self.txtColor.delete(0, END)
        self.txtEspesor.delete(0, END)

        self.requerimientos['Moldes'].set(value=False)
        self.requerimientos['Arenados'].set(value=False)
        for broca in self.brocas:
            self.brocas[broca].set(value=False)

        self.txtTotalPago.delete(0, END)
        self.txtAbono.delete(0, END)
        self.cboxMetodo.current(0)
        self.txtDocumento.delete(0, END)
        self.cboxDocumento.current(0)
        self.diasPlazo.set(value=2)

        self.frmCrearPlano.config(text='Crear Plano')

    def guardar_plano(self):
        if self.verificar_campos():

            cliente=self.txtCliente.get()
            destino=self.txtDestino.get()
            if self.bandera:
                color=None
                espesor=None
                id_materia_prima=self.dictMateriaPrimaInvert[self.cboxMateriaPrima.get()]
            else:
                color=self.txtColor.get()
                espesor=self.txtEspesor.get()
                id_materia_prima=None
            cantidad_vidrios=self.spboxCantidad.get()
            area=Decimal(self.txtArea.get())
            tipo_templado=self.cboxTT.get()
            sello=self.cboxSello.get()
            estado='Creado'
            comentario1=self.txtNota1.get()
            comentario2=self.txtNota2.get()
            comentario3=self.txtNota3.get()
            id_empleado=self.id_empleado
            check_ventas=None
            check_caja=None
            marca_creacion=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            marca_plazo=self.fecha_plazo.strftime("%Y-%m-%d %H:%M:%S")
            total=Decimal(self.txtTotalPago.get())
            abono=Decimal(self.txtAbono.get())
            saldo=Decimal(self.txtTotalPago.get())-Decimal(self.txtAbono.get())
            metodo=self.cboxMetodo.get()
            tipo_doc=self.cboxDocumento.get()
            no_doc=self.txtDocumento.get()
            if self.requerimientos['Moldes'].get():
                moldes=1
            else:
                moldes=0
            if self.requerimientos['Arenados'].get():
                arenados=1
            else:
                arenados=0

            Aux=''
            for broca in self.brocas:
                if self.brocas[broca].get():
                    Aux=Aux+'/'+ broca
            brocas=Aux
            #----------------Actualizar la instancia en la base de datos---------------
            sentencia = (
            "UPDATE planos SET id_empleado=%s, id_materia_prima=%s, cliente=%s, destino=%s, cantidad_vidrios=%s, marca_creacion=%s, dias_plazo=%s, marca_plazo=%s, comentario1=%s, comentario2=%s, comentario3=%s, tipo_templado=%s, estado=%s, color=%s, espesor=%s, sello=%s, area=%s, moldes=%s, arenados=%s, brocas=%s, total=%s, abono=%s, saldo=%s, metodo=%s, tipo_doc=%s, no_doc=%s WHERE id=%s;"
            )
            datos=(id_empleado, id_materia_prima, cliente.title(), destino, cantidad_vidrios, marca_creacion, self.diasPlazo.get(), marca_plazo, comentario1, comentario2, comentario3, tipo_templado, estado, color, espesor, sello, area, moldes, arenados, brocas, total, abono, saldo, metodo, tipo_doc, no_doc, self.id_plano)
            
            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            cursor.execute(sentencia, datos)
            conn.commit()
            conn.close()
            
            self.callback()
            self.destroy()

    def cancelar(self):
        self.destroy()