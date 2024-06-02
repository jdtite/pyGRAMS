import ttkbootstrap as tkb
from tkinter import messagebox
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
import os
from decimal import Decimal
import create_document
from connection import connection
import win32print  
import win32api

class Main(tkb.Frame):
    def __init__(self, *args, ventana:tkb.Window, register:tkb.Window.register, id_empleado:int, usuario:str, **kwargs):
        super().__init__(*args, **kwargs)
        #Carpeta principal
        self.carpeta_principal = os.path.dirname(__file__)
        #.\igrams\gui\documentos
        self.carpeta_documentos = os.path.join(self.carpeta_principal, "documentos")
        #.\igrams\gui\images
        self.carpeta_imagenes = os.path.join(self.carpeta_principal, "images")

        #instanciar cursor
        self.ventana=ventana

        self.NoPlano=tkb.StringVar()
        self.registro=register
        #Define el Frame como contenedor total que se asocia al notebook
        self.contenedor=ScrolledFrame(self)
        self.contenedor.pack(padx=10,pady=10, fill='both', expand=True)
        #Define la primera columna del contenedor
        self.colum1=tkb.Frame(self.contenedor)
        self.colum1.pack(padx=10, pady=10, side='left', fill='y')

        #Crear y ubicar el Frame para la búsqueda
        self.frmBuscar=tkb.Labelframe(self.colum1, text='Plano No:', relief='flat')
        self.frmBuscar.pack(padx=5, pady=5,  fill='x')
        #Crear y ubicar el cuadro de texto para el No Plano
        self.txtNoPlano=tkb.Entry(self.frmBuscar, width=10,font=("Times 12"),textvariable=self.NoPlano,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"))
        self.txtNoPlano.pack(padx=5, pady=5, side='left')
        self.txtNoPlano.bind("<Return>", self.onEnter)
        self.txtNoPlano.focus()
        #Crear y ubicar el botón para buscar
        self.btnBuscar=tkb.Button(self.frmBuscar,text="Buscar",command=self.consultar_plano)
        self.btnBuscar.pack(padx=5, pady=5,side='left')

        #Crear y ubicar el Frame para la información del Plano
        self.frmInfo=tkb.LabelFrame(self.colum1,text='Información del Plano')
        self.frmInfo.pack(padx=5, pady=5, fill='x')
        #Crear y ubicar la etiqueta para el No. de Plano
        tkb.Label(self.frmInfo,text="Plano No:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=0,column=0, padx=5,pady=3,sticky='w')
        self.lblNoPlano=tkb.Label(self.frmInfo,text="",font=("Arial 20 bold"),anchor="w", justify="left")
        self.lblNoPlano.grid(row=0,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de vidrios
        tkb.Label(self.frmInfo,text="Cantidad:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=1,column=0, padx=5,pady=3,sticky='w')
        self.lblCantidad=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblCantidad.grid(row=1,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para la materia prima
        tkb.Label(self.frmInfo,text="Materia:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=2,column=0,padx=5,pady=3,sticky='w')
        self.lblMateriaPrima=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblMateriaPrima.grid(row=2,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para el Tipo de Templado
        tkb.Label(self.frmInfo,text="Tipo Temp.:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=3,column=0,padx=5,pady=3,sticky='w')
        self.lblTT=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblTT.grid(row=3,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para el Cliente
        tkb.Label(self.frmInfo,text="Cliente:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=4,column=0,padx=5,pady=3,sticky='w')
        self.lblCliente=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left",width=25)
        self.lblCliente.grid(row=4,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para el Inicio
        tkb.Label(self.frmInfo,text="Inicio:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=5,column=0,padx=5,pady=3,sticky='w')
        self.lblInicio=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblInicio.grid(row=5,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para el Plazo
        tkb.Label(self.frmInfo,text="Plazo:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=6,column=0,padx=5,pady=3,sticky='w')
        self.lblPlazo=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblPlazo.grid(row=6,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para el Vendedor
        tkb.Label(self.frmInfo,text="Vendedor:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=7,column=0,padx=5,pady=3,sticky='w')
        self.lblVendedor=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblVendedor.grid(row=7,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para el Área vendida del plano
        tkb.Label(self.frmInfo,text="Area V.:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=8,column=0,padx=5,pady=3,sticky='w')
        self.lblAreaVendida=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblAreaVendida.grid(row=8,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para el Área calculada del plano
        tkb.Label(self.frmInfo,text="Area R.:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=9,column=0,padx=5,pady=3,sticky='w')
        self.lblAreaCalculada=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblAreaCalculada.grid(row=9,column=1,pady=3,sticky='w')

        #Crear y ubicar el marco para contener los botones.
        self.frmBotonesPlano = tkb.Frame(self.frmInfo)
        self.frmBotonesPlano.grid(row=10, column=0, columnspan=2, padx=5, pady=5)
        #Crear y ubicar el checkbutton para los Moldes
        self.bloqueo=tkb.BooleanVar(value=False)
        self.cbtnBloqueo=tkb.Checkbutton(self.frmBotonesPlano, text='Bloquear', variable=self.bloqueo, bootstyle="round-toggle",command=self.bloquear_edicion)
        self.cbtnBloqueo.pack(padx=5, pady=5)
        #Crear y ubicar el botón para ver la portada
        self.btnImprimirTodo=tkb.Button(self.frmBotonesPlano,text="Imprimir todo",command=self.imprimir_todo, state='disabled',bootstyle="disabled")
        self.btnImprimirTodo.pack(padx=5, pady=5)

        #-----------------------------------------------------------------------------------------------
        
        #Define la segunda columna del contenedor
        self.colum2=tkb.Frame(self.contenedor)
        self.colum2.pack(padx=10, pady=10, side='left', fill='y')

        #Definir impresoras---------------------------------------------------
        impresoras=win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME, None, 2)
        listaImpresoras=list()
        for impresora in impresoras:
            listaImpresoras.append(impresora['pPrinterName'])

        tkb.Label(self.colum2,text="Impresoras",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.cboxImpresoras=tkb.Combobox(self.colum2,font=("Arial 10"),values=listaImpresoras)
        self.cboxImpresoras.pack(padx=3,pady=3,anchor='w')
        self.cboxImpresoras.insert(0,win32print.GetDefaultPrinter())
        #----------------------------------------------------------------------

        #Crear y ubicar el Frame para el formulario de los vidrios
        self.frmInfoVidrios=tkb.LabelFrame(self.colum2,text='Lista de vidrios')
        self.frmInfoVidrios.pack(padx=5, pady=5, ipadx=5, ipady=5)

        #Primera columna del formulario---------------------------------------------
        c1=tkb.Frame(self.frmInfoVidrios)
        c1.pack(side='left',anchor='n', padx=5)

        #Crear y ubicar la etiqueta para el número de vidrio
        self.frmNoVidrio=tkb.Labelframe(c1, text='Vidrio No', relief='sunken')
        self.frmNoVidrio.grid(row=0,column=0, padx=5, pady=5, columnspan=2)
        self.lblNoVidrio=tkb.Label(self.frmNoVidrio,font=("Times 16"), anchor='center')
        self.lblNoVidrio.pack(fill='x')

        #Crear y ubicar la etiqueta para las unidades
        self.frmUnidades=tkb.Frame(c1)
        self.frmUnidades.grid(row=1,column=0, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmUnidades,text="Unidades",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.spboxUnidades=tkb.Spinbox(self.frmUnidades, width=4, font=("Arial 10"),from_=1,to=1000,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"), state='readonly')
        self.spboxUnidades.pack(padx=3,pady=3,anchor='w')
        self.spboxUnidades.set(1)

        #Crear y ubicar la entrada para el Tipo de Borde
        self.frmBorde=tkb.Frame(c1)
        self.frmBorde.grid(row=1,column=1, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmBorde,text="Borde",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.cboxBorde=tkb.Combobox(self.frmBorde,font=("Arial 10"),width=5,values=['BPB','BPM','FM', 'V/P'], state='readonly')
        self.cboxBorde.pack(padx=3,pady=3,anchor='w')
        self.cboxBorde.current(0)

        #Crear y ubicar la entrada para la forma del vidrio
        self.frmForma=tkb.Frame(c1)
        self.frmForma.grid(row=2,column=0, padx=5, pady=5, columnspan=2, sticky='we')
        tkb.Label(self.frmForma,text="Forma",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.cboxForma=tkb.Combobox(self.frmForma,font=("Arial 10"),values=['Rectangular','Molde','Irregular', 'Circular'], state='readonly')
        self.cboxForma.pack(padx=3,pady=3,anchor='w')
        self.cboxForma.current(0)

        #Crear y ubicar la entrada para el ancho del vidrio
        self.frmAncho=tkb.Frame(c1)
        self.frmAncho.grid(row=3,column=0, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmAncho,text="Ancho",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.descAncho=tkb.BooleanVar(value=False)
        self.cbtnDescAncho=tkb.Checkbutton(self.frmAncho, text='Desc.', variable=self.descAncho, bootstyle="round-toggle",command=self.descuadre_ancho, state='readonly')
        self.cbtnDescAncho.pack(padx=3,pady=3,anchor='w')
        self.spboxAncho1=tkb.Spinbox(self.frmAncho, width=5, font=("Arial 10"),from_=1,to=3660,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"), state='readonly')
        self.spboxAncho1.pack(padx=3,pady=3,anchor='w')
        self.spboxAncho1.set(0)
        self.spboxAncho2=tkb.Spinbox(self.frmAncho, width=5, font=("Arial 10"),from_=1,to=3660,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"), state='readonly')
        self.spboxAncho2.set(0)

        #Crear y ubicar la entrada para el alto del vidrio
        self.frmAlto=tkb.Frame(c1)
        self.frmAlto.grid(row=3,column=1, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmAlto,text="Alto",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.descAlto=tkb.BooleanVar(value=False)
        self.cbtnDescAlto=tkb.Checkbutton(self.frmAlto, text='Desc.', variable=self.descAlto, bootstyle="round-toggle",command=self.descuadre_alto, state='readonly')
        self.cbtnDescAlto.pack(padx=3,pady=3,anchor='w')
        self.spboxAlto1=tkb.Spinbox(self.frmAlto, width=5, font=("Arial 10"),from_=1,to=3660,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"), state='readonly')
        self.spboxAlto1.pack(padx=3,pady=3,anchor='w')
        self.spboxAlto1.set(0)
        self.spboxAlto2=tkb.Spinbox(self.frmAlto, width=5, font=("Arial 10"),from_=1,to=3660,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"), state='readonly')
        self.spboxAlto2.set(0)

        #Crear y ubicar la entrada para la cantidad de saques
        self.frmSaques=tkb.Frame(c1)
        self.frmSaques.grid(row=4,column=0, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmSaques,text="Saques",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.spboxSaques=tkb.Spinbox(self.frmSaques, width=4, font=("Arial 10"),from_=1,to=999,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"), state='readonly')
        self.spboxSaques.pack(padx=3,pady=3,anchor='w')
        self.spboxSaques.set(0)

        #Crear y ubicar la entrada para la cantidad de perforaciones
        self.frmPerforaciones=tkb.Frame(c1)
        self.frmPerforaciones.grid(row=4,column=1, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmPerforaciones,text="Perf.",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.spboxPerforaciones=tkb.Spinbox(self.frmPerforaciones, width=4, font=("Arial 10"),from_=1,to=999,increment=1,validate='key',validatecommand=(self.registro(lambda text: text.isdecimal()), "%S"), state='readonly')
        self.spboxPerforaciones.pack(padx=3,pady=3,anchor='w')
        self.spboxPerforaciones.set(0)

        #Crear y ubicar la entrada para el tipo de arenado
        self.frmArenado=tkb.Frame(c1)
        self.frmArenado.grid(row=5,column=0, columnspan=2, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmArenado,text="Arenado",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.cboxArenado=tkb.Combobox(self.frmArenado,font=("Arial 10"),values=['Ninguno','Total','Franjas', 'Diseño'], state='readonly')
        self.cboxArenado.pack(padx=3,pady=3,anchor='w')
        self.cboxArenado.current(0)

        #Segunda columna del formulario---------------------------------------------
        c2=tkb.Frame(self.frmInfoVidrios)
        c2.pack(side='left',anchor='n', padx=5)

        #Crear y ubicar el marco para contener la tabla y la barra de desplazamiento
        self.frmTablaVidrios = tkb.Frame(c2)
        self.frmTablaVidrios.grid(row=0, column=0, sticky='we')
        self.frmTablaVidrios.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmTablaVidrios.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmTablaVidrios, orient=VERTICAL)
        scrollbarV.grid(row=0,column=1, pady=5, sticky='ns')

        fuente=tkb.font.Font(family="times", size=11)
        style = tkb.Style()
        style.configure("mystyle.Treeview", font=('times', 11))

        #Crear y ubicar la tabla de los vidrios
        self.tableInfoVidrios=tkb.Treeview(self.frmTablaVidrios, show='headings', height=10, yscrollcommand=scrollbarV.set, style="mystyle.Treeview", )
        self.tableInfoVidrios.configure(columns=('vidrio', 'medidas'))
        self.tableInfoVidrios.column('vidrio', stretch=False, anchor='center', width=fuente.measure(' 0000000 '))
        self.tableInfoVidrios.column('medidas', stretch=False, anchor='center', width=fuente.measure(' 00000 X 00000 '))
        
        for col in self.tableInfoVidrios['columns']:
            self.tableInfoVidrios.heading(col, text=col.title(), anchor=W)

        self.tableInfoVidrios.grid(row=0,column=0, padx=5, pady=5)
        self.tableInfoVidrios.tag_configure('listo', background='sky blue')
        self.tableInfoVidrios.bind("<<TreeviewSelect>>", self.tableVidrios_change)

        scrollbarV.config(command=self.tableInfoVidrios.yview)

        #Crear y ubicar la entrada para la referencia
        self.frmReferencia=tkb.Frame(c2)
        self.frmReferencia.grid(row=1,column=0, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmReferencia,text="Referencia",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.varRefencia=tkb.StringVar()
        self.txtRefencia=tkb.Entry(self.frmReferencia, font=("Arial 10"),textvariable=self.varRefencia, state='readonly')
        self.txtRefencia.pack(padx=3,pady=3,anchor='w')

        #Crear y ubicar la entrada para la observación
        self.frmObservacion=tkb.Frame(c2)
        self.frmObservacion.grid(row=2,column=0, padx=5, pady=5, sticky='we')
        tkb.Label(self.frmObservacion,text="Observación",font=("Arial 10 bold"),anchor='w').pack(fill='x')
        self.varObservacion=tkb.StringVar()
        self.txtObservacion=tkb.Entry(self.frmObservacion, font=("Arial 10"),textvariable=self.varObservacion, state='readonly')
        self.txtObservacion.pack(padx=3,pady=3,anchor='w')

        #Crear y ubicar el botón para Guardar el vidrio
        self.btnGuardar=tkb.Button(c2,text="Guardar", width=15, command=self.guardar_vidrio, state='disabled')
        self.btnGuardar.grid(row=3, column=0, pady=5)

        #Crear y ubicar el botón para Actualizar el vidrio
        self.btnActualizar=tkb.Button(c2,text="Actualizar", width=15, command=self.actualizar_vidrio)
        #Crear y ubicar el botón para Cancelar la actualización
        self.btnCancelar=tkb.Button(c2,text="Cancelar", width=15, command=self.cancelar_edicion, bootstyle="danger")

        #Crear y ubicar el botón para editar el vidrio
        self.btnEditar=tkb.Button(c2,text="Editar", width=15, command=self.editar_vidrio)
        #Crear y ubicar el botón para imprimir la etiqueta
        self.btnImprimir=tkb.Button(c2,text="Imprimir", width=15, command=self.imprimir_etiqueta, bootstyle="info")

    def onEnter(self, event):
        self.consultar_plano()

    def consultar_plano(self):
        if  self.NoPlano.get() != '':
            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            cursor.execute("SELECT planos.*, empleados.nombre1, empleados.apellido1 FROM planos INNER JOIN empleados ON planos.id_empleado = empleados.id WHERE planos.id = %s;",(int(self.NoPlano.get()),))
            self.datosPlano=cursor.fetchone()
            conn.commit()

            if self.datosPlano is not None: #Si el plano existe
                self.lblNoPlano.config(text=self.NoPlano.get())
                self.NoPlano.set('')

                self.lblCantidad.config(text=str(self.datosPlano['cantidad_vidrios']))
                #Cosultar la materia prima
                if self.datosPlano['id_materia_prima'] is None:
                    self.lblMateriaPrima.config(text=self.datosPlano['color'] + ' ' + str(self.datosPlano['espesor']) + ' mm')
                else:
                    cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(self.datosPlano['id_materia_prima'],))
                    materia_prima=cursor.fetchone()
                    conn.commit()
                    self.lblMateriaPrima.config(text=materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')
                
                self.lblTT.config(text=self.datosPlano['tipo_templado'])
                self.lblCliente.config(text=self.datosPlano['cliente'])
                self.lblInicio.config(text=self.datosPlano['marca_etiquetado'])
                self.lblPlazo.config(text=self.datosPlano['marca_plazo'])

                self.lblVendedor.config(text=self.datosPlano['nombre1'] + ' ' + self.datosPlano['apellido1'])

                self.lblAreaVendida.config(text=str(self.datosPlano['area']) + ' m²')

                self.btnGuardar.configure(state='enabled')

                #Consultar si el plano ya ha sido etiquetado
                self.bloqueo.set(value=bool(self.datosPlano['marca_etiquetado']))

                self.consultar_vidrios()
                #Mostrar el área calculada
                self.lblAreaCalculada.configure(text=str(round(self.calcular_area(),2))+' m²')
            else:
                messagebox.showerror(title='Plano no encontrado', message='Su plano no se encuentra registrado.')
            conn.close()
        else:
            messagebox.showerror(title='Plano no definido', message='La casilla de número de plano no puede quedar vacía.')
    
    def consultar_vidrios(self):
        #Crear el contenedor para la cantidad de vidrios definida en el plano
        self.vidrios=list()
        for i in range(self.datosPlano['cantidad_vidrios']):
            self.vidrios.append(dict.fromkeys(('id','unidades', 'tipo_borde', 'forma', 'ancho1', 'alto1', 'ancho2', 'alto2', 'saques', 'perforaciones', 'tipo_arenado', 'referencia', 'observacion', 'marca_etiquetado')))

        #Buscar y listar todos los vidrios--------------------------------------------------
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM vidrios WHERE id_plano=%s;",(int(self.lblNoPlano.cget('text')),))
        self.datosVidrios=cursor.fetchall()
        conn.commit()
        #Consultar el último vidrio registrado
        cursor.execute("SELECT MAX(id) from vidrios;")
        max_id=cursor.fetchone()
        conn.commit()
        conn.close()
        
        self.vidrios[0]['id']=max_id['MAX(id)']+1

        diferencia=len(self.vidrios) - len(self.datosVidrios)
        if diferencia>=0:
            i=-1
            for i in range(len(self.datosVidrios)):
                #Copiar la infomación consultada en el contenedor local
                self.vidrios[i]['id']=self.datosVidrios[i]['id']
                self.vidrios[i]['unidades']=self.datosVidrios[i]['unidades']
                self.vidrios[i]['tipo_borde']=self.datosVidrios[i]['tipo_borde']
                self.vidrios[i]['forma']=self.datosVidrios[i]['forma']
                self.vidrios[i]['ancho1']=self.datosVidrios[i]['ancho1']
                self.vidrios[i]['alto1']=self.datosVidrios[i]['alto1']
                self.vidrios[i]['ancho2']=self.datosVidrios[i]['ancho2']
                self.vidrios[i]['alto2']=self.datosVidrios[i]['alto2']
                self.vidrios[i]['saques']=self.datosVidrios[i]['saques']
                self.vidrios[i]['perforaciones']=self.datosVidrios[i]['perforaciones']
                self.vidrios[i]['tipo_arenado']=self.datosVidrios[i]['tipo_arenado']
                self.vidrios[i]['referencia']=self.datosVidrios[i]['referencia']
                self.vidrios[i]['observacion']=self.datosVidrios[i]['observacion']
                self.vidrios[i]['marca_etiquetado']=self.datosVidrios[i]['marca_etiquetado']

            if diferencia>0:
                #Si existen más vidrios por guardar, el siguiente se numera automáticamente
                self.vidrios[i+1]['id']=max_id['MAX(id)']+1
                self.actualizar_lista(indice=i+1)
            else:
                self.actualizar_lista(indice=0)
            #self.tableInfoVidrios.focus(self.tableInfoVidrios.get_children()[0])
            #self.tableInfoVidrios.selection_set(self.tableInfoVidrios.get_children()[0])

        else:
            messagebox.showerror(title='Inconsistencia detectada', message='La cantidad de vidrios en el plano es menor a la cantidad de vidrios registrados')
        #-----------------------------------------------------------------------------------

    def actualizar_lista(self, indice: int):
        fila=[]
        self.tableInfoVidrios.delete(*self.tableInfoVidrios.get_children())
        for vidrio in self.vidrios:
            #Escribir las filas de la lista con los vidrios guardados
            fila.clear()
            if vidrio['id']:
                fila.append(str(vidrio['id'])[-4:])
            else:
                fila.append('---')

            if vidrio['marca_etiquetado']:
                fila.append(str(vidrio['ancho1'])+ " X " + str(vidrio['alto1']))
            else:
                fila.append('---')

            self.tableInfoVidrios.insert('', index='end', values=fila)

        self.tableInfoVidrios.focus(self.tableInfoVidrios.get_children()[indice])
        self.tableInfoVidrios.selection_set(self.tableInfoVidrios.get_children()[indice])

    def tableVidrios_change(self, Event):
        #Si hay selección en la lista, se envía el índice al formulario para actualizar
        if self.tableInfoVidrios.selection():
            indice=self.tableInfoVidrios.selection()[0]
            self.actualizar_formulario(self.tableInfoVidrios.index(indice))

    def actualizar_formulario(self, indice: int):
        #Si tiene la marca de etiquetado
        if self.vidrios[indice]['marca_etiquetado']:
            #Se bloquean las entradas para mostrar la información
            self.spboxUnidades.configure(state='disabled')
            self.cboxBorde.configure(state='disabled')
            self.cboxForma.configure(state='disabled')
            self.cbtnDescAncho.configure(state='disabled')
            self.cbtnDescAlto.configure(state='disabled')
            self.spboxAncho1.configure(state='disabled')
            self.spboxAncho2.configure(state='disabled')
            self.spboxAlto1.configure(state='disabled')
            self.spboxAlto2.configure(state='disabled')
            self.spboxSaques.configure(state='disabled')
            self.spboxPerforaciones.configure(state='disabled')
            self.cboxArenado.configure(state='disabled')
            self.txtRefencia.configure(state='disabled')
            self.txtObservacion.configure(state='disabled')

            #Se establecen los valores guardados en el contenedor
            self.lblNoVidrio.configure(text=str(self.vidrios[indice]['id'])[-4:])
            self.spboxUnidades.set(self.vidrios[indice]['unidades'])
            self.cboxBorde.set(self.vidrios[indice]['tipo_borde'])
            self.cboxForma.set(self.vidrios[indice]['forma'])
            self.spboxAncho1.set(self.vidrios[indice]['ancho1'])
            self.spboxAncho2.set(self.vidrios[indice]['ancho2'])
            if self.vidrios[indice]['ancho2'] != 0:
                self.descAncho.set(True)
            else:
                self.descAncho.set(False)
            self.descuadre_ancho()
            self.spboxAlto1.set(self.vidrios[indice]['alto1'])
            self.spboxAlto2.set(self.vidrios[indice]['alto2'])
            if self.vidrios[indice]['alto2'] != 0:
                self.descAlto.set(True)
            else:
                self.descAlto.set(False)
            self.descuadre_alto()
            self.spboxSaques.set(self.vidrios[indice]['saques'])
            self.spboxPerforaciones.set(self.vidrios[indice]['perforaciones'])
            self.cboxArenado.set(self.vidrios[indice]['tipo_arenado'])
            self.varRefencia.set(self.vidrios[indice]['referencia'])
            self.varObservacion.set(self.vidrios[indice]['observacion'])

            #Se muestran los botones para editar e imprimir
            self.btnGuardar.grid_forget()
            self.btnActualizar.grid_forget()
            self.btnCancelar.grid_forget()
            self.btnEditar.grid(row=3, column=0, pady=5)
            self.btnImprimir.grid(row=4, column=0, pady=5)

        else:
            if self.vidrios[indice]['id']:
                self.lblNoVidrio.configure(text=str(self.vidrios[indice]['id'])[-4:])
            else:
                self.lblNoVidrio.configure(text='---')
            self.spboxUnidades.set(1)

            #Se habilitan las entradas del formulario
            self.spboxUnidades.configure(state='enabled')
            self.cboxBorde.configure(state='enabled')
            self.cboxForma.configure(state='enabled')
            self.cbtnDescAncho.configure(state='enabled')
            self.cbtnDescAlto.configure(state='enabled')
            self.spboxAncho1.configure(state='enabled')
            self.spboxAncho2.configure(state='enabled')
            self.spboxAlto1.configure(state='enabled')
            self.spboxAlto2.configure(state='enabled')
            self.spboxSaques.configure(state='enabled')
            self.spboxPerforaciones.configure(state='enabled')
            self.cboxArenado.configure(state='enabled')
            self.txtRefencia.configure(state='enabled')
            self.txtObservacion.configure(state='enabled')

            #Se establecen los valores por defecto
            self.spboxUnidades.set(1)
            self.cboxBorde.current(0)
            self.cboxForma.current(0)
            self.spboxAncho1.set(0)
            self.spboxAncho2.set(0)
            self.spboxAlto1.set(0)
            self.spboxAlto2.set(0)
            self.spboxSaques.set(0)
            self.spboxPerforaciones.set(0)
            self.cboxArenado.current(0)
            self.varRefencia.set('')
            self.varObservacion.set('')

            #Se muestra el botón para guardar
            self.btnActualizar.grid_forget()
            self.btnCancelar.grid_forget()
            self.btnEditar.grid_forget()
            self.btnImprimir.grid_forget()
            self.btnGuardar.grid(row=3, column=0, pady=5)

    def bloquear_edicion(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        if self.bloqueo.get():
            self.btnImprimirTodo.configure(state='enabled')
            cursor.execute('UPDATE planos SET marca_etiquetado=NOW() WHERE id=%s',(self.datosPlano['id'],))
        else:
            self.btnImprimirTodo.configure(state='disabled')
            cursor.execute('UPDATE planos SET marca_etiquetado=NULL WHERE id=%s',(self.datosPlano['id'],))
        conn.commit()
        conn.close()
    
    def descuadre_ancho(self):
        if self.descAncho.get():
            self.spboxAncho2.pack(padx=3,pady=3,anchor='w')
        else:
            self.spboxAncho2.pack_forget()
        
    def descuadre_alto(self):
        if self.descAlto.get():
            self.spboxAlto2.pack(padx=3,pady=3,anchor='w')
        else:
            self.spboxAlto2.pack_forget()

    def imprimir_etiqueta(self):
        indice=self.tableInfoVidrios.selection()[0]
        etiqueta=create_document.BarcodeLabel(datosPlano=self.datosPlano, datosVidrio=self.vidrios[self.tableInfoVidrios.index(indice)],materia=self.lblMateriaPrima.cget('text'), contador=self.tableInfoVidrios.index(indice))
        
        pdf_path=etiqueta.get_path()

        printer_handle = win32print.OpenPrinter(self.cboxImpresoras.get())
        job_info=win32print.StartDocPrinter(printer_handle,1,(pdf_path,None,'RAW'))
        win32print.StartPagePrinter(printer_handle)

        win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)

        win32print.EndPagePrinter(printer_handle)
        win32print.EndDocPrinter(printer_handle)
        win32print.ClosePrinter(printer_handle)

    def imprimir_todo(self):
        printer_handle = win32print.OpenPrinter(self.cboxImpresoras.get())
        i=0
        for vidrio in self.vidrios:
            etiqueta=create_document.BarcodeLabel(datosPlano=self.datosPlano, datosVidrio=vidrio, materia=self.lblMateriaPrima.cget('text'), contador=i)
            pdf_path=etiqueta.get_path()

            win32print.StartDocPrinter(printer_handle,1,(pdf_path,None,'RAW'))
            win32print.StartPagePrinter(printer_handle)

            win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)

            win32print.EndPagePrinter(printer_handle)
            win32print.EndDocPrinter(printer_handle)
            i+=1
        
        win32print.ClosePrinter(printer_handle)
        

    def guardar_vidrio(self):
        indice=self.tableInfoVidrios.selection()[0]

        if self.verificar_formulario():
            if self.tableInfoVidrios.index(indice) + int(self.spboxUnidades.get()) <= len(self.vidrios): #Si la unidades no exceden al máximo de vidrios
                #Copiar la infomación del formulario a una variable local
                aux=dict()
                aux['unidades']=int(self.spboxUnidades.get())
                aux['tipo_borde']=self.cboxBorde.get()
                aux['forma']=self.cboxForma.get()
                aux['ancho1']=int(self.spboxAncho1.get())
                aux['alto1']=int(self.spboxAlto1.get())
                aux['ancho2']=int(self.spboxAncho2.get())
                aux['alto2']=int(self.spboxAlto2.get())
                aux['saques']=int(self.spboxSaques.get())
                aux['perforaciones']=int(self.spboxPerforaciones.get())
                aux['tipo_arenado']=self.cboxArenado.get()
                aux['referencia']=self.txtRefencia.get()
                aux['observacion']=self.txtObservacion.get()

                for i in range(int(self.spboxUnidades.get())):
                    #----------------Crear la instancia en la base de datos---------------
                    conn=connection(self.ventana)
                    cursor=conn.cursor(dictionary=True)
                    cursor.execute("INSERT INTO vidrios (id_plano, unidades, tipo_borde, forma, ancho1, alto1, ancho2, alto2, saques, perforaciones, tipo_arenado, referencia, observacion, marca_etiquetado) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, NOW());", 
                                   (self.datosPlano['id'],
                                    aux['unidades'], 
                                    aux['tipo_borde'], 
                                    aux['forma'], 
                                    aux['ancho1'], 
                                    aux['alto1'], 
                                    aux['ancho2'], 
                                    aux['alto2'], 
                                    aux['saques'], 
                                    aux['perforaciones'], 
                                    aux['tipo_arenado'], 
                                    aux['referencia'], 
                                    aux['observacion']))
                    conn.commit()
                    conn.close()

                self.consultar_vidrios()
                self.lblAreaCalculada.configure(text=str(round(self.calcular_area(),2))+' m²')
            else:
                messagebox.showerror(title='Error en unidades', message='Las unidades de este vidrios excederían la cantidad de vidrios de este Plano')

    def editar_vidrio(self):
        #Se habilitan las entradas del formulario
        self.spboxUnidades.configure(state='enabled')
        self.cboxBorde.configure(state='enabled')
        self.cboxForma.configure(state='enabled')
        self.cbtnDescAncho.configure(state='enabled')
        self.cbtnDescAlto.configure(state='enabled')
        self.spboxAncho1.configure(state='enabled')
        self.spboxAncho2.configure(state='enabled')
        self.spboxAlto1.configure(state='enabled')
        self.spboxAlto2.configure(state='enabled')
        self.spboxSaques.configure(state='enabled')
        self.spboxPerforaciones.configure(state='enabled')
        self.cboxArenado.configure(state='enabled')
        self.txtRefencia.configure(state='enabled')
        self.txtObservacion.configure(state='enabled')

        #Se muestran los botones para actualizar y cancelar
        self.btnGuardar.grid_forget()
        self.btnEditar.grid_forget()
        self.btnImprimir.grid_forget()
        self.btnActualizar.grid(row=3, column=0, pady=5)
        self.btnCancelar.grid(row=4, column=0, pady=5)
    
    def actualizar_vidrio(self):
        if self.verificar_formulario():
            indice=self.tableInfoVidrios.index(self.tableInfoVidrios.selection()[0])

            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            cursor.execute('UPDATE vidrios SET unidades=%s, tipo_borde=%s, forma=%s, ancho1=%s, alto1=%s, ancho2=%s, alto2=%s, saques=%s, perforaciones=%s, tipo_arenado=%s, referencia=%s, observacion=%s WHERE id=%s;',
                        (int(self.spboxUnidades.get()),
                        self.cboxBorde.get(),
                        self.cboxForma.get(),
                        int(self.spboxAncho1.get()),
                        int(self.spboxAlto1.get()),
                        int(self.spboxAncho2.get()),
                        int(self.spboxAlto2.get()),
                        int(self.spboxSaques.get()),
                        int(self.spboxPerforaciones.get()),
                        self.cboxArenado.get(),
                        self.txtRefencia.get(),
                        self.txtObservacion.get(),
                        self.vidrios[indice]['id']))
            conn.commit()
            conn.close()
            
            self.consultar_vidrios()
            self.actualizar_lista(indice)

    def cancelar_edicion(self):
        if self.tableInfoVidrios.selection():
            indice=self.tableInfoVidrios.selection()[0]
            self.actualizar_lista(self.tableInfoVidrios.index(indice))

    def verificar_formulario(self)->bool:
        if not self.spboxUnidades.get().isdigit():
            messagebox.showerror(title='Entrada no válida', message='La casilla de las unidades del vidrio, debe ser un número entero.')
            return False
        if self.cboxBorde.current() == -1:
            messagebox.showerror(title='Entrada no válida', message='Debe seleccionar una de las opciones para el tipo de borde.')
            return False
        if self.cboxForma.current() == -1:
            messagebox.showerror(title='Entrada no válida', message='Debe seleccionar una de las opciones para la forma del vidrio.')
            return False
        
        if not self.spboxAncho1.get().isdigit():
            messagebox.showerror(title='Entrada no válida', message='La casilla del ancho del vidrio, debe ser un número entero.')
            return False
        if not self.spboxAlto1.get().isdigit():
            messagebox.showerror(title='Entrada no válida', message='La casilla del alto del vidrio, debe ser un número entero.')
            return False
        
        if self.descAncho.get():
            if not self.spboxAncho2.get().isdigit():
                messagebox.showerror(title='Entrada no válida', message='La casilla del ancho 2 del vidrio, debe ser un número entero.')
                return False
        if self.descAlto.get():
            if not self.spboxAlto2.get().isdigit():
                messagebox.showerror(title='Entrada no válida', message='La casilla del alto 2 del vidrio, debe ser un número entero.')
                return False
            
        if not self.spboxSaques.get().isdigit():
            messagebox.showerror(title='Entrada no válida', message='La casilla la cantidad de saques del vidrio, debe ser un número entero.')
            return False
        if not self.spboxPerforaciones.get().isdigit():
            messagebox.showerror(title='Entrada no válida', message='La casilla la cantidad de perforaciones del vidrio, debe ser un número entero.')
            return False
        
        if self.cboxArenado.current() == -1:
            messagebox.showerror(title='Entrada no válida', message='Debe seleccionar una de las opciones para el tipo de arenado.')
            return False
        
        return True
    
    def calcular_area(self)->Decimal:
        area=Decimal(0.00)
        for vidrio in self.datosVidrios:
            area+=Decimal(max(vidrio['ancho1'], vidrio['ancho2'])/1000)*Decimal(max(vidrio['alto1'], vidrio['alto2'])/1000)
        return area