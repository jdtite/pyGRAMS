import ttkbootstrap as tkb
from tkinter import messagebox
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
import os
import tempfile
from ftplib import FTP
import create_document
from connection import connection
from connection import ftpConnection

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
        self.btnBuscar=tkb.Button(self.frmBuscar,text="Buscar",command=self.consultar)
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
        #Crear y ubicar la etiqueta para el Área total del plano
        tkb.Label(self.frmInfo,text="Area V.:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=8,column=0,padx=5,pady=3,sticky='w')
        self.lblAreaVendida=tkb.Label(self.frmInfo,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblAreaVendida.grid(row=8,column=1,pady=3,sticky='w')

        #Crear y ubicar el marco para contener los botones.
        self.frmBotonesPlano = tkb.Frame(self.frmInfo)
        self.frmBotonesPlano.grid(row=9, column=0, columnspan=2, padx=5, pady=5)
        #Crear y ubicar el botón para ver la portada
        self.btnVerPortada=tkb.Button(self.frmBotonesPlano,text="Ver Portada",command=self.verPortada, state='disabled',bootstyle="disabled")
        self.btnVerPortada.pack(padx=5, pady=5, side='left')
        #Crear y ubicar el botón para ver el plano.
        self.btnVerPlano=tkb.Button(self.frmBotonesPlano,text="Ver Plano",command=self.verPlano, state='disabled',bootstyle="disabled")
        self.btnVerPlano.pack(padx=5, pady=5, side='left')
        #Crear y ubicar el botón para ver la portada
        self.btnExportarDatos=tkb.Button(self.frmBotonesPlano,text="Exportar",command=self.exportar_datos, state='disabled',bootstyle="info")
        self.btnExportarDatos.pack(padx=5, pady=5, side='left')

        #-----------------------------------------------------------------------------------------------
        #Define la segunda columna del contenedor
        self.colum2=tkb.Frame(self.contenedor)
        self.colum2.pack(padx=10, pady=10, side='left', fill='y')
        #Crear y ubicar el Frame para la información de los vidrios
        self.frmInfoVidrios=tkb.LabelFrame(self.colum2,text='Lista de vidrios')
        self.frmInfoVidrios.pack(padx=5, pady=5,ipadx=5,ipady=5)

        #Crear y ubicar el marco para contener la tabla y la barra de desplazamiento
        self.frmTablaVidrios = tkb.Frame(self.frmInfoVidrios, relief="flat")
        self.frmTablaVidrios.grid(row=0, column=0, columnspan=2, sticky='we')
        self.frmTablaVidrios.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmTablaVidrios.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmTablaVidrios, orient=VERTICAL)
        scrollbarV.grid(row=0,column=1, pady=5, sticky='ns')

        fuente=tkb.font.Font(family="times", size=11)
        style = tkb.Style()
        style.configure("mystyle.Treeview", font=('times', 11))

        #Crear y ubicar la tabla de los vidrios
        self.tableInfoVidrios=tkb.Treeview(self.frmTablaVidrios, show='headings', height=15, yscrollcommand=scrollbarV.set, style="mystyle.Treeview")
        self.tableInfoVidrios.configure(columns=('vidrio', 'medidas', 'etiquetado', 'marcado', 'almacenado', 'arenado'))
        self.tableInfoVidrios.column('vidrio', stretch=False, anchor='center', width=fuente.measure(' 0000000 '))
        self.tableInfoVidrios.column('medidas', stretch=False, anchor='center', width=fuente.measure(' 0000X0000 '))
        self.tableInfoVidrios.column('etiquetado', stretch=False, anchor='w', width=fuente.measure('0 0000-00-00 00:00:00 0'))
        self.tableInfoVidrios.column('marcado', stretch=False, anchor='w', width=fuente.measure('0 0000-00-00 00:00:00 0'))
        self.tableInfoVidrios.column('almacenado', stretch=False, anchor='w', width=fuente.measure('0 0000-00-00 00:00:00 0'))
        self.tableInfoVidrios.column('arenado', stretch=False, anchor='w', width=fuente.measure('0 0000-00-00 00:00:00 0'))
        
        for col in self.tableInfoVidrios['columns']:
            self.tableInfoVidrios.heading(col, text=col.title(), anchor=W)

        self.tableInfoVidrios.grid(row=0,column=0, padx=5, pady=5)
        self.tableInfoVidrios.tag_configure('listo', background='sky blue')
        self.tableInfoVidrios.bind("<<TreeviewSelect>>", self.lboxVidrios_change)

        scrollbarV.config(command=self.tableInfoVidrios.yview)
        
        #Crear y ubicar el marco para contener el el resumen de la información
        self.frmResumenInfo = tkb.Frame(self.frmInfoVidrios, relief="flat")
        self.frmResumenInfo.grid(row=1, column=0, sticky='we')

        #Crear y ubicar la etiqueta para la Forma
        tkb.Label(self.frmResumenInfo,text="Forma:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=1,column=0,padx=5,sticky='w')
        self.lblForma=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"),anchor="w", width=10)
        self.lblForma.grid(row=1,column=1,sticky='w')
        #Crear y ubicar la etiqueta para Arenado
        tkb.Label(self.frmResumenInfo,text="Aren:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=2,column=0,padx=5,sticky='w')
        self.lblArenado=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"),anchor="w")
        self.lblArenado.grid(row=2,column=1,sticky='w')
        #Crear y ubicar la etiqueta para las referencias
        tkb.Label(self.frmResumenInfo,text="Ref:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=3,column=0,padx=5,sticky='w')
        self.lblReferencia=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"),anchor="w")
        self.lblReferencia.grid(row=3,column=1,sticky='w')
        #Crear y ubicar la etiqueta para las observaciones
        tkb.Label(self.frmResumenInfo,text="Obs:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=4,column=0,padx=5,sticky='w')
        self.lblObservaciones=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"),anchor="w")
        self.lblObservaciones.grid(row=4,column=1,sticky='w')

        #Crear y ubicar el marco para contener los botones.
        self.lblAreaReal = tkb.Label(self.frmInfoVidrios, font=("Arial 20"))
        self.lblAreaReal.grid(row=1, column=1, columnspan=2, padx=5, pady=5)
        

        #-----------------------------------------------------------------------------

    def onEnter(self, event):
        self.consultar()

    def consultar(self):
        if  self.NoPlano.get() != '':
            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            cursor.execute("SELECT planos.*, empleados.nombre1, empleados.apellido1 FROM planos INNER JOIN empleados ON planos.id_empleado = empleados.id WHERE planos.id = %s;",(int(self.NoPlano.get()),))
            self.datosPlano=cursor.fetchone()
            conn.commit()

            if self.datosPlano is not None: #Si el plano existe
                if self.datosPlano['marca_etiquetado'] is not None: #Si el plano ha sido etiquetado
                    self.lblNoPlano.config(text=self.NoPlano.get())
                    self.NoPlano.set('')

                    self.lblCantidad.config(text=str(self.datosPlano['cantidad_vidrios']))

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

                    #Activar los botones
                    self.btnVerPortada.config(state='active',bootstyle="primary")
                    self.btnVerPlano.config(state='active',bootstyle="primary")
                    self.btnExportarDatos.config(state='active')

                    #Buscar y listar todos los vidrios--------------------------------------------------
                    cursor.execute("SELECT * FROM vidrios WHERE id_plano=%s;",(int(self.lblNoPlano.cget('text')),))
                    self.datosVidrios=cursor.fetchall()
                    conn.commit()
                    
                    if self.datosVidrios:
                        fila=[]
                        self.tableInfoVidrios.delete(*self.tableInfoVidrios.get_children())
                        for vidrio in self.datosVidrios:
                            fila.clear()
                            fila.append(str(vidrio['id']))
                            fila.append(str(vidrio['ancho1'])+ "x" + str(vidrio['alto1']))
                            fila.append(str(vidrio['marca_etiquetado']))
                            if vidrio['marca_marcado']:
                                fila.append(str(vidrio['marca_marcado']))
                            else:
                                fila.append('')
                            if vidrio['marca_almacenado']:
                                fila.append(str(vidrio['marca_almacenado']))
                            else:
                                fila.append('')
                            if vidrio['marca_arenado']:
                                fila.append(str(vidrio['marca_arenado']))
                            else:
                                fila.append('')
                            #Agregar la fila y resaltarla si el vidrio está listo
                            if vidrio['tipo_arenado']!='Ninguno':
                                if vidrio['marca_almacenado'] and vidrio['marca_arenado']:
                                    self.tableInfoVidrios.insert('', index='end', values=fila, tags=('listo',))
                                else:
                                    self.tableInfoVidrios.insert('', index='end', values=fila)
                            else:
                                if vidrio['marca_almacenado']:
                                    self.tableInfoVidrios.insert('', index='end', values=fila, tags=('listo',))
                                else:
                                    self.tableInfoVidrios.insert('', index='end', values=fila)
                            #self.tableInfoVidrios.insert('', index='end', values=fila)
                        #----------------Consultar el área total del plano--------------------------------------------------
                        cursor.execute("SELECT SUM(vidrios.area) FROM vidrios WHERE id_plano=%s;",(int(self.lblNoPlano.cget('text')),))
                        aux=cursor.fetchone()
                        conn.commit()
                        area_total=aux['SUM(vidrios.area)']
                        self.lblAreaReal.configure(text=str(area_total) + ' m²')
                else:
                    messagebox.showerror(title='Plano no etiquetado', message='Su plano aún no inicia el proceso.')
            else:
                messagebox.showerror(title='Plano no encontrado', message='Su plano no se encuentra registrado.')
            conn.close()
        else:
            messagebox.showerror(title='Plano no definido', message='No sea estúpido, la casilla del número de plano no puede quedar vacía.')
    
    def verPortada(self):
        if self.datosPlano['id_materia_prima']:
            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(self.datosPlano['id_materia_prima'],))
            materia_prima=cursor.fetchone()
            conn.commit()
            conn.close()
            materia=materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm'

        else:
            materia=self.datosPlano['color'] + ' ' + str(self.datosPlano['espesor']) + ' mm'

        create_document.Cover(datosPlano=self.datosPlano,materia=materia)

    def exportar_datos(self):
        if self.datosPlano['id_materia_prima']:
            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(self.datosPlano['id_materia_prima'],))
            materia_prima=cursor.fetchone()
            conn.commit()
            conn.close()

            color=materia_prima['id_color']
            espesor=materia_prima['id_espesor']

        else:
            color=self.datosPlano['color']
            espesor=self.datosPlano['espesor']

        filename = os.path.join(tempfile.gettempdir(), "plano" + str(self.datosPlano['id']) +".csv")
        file = open(filename, "w")
        file.write('Fecha;Templado;No. Plano;No. Vidrio;Color;Espesor;Ancho1;Ancho2;Alto1;Alto2;Referencia;Observacion;Cliente\n')
        for vidrio in self.datosVidrios:
            file.write(self.datosPlano['marca_etiquetado'].strftime('%d/%m/%Y')+';'
                    +self.datosPlano['tipo_templado']+';'
                    +str(self.datosPlano['id'])+';'
                    +str(vidrio['id'])+';'
                    +color+';'
                    +str(espesor)+';'
                    +str(vidrio['ancho1'])+';'
                    +str(vidrio['ancho2'])+';'
                    +str(vidrio['alto1'])+';'
                    +str(vidrio['alto2'])+';'
                    +vidrio['referencia']+';'
                    +vidrio['observacion']+';'
                    +self.datosPlano['cliente']+'\n')
        file.close()
        os.startfile(filename)

    def verPlano(self):
        ftp_conn = ftpConnection()
        if ftp_conn is not None:
            try:
                # Crear un archivo vacío
                with open(os.path.join(tempfile.gettempdir(), f"plano{self.lblNoPlano.cget('text')}.pdf"), "wb") as fp:
                    # Transferir el archivo del servidor FTP
                    ftp_conn.retrbinary(f"RETR p{self.lblNoPlano.cget('text')}.pdf", fp.write)
                #Abrir el archivo
                os.startfile(os.path.join(tempfile.gettempdir(), f"plano{self.lblNoPlano.cget('text')}.pdf"))

            except:
                messagebox.showerror(title='Error de lectura', message='Plano no encontrado')
            ftp_conn.close()

    def lboxVidrios_change(self, Event):
        if self.tableInfoVidrios.selection():
            item=self.tableInfoVidrios.selection()[0]
            indice=self.tableInfoVidrios.index(item)

            self.lblForma.config(text=self.datosVidrios[indice]['forma'])
            self.lblArenado.config(text=self.datosVidrios[indice]['tipo_arenado'])
            self.lblReferencia.config(text=self.datosVidrios[indice]['referencia'])
            self.lblObservaciones.config(text=self.datosVidrios[indice]['observacion'])
