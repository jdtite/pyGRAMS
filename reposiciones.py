import ttkbootstrap as tkb
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
from datetime import datetime
import os
from decimal import Decimal
import create_document
from collections import defaultdict
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends._backend_tk import NavigationToolbar2Tk
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

        self.id_empleado=id_empleado
        self.usuario=usuario
        
        #Define el Frame como contenedor total que se asocia al notebook
        self.contenedor=ScrolledFrame(self)
        self.contenedor.pack(padx=10,pady=10, fill='both', expand=True)

        #-------------------Crear y ubicar el Frame para ver la lista de reposiciones creados---------------------------
        #Define la primera columna del contenedor
        self.colum1=tkb.Frame(self.contenedor)
        self.colum1.pack(padx=10, pady=10, side='left', fill='y')

        self.frmReposicionesCreadas=tkb.LabelFrame(self.colum1,text='Lista de reposiciones creadas', bootstyle="primary")
        self.frmReposicionesCreadas.pack(padx=5, pady=5, ipadx=10, ipady=10)
        # Crear y ubicar la entrada de fecha de inicio para la búsqueda
        tkb.Label(self.frmReposicionesCreadas,text="Desde:",font=("Arial 10 bold")).grid(row=0,column=0, padx=5,pady=3,sticky='e')
        self.dtDesde=tkb.DateEntry(self.frmReposicionesCreadas, dateformat="%Y-%m-%d", width=12)
        self.dtDesde.grid(row=0,column=1,pady=3,sticky='w')
        # Crear y ubicar la entrada de fecha de fin para la búsqueda
        tkb.Label(self.frmReposicionesCreadas,text="Hasta:",font=("Arial 10 bold")).grid(row=0,column=2, padx=5,pady=3,sticky='e')
        self.dtHasta=tkb.DateEntry(self.frmReposicionesCreadas, dateformat="%Y-%m-%d", width=12)
        self.dtHasta.grid(row=0,column=3,pady=3,sticky='w')
        #Crear y ubicar el botón para actualizar los valores
        self.btnActualizar=tkb.Button(self.frmReposicionesCreadas,text="Actualizar",command=self.consultar_reposiciones_creadas)
        self.btnActualizar.grid(row=0, column=4, padx=5, pady=5)

        #Crear y ubicar el marco para contener el listbox y la barra de desplazamiento
        self.frmListaReposiciones = tkb.Frame(self.frmReposicionesCreadas, relief="flat")
        self.frmListaReposiciones.grid(row=1, column=0, columnspan=5, sticky='we')
        self.frmListaReposiciones.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmListaReposiciones.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmListaReposiciones, orient=VERTICAL)
        scrollbarV.grid(row=0,column=1, pady=5, sticky='ns')

        fuente=tkb.font.Font(family="times", size=11)
        style = tkb.Style()
        style.configure("mystyle.Treeview", font=('times', 11))
        
        #Crear y ubicar la tabla de los planos creados
        self.tableReposicionesCreadas=tkb.Treeview(self.frmListaReposiciones, show='headings', height=10, yscrollcommand=scrollbarV.set, style="mystyle.Treeview")
        self.tableReposicionesCreadas.configure(columns=('no-rep', 'vidrio', 'fecha-hora', 'estacion', 'vendedor'))
        self.tableReposicionesCreadas.column('no-rep', stretch=False, anchor='w', width=fuente.measure(' 0000000 '))
        self.tableReposicionesCreadas.column('vidrio', stretch=False, anchor='w', width=fuente.measure(' 0000000 '))
        self.tableReposicionesCreadas.column('fecha-hora', stretch=False, anchor='w', width=fuente.measure('0 0000-00-00 00:00:00 0'))
        self.tableReposicionesCreadas.column('estacion', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxx'))
        self.tableReposicionesCreadas.column('vendedor', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxxxxxx'))

        for col in self.tableReposicionesCreadas['columns']:
            self.tableReposicionesCreadas.heading(col, text=col.title(), anchor=W)

        self.tableReposicionesCreadas.grid(row=0,column=0, padx=5, pady=5)
        self.tableReposicionesCreadas.bind("<Double-Button-1>", self.ver_reposicion)

        scrollbarV.config(command=self.tableReposicionesCreadas.yview)

        #Crear y ubicar el marco para contener el resumen de la información
        self.frmResumenInfo = tkb.Frame(self.frmReposicionesCreadas, relief="flat")
        self.frmResumenInfo.grid(row=2, column=0, columnspan=4, sticky='we')
        #Crear y ubicar la etiqueta para la cantidad de vidrios repuestos
        tkb.Label(self.frmResumenInfo,text="Vidrios repuestos:",font=("Arial 10 bold"), justify="left").grid(row=0,column=0,padx=5,sticky='w')
        self.lblVidriosRepuestos=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblVidriosRepuestos.grid(row=0,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de área original
        tkb.Label(self.frmResumenInfo,text="Área original:",font=("Arial 10 bold"), justify="left").grid(row=1,column=0,padx=5,sticky='w')
        self.lblAreaOriginal=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblAreaOriginal.grid(row=1,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de área recuperada
        tkb.Label(self.frmResumenInfo,text="Área recuperada:",font=("Arial 10 bold"), justify="left").grid(row=2,column=0,padx=5,sticky='w')
        self.lblAreaRecuperada=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblAreaRecuperada.grid(row=2,column=1,sticky='w')

        #Crear y ubicar el marco para contener los botones.
        self.frmBotonesLista = tkb.Frame(self.frmReposicionesCreadas)
        self.frmBotonesLista.grid(row=2, column=4, padx=5, pady=5)
        #Crear el botón para Crear una nueva reposición.
        self.btnNuevo=tkb.Button(self.frmBotonesLista,text="Nueva", width=10,command=self.nueva_reposicion, bootstyle="success")
        self.btnNuevo.pack(padx=5, pady=5, side='left')

        #-------------------Crear y ubicar el Frame para ver la gráfica de reposiciones creados---------------------------
        #Define la segunda columna del contenedor
        self.colum2=tkb.Frame(self.contenedor)
        self.colum2.pack(padx=10, pady=10, side='left', fill='y')

        self.frmBarchart=tkb.LabelFrame(self.colum2, text='Reposiciones por Estación de trabajo', bootstyle="primary")
        self.frmBarchart.pack(padx=5, pady=5)

        self.fig1, self.ax1 = plt.subplots()
        self.fig1.set_dpi(68)
        self.fig1.set_size_inches(5,8.2)
        self.fig1.subplots_adjust(left=0.23, right=0.95, top=0.98, bottom=0.08)
        #self.fig, self.axs = plt.subplots(ncols=2)
        #self.ax1=self.fig1.add_axes([0.21, 0.07, 0.74, 0.88])
        self.canvas1 = FigureCanvasTkAgg(self.fig1, master = self.frmBarchart)
        self.canvas1.get_tk_widget().pack()
        
        self.toolbar1 = NavigationToolbar2Tk(self.canvas1, self.frmBarchart)
        self.toolbar1._message_label.pack_forget()
        self.toolbar1._buttons['Back'].pack_forget()
        self.toolbar1._buttons['Forward'].pack_forget()
        self.toolbar1.pack()
 
        #-------------------------------------------------------------------------------------
        #Define la tercera columna del contenedor
        self.colum3=tkb.Frame(self.contenedor)
        self.colum3.pack(padx=10, pady=10, side='left', fill='y')

        self.frmPiechart=tkb.LabelFrame(self.colum3,text='Porcentaje de incidencias', bootstyle="primary")
        self.frmPiechart.pack(padx=5, pady=5)

        self.fig2, self.ax2 = plt.subplots()
        self.fig2.set_dpi(70)
        self.fig2.set_size_inches(4.5,3.2)
        self.fig2.subplots_adjust(left=0.1, right=0.9, bottom=0, top=1)
        self.ax2.set_axis_off()

        self.canvas2 = FigureCanvasTkAgg(self.fig2, master = self.frmPiechart)
        self.canvas2.get_tk_widget().pack()
        
        self.toolbar2 = NavigationToolbar2Tk(self.canvas2, self.frmPiechart)
        self.toolbar2._message_label.pack_forget()
        self.toolbar2._buttons['Back'].pack_forget()
        self.toolbar2._buttons['Forward'].pack_forget()
        self.toolbar2.pack()
        #-------------------------------------------------------------------------------------

        self.frmStemchart=tkb.LabelFrame(self.colum1,text='Incidencias por día', bootstyle="primary")
        self.frmStemchart.pack(padx=5, pady=5)

        self.fig3, self.ax3 = plt.subplots()
        self.fig3.set_dpi(70)
        self.fig3.set_size_inches(7.5,3.2)
        self.fig3.subplots_adjust(left=0.07, right=0.95, bottom=0.3, top=0.95)

        self.canvas3 = FigureCanvasTkAgg(self.fig3, master = self.frmStemchart)
        self.canvas3.get_tk_widget().pack()

        self.toolbar3 = NavigationToolbar2Tk(self.canvas3, self.frmStemchart)
        self.toolbar3._message_label.pack_forget()
        self.toolbar3._buttons['Back'].pack_forget()
        self.toolbar3._buttons['Forward'].pack_forget()
        self.toolbar3.pack()

#---------------------------------------------------------------------------------------
    def nueva_reposicion(self):
        #Abrir la ventana secundaria con el formulario
        FormNuevaReposicion(ventana=self.ventana, register=self.registro, id_empleado=self.id_empleado, usuario=self.usuario, callback=self.consultar_reposiciones_creadas, callback2=self.ver_reposicion)
    
    def consultar_reposiciones_creadas(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT reposiciones.*, vidrios.id_plano, vidrios.ancho1, vidrios.alto1,  vidrios.ancho2, vidrios.alto2, planos.cliente, planos.id_materia_prima, planos.color, planos.espesor, empleados.nombre1, empleados.apellido1 FROM reposiciones INNER JOIN vidrios ON reposiciones.id_vidrio=vidrios.id INNER JOIN planos ON vidrios.id_plano=planos.id INNER JOIN empleados ON planos.id_empleado=empleados.id WHERE reposiciones.marca_creacion BETWEEN %s AND %s;",(self.dtDesde.entry.get() + " 00:00:00",self.dtHasta.entry.get() + " 23:59:59"))
        self.ReposicionesCreadas=cursor.fetchall()
        conn.commit()
        conn.close()

        fila=[]
        self.tableReposicionesCreadas.delete(*self.tableReposicionesCreadas.get_children())
        self.area_recuperada=Decimal(0.00)
        self.area_original=Decimal(0.00)
        for reposicion in self.ReposicionesCreadas:
            fila.clear()
            fila.append(str(reposicion['id']))
            fila.append(str(reposicion['id_vidrio']))
            fila.append(str(reposicion['marca_creacion']))
            fila.append(reposicion['estacion_trabajo'])
            fila.append(reposicion['nombre1'] + ' ' + reposicion['apellido1'])

            self.area_recuperada=self.area_recuperada+(Decimal(reposicion['ancho_recuperable']/1000)*Decimal(reposicion['alto_recuperable']/1000))

            if reposicion['ancho1'] > reposicion['ancho2']:
                if reposicion['alto1'] > reposicion['alto2']:
                    self.area_original=self.area_original+(Decimal(reposicion['ancho1']/1000)*Decimal(reposicion['alto1']/1000))
                else:
                    self.area_original=self.area_original+(Decimal(reposicion['ancho1']/1000)*Decimal(reposicion['alto2']/1000))
            else:
                if reposicion['alto1'] > reposicion['alto2']:
                    self.area_original=self.area_original+(Decimal(reposicion['ancho2']/1000)*Decimal(reposicion['alto1']/1000))
                else:
                    self.area_original=self.area_original+(Decimal(reposicion['ancho2']/1000)*Decimal(reposicion['alto2']/1000))

            self.tableReposicionesCreadas.insert('', 0, values=fila)

        self.lblVidriosRepuestos.config(text=str(len(self.ReposicionesCreadas)))
        self.lblAreaRecuperada.config(text=str(round(self.area_recuperada,2)) + ' m²')
        self.lblAreaOriginal.config(text=str(round(self.area_original,2)) + ' m²')

        self.graficar_reposiciones()

    def graficar_reposiciones(self):
        #Limpiar gráfica
        self.ax1.clear()
        self.ax2.clear()
        self.ax3.clear()

        datosAreaOriginal=defaultdict(Decimal)
        datosAreaRecuperable=defaultdict(Decimal)
        
        for reposicion in self.ReposicionesCreadas:
            if reposicion['ancho1'] > reposicion['ancho2']:
                if reposicion['alto1'] > reposicion['alto2']:
                    datosAreaOriginal[reposicion['estacion_trabajo']] += (Decimal(reposicion['ancho1']/1000)*Decimal(reposicion['alto1']/1000))
                else:
                    datosAreaOriginal[reposicion['estacion_trabajo']] += (Decimal(reposicion['ancho1']/1000)*Decimal(reposicion['alto2']/1000))
            else:
                if reposicion['alto1'] > reposicion['alto2']:
                    datosAreaOriginal[reposicion['estacion_trabajo']] += (Decimal(reposicion['ancho2']/1000)*Decimal(reposicion['alto1']/1000))
                else:
                    datosAreaOriginal[reposicion['estacion_trabajo']] += (Decimal(reposicion['ancho2']/1000)*Decimal(reposicion['alto2']/1000))
            datosAreaRecuperable[reposicion['estacion_trabajo']] += Decimal(reposicion['ancho_recuperable']/1000)*Decimal(reposicion['alto_recuperable']/1000)

        datosConjuntos=dict()
        for estacion in datosAreaOriginal:
            datosAreaOriginal[estacion]=round(datosAreaOriginal[estacion],2)
            datosAreaRecuperable[estacion]=round(datosAreaRecuperable[estacion],2)
            datosConjuntos[estacion]=(datosAreaOriginal[estacion],datosAreaRecuperable[estacion])

        listaAux=list(datosConjuntos.items())
        listaAux.sort(key=lambda item: item[1][0], reverse=False)
        datosOrdenados=dict(listaAux)

        self.datosIncidencias=defaultdict(int)
        for reposicion in self.ReposicionesCreadas:
            self.datosIncidencias[str(reposicion['marca_creacion'])[:10]]+=1
        #print(datosIncidencias)

        #Gráfica de barras----------------------
        ancho=0.3
        self.ax1.barh(list(datosOrdenados.keys()), [x[0] for x in list(datosOrdenados.values())], height=ancho, edgecolor="black", align='edge', color='coral', label='Área original')
        self.ax1.barh(list(datosOrdenados.keys()), [x[1] for x in list(datosOrdenados.values())], height=-1*ancho, edgecolor="black", align='edge', color='SkyBlue', label='Área recuperable')
        self.ax1.set_xlim(xmin=0,xmax=max(datosAreaOriginal.values())*Decimal(1.25))

        for item in range(len(datosOrdenados)):
            self.ax1.text(x=listaAux[item][1][0]+max(datosAreaOriginal.values())*Decimal(0.01), y=item+(ancho/2), s=str(listaAux[item][1][0]) + ' m²', horizontalalignment='left', verticalalignment='center')
            self.ax1.text(x=listaAux[item][1][1]+max(datosAreaOriginal.values())*Decimal(0.01), y=item-(ancho/2), s=str(listaAux[item][1][1]) + ' m²', horizontalalignment='left', verticalalignment='center')

        self.ax1.set_ylabel('Estación de trabajo',fontsize='large')
        self.ax1.set_xlabel('Área [m²]',fontsize='large')
        
        self.ax1.legend()
        #self.ax2.legend()

        self.toolbar1.update()
        self.canvas1.draw()

        #Gráfica de pastel----------------------
        self.ax2.pie([x[0] for x in list(datosOrdenados.values())], labels=list(datosOrdenados.keys()), autopct = '%1.1f%%', wedgeprops=dict(width=0.5, edgecolor='black'), pctdistance = 0.75)
        self.toolbar2.update()
        self.canvas2.draw()

        #Gráfica de Stem----------------------
        self.ax3.stem(list(self.datosIncidencias.keys()), list(self.datosIncidencias.values()), label='Incidencias')
        self.fig3.autofmt_xdate(rotation=90,ha='center',bottom=0.3)

        self.toolbar3.update()
        self.canvas3.draw()

    def ver_reposicion(self, Event=None, id_reposicion=None):
        if id_reposicion is None:
            indice=self.tableReposicionesCreadas.selection()[0]
            id_reposicion=self.tableReposicionesCreadas.item(indice,'values')[0]
        
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT reposiciones.*, vidrios.id_plano, vidrios.ancho1, vidrios.alto1,  vidrios.ancho2, vidrios.alto2, planos.cliente, planos.id_materia_prima, planos.color, planos.espesor, empleados.nombre1, empleados.apellido1 FROM reposiciones INNER JOIN vidrios ON reposiciones.id_vidrio=vidrios.id INNER JOIN planos ON vidrios.id_plano=planos.id INNER JOIN empleados ON planos.id_empleado=empleados.id WHERE reposiciones.id = %s;",(id_reposicion,))
        datosReposicion=cursor.fetchone()
        conn.commit()

        if datosReposicion['id_materia_prima']:
            cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(datosReposicion['id_materia_prima'],))
            materia_prima=cursor.fetchone()
            conn.commit()
            materia=materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm'
        else:
            materia=datosReposicion['color'] + ' ' + str(datosReposicion['espesor']) + ' mm'
        conn.close()
        create_document.Replacement(datosReposicion=datosReposicion,materia=materia)

class FormNuevaReposicion(tkb.Toplevel):   
    def __init__(self, *args, ventana:tkb.Window, register:tkb.Window.register, id_empleado:int, usuario:str, callback, callback2, **kwargs):
        super().__init__(*args, **kwargs)
        #instanciar cursor
        self.ventana=ventana
        self.registro=register

        self.usuario=usuario
        self.id_empleado=id_empleado
        self.callback=callback
        self.callback2=callback2

        self.title('Formulario de reposición')
        self.resizable(False,False)
        self.attributes('-topmost', 1)
        self.attributes('-toolwindow', True)
        self.iconbitmap(os.path.join(carpeta_imagenes, "igrams.ico"))

        #Define el Frame como contenedor general
        self.contenedor=tkb.Frame(self)
        self.contenedor.pack(padx=10,pady=10,side='left', fill='y')

        #--------------------------------------------------------------------------------------------

        #Define la primera columna del contenedor
        self.colum1=tkb.Frame(self.contenedor)
        self.colum1.pack(padx=10, pady=10, side='left', fill='y')

        #Crear y ubicar el Frame para la búsqueda
        self.frmBuscar=tkb.Labelframe(self.colum1, text='Vidrio No:', relief='flat')
        self.frmBuscar.pack(padx=5, pady=5,  fill='x')
        #Crear y ubicar el cuadro de texto para buscar por No Plano
        self.NoVidrio=tkb.StringVar()
        self.txtNoPlano=tkb.Entry(self.frmBuscar, width=10,font=("Times 12"),textvariable=self.NoVidrio,validate="key",validatecommand=(self.registro(lambda text: text.isdigit()), "%S"))
        self.txtNoPlano.pack(padx=5, pady=5, side='left')
        self.txtNoPlano.bind("<Return>", self.onEnter)
        self.txtNoPlano.focus()
        #Crear y ubicar el botón para buscar
        self.btnBuscarPlano=tkb.Button(self.frmBuscar,text="Buscar",command=self.buscar_vidrio)
        self.btnBuscarPlano.pack(padx=5, pady=5,side='left')

        #Crear y ubicar el Frame para la información del vidrio
        self.frmInfoPlano=tkb.LabelFrame(self.colum1,text='Información del vidrio',bootstyle="primary")
        self.frmInfoPlano.pack(padx=5, pady=5, ipadx=10, ipady=10, fill='x')
        #Crear y ubicar la etiqueta para el No. de Plano
        tkb.Label(self.frmInfoPlano,text="Plano No:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=0,column=0, padx=5,pady=3,sticky='w')
        self.lblNoPlano=tkb.Label(self.frmInfoPlano,text="",font=("Arial 20 bold"),anchor="w",width=6)
        self.lblNoPlano.grid(row=0,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de vidrios
        tkb.Label(self.frmInfoPlano,text="Vidrio No:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=1,column=0, padx=5,pady=3,sticky='w')
        self.lblNoVidrio=tkb.Label(self.frmInfoPlano, width=10, font=("Arial 14 bold"))
        self.lblNoVidrio.grid(row=1,column=1,pady=3,sticky='w')
        #Crear y ubicar la etiqueta para las medidas del vidrio
        tkb.Label(self.frmInfoPlano,text="Medidas:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=2,column=0,padx=5,pady=3,sticky='we')
        self.lblMedidas=tkb.Label(self.frmInfoPlano,font=("Arial 10"))
        self.lblMedidas.grid(row=2,column=1,pady=3,sticky='we')
        #Crear y ubicar la etiqueta para la materia prima
        tkb.Label(self.frmInfoPlano,text="Materia:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=3,column=0,padx=5,pady=3,sticky='we')
        self.lblMateria=tkb.Label(self.frmInfoPlano, font=("Arial 10"))
        self.lblMateria.grid(row=3,column=1,pady=5,sticky='we')
        #Crear y ubicar la etiqueta para el Cliente
        tkb.Label(self.frmInfoPlano,text="Cliente:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=4,column=0,padx=5,pady=3,sticky='we')
        self.lblCliente=tkb.Label(self.frmInfoPlano, font=("Arial 10"))
        self.lblCliente.grid(row=4,column=1,pady=3,sticky='we')
        #Crear y ubicar la etiqueta para el Vendedor
        tkb.Label(self.frmInfoPlano,text="Vendedor:",font=("Arial 10 bold"),anchor="w", justify="left").grid(row=5,column=0,padx=5,pady=3,sticky='we')
        self.lblVendedor=tkb.Label(self.frmInfoPlano,text="",font=("Arial 10"),anchor="w", justify="left")
        self.lblVendedor.grid(row=5,column=1,pady=3,sticky='we')

        #--------------------------------------------------------------------------------------------

        #Define la tercera columna del contenedor
        self.colum2=tkb.Frame(self.contenedor)
        self.colum2.pack(padx=10, pady=10, side='left', fill='y')

        #Crear y ubicar el Frame para la información del incidente
        self.frmInfoIncidente=tkb.LabelFrame(self.colum2,text='Información del incidente',bootstyle="primary")
        self.frmInfoIncidente.pack(padx=5, pady=10, ipadx=10,  fill='x')
        
        #Crear y ubicar la entrada para la estación de trabajo involucrada
        tkb.Label(self.frmInfoIncidente,text="Estación:",font=("Arial 10 bold"), justify="left").grid(row=0,column=0,padx=3,pady=3,sticky='w')
        self.cboxEstacion=tkb.Combobox(self.frmInfoIncidente,font=("Arial 10"),width=10,values=['Corte','Rectilinea', 'Pulpo', 'Marcado', 'Perforado', 'Saques', 'Lavadora', 'Templado', 'Despacho', 'Arenado', 'Caballete'])
        self.cboxEstacion.grid(row=0,column=1,pady=3,sticky='w')
        self.cboxEstacion.current(0)

        #Crear y ubicar el switch para asignar otra estación
        self.OtraEstacion=tkb.BooleanVar(value=False)
        self.cbtnOtraEstacion=tkb.Checkbutton(self.frmInfoIncidente, text='Otra', variable=self.OtraEstacion, command=self.alternar_estacion)
        self.cbtnOtraEstacion.grid(row=0,column=2, padx=5, pady=5)
        self.txtOtraEstacion=tkb.Entry(self.frmInfoIncidente, width=10, font=("Arial 10"))

        #Crear y ubicar la entrada para el motivo
        tkb.Label(self.frmInfoIncidente,text="Motivo:",font=("Arial 10 bold"), justify="left").grid(row=1,column=0,padx=5,pady=3,sticky='w')
        self.txtMotivo=tkb.Entry(self.frmInfoIncidente, font=("Arial 10"))
        self.txtMotivo.grid(row=1,column=1,columnspan=3,pady=5,sticky='we')
        #Crear y ubicar la entrada para la observación
        tkb.Label(self.frmInfoIncidente,text="Observación:",font=("Arial 10 bold"), justify="left").grid(row=2,column=0,padx=5,pady=3,sticky='we')
        self.txtObservacion=tkb.Entry(self.frmInfoIncidente, font=("Arial 10"))
        self.txtObservacion.grid(row=2,column=1,columnspan=3,pady=5,sticky='we')

        tkb.Label(self.frmInfoIncidente,text="Medidas recuperables",font=("Arial 10 bold"),justify="left").grid(row=3,column=0,padx=5,pady=3,columnspan=3,sticky='we')
        #Crear y ubicar la entrada de texto para el Ancho recuperable
        tkb.Label(self.frmInfoIncidente,text="Ancho:",font=("Arial 10 bold")).grid(row=4,column=0,padx=5,pady=3,sticky='e')
        self.spboxAnchoRecuperable=tkb.Spinbox(self.frmInfoIncidente, width=6, font=("Arial 10"),from_=0,to=3660,increment=1,validate="key",validatecommand=(self.registro(lambda text: text.isdigit()), "%S"))
        self.spboxAnchoRecuperable.grid(row=4,column=1,pady=5,sticky='w')
        self.spboxAnchoRecuperable.set(0)
        #Crear y ubicar la entrada de texto para el Alto recuperable
        tkb.Label(self.frmInfoIncidente,text="Alto:", font=("Arial 10 bold")).grid(row=4, column=2, padx=5, pady=3, sticky='e')
        self.spboxAltoRecuperable=tkb.Spinbox(self.frmInfoIncidente, width=6, font=("Arial 10"),from_=0,to=3660,increment=1,validate="key",validatecommand=(self.registro(lambda text: text.isdigit()), "%S"))
        self.spboxAltoRecuperable.grid(row=4,column=3,pady=5,sticky='w')
        self.spboxAltoRecuperable.set(0)
        #Crear y ubicar la entrada para el nombre del responsable
        tkb.Label(self.frmInfoIncidente,text="Responsable:",font=("Arial 10 bold"), justify="left").grid(row=5,column=0,padx=5,pady=3,sticky='we')
        self.txtResponsable=tkb.Entry(self.frmInfoIncidente, font=("Arial 10"))
        self.txtResponsable.grid(row=5,column=1,columnspan=3,pady=5,sticky='we')

        #Crear y ubicar el botón para crear la resposicion
        self.btnCrearReposicion=tkb.Button(self.colum2,text="Crear", command=self.crear_reposicion, width=25, bootstyle="success")
        self.btnCrearReposicion.pack(padx=3,pady=3)

    def onEnter(self, event):
        self.buscar_vidrio()

    def buscar_vidrio(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT id_plano, ancho1, alto1, ancho2, alto2, forma FROM vidrios WHERE id = %s;",(int(self.NoVidrio.get()),))
        self.datosVidrio=cursor.fetchone()
        conn.commit()

        if self.datosVidrio:
            self.lblNoPlano.config(text=str(self.datosVidrio['id_plano']))
            self.lblNoVidrio.config(text=self.NoVidrio.get())
            if self.datosVidrio['ancho1'] > self.datosVidrio['ancho2']:
                if self.datosVidrio['alto1'] > self.datosVidrio['alto2']:
                    self.lblMedidas.config(text="("+str(self.datosVidrio['ancho1'])+" x "+str(self.datosVidrio['alto1'])+")")
                else:
                    self.lblMedidas.config(text="("+str(self.datosVidrio['ancho1'])+" x "+str(self.datosVidrio['alto2'])+")")
            else:
                if self.datosVidrio['alto1'] > self.datosVidrio['alto2']:
                    self.lblMedidas.config(text="("+str(self.datosVidrio['ancho2'])+" x "+str(self.datosVidrio['alto1'])+")")
                else:
                    self.lblMedidas.config(text="("+str(self.datosVidrio['ancho2'])+" x "+str(self.datosVidrio['alto2'])+")")
            
            cursor.execute("SELECT id_materia_prima, color, espesor, cliente, id_empleado FROM planos WHERE id = %s;",(self.datosVidrio['id_plano'],))
            datosPlano=cursor.fetchone()
            conn.commit()

            if datosPlano['id_materia_prima']:
                cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(datosPlano['id_materia_prima'],))
                materia_prima=cursor.fetchone()
                conn.commit()
                self.lblMateria.config(text=materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')
            else:
                self.lblMateria.config(text=datosPlano['color'] + ' ' + str(datosPlano['espesor']) + ' mm')

            self.lblCliente.config(text=datosPlano['cliente'])

            cursor.execute("SELECT nombre1, apellido1 FROM empleados WHERE id = %s;",(datosPlano['id_empleado'],))
            empleado=cursor.fetchone()
            conn.commit()
            self.lblVendedor.config(text=empleado['nombre1']+ ' ' + empleado['apellido1'])

            self.NoVidrio.set('')
        else:
            Messagebox.show_info(title='Búsqueda finalizada', message='Vidrio no encontrado', parent=self)
        conn.close

    def verificar_formulario(self)->bool:
        if self.OtraEstacion.get():
            if not bool(self.txtOtraEstacion.get()):
                Messagebox.show_error(title='Entrada no válida', message='La casilla de la estación de trabajo.', parent=self)
                return False
        else:
            if self.cboxEstacion.current() == -1:
                Messagebox.show_error(title='Entrada no válida', message='Debe seleccionar una de las estaciones de trabajo.', parent=self)
                return False
        
        if not bool(self.txtMotivo.get()):
            Messagebox.show_error(title='Entrada no válida', message='Debe especificar el motivo del incidente.', parent=self)
            return False
        
        if not self.spboxAnchoRecuperable.get().isdigit():
            Messagebox.show_error(title='Entrada no válida', message='La casilla del ancho recuperable debe ser un número entero.', parent=self)
            return False
        if not self.spboxAltoRecuperable.get().isdigit():
            Messagebox.show_error(title='Entrada no válida', message='La casilla del alto recuperable debe ser un número entero.', parent=self)
            return False
        return True

    def crear_reposicion(self):
        if self.verificar_formulario():
            if self.OtraEstacion.get():
                estacion=self.txtOtraEstacion.get()
            else:
                estacion=self.cboxEstacion.get()
            crear = Messagebox.yesno(message="¿Está seguro que desea crear una nueva reposición?", title="Confirmar creación", parent=self)
            if crear=='Yes':
                conn=connection(self.ventana)
                cursor=conn.cursor(dictionary=True)
                cursor.execute("INSERT INTO reposiciones (id_vidrio, marca_creacion, estacion_trabajo, motivo, responsable, ancho_recuperable, alto_recuperable, observacion) VALUES (%s,%s,%s,%s,%s,%s,%s,%s);", (int(self.lblNoVidrio.cget("text")), datetime.now().strftime("%Y-%m-%d %H:%M:%S"), estacion, self.txtMotivo.get(), self.txtResponsable.get(), self.spboxAnchoRecuperable.get(), self.spboxAltoRecuperable.get(), self.txtObservacion.get()))
                cursor.execute("COMMIT;")

                cursor.execute("UPDATE vidrios SET marca_pulido=NULL, marca_marcado=NULL, marca_templado=NULL, marca_almacenado=NULL, marca_despacho=NULL WHERE id=%s;",(int(self.lblNoVidrio.cget("text")),))
                conn.commit()
                conn.close()

            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            self.limpiar_formulario()
            cursor.execute("SELECT LAST_INSERT_ID();")
            aux=cursor.fetchone()
            conn.commit()
            conn.close()
            
            self.callback()
            self.callback2(id_reposicion=aux['LAST_INSERT_ID()'])
            self.destroy()
        
    def limpiar_formulario(self):
        #Limpiar el formulario
        self.OtraEstacion.set(False)
        self.cboxEstacion.current(0)
        self.txtOtraEstacion.delete(0, END)

        self.txtMotivo.delete(0, END)
        self.txtObservacion.delete(0, END)
        self.txtResponsable.delete(0, END)

        self.spboxAnchoRecuperable.set(0)
        self.spboxAltoRecuperable.set(0)

        #Ocultar el formulario
        self.frmInfoIncidente.pack_forget()

        #Limpiar la información del vidrio
        self.lblNoPlano.config(text='')
        self.lblNoVidrio.config(text='')
        self.lblMedidas.config(text='')
        self.lblMateria.config(text='')
        self.lblCliente.config(text='')
        self.lblVendedor.config(text='')

    def alternar_estacion(self):
        if self.OtraEstacion.get():
            self.cboxEstacion.current(0)
            self.cboxEstacion.grid_forget()
            self.txtOtraEstacion.grid(row=0,column=1,pady=3,sticky='w')
        else:
            self.txtOtraEstacion.delete(0, END)
            self.txtOtraEstacion.grid_forget()
            self.cboxEstacion.grid(row=0,column=1,pady=3,sticky='w')
    