import ttkbootstrap as tkb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
import os
from decimal import Decimal
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends._backend_tk import NavigationToolbar2Tk
from collections import defaultdict
import operator
from connection import connection

class Main(tkb.Frame):
    def __init__(self, *args, ventana:tkb.Window, register:tkb.Window.register, id_empleado:int, usuario:str, **kwargs):
        super().__init__(*args, **kwargs)
        #Carpeta principal
        self.carpeta_principal = os.path.dirname(__file__)
        #.\igrams\gui\documentos
        self.carpeta_documentos = os.path.join(self.carpeta_principal, "documentos")
        #.\igrams\gui\images
        self.carpeta_imagenes = os.path.join(self.carpeta_principal, "images")
        self.ventana=ventana

        #Define el Frame como contenedor total que se asocia al notebook
        self.contenedor=ScrolledFrame(self)
        self.contenedor.pack(padx=10,pady=10, fill='both', expand=True)

        #Define la primera columna del contenedor
        self.colum1=tkb.Frame(self.contenedor)
        self.colum1.pack(padx=10, pady=10, side='left', fill='y')

        #-------------------Crear y ubicar el Frame para Ver la lista de planos creados---------------------------
        self.frmPlanosCreados=tkb.LabelFrame(self.colum1,text='Planos en Producción', bootstyle="primary")
        self.frmPlanosCreados.pack(padx=5, pady=5, ipadx=10, ipady=10)
        # Crear y ubicar la entrada de fecha de inicio para la búsqueda
        tkb.Label(self.frmPlanosCreados,text="Desde:",font=("Arial 10 bold")).grid(row=0,column=0, padx=5,pady=3,sticky='e')
        self.dtDesdePlanos=tkb.DateEntry(self.frmPlanosCreados, dateformat="%Y-%m-%d", width=12)
        self.dtDesdePlanos.grid(row=0,column=1,pady=3,sticky='w')
        # Crear y ubicar la entrada de fecha de fin para la búsqueda
        tkb.Label(self.frmPlanosCreados,text="Hasta:",font=("Arial 10 bold")).grid(row=0,column=2, padx=5,pady=3,sticky='e')
        self.dtHastaPlanos=tkb.DateEntry(self.frmPlanosCreados, dateformat="%Y-%m-%d", width=12)
        self.dtHastaPlanos.grid(row=0,column=3,pady=3,sticky='w')
        #Crear y ubicar el botón para actualizar los valores
        self.btnActualizar=tkb.Button(self.frmPlanosCreados,text="Actualizar",command=self.consultar_planos_ingresados)
        self.btnActualizar.grid(row=0, column=4, padx=5, pady=5)

        #Crear y ubicar el marco para contener la tabla y la barra de desplazamiento
        self.frmListaPlanos = tkb.Frame(self.frmPlanosCreados, relief="flat")
        self.frmListaPlanos.grid(row=1, column=0, columnspan=5, sticky='we')
        self.frmListaPlanos.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmListaPlanos.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmListaPlanos)
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
        self.tablePlanosCreados.tag_configure('listo', background='sky blue')
        self.tablePlanosCreados.bind("<<TreeviewSelect>>", self.lboxPlanosCreados_change)

        scrollbarV.config(command=self.tablePlanosCreados.yview)

        #Crear y ubicar el marco para contener el el resumen de la información
        self.frmResumenInfo = tkb.Frame(self.frmPlanosCreados, relief="flat")
        self.frmResumenInfo.grid(row=2, column=0, columnspan=3, sticky='we')
        
        #Crear y ubicar la etiqueta para la cantidad de planos ingresados
        tkb.Label(self.frmResumenInfo,text="Planos ingresados:",font=("Arial 10 bold"), justify="left").grid(row=0,column=0,padx=5,sticky='w')
        self.lblPlanosIngresados=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblPlanosIngresados.grid(row=0,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de planos terminados
        tkb.Label(self.frmResumenInfo,text="Planos terminados:",font=("Arial 10 bold"), justify="left").grid(row=1,column=0,padx=5,sticky='w')
        self.lblPlanosTerminados=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblPlanosTerminados.grid(row=1,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de área vendida
        tkb.Label(self.frmResumenInfo,text="Área vendida:",font=("Arial 10 bold"), justify="left").grid(row=2,column=0,padx=5,sticky='w')
        self.lblAreaVendida=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        self.lblAreaVendida.grid(row=2,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de área real
        #tkb.Label(self.frmResumenInfo,text="Área real:",font=("Arial 10 bold"), justify="left").grid(row=3,column=0,padx=5,sticky='w')
        #self.lblAreaReal=tkb.Label(self.frmResumenInfo,text="",font=("Arial 10"))
        #self.lblAreaReal.grid(row=3,column=1,sticky='w')

        #------------------------------Gráficas para producción ingresada-----------------------------------------------------
        #Define la segunda columna del contenedor
        self.colum2=tkb.Frame(self.contenedor)
        self.colum2.pack(padx=10, pady=10, side='left', fill='y')

        self.frmBarchart=tkb.LabelFrame(self.colum2, text='Área según vendedor', bootstyle="primary")
        self.frmBarchart.pack(padx=5, pady=5)

        self.fig1, self.ax1 = plt.subplots()
        self.fig1.set_dpi(70)
        self.fig1.set_size_inches(6,4)
        self.fig1.subplots_adjust(left=0.27, right=0.95, top=0.95, bottom=0.15)
        self.canvas1 = FigureCanvasTkAgg(self.fig1, master = self.frmBarchart)
        self.canvas1.get_tk_widget().pack()
        
        self.toolbar1 = NavigationToolbar2Tk(self.canvas1, self.frmBarchart)
        self.toolbar1._message_label.pack_forget()
        self.toolbar1._buttons['Back'].pack_forget()
        self.toolbar1._buttons['Forward'].pack_forget()
        self.toolbar1.pack()
        #--------------------------------------------------------------------------------------

        self.frmPiechart=tkb.LabelFrame(self.colum2, text='Porcentajes de ventas', bootstyle="primary")
        self.frmPiechart.pack(padx=5, pady=5)

        self.fig2, self.ax2 = plt.subplots()
        self.fig2.set_dpi(70)
        self.fig2.set_size_inches(6,3.5)
        self.fig2.subplots_adjust(left=0.05, right=0.95, top=1, bottom=0)
        self.ax2.set_axis_off()

        self.canvas2 = FigureCanvasTkAgg(self.fig2, master = self.frmPiechart)
        self.canvas2.get_tk_widget().pack()
        
        self.toolbar2 = NavigationToolbar2Tk(self.canvas2, self.frmPiechart)
        self.toolbar2._message_label.pack_forget()
        self.toolbar2._buttons['Back'].pack_forget()
        self.toolbar2._buttons['Forward'].pack_forget()
        self.toolbar2.pack()

        #-------------------------------------------------------------------------------------

        self.frmStemchart=tkb.LabelFrame(self.colum1,text='Ventas diarias', bootstyle="primary")
        self.frmStemchart.pack(padx=5, pady=5)

        self.fig3, self.ax3 = plt.subplots()
        self.fig3.set_dpi(70)
        self.fig3.set_size_inches(7.5,3.5)
        self.fig3.subplots_adjust(left=0.07, right=0.95, bottom=0.3, top=0.95)

        self.canvas3 = FigureCanvasTkAgg(self.fig3, master = self.frmStemchart)
        self.canvas3.get_tk_widget().pack()

        self.toolbar3 = NavigationToolbar2Tk(self.canvas3, self.frmStemchart)
        self.toolbar3._message_label.pack_forget()
        self.toolbar3._buttons['Back'].pack_forget()
        self.toolbar3._buttons['Forward'].pack_forget()
        self.toolbar3.pack()
        #--------------------------------------------------------------------------------------
    

    #-----------------------------------------------------------------------------------------------------
    def consultar_planos_ingresados(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT planos.*, empleados.nombre1, empleados.apellido1 FROM planos INNER JOIN empleados ON planos.id_empleado = empleados.id WHERE planos.marca_etiquetado BETWEEN %s AND %s;",(self.dtDesdePlanos.entry.get() + " 00:00:00",self.dtHastaPlanos.entry.get() + " 23:59:59"))
        self.PlanosIngresados=cursor.fetchall()
        conn.commit()

        fila=[]
        self.tablePlanosCreados.delete(*self.tablePlanosCreados.get_children())
        areaVendida=Decimal(0.00)
        cantidadTerminados=0
        for plano in self.PlanosIngresados:
            fila.clear()
            fila.append(str(plano['id']))
            fila.append(str(plano['marca_etiquetado']))

            if plano['id_materia_prima']:
                cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(plano['id_materia_prima'],))
                materia_prima=cursor.fetchone()
                conn.commit()
                fila.append(materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')            
            else:
                fila.append(str(plano['color']) + ' ' + str(plano['espesor']) + ' mm')
            fila.append(str(plano['cliente']))

            areaVendida=areaVendida+Decimal(plano['area'])

            #Consultar todos los vidrios ingresados para este plano
            cursor.execute("SELECT SUM(vidrios.area), COUNT(IF(vidrios.marca_almacenado IS NOT NULL, 1, null)) FROM vidrios WHERE id_plano=%s;",(plano['id'],))
            aux=cursor.fetchone()
            conn.commit()

            area_real=aux['SUM(vidrios.area)']
            vidrios_templados=aux['COUNT(IF(vidrios.marca_almacenado IS NOT NULL, 1, null))']
            
            if vidrios_templados==plano['cantidad_vidrios']:
                cantidadTerminados=cantidadTerminados+1
                self.tablePlanosCreados.insert('', 0, values=fila,tags=('listo',))
            else:
                self.tablePlanosCreados.insert('', 0, values=fila)

        self.lblPlanosIngresados.config(text=str(len(self.PlanosIngresados)))
        self.lblPlanosTerminados.config(text=str(cantidadTerminados))
        self.lblAreaVendida.config(text=str(round(areaVendida,2)) + ' m²')
        #self.lblAreaReal.config(text=str(area_real) + ' m²')
        
        conn.close()

        self.graficar_area_ingresada()

    def lboxPlanosCreados_change(self, Event):
        print('estoy cambiando de plano')

    def graficar_area_ingresada(self):
        datosAreaReal=defaultdict(Decimal)
        datosAreaVendida=defaultdict(Decimal)
        datosVentasDiarias=defaultdict(Decimal)
        for plano in self.PlanosIngresados:
            if plano['marca_etiquetado']:
                datosAreaVendida[plano['nombre1'] + ' ' + plano['apellido1']] += Decimal(plano['area'])
                datosVentasDiarias[str(plano['marca_etiquetado'])[:10]]+=Decimal(plano['area'])
                """ #Consultar todos los vidrios ingresados en el rango de fechas
                conn=connection(self.ventana)
                cursor=conn.cursor(dictionary=True)
                cursor.execute("SELECT id, ancho1, alto1, ancho2, alto2 FROM vidrios WHERE id_plano=%s;",(plano['id'],))
                vidrios=cursor.fetchall()
                conn.commit()
                conn.close()
                for vidrio in vidrios:
                    if vidrio['ancho1'] > vidrio['ancho2']:
                        if vidrio['alto1'] > vidrio['alto2']:
                            datosAreaReal[plano['nombre1'] + ' ' + plano['apellido1']] += (Decimal(vidrio['ancho1']/1000)*Decimal(vidrio['alto1']/1000))
                        else:
                            datosAreaReal[plano['nombre1'] + ' ' + plano['apellido1']] += (Decimal(vidrio['ancho1']/1000)*Decimal(vidrio['alto2']/1000))
                    else:
                        if vidrio['alto1'] > vidrio['alto2']:
                            datosAreaReal[plano['nombre1'] + ' ' + plano['apellido1']] += (Decimal(vidrio['ancho2']/1000)*Decimal(vidrio['alto1']/1000))
                        else:
                            datosAreaReal[plano['nombre1'] + ' ' + plano['apellido1']] += (Decimal(vidrio['ancho2']/1000)*Decimal(vidrio['alto2']/1000))
        for vendedor in datosAreaReal:
            datosAreaReal[vendedor]=round(datosAreaReal[vendedor],2)

        datosAreaReal = dict(sorted(datosAreaReal.items(), key=operator.itemgetter(1), reverse=True)) """
        if datosAreaVendida:
            datosAreaVendida = dict(sorted(datosAreaVendida.items(), key=operator.itemgetter(1), reverse=False))

            self.ax1.clear()
            self.ax1.barh(y=list(datosAreaVendida.keys()),width=list(datosAreaVendida.values()), height=0.7, edgecolor="black", align='center', color='cyan')

            self.ax1.set_xlabel('Área [m²]', fontsize='large')
            self.ax1.set_ylabel('Vendedor', fontsize='large')
            self.ax1.set_xlim(xmin=0,xmax=max(datosAreaVendida.values())*Decimal(1.25))

            for item in range(len(datosAreaVendida)):
                self.ax1.text(x=list(datosAreaVendida.values())[item] + max(datosAreaVendida.values())*Decimal(0.01), y=item, s=str(list(datosAreaVendida.values())[item]) + ' m²', horizontalalignment='left', verticalalignment='center')

            self.toolbar1.update()
            self.canvas1.draw()

            #Gráfica de pastel----------------------
            self.ax2.clear()
            self.ax2.pie(list(datosAreaVendida.values()), labels=list(datosAreaVendida.keys()), autopct = '%1.1f%%', wedgeprops=dict(width=0.5, edgecolor='black'), pctdistance = 0.75)
            self.toolbar2.update()
            self.canvas2.draw()

        if datosVentasDiarias:
            #Gráfica de Stem----------------------
            self.ax3.clear()
            self.ax3.stem(list(datosVentasDiarias.keys()), list(datosVentasDiarias.values()), label='Ventas')
            self.fig3.autofmt_xdate(rotation=90,ha='center',bottom=0.3)

            self.toolbar3.update()
            self.canvas3.draw()
