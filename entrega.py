import ttkbootstrap as tkb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
import os
from decimal import Decimal
import matplotlib.pyplot as plt
from collections import defaultdict
from connection import connection
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends._backend_tk import NavigationToolbar2Tk

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

        #-------------------Crear y ubicar el Frame para ver la tabla de planos a entregar---------------------------
        #Define la primera columna del contenedor
        self.colum1=tkb.Frame(self.contenedor)
        self.colum1.pack(padx=10, pady=10, side='left', fill='y')
        
        self.frmPlanosEntrega=tkb.LabelFrame(self.colum1,text='Planos a entregar', bootstyle="primary")
        self.frmPlanosEntrega.pack(padx=5, pady=5, ipadx=10, ipady=10)
        # Crear y ubicar la entrada de fecha de inicio para la búsqueda
        tkb.Label(self.frmPlanosEntrega,text="Desde:",font=("Arial 10 bold")).grid(row=0,column=0, padx=5,pady=3,sticky='w')
        self.dtDesdePEntrega=tkb.DateEntry(self.frmPlanosEntrega, dateformat="%Y-%m-%d", width=12)
        self.dtDesdePEntrega.grid(row=0,column=1,pady=3,sticky='w')
        # Crear y ubicar la entrada de fecha de fin para la búsqueda
        tkb.Label(self.frmPlanosEntrega,text="Hasta:",font=("Arial 10 bold")).grid(row=0,column=2, padx=5,pady=3,sticky='w')
        self.dtHastaPEntrega=tkb.DateEntry(self.frmPlanosEntrega, dateformat="%Y-%m-%d", width=12)
        self.dtHastaPEntrega.grid(row=0,column=3,pady=3,sticky='w')
        #Crear y ubicar el botón para actualizar los valores
        self.btnActualizarPEntrega=tkb.Button(self.frmPlanosEntrega,text="Actualizar",command=self.consultar_planos_entrega)
        self.btnActualizarPEntrega.grid(row=0, column=4, columnspan=2, padx=5, pady=5)

        #Crear y ubicar el marco para contener la tabla y la barra de desplazamiento
        self.frmTablaPlanosEntrega = tkb.Frame(self.frmPlanosEntrega, relief="flat")
        self.frmTablaPlanosEntrega.grid(row=1, column=0, columnspan=5, sticky='we')
        self.frmTablaPlanosEntrega.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmTablaPlanosEntrega.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmTablaPlanosEntrega, orient=VERTICAL)
        scrollbarV.grid(row=0,column=1, pady=5, sticky='ns')

        fuente=tkb.font.Font(family="times", size=11)
        style = tkb.Style()
        style.configure("mystyle.Treeview", font=('times', 11))

        #Crear y ubicar la tabla de los planos a entregar
        self.tablePlanosEntrega=tkb.Treeview(self.frmTablaPlanosEntrega, show='headings', height=10, yscrollcommand=scrollbarV.set, style="mystyle.Treeview")
        self.tablePlanosEntrega.configure(columns=('plano', 'estado', 'materia-prima', 'fecha-entrega', 'cliente', 'vendedor'))
        self.tablePlanosEntrega.column('plano', stretch=False, anchor='center', width=fuente.measure(' 0000000 '))
        self.tablePlanosEntrega.column('estado', stretch=False, anchor='center', width=fuente.measure(' 000/000 '))
        self.tablePlanosEntrega.column('materia-prima', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxx'))
        self.tablePlanosEntrega.column('fecha-entrega', stretch=False, anchor='w', width=fuente.measure('00 0000-00-00 00'))
        self.tablePlanosEntrega.column('cliente', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxxxxxxxxx'))
        self.tablePlanosEntrega.column('vendedor', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxxxxxx'))

        for col in self.tablePlanosEntrega['columns']:
            self.tablePlanosEntrega.heading(col, text=col.title(), anchor=W)

        self.tablePlanosEntrega.grid(row=0,column=0, padx=5, pady=5)
        self.tablePlanosEntrega.tag_configure('listo', background='sky blue')
        self.tablePlanosEntrega.bind("<<TreeviewSelect>>", self.tablePlanosEntrega_change)

        scrollbarV.config(command=self.tablePlanosEntrega.yview)

        #Crear y ubicar los indicadores para el área templada y los planos listos
        self.meterAreaTemplada =tkb.Meter(self.frmPlanosEntrega, metertype='semi', padding=5, interactive=False, metersize=125, textright ='m²')
        self.meterAreaTemplada.grid(row=2, column=0, columnspan=2, sticky='we')
        self.meterPlanosTemplados =tkb.Meter(self.frmPlanosEntrega, metertype='semi', padding=5, interactive=False, metersize=125, stripethickness=10, textright ='P.')
        self.meterPlanosTemplados.grid(row=2, column=2, columnspan=2, sticky='we')

        #Crear y ubicar el marco para contener la tabla y la barra de desplazamiento
        self.frmTablaVidrios = tkb.Frame(self.frmPlanosEntrega, relief="flat")
        self.frmTablaVidrios.grid(row=3, column=0, columnspan=5, sticky='we')
        self.frmTablaVidrios.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmTablaVidrios.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmTablaVidrios, orient=VERTICAL)
        scrollbarV.grid(row=0,column=1, pady=5, sticky='ns')

        #Crear y ubicar la tabla de los vidrios
        self.tableInfoVidrios=tkb.Treeview(self.frmTablaVidrios, show='headings', height=10, yscrollcommand=scrollbarV.set, style="mystyle.Treeview")
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
        self.tableInfoVidrios.bind("<<TreeviewSelect>>", self.tableVidrios_change)

        scrollbarV.config(command=self.tableInfoVidrios.yview)

        #-------------------Crear y ubicar el Frame para ver la gráfica del área a templar---------------------------
        #Define la segunda columna del contenedor
        self.colum2=tkb.Frame(self.contenedor)
        #self.colum2.pack(padx=10, pady=10, side='left', fill='y')

        self.frmBarchart=tkb.LabelFrame(self.colum2, text='Producción a entregar', bootstyle="primary")
        self.frmBarchart.pack(padx=5, pady=5)

        self.fig1, self.ax1 = plt.subplots()
        self.fig1.set_dpi(70)
        self.fig1.set_size_inches(5,7.5)
        self.fig1.subplots_adjust(left=0.33, right=0.95, top=0.98, bottom=0.08)

        self.canvas1 = FigureCanvasTkAgg(self.fig1, master = self.frmBarchart)
        self.canvas1.get_tk_widget().pack()

        self.toolbar1 = NavigationToolbar2Tk(self.canvas1, self.frmBarchart)
        self.toolbar1._message_label.pack_forget()
        self.toolbar1._buttons['Back'].pack_forget()
        self.toolbar1._buttons['Forward'].pack_forget()
        self.toolbar1.pack()
        
    #----------------------------------------------------------------------------------------------------
    def consultar_planos_entrega(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT planos.id, planos.cantidad_vidrios, planos.cliente, planos.id_materia_prima, planos.color, planos.espesor, planos.area, planos.marca_plazo, empleados.nombre1, empleados.apellido1 FROM planos INNER JOIN empleados ON planos.id_empleado = empleados.id WHERE planos.marca_etiquetado IS NOT NULL AND planos.marca_plazo BETWEEN %s AND %s;",(self.dtDesdePEntrega.entry.get() + " 00:00:00", self.dtHastaPEntrega.entry.get() + " 23:59:59"))
        self.PlanosEntrega=cursor.fetchall()
        conn.commit()
        #Consultar el área total y área templada
        cursor.execute("SELECT SUM(vidrios.area), SUM(IF(vidrios.marca_almacenado IS NOT NULL, vidrios.area, 0)) FROM vidrios INNER JOIN planos WHERE vidrios.id_plano = planos.id AND planos.marca_plazo BETWEEN %s AND %s;",(self.dtDesdePEntrega.entry.get() + " 00:00:00", self.dtHastaPEntrega.entry.get() + " 23:59:59"))
        aux=cursor.fetchone()
        conn.commit()
        area_total=aux['SUM(vidrios.area)']
        area_templada=aux['SUM(IF(vidrios.marca_almacenado IS NOT NULL, vidrios.area, 0))']

        fila=[]
        planos_templados=0
        planos_etiquetados=len(self.PlanosEntrega)

        self.tablePlanosEntrega.delete(*self.tablePlanosEntrega.get_children())
        self.tableInfoVidrios.delete(*self.tableInfoVidrios.get_children())

        i=0
        for plano in self.PlanosEntrega:
            i+=1
            
            fila.clear()
            fila.append(str(plano['id']))

            #Consultar todos los vidrios templados
            cursor.execute("SELECT COUNT(marca_almacenado) FROM vidrios WHERE id_plano=%s;",(plano['id'],))
            vidrios_templados=cursor.fetchone()
            conn.commit()

            fila.append(str(vidrios_templados['COUNT(marca_almacenado)']) + '/' + str(plano['cantidad_vidrios']))  

            #Consultar materia prima
            if plano['id_materia_prima']:
                cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(plano['id_materia_prima'],))
                materia_prima=cursor.fetchone()
                conn.commit()
                fila.append(materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')
            else:
                fila.append(str(plano['color']) + ' ' + str(plano['espesor']) + ' mm')

            fila.append(str(plano['marca_plazo'])[:10])
            fila.append(str(plano['cliente']))
            fila.append(plano['nombre1'] + ' ' + plano['apellido1'])

            if vidrios_templados['COUNT(marca_almacenado)'] == plano['cantidad_vidrios']:
                self.tablePlanosEntrega.insert('', END, values=fila, tags=('listo',))
                planos_templados+=1
            else:
                self.tablePlanosEntrega.insert('', END, values=fila)
        conn.close()

        #Se configura el indicador visual
        if area_total is None:
            self.meterAreaTemplada.configure(amountused=0, subtext='de 0 m²')
            self.meterPlanosTemplados.configure(amountused=0, subtext='de 0 Planos')
        elif int(area_total)==0:
            self.meterAreaTemplada.configure(amountused=0, subtext='de 0 m²')
            self.meterPlanosTemplados.configure(amountused=0, subtext='de 0 Planos')
        else: 
            self.meterPlanosTemplados.configure(amounttotal=planos_etiquetados,amountused=planos_templados,subtext=f'de {planos_etiquetados} Planos')
            self.meterAreaTemplada.configure(amounttotal= int(area_total),amountused=int(area_templada),subtext=f'de {int(area_total)} m²')

    def tablePlanosEntrega_change(self, Event):
        if self.tablePlanosEntrega.selection():
            indice=self.tablePlanosEntrega.selection()[0]
            id_plano=self.tablePlanosEntrega.item(indice,'values')[0]
            
            #Buscar y listar todos los vidrios--------------------------------------------------
            conn=connection(self.ventana)
            cursor=conn.cursor(dictionary=True)
            cursor.execute("SELECT id, ancho1, alto1, ancho2, alto2, tipo_arenado, forma, referencia, observacion, marca_etiquetado, marca_pulido, marca_marcado, marca_templado, marca_almacenado, marca_arenado, marca_despacho FROM vidrios WHERE id_plano=%s;",(int(id_plano),))
            self.datosVidrios=cursor.fetchall()
            conn.commit()
            conn.close()
            
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

    def tableVidrios_change(self, Event):
        print('estoy cambiando de vidrio')

    def graficar_produccion(self):
        self.ax1.clear()

        datosAreaTotal=defaultdict(Decimal)
        datosAreaTemplada=defaultdict(Decimal)
        for plano in self.PlanosEntrega:
            if bool(plano['marca_etiquetado']):
                if plano['id_materia_prima']:
                    conn=connection(self.ventana)
                    cursor=conn.cursor(dictionary=True)
                    cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(plano['id_materia_prima'],))
                    materia_prima=cursor.fetchone()
                    conn.commit()
                    conn.close()

                    datosAreaTotal[materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm'] += plano['area']
                    if plano['marca_almacenado']:
                        datosAreaTemplada[materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm'] += plano['area']
                    else:
                        datosAreaTemplada[materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm'] += Decimal(0.0)
                else:
                    datosAreaTotal[str(plano['color']) + ' ' + str(plano['espesor']) + ' mm'] += plano['area']
                    if plano['marca_almacenado']:
                        datosAreaTemplada[str(plano['color']) + ' ' + str(plano['espesor']) + ' mm'] += plano['area']
                    else:
                        datosAreaTemplada[str(plano['color']) + ' ' + str(plano['espesor']) + ' mm'] += Decimal(0.0)

        #Gráfica de barras----------------------
        ancho=0.3
        self.ax1.barh(list(datosAreaTotal.keys()), list(datosAreaTotal.values()), height=ancho, edgecolor="black", align='edge', color='coral', label='Área total')
        self.ax1.barh(list(datosAreaTemplada.keys()), list(datosAreaTemplada.values()), height=-1*ancho, edgecolor="black", align='edge', color='SkyBlue', label='Área templada')
        self.ax1.set_xlim(xmin=0,xmax=max(datosAreaTotal.values())*Decimal(1.25))

        for item in range(len(datosAreaTotal)):
            self.ax1.text(x=list(datosAreaTotal.values())[item]+max(datosAreaTotal.values())*Decimal(0.01), y=item+(ancho/2), s=str(list(datosAreaTotal.values())[item]) + ' m²', horizontalalignment='left', verticalalignment='center')
            self.ax1.text(x=list(datosAreaTemplada.values())[item]+max(datosAreaTotal.values())*Decimal(0.01), y=item-(ancho/2), s=str(list(datosAreaTemplada.values())[item]) + ' m²', horizontalalignment='left', verticalalignment='center')

        self.toolbar1.update()
        self.canvas1.draw()