import ttkbootstrap as tkb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledFrame
import os
from decimal import Decimal
import matplotlib.pyplot as plt
from collections import defaultdict
import operator
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

        #-------------------Crear y ubicar el Frame para ver la tabla de vidrios templados---------------------------
        #Define la primera columna del contenedor
        self.colum1=tkb.Frame(self.contenedor)
        self.colum1.pack(padx=10, pady=10, side='left', fill='y')
        
        self.frmVidriosTemplados=tkb.LabelFrame(self.colum1,text='Lista de vidrios templados', bootstyle="primary")
        self.frmVidriosTemplados.pack(padx=5, pady=5, ipadx=10, ipady=10)
        # Crear y ubicar la entrada de fecha de inicio para la búsqueda
        tkb.Label(self.frmVidriosTemplados,text="Desde:",font=("Arial 10 bold")).grid(row=0,column=0, padx=5,pady=3,sticky='w')
        self.dtDesdeVTemplados=tkb.DateEntry(self.frmVidriosTemplados, dateformat="%Y-%m-%d", width=12)
        self.dtDesdeVTemplados.grid(row=0,column=1,pady=3,sticky='w')
        # Crear y ubicar la entrada de fecha de fin para la búsqueda
        tkb.Label(self.frmVidriosTemplados,text="Hasta:",font=("Arial 10 bold")).grid(row=0,column=2, padx=5,pady=3,sticky='w')
        self.dtHastaVTemplados=tkb.DateEntry(self.frmVidriosTemplados, dateformat="%Y-%m-%d", width=12)
        self.dtHastaVTemplados.grid(row=0,column=3,pady=3,sticky='w')
        #Crear y ubicar el botón para actualizar los valores
        self.btnActualizarVTemplados=tkb.Button(self.frmVidriosTemplados,text="Actualizar",command=self.consultar_vidrios_templados)
        self.btnActualizarVTemplados.grid(row=0, column=4, columnspan=2, padx=5, pady=5)

        #Crear y ubicar el marco para contener el listbox y la barra de desplazamiento
        self.frmListaVTemplados = tkb.Frame(self.frmVidriosTemplados, relief="flat")
        self.frmListaVTemplados.grid(row=1, column=0, columnspan=5, sticky='we')
        self.frmListaVTemplados.bind("<Enter>", lambda event: self.contenedor.disable_scrolling())
        self.frmListaVTemplados.bind("<Leave>", lambda event: self.contenedor.enable_scrolling())
        #Crear una barra de deslizamiento con orientación vertical
        scrollbarV = tkb.Scrollbar(self.frmListaVTemplados, orient=VERTICAL)
        scrollbarV.grid(row=0,column=1, pady=5, sticky='ns')

        fuente=tkb.font.Font(family="times", size=11)
        style = tkb.Style()
        style.configure("mystyle.Treeview", font=('times', 11))

        #Crear y ubicar la tabla de los planos creados
        self.tableVidriosTemplados=tkb.Treeview(self.frmListaVTemplados, show='headings', height=9, yscrollcommand=scrollbarV.set, style="mystyle.Treeview")
        self.tableVidriosTemplados.configure(columns=('plano', 'vidrio', 'fecha-hora', 'materia-prima', 'medidas'))
        self.tableVidriosTemplados.column('plano', stretch=False, anchor='w', width=fuente.measure(' 0000000 '))
        self.tableVidriosTemplados.column('vidrio', stretch=False, anchor='w', width=fuente.measure(' 0000000 '))
        self.tableVidriosTemplados.column('fecha-hora', stretch=False, anchor='w', width=fuente.measure('0 0000-00-00 00:00:00 0'))
        self.tableVidriosTemplados.column('materia-prima', stretch=False, anchor='w', width=fuente.measure('xxxxxxxxxxxxxxxxxxx'))
        self.tableVidriosTemplados.column('medidas', stretch=False, anchor='center', width=fuente.measure(' 0000X0000 '))

        for col in self.tableVidriosTemplados['columns']:
            self.tableVidriosTemplados.heading(col, text=col.title(), anchor=W)

        self.tableVidriosTemplados.grid(row=0,column=0, padx=5, pady=5)
        self.tableVidriosTemplados.bind("<<TreeviewSelect>>", self.lboxVidriosTemplados_change)

        scrollbarV.config(command=self.tableVidriosTemplados.yview)

        #Crear y ubicar el marco para contener el resumen de la información
        self.frmResumenInfoVTemplados = tkb.Frame(self.frmVidriosTemplados, relief="flat")
        self.frmResumenInfoVTemplados.grid(row=2, column=0, columnspan=3, sticky='we')
        #Crear y ubicar la etiqueta para la cantidad de área templada
        tkb.Label(self.frmResumenInfoVTemplados,text="Vidrios templados:",font=("Arial 10 bold"), justify="left").grid(row=0,column=0,padx=5,sticky='w')
        self.lblVidriosTemplados=tkb.Label(self.frmResumenInfoVTemplados,text="",font=("Arial 10"))
        self.lblVidriosTemplados.grid(row=0,column=1,sticky='w')
        #Crear y ubicar la etiqueta para la cantidad de área templada
        tkb.Label(self.frmResumenInfoVTemplados,text="Área templada:",font=("Arial 10 bold"), justify="left").grid(row=1,column=0,padx=5,sticky='w')
        self.lblAreaTemplada=tkb.Label(self.frmResumenInfoVTemplados,text="",font=("Arial 10"))
        self.lblAreaTemplada.grid(row=1,column=1,sticky='w')
        #-------------------Crear y ubicar el Frame para ver la lista de entrega por día---------------------------

        self.frmPlanosDiarios=tkb.LabelFrame(self.colum1,text='Lista de entrega diaria', bootstyle="primary")
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
        self.tablePlanosEntrega=tkb.Treeview(self.frmListaPDiarios, show='headings', height=9, yscrollcommand=scrollbarV.set, style="mystyle.Treeview")
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

        #------------------------------Gráficas para templado-----------------------------------------------------
        #Define la segunda columna del contenedor
        self.colum2=tkb.Frame(self.contenedor)
        self.colum2.pack(padx=10, pady=10, side='left', fill='y')

        self.frmBarchart=tkb.LabelFrame(self.colum2, text='Área templada por hora', bootstyle="primary")
        self.frmBarchart.pack(padx=5, pady=5)

        self.fig1, self.ax1 = plt.subplots()
        self.fig1.set_dpi(68)
        self.fig1.set_size_inches(5,8)
        self.fig1.subplots_adjust(left=0.17, right=0.95, top=0.98, bottom=0.08)
        self.canvas1 = FigureCanvasTkAgg(self.fig1, master = self.frmBarchart)
        self.canvas1.get_tk_widget().pack()
        
        self.toolbar1 = NavigationToolbar2Tk(self.canvas1, self.frmBarchart)
        self.toolbar1._message_label.pack_forget()
        self.toolbar1._buttons['Back'].pack_forget()
        self.toolbar1._buttons['Forward'].pack_forget()
        self.toolbar1.pack()
        #--------------------------------------------------------------------------------------
    
    def consultar_vidrios_templados(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT vidrios.id, vidrios.id_plano, vidrios.ancho1, vidrios.alto1, vidrios.ancho2, vidrios.alto2, vidrios.marca_almacenado, planos.id_materia_prima, planos.color, planos.espesor FROM vidrios INNER JOIN planos ON vidrios.id_plano = planos.id WHERE vidrios.marca_almacenado BETWEEN %s AND %s;",(self.dtDesdeVTemplados.entry.get() + " 00:00:00",self.dtHastaVTemplados.entry.get() + " 23:59:59"))
        self.VidriosTemplados=cursor.fetchall()
        conn.commit()

        fila=[]
        self.tableVidriosTemplados.delete(*self.tableVidriosTemplados.get_children())
        area=Decimal(0.00)
        for vidrio in self.VidriosTemplados:
            fila.clear()
            fila.append(str(vidrio['id_plano']))
            fila.append(str(vidrio['id']))
            fila.append(str(vidrio['marca_almacenado']))

            #Consultar materia prima
            if vidrio['id_materia_prima']:
                cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(vidrio['id_materia_prima'],))
                materia_prima=cursor.fetchone()
                conn.commit()
                fila.append(materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')
            else:
                fila.append(str(vidrio['color']) + ' ' + str(vidrio['espesor']) + ' mm')
            
            #Calcular la suma de áreas
            if vidrio['ancho1'] > vidrio['ancho2']:
                if vidrio['alto1'] > vidrio['alto2']:
                    area=area+(Decimal(vidrio['ancho1']/1000)*Decimal(vidrio['alto1']/1000))
                    fila.append(str(vidrio['ancho1'])+"x"+str(vidrio['alto1']))
                else:
                    area=area+(Decimal(vidrio['ancho1']/1000)*Decimal(vidrio['alto2']/1000))
                    fila.append(str(vidrio['ancho1'])+"x"+str(vidrio['alto2']))
            else:
                if vidrio['alto1'] > vidrio['alto2']:
                    area=area+(Decimal(vidrio['ancho2']/1000)*Decimal(vidrio['alto1']/1000))
                    fila.append(str(vidrio['ancho2'])+"x"+str(vidrio['alto1']))
                else:
                    area=area+(Decimal(vidrio['ancho2']/1000)*Decimal(vidrio['alto2']/1000))
                    fila.append(str(vidrio['ancho2'])+"x"+str(vidrio['alto2']))
            
            self.tableVidriosTemplados.insert('', 0, values=fila)
        
        self.lblVidriosTemplados.config(text=str(len(self.VidriosTemplados)))
        self.lblAreaTemplada.config(text=str(round(area,2)) + ' m²')

        conn.close()
        self.graficar_area_templada()

    def lboxVidriosTemplados_change(self, Event):
        print('estoy cambiando de vidrio')

    def graficar_area_templada(self):
        horas=['00:00','01:00','02:00','03:00','04:00','05:00','06:00','07:00','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00','18:00','19:00','20:00','21:00','22:00','23:00']
        datosGrafica = dict.fromkeys(horas, Decimal(0.00))

        for vidrio in self.VidriosTemplados:
            #Calcular el área
            if vidrio['ancho1'] > vidrio['ancho2']:
                if vidrio['alto1'] > vidrio['alto2']:
                    area=(Decimal(vidrio['ancho1']/1000)*Decimal(vidrio['alto1']/1000))
                else:
                    area=(Decimal(vidrio['ancho1']/1000)*Decimal(vidrio['alto2']/1000))
            else:
                if vidrio['alto1'] > vidrio['alto2']:
                    area=(Decimal(vidrio['ancho2']/1000)*Decimal(vidrio['alto1']/1000))
                else:
                    area=(Decimal(vidrio['ancho2']/1000)*Decimal(vidrio['alto2']/1000))

            # Si la hora de almacenado ya está en el diccionario, suma el valor
            hora=str(vidrio['marca_almacenado'])[11:13]+':00'
            if hora in datosGrafica:
                datosGrafica[hora] = round(datosGrafica[hora]+area,2)
            # Si no está en el diccionario, agrega el nombre y el valor
            else:
                datosGrafica[hora] = round(area,2)

        self.ax1.clear()
        self.ax1.barh(list(datosGrafica.keys()), list(datosGrafica.values()), height=1, edgecolor="black",align='edge', color='purple')
        self.ax1.set_xlim(xmin=0,xmax=max(datosGrafica.values())*Decimal(1.25))
        self.ax1.set_xlabel('Área [m²]', fontsize='large')
        self.ax1.set_ylabel('Hora', fontsize='large')

        for item in range(len(datosGrafica)):
            self.ax1.text(x=list(datosGrafica.values())[item] + max(datosGrafica.values())*Decimal(0.01), y=item+0.5, s=str(list(datosGrafica.values())[item]) + ' m²', horizontalalignment='left', verticalalignment='center')

        self.toolbar1.update()
        self.canvas1.draw()

    def consultar_planos_diarios(self):
        conn=connection(self.ventana)
        cursor=conn.cursor(dictionary=True)
        cursor.execute("SELECT id, id_empleado, cantidad_vidrios, cliente, id_materia_prima, color, espesor, marca_etiquetado FROM planos WHERE marca_plazo BETWEEN %s AND %s;",(self.dtFechaEntrega.entry.get() + " 00:00:00",self.dtFechaEntrega.entry.get() + " 23:59:59"))
        self.PlanosDiarios=cursor.fetchall()
        conn.commit()

        fila=[]
        i=0
        area=Decimal(0.00)
        
        self.tablePlanosEntrega.delete(*self.tablePlanosEntrega.get_children())

        for plano in self.PlanosDiarios:
            if bool(plano['marca_etiquetado']):
                fila.clear()
                fila.append(str(plano['id']))
                i+=1

                #Consultar todos los vidrios que deben salir según la fecha
                cursor.execute("SELECT id, ancho1, alto1, ancho2, alto2, marca_almacenado FROM vidrios WHERE id_plano=%s;",(plano['id'],))
                vidrios=cursor.fetchall()
                conn.commit()
                listo=True
                j=0
                for vidrio in vidrios:
                    listo = listo and bool(vidrio['marca_almacenado'])
                    if bool(vidrio['marca_almacenado']):
                        j+=1
                    if vidrio['ancho1'] > vidrio['ancho2']:
                        if vidrio['alto1'] > vidrio['alto2']:
                            area=area+(Decimal(vidrio['ancho1']/1000)*Decimal(vidrio['alto1']/1000))
                        else:
                            area=area+(Decimal(vidrio['ancho1']/1000)*Decimal(vidrio['alto2']/1000))
                    else:
                        if vidrio['alto1'] > vidrio['alto2']:
                            area=area+(Decimal(vidrio['ancho2']/1000)*Decimal(vidrio['alto1']/1000))
                        else:
                            area=area+(Decimal(vidrio['ancho2']/1000)*Decimal(vidrio['alto2']/1000))
                fila.append(str(j) + '/' + str(plano['cantidad_vidrios']))
                #Consultar materia prima
                if plano['id_materia_prima']:
                    cursor.execute("SELECT id_color, id_espesor FROM materia_prima WHERE id = %s;",(plano['id_materia_prima'],))
                    materia_prima=cursor.fetchone()
                    conn.commit()
                    fila.append(materia_prima['id_color'] + ' ' + str(materia_prima['id_espesor']) + ' mm')
                else:
                    fila.append(str(plano['color']) + ' ' + str(plano['espesor']) + ' mm')
                
                fila.append(str(plano['cliente']))

                if listo:
                    self.tablePlanosEntrega.insert('', END, values=fila,tags=('listo',))
                else:
                    self.tablePlanosEntrega.insert('', END, values=fila)
        
        self.lblPlanosDiarios.config(text=str(i))
        self.lblAreaDiaria.config(text=str(round(area,2)) + ' m²')
        conn.close()

    def tablePlanosEntrega_change(self, Event):
        print('estoy cambiando de vidrio')