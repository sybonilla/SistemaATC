import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font
import os
import main


meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
class AplicacionATC:
    def __init__(self):
        self.ventana = tk.Tk()
        self.fecha_actual = datetime.datetime.now().date().strftime("%d/%m/%Y")
        self.mes_seleccionado = None
        self.año_seleccionado = None
        self.configurar_ventana()
        self.mostrar_pantalla_inicio()
        
    def configurar_ventana(self):
        """Configura los parámetros de la ventana principal"""
        self.ventana.title("ATC - Sistema de Control de Tráfico Aéreo")
        self.ventana.geometry("800x600+200+40")
        self.ventana.minsize(800, 600)
        self.ventana.maxsize(800, 600)
        self.ventana.resizable(False, False)
        self.ventana.configure(bg="#2B4263")
    
    def mostrar_pantalla_inicio(self):
        """Muestra la pantalla inicial con opciones de nuevo archivo o editar"""
        # Limpiar ventana si ya hay widgets
        for widget in self.ventana.winfo_children():
            widget.destroy()
        
        # Título principal
        titulo = tk.Label(
            self.ventana, 
            text="SISTEMA ATC", 
            fg="white", 
            bg="#2B4263", 
            font=("Arial", 24, "bold")
        )
        titulo.pack(pady=40)
        
        # Subtítulo
        subtitulo = tk.Label(
            self.ventana, 
            text="Seleccione una opción", 
            fg="white", 
            bg="#2B4263", 
            font=("Arial", 16)
        )
        subtitulo.pack(pady=10)
        
        # Frame para botones principales
        frame_botones_principales = tk.Frame(self.ventana, bg="#2B4263")
        frame_botones_principales.pack(pady=50)
        
        # Botón Nuevo Archivo
        boton_nuevo = tk.Button(
            frame_botones_principales,
            text="NUEVO ARCHIVO",
            command=self.crear_nuevo_archivo,
            fg="white",
            bg="#4CAF50",
            font=("Arial", 14, "bold"),
            width=20,
            height=2
        )
        boton_nuevo.pack(pady=15)
        
        # Botón Editar Archivo
        boton_editar = tk.Button(
            frame_botones_principales,
            text="EDITAR ARCHIVO EXISTENTE",
            command=self.editar_archivo_existente,
            fg="white",
            bg="#2196F3",
            font=("Arial", 14, "bold"),
            width=20,
            height=2
        )
        boton_editar.pack(pady=15)
    
    def crear_nuevo_archivo(self):
        """Muestra diálogo para crear nuevo archivo"""
        self.mostrar_dialogo_mes_año("crear")
    
    def editar_archivo_existente(self):
        """Muestra diálogo para editar archivo existente"""
        self.mostrar_dialogo_mes_año("editar")
    
    def mostrar_dialogo_mes_año(self, accion):
        """Muestra diálogo para seleccionar mes y año"""
        def procesar_seleccion():
            mes = combo_mes.get()
            año = entry_año.get()
            
            if not mes or not año:
                messagebox.showwarning("Advertencia", "Por favor seleccione mes e ingrese año")
                return
            
            if not año.isdigit() or len(año) != 4:
                messagebox.showwarning("Advertencia", "Por favor ingrese un año válido (4 dígitos)")
                return
            
            # Guardar mes y año seleccionados
            self.mes_seleccionado = meses.index(mes)
            self.año_seleccionado = año
            
            # Cerrar diálogo
            dialogo.destroy()
            
            if accion == "crear":
                # Crear nuevo archivo
                exito, mensaje = main.crearNuevoArchivo(self.mes_seleccionado, self.año_seleccionado)
                if exito:
                    messagebox.showinfo("Éxito", mensaje)
                    self.mostrar_pantalla_ingreso_datos(mes, año)
                else:
                    messagebox.showerror("Error", mensaje)
            else:
                # Verificar si el archivo existe
                ruta = os.path.join(os.path.expanduser('~'), 'Desktop', f'{mes}_{año}.xlsx')
                if os.path.exists(ruta):
                    self.mostrar_pantalla_ingreso_datos(mes, año)
                else:
                    messagebox.showerror("Error", f"El archivo {mes}_{año}.xlsx no existe en el Escritorio")
        
        # Crear diálogo
        dialogo = tk.Toplevel(self.ventana)
        dialogo.title("Nuevo Archivo" if accion == "crear" else "Editar Archivo")
        dialogo.geometry("400x250")
        dialogo.configure(bg="#2B4263")
        dialogo.resizable(False, False)
        dialogo.transient(self.ventana)
        dialogo.grab_set()
        
        # Centrar diálogo
        dialogo.update_idletasks()
        x = (self.ventana.winfo_width() - dialogo.winfo_width()) // 2 + self.ventana.winfo_x()
        y = (self.ventana.winfo_height() - dialogo.winfo_height()) // 2 + self.ventana.winfo_y()
        dialogo.geometry(f"+{x}+{y}")
        
        # Título del diálogo
        titulo_dialogo = tk.Label(
            dialogo,
            text="NUEVO ARCHIVO" if accion == "crear" else "EDITAR ARCHIVO",
            fg="white",
            bg="#2B4263",
            font=("Arial", 16, "bold")
        )
        titulo_dialogo.pack(pady=20)
        
        # Frame para mes
        frame_mes = tk.Frame(dialogo, bg="#2B4263")
        frame_mes.pack(pady=10)
        
        tk.Label(
            frame_mes,
            text="Mes:",
            fg="white",
            bg="#2B4263",
            font=("Arial", 12, "bold")
        ).pack(side="left", padx=(0, 10))
        
        combo_mes = ttk.Combobox(
            frame_mes,
            values=meses,
            state="readonly",
            font=("Arial", 12),
            width=15
        )
        combo_mes.pack(side="left")
        
        # Establecer mes actual por defecto
        mes_actual = datetime.datetime.now().month - 1
        combo_mes.current(mes_actual)
        
        # Frame para año
        frame_año = tk.Frame(dialogo, bg="#2B4263")
        frame_año.pack(pady=10)
        
        tk.Label(
            frame_año,
            text="Año:",
            fg="white",
            bg="#2B4263",
            font=("Arial", 12, "bold")
        ).pack(side="left", padx=(0, 10))
        
        entry_año = tk.Entry(
            frame_año,
            font=("Arial", 12),
            width=10,
            justify="center"
        )
        entry_año.pack(side="left")
        
        # Establecer año actual por defecto
        año_actual = datetime.datetime.now().year
        entry_año.insert(0, str(año_actual))
        
        # Frame para botones
        frame_botones = tk.Frame(dialogo, bg="#2B4263")
        frame_botones.pack(pady=20)
        
        boton_aceptar = tk.Button(
            frame_botones,
            text="ACEPTAR",
            command=procesar_seleccion,
            fg="white",
            bg="#4CAF50",
            font=("Arial", 12, "bold"),
            width=10
        )
        boton_aceptar.pack(side="left", padx=10)
        
        boton_cancelar = tk.Button(
            frame_botones,
            text="CANCELAR",
            command=dialogo.destroy,
            fg="white",
            bg="#F44336",
            font=("Arial", 12, "bold"),
            width=10
        )
        boton_cancelar.pack(side="left", padx=10)
    
    def mostrar_pantalla_ingreso_datos(self, mes, año):
        """Muestra la pantalla para ingresar datos de vuelos"""
        # Limpiar ventana
        for widget in self.ventana.winfo_children():
            widget.destroy()
        
        # Título principal con información del archivo
        titulo = tk.Label(
            self.ventana, 
            text=f"SISTEMA ATC - {mes.upper()} {año}", 
            fg="white", 
            bg="#2B4263", 
            font=("Arial", 20, "bold")
        )
        titulo.pack(pady=10)
        
        # Botón para volver al menú principal
        frame_superior = tk.Frame(self.ventana, bg="#2B4263")
        frame_superior.pack(fill="x", padx=20, pady=5)
        
        boton_volver = tk.Button(
            frame_superior,
            text="VOLVER AL MENÚ",
            command=self.mostrar_pantalla_inicio,
            fg="white",
            bg="#FF9800",
            font=("Arial", 10, "bold")
        )
        boton_volver.pack(side="left")
        
        # Crear frame principal para organizar mejor los elementos
        frame_principal = tk.Frame(self.ventana, bg="#2B4263")
        frame_principal.pack(expand=True, fill="both", padx=20, pady=10)
        
        # Lista de campos con sus configuraciones
        campos = [
            ("FECHA", "fecha", True),
            ("CALLSIGN", "callsign", True),
            ("TYP", "typ", True),
            ("REG", "reg", True),
            ("DEP AD", "dep_ad", False),
            ("DEST AD", "dest_ad", False)
        ]
        
        self.entradas = {}
        self.campos_obligatorios = []
        
        for texto, clave, obligatorio in campos:
            self.crear_campo(frame_principal, texto, clave, obligatorio)
        
        # Frame para botones
        frame_botones = tk.Frame(self.ventana, bg="#2B4263")
        frame_botones.pack(pady=20)
        
        # Botón enviar datos
        boton_enviar = tk.Button(
            frame_botones,
            text="ENVIAR DATOS",
            command=self.enviar_todos_los_datos,
            fg="white",
            bg="#2196F3",
            font=("Arial", 12, "bold"),
            width=15,
            height=1
        )
        boton_enviar.pack(side="left", padx=10)
        
        # Botón limpiar
        boton_limpiar = tk.Button(
            frame_botones,
            text="LIMPIAR",
            command=self.limpiar_campos,
            fg="white",
            bg="#FF9800",
            font=("Arial", 12, "bold"),
            width=15,
            height=1
        )
        boton_limpiar.pack(side="left", padx=10)
    
    def crear_campo(self, padre, texto, clave, obligatorio):
        """Crea un campo de entrada con su etiqueta"""
        frame_campo = tk.Frame(padre, bg="#2B4263")
        frame_campo.pack(fill="x", pady=8)
        
        # Etiqueta con indicador de campo obligatorio
        texto_etiqueta = texto + " *" if obligatorio else texto
        etiqueta = tk.Label(
            frame_campo,
            text=texto_etiqueta,
            fg="white",
            bg="#2B4263",
            font=("Arial", 14, "bold"),
            width=12,
            anchor="w"
        )
        etiqueta.pack(side="left")
        
        # Frame para entrada y botones (especial para fecha)
        frame_entrada_botones = tk.Frame(frame_campo, bg="#2B4263")
        frame_entrada_botones.pack(side="left", fill="x", expand=True, padx=(10, 0))
        
        # Campo de entrada
        entrada = tk.Entry(
            frame_entrada_botones,
            fg="black",
            bg="#E8F4FD",
            font=("Arial", 14),
            relief="solid",
            bd=2
        )
        entrada.pack(side="left", fill="x", expand=True)
        
        # Configurar navegación con Enter
        entrada.bind('<Return>', self.siguiente_campo)
        
        self.entradas[clave] = entrada
        
        # Si es campo obligatorio, agregar a la lista
        if obligatorio:
            self.campos_obligatorios.append(clave)
        
        # Botones especiales solo para el campo FECHA
        if clave == "fecha":
            frame_botones_fecha = tk.Frame(frame_entrada_botones, bg="#2B4263")
            frame_botones_fecha.pack(side="left", padx=(5, 0))
            
            # Botón para fecha de hoy
            boton_hoy = tk.Button(
                frame_botones_fecha,
                text="HOY",
                command=self.establecer_fecha_actual,
                fg="white",
                bg="#4CAF50",
                font=("Arial", 10, "bold"),
                width=6,
                height=1
            )
            boton_hoy.pack(side="left", padx=2)
            
            # Botón para otra fecha
            boton_otra_fecha = tk.Button(
                frame_botones_fecha,
                text="OTRA",
                command=self.establecer_otra_fecha,
                fg="white",
                bg="#FF9800",
                font=("Arial", 10, "bold"),
                width=6,
                height=1
            )
            boton_otra_fecha.pack(side="left", padx=2)
            
            # Establecer fecha actual por defecto
            entrada.insert(0, self.fecha_actual)
    
    def siguiente_campo(self, event):
        """Salta al siguiente campo cuando se presiona Enter"""
        widget_actual = event.widget
        todos_campos = list(self.entradas.values())
        
        # Encontrar el índice del campo actual
        try:
            indice_actual = todos_campos.index(widget_actual)
            # Ir al siguiente campo, o al primero si es el último
            siguiente_indice = (indice_actual + 1) % len(todos_campos)
            todos_campos[siguiente_indice].focus_set()
        except ValueError:
            # Si no se encuentra el widget, ir al primer campo
            todos_campos[0].focus_set()
        
        return "break"  # Prevenir el comportamiento por defecto de Enter
    
    def establecer_fecha_actual(self):
        """Establece la fecha actual en el campo correspondiente"""
        self.entradas["fecha"].delete(0, tk.END)
        self.entradas["fecha"].insert(0, self.fecha_actual)
    
    def establecer_otra_fecha(self):
        """Abre una ventana para seleccionar otra fecha"""
        def aplicar_fecha():
            try:
                dia = entry_dia.get()
                mes = entry_mes.get()
                año = entry_año.get()
                
                # Validar y formatear la fecha
                if dia and mes and año:
                    fecha = f"{int(dia):02d}/{int(mes):02d}/{año}"
                    self.entradas["fecha"].delete(0, tk.END)
                    self.entradas["fecha"].insert(0, fecha)
                    dialogo.destroy()
                else:
                    messagebox.showwarning("Advertencia", "Por favor complete todos los campos de fecha")
            except ValueError:
                messagebox.showerror("Error", "Por favor ingrese valores numéricos válidos")
        
        # Crear diálogo para fecha personalizada
        dialogo = tk.Toplevel(self.ventana)
        dialogo.title("Seleccionar Fecha")
        dialogo.geometry("300x150")
        dialogo.configure(bg="#2B4263")
        dialogo.resizable(False, False)
        dialogo.transient(self.ventana)
        dialogo.grab_set()
        
        # Frame para entradas
        frame_fecha = tk.Frame(dialogo, bg="#2B4263")
        frame_fecha.pack(pady=20)
        
        tk.Label(frame_fecha, text="Día:", fg="white", bg="#2B4263", font=("Arial", 10)).pack(side="left")
        entry_dia = tk.Entry(frame_fecha, width=3, font=("Arial", 10))
        entry_dia.pack(side="left", padx=5)
        
        tk.Label(frame_fecha, text="Mes:", fg="white", bg="#2B4263", font=("Arial", 10)).pack(side="left")
        entry_mes = tk.Entry(frame_fecha, width=3, font=("Arial", 10))
        entry_mes.pack(side="left", padx=5)
        
        tk.Label(frame_fecha, text="Año:", fg="white", bg="#2B4263", font=("Arial", 10)).pack(side="left")
        entry_año = tk.Entry(frame_fecha, width=5, font=("Arial", 10))
        entry_año.pack(side="left", padx=5)
        
        # Botones
        frame_botones = tk.Frame(dialogo, bg="#2B4263")
        frame_botones.pack(pady=10)
        
        tk.Button(
            frame_botones, 
            text="Aplicar", 
            command=aplicar_fecha,
            fg="white",
            bg="#4CAF50",
            font=("Arial", 10, "bold")
        ).pack(side="left", padx=10)
        
        tk.Button(
            frame_botones, 
            text="Cancelar", 
            command=dialogo.destroy,
            fg="white",
            bg="#F44336",
            font=("Arial", 10, "bold")
        ).pack(side="left", padx=10)
        
        # Establecer fecha actual como valores por defecto
        hoy = datetime.datetime.now()
        entry_dia.insert(0, str(hoy.day))
        entry_mes.insert(0, str(hoy.month))
        entry_año.insert(0, str(hoy.year))
    
    def validar_campos_obligatorios(self):
        """Valida que todos los campos obligatorios estén completos"""
        campos_vacios = []
        
        for campo in self.campos_obligatorios:
            if not self.entradas[campo].get().strip():
                campos_vacios.append(campo)
        
        return campos_vacios
    
    def enviar_todos_los_datos(self):
        """Envía y procesa todos los datos ingresados"""
        # Validar campos obligatorios
        campos_vacios = self.validar_campos_obligatorios()
        
        if campos_vacios:
            mensaje = "Los siguientes campos obligatorios están vacíos:\n"
            for campo in campos_vacios:
                nombre_campo = {
                    "fecha": "FECHA",
                    "callsign": "CALLSIGN", 
                    "typ": "TYP",
                    "reg": "REG"
                }.get(campo, campo)
                mensaje += f"- {nombre_campo}\n"
            
            messagebox.showwarning("Campos Obligatorios", mensaje)
            
            # Poner foco en el primer campo vacío
            if campos_vacios:
                self.entradas[campos_vacios[0]].focus_set()
            return
        
        # Recopilar datos para el archivo Excel
        datos_vuelo = {
            'fecha': self.entradas["fecha"].get().strip(),
            'callsign': self.entradas["callsign"].get().strip(),
            'typ': self.entradas["typ"].get().strip(),
            'reg': self.entradas["reg"].get().strip(),
            'dep_ad': self.entradas["dep_ad"].get().strip(),
            'dest_ad': self.entradas["dest_ad"].get().strip()
        }
        
        # Agregar datos al archivo Excel
        exito, mensaje = main.editarArchivoExistente(
            self.mes_seleccionado, 
            self.año_seleccionado, 
            datos_vuelo
        )
        
        if exito:
            messagebox.showinfo("Éxito", mensaje)
            # Limpiar campos después de enviar exitosamente
            self.limpiar_campos()
            self.establecer_fecha_actual()  # Volver a establecer fecha actual
        else:
            messagebox.showerror("Error", mensaje)
    
    def limpiar_campos(self):
        """Limpia todos los campos de entrada"""
        for entrada in self.entradas.values():
            entrada.delete(0, tk.END)
    
    def ejecutar(self):
        """Inicia la aplicación"""
        self.ventana.mainloop()

# Ejecutar la aplicación
if __name__ == "__main__":
    app = AplicacionATC()
    app.ejecutar()