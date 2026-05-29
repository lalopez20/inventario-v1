import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import os
from datetime import datetime
from PIL import Image, ImageTk

# Librerías de Reportes Externas
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Librerías para Gráficas
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class InventarioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sushi Stock Premium - Control de Almacén")
        
        # VENTANA PRINCIPAL EN PANTALLA COMPLETA
        try:
            self.root.state('zoomed') 
        except tk.TclError:
            self.root.attributes('-zoomed', True) 
            
        # ================= PALETA DE COLORES SOFISTICADA POR CAPAS =================
        self.bg_base = "#F0F2F5"        # Gris muy claro neutro (Fondo absoluto)
        self.bg_panel = "#FFFFFF"       # Blanco puro (Para las secciones/tarjetas)
        self.text_dark = "#1A252F"      # Azul marino casi negro (Texto Principal)
        self.header_bg = "#2C3E50"      # Azul noche (Encabezado superior)
        self.btn_bg = "#34495E"         # Azul pizarra (Botones del menú normales)
        self.btn_active = "#2C3E50"     # Azul noche (Botón menú presionado)
        self.border_color = "#E0E0E0"   # Gris sutil para bordes de paneles
        
        self.root.configure(bg=self.bg_base)
        
        # CONFIGURACIÓN DE ESTILOS MODERNOS
        self.estilo = ttk.Style()
        self.estilo.theme_use("clam")
        
        # Estilos Generales base
        self.estilo.configure(".", background=self.bg_base, foreground=self.text_dark, font=("Segoe UI", 10))
        
        # Estilo para Paneles de Contenido (LabelFrames blancos)
        self.estilo.configure("TitledPanel.TLabelframe", background=self.bg_panel, bordercolor=self.border_color, relief="solid", borderwidth=1)
        self.estilo.configure("TitledPanel.TLabelframe.Label", background=self.bg_panel, foreground="#7F8C8D", font=("Segoe UI", 10, "bold"))
        
        # Estilo para elementos de entrada
        self.estilo.configure("TCombobox", fieldbackground="#FFFFFF", background="#ECF0F1", bordercolor="#BDC3C7")
        
        # Estilos de Tabla (Treeview)
        self.estilo.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF", foreground=self.text_dark, rowheight=30, bordercolor="#ECF0F1", relief="flat")
        self.estilo.configure("Treeview.Heading", background=self.header_bg, foreground="#FFFFFF", font=("Segoe UI", 10, "bold"), relief="flat")
        self.estilo.map("Treeview.Heading", background=[("active", "#34495E")])

        self.init_db()
        
        # --- PANEL SUPERIOR (Encabezado Oscuro) ---
        frame_superior = tk.Frame(root, bg=self.bg_base, pady=10, padx=20)
        frame_superior.pack(fill=tk.X)
        
        banner_top = tk.Frame(frame_superior, bg=self.header_bg, height=55, relief="raised", bd=1)
        banner_top.pack(fill=tk.X)
        
        lbl_titulo_top = tk.Label(banner_top, text=" 🍣 SUSHI STOCK | GESTIÓN PREMIUM", font=("Segoe UI", 15, "bold"), bg=self.header_bg, fg="#FFFFFF")
        lbl_titulo_top.pack(side=tk.LEFT, padx=20, pady=12)
        
        self.lbl_reloj = tk.Label(banner_top, text="", font=("Segoe UI", 12, "bold"), bg=self.header_bg, fg="#ECF0F1")
        self.lbl_reloj.pack(side=tk.RIGHT, padx=20, pady=12)
        self.actualizar_reloj()
        
        # --- CONTENEDOR PRINCIPAL (PanedWindow sobre fondo gris base) ---
        main_container = tk.Frame(root, bg=self.bg_base)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        self.paned_window = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True)
        
        # COLUMNA IZQUIERDA: Contenedor del Menú (Fondo Gris Base)
        self.container_izquierdo = tk.Frame(self.paned_window, bg=self.bg_base)
        self.paned_window.add(self.container_izquierdo, weight=1)
        
        # TARJETA DEL MENÚ (Blanca)
        self.frame_menu = ttk.LabelFrame(self.container_izquierdo, text="  PANEL DE CONTROL  ", style="TitledPanel.TLabelframe")
        self.frame_menu.pack(fill=tk.BOTH, expand=True, padx=(0, 10), ipadx=5, ipady=5)
        
        # Función auxiliar para crear botones mejor organizados y compactos
        def crear_boton_menu(padre, texto, comando):
            btn = tk.Button(padre, text=texto, font=("Segoe UI", 11), bg=self.btn_bg, fg="#FFFFFF",
                            activebackground=self.btn_active, activeforeground="#FFFFFF",
                            relief="flat", bd=0, anchor="w", padx=20, pady=8, cursor="hand2", command=comando)
            btn.pack(fill=tk.X, pady=3, padx=10) 
            return btn

        # BLOQUE 1: Gestión de Insumos
        crear_boton_menu(self.frame_menu, "📥   Ingresar Insumo", self.abrir_ventana_alta)
        crear_boton_menu(self.frame_menu, "🔪   Registrar Consumo", self.abrir_ventana_registrar)
        crear_boton_menu(self.frame_menu, "✏️   Editar Artículo", self.abrir_ventana_editar_directo)
        crear_boton_menu(self.frame_menu, "❌   Eliminar Artículo", self.abrir_ventana_eliminar)
        
        # Separador visual
        tk.Frame(self.frame_menu, bg=self.border_color, height=1).pack(fill=tk.X, pady=10, padx=25)
        
        # BLOQUE 2: Consultas e Inventario
        crear_boton_menu(self.frame_menu, "📦   Inventario Físico", self.abrir_ventana_lista)
        crear_boton_menu(self.frame_menu, "🖨️   Imprimir Inventario", self.exportar_pdf_inventario)
        
        tk.Frame(self.frame_menu, bg=self.border_color, height=1).pack(fill=tk.X, pady=10, padx=25)
        
        # BLOQUE 3: Reportes
        crear_boton_menu(self.frame_menu, "📊   Consumos de Hoy", self.abrir_ventana_reporte_diario)
        crear_boton_menu(self.frame_menu, "📓   Bitácora General", self.abrir_ventana_historial)
        crear_boton_menu(self.frame_menu, "🛒   Pedido Urgente", self.abrir_ventana_pedido_hoy)
        
        # Espaciador para empujar el botón rojo hacia abajo
        tk.Frame(self.frame_menu, bg=self.bg_panel).pack(fill=tk.BOTH, expand=True)
        
        # --- BOTÓN DE CIERRE (ROJO DESTACADO) - UNA SOLA LÍNEA ---
        self.btn_Cierre = tk.Button(self.frame_menu, text="🛑 CERRAR JORNADA Y PEDIDO", 
                                   font=("Segoe UI", 11, "bold"), bg="#C0392B", fg="#FFFFFF", 
                                   activebackground="#A93226", activeforeground="#FFFFFF",
                                   relief="flat", bd=0, pady=12, command=self.finalizar_dia, cursor="hand2")
        self.btn_Cierre.pack(fill=tk.X, padx=10, pady=(5, 15), side=tk.BOTTOM)

        # COLUMNA DERECHA: Contenedor de la Gráfica (Fondo Gris Base)
        self.container_derecho = tk.Frame(self.paned_window, bg=self.bg_base)
        self.paned_window.add(self.container_derecho, weight=3)

        # TARJETA DE LA GRÁFICA (Blanca)
        self.frame_grafica = ttk.LabelFrame(self.container_derecho, text="  MÉTRICAS DE CONSUMO (HOY)  ", style="TitledPanel.TLabelframe")
        self.frame_grafica.pack(fill=tk.BOTH, expand=True, padx=(10, 0))

        self.canvas_grafica = None
        self.actualizar_grafica_barras(self.bg_panel)

    def init_db(self):
        conn = sqlite3.connect('inventario_avanzado.db')
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS articulos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                cantidad REAL NOT NULL,
                unidad TEXT NOT NULL,
                imagen_path TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                articulo_id INTEGER,
                tipo TEXT,
                cantidad REAL,
                fecha TEXT,
                FOREIGN KEY(articulo_id) REFERENCES articulos(id)
            )
        ''')
        conn.commit()
        conn.close()

    def actualizar_reloj(self):
        ahora = datetime.now().strftime("%A, %d de %B  |  %H:%M:%S")
        self.lbl_reloj.config(text=ahora.capitalize())
        self.root.after(1000, self.actualizar_reloj)

    def actualizar_grafica_barras(self, color_fondo_panel):
        if hasattr(self, 'canvas_grafica') and self.canvas_grafica:
            self.canvas_grafica.get_tk_widget().destroy()

        hoy = datetime.now().strftime('%Y-%m-%d')
        conn = sqlite3.connect('inventario_avanzado.db')
        cursor = conn.cursor()
        cursor.execute('''
            SELECT a.nombre, SUM(h.cantidad) FROM historial h 
            JOIN articulos a ON h.articulo_id = a.id 
            WHERE h.tipo = 'CONSUMO' AND h.fecha LIKE ? GROUP BY a.id
            ORDER BY SUM(h.cantidad) DESC
        ''', (f"{hoy}%",))
        datos = cursor.fetchall()
        conn.close()

        fig, ax = plt.subplots(figsize=(6, 4), dpi=100)
        fig.patch.set_facecolor(color_fondo_panel)
        ax.set_facecolor(color_fondo_panel)

        if datos:
            nombres = [d[0] for d in datos]
            cantidades = [d[1] for d in datos]
            
            barras = ax.bar(nombres, cantidades, color='#3498DB', edgecolor='none', width=0.6)
            
            ax.set_ylabel('Cantidad Descontada', fontsize=10, fontweight='bold', color=self.text_dark)
            ax.tick_params(axis='x', rotation=30, labelsize=9, colors=self.text_dark)
            ax.tick_params(axis='y', colors=self.text_dark)
            
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['left'].set_color('#BDC3C7')
            ax.spines['bottom'].set_color('#BDC3C7')
            
            for barra in barras:
                yval = barra.get_height()
                ax.text(barra.get_x() + barra.get_width()/2, yval + (max(cantidades)*0.02), round(yval, 2), 
                        ha='center', va='bottom', fontsize=9, fontweight='bold', color=self.text_dark)
        else:
            ax.text(0.5, 0.5, "✨ AÚN NO HAY MERMAS REGISTRADAS HOY ✨", 
                    ha='center', va='center', fontsize=12, color='#7F8C8D', weight='bold')
            ax.set_axis_off()

        plt.tight_layout()
        self.canvas_grafica = FigureCanvasTkAgg(fig, master=self.frame_grafica)
        self.canvas_grafica.draw()
        self.canvas_grafica.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        plt.close(fig)

    def centrar_y_rellenar(self, pil_image, target_size):
        pil_image.thumbnail(target_size, Image.Resampling.LANCZOS)
        new_image = Image.new('RGB', target_size, (255, 255, 255))
        new_image.paste(pil_image, ((target_size[0] - pil_image.size[0]) // 2,
                                     (target_size[1] - pil_image.size[1]) // 2))
        return new_image

    def seleccionar_imagen(self):
        ruta = filedialog.askopenfilename(filetypes=[("Imágenes", "*.png *.jpg *.jpeg")])
        if ruta:
            try:
                img_original = Image.open(ruta).convert('RGB')
                img_fijada = self.centrar_y_rellenar(img_original, (250, 200))
                foto_redimensionada = ImageTk.PhotoImage(img_fijada)
                if hasattr(self, 'lbl_previsualizacion'):
                    self.lbl_previsualizacion.config(image=foto_redimensionada, text="", padx=0, pady=0)
                    self.lbl_previsualizacion.image = foto_redimensionada
                self.ruta_imagen_temporal = ruta
                if hasattr(self, 'btn_img'):
                    self.btn_img.config(text="✅ Imagen Vinculada")
            except Exception as e:
                messagebox.showerror("Error", f"Error al procesar archivo: {e}")

    # ==========================================
    # 1. VENTANA DE AGREGAR INSUMO
    # ==========================================
    def abrir_ventana_alta(self):
        v_alta = tk.Toplevel(self.root)
        v_alta.title("📦 Ingresar Nueva Materia Prima")
        v_alta.geometry("530x560")
        v_alta.configure(bg=self.bg_panel) 
        
        frame_alta = ttk.LabelFrame(v_alta, text="  FORMULARIO DE INGRESO  ", style="TitledPanel.TLabelframe")
        frame_alta.pack(fill=tk.BOTH, expand=True, padx=25, pady=25, ipadx=10, ipady=10)
        frame_alta.columnconfigure(1, weight=1)
        
        ttk.Label(frame_alta, text="Nombre del Insumo:", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, padx=8, pady=12, sticky=tk.W)
        self.entry_nombre = ttk.Entry(frame_alta, font=("Segoe UI", 11))
        self.entry_nombre.grid(row=0, column=1, padx=8, pady=12, sticky=tk.EW)
        
        ttk.Label(frame_alta, text="Cantidad que Entra:", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, padx=8, pady=12, sticky=tk.W)
        self.entry_cantidad = ttk.Entry(frame_alta, font=("Segoe UI", 11))
        self.entry_cantidad.grid(row=1, column=1, padx=8, pady=12, sticky=tk.EW)
        
        ttk.Label(frame_alta, text="Unidad de Medida:", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, padx=8, pady=12, sticky=tk.W)
        self.combo_unidad = ttk.Combobox(frame_alta, values=["pieza", "kilo"], state="readonly", font=("Segoe UI", 10))
        self.combo_unidad.set("pieza")
        self.combo_unidad.grid(row=2, column=1, padx=8, pady=12, sticky=tk.EW)
        
        ttk.Label(frame_alta, text="Imagen (Opcional):", font=("Segoe UI", 10)).grid(row=3, column=0, padx=8, pady=12, sticky=tk.NW)
        self.lbl_previsualizacion = tk.Label(frame_alta, text="[ Sin Imagen ]", font=("Segoe UI", 9), bg="#F8F9F9", fg="#BDC3C7", relief="solid", borderwidth=1, padx=30, pady=30)
        self.lbl_previsualizacion.grid(row=3, column=1, padx=8, pady=12, sticky=tk.W)
        
        self.ruta_imagen_temporal = ""
        self.btn_img = tk.Button(frame_alta, text="🖼️ Examinar Imagen...", font=("Segoe UI", 10), bg="#EDF2F7", fg=self.text_dark, relief="flat", padx=15, pady=8, cursor="hand2", command=self.seleccionar_imagen)
        self.btn_img.grid(row=4, column=0, columnspan=2, pady=15, padx=8, sticky=tk.EW)
        
        tk.Button(frame_alta, text="GUARDAR NUEVO INSUMO", font=("Segoe UI", 11, "bold"), bg="#27AE60", fg="white", activebackground="#1E8449", relief="flat", pady=12, command=lambda: self.registrar_alta(v_alta), cursor="hand2").grid(row=5, column=0, columnspan=2, pady=(15, 5), padx=8, sticky=tk.EW)

    def registrar_alta(self, ventana_top):
        nombre = self.entry_nombre.get().strip().upper()
        cantidad_str = self.entry_cantidad.get().replace(',', '.') 
        text_unidad = self.combo_unidad.get()
        
        if not nombre or not cantidad_str:
            messagebox.showerror("Campos Vacíos", "Rellene el nombre del insumo y la cantidad.")
            return
        try:
            cantidad = float(cantidad_str)
        except ValueError:
            messagebox.showerror("Formato Incorrecto", "La cantidad ingresada debe ser un número.")
            return

        conn = sqlite3.connect('inventario_avanzado.db')
        cursor = conn.cursor()
        cursor.execute("SELECT id, cantidad FROM articulos WHERE nombre = ?", (nombre,))
        item = cursor.fetchone()
        
        if item:
            nuevo_stock = item[1] + cantidad
            cursor.execute("UPDATE articulos SET cantidad = ?, imagen_path = ? WHERE id = ?", (nuevo_stock, self.ruta_imagen_temporal, item[0]))
            art_id = item[0]
        else:
            cursor.execute("INSERT INTO articulos (nombre, cantidad, unidad, imagen_path) VALUES (?, ?, ?, ?)", (nombre, cantidad, text_unidad, self.ruta_imagen_temporal))
            art_id = cursor.lastrowid
            
        cursor.execute("INSERT INTO historial (articulo_id, tipo, cantidad, fecha) VALUES (?, 'ALTA', ?, ?)", (art_id, cantidad, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("Almacén Actualizado", f"✅ Ingreso exitoso de: {nombre}")
        self.actualizar_grafica_barras(self.bg_panel)
        ventana_top.destroy()

    def obtener_nombres_articulos(self):
        conn = sqlite3.connect('inventario_avanzado.db')
        cursor = conn.cursor()
        cursor.execute("SELECT nombre FROM articulos ORDER BY nombre ASC")
        nombres = [row[0] for row in cursor.fetchall()]
        conn.close()
        return nombres

    # ==========================================
    # 2. VENTANA DE REGISTRAR CONSUMO (CON AUTOSELECCIÓN)
    # ==========================================
    def abrir_ventana_registrar(self):
        v_reg = tk.Toplevel(self.root)
        v_reg.title("🔪 Registrar Consumo o Merma")
        v_reg.geometry("420x400")
        v_reg.configure(bg=self.bg_panel)

        # --- SECCIÓN DE BÚSQUEDA ---
        f_busqueda = tk.Frame(v_reg, bg=self.bg_panel)
        f_busqueda.pack(fill=tk.X, padx=40, pady=(25, 5))
        
        tk.Label(f_busqueda, text="🔍 Buscar:", font=("Segoe UI", 10, "bold"), bg=self.bg_panel, fg=self.text_dark).pack(side=tk.LEFT)
        ent_busqueda = ttk.Entry(f_busqueda, font=("Segoe UI", 11))
        ent_busqueda.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        # --- SECCIÓN DEL COMBOBOX (LISTA DESPLEGABLE) ---
        tk.Label(v_reg, text="Selecciona el Artículo:", font=("Segoe UI", 11, "bold"), bg=self.bg_panel, fg=self.text_dark).pack(pady=(15, 8))
        
        todos_los_articulos = self.obtener_nombres_articulos()
        combo_articulos = ttk.Combobox(v_reg, values=todos_los_articulos, state="readonly", font=("Segoe UI", 11))
        combo_articulos.pack(fill=tk.X, padx=40, pady=5)

        lbl_info_stock = tk.Label(v_reg, text="Stock actual: --", font=("Segoe UI", 10, "italic"), bg=self.bg_panel, fg="#7F8C8D")
        lbl_info_stock.pack()

        # Función para actualizar la etiqueta de stock
        def actualizar_info_stock(event=None):
            nombre = combo_articulos.get()
            if nombre:
                conn = sqlite3.connect('inventario_avanzado.db')
                cursor = conn.cursor()
                cursor.execute("SELECT cantidad, unidad FROM articulos WHERE nombre=?", (nombre,))
                res = cursor.fetchone()
                conn.close()
                if res:
                    lbl_info_stock.config(text=f"📦 Stock actual: {res[0]} {res[1]}(s)")

        combo_articulos.bind("<<ComboboxSelected>>", actualizar_info_stock)

        # Función para filtrar y AUTO-SELECCIONAR
        def filtrar_combobox(event):
            termino = ent_busqueda.get().strip().upper()
            if termino:
                filtrados = [art for art in todos_los_articulos if termino in art]
                combo_articulos.config(values=filtrados)
                
                # --- MAGIA DE AUTOSELECCIÓN ---
                if len(filtrados) == 1:
                    # Si solo queda 1 opción, se selecciona sola
                    combo_articulos.set(filtrados[0])
                    actualizar_info_stock()
                elif termino in todos_los_articulos:
                    # Si escribiste el nombre exacto, se selecciona sola
                    combo_articulos.set(termino)
                    actualizar_info_stock()
                else:
                    # Si no hay autoselección y lo que está seleccionado ya no encaja, lo limpiamos
                    if combo_articulos.get() not in filtrados:
                        combo_articulos.set("")
                        lbl_info_stock.config(text="Stock actual: --")
            else:
                # Si borran la búsqueda, mostramos todo y limpiamos la selección
                combo_articulos.config(values=todos_los_articulos)
                combo_articulos.set("")
                lbl_info_stock.config(text="Stock actual: --")

        ent_busqueda.bind("<KeyRelease>", filtrar_combobox)

        # --- SECCIÓN DE CANTIDAD A DESCONTAR ---
        tk.Label(v_reg, text="Cantidad a Descontar:", font=("Segoe UI", 11, "bold"), bg=self.bg_panel, fg=self.text_dark).pack(pady=(20, 8))
        entry_cantidad = ttk.Entry(v_reg, font=("Segoe UI", 12), justify="center")
        entry_cantidad.pack(fill=tk.X, padx=60, pady=5)

        def ejecutar_descuento():
            nombre = combo_articulos.get()
            cant_str = entry_cantidad.get().replace(',', '.')

            if not nombre or not cant_str:
                messagebox.showerror("Error", "Selecciona un artículo y escribe la cantidad.")
                return

            try:
                gasto = float(cant_str)
            except ValueError:
                messagebox.showerror("Error", "La cantidad debe ser un número válido.")
                return

            conn = sqlite3.connect('inventario_avanzado.db')
            cursor = conn.cursor()
            cursor.execute("SELECT id, cantidad, unidad FROM articulos WHERE nombre=?", (nombre,))
            res = cursor.fetchone()
            
            if not res:
                conn.close()
                return

            art_id, stock_actual, unidad = res
            if gasto > stock_actual:
                messagebox.showwarning("Sin Stock", f"Solo quedan {stock_actual} {unidad}(s). No puedes descontar {gasto}.")
                conn.close()
                return

            nuevo_stock = stock_actual - gasto
            cursor.execute("UPDATE articulos SET cantidad=? WHERE id=?", (nuevo_stock, art_id))
            cursor.execute("INSERT INTO historial (articulo_id, tipo, cantidad, fecha) VALUES (?, 'CONSUMO', ?, ?)", 
                           (art_id, gasto, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            conn.commit()
            conn.close()

            messagebox.showinfo("Éxito", f"📉 Se descontaron {gasto} {unidad}(s) de {nombre}.")
            self.actualizar_grafica_barras(self.bg_panel)
            v_reg.destroy()

        btn_guardar = tk.Button(v_reg, text="CONFIRMAR DESCUENTO", font=("Segoe UI", 11, "bold"), bg=self.btn_bg, fg="white", activebackground=self.btn_active, relief="flat", pady=12, command=ejecutar_descuento, cursor="hand2")
        btn_guardar.pack(fill=tk.X, padx=40, pady=25)

    # ==========================================
    # 3. VENTANA DE EDITAR ARTÍCULO
    # ==========================================
    def abrir_ventana_editar_directo(self):
        v_edit = tk.Toplevel(self.root)
        v_edit.title("📝 Editar Artículo")
        v_edit.geometry("480x680")
        v_edit.configure(bg=self.bg_panel)

        tk.Label(v_edit, text="Selecciona el Artículo a Editar:", font=("Segoe UI", 11, "bold"), bg=self.bg_panel).pack(pady=(25, 8))
        combo_articulos = ttk.Combobox(v_edit, values=self.obtener_nombres_articulos(), state="readonly", font=("Segoe UI", 11))
        combo_articulos.pack(fill=tk.X, padx=30, pady=5)

        f_datos_text = ttk.LabelFrame(v_edit, text="  DATOS DEL ARTÍCULO  ", style="TitledPanel.TLabelframe")
        f_datos_text.pack(fill=tk.X, padx=30, pady=20, ipadx=5, ipady=10)

        tk.Label(f_datos_text, text="Nuevo Nombre:", bg=self.bg_panel, font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, padx=10)
        ent_edit_nombre = ttk.Entry(f_datos_text, font=("Segoe UI", 11))
        ent_edit_nombre.pack(fill=tk.X, pady=(2, 12), padx=10)
        
        tk.Label(f_datos_text, text="Stock Físico Real:", bg=self.bg_panel, font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, padx=10)
        ent_edit_stock = ttk.Entry(f_datos_text, font=("Segoe UI", 11))
        ent_edit_stock.pack(fill=tk.X, pady=(2, 12), padx=10)
        
        tk.Label(f_datos_text, text="Unidad de Medida:", bg=self.bg_panel, font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, padx=10)
        combo_edit_unidad = ttk.Combobox(f_datos_text, values=["pieza", "kilo"], state="readonly", font=("Segoe UI", 11))
        combo_edit_unidad.pack(fill=tk.X, pady=(2, 12), padx=10)

        lbl_foto_edit_prev = tk.Label(v_edit, text="[ Sin Imagen ]", bg="#F8F9F9", fg="#BDC3C7", relief="solid", borderwidth=1, padx=30, pady=30)
        lbl_foto_edit_prev.pack(pady=5)

        self.ruta_imagen_edicion = ""
        nombre_original_seleccionado = {"nombre": ""}

        def cargar_datos_articulo(event):
            nombre = combo_articulos.get()
            if not nombre: return
            
            conn = sqlite3.connect('inventario_avanzado.db')
            cursor = conn.cursor()
            cursor.execute("SELECT nombre, cantidad, unidad, imagen_path FROM articulos WHERE nombre=?", (nombre,))
            res = cursor.fetchone()
            conn.close()

            if res:
                ent_edit_nombre.delete(0, tk.END)
                ent_edit_nombre.insert(0, res[0])
                ent_edit_stock.delete(0, tk.END)
                ent_edit_stock.insert(0, res[1])
                combo_edit_unidad.set(res[2])
                self.ruta_imagen_edicion = res[3] if res[3] else ""
                nombre_original_seleccionado["nombre"] = res[0]

                if self.ruta_imagen_edicion and os.path.exists(self.ruta_imagen_edicion):
                    try:
                        img_current = Image.open(self.ruta_imagen_edicion).convert('RGB')
                        img_edit_view = self.centrar_y_rellenar(img_current, (250, 200))
                        foto_current = ImageTk.PhotoImage(img_edit_view)
                        lbl_foto_edit_prev.config(image=foto_current, text="", padx=0, pady=0)
                        lbl_foto_edit_prev.image = foto_current
                    except:
                        lbl_foto_edit_prev.config(image="", text="[ Error archivo ]", padx=30, pady=30)
                else:
                    lbl_foto_edit_prev.config(image="", text="[ Sin Imagen ]", padx=30, pady=30)

        combo_articulos.bind("<<ComboboxSelected>>", cargar_datos_articulo)

        def seleccionar_nueva_imagen_edit():
            ruta = filedialog.askopenfilename(filetypes=[("Imágenes", "*.png *.jpg *.jpeg")])
            if ruta:
                try:
                    img_current = Image.open(ruta).convert('RGB')
                    img_edit_view = self.centrar_y_rellenar(img_current, (250, 200))
                    foto_current = ImageTk.PhotoImage(img_edit_view)
                    lbl_foto_edit_prev.config(image=foto_current, text="", padx=0, pady=0)
                    lbl_foto_edit_prev.image = foto_current
                    self.ruta_imagen_edicion = ruta
                except Exception as e:
                    messagebox.showerror("Error", f"Error al abrir la imagen: {e}")

        tk.Button(v_edit, text="🖼️ Cambiar Foto...", font=("Segoe UI", 10), bg="#EDF2F7", fg=self.text_dark, relief="flat", padx=15, pady=8, command=seleccionar_nueva_imagen_edit, cursor="hand2").pack(pady=10, padx=30, fill=tk.X)

        def guardar_cambios():
            viejo_nombre = nombre_original_seleccionado["nombre"]
            if not viejo_nombre:
                messagebox.showerror("Error", "Primero selecciona un artículo para editar.")
                return

            nuevo_nom = ent_edit_nombre.get().strip().upper()
            nuevo_stk_str = ent_edit_stock.get().strip().replace(',', '.')
            nueva_uni = combo_edit_unidad.get()
            
            if not nuevo_nom or not nuevo_stk_str:
                messagebox.showerror("Campos Vacíos", "No dejes casillas en blanco.")
                return
            try:
                nuevo_stk = float(nuevo_stk_str)
            except ValueError:
                messagebox.showerror("Error", "El stock debe ser numérico.")
                return

            conn = sqlite3.connect('inventario_avanzado.db')
            cursor = conn.cursor()
            try:
                cursor.execute("""
                    UPDATE articulos 
                    SET nombre = ?, cantidad = ?, unidad = ?, imagen_path = ?
                    WHERE nombre = ?
                """, (nuevo_nom, nuevo_stk, nueva_uni, self.ruta_imagen_edicion, viejo_nombre))
                conn.commit()
                messagebox.showinfo("Éxito", "✨ El artículo se actualizó correctamente.")
                self.actualizar_grafica_barras(self.bg_panel)
                v_edit.destroy()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Ya existe otro insumo con ese nombre.")
            finally:
                conn.close()

        tk.Button(v_edit, text="GUARDAR CAMBIOS", font=("Segoe UI", 11, "bold"), bg="#F39C12", fg="white", activebackground="#D68910", relief="flat", pady=12, command=guardar_cambios, cursor="hand2").pack(fill=tk.X, padx=30, pady=25)

    # ==========================================
    # 4. VENTANA DE ELIMINAR ARTÍCULO
    # ==========================================
    def abrir_ventana_eliminar(self):
        v_elim = tk.Toplevel(self.root)
        v_elim.title("🗑️ Eliminar Artículo")
        v_elim.geometry("400x240")
        v_elim.configure(bg=self.bg_panel)

        tk.Label(v_elim, text="Selecciona el Artículo a Borrar:", font=("Segoe UI", 11, "bold"), bg=self.bg_panel, fg="#C0392B").pack(pady=(30, 10))
        combo_articulos = ttk.Combobox(v_elim, values=self.obtener_nombres_articulos(), state="readonly", font=("Segoe UI", 11))
        combo_articulos.pack(fill=tk.X, padx=40, pady=10)

        def ejecutar_borrado():
            nombre = combo_articulos.get()
            if not nombre:
                messagebox.showerror("Error", "Selecciona un artículo de la lista.")
                return

            confirmar = messagebox.askyesno("Confirmar", f"⚠️ ¿Estás seguro de que deseas eliminar '{nombre}' permanentemente?")
            if confirmar:
                conn = sqlite3.connect('inventario_avanzado.db')
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM articulos WHERE nombre=?", (nombre,))
                art_id = cursor.fetchone()
                if art_id:
                    cursor.execute("DELETE FROM historial WHERE articulo_id=?", (art_id[0],))
                    cursor.execute("DELETE FROM articulos WHERE id=?", (art_id[0],))
                    conn.commit()
                conn.close()

                messagebox.showinfo("Borrado", f"🗑️ El artículo '{nombre}' ha sido eliminado.")
                self.actualizar_grafica_barras(self.bg_panel)
                v_elim.destroy()

        tk.Button(v_elim, text="ELIMINAR DEFINITIVAMENTE", font=("Segoe UI", 11, "bold"), bg="#C0392B", fg="white", activebackground="#A93226", relief="flat", pady=12, command=ejecutar_borrado, cursor="hand2").pack(fill=tk.X, padx=40, pady=25)

    # ==========================================
    # 5. VENTANA DE VER TABLA DE INVENTARIO
    # ==========================================
    def abrir_ventana_lista(self):
        try:
            v_lista = tk.Toplevel(self.root)
            v_lista.title("📋 Tabla Completa de Inventario")
            v_lista.geometry("900x680") 
            v_lista.configure(bg=self.bg_panel) 
            
            f_tabla = ttk.LabelFrame(v_lista, text="  EXISTENCIA REAL EN ALMACÉN  ", style="TitledPanel.TLabelframe")
            f_tabla.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)
            
            f_busqueda = tk.Frame(f_tabla, bg=self.bg_panel)
            f_busqueda.pack(fill=tk.X, padx=15, pady=20)
            
            tk.Label(f_busqueda, text="🔍 Buscar Insumo:", font=("Segoe UI", 11, "bold"), bg=self.bg_panel, fg=self.text_dark).pack(side=tk.LEFT, padx=5)
            ent_busqueda = ttk.Entry(f_busqueda, font=("Segoe UI", 11))
            ent_busqueda.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=15)

            tree = ttk.Treeview(f_tabla, columns=("N", "S", "M"), show="headings", style="Treeview")
            tree.heading("N", text="INGREDIENTE"); tree.heading("S", text="STOCK ACTUAL"); tree.heading("M", text="MEDIDA")
            tree.column("N", width=380); tree.column("S", width=160, anchor=tk.CENTER); tree.column("M", width=140, anchor=tk.CENTER)
            tree.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))

            def actualizar_tabla(evento=None):
                termino = ent_busqueda.get().strip().upper()
                for r in tree.get_children(): tree.delete(r)
                
                conn = sqlite3.connect('inventario_avanzado.db')
                cursor = conn.cursor()
                if termino:
                    cursor.execute("SELECT nombre, cantidad, unidad FROM articulos WHERE nombre LIKE ?", (f"%{termino}%",))
                else:
                    cursor.execute("SELECT nombre, cantidad, unidad FROM articulos")
                    
                for a in cursor.fetchall(): tree.insert("", tk.END, values=a)
                conn.close()

            ent_busqueda.bind("<KeyRelease>", actualizar_tabla)
            actualizar_tabla()
        except Exception as e:
            pass

    def exportar_pdf_inventario(self):
        conn = sqlite3.connect('inventario_avanzado.db')
        cursor = conn.cursor()
        cursor.execute("SELECT nombre, cantidad, unidad FROM articulos ORDER BY nombre ASC")
        datos_inventario = cursor.fetchall()
        conn.close()

        if not datos_inventario:
            messagebox.showwarning("Vacio", "No hay artículos registrados.")
            return

        fecha_str_pdf = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        nombre_archivo_pdf = f"Inventario_Actual_{fecha_str_pdf}.pdf"
        
        doc = SimpleDocTemplate(nombre_archivo_pdf, pagesize=letter)
        styles = getSampleStyleSheet()
        elementos = []
        
        elementos.append(Paragraph(f"<b>Inventario Físico Actual - Sushi Stock</b>", styles['Title']))
        elementos.append(Paragraph(f"Fecha de impresión: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", styles['Normal']))
        elementos.append(Spacer(1, 15))
        
        tabla_datos_pdf = [["Ingrediente en Almacén", "Stock Disponible", "Unidad"]]
        for nom, can, uni in datos_inventario:
            tabla_datos_pdf.append([nom, str(can), uni])
            
        t = Table(tabla_datos_pdf, colWidths=[250, 120, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(self.header_bg)),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 1, colors.grey)
        ]))
        elementos.append(t)
        doc.build(elementos)
        messagebox.showinfo("PDF Creado", f"🖨️ Reporte guardado como:\n\n{nombre_archivo_pdf}")

    def abrir_ventana_reporte_diario(self):
        try:
            v_rep = tk.Toplevel(self.root)
            v_rep.title("📊 Consumos de la Jornada Actual")
            v_rep.geometry("750x520") 
            v_rep.configure(bg=self.bg_panel) 
            
            f_rep = ttk.LabelFrame(v_rep, text="  RESUMEN DE GASTOS HOY  ", style="TitledPanel.TLabelframe")
            f_rep.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

            tree = ttk.Treeview(f_rep, columns=("Art", "Total", "Uni"), show="headings", style="Treeview")
            tree.heading("Art", text="INGREDIENTE OCUPADO")
            tree.heading("Total", text="TOTAL GASTADO HOY")
            tree.heading("Uni", text="UNIDAD")
            tree.column("Art", width=320); tree.column("Total", width=200, anchor=tk.CENTER); tree.column("Uni", width=140, anchor=tk.CENTER)
            tree.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
            
            hoy = datetime.now().strftime('%Y-%m-%d')
            conn = sqlite3.connect('inventario_avanzado.db')
            cursor = conn.cursor()
            cursor.execute('''
                SELECT a.nombre, SUM(h.cantidad), a.unidad FROM historial h 
                JOIN articulos a ON h.articulo_id = a.id 
                WHERE h.tipo = 'CONSUMO' AND h.fecha LIKE ? GROUP BY a.id
            ''', (f"{hoy}%",))
            for r in cursor.fetchall(): tree.insert("", tk.END, values=r)
            conn.close()
        except Exception as e:
            pass

    def abrir_ventana_historial(self):
        try:
            v_hist = tk.Toplevel(self.root)
            v_hist.title("📓 Bitácora General de Almacén")
            v_hist.geometry("900x580") 
            v_hist.configure(bg=self.bg_panel)
            
            f_hist = ttk.LabelFrame(v_hist, text="  HISTORIAL COMPLETO DE MOVIMIENTOS  ", style="TitledPanel.TLabelframe")
            f_hist.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

            tree = ttk.Treeview(f_hist, columns=("T", "A", "C", "F"), show="headings", style="Treeview")
            tree.heading("T", text="MOVIMIENTO")
            tree.heading("A", text="INGREDIENTE")
            tree.heading("C", text="CANTIDAD")
            tree.heading("F", text="FECHA Y HORA")
            tree.column("T", width=160, anchor=tk.CENTER); tree.column("A", width=280); tree.column("C", width=200, anchor=tk.CENTER); tree.column("F", width=220, anchor=tk.CENTER)
            tree.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
            
            conn = sqlite3.connect('inventario_avanzado.db')
            cursor = conn.cursor()
            cursor.execute('''
                SELECT h.tipo, a.nombre, h.cantidad || ' ' || a.unidad, h.fecha 
                FROM historial h JOIN articulos a ON h.articulo_id = a.id ORDER BY h.id DESC
            ''')
            for r in cursor.fetchall(): tree.insert("", tk.END, values=r)
            conn.close()
        except Exception as e:
            pass

    def abrir_ventana_pedido_hoy(self):
        try:
            v_pedido_hoy = tk.Toplevel(self.root)
            v_pedido_hoy.title("🛒 Generar Pedido de Hoy")
            v_pedido_hoy.geometry("750x680") 
            v_pedido_hoy.configure(bg=self.bg_panel) 
            
            f_main = ttk.LabelFrame(v_pedido_hoy, text="  SOLICITUD DE INSUMOS URGENTES  ", style="TitledPanel.TLabelframe")
            f_main.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

            lbl_info = tk.Label(f_main, text="SELECCIÓN DE INSUMOS PARA HOY", 
                                font=("Segoe UI", 12, "bold"), bg=self.bg_panel, fg=self.header_bg)
            lbl_info.pack(fill=tk.X, pady=(15, 5))

            f_busqueda = tk.Frame(f_main, bg=self.bg_panel)
            f_busqueda.pack(fill=tk.X, padx=20, pady=10)
            
            tk.Label(f_busqueda, text="🔍 Filtrar:", font=("Segoe UI", 10, "bold"), bg=self.bg_panel, fg=self.text_dark).pack(side=tk.LEFT, padx=5)
            ent_busqueda = ttk.Entry(f_busqueda, font=("Segoe UI", 11))
            ent_busqueda.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

            frame_lista = tk.Frame(f_main, bg=self.bg_panel)
            frame_lista.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

            canvas_scroll = tk.Canvas(frame_lista, bg=self.bg_panel, highlightthickness=0)
            scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=canvas_scroll.yview)
            scrollable_frame = tk.Frame(canvas_scroll, bg=self.bg_panel)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all"))
            )
            
            canvas_window = canvas_scroll.create_window((0, 0), window=scrollable_frame, anchor="nw")
            def on_canvas_configure(event):
                canvas_scroll.itemconfig(canvas_window, width=event.width)
            canvas_scroll.bind("<Configure>", on_canvas_configure)

            canvas_scroll.configure(yscrollcommand=scrollbar.set)

            conn = sqlite3.connect('inventario_avanzado.db')
            cursor = conn.cursor()
            cursor.execute("SELECT nombre, cantidad, unidad FROM articulos")
            lista_articulos = cursor.fetchall()
            conn.close()

            referencias_inputs = []
            filas_ui = [] 

            for nombre, stock, unidad in lista_articulos:
                frame_item = tk.Frame(scrollable_frame, bg="#FFFFFF", pady=10, relief="flat", highlightbackground="#F2F3F4", highlightthickness=1)
                frame_item.pack(fill=tk.X, expand=True, padx=10, pady=5)
                
                var_check = tk.BooleanVar()
                chk = ttk.Checkbutton(frame_item, text=f"  {nombre}  (Stock: {stock} {unidad}s)", variable=var_check)
                chk.pack(side=tk.LEFT, padx=10)
                
                frame_derecho_input = tk.Frame(frame_item, bg="#FFFFFF")
                frame_derecho_input.pack(side=tk.RIGHT, padx=15)
                
                lbl_cant = tk.Label(frame_derecho_input, text="Pedir:", font=("Segoe UI", 10, "bold"), bg="#FFFFFF")
                lbl_cant.pack(side=tk.LEFT, padx=5)
                
                ent_cant = ttk.Entry(frame_derecho_input, width=10, font=("Segoe UI", 11), justify="center")
                ent_cant.pack(side=tk.LEFT, padx=5)
                
                lbl_uni = tk.Label(frame_derecho_input, text=f"{unidad}(s)", width=8, bg="#FFFFFF", fg="#7F8C8D")
                lbl_uni.pack(side=tk.LEFT, padx=2)
                
                referencias_inputs.append((var_check, ent_cant, nombre, unidad))
                filas_ui.append((frame_item, nombre)) 

            canvas_scroll.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            def filtrar_elementos(evento=None):
                termino = ent_busqueda.get().strip().upper()
                for frame_item, nombre_item in filas_ui:
                    if termino in nombre_item:
                        frame_item.pack(fill=tk.X, expand=True, padx=10, pady=5)
                    else:
                        frame_item.pack_forget() 
                v_pedido_hoy.update_idletasks() 
                
            ent_busqueda.bind("<KeyRelease>", filtrar_elementos)

            def procesar_pedido_hoy():
                ingredientes_pedidos = []
                for var_chk, ent_cant, nombre, unidad in referencias_inputs:
                    if var_chk.get():
                        cant_pedida = ent_cant.get().strip()
                        if not cant_pedida:
                            messagebox.showerror("Faltan Cantidades", f"Ingresa cuánto pedir de '{nombre}'.")
                            return
                        ingredientes_pedidos.append((nombre, cant_pedida, unidad))
                
                if not ingredientes_pedidos:
                    messagebox.showwarning("Atención", "No marcaste ningún insumo. No se creará PDF.")
                    v_pedido_hoy.destroy()
                    return

                fecha_actual = datetime.now().strftime('%d/%m/%Y')
                hora_str = datetime.now().strftime('%H:%M:%S')
                timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                nombre_pdf_pedido = f"Pedido_Hoy_{timestamp}.pdf"
                
                doc = SimpleDocTemplate(nombre_pdf_pedido, pagesize=letter)
                styles = getSampleStyleSheet()
                elementos = []
                
                estilo_titulo = styles['Title']
                estilo_titulo.textColor = colors.HexColor(self.header_bg)
                
                elementos.append(Paragraph(f"<b>Pedido para el día de hoy: {fecha_actual}</b>", estilo_titulo))
                elementos.append(Paragraph(f"<b>Hora de Solicitud:</b> {hora_str}", styles['Normal']))
                elementos.append(Spacer(1, 20))
                
                mensaje_cuerpo = "A continuación, la lista de insumos requeridos:"
                elementos.append(Paragraph(f"<font size=12><b>{mensaje_cuerpo}</b></font>", styles['Normal']))
                elementos.append(Spacer(1, 15))
                
                for index, (nom, can, uni) in enumerate(ingredientes_pedidos, start=1):
                    item_texto = f"{index}) {nom} -> Cantidad solicitada: {can} {uni}(s)"
                    elementos.append(Paragraph(f"<font size=11>{item_texto}</font>", styles['Normal']))
                    elementos.append(Spacer(1, 6))
                    
                doc.build(elementos)
                messagebox.showinfo("Pedido Generado", f"✅ Pedido creado exitosamente.\n\nArchivo guardado como:\n{nombre_pdf_pedido}")
                v_pedido_hoy.destroy() 

            tk.Button(f_main, text="🖨️ GENERAR PDF DE PEDIDO", 
                                             font=("Segoe UI", 11, "bold"), bg=self.btn_bg, fg="#ffffff", 
                                             activebackground=self.btn_active, activeforeground="#ffffff",
                                             relief="flat", pady=12, command=procesar_pedido_hoy, cursor="hand2").pack(fill=tk.X, padx=20, pady=20, side=tk.BOTTOM)
        except Exception as e:
            pass

    def abrir_ventana_pedido_manana(self, nombre_base_reportes):
        try:
            v_pedido = tk.Toplevel(self.root)
            v_pedido.title("📝 Lista de Insumos para Mañana")
            v_pedido.geometry("750x680") 
            v_pedido.configure(bg=self.bg_panel) 
            
            f_main = ttk.LabelFrame(v_pedido, text="  REQUERIMIENTO DE PRODUCCIÓN  ", style="TitledPanel.TLabelframe")
            f_main.pack(fill=tk.BOTH, expand=True, padx=25, pady=25)

            lbl_info = tk.Label(f_main, text="SELECCIÓN DE PEDIDO PARA MAÑANA", 
                                font=("Segoe UI", 12, "bold"), bg=self.bg_panel, fg=self.header_bg, pady=15)
            lbl_info.pack(fill=tk.X)

            frame_lista_manana = tk.Frame(f_main, bg=self.bg_panel)
            frame_lista_manana.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

            canvas_scroll = tk.Canvas(frame_lista_manana, bg=self.bg_panel, highlightthickness=0)
            scrollbar = ttk.Scrollbar(frame_lista_manana, orient="vertical", command=canvas_scroll.yview)
            scrollable_frame = tk.Frame(canvas_scroll, bg=self.bg_panel)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas_scroll.configure(scrollregion=canvas_scroll.bbox("all"))
            )
            
            canvas_window = canvas_scroll.create_window((0, 0), window=scrollable_frame, anchor="nw")
            def on_canvas_configure(event):
                canvas_scroll.itemconfig(canvas_window, width=event.width)
            canvas_scroll.bind("<Configure>", on_canvas_configure)

            canvas_scroll.configure(yscrollcommand=scrollbar.set)

            conn = sqlite3.connect('inventario_avanzado.db')
            cursor = conn.cursor()
            cursor.execute("SELECT nombre, cantidad, unidad FROM articulos")
            lista_articulos = cursor.fetchall()
            conn.close()

            referencias_inputs = []

            for nombre, stock, unidad in lista_articulos:
                frame_item = tk.Frame(scrollable_frame, bg="#FFFFFF", pady=10, relief="flat", highlightbackground="#F2F3F4", highlightthickness=1)
                frame_item.pack(fill=tk.X, expand=True, padx=10, pady=5)
                
                var_check = tk.BooleanVar()
                chk = ttk.Checkbutton(frame_item, text=f"  {nombre}  (Stock: {stock} {unidad}s)", variable=var_check)
                chk.pack(side=tk.LEFT, padx=10)
                
                frame_derecho_input = tk.Frame(frame_item, bg="#FFFFFF")
                frame_derecho_input.pack(side=tk.RIGHT, padx=15)
                
                lbl_cant = tk.Label(frame_derecho_input, text="Pedir:", font=("Segoe UI", 10, "bold"), bg="#FFFFFF")
                lbl_cant.pack(side=tk.LEFT, padx=5)
                
                ent_cant = ttk.Entry(frame_derecho_input, width=10, font=("Segoe UI", 11), justify="center")
                ent_cant.pack(side=tk.LEFT, padx=5)
                
                lbl_uni = tk.Label(frame_derecho_input, text=f"{unidad}(s)", width=8, bg="#FFFFFF", fg="#7F8C8D")
                lbl_uni.pack(side=tk.LEFT, padx=2)
                
                referencias_inputs.append((var_check, ent_cant, nombre, unidad))

            canvas_scroll.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            def procesar_pedido_final():
                ingredientes_pedidos = []
                for var_chk, ent_cant, nombre, unidad in referencias_inputs:
                    if var_chk.get():
                        cant_pedida = ent_cant.get().strip()
                        if not cant_pedida:
                            messagebox.showerror("Faltan Cantidades", f"Ingresa cuánto pedir de '{nombre}'.")
                            return
                        ingredientes_pedidos.append((nombre, cant_pedida, unidad))
                
                if not ingredientes_pedidos:
                    messagebox.showwarning("Atención", "No marcaste ningún insumo. No se creará archivo de pedido.")
                    self.root.destroy() 
                    return

                hoy_str = datetime.now().strftime('%Y-%m-%d')
                hora_str = datetime.now().strftime('%H:%M:%S')
                nombre_pdf_pedido = f"Pedido_Manana_{hoy_str}.pdf"
                
                doc = SimpleDocTemplate(nombre_pdf_pedido, pagesize=letter)
                styles = getSampleStyleSheet()
                elementos = []
                
                estilo_titulo = styles['Title']
                estilo_titulo.textColor = colors.HexColor(self.header_bg)
                
                elementos.append(Paragraph(f"<b>REQUERIMIENTO DE PRODUCCIÓN (MAÑANA)</b>", estilo_titulo))
                elementos.append(Paragraph(f"<b>Fecha de Finalización:</b> {datetime.now().strftime('%d/%m/%Y')}", styles['Normal']))
                elementos.append(Paragraph(f"<b>Hora de Cierre:</b> {hora_str}", styles['Normal']))
                elementos.append(Spacer(1, 20))
                
                mensaje_cuerpo = "Buenas noches, para mañana se ocupará:"
                elementos.append(Paragraph(f"<font size=12><b>{mensaje_cuerpo}</b></font>", styles['Normal']))
                elementos.append(Spacer(1, 15))
                
                for index, (nom, can, uni) in enumerate(ingredientes_pedidos, start=1):
                    item_texto = f"{index}) {nom} -> Cantidad a ocupar: {can} {uni}(s)"
                    elementos.append(Paragraph(f"<font size=11>{item_texto}</font>", styles['Normal']))
                    elementos.append(Spacer(1, 6))
                    
                doc.build(elementos)
                
                messagebox.showinfo("Cierre Exitoso", 
                                    f"✨ Jornada Finalizada.\n\nDocumentos listos:\n"
                                    f"1. {nombre_base_reportes}.xlsx (Gastos)\n"
                                    f"2. {nombre_base_reportes}.pdf (Gastos)\n"
                                    f"3. {nombre_pdf_pedido} (Pedido Matutino)\n\nCerrando sistema.")
                
                self.root.destroy() 

            tk.Button(f_main, text="🏁 GENERAR REQUERIMIENTO Y CERRAR", 
                                             font=("Segoe UI", 11, "bold"), bg="#27AE60", fg="#ffffff", 
                                             activebackground="#1E8449", activeforeground="#ffffff",
                                             relief="flat", pady=12, command=procesar_pedido_final, cursor="hand2").pack(fill=tk.X, padx=20, pady=20, side=tk.BOTTOM)
        except Exception as e:
            pass

    def finalizar_dia(self):
        confirmar = messagebox.askyesno("Confirmar Cierre", "⚠️ ¿Desea cerrar el día actual? Esto compilará las mermas del día.")
        if not confirmar: return
        
        hoy = datetime.now().strftime('%Y-%m-%d')
        conn = sqlite3.connect('inventario_avanzado.db')
        cursor = conn.cursor()
        cursor.execute('''
            SELECT a.nombre, SUM(h.cantidad), a.unidad FROM historial h 
            JOIN articulos a ON h.articulo_id = a.id 
            WHERE h.tipo = 'CONSUMO' AND h.fecha LIKE ? GROUP BY a.id
        ''', (f"{hoy}%",))
        datos_hoy = cursor.fetchall()
        conn.close()
        
        nombre_base = f"Reporte_Consumo_{hoy}"
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Consumos"
        ws.append(["REPORTE GENERAL DE ARTICULOS GASTADOS"])
        ws.append(["Fecha:", hoy])
        ws.append([])
        ws.append(["Articulo", "Cantidad Gastada", "Medida"])
        if datos_hoy:
            for art, cant, uni in datos_hoy: ws.append([art, cant, uni])
        else:
            ws.append(["Sin mermas registradas hoy", "", ""])
        wb.save(f"{nombre_base}.xlsx")
        
        doc = SimpleDocTemplate(f"{nombre_base}.pdf", pagesize=letter)
        styles = getSampleStyleSheet()
        elementos = []
        elementos.append(Paragraph(f"<b>Cierre de Inventario Diario - Consumos</b>", styles['Title']))
        elementos.append(Paragraph(f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", styles['Normal']))
        elementos.append(Spacer(1, 15))
        
        tabla_datos = [["Articulo", "Cantidad Gastada", "Medida"]]
        if datos_hoy:
            for art, cant, uni in datos_hoy: tabla_datos.append([art, str(cant), uni])
        else:
            tabla_datos.append(["Sin gastos hoy", "0", "-"])
            
        t = Table(tabla_datos, colWidths=[220, 110, 100])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(self.header_bg)),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 1, colors.grey)
        ]))
        elementos.append(t)
        doc.build(elementos)
        
        self.abrir_ventana_pedido_manana(nombre_base)

if __name__ == "__main__":
    root = tk.Tk()
    root.configure(bg="#F0F2F5")
    app = InventarioApp(root)
    root.mainloop()
