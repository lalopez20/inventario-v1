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
        self.root.title("Sushi Stock Premium - Control de Almacen")
        
        # VENTANA PRINCIPAL EN PANTALLA COMPLETA
        try:
            self.root.state('zoomed') 
        except tk.TclError:
            self.root.attributes('-zoomed', True) 
            
        self.root.configure(bg="#fbfaf7")
        
        # CONFIGURACIÓN DE COLORES
        self.estilo = ttk.Style()
        self.estilo.theme_use("clam")
        
        self.estilo.configure(".", background="#fbfaf7", foreground="#2c3e50", font=("Helvetica", 10))
        self.estilo.configure("TLabelframe", background="#fbfaf7", bordercolor="#d5dbdb", relief="solid", borderwidth=1)
        self.estilo.configure("TLabelframe.Label", background="#fbfaf7", foreground="#e67e22", font=("Helvetica", 11, "bold"))
        self.estilo.configure("TCombobox", fieldbackground="#ffffff", background="#ecf0f1")
        
        self.estilo.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground="#2c3e50", rowheight=28)
        self.estilo.configure("Treeview.Heading", background="#2c3e50", foreground="#ffffff", font=("Helvetica", 10, "bold"))
        self.estilo.map("Treeview.Heading", background=[("active", "#34495e")])

        self.estilo.configure("BtnGuardar.TButton", background="#27ae60", foreground="#ffffff", font=("Helvetica", 10, "bold"), padding=6)
        self.estilo.map("BtnGuardar.TButton", background=[("active", "#219653")])
        
        self.estilo.configure("BtnOperacion.TButton", background="#3498db", foreground="#ffffff", font=("Helvetica", 10, "bold"), padding=8)
        self.estilo.map("BtnOperacion.TButton", background=[("active", "#2980b9")])
        
        self.estilo.configure("BtnExaminar.TButton", background="#7f8c8d", foreground="#ffffff", font=("Helvetica", 9), padding=5)
        self.estilo.map("BtnExaminar.TButton", background=[("active", "#95a5a6")])

        self.init_db()
        
        # --- PANEL SUPERIOR ---
        frame_superior = ttk.Frame(root, padding=12)
        frame_superior.pack(fill=tk.X)
        frame_superior.master.configure(bg="#fbfaf7")
        
        banner_top = tk.Frame(frame_superior, bg="#2c3e50", height=45)
        banner_top.pack(fill=tk.X)
        
        lbl_titulo_top = tk.Label(banner_top, text="  SUSHI STOCK | Gestion de Almacen", font=("Helvetica", 12, "bold"), bg="#2c3e50", fg="#ffffff")
        lbl_titulo_top.pack(side=tk.LEFT, padx=15, pady=8)
        
        self.lbl_reloj = tk.Label(banner_top, text="", font=("Helvetica", 11, "bold"), bg="#2c3e50", fg="#e67e22")
        self.lbl_reloj.pack(side=tk.RIGHT, padx=15, pady=8)
        self.actualizar_reloj()
        
        # --- CONTENEDOR PRINCIPAL ---
        paned_window = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # COLUMNA IZQUIERDA: Formulario
        frame_izquierdo = tk.Frame(paned_window, bg="#fbfaf7")
        paned_window.add(frame_izquierdo, weight=1)
        
        # COLUMNA DERECHA: Gráfica de Barras
        self.frame_grafica = ttk.LabelFrame(paned_window, text="  CONSUMOS DE HOY  ")
        paned_window.add(self.frame_grafica, weight=2)
        
        # --- FORMULARIO DE ALTA ---
        frame_alta = ttk.LabelFrame(frame_izquierdo, text="  INGRESO DE MATERIA PRIMA  ")
        frame_alta.pack(fill=tk.X, pady=10, ipadx=10, ipady=10)
        frame_alta.columnconfigure(1, weight=1)
        
        ttk.Label(frame_alta, text="Nombre del Insumo:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, padx=8, pady=6, sticky=tk.W)
        self.entry_nombre = ttk.Entry(frame_alta, font=("Helvetica", 11))
        self.entry_nombre.grid(row=0, column=1, padx=8, pady=6, sticky=tk.EW)
        
        ttk.Label(frame_alta, text="Cantidad que Entra:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, padx=8, pady=6, sticky=tk.W)
        self.entry_cantidad = ttk.Entry(frame_alta, font=("Helvetica", 11))
        self.entry_cantidad.grid(row=1, column=1, padx=8, pady=6, sticky=tk.EW)
        
        ttk.Label(frame_alta, text="Unidad de Medida:", font=("Helvetica", 10, "bold")).grid(row=2, column=0, padx=8, pady=6, sticky=tk.W)
        self.combo_unidad = ttk.Combobox(frame_alta, values=["pieza", "kilo"], state="readonly", font=("Helvetica", 10))
        self.combo_unidad.set("pieza")
        self.combo_unidad.grid(row=2, column=1, padx=8, pady=6, sticky=tk.EW)
        
        ttk.Label(frame_alta, text="Imagen de Referencia:", font=("Helvetica", 10)).grid(row=3, column=0, padx=8, pady=6, sticky=tk.NW)
        # Se remueven las restricciones de width y height para que la imagen no se deforme
        self.lbl_previsualizacion = tk.Label(frame_alta, text="[ Sin Imagen ]", font=("Helvetica", 9), bg="#f2f4f4", fg="#7f8c8d", relief="solid", borderwidth=1, padx=30, pady=30)
        self.lbl_previsualizacion.grid(row=3, column=1, padx=8, pady=6, sticky=tk.W)
        
        self.ruta_imagen_temporal = ""
        self.btn_img = ttk.Button(frame_alta, text="[+] Buscar Imagen de Insumo...", style="BtnExaminar.TButton", command=self.seleccionar_imagen)
        self.btn_img.grid(row=4, column=0, columnspan=2, pady=6, padx=8, sticky=tk.EW)
        
        ttk.Button(frame_alta, text="Guardar e Incrementar Inventario", style="BtnGuardar.TButton", command=self.registrar_alta).grid(row=5, column=0, columnspan=2, pady=6, padx=8, sticky=tk.EW)
        
        # --- PANEL DE ACCIONES ---
        frame_menu = ttk.LabelFrame(frame_izquierdo, text="  CONTROL Y CONSULTAS INTERNAS  ")
        frame_menu.pack(fill=tk.BOTH, expand=True, ipadx=10, ipady=10)
        
        ttk.Button(frame_menu, text="VER STOCK, EDITAR Y REGISTRAR GASTOS", style="BtnOperacion.TButton", command=self.abrir_ventana_lista).pack(fill=tk.X, pady=6, padx=8)
        ttk.Button(frame_menu, text="HISTORIAL DE CONSUMOS DE HOY", style="BtnOperacion.TButton", command=self.abrir_ventana_reporte_diario).pack(fill=tk.X, pady=6, padx=8)
        ttk.Button(frame_menu, text="BITACORA GENERAL DE MOVIMIENTOS", style="BtnOperacion.TButton", command=self.abrir_ventana_historial).pack(fill=tk.X, pady=6, padx=8)
        ttk.Button(frame_menu, text="CREAR PEDIDO PARA HOY (PDF)", style="BtnOperacion.TButton", command=self.abrir_ventana_pedido_hoy).pack(fill=tk.X, pady=6, padx=8)
        
        # --- BOTÓN DE CIERRE ---
        self.btn_Cierre = tk.Button(root, text="FINALIZAR JORNADA, GENERAR REPORTES Y PEDIDO MANANA", 
                                   font=("Helvetica", 11, "bold"), bg="#c0392b", fg="#ffffff", 
                                   activebackground="#962d22", activeforeground="#ffffff",
                                   relief="raised", bd=2, pady=12, command=self.finalizar_dia)
        self.btn_Cierre.pack(fill=tk.X, padx=20, pady=15)

        self.canvas_grafica = None
        self.actualizar_grafica_barras()

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
        ahora = datetime.now().strftime("%A, %d de %B - %H:%M:%S")
        self.lbl_reloj.config(text=ahora)
        self.root.after(1000, self.actualizar_reloj)

    def actualizar_grafica_barras(self):
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
        fig.patch.set_facecolor('#fbfaf7')
        ax.set_facecolor('#fbfaf7')

        if datos:
            nombres = [d[0] for d in datos]
            cantidades = [d[1] for d in datos]
            
            barras = ax.bar(nombres, cantidades, color='#3498db', edgecolor='#2980b9')
            
            ax.set_ylabel('Cantidad Descontada', fontsize=10, fontweight='bold', color='#2c3e50')
            ax.tick_params(axis='x', rotation=30, labelsize=9)
            
            for barra in barras:
                yval = barra.get_height()
                ax.text(barra.get_x() + barra.get_width()/2, yval + (max(cantidades)*0.02), round(yval, 2), ha='center', va='bottom', fontsize=9, fontweight='bold', color='#2c3e50')
        else:
            ax.text(0.5, 0.5, "AUN NO HAY REGISTROS\nDE MERMAS O GASTOS HOY", 
                    ha='center', va='center', fontsize=11, color='#2c3e50', weight='bold')
            ax.set_axis_off()

        plt.tight_layout()
        self.canvas_grafica = FigureCanvasTkAgg(fig, master=self.frame_grafica)
        self.canvas_grafica.draw()
        self.canvas_grafica.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        plt.close(fig)

    def centrar_y_rellenar(self, pil_image, target_size):
        pil_image.thumbnail(target_size, Image.Resampling.LANCZOS)
        new_image = Image.new('RGB', target_size, (255, 255, 255))
        new_image.paste(pil_image, ((target_size[0] - pil_image.size[0]) // 2,
                                     (target_size[1] - pil_image.size[1]) // 2))
        return new_image

    def seleccionar_imagen(self):
        ruta = filedialog.askopenfilename(filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.gif *.bmp")])
        if ruta:
            try:
                img_original = Image.open(ruta).convert('RGB')
                img_fijada = self.centrar_y_rellenar(img_original, (250, 200))
                foto_redimensionada = ImageTk.PhotoImage(img_fijada)
                # Al poner la imagen, quitamos el padding temporal
                self.lbl_previsualizacion.config(image=foto_redimensionada, text="", padx=0, pady=0)
                self.lbl_previsualizacion.image = foto_redimensionada
                self.ruta_imagen_temporal = ruta
                self.btn_img.config(text="Imagen Vinculada")
            except Exception as e:
                messagebox.showerror("Error", f"Error al procesar archivo: {e}")

    def registrar_alta(self):
        nombre = self.entry_nombre.get().strip().upper()
        cantidad_str = self.entry_cantidad.get().replace(',', '.') 
        text_unidad = self.combo_unidad.get()
        
        if not nombre or not cantidad_str:
            messagebox.showerror("Campos Vacios", "Rellene el nombre del insumo y la cantidad.")
            return
        try:
            cantidad = float(cantidad_str)
        except ValueError:
            messagebox.showerror("Formato Incorrecto", "La cantidad ingresada debe ser un numero.")
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
        
        self.entry_nombre.delete(0, tk.END)
        self.entry_cantidad.delete(0, tk.END)
        self.ruta_imagen_temporal = ""
        self.lbl_previsualizacion.config(image="", text="[ Sin Imagen ]", padx=30, pady=30)
        self.btn_img.config(text="Buscar Imagen de Insumo...")
        messagebox.showinfo("Almacen Actualizado", f"Ingreso exitoso de: {nombre}")
        self.actualizar_grafica_barras()

    def abrir_ventana_lista(self):
        try:
            v_lista = tk.Toplevel(self.root)
            v_lista.title("Inventario Fisico y Descuento de Insumos")
            v_lista.geometry("1150x640") 
            v_lista.configure(bg="#fbfaf7")
            
            f_tabla = ttk.LabelFrame(v_lista, text="  EXISTENCIA REAL EN COCINA  ")
            f_tabla.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=15, pady=15)
            
            f_acciones = ttk.LabelFrame(v_lista, text="  REGISTRAR CONSUMO / MERMA  ")
            f_acciones.pack(side=tk.RIGHT, fill=tk.Y, padx=15, pady=15, ipadx=15)
            
            f_busqueda = tk.Frame(f_tabla, bg="#fbfaf7")
            f_busqueda.pack(fill=tk.X, padx=10, pady=10)
            
            tk.Label(f_busqueda, text="Buscar Insumo:", font=("Helvetica", 10, "bold"), bg="#fbfaf7", fg="#2c3e50").pack(side=tk.LEFT, padx=5)
            ent_busqueda = ttk.Entry(f_busqueda, font=("Helvetica", 11))
            ent_busqueda.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Recuadro dinámico para la foto
            lbl_foto = tk.Label(f_acciones, text="[ Foto del Insumo ]", bg="#f2f4f4", fg="#7f8c8d", relief="solid", borderwidth=1, padx=30, pady=30)
            lbl_foto.pack(pady=10, padx=5)
            
            lbl_art = ttk.Label(f_acciones, text="Ningun insumo seleccionado", font=("Helvetica", 10, "bold"))
            lbl_art.pack(anchor=tk.W, pady=5, padx=5)
            
            lbl_medida = ttk.Label(f_acciones, text="Cantidad a descontar:")
            lbl_medida.pack(anchor=tk.W, padx=5)
            
            entry_gasto = ttk.Entry(f_acciones, font=("Helvetica", 11))
            entry_gasto.pack(fill=tk.X, pady=5, padx=5)
            
            tree = ttk.Treeview(f_tabla, columns=("N", "S", "M"), show="headings")
            tree.heading("N", text="INGREDIENTE"); tree.heading("S", text="STOCK ACTUAL"); tree.heading("M", text="MEDIDA")
            tree.column("N", width=320); tree.column("S", width=130, anchor=tk.CENTER); tree.column("M", width=120, anchor=tk.CENTER)
            tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

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

            def mostrar_detalles(event):
                sel = tree.selection()
                if not sel: return
                nombre, stock, unidad = tree.item(sel[0])['values']
                lbl_art.config(text=f"Insumo: {nombre}\nEn Almacen: {stock} {unidad}s")
                lbl_medida.config(text=f"Cantidad utilizada ({unidad}s):")
                
                conn = sqlite3.connect('inventario_avanzado.db')
                cursor = conn.cursor()
                cursor.execute("SELECT imagen_path FROM articulos WHERE nombre = ?", (nombre,))
                path = cursor.fetchone()[0]
                conn.close()
                
                if path and os.path.exists(path):
                    try:
                        img = Image.open(path).convert('RGB')
                        img_catalogo = self.centrar_y_rellenar(img, (250, 200))
                        foto = ImageTk.PhotoImage(img_catalogo)
                        lbl_foto.config(image=foto, text="", padx=0, pady=0)
                        lbl_foto.image = foto
                    except:
                        lbl_foto.config(image="", text="[ Error de Archivo ]", padx=30, pady=30)
                else:
                    lbl_foto.config(image="", text="[ Sin Foto ]", padx=30, pady=30)

            tree.bind("<<TreeviewSelect>>", mostrar_detalles)

            def ejecutar_baja():
                sel = tree.selection()
                if not sel: return
                nombre_sel, stock, unidad = tree.item(sel[0])['values']
                
                gasto_str = entry_gasto.get().replace(',', '.')
                try:
                    gasto = float(gasto_str)
                except:
                    messagebox.showerror("Error", "Ingrese un valor numerico.")
                    return
                    
                if gasto > float(stock):
                    messagebox.showwarning("Falta de Stock", f"No puedes descontar {gasto} porque solo quedan {stock}.")
                    return
                    
                conn = sqlite3.connect('inventario_avanzado.db')
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM articulos WHERE nombre = ?", (nombre_sel,))
                art_id = cursor.fetchone()[0]
                cursor.execute("UPDATE articulos SET cantidad = ? WHERE id = ?", (float(stock)-gasto, art_id))
                cursor.execute("INSERT INTO historial (articulo_id, tipo, cantidad, fecha) VALUES (?, 'CONSUMO', ?, ?)", 
                               (art_id, gasto, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                conn.commit()
                conn.close()
                
                entry_gasto.delete(0, tk.END)
                actualizar_tabla()
                lbl_foto.config(image="", text="[ Foto del Insumo ]", padx=30, pady=30)
                lbl_art.config(text="Ningun insumo seleccionado")
                self.actualizar_grafica_barras()
                messagebox.showinfo("Inventario Actualizado", f"Se desconto correctamente.")

            def abrir_ventana_editar():
                sel = tree.selection()
                if not sel:
                    messagebox.showwarning("Seleccion Requerida", "Selecciona un articulo.")
                    return
                
                nombre_actual, stock_actual, unidad_actual = tree.item(sel[0])['values']
                
                v_edit = tk.Toplevel(v_lista)
                v_edit.title(f"Editar: {nombre_actual}")
                v_edit.geometry("450x650") 
                v_edit.configure(bg="#fbfaf7")
                
                tk.Label(v_edit, text=f"Modificar Datos", font=("Helvetica", 11, "bold"), bg="#fbfaf7", fg="#e67e22").pack(pady=10)
                
                f_datos_text = tk.Frame(v_edit, bg="#fbfaf7")
                f_datos_text.pack(fill=tk.X, padx=20)
                
                tk.Label(f_datos_text, text="Nombre del Articulo:", bg="#fbfaf7").pack(anchor=tk.W)
                ent_edit_nombre = ttk.Entry(f_datos_text, font=("Helvetica", 10))
                ent_edit_nombre.insert(0, nombre_actual)
                ent_edit_nombre.pack(fill=tk.X, pady=5)
                
                tk.Label(f_datos_text, text="Stock Fisico Real:", bg="#fbfaf7").pack(anchor=tk.W)
                ent_edit_stock = ttk.Entry(f_datos_text, font=("Helvetica", 10))
                ent_edit_stock.insert(0, stock_actual)
                ent_edit_stock.pack(fill=tk.X, pady=5)
                
                tk.Label(f_datos_text, text="Unidad de Medida:", bg="#fbfaf7").pack(anchor=tk.W)
                combo_edit_unidad = ttk.Combobox(f_datos_text, values=["pieza", "kilo"], state="readonly", font=("Helvetica", 10))
                combo_edit_unidad.set(unidad_actual)
                combo_edit_unidad.pack(fill=tk.X, pady=5)
                
                tk.Label(v_edit, text="Imagen del Insumo:", font=("Helvetica", 10, "bold"), bg="#fbfaf7").pack(pady=10)
                
                conn = sqlite3.connect('inventario_avanzado.db')
                cursor = conn.cursor()
                cursor.execute("SELECT imagen_path FROM articulos WHERE nombre = ?", (nombre_actual,))
                path_imagen_actual_db = cursor.fetchone()[0]
                conn.close()
                
                self.ruta_imagen_edicion = path_imagen_actual_db if path_imagen_actual_db else ""
                
                # Aquí se arregla la foto de la ventana de edición quitando width y height
                lbl_foto_edit_prev = tk.Label(v_edit, text="[ Sin Imagen ]", bg="#f2f4f4", fg="#7f8c8d", relief="solid", borderwidth=1, padx=30, pady=30)
                lbl_foto_edit_prev.pack(pady=5)
                
                if path_imagen_actual_db and os.path.exists(path_imagen_actual_db):
                    try:
                        img_current = Image.open(path_imagen_actual_db).convert('RGB')
                        img_edit_view = self.centrar_y_rellenar(img_current, (250, 200))
                        foto_current = ImageTk.PhotoImage(img_edit_view)
                        lbl_foto_edit_prev.config(image=foto_current, text="", padx=0, pady=0)
                        lbl_foto_edit_prev.image = foto_current
                    except:
                        lbl_foto_edit_prev.config(text="[ Error archivo ]", padx=30, pady=30)
                        
                def seleccionar_nueva_imagen_edit():
                    ruta = filedialog.askopenfilename(filetypes=[("Imagenes", "*.png *.jpg *.jpeg *.gif *.bmp")])
                    if ruta:
                        try:
                            img_new_orig = Image.open(ruta).convert('RGB')
                            img_new_fijada = self.centrar_y_rellenar(img_new_orig, (250, 200))
                            foto_new_resized = ImageTk.PhotoImage(img_new_fijada)
                            lbl_foto_edit_prev.config(image=foto_new_resized, text="", padx=0, pady=0)
                            lbl_foto_edit_prev.image = foto_new_resized
                            
                            self.ruta_imagen_edicion = ruta 
                            btn_change_img.config(text="Foto Cambiada")
                        except Exception as e:
                            messagebox.showerror("Error", f"Error al procesar: {e}")
                            
                btn_change_img = ttk.Button(v_edit, text="Cambiar o Anadir Foto...", style="BtnExaminar.TButton", command=seleccionar_nueva_imagen_edit)
                btn_change_img.pack(pady=5, padx=20, fill=tk.X)
                
                def guardar_cambios_articulo():
                    nuevo_nom = ent_edit_nombre.get().strip().upper()
                    nuevo_stk_str = ent_edit_stock.get().strip().replace(',', '.')
                    nueva_uni = combo_edit_unidad.get()
                    nueva_ruta_img = self.ruta_imagen_edicion 
                    
                    if not nuevo_nom or not nuevo_stk_str:
                        messagebox.showerror("Campos Vacios", "No dejes casillas en blanco.")
                        return
                    try:
                        nuevo_stk = float(nuevo_stk_str)
                    except ValueError:
                        messagebox.showerror("Error", "El stock debe ser numerico.")
                        return
                        
                    conn = sqlite3.connect('inventario_avanzado.db')
                    cursor = conn.cursor()
                    try:
                        cursor.execute("""
                            UPDATE articulos 
                            SET nombre = ?, cantidad = ?, unidad = ?, imagen_path = ?
                            WHERE nombre = ?
                        """, (nuevo_nom, nuevo_stk, nueva_uni, nueva_ruta_img, nombre_actual))
                        conn.commit()
                        messagebox.showinfo("Modificado", "Se actualizo correctamente.")
                        v_edit.destroy()
                        ent_busqueda.delete(0, tk.END)
                        actualizar_tabla()
                        self.actualizar_grafica_barras()
                    except sqlite3.IntegrityError:
                        messagebox.showerror("Error", "Ya existe otro insumo con ese nombre.")
                    finally:
                        conn.close()
                
                tk.Button(v_edit, text="Guardar Cambios", font=("Helvetica", 10, "bold"), bg="#27ae60", fg="white", activebackground="#219653", relief="raised", pady=8, command=guardar_cambios_articulo).pack(fill=tk.X, padx=20, pady=15)

            btn_confirmar_baja = tk.Button(f_acciones, text="DESCONTAR DE EXISTENCIAS", font=("Helvetica", 10, "bold"), bg="#e67e22", fg="white", activebackground="#d35400", relief="raised", pady=6, command=ejecutar_baja)
            btn_confirmar_baja.pack(fill=tk.X, pady=10, padx=5)
            
            btn_editar_art = tk.Button(f_acciones, text="EDITAR PROPIEDADES / STOCK", font=("Helvetica", 10, "bold"), bg="#f1c40f", fg="#2c3e50", activebackground="#f39c12", relief="raised", pady=6, command=abrir_ventana_editar)
            btn_editar_art.pack(fill=tk.X, pady=5, padx=5)
            
            def exportar_pdf_inventario():
                conn = sqlite3.connect('inventario_avanzado.db')
                cursor = conn.cursor()
                cursor.execute("SELECT nombre, cantidad, unidad FROM articulos ORDER BY nombre ASC")
                datos_inventario = cursor.fetchall()
                conn.close()

                if not datos_inventario:
                    messagebox.showwarning("Vacio", "No hay articulos registrados.")
                    return

                fecha_str_pdf = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
                nombre_archivo_pdf = f"Inventario_Actual_{fecha_str_pdf}.pdf"
                
                doc = SimpleDocTemplate(nombre_archivo_pdf, pagesize=letter)
                styles = getSampleStyleSheet()
                elementos = []
                
                elementos.append(Paragraph(f"<b>Inventario Fisico Actual - Sushi Stock</b>", styles['Title']))
                elementos.append(Paragraph(f"Fecha de impresion: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", styles['Normal']))
                elementos.append(Spacer(1, 15))
                
                tabla_datos_pdf = [["Ingrediente en Almacen", "Stock Disponible", "Unidad"]]
                for nom, can, uni in datos_inventario:
                    tabla_datos_pdf.append([nom, str(can), uni])
                    
                t = Table(tabla_datos_pdf, colWidths=[250, 120, 100])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c3e50')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('GRID', (0,0), (-1,-1), 1, colors.grey)
                ]))
                elementos.append(t)
                doc.build(elementos)
                
                messagebox.showinfo("PDF Creado", f"Inventario guardado como:\n\n{nombre_archivo_pdf}")

            btn_pdf_inv = tk.Button(f_acciones, text="EXPORTAR INVENTARIO A PDF", font=("Helvetica", 10, "bold"), bg="#8e44ad", fg="white", activebackground="#732d91", relief="raised", pady=6, command=exportar_pdf_inventario)
            btn_pdf_inv.pack(fill=tk.X, pady=15, padx=5)
            
            actualizar_tabla()
        except Exception as e:
            messagebox.showerror("Error del Sistema", f"No se pudo cargar la ventana:\n{str(e)}")

    def abrir_ventana_reporte_diario(self):
        try:
            v_rep = tk.Toplevel(self.root)
            v_rep.title("Consumos de la Jornada Actual")
            v_rep.geometry("650x450") 
            v_rep.configure(bg="#fbfaf7")
            
            tree = ttk.Treeview(v_rep, columns=("Art", "Total", "Uni"), show="headings")
            tree.heading("Art", text="INGREDIENTE OCUPADO")
            tree.heading("Total", text="TOTAL GASTADO HOY")
            tree.heading("Uni", text="UNIDAD")
            tree.column("Art", width=260); tree.column("Total", width=160, anchor=tk.CENTER); tree.column("Uni", width=120, anchor=tk.CENTER)
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
            messagebox.showerror("Error del Sistema", f"No se pudo cargar la ventana:\n{str(e)}")

    def abrir_ventana_historial(self):
        try:
            v_hist = tk.Toplevel(self.root)
            v_hist.title("Bitacora General de Almacen")
            v_hist.geometry("800x520") 
            v_hist.configure(bg="#fbfaf7")
            
            tree = ttk.Treeview(v_hist, columns=("T", "A", "C", "F"), show="headings")
            tree.heading("T", text="MOVIMIENTO")
            tree.heading("A", text="INGREDIENTE")
            tree.heading("C", text="CANTIDAD")
            tree.heading("F", text="FECHA Y HORA EXACTA")
            tree.column("T", width=140, anchor=tk.CENTER); tree.column("A", width=220); tree.column("C", width=180, anchor=tk.CENTER); tree.column("F", width=190, anchor=tk.CENTER)
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
            messagebox.showerror("Error del Sistema", f"No se pudo cargar la ventana:\n{str(e)}")

    def abrir_ventana_pedido_hoy(self):
        try:
            v_pedido_hoy = tk.Toplevel(self.root)
            v_pedido_hoy.title("Generar Pedido de Hoy")
            v_pedido_hoy.geometry("700x620") 
            v_pedido_hoy.configure(bg="#fbfaf7")
            
            lbl_info = tk.Label(v_pedido_hoy, text="SELECCION DE INSUMOS URGENTES (HOY)", 
                                font=("Helvetica", 12, "bold"), bg="#fbfaf7", fg="#3498db")
            # Corrección de paddings asimétricos para Linux
            lbl_info.pack(fill=tk.X, pady=10)

            f_busqueda = tk.Frame(v_pedido_hoy, bg="#fbfaf7")
            f_busqueda.pack(fill=tk.X, padx=25, pady=10)
            
            tk.Label(f_busqueda, text="Filtrar ingrediente:", font=("Helvetica", 10, "bold"), bg="#fbfaf7", fg="#2c3e50").pack(side=tk.LEFT, padx=5)
            ent_busqueda = ttk.Entry(f_busqueda, font=("Helvetica", 11))
            ent_busqueda.pack(side=tk.LEFT, fill=tk.X, expand=True)

            frame_lista = tk.Frame(v_pedido_hoy, bg="#fbfaf7")
            frame_lista.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

            canvas_scroll = tk.Canvas(frame_lista, bg="#fbfaf7", highlightthickness=0)
            scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=canvas_scroll.yview)
            scrollable_frame = tk.Frame(canvas_scroll, bg="#fbfaf7")

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
                frame_item = tk.Frame(scrollable_frame, bg="#ffffff", pady=6, relief="groove", borderwidth=1)
                frame_item.pack(fill=tk.X, expand=True, padx=15, pady=3)
                
                var_check = tk.BooleanVar()
                chk = ttk.Checkbutton(frame_item, text=f" [ - ] {nombre}  (Stock actual: {stock} {unidad}s)", variable=var_check)
                chk.pack(side=tk.LEFT, padx=10)
                
                frame_derecho_input = tk.Frame(frame_item, bg="#ffffff")
                frame_derecho_input.pack(side=tk.RIGHT, padx=15)
                
                lbl_cant = ttk.Label(frame_derecho_input, text="Pedir:", font=("Helvetica", 10, "bold"))
                lbl_cant.pack(side=tk.LEFT, padx=5)
                
                ent_cant = ttk.Entry(frame_derecho_input, width=10, font=("Helvetica", 10))
                ent_cant.pack(side=tk.LEFT, padx=5)
                
                lbl_uni = ttk.Label(frame_derecho_input, text=f"{unidad}(s)", width=8)
                lbl_uni.pack(side=tk.LEFT, padx=2)
                
                referencias_inputs.append((var_check, ent_cant, nombre, unidad))
                filas_ui.append((frame_item, nombre)) 

            canvas_scroll.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            def filtrar_elementos(evento=None):
                termino = ent_busqueda.get().strip().upper()
                for frame_item, nombre_item in filas_ui:
                    if termino in nombre_item:
                        frame_item.pack(fill=tk.X, expand=True, padx=15, pady=3)
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
                            messagebox.showerror("Faltan Cantidades", f"Ingresa cuanto vas a pedir de '{nombre}'.")
                            return
                        ingredientes_pedidos.append((nombre, cant_pedida, unidad))
                
                if not ingredientes_pedidos:
                    messagebox.showwarning("Atencion", "No marcaste ningun insumo. No se creara PDF.")
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
                estilo_titulo.textColor = colors.HexColor('#2c3e50')
                
                elementos.append(Paragraph(f"<b>Pedido para el dia de hoy: {fecha_actual}</b>", estilo_titulo))
                elementos.append(Paragraph(f"<b>Hora de Solicitud:</b> {hora_str}", styles['Normal']))
                elementos.append(Spacer(1, 20))
                
                mensaje_cuerpo = "A continuacion, la lista de insumos requeridos:"
                elementos.append(Paragraph(f"<font size=12><b>{mensaje_cuerpo}</b></font>", styles['Normal']))
                elementos.append(Spacer(1, 15))
                
                for index, (nom, can, uni) in enumerate(ingredientes_pedidos, start=1):
                    item_texto = f"{index}) {nom} -> Cantidad solicitada: {can} {uni}(s)"
                    elementos.append(Paragraph(f"<font size=11>{item_texto}</font>", styles['Normal']))
                    elementos.append(Spacer(1, 6))
                    
                doc.build(elementos)
                
                messagebox.showinfo("Pedido Generado", f"El pedido se ha creado exitosamente.\n\nArchivo guardado como:\n{nombre_pdf_pedido}")
                v_pedido_hoy.destroy() 

            btn_confirmar_pedido = tk.Button(v_pedido_hoy, text="GENERAR PDF DE PEDIDO PARA HOY", 
                                             font=("Helvetica", 10, "bold"), bg="#3498db", fg="#ffffff", 
                                             activebackground="#2980b9", activeforeground="#ffffff",
                                             pady=12, command=procesar_pedido_hoy)
            btn_confirmar_pedido.pack(fill=tk.X, padx=20, pady=15, side=tk.BOTTOM)
        
        except Exception as e:
            messagebox.showerror("Error del Sistema", f"Ocurrio un error al intentar dibujar la ventana:\n\n{str(e)}")

    def abrir_ventana_pedido_manana(self, nombre_base_reportes):
        try:
            v_pedido = tk.Toplevel(self.root)
            v_pedido.title("Lista de Insumos para Manana")
            v_pedido.geometry("700x620") 
            v_pedido.configure(bg="#fbfaf7")
            
            lbl_info = tk.Label(v_pedido, text="SELECCION DE PEDIDO PARA MANANA", 
                                font=("Helvetica", 12, "bold"), bg="#fbfaf7", fg="#e67e22", pady=15)
            lbl_info.pack(fill=tk.X)

            frame_lista_manana = tk.Frame(v_pedido, bg="#fbfaf7")
            frame_lista_manana.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

            canvas_scroll = tk.Canvas(frame_lista_manana, bg="#fbfaf7", highlightthickness=0)
            scrollbar = ttk.Scrollbar(frame_lista_manana, orient="vertical", command=canvas_scroll.yview)
            scrollable_frame = tk.Frame(canvas_scroll, bg="#fbfaf7")

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
                frame_item = tk.Frame(scrollable_frame, bg="#ffffff", pady=6, relief="groove", borderwidth=1)
                frame_item.pack(fill=tk.X, expand=True, padx=15, pady=3)
                
                var_check = tk.BooleanVar()
                chk = ttk.Checkbutton(frame_item, text=f" [ - ] {nombre}  (Hay en stock: {stock} {unidad}s)", variable=var_check)
                chk.pack(side=tk.LEFT, padx=10)
                
                frame_derecho_input = tk.Frame(frame_item, bg="#ffffff")
                frame_derecho_input.pack(side=tk.RIGHT, padx=15)
                
                lbl_cant = ttk.Label(frame_derecho_input, text="Pedir:", font=("Helvetica", 10, "bold"))
                lbl_cant.pack(side=tk.LEFT, padx=5)
                
                ent_cant = ttk.Entry(frame_derecho_input, width=10, font=("Helvetica", 10))
                ent_cant.pack(side=tk.LEFT, padx=5)
                
                lbl_uni = ttk.Label(frame_derecho_input, text=f"{unidad}(s)", width=8)
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
                            messagebox.showerror("Faltan Cantidades", f"Ingresa cuanto vas a ocupar de '{nombre}'.")
                            return
                        ingredientes_pedidos.append((nombre, cant_pedida, unidad))
                
                if not ingredientes_pedidos:
                    messagebox.showwarning("Atencion", "No marcaste ningun insumo. No se creara archivo de pedido.")
                    self.root.destroy() 
                    return

                hoy_str = datetime.now().strftime('%Y-%m-%d')
                hora_str = datetime.now().strftime('%H:%M:%S')
                nombre_pdf_pedido = f"Pedido_Manana_{hoy_str}.pdf"
                
                doc = SimpleDocTemplate(nombre_pdf_pedido, pagesize=letter)
                styles = getSampleStyleSheet()
                elementos = []
                
                estilo_titulo = styles['Title']
                estilo_titulo.textColor = colors.HexColor('#2c3e50')
                
                elementos.append(Paragraph(f"<b>REQUERIMIENTO DE PRODUCCION (MANANA)</b>", estilo_titulo))
                elementos.append(Paragraph(f"<b>Fecha de Finalizacion:</b> {datetime.now().strftime('%d/%m/%Y')}", styles['Normal']))
                elementos.append(Paragraph(f"<b>Hora de Cierre:</b> {hora_str}", styles['Normal']))
                elementos.append(Spacer(1, 20))
                
                mensaje_cuerpo = "Buenas noches, para manana se ocupara:"
                elementos.append(Paragraph(f"<font size=12><b>{mensaje_cuerpo}</b></font>", styles['Normal']))
                elementos.append(Spacer(1, 15))
                
                for index, (nom, can, uni) in enumerate(ingredientes_pedidos, start=1):
                    item_texto = f"{index}) {nom} -> Cantidad a ocupar: {can} {uni}(s)"
                    elementos.append(Paragraph(f"<font size=11>{item_texto}</font>", styles['Normal']))
                    elementos.append(Spacer(1, 6))
                    
                doc.build(elementos)
                
                messagebox.showinfo("Cierre Guardado", 
                                    f"Operacion completada.\n\nDocumentos listos:\n"
                                    f"1. {nombre_base_reportes}.xlsx (Gastos)\n"
                                    f"2. {nombre_base_reportes}.pdf (Gastos)\n"
                                    f"3. {nombre_pdf_pedido} (Pedido Matutino)\n\nCerrando software.")
                
                self.root.destroy() 

            btn_confirmar_pedido = tk.Button(v_pedido, text="GENERAR REQUERIMIENTO EN PDF Y CERRAR", 
                                             font=("Helvetica", 10, "bold"), bg="#27ae60", fg="#ffffff", 
                                             activebackground="#219653", activeforeground="#ffffff",
                                             pady=12, command=procesar_pedido_final)
            btn_confirmar_pedido.pack(fill=tk.X, padx=20, pady=15, side=tk.BOTTOM)
            
        except Exception as e:
            messagebox.showerror("Error del Sistema", f"No se pudo cargar la ventana:\n{str(e)}")

    def finalizar_dia(self):
        confirmar = messagebox.askyesno("Confirmar Cierre", "Desea cerrar el dia actual? Esto compilara las mermas del dia.")
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
        
        # Guardar Excel de Gastos
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
        
        # Guardar PDF de Gastos
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
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2c3e50')),
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
    app = InventarioApp(root)
    root.mainloop()
