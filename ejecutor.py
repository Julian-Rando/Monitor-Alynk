import tkinter as tk
from tkinter import messagebox
import xlwings as xw
from datetime import datetime
from PIL import Image, ImageTk 
import os
import sys  # <--- NUEVO: Para leer parámetros

# --- PROCESAMIENTO DE PARÁMETROS ---
# Lógica: Si me pasan una ruta, la uso. Si no, uso el default.
if len(sys.argv) > 1:
    # El usuario pasó una ruta (ej: "C:\Datos\Finanzas.xlsx")
    ruta_completa = sys.argv[1]
    # xlwings conecta a libros ABIERTOS por su nombre de archivo, no por la ruta entera.
    # Extraemos solo el nombre (ej: "Finanzas.xlsx") de la ruta completa.
    NOMBRE_ARCHIVO_EXCEL = os.path.basename(ruta_completa)
    print(f"Modo Parámetro: Buscando '{NOMBRE_ARCHIVO_EXCEL}'")
else:
    # No pasaron nada, usamos el default
    NOMBRE_ARCHIVO_EXCEL = 'excel.xlsx'
    print(f"Modo Default: Buscando '{NOMBRE_ARCHIVO_EXCEL}'")

# --- RESTO DE LA CONFIGURACIÓN ---
NOMBRE_HOJA = 'MEP-CANJE-Arbitraje'
NOMBRE_LOGO = 'alynk logo.png' 

# TIPOS DE DATO: "P" (Precio $) | "S" (Spread %)
MAPA_DATOS = {
    "AL30": {
        "C":   ("B3", "P"), "V": ("E3", "P"), "CCL": ("H3", "P")
    },
    "GD30": {
        "C":   ("B4", "P"), "V": ("E4", "P"), "CCL": ("H4", "P")
    },
    "SEP": None,
    "CANJE COMPRA": {
        "C":   ("B7", "S"), "V": ("E7", "S"), "CCL": ("H7", "P")
    },
    "CANJE VENTA": {
        "C":   ("B8", "S"), "V": ("E8", "S"), "CCL": ("H8", "P")
    }
}

TIEMPO_REFRESCO = 1000 

# --- COLORES ---
COLOR_FONDO = "#121212"
COLOR_TABLA_BG = "#1e1e1e"
COLOR_HEADER = "#263238"
COLOR_PRECIO = "#00e676"
COLOR_SPREAD = "#4fc3f7"
COLOR_TEXTO = "#ffffff"

class MonitorProApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"ALYNK - Monitor ({NOMBRE_ARCHIVO_EXCEL})") # Muestra el archivo en el título
        self.root.geometry("1000x650") 
        self.root.configure(bg=COLOR_FONDO)

        # HEADER
        header_frame = tk.Frame(root, bg=COLOR_FONDO, pady=15)
        header_frame.pack(fill="x")
        
        tk.Label(header_frame, text="MONITOR DE MERCADO EN VIVO", font=("Arial", 18, "bold"), 
                 bg=COLOR_FONDO, fg="white").pack()
        
        tk.Label(header_frame, text="Conexión directa a memoria de mercado", font=("Arial", 10), 
                 bg=COLOR_FONDO, fg="#b0bec5").pack()

        # TABLA
        self.frame_tabla = tk.Frame(root, bg=COLOR_FONDO)
        self.frame_tabla.pack(expand=True, fill="both", padx=30, pady=10)
        
        self.labels_cache = {}
        self.construir_tabla()
        
        # LOGO
        self.agregar_logo_pie()

        # STATUS BAR
        self.lbl_status = tk.Label(root, text="Iniciando...", font=("Consolas", 9), bg="#000000", fg="gray", pady=5)
        self.lbl_status.pack(side="bottom", fill="x")

        self.actualizar_datos()

    def agregar_logo_pie(self):
        # Busca el logo en la misma carpeta donde está el script, no importa desde donde se ejecute
        base_path = os.path.dirname(os.path.abspath(__file__))
        ruta_logo = os.path.join(base_path, NOMBRE_LOGO)
        
        if not os.path.exists(ruta_logo):
            return
        try:
            pil_image = Image.open(ruta_logo)
            target_height = 70
            aspect_ratio = pil_image.width / pil_image.height
            target_width = int(target_height * aspect_ratio)
            resized_image = pil_image.resize((target_width, target_height), Image.Resampling.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(resized_image)
            tk.Label(self.root, image=self.logo_img, bg=COLOR_FONDO).pack(side="bottom", pady=(10, 20))
        except: pass

    def construir_tabla(self):
        headers = ["INSTRUMENTO", "COMPRA", "VENTA", "CABLE (CCL)"]
        for col, text in enumerate(headers):
            tk.Label(self.frame_tabla, text=text, font=("Arial", 11, "bold"), 
                     bg=COLOR_HEADER, fg="#cfd8dc", pady=8).grid(row=0, column=col, sticky="nsew", padx=1, pady=0)

        row_idx = 1
        for key, config in MAPA_DATOS.items():
            if key == "SEP":
                tk.Label(self.frame_tabla, text="", bg=COLOR_FONDO, height=1).grid(row=row_idx, columnspan=4)
                row_idx += 1
                continue

            es_fila_precio = config["C"][1] == "P"
            bg_row = COLOR_TABLA_BG
            font_inst = ("Arial", 14, "bold") if es_fila_precio else ("Arial", 12, "bold")
            fg_inst = "#ffffff" if es_fila_precio else "#b0bec5"

            tk.Label(self.frame_tabla, text=key, font=font_inst, 
                     bg=bg_row, fg=fg_inst, anchor="w", padx=15).grid(row=row_idx, column=0, sticky="nsew", padx=1, pady=2)

            self.labels_cache[(key, "C")] = self.crear_celda(row_idx, 1, config["C"][1], bg_row)
            self.labels_cache[(key, "V")] = self.crear_celda(row_idx, 2, config["V"][1], bg_row)
            self.labels_cache[(key, "CCL")] = self.crear_celda(row_idx, 3, config["CCL"][1], bg_row)
            row_idx += 1

        for i in range(4): self.frame_tabla.grid_columnconfigure(i, weight=1)

    def crear_celda(self, r, c, tipo, bg_color):
        fg_color = COLOR_PRECIO if tipo == "P" else COLOR_SPREAD
        font_val = ("Consolas", 18, "bold") if tipo == "P" else ("Consolas", 14)
        lbl = tk.Label(self.frame_tabla, text="---", font=font_val, bg=bg_color, fg=fg_color)
        lbl.grid(row=r, column=c, sticky="nsew", padx=1, pady=2)
        return lbl

    def actualizar_datos(self):
        try:
            libro = None
            try: libro = xw.books[NOMBRE_ARCHIVO_EXCEL]
            except:
                try: libro = xw.books.active
                except: pass
            
            if not libro:
                self.lbl_status.config(text=f"❌ NO SE DETECTA '{NOMBRE_ARCHIVO_EXCEL}' ABIERTO", fg="orange")
                self.root.after(2000, self.actualizar_datos)
                return

            hoja = libro.sheets[NOMBRE_HOJA]
            
            for key, config in MAPA_DATOS.items():
                if key == "SEP": continue
                
                raw_c = hoja.range(config["C"][0]).value
                raw_v = hoja.range(config["V"][0]).value
                raw_ccl = hoja.range(config["CCL"][0]).value
                
                self.set_valor(key, "C", raw_c, config["C"][1])
                self.set_valor(key, "V", raw_v, config["V"][1])
                self.set_valor(key, "CCL", raw_ccl, config["CCL"][1])

            self.lbl_status.config(text=f"🟢 ONLINE: {NOMBRE_ARCHIVO_EXCEL} - {datetime.now().strftime('%H:%M:%S')}", fg="#00c853")

        except Exception as e:
            self.lbl_status.config(text=f"⚠️ ERROR: {str(e)}", fg="red")

        self.root.after(TIEMPO_REFRESCO, self.actualizar_datos)

    def set_valor(self, key, col_type, valor, tipo_dato):
        label = self.labels_cache[(key, col_type)]
        texto = "-"
        if valor is not None and valor != "":
            try:
                val_float = float(valor)
                if tipo_dato == "P":
                    texto = f"$ {val_float:,.2f}"
                elif tipo_dato == "S":
                    texto = f"{val_float * 100:.2f}%"
            except:
                texto = str(valor)
        
        if label.cget("text") != texto:
            label.config(text=texto)

if __name__ == "__main__":
    try:
        import PIL
    except ImportError:
        messagebox.showerror("Error", "Falta pillow. Ejecuta: pip install pillow")
        exit()

    root = tk.Tk()
    app = MonitorProApp(root)
    root.mainloop()