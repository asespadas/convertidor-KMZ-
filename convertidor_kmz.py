import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import os
import re
import shutil
import sys
import threading
from datetime import datetime

class ConvertidorKMZ:
    def __init__(self, root):
        self.root = root
        self.root.title("🔄 Convertidor KMZ/KML a Excel")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Configurar estilo
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.archivo_seleccionado = None
        self.carpeta_destino = os.path.join(os.path.expanduser("~"), "Desktop", "KMZ_Exportaciones")
        
        # Crear interfaz
        self.crear_interfaz()
        
    def crear_interfaz(self):
        # Título
        titulo = tk.Label(self.root, text="🔄 Convertidor KMZ/KML a Excel", 
                         font=("Arial", 20, "bold"), bg='#f0f0f0', fg='#2c3e50')
        titulo.pack(pady=20)
        
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Sección de selección de archivo
        archivo_frame = ttk.LabelFrame(main_frame, text="📁 Archivo de entrada", padding="10")
        archivo_frame.pack(fill=tk.X, pady=10)
        
        btn_seleccionar = ttk.Button(archivo_frame, text="Seleccionar archivo KMZ/KML", 
                                     command=self.seleccionar_archivo)
        btn_seleccionar.pack(side=tk.LEFT, padx=5)
        
        self.lbl_archivo = ttk.Label(archivo_frame, text="Ningún archivo seleccionado", 
                                     foreground="gray")
        self.lbl_archivo.pack(side=tk.LEFT, padx=10)
        
        # Sección de carpeta de destino
        destino_frame = ttk.LabelFrame(main_frame, text="📂 Carpeta de destino", padding="10")
        destino_frame.pack(fill=tk.X, pady=10)
        
        btn_carpeta = ttk.Button(destino_frame, text="Seleccionar carpeta", 
                                 command=self.seleccionar_carpeta)
        btn_carpeta.pack(side=tk.LEFT, padx=5)
        
        self.lbl_carpeta = ttk.Label(destino_frame, text=self.carpeta_destino, 
                                     foreground="blue")
        self.lbl_carpeta.pack(side=tk.LEFT, padx=10)
        
        # Botón de procesar
        btn_procesar = ttk.Button(main_frame, text="🚀 PROCESAR ARCHIVO", 
                                  command=self.iniciar_procesamiento,
                                  style="Accent.TButton")
        btn_procesar.pack(pady=20)
        
        # Área de log
        log_frame = ttk.LabelFrame(main_frame, text="📋 Progreso", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, height=15, 
                                                  font=("Consolas", 9))
        self.log_area.pack(fill=tk.BOTH, expand=True)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=5)
        
        # Frame de créditos
        creditos_frame = ttk.Frame(main_frame)
        creditos_frame.pack(fill=tk.X, pady=5)
        
        creditos = tk.Label(creditos_frame, 
                           text="Creado por: Saul Armando Espinoza Garcia | v1.0", 
                           font=("Arial", 8), fg='gray')
        creditos.pack(side=tk.RIGHT)
        
    def log(self, mensaje, tipo="info"):
        """Agregar mensaje al área de log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        colores = {
            "info": "black",
            "success": "green",
            "error": "red",
            "warning": "orange"
        }
        color = colores.get(tipo, "black")
        
        self.log_area.insert(tk.END, f"[{timestamp}] ", "timestamp")
        self.log_area.insert(tk.END, f"{mensaje}\n", tipo)
        self.log_area.tag_config("timestamp", foreground="gray")
        self.log_area.tag_config("info", foreground="black")
        self.log_area.tag_config("success", foreground="green")
        self.log_area.tag_config("error", foreground="red")
        self.log_area.tag_config("warning", foreground="orange")
        
        self.log_area.see(tk.END)
        self.root.update()
        
    def seleccionar_archivo(self):
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo KMZ o KML",
            filetypes=[("Archivos KMZ/KML", "*.kmz *.kml"), 
                      ("Archivos KMZ", "*.kmz"), 
                      ("Archivos KML", "*.kml"),
                      ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.archivo_seleccionado = archivo
            nombre_archivo = os.path.basename(archivo)
            self.lbl_archivo.config(text=nombre_archivo, foreground="green")
            self.log(f"✅ Archivo seleccionado: {nombre_archivo}")
            
    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(
            title="Seleccionar carpeta de destino"
        )
        if carpeta:
            self.carpeta_destino = carpeta
            self.lbl_carpeta.config(text=carpeta)
            self.log(f"📂 Carpeta destino: {carpeta}")
            
    def iniciar_procesamiento(self):
        if not self.archivo_seleccionado:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo primero")
            return
        
        # Crear carpeta destino si no existe
        if not os.path.exists(self.carpeta_destino):
            os.makedirs(self.carpeta_destino)
            self.log(f"📁 Carpeta creada: {self.carpeta_destino}")
        
        # Iniciar procesamiento en hilo separado
        self.progress.start()
        thread = threading.Thread(target=self.procesar_archivo)
        thread.daemon = True
        thread.start()
        
    def procesar_kml(self, kml_content):
        """Función universal para procesar KMLs"""
        polygon_data = []
        pin_data = []
        
        # MÉTODO 1: Parseo XML
        try:
            if 'xmlns:kml=' not in kml_content and 'xmlns="http://www.opengis.net/kml/2.2"' not in kml_content:
                kml_content = re.sub(r'(<kml[^>]*)', r'\1 xmlns:kml="http://www.opengis.net/kml/2.2"', kml_content)
            
            root = ET.fromstring(kml_content)
            namespaces = {'kml': 'http://www.opengis.net/kml/2.2'}
            
            placemarks = root.findall('.//kml:Placemark', namespaces)
            if not placemarks:
                placemarks = root.findall('.//Placemark')
            
            self.log(f"📊 Método XML: {len(placemarks)} elementos encontrados")
            
            for placemark in placemarks:
                name_elem = placemark.find('kml:name', namespaces)
                if name_elem is None:
                    name_elem = placemark.find('name')
                name = name_elem.text if name_elem is not None else None
                
                # Buscar polígono
                polygon = placemark.find('.//kml:Polygon', namespaces)
                if polygon is None:
                    polygon = placemark.find('.//Polygon')
                
                if polygon is not None:
                    coords_elem = polygon.find('.//kml:coordinates', namespaces)
                    if coords_elem is None:
                        coords_elem = polygon.find('.//coordinates')
                    
                    if coords_elem is not None and coords_elem.text:
                        coords_text = coords_elem.text.strip()
                        coords_list = coords_text.split()
                        coordinates = []
                        
                        for coord in coords_list:
                            parts = coord.strip().split(',')
                            if len(parts) >= 2:
                                lon, lat = float(parts[0]), float(parts[1])
                                coordinates.append((lat, lon))
                        
                        if coordinates:
                            polygon_data.append({
                                'Name': name if name else f'Polígono {len(polygon_data) + 1}',
                                'Coordinates': coordinates
                            })
                            continue
                
                # Buscar punto
                point = placemark.find('.//kml:Point', namespaces)
                if point is None:
                    point = placemark.find('.//Point')
                
                if point is not None:
                    coords_elem = point.find('.//kml:coordinates', namespaces)
                    if coords_elem is None:
                        coords_elem = point.find('.//coordinates')
                    
                    if coords_elem is not None and coords_elem.text:
                        coords_text = coords_elem.text.strip()
                        parts = coords_text.strip().split(',')
                        
                        if len(parts) >= 2:
                            lon, lat = float(parts[0]), float(parts[1])
                            pin_data.append({
                                'Name': name if name else 'Pin sin nombre',
                                'Latitude': lat,
                                'Longitude': lon
                            })
                            
        except Exception as e:
            self.log(f"⚠️ Error en método XML: {str(e)}", "warning")
        
        # Si no se encontró nada, usar regex
        if not polygon_data and not pin_data:
            self.log("🔍 Usando método alternativo (regex)...")
            
            coord_pattern = r'<coordinates[^>]*>(.*?)</coordinates>'
            name_pattern = r'<name[^>]*>(.*?)</name>'
            
            coord_blocks = re.findall(coord_pattern, kml_content, re.DOTALL | re.IGNORECASE)
            names = re.findall(name_pattern, kml_content, re.DOTALL | re.IGNORECASE)
            
            for i, coord_block in enumerate(coord_blocks):
                coord_block = coord_block.strip()
                coords_list = re.split(r'[\s]+', coord_block)
                coordinates = []
                
                for coord in coords_list:
                    coord = coord.strip()
                    if coord and ',' in coord:
                        parts = coord.split(',')
                        if len(parts) >= 2:
                            try:
                                lon, lat = float(parts[0]), float(parts[1])
                                coordinates.append((lat, lon))
                            except ValueError:
                                continue
                
                name = names[i] if i < len(names) else f"Elemento {i+1}"
                
                if len(coordinates) > 1:
                    polygon_data.append({
                        'Name': name,
                        'Coordinates': coordinates
                    })
                elif len(coordinates) == 1:
                    pin_data.append({
                        'Name': name,
                        'Latitude': coordinates[0][0],
                        'Longitude': coordinates[0][1]
                    })
        
        return polygon_data, pin_data
        
    def procesar_archivo(self):
        try:
            nombre_base = os.path.splitext(os.path.basename(self.archivo_seleccionado))[0]
            extension = os.path.splitext(self.archivo_seleccionado)[1].lower()
            
            self.log(f"\n🔄 Procesando: {os.path.basename(self.archivo_seleccionado)}")
            
            temp_dir = os.path.join(os.environ['TEMP'], f"kmz_convert_{nombre_base}")
            os.makedirs(temp_dir, exist_ok=True)
            
            kml_content = None
            
            if extension == '.kmz':
                self.log("📦 Extrayendo KMZ...")
                with zipfile.ZipFile(self.archivo_seleccionado, 'r') as kmz:
                    kmz.extractall(temp_dir)
                
                kml_files = []
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        if file.endswith('.kml'):
                            kml_files.append(os.path.join(root, file))
                
                if not kml_files:
                    raise Exception("No se encontró archivo KML dentro del KMZ")
                
                kml_path = kml_files[0]
                self.log(f"📄 KML encontrado: {os.path.basename(kml_path)}")
                
                with open(kml_path, 'r', encoding='utf-8') as f:
                    kml_content = f.read()
                    
            elif extension == '.kml':
                self.log("📄 Procesando KML directamente...")
                with open(self.archivo_seleccionado, 'r', encoding='utf-8') as f:
                    kml_content = f.read()
            
            # Procesar contenido
            polygon_data, pin_data = self.procesar_kml(kml_content)
            
            # Crear DataFrames
            pin_df = pd.DataFrame(pin_data)
            
            polygon_rows = []
            for polygon in polygon_data:
                row = {'Name': polygon['Name']}
                for i, (lat, lon) in enumerate(polygon['Coordinates']):
                    row[f'Latitud {i+1}'] = round(lat, 5)
                    row[f'Longitud {i+1}'] = round(lon, 5)
                polygon_rows.append(row)
            
            polygon_df = pd.DataFrame(polygon_rows)
            
            # Mostrar resultados
            self.log(f"\n📊 RESULTADOS:")
            self.log(f"   📍 Pines encontrados: {len(pin_data)}")
            self.log(f"   📐 Polígonos encontrados: {len(polygon_data)}")
            
            if len(pin_data) == 0 and len(polygon_data) == 0:
                self.log("⚠️ No se encontraron datos en el archivo", "warning")
                messagebox.showwarning("Sin datos", "No se encontraron coordenadas en el archivo")
            else:
                # Guardar Excel
                output_filename = os.path.join(self.carpeta_destino, f"{nombre_base}.xlsx")
                
                with pd.ExcelWriter(output_filename) as writer:
                    if not pin_df.empty:
                        pin_df.to_excel(writer, sheet_name='Pins', index=False, float_format="%.5f")
                    if not polygon_df.empty:
                        polygon_df.to_excel(writer, sheet_name='Polygons', index=False, float_format="%.5f")
                
                self.log(f"✅ Archivo guardado: {output_filename}", "success")
                messagebox.showinfo("Éxito", f"Archivo procesado correctamente\nGuardado en:\n{output_filename}")
            
            # Limpiar temporales
            shutil.rmtree(temp_dir, ignore_errors=True)
            self.log("🧹 Archivos temporales eliminados")
            
        except Exception as e:
            self.log(f"❌ Error: {str(e)}", "error")
            messagebox.showerror("Error", f"Error al procesar el archivo:\n{str(e)}")
        finally:
            self.progress.stop()

if __name__ == "__main__":
    root = tk.Tk()
    app = ConvertidorKMZ(root)
    root.mainloop()