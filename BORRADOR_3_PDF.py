import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import PyPDF2
import requests
import base64
import os
from pathlib import Path
import tempfile
import re
import threading
import json
from datetime import datetime

class PDFTextExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Texto PDF Mejorado v2.0")
        self.root.geometry("1000x800")
        
        # Variable para almacenar el texto extraído
        self.extracted_text = ""
        
        # API Key para OCR Space (opcional - puedes usar la gratuita)
        self.ocr_api_key = "K83967071688957"  # Tu API key
        
        # Variables de control
        self.is_processing = False
        self.processing_thread = None
        
        # Configuración de mejoras
        self.load_settings()
        
        self.setup_ui()
        
        # Centrar ventana
        self.center_window()
    
    def center_window(self):
        """Centrar la ventana en la pantalla"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def load_settings(self):
        """Cargar configuraciones guardadas"""
        self.settings_file = "pdf_extractor_settings.json"
        self.default_settings = {
            "use_ocr": False,
            "improve_formulas": True,
            "ocr_engine": "2",
            "auto_save": False,
            "output_format": "txt",
            "language": "eng"
        }
        
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    self.settings = json.load(f)
            else:
                self.settings = self.default_settings.copy()
        except:
            self.settings = self.default_settings.copy()
    
    def save_settings(self):
        """Guardar configuraciones"""
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=2)
        except:
            pass
    
    def setup_ui(self):
        # Frame principal con scrollbar
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Frame principal con padding
        main_frame = ttk.Frame(scrollable_frame, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(1, weight=1)
        
        # Título con estilo
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky=(tk.W, tk.E))
        
        title_label = ttk.Label(title_frame, text="🔍 Extractor de Texto PDF Mejorado", 
                               font=("Arial", 18, "bold"))
        title_label.pack(side=tk.LEFT)
        
        version_label = ttk.Label(title_frame, text="v2.0", 
                                 font=("Arial", 10), foreground="gray")
        version_label.pack(side=tk.RIGHT, padx=(10, 0))
        
        # Sección de archivo
        file_frame = ttk.LabelFrame(main_frame, text="📄 Selección de Archivo", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        file_frame.columnconfigure(1, weight=1)
        
        self.select_btn = ttk.Button(file_frame, text="📁 Seleccionar PDF", 
                                    command=self.select_file)
        self.select_btn.grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        
        self.file_label = ttk.Label(file_frame, text="Ningún archivo seleccionado", 
                                   foreground="gray")
        self.file_label.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # Frame para información del archivo
        info_frame = ttk.Frame(file_frame)
        info_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        info_frame.columnconfigure(0, weight=1)
        
        self.file_info_label = ttk.Label(info_frame, text="", font=("Arial", 9), 
                                        foreground="blue")
        self.file_info_label.grid(row=0, column=0, sticky=tk.W)
        
        # Sección de páginas
        page_frame = ttk.LabelFrame(main_frame, text="📑 Selección de Páginas", padding="10")
        page_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        page_frame.columnconfigure(2, weight=1)
        
        self.all_pages = tk.BooleanVar(value=True)
        self.all_pages_checkbox = ttk.Checkbutton(page_frame, text="Todas las páginas", 
                                                 variable=self.all_pages,
                                                 command=self.toggle_page_selection)
        self.all_pages_checkbox.grid(row=0, column=0, sticky=tk.W)
        
        self.pages_label = ttk.Label(page_frame, text="Páginas específicas:")
        self.pages_label.grid(row=0, column=1, padx=(20, 5), sticky=tk.W)
        
        self.pages_entry = ttk.Entry(page_frame, width=25, state="disabled")
        self.pages_entry.grid(row=0, column=2, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.total_pages_label = ttk.Label(page_frame, text="", foreground="gray")
        self.total_pages_label.grid(row=0, column=3, sticky=tk.W)
        
        # Ejemplo de formato
        example_label = ttk.Label(page_frame, text="Formato: 1,3,5-8,12", 
                                 font=("Arial", 8), foreground="gray")
        example_label.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=(2, 0))
        
        # Sección OCR
        ocr_frame = ttk.LabelFrame(main_frame, text="🤖 Configuración OCR", padding="10")
        ocr_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        ocr_frame.columnconfigure(1, weight=1)
        
        self.use_ocr = tk.BooleanVar(value=self.settings.get("use_ocr", False))
        self.ocr_checkbox = ttk.Checkbutton(ocr_frame, text="Usar OCR (para PDFs escaneados)", 
                                           variable=self.use_ocr,
                                           command=self.toggle_ocr_options)
        self.ocr_checkbox.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        # Opciones OCR avanzadas
        self.ocr_options_frame = ttk.Frame(ocr_frame)
        self.ocr_options_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        self.ocr_options_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.ocr_options_frame, text="Motor OCR:").grid(row=0, column=0, sticky=tk.W)
        self.ocr_engine = tk.StringVar(value=self.settings.get("ocr_engine", "2"))
        ocr_combo = ttk.Combobox(self.ocr_options_frame, textvariable=self.ocr_engine, 
                                values=["1 (Básico)", "2 (Avanzado)", "3 (Beta)"], 
                                width=15, state="readonly")
        ocr_combo.grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        
        ttk.Label(self.ocr_options_frame, text="Idioma:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        self.ocr_language = tk.StringVar(value=self.settings.get("language", "eng"))
        lang_combo = ttk.Combobox(self.ocr_options_frame, textvariable=self.ocr_language,
                                 values=["eng (Inglés)", "spa (Español)", "fre (Francés)", "ger (Alemán)"],
                                 width=15, state="readonly")
        lang_combo.grid(row=0, column=3, sticky=tk.W, padx=(5, 0))
        
        # Sección de mejoras
        enhance_frame = ttk.LabelFrame(main_frame, text="✨ Mejoras de Texto", padding="10")
        enhance_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        enhance_frame.columnconfigure(1, weight=1)
        
        self.improve_formulas = tk.BooleanVar(value=self.settings.get("improve_formulas", True))
        self.formulas_checkbox = ttk.Checkbutton(enhance_frame, text="Mejorar reconocimiento de fórmulas matemáticas", 
                                                variable=self.improve_formulas)
        self.formulas_checkbox.grid(row=0, column=0, sticky=tk.W)
        
        self.clean_text = tk.BooleanVar(value=True)
        self.clean_checkbox = ttk.Checkbutton(enhance_frame, text="Limpiar y formatear texto automáticamente", 
                                             variable=self.clean_text)
        self.clean_checkbox.grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        
        self.preserve_structure = tk.BooleanVar(value=True)
        self.structure_checkbox = ttk.Checkbutton(enhance_frame, text="Preservar estructura de párrafos", 
                                                 variable=self.preserve_structure)
        self.structure_checkbox.grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
        # Frame de botones principales
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=(0, 15))
        
        self.extract_btn = ttk.Button(button_frame, text="🚀 Extraer Texto", 
                                     command=self.start_extraction, state="disabled")
        self.extract_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.cancel_btn = ttk.Button(button_frame, text="❌ Cancelar", 
                                    command=self.cancel_extraction, state="disabled")
        self.cancel_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.preview_btn = ttk.Button(button_frame, text="👁️ Vista Previa", 
                                     command=self.preview_text, state="disabled")
        self.preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Área de texto mejorada
        text_frame = ttk.LabelFrame(main_frame, text="📝 Texto Extraído", padding="10")
        text_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        # Frame para controles del texto
        text_controls = ttk.Frame(text_frame)
        text_controls.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        text_controls.columnconfigure(0, weight=1)
        
        # Estadísticas del texto
        self.text_stats_label = ttk.Label(text_controls, text="", font=("Arial", 9), 
                                         foreground="gray")
        self.text_stats_label.grid(row=0, column=0, sticky=tk.W)
        
        # Botones de formato
        format_buttons = ttk.Frame(text_controls)
        format_buttons.grid(row=0, column=1, sticky=tk.E)
        
        ttk.Button(format_buttons, text="🔍 Buscar", 
                  command=self.open_search_dialog, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(format_buttons, text="📋 Copiar", 
                  command=self.copy_text, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(format_buttons, text="🔄 Limpiar", 
                  command=self.clear_text, width=8).pack(side=tk.LEFT, padx=2)
        
        # Área de texto con numeración de líneas
        text_container = ttk.Frame(text_frame)
        text_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        text_container.columnconfigure(1, weight=1)
        text_container.rowconfigure(0, weight=1)
        
        # Numeración de líneas
        self.line_numbers = tk.Text(text_container, width=4, padx=3, takefocus=0,
                                   border=0, state='disabled', wrap='none',
                                   background='#f0f0f0', foreground='gray')
        self.line_numbers.grid(row=0, column=0, sticky=(tk.N, tk.S))
        
        # Área de texto principal
        self.text_area = scrolledtext.ScrolledText(text_container, wrap=tk.WORD, 
                                                  width=80, height=20, font=("Consolas", 10),
                                                  undo=True)
        self.text_area.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Sincronizar scroll
        self.text_area.bind('<KeyPress>', self.update_line_numbers)
        self.text_area.bind('<Button-1>', self.update_line_numbers)
        self.text_area.bind('<MouseWheel>', self.sync_scroll)
        
        # Frame para botones de guardado
        save_frame = ttk.Frame(main_frame)
        save_frame.grid(row=7, column=0, columnspan=3, pady=(0, 15))
        
        self.save_btn = ttk.Button(save_frame, text="💾 Guardar como TXT", 
                                  command=self.save_text, state="disabled")
        self.save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_formatted_btn = ttk.Button(save_frame, text="📄 Guardar Formateado", 
                                           command=self.save_formatted_text, state="disabled")
        self.save_formatted_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.export_btn = ttk.Button(save_frame, text="📤 Exportar", 
                                    command=self.show_export_options, state="disabled")
        self.export_btn.pack(side=tk.LEFT)
        
        # Barra de progreso mejorada
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.progress_label = ttk.Label(progress_frame, text="", font=("Arial", 8))
        self.progress_label.grid(row=1, column=0, pady=(2, 0))
        
        # Label de estado mejorado
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=9, column=0, columnspan=3, sticky=(tk.W, tk.E))
        status_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, text="✅ Listo para usar", 
                                     font=("Arial", 10))
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        self.timestamp_label = ttk.Label(status_frame, text="", 
                                        font=("Arial", 8), foreground="gray")
        self.timestamp_label.grid(row=0, column=1, sticky=tk.E)
        
        # Variables de estado
        self.selected_file = None
        self.total_pages = 0
        
        # Configurar estado inicial
        self.toggle_ocr_options()
        
        # Configurar cierre de ventana
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def sync_scroll(self, event):
        """Sincronizar scroll entre área de texto y numeración"""
        self.line_numbers.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"
    
    def update_line_numbers(self, event=None):
        """Actualizar numeración de líneas"""
        def update():
            self.line_numbers.config(state='normal')
            self.line_numbers.delete('1.0', 'end')
            
            line_count = int(self.text_area.index('end-1c').split('.')[0])
            line_numbers_string = "\n".join(str(i) for i in range(1, line_count))
            self.line_numbers.insert('1.0', line_numbers_string)
            self.line_numbers.config(state='disabled')
            
            # Sincronizar vista
            self.line_numbers.yview_moveto(self.text_area.yview()[0])
        
        self.root.after_idle(update)
    
    def toggle_ocr_options(self):
        """Mostrar/ocultar opciones OCR"""
        if self.use_ocr.get():
            for widget in self.ocr_options_frame.winfo_children():
                widget.configure(state="normal")
        else:
            for widget in self.ocr_options_frame.winfo_children():
                if isinstance(widget, (ttk.Combobox, ttk.Entry)):
                    widget.configure(state="disabled")
    
    def toggle_page_selection(self):
        """Habilitar/deshabilitar la entrada de páginas específicas"""
        if self.all_pages.get():
            self.pages_entry.config(state="disabled")
            self.pages_entry.delete(0, tk.END)
        else:
            self.pages_entry.config(state="normal")
            self.pages_entry.focus()
    
    def get_total_pages(self, file_path):
        """Obtener el número total de páginas del PDF"""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                return len(pdf_reader.pages)
        except Exception as e:
            print(f"Error getting page count: {e}")
            return 0
    
    def get_file_info(self, file_path):
        """Obtener información detallada del archivo"""
        try:
            stat = os.stat(file_path)
            size_mb = stat.st_size / (1024 * 1024)
            modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
            return f"Tamaño: {size_mb:.1f} MB | Modificado: {modified}"
        except:
            return ""
    
    def parse_page_ranges(self, page_input, total_pages):
        """Parsear la entrada de páginas y devolver una lista de números de página"""
        pages = set()
        
        if not page_input.strip():
            return list(range(1, total_pages + 1))
        
        try:
            # Dividir por comas
            parts = page_input.split(',')
            
            for part in parts:
                part = part.strip()
                
                if '-' in part:
                    # Rango de páginas (ej: 3-7)
                    start, end = part.split('-')
                    start = int(start.strip())
                    end = int(end.strip())
                    
                    # Validar rango
                    if start < 1 or end > total_pages or start > end:
                        raise ValueError(f"Rango inválido: {part}")
                    
                    pages.update(range(start, end + 1))
                else:
                    # Página individual
                    page_num = int(part)
                    if page_num < 1 or page_num > total_pages:
                        raise ValueError(f"Página fuera de rango: {page_num}")
                    pages.add(page_num)
            
            return sorted(list(pages))
        
        except ValueError as e:
            raise ValueError(f"Formato de páginas inválido: {str(e)}")
        except Exception:
            raise ValueError("Formato de páginas inválido. Use formato: 1,3,5-8")
    
    def select_file(self):
        """Seleccionar archivo PDF"""
        file_path = filedialog.askopenfilename(
            title="Selecciona un archivo PDF",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
        )
        
        if file_path:
            self.selected_file = file_path
            filename = os.path.basename(file_path)
            self.file_label.config(text=f"📄 {filename}")
            self.extract_btn.config(state="normal")
            
            # Obtener información del archivo
            self.total_pages = self.get_total_pages(file_path)
            file_info = self.get_file_info(file_path)
            
            if self.total_pages > 0:
                self.total_pages_label.config(text=f"({self.total_pages} páginas)")
                self.status_label.config(text=f"✅ Archivo listo: {filename}")
            else:
                self.total_pages_label.config(text="(Error al leer páginas)")
                self.status_label.config(text="⚠️ Error al leer el archivo")
            
            self.file_info_label.config(text=file_info)
            self.timestamp_label.config(text=datetime.now().strftime("%H:%M:%S"))
    
    def improve_mathematical_formulas(self, text):
        """Mejorar el reconocimiento de fórmulas matemáticas"""
        if not self.improve_formulas.get():
            return text
        
        # Patrones mejorados para fórmulas matemáticas
        improvements = [
            # Fracciones comunes
            (r'(\w+)\s*×\s*\(\s*(\w+)\s*-\s*(\w+)\s*\)\s*r\s*=\s*(\w+)\s*\+\s*(\w+)', 
             r'r = \1 × (\2 - \3) / (\4 + \5)'),
            
            # Patrón específico mejorado
            (r'N\s*×\s*\(\s*B\s*-\s*C\s*\)\s*r\s*=\s*N\s*\+\s*V', 
             r'r = N × (B - C) / (N + V)'),
            
            # Operadores matemáticos
            (r'(\w)\s*=\s*(\w)', r'\1 = \2'),
            (r'(\w)\s*\+\s*(\w)', r'\1 + \2'),
            (r'(\w)\s*-\s*(\w)', r'\1 - \2'),
            (r'(\w)\s*×\s*(\w)', r'\1 × \2'),
            (r'(\w)\s*/\s*(\w)', r'\1 / \2'),
            (r'(\w)\s*\*\s*(\w)', r'\1 × \2'),
            
            # Paréntesis
            (r'\(\s*([^)]+)\s*\)', r'(\1)'),
            
            # Exponentes (reconocer patrones como x^2)
            (r'(\w)\s*\^\s*(\w)', r'\1^\2'),
            
            # Subíndices (reconocer patrones como H_2O)
            (r'(\w)\s*_\s*(\w)', r'\1_\2'),
            
            # Variables con definiciones
            (r'([A-Z]):\s*([^A-Z\n]+?)(?=[A-Z]:|\n\n|\Z)', r'\1: \2\n'),
            
            # Limpiar espacios múltiples
            (r'\s+', ' '),
        ]
        
        improved_text = text
        for pattern, replacement in improvements:
            improved_text = re.sub(pattern, replacement, improved_text, flags=re.IGNORECASE | re.MULTILINE)
        
        return improved_text
    
    def clean_and_format_text(self, text):
        """Limpiar y formatear el texto extraído"""
        if not self.clean_text.get():
            return text
        
        # Limpieza general
        cleaned = text
        
        # Remover caracteres extraños
        cleaned = re.sub(r'[^\w\s\.,;:!?()[\]{}\-=+*/\\<>@#$%^&|\'\"áéíóúñüÁÉÍÓÚÑÜ]', '', cleaned)
        
        # Normalizar espacios
        cleaned = re.sub(r'\s+', ' ', cleaned)
        
        # Normalizar saltos de línea
        cleaned = re.sub(r'\n\s*\n\s*\n+', '\n\n', cleaned)
        
        # Preservar estructura si está activado
        if self.preserve_structure.get():
            # Mantener párrafos
            cleaned = re.sub(r'([.!?])\s*\n\s*([A-Z])', r'\1\n\n\2', cleaned)
        
        return cleaned.strip()
    
    def post_process_ocr_text(self, text):
        """Post-procesamiento específico para texto OCR"""
        # Correcciones comunes de OCR
        corrections = {
            'CINs': 'CINS',
            'ICATI CADE': 'ICADE',
            'rn': 'm',
            'vv': 'w',
            '0': 'O',  # En contextos apropiados
            '1': 'I',  # En contextos apropiados
            '•': '■',
            '→': '->',
            '←': '<-',
        }
        
        processed_text = text
        
        # Aplicar correcciones contextualmente
        for wrong, correct in corrections.items():
            processed_text = processed_text.replace(wrong, correct)
        
        # Mejorar fórmulas matemáticas
        processed_text = self.improve_mathematical_formulas(processed_text)
        
        # Limpiar y formatear
        processed_text = self.clean_and_format_text(processed_text)
        
        return processed_text
    
    def extract_text_from_pdf(self, file_path, selected_pages=None, progress_callback=None):
        """Extraer texto usando PyPDF2"""
        text = ""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                
                if selected_pages is None:
                    selected_pages = list(range(1, total_pages + 1))
                
                for i, page_num in enumerate(selected_pages):
                    if self.is_processing:  # Verificar si se canceló
                        return ""
                    
                    if 1 <= page_num <= total_pages:
                        page = pdf_reader.pages[page_num - 1]
                        page_text = page.extract_text()
                        if page_text.strip():
                            text += f"\n--- Página {page_num} ---\n"
                            text += page_text + "\n"
                    
                    # Actualizar progreso
                    if progress_callback:
                        progress = (i + 1) / len(selected_pages) * 100
                        progress_callback(f"Procesando página {page_num}...", progress)
        
        except Exception as e:
            raise Exception(f"Error al leer PDF: {str(e)}")
        
        return text
    
    def extract_text_with_ocr(self, file_path, selected_pages=None, progress_callback=None):
        """Extraer texto usando OCR Space API mejorado"""
        try:
            url = 'https://api.ocr.space/parse/image'
            temp_file = None
            
            # Preparar archivo para OCR
            pdf_to_send = file_path
            if selected_pages is not None:
                from PyPDF2 import PdfWriter, PdfReader
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                writer = PdfWriter()
                with open(file_path, 'rb') as infile:
                    reader = PdfReader(infile)
                    for page_num in selected_pages:
                        if 1 <= page_num <= len(reader.pages):
                            writer.add_page(reader.pages[page_num - 1])
                    writer.write(temp_file)
                temp_file.close()
                pdf_to_send = temp_file.name
            
            if progress_callback:
                progress_callback("Enviando archivo para OCR...", 10)
            
            with open(pdf_to_send, 'rb') as file:
                files = {'file': file}
                
                # Obtener configuración OCR
                engine = self.ocr_engine.get().split()[0]
                language = self.ocr_language.get().split()[0]
                
                data = {
                    'apikey': self.ocr_api_key,
                    'language': language,
                    'detectOrientation': 'true',
                    'scale': 'true',
                    'OCREngine': engine,
                    'filetype': 'PDF',
                    'isTable': 'true',
                }
                
                if progress_callback:
                    progress_callback("Procesando con OCR...", 50)
                
                response = requests.post(url, files=files, data=data, timeout=60)
                result = response.json()

                if result.get('IsErroredOnProcessing'):
                    raise Exception(f"Error en OCR: {result.get('ErrorMessage', 'Error desconocido')}")

                text = ""
                if 'ParsedResults' in result:
                    total_results = len(result['ParsedResults'])
                    for i, page_result in enumerate(result['ParsedResults']):
                        if not self.is_processing:  # Verificar cancelación
                            return ""
                        
                        if 'ParsedText' in page_result:
                            page_text = page_result['ParsedText']
                            page_text = self.post_process_ocr_text(page_text)
                            page_num = selected_pages[i] if selected_pages else i + 1
                            text += f"\n--- Página {page_num} (OCR) ---\n"
                            text += page_text + "\n"
                        
                        if progress_callback:
                            progress = 50 + (i + 1) / total_results * 40
                            progress_callback(f"Procesando resultado OCR {i+1}/{total_results}...", progress)

                return text

        except Exception as e:
            raise Exception(f"Error en OCR: {str(e)}")
        finally:
            if temp_file is not None:
                try:
                    os.unlink(temp_file.name)
                except Exception:
                    pass
    
    def extraction_worker(self, file_path, selected_pages, use_ocr):
        """Worker thread para extracción de texto"""
        try:
            def update_progress(message, percent=None):
                if self.is_processing:
                    self.root.after(0, lambda: self.progress_label.config(text=message))
                    if percent is not None:
                        self.root.after(0, lambda: self.progress.config(value=percent, mode='determinate'))
            
            if use_ocr:
                self.extracted_text = self.extract_text_with_ocr(file_path, selected_pages, update_progress)
                if not self.extracted_text.strip():
                    update_progress("OCR falló, intentando extracción normal...", 0)
                    self.extracted_text = self.extract_text_from_pdf(file_path, selected_pages, update_progress)
                    method = "extracción directa (fallback)"
                else:
                    method = "OCR mejorado"
            else:
                self.extracted_text = self.extract_text_from_pdf(file_path, selected_pages, update_progress)
                if self.improve_formulas.get() or self.clean_text.get():
                    update_progress("Aplicando mejoras al texto...", 90)
                    self.extracted_text = self.post_process_ocr_text(self.extracted_text)
                method = "extracción directa"
            
            # Finalizar en el hilo principal
            self.root.after(0, lambda: self.extraction_completed(method, selected_pages))
            
        except Exception as e:
            self.root.after(0, lambda: self.extraction_failed(str(e)))
    
    def start_extraction(self):
        """Iniciar extracción de texto en thread separado"""
        if not self.selected_file:
            messagebox.showwarning("Advertencia", "Por favor selecciona un archivo PDF primero.")
            return
        
        # Validar páginas
        selected_pages = None
        if not self.all_pages.get():
            try:
                page_input = self.pages_entry.get()
                selected_pages = self.parse_page_ranges(page_input, self.total_pages)
                if not selected_pages:
                    messagebox.showwarning("Advertencia", "No se especificaron páginas válidas.")
                    return
            except ValueError as e:
                messagebox.showerror("Error", str(e))
                return
        
        # Configurar UI para procesamiento
        self.is_processing = True
        self.progress.config(mode='indeterminate')
        self.progress.start(10)
        self.extract_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.status_label.config(text="🔄 Extrayendo texto...")
        
        # Guardar configuraciones
        self.settings.update({
            "use_ocr": self.use_ocr.get(),
            "improve_formulas": self.improve_formulas.get(),
            "ocr_engine": self.ocr_engine.get(),
            "language": self.ocr_language.get()
        })
        self.save_settings()
        
        # Iniciar thread de extracción
        self.processing_thread = threading.Thread(
            target=self.extraction_worker,
            args=(self.selected_file, selected_pages, self.use_ocr.get())
        )
        self.processing_thread.daemon = True
        self.processing_thread.start()
    
    def cancel_extraction(self):
        """Cancelar extracción en progreso"""
        self.is_processing = False
        self.progress.stop()
        self.progress.config(mode='indeterminate')
        self.progress_label.config(text="")
        self.extract_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        self.status_label.config(text="❌ Extracción cancelada")
        self.timestamp_label.config(text=datetime.now().strftime("%H:%M:%S"))
    
    def extraction_completed(self, method, selected_pages):
        """Finalizar extracción exitosa"""
        self.is_processing = False
        self.progress.stop()
        self.progress.config(mode='indeterminate')
        self.progress_label.config(text="")
        self.extract_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        
        # Mostrar texto
        self.text_area.delete(1.0, tk.END)
        if self.extracted_text.strip():
            self.text_area.insert(1.0, self.extracted_text)
            self.save_btn.config(state="normal")
            self.save_formatted_btn.config(state="normal")
            self.export_btn.config(state="normal")
            self.preview_btn.config(state="normal")
            
            # Actualizar estadísticas
            self.update_text_statistics()
            
            pages_count = len(selected_pages) if selected_pages else self.total_pages
            self.status_label.config(text=f"✅ Texto extraído con {method} de {pages_count} página(s)")
        else:
            self.show_extraction_help()
            self.status_label.config(text="⚠️ No se encontró texto en las páginas seleccionadas")
        
        self.timestamp_label.config(text=datetime.now().strftime("%H:%M:%S"))
        self.update_line_numbers()
    
    def extraction_failed(self, error_message):
        """Manejar error en extracción"""
        self.is_processing = False
        self.progress.stop()
        self.progress.config(mode='indeterminate')
        self.progress_label.config(text="")
        self.extract_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        
        messagebox.showerror("Error", f"Error al extraer texto:\n{error_message}")
        self.status_label.config(text="❌ Error en la extracción")
        self.timestamp_label.config(text=datetime.now().strftime("%H:%M:%S"))
    
    def show_extraction_help(self):
        """Mostrar ayuda cuando no se encuentra texto"""
        help_text = """No se pudo extraer texto del PDF.

Posibles soluciones:

🔍 Para PDFs escaneados:
   • Activa la opción 'Usar OCR'
   • Prueba diferentes motores OCR (1, 2, o 3)
   • Verifica que el idioma sea correcto

📄 Para PDFs con texto:
   • Verifica que las páginas especificadas sean correctas
   • Algunos PDFs tienen texto en formato imagen
   • Intenta con diferentes páginas

⚙️ Configuración avanzada:
   • Cambia el motor OCR si usas OCR
   • Verifica tu conexión a internet para OCR
   • Prueba desactivar las mejoras de texto

Si el problema persiste, el PDF podría estar protegido o dañado."""
        
        self.text_area.insert(1.0, help_text)
    
    def update_text_statistics(self):
        """Actualizar estadísticas del texto"""
        if not self.extracted_text:
            self.text_stats_label.config(text="")
            return
        
        lines = len(self.extracted_text.split('\n'))
        words = len(self.extracted_text.split())
        chars = len(self.extracted_text)
        chars_no_spaces = len(self.extracted_text.replace(' ', ''))
        
        stats = f"Líneas: {lines:,} | Palabras: {words:,} | Caracteres: {chars:,} ({chars_no_spaces:,} sin espacios)"
        self.text_stats_label.config(text=stats)
    
    def preview_text(self):
        """Mostrar vista previa del texto en ventana separada"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para previsualizar.")
            return
        
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Vista Previa del Texto")
        preview_window.geometry("800x600")
        
        # Área de texto de solo lectura
        preview_text = scrolledtext.ScrolledText(preview_window, wrap=tk.WORD, 
                                               font=("Arial", 11), state='disabled')
        preview_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        preview_text.config(state='normal')
        preview_text.insert(1.0, self.extracted_text)
        preview_text.config(state='disabled')
        
        # Botones
        button_frame = ttk.Frame(preview_window)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(button_frame, text="Cerrar", 
                  command=preview_window.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Copiar Todo", 
                  command=lambda: self.copy_to_clipboard(self.extracted_text)).pack(side=tk.RIGHT, padx=(0, 10))
    
    def open_search_dialog(self):
        """Abrir diálogo de búsqueda"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto donde buscar.")
            return
        
        search_window = tk.Toplevel(self.root)
        search_window.title("Buscar en el texto")
        search_window.geometry("400x150")
        search_window.resizable(False, False)
        
        # Centrar ventana
        search_window.transient(self.root)
        search_window.grab_set()
        
        frame = ttk.Frame(search_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Buscar:").pack(anchor=tk.W)
        search_entry = ttk.Entry(frame, width=40)
        search_entry.pack(fill=tk.X, pady=(5, 10))
        search_entry.focus()
        
        # Variables de búsqueda
        case_sensitive = tk.BooleanVar()
        whole_word = tk.BooleanVar()
        
        ttk.Checkbutton(frame, text="Distinguir mayúsculas/minúsculas", 
                       variable=case_sensitive).pack(anchor=tk.W)
        ttk.Checkbutton(frame, text="Palabra completa", 
                       variable=whole_word).pack(anchor=tk.W)
        
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        def search_text():
            query = search_entry.get()
            if query:
                self.search_in_text(query, case_sensitive.get(), whole_word.get())
                search_window.destroy()
        
        ttk.Button(button_frame, text="Buscar", 
                  command=search_text).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Cancelar", 
                  command=search_window.destroy).pack(side=tk.RIGHT, padx=(0, 10))
        
        search_entry.bind('<Return>', lambda e: search_text())
    
    def search_in_text(self, query, case_sensitive, whole_word):
        """Buscar texto en el área de texto"""
        # Limpiar selecciones anteriores
        self.text_area.tag_remove('search', '1.0', tk.END)
        
        if not query:
            return
        
        text_content = self.text_area.get('1.0', tk.END)
        if not case_sensitive:
            text_content = text_content.lower()
            query = query.lower()
        
        start_pos = '1.0'
        count = 0
        
        while True:
            pos = self.text_area.search(query, start_pos, tk.END, regexp=whole_word)
            if not pos:
                break
            
            end_pos = f"{pos}+{len(query)}c"
            self.text_area.tag_add('search', pos, end_pos)
            start_pos = end_pos
            count += 1
        
        # Configurar estilo de búsqueda
        self.text_area.tag_config('search', background='yellow', foreground='black')
        
        if count > 0:
            # Ir a la primera coincidencia
            first_match = self.text_area.search(query, '1.0', tk.END, regexp=whole_word)
            self.text_area.see(first_match)
            messagebox.showinfo("Búsqueda", f"Se encontraron {count} coincidencia(s)")
        else:
            messagebox.showinfo("Búsqueda", "No se encontraron coincidencias")
    
    def copy_text(self):
        """Copiar texto seleccionado o todo el texto"""
        try:
            selected_text = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            selected_text = self.text_area.get('1.0', tk.END).strip()
        
        if selected_text:
            self.copy_to_clipboard(selected_text)
            messagebox.showinfo("Copiado", "Texto copiado al portapapeles")
        else:
            messagebox.showwarning("Advertencia", "No hay texto para copiar")
    
    def copy_to_clipboard(self, text):
        """Copiar texto al portapapeles"""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
    
    def save_text(self):
        """Guardar el texto extraído en un archivo TXT"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para guardar.")
            return
        
        self._save_file("txt", "Archivo de texto", self.extracted_text)
    
    def save_formatted_text(self):
        """Guardar texto con formato mejorado"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para guardar.")
            return
        
        # Aplicar formato adicional
        formatted_text = self._format_for_save(self.extracted_text)
        self._save_file("txt", "Archivo de texto formateado", formatted_text, "_formateado")
    
    def _format_for_save(self, text):
        """Aplicar formato adicional para guardado"""
        formatted = text
        
        # Agregar encabezado
        header = f"""Texto extraído de PDF
Fecha: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
Archivo: {os.path.basename(self.selected_file) if self.selected_file else "Desconocido"}
Método: {"OCR" if self.use_ocr.get() else "Extracción directa"}

{"="*60}

"""
        
        formatted = header + formatted
        
        # Agregar pie de página
        footer = f"""

{"="*60}
Procesado con PDF Text Extractor v2.0
Total de caracteres: {len(text):,}
"""
        
        formatted += footer
        
        return formatted
    
    def show_export_options(self):
        """Mostrar opciones de exportación"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para exportar.")
            return
        
        export_window = tk.Toplevel(self.root)
        export_window.title("Opciones de Exportación")
        export_window.geometry("400x300")
        export_window.resizable(False, False)
        
        export_window.transient(self.root)
        export_window.grab_set()
        
        frame = ttk.Frame(export_window, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Selecciona el formato de exportación:", 
                 font=("Arial", 12, "bold")).pack(pady=(0, 15))
        
        # Opciones de formato
        export_format = tk.StringVar(value="txt")
        
        formats = [
            ("txt", "Archivo de texto (.txt)", "Texto plano sin formato"),
            ("rtf", "Texto enriquecido (.rtf)", "Con formato básico"),
            ("html", "Página web (.html)", "Para visualización web"),
            ("json", "Datos JSON (.json)", "Para procesamiento automático")
        ]
        
        for value, text, desc in formats:
            frame_option = ttk.Frame(frame)
            frame_option.pack(fill=tk.X, pady=2)
            
            ttk.Radiobutton(frame_option, text=text, value=value, 
                           variable=export_format).pack(anchor=tk.W)
            ttk.Label(frame_option, text=desc, font=("Arial", 9), 
                     foreground="gray").pack(anchor=tk.W, padx=(20, 0))
        
        # Opciones adicionales
        ttk.Separator(frame, orient='horizontal').pack(fill=tk.X, pady=15)
        
        include_metadata = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="Incluir metadatos (fecha, archivo, etc.)", 
                       variable=include_metadata).pack(anchor=tk.W)
        
        # Botones
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        def export_file():
            format_type = export_format.get()
            include_meta = include_metadata.get()
            self._export_text(format_type, include_meta)
            export_window.destroy()
        
        ttk.Button(button_frame, text="Exportar", 
                  command=export_file).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Cancelar", 
                  command=export_window.destroy).pack(side=tk.RIGHT, padx=(0, 10))
    
    def _export_text(self, format_type, include_metadata):
        """Exportar texto en el formato especificado"""
        if format_type == "txt":
            content = self._format_for_save(self.extracted_text) if include_metadata else self.extracted_text
            self._save_file("txt", "Archivo de texto", content)
        
        elif format_type == "rtf":
            content = self._convert_to_rtf(self.extracted_text, include_metadata)
            self._save_file("rtf", "Archivo RTF", content)
        
        elif format_type == "html":
            content = self._convert_to_html(self.extracted_text, include_metadata)
            self._save_file("html", "Archivo HTML", content)
        
        elif format_type == "json":
            content = self._convert_to_json(self.extracted_text, include_metadata)
            self._save_file("json", "Archivo JSON", content)
    
    def _convert_to_rtf(self, text, include_metadata):
        """Convertir texto a formato RTF"""
        rtf_header = r"{\rtf1\ansi\deff0 {\fonttbl {\f0 Times New Roman;}}}"
        rtf_content = text.replace('\n', r'\par ')
        
        if include_metadata:
            metadata = f"Extraído de: {os.path.basename(self.selected_file) if self.selected_file else 'Desconocido'}\\par Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\\par\\par "
            rtf_content = metadata + rtf_content
        
        return f"{rtf_header}{rtf_content}}}"
    
    def _convert_to_html(self, text, include_metadata):
        """Convertir texto a formato HTML"""
        html_content = text.replace('\n', '<br>\n')
        
        metadata_html = ""
        if include_metadata:
            metadata_html = f"""
        <div style="background-color: #f0f0f0; padding: 10px; margin-bottom: 20px; border-left: 4px solid #007acc;">
            <h3>Información del archivo</h3>
            <p><strong>Archivo:</strong> {os.path.basename(self.selected_file) if self.selected_file else 'Desconocido'}</p>
            <p><strong>Fecha de extracción:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p><strong>Método:</strong> {"OCR" if self.use_ocr.get() else "Extracción directa"}</p>
        </div>
        """
        
        return f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Texto Extraído de PDF</title>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; margin: 40px; }}
        .content {{ max-width: 800px; margin: 0 auto; }}
        pre {{ white-space: pre-wrap; font-family: 'Consolas', monospace; }}
    </style>
</head>
<body>
    <div class="content">
        <h1>Texto Extraído de PDF</h1>
        {metadata_html}
        <div style="border: 1px solid #ddd; padding: 20px; background-color: #fafafa;">
            <pre>{html_content}</pre>
        </div>
    </div>
</body>
</html>"""
    
    def _convert_to_json(self, text, include_metadata):
        """Convertir texto a formato JSON"""
        data = {
            "extracted_text": text,
            "extraction_info": {
                "timestamp": datetime.now().isoformat(),
                "method": "OCR" if self.use_ocr.get() else "Direct extraction",
                "total_characters": len(text),
                "total_words": len(text.split()),
                "total_lines": len(text.split('\n'))
            }
        }
        
        if include_metadata and self.selected_file:
            data["file_info"] = {
                "filename": os.path.basename(self.selected_file),
                "total_pages": self.total_pages,
                "file_size": os.path.getsize(self.selected_file)
            }
        
        return json.dumps(data, indent=2, ensure_ascii=False)
    
    def _save_file(self, extension, file_type, content, suffix=""):
        """Guardar archivo con el contenido especificado"""
        if self.selected_file:
            pdf_name = Path(self.selected_file).stem
            default_name = f"{pdf_name}_texto{suffix}.{extension}"
        else:
            default_name = f"texto_extraido{suffix}.{extension}"
        
        file_path = filedialog.asksaveasfilename(
            title=f"Guardar como {file_type}",
            defaultextension=f".{extension}",
            initialvalue=default_name,
            filetypes=[(file_type, f"*.{extension}"), ("Todos los archivos", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(content)
                messagebox.showinfo("Éxito", f"Archivo guardado:\n{file_path}")
                self.status_label.config(text=f"💾 Guardado: {os.path.basename(file_path)}")
                self.timestamp_label.config(text=datetime.now().strftime("%H:%M:%S"))
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar archivo:\n{str(e)}")
    
    def clear_text(self):
        """Limpiar el área de texto y resetear"""
        self.text_area.delete(1.0, tk.END)
        self.extracted_text = ""
        self.selected_file = None
        self.total_pages = 0
        self.file_label.config(text="Ningún archivo seleccionado", foreground="gray")
        self.file_info_label.config(text="")
        self.total_pages_label.config(text="")
        self.text_stats_label.config(text="")
        self.pages_entry.delete(0, tk.END)
        self.all_pages.set(True)
        self.pages_entry.config(state="disabled")
        self.extract_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.save_formatted_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.preview_btn.config(state="disabled")
        self.status_label.config(text="✅ Listo para usar")
        self.timestamp_label.config(text="")
        self.update_line_numbers()
    
    def on_closing(self):
        """Manejar cierre de la aplicación"""
        if self.is_processing:
            if messagebox.askokcancel("Cerrar", "Hay una extracción en progreso.\n¿Deseas cancelar y cerrar la aplicación?"):
                self.is_processing = False
                self.root.destroy()
        else:
            self.save_settings()
            self.root.destroy()

def main():
    """Función principal"""
    # Verificar dependencias
    try:
        import PyPDF2
        import requests
    except ImportError as e:
        print("Error: Faltan dependencias requeridas.")
        print("Instala las dependencias con:")
        print("pip install PyPDF2 requests")
        input("Presiona Enter para salir...")
        return
    
    try:
        root = tk.Tk()
        app = PDFTextExtractor(root)
        
        # Configurar tema si está disponible
        try:
            root.tk.call("source", "azure.tcl")
            root.tk.call("set_theme", "light")
        except:
            pass  # Continuar sin tema personalizado
        
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror("Error Fatal", f"Error al inicializar la aplicación:\n{str(e)}")

if __name__ == "__main__":
    main()