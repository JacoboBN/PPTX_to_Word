import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import PyPDF2
import requests
import base64
import os
from pathlib import Path
import tempfile
import re
import json
from PIL import Image, ImageEnhance, ImageFilter
import fitz  # PyMuPDF
import io

class PDFFormulaExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor PDF con Reconocimiento de Fórmulas")
        self.root.geometry("1000x800")
        
        # Variable para almacenar el texto extraído
        self.extracted_text = ""
        
        # API Keys para diferentes servicios OCR
        self.ocr_space_key = "K83967071688957"  # Tu API key actual
        
        # Patrones de fórmulas matemáticas comunes
        self.math_patterns = self.initialize_math_patterns()
        
        self.setup_ui()
    
    def initialize_math_patterns(self):
        """Inicializar patrones de reconocimiento de fórmulas matemáticas"""
        return {
            # Patrones de símbolos matemáticos mal reconocidos por OCR
            'symbol_corrections': {
                # Letras griegas comunes
                'α': ['a', 'α', 'alpha'],
                'β': ['b', 'β', 'beta', 'B'],
                'γ': ['y', 'γ', 'gamma'],
                'δ': ['d', 'δ', 'delta'],
                'ε': ['e', 'ε', 'epsilon'],
                'π': ['pi', 'π', 'TI', 'n'],
                'σ': ['o', 'σ', 'sigma'],
                'μ': ['u', 'μ', 'mu'],
                'λ': ['λ', 'lambda'],
                'Σ': ['E', 'Σ', 'sum'],
                
                # Operadores matemáticos
                '×': ['x', '*', 'X'],
                '÷': ['/', ':'],
                '≈': ['~', '≈', 'approx'],
                '≠': ['!=', '≠'],
                '≤': ['<=', '≤'],
                '≥': ['>=', '≥'],
                '∞': ['inf', '∞', 'infinity'],
                '√': ['sqrt', '√', 'V'],
                '∫': ['integral', '∫', 'f'],
                '∂': ['partial', '∂', 'd'],
                '∇': ['nabla', '∇', 'V'],
                
                # Superíndices/subíndices comunes
                '²': ['^2', '2'],
                '³': ['^3', '3'],
                '¹': ['^1', '1'],
                '⁻¹': ['^-1', '^(-1)'],
            },
            
            # Patrones de estructura de fórmulas
            'formula_structures': [
                # Fracciones
                r'(\w+)\s*/\s*(\w+)',  # a/b
                r'(\w+)\s*÷\s*(\w+)',  # a÷b
                
                # Exponentes
                r'(\w+)\^(\d+)',  # a^2
                r'(\w+)\^{([^}]+)}',  # a^{n+1}
                
                # Raíces
                r'√\(([^)]+)\)',  # √(x)
                r'√(\w+)',  # √x
                
                # Sumatorias
                r'Σ\s*(\w+)',  # Σx
                r'∑\s*(\w+)',  # ∑x
                
                # Integrales
                r'∫\s*([^d]+)\s*d(\w+)',  # ∫f(x)dx
                
                # Logaritmos
                r'log\(([^)]+)\)',  # log(x)
                r'ln\(([^)]+)\)',   # ln(x)
                
                # Trigonométricas
                r'(sin|cos|tan|sec|csc|cot)\(([^)]+)\)',
                
                # Matrices/vectores
                r'\[([^\]]+)\]',  # [a, b, c]
                r'\(([^)]+)\)',   # (a, b, c)
            ]
        }
    
    def setup_ui(self):
        # Frame principal con scroll
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)
        
        main_canvas.configure(yscrollcommand=scrollbar.set)
        main_canvas.bind('<Configure>', lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Frame principal
        main_frame = ttk.Frame(scrollable_frame, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="Extractor PDF con Reconocimiento Avanzado de Fórmulas", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Frame de selección de archivo
        file_frame = ttk.LabelFrame(main_frame, text="Selección de Archivo", padding="10")
        file_frame.pack(fill="x", pady=(0, 10))
        
        file_controls = ttk.Frame(file_frame)
        file_controls.pack(fill="x")
        
        self.select_btn = ttk.Button(file_controls, text="Seleccionar PDF", command=self.select_file)
        self.select_btn.pack(side="left", padx=(0, 10))
        
        self.file_label = ttk.Label(file_controls, text="Ningún archivo seleccionado")
        self.file_label.pack(side="left", fill="x", expand=True)
        
        self.total_pages_label = ttk.Label(file_controls, text="", foreground="gray")
        self.total_pages_label.pack(side="right")
        
        # Frame de configuración de extracción
        config_frame = ttk.LabelFrame(main_frame, text="Configuración de Extracción", padding="10")
        config_frame.pack(fill="x", pady=(0, 10))
        
        # Método de extracción
        method_frame = ttk.Frame(config_frame)
        method_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(method_frame, text="Método de extracción:").pack(side="left")
        
        self.extraction_method = tk.StringVar(value="hybrid")
        methods = [
            ("Híbrido (Recomendado)", "hybrid"),
            ("Solo Texto Directo", "direct"),
            ("Solo OCR", "ocr"),
            ("OCR + Preprocesamiento de Imágenes", "ocr_enhanced")
        ]
        
        for text, value in methods:
            ttk.Radiobutton(method_frame, text=text, variable=self.extraction_method, 
                           value=value).pack(side="left", padx=(10, 0))
        
        # Configuración de fórmulas
        formula_frame = ttk.LabelFrame(config_frame, text="Reconocimiento de Fórmulas", padding="5")
        formula_frame.pack(fill="x", pady=(10, 0))
        
        # Opciones de mejora de fórmulas
        self.enhance_formulas = tk.BooleanVar(value=True)
        ttk.Checkbutton(formula_frame, text="Mejorar reconocimiento de fórmulas matemáticas", 
                       variable=self.enhance_formulas).pack(anchor="w")
        
        self.fix_symbols = tk.BooleanVar(value=True)
        ttk.Checkbutton(formula_frame, text="Corregir símbolos matemáticos mal reconocidos", 
                       variable=self.fix_symbols).pack(anchor="w")
        
        self.detect_structures = tk.BooleanVar(value=True)
        ttk.Checkbutton(formula_frame, text="Detectar estructuras matemáticas (fracciones, exponentes, etc.)", 
                       variable=self.detect_structures).pack(anchor="w")
        
        self.preserve_layout = tk.BooleanVar(value=True)
        ttk.Checkbutton(formula_frame, text="Preservar layout de ecuaciones", 
                       variable=self.preserve_layout).pack(anchor="w")
        
        # Configuración OCR
        ocr_frame = ttk.LabelFrame(config_frame, text="Configuración OCR", padding="5")
        ocr_frame.pack(fill="x", pady=(10, 0))
        
        ocr_controls = ttk.Frame(ocr_frame)
        ocr_controls.pack(fill="x")
        
        ttk.Label(ocr_controls, text="Motor OCR:").pack(side="left")
        self.ocr_engine = tk.StringVar(value="2")
        ocr_combo = ttk.Combobox(ocr_controls, textvariable=self.ocr_engine, 
                                values=["1 (Básico)", "2 (Avanzado)", "3 (Beta - Mejor para fórmulas)"], 
                                width=25, state="readonly")
        ocr_combo.pack(side="left", padx=(5, 20))
        
        ttk.Label(ocr_controls, text="Idioma:").pack(side="left")
        self.ocr_language = tk.StringVar(value="eng")
        lang_combo = ttk.Combobox(ocr_controls, textvariable=self.ocr_language, 
                                 values=["eng", "spa", "fre", "ger"], 
                                 width=10, state="readonly")
        lang_combo.pack(side="left", padx=(5, 0))
        
        # Selección de páginas
        page_frame = ttk.LabelFrame(config_frame, text="Selección de Páginas", padding="5")
        page_frame.pack(fill="x", pady=(10, 0))
        
        page_controls = ttk.Frame(page_frame)
        page_controls.pack(fill="x")
        
        self.all_pages = tk.BooleanVar(value=True)
        self.all_pages_checkbox = ttk.Checkbutton(page_controls, text="Todas las páginas", 
                                                 variable=self.all_pages,
                                                 command=self.toggle_page_selection)
        self.all_pages_checkbox.pack(side="left")
        
        ttk.Label(page_controls, text="Páginas específicas (ej: 1,3,5-8):").pack(side="left", padx=(20, 5))
        
        self.pages_entry = ttk.Entry(page_controls, width=20, state="disabled")
        self.pages_entry.pack(side="left")
        
        # Área de texto para mostrar el contenido
        text_frame = ttk.LabelFrame(main_frame, text="Texto Extraído", padding="10")
        text_frame.pack(fill="both", expand=True, pady=(10, 0))
        
        # Crear notebook para diferentes vistas
        self.notebook = ttk.Notebook(text_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # Pestaña de texto completo
        self.text_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.text_frame, text="Texto Completo")
        
        self.text_area = scrolledtext.ScrolledText(self.text_frame, wrap=tk.WORD, 
                                                  width=80, height=20, font=("Consolas", 10))
        self.text_area.pack(fill="both", expand=True)
        
        # Pestaña de fórmulas detectadas
        self.formula_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.formula_frame, text="Fórmulas Detectadas")
        
        self.formula_area = scrolledtext.ScrolledText(self.formula_frame, wrap=tk.WORD, 
                                                     width=80, height=20, font=("Consolas", 10))
        self.formula_area.pack(fill="both", expand=True)
        
        # Frame para botones y controles
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill="x", pady=(20, 0))
        
        # Botones de acción
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(side="left")
        
        self.extract_btn = ttk.Button(button_frame, text="Extraer Texto y Fórmulas", 
                                     command=self.extract_text, state="disabled")
        self.extract_btn.pack(side="left", padx=(0, 10))
        
        self.save_btn = ttk.Button(button_frame, text="Guardar Texto", 
                                  command=self.save_text, state="disabled")
        self.save_btn.pack(side="left", padx=(0, 10))
        
        self.save_formulas_btn = ttk.Button(button_frame, text="Guardar Fórmulas", 
                                           command=self.save_formulas, state="disabled")
        self.save_formulas_btn.pack(side="left", padx=(0, 10))
        
        self.clear_btn = ttk.Button(button_frame, text="Limpiar", command=self.clear_text)
        self.clear_btn.pack(side="left")
        
        # Barra de progreso y estado
        progress_frame = ttk.Frame(control_frame)
        progress_frame.pack(side="right", fill="x", expand=True)
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress.pack(fill="x", pady=(0, 5))
        
        self.status_label = ttk.Label(progress_frame, text="Listo")
        self.status_label.pack()
        
        # Variables de instancia
        self.selected_file = None
        self.total_pages = 0
        self.extracted_formulas = []
    
    def preprocess_image_for_ocr(self, image):
        """Preprocesar imagen para mejorar reconocimiento OCR de fórmulas"""
        try:
            # Convertir a escala de grises si no lo está
            if image.mode != 'L':
                image = image.convert('L')
            
            # Aumentar contraste
            enhancer = ImageEnhance.Contrast(image)
            image = enhancer.enhance(2.0)
            
            # Aumentar nitidez
            enhancer = ImageEnhance.Sharpness(image)
            image = enhancer.enhance(2.0)
            
            # Aplicar filtro para reducir ruido
            image = image.filter(ImageFilter.MedianFilter())
            
            # Redimensionar si es muy pequeño (mejora OCR)
            width, height = image.size
            if width < 300 or height < 300:
                scale_factor = max(300/width, 300/height)
                new_size = (int(width * scale_factor), int(height * scale_factor))
                image = image.resize(new_size, Image.LANCZOS)
            
            return image
        except Exception as e:
            print(f"Error en preprocesamiento de imagen: {e}")
            return image
    
    def extract_with_pymupdf(self, file_path, selected_pages=None):
        """Extraer texto usando PyMuPDF con mejor soporte para fórmulas"""
        text = ""
        try:
            doc = fitz.open(file_path)
            total_pages = len(doc)
            
            if selected_pages is None:
                selected_pages = list(range(1, total_pages + 1))
            
            for page_num in selected_pages:
                if 1 <= page_num <= total_pages:
                    page = doc[page_num - 1]
                    
                    # Extraer texto con información de formato
                    blocks = page.get_text("dict")
                    page_text = self.process_pymupdf_blocks(blocks, page_num)
                    
                    if page_text.strip():
                        text += f"\n--- Página {page_num} ---\n"
                        text += page_text + "\n"
            
            doc.close()
            return text
            
        except Exception as e:
            raise Exception(f"Error con PyMuPDF: {str(e)}")
    
    def process_pymupdf_blocks(self, blocks, page_num):
        """Procesar bloques de texto de PyMuPDF para preservar formato de fórmulas"""
        text = ""
        
        try:
            for block in blocks["blocks"]:
                if "lines" in block:
                    block_text = ""
                    for line in block["lines"]:
                        line_text = ""
                        for span in line["spans"]:
                            span_text = span["text"]
                            
                            # Detectar si puede ser una fórmula por formato
                            font = span.get("font", "").lower()
                            size = span.get("size", 0)
                            
                            # Características de fórmulas: fuentes math, símbolos especiales
                            is_math = any(indicator in font for indicator in 
                                        ["math", "symbol", "times", "cambria"])
                            
                            if is_math or self.contains_math_symbols(span_text):
                                span_text = f"[MATH]{span_text}[/MATH]"
                            
                            line_text += span_text
                        
                        if line_text.strip():
                            block_text += line_text + "\n"
                    
                    if block_text.strip():
                        text += block_text + "\n"
        
        except Exception as e:
            print(f"Error procesando bloques PyMuPDF: {e}")
            # Fallback a extracción simple
            return blocks.get("text", "")
        
        return text
    
    def contains_math_symbols(self, text):
        """Detectar si el texto contiene símbolos matemáticos"""
        math_indicators = [
            '=', '+', '-', '×', '÷', '/', '^', '²', '³', '√', '∫', '∑', 'Σ',
            'α', 'β', 'γ', 'δ', 'ε', 'π', 'σ', 'μ', 'λ', '≈', '≠', '≤', '≥',
            '∞', '∂', '∇', 'sin', 'cos', 'tan', 'log', 'ln', 'exp'
        ]
        
        return any(symbol in text for symbol in math_indicators)
    
    def enhanced_ocr_extraction(self, file_path, selected_pages=None):
        """OCR mejorado con preprocesamiento de imágenes para fórmulas"""
        try:
            # Convertir PDF a imágenes con alta resolución
            doc = fitz.open(file_path)
            text = ""
            
            if selected_pages is None:
                selected_pages = list(range(1, len(doc) + 1))
            
            for page_num in selected_pages:
                if 1 <= page_num <= len(doc):
                    page = doc[page_num - 1]
                    
                    # Convertir página a imagen con alta resolución
                    matrix = fitz.Matrix(3.0, 3.0)  # Escala 3x para mejor calidad
                    pix = page.get_pixmap(matrix=matrix)
                    img_data = pix.tobytes("png")
                    
                    # Cargar imagen con PIL
                    image = Image.open(io.BytesIO(img_data))
                    
                    # Preprocesar imagen
                    if self.extraction_method.get() == "ocr_enhanced":
                        image = self.preprocess_image_for_ocr(image)
                    
                    # Convertir a bytes para OCR
                    img_bytes = io.BytesIO()
                    image.save(img_bytes, format='PNG')
                    img_bytes = img_bytes.getvalue()
                    
                    # Realizar OCR
                    page_text = self.ocr_space_api(img_bytes, page_num)
                    
                    if page_text.strip():
                        text += f"\n--- Página {page_num} (OCR Mejorado) ---\n"
                        text += page_text + "\n"
            
            doc.close()
            return text
            
        except Exception as e:
            raise Exception(f"Error en OCR mejorado: {str(e)}")
    
    def ocr_space_api(self, image_data, page_num):
        """Llamada a OCR Space API con configuración optimizada para fórmulas"""
        try:
            url = 'https://api.ocr.space/parse/image'
            
            files = {'file': ('page.png', image_data, 'image/png')}
            
            engine = self.ocr_engine.get().split()[0]
            language = self.ocr_language.get()
            
            data = {
                'apikey': self.ocr_space_key,
                'language': language,
                'detectOrientation': 'true',
                'scale': 'true',
                'OCREngine': engine,
                'isTable': 'true',
                'detectCheckbox': 'false',
                'checkboxTemplate': '0',
            }
            
            # Motor 3 tiene mejores capacidades para fórmulas
            if engine == "3":
                data['isTable'] = 'false'  # Mejor para fórmulas continuas
            
            response = requests.post(url, files=files, data=data, timeout=60)
            result = response.json()
            
            if result.get('IsErroredOnProcessing'):
                raise Exception(f"Error en OCR: {result.get('ErrorMessage', 'Error desconocido')}")
            
            text = ""
            if 'ParsedResults' in result and result['ParsedResults']:
                parsed_text = result['ParsedResults'][0].get('ParsedText', '')
                text = self.post_process_ocr_text(parsed_text)
            
            return text
            
        except Exception as e:
            raise Exception(f"Error en OCR API: {str(e)}")
    
    def fix_mathematical_symbols(self, text):
        """Corregir símbolos matemáticos mal reconocidos"""
        if not self.fix_symbols.get():
            return text
        
        corrections = self.math_patterns['symbol_corrections']
        
        # Aplicar correcciones de símbolos
        for correct_symbol, wrong_variants in corrections.items():
            for wrong in wrong_variants:
                # Usar contexto para evitar correcciones erróneas
                # Por ejemplo, solo corregir 'x' por '×' si está entre números
                if correct_symbol == '×' and wrong == 'x':
                    text = re.sub(r'(\d+)\s*x\s*(\d+)', r'\1 × \2', text)
                    text = re.sub(r'(\w+)\s*x\s*\(', r'\1 × (', text)
                else:
                    text = text.replace(wrong, correct_symbol)
        
        return text
    
    def detect_mathematical_structures(self, text):
        """Detectar y mejorar estructuras matemáticas"""
        if not self.detect_structures.get():
            return text
        
        # Detectar fracciones escritas como "a/b" y mejorar formato
        text = re.sub(r'(\w+)\s*/\s*(\w+)', r'(\1)/(\2)', text)
        
        # Detectar exponentes
        text = re.sub(r'(\w+)\^(\d+)', r'\1^\2', text)
        text = re.sub(r'(\w+)\^{([^}]+)}', r'\1^{\2}', text)
        
        # Detectar raíces cuadradas
        text = re.sub(r'sqrt\(([^)]+)\)', r'√(\1)', text)
        text = re.sub(r'√\s*(\w+)', r'√\1', text)
        
        # Mejorar sumatorias
        text = re.sub(r'(sum|SUM|Σ|∑)\s*(\w+)', r'Σ\2', text)
        
        # Mejorar integrales
        text = re.sub(r'integral\s*([^d]+)\s*d(\w+)', r'∫\1 d\2', text)
        
        # Mejorar logaritmos
        text = re.sub(r'log\s*\(\s*([^)]+)\s*\)', r'log(\1)', text)
        text = re.sub(r'ln\s*\(\s*([^)]+)\s*\)', r'ln(\1)', text)
        
        return text
    
    def extract_formulas(self, text):
        """Extraer fórmulas identificadas del texto"""
        formulas = []
        
        # Buscar bloques marcados como matemáticos
        math_blocks = re.findall(r'\[MATH\](.*?)\[/MATH\]', text)
        formulas.extend(math_blocks)
        
        # Buscar patrones de fórmulas
        formula_patterns = [
            r'[^.!?]*[=≈≠≤≥].*?[^.!?]*',  # Líneas con operadores de igualdad
            r'[^.!?]*[∫∑Σ].*?[^.!?]*',    # Líneas con integrales o sumatorias
            r'[^.!?]*\b\w+\^[\d{].*?[^.!?]*',  # Líneas con exponentes
            r'[^.!?]*√.*?[^.!?]*',        # Líneas con raíces
            r'[^.!?]*\([^)]*[+\-×÷*/][^)]*\).*?[^.!?]*',  # Expresiones en paréntesis
        ]
        
        for pattern in formula_patterns:
            matches = re.findall(pattern, text, re.MULTILINE)
            for match in matches:
                match = match.strip()
                if len(match) > 3 and match not in formulas:
                    formulas.append(match)
        
        # Filtrar fórmulas que son muy largas (probablemente no son fórmulas)
        formulas = [f for f in formulas if len(f) < 200]
        
        return formulas
    
    def post_process_ocr_text(self, text):
        """Post-procesamiento completo del texto OCR"""
        # Correcciones básicas de OCR
        text = text.replace('CINs', 'CINS')
        text = text.replace('ICATI CADE', 'ICADE')
        text = text.replace('•', '■')
        
        # Aplicar mejoras de fórmulas si están activadas
        if self.enhance_formulas.get():
            text = self.fix_mathematical_symbols(text)
            text = self.detect_mathematical_structures(text)
        
        # Mejorar formato general
        if self.preserve_layout.get():
            # Preservar espaciado de ecuaciones
            text = re.sub(r'([=≈≠≤≥])\s*([^=≈≠≤≥\n]*)', r'\1 \2', text)
            text = re.sub(r'([+\-×÷*/])\s*', r' \1 ', text)
            text = re.sub(r'\s+([+\-×÷*/])\s+', r' \1 ', text)
        
        return text
    
    def hybrid_extraction(self, file_path, selected_pages=None):
        """Método híbrido: combina extracción directa y OCR"""
        text = ""
        
        try:
            # Intentar PyMuPDF primero
            try:
                pymupdf_text = self.extract_with_pymupdf(file_path, selected_pages)
                if pymupdf_text.strip():
                    text += "=== EXTRACCIÓN DIRECTA (PyMuPDF) ===\n"
                    text += pymupdf_text + "\n"
            except:
                pass
            
            # Si no hay suficiente texto, o hay páginas que parecen escaneadas, usar OCR
            if len(text.strip()) < 100:  # Muy poco texto extraído
                try:
                    ocr_text = self.enhanced_ocr_extraction(file_path, selected_pages)
                    if ocr_text.strip():
                        text += "\n=== EXTRACCIÓN OCR ===\n"
                        text += ocr_text
                except Exception as e:
                    print(f"OCR falló: {e}")
            
            # Fallback a PyPDF2 si todo falla
            if not text.strip():
                text = self.extract_text_from_pdf(file_path, selected_pages)
                if text.strip():
                    text = "=== EXTRACCIÓN PYPDF2 ===\n" + text
            
            return text
            
        except Exception as e:
            raise Exception(f"Error en extracción híbrida: {str(e)}")
    
    def extract_text_from_pdf(self, file_path, selected_pages=None):
        """Extraer texto usando PyPDF2 (método original)"""
        text = ""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                
                if selected_pages is None:
                    selected_pages = list(range(1, total_pages + 1))
                
                for page_num in selected_pages:
                    if 1 <= page_num <= total_pages:
                        page = pdf_reader.pages[page_num - 1]
                        page_text = page.extract_text()
                        if page_text.strip():
                            text += f"\n--- Página {page_num} ---\n"
                            text += page_text + "\n"
        
        except Exception as e:
            raise Exception(f"Error al leer PDF con PyPDF2: {str(e)}")
        
        return text
    
    def toggle_page_selection(self):
        """Habilitar/deshabilitar la entrada de páginas específicas"""
        if self.all_pages.get():
            self.pages_entry.config(state="disabled")
            self.pages_entry.delete(0, tk.END)
        else:
            self.pages_entry.config(state="normal")
    
    def get_total_pages(self, file_path):
        """Obtener el número total de páginas del PDF"""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                return len(pdf_reader.pages)
        except:
            return 0
    
    def parse_page_ranges(self, page_input, total_pages):
        """Parsear la entrada de páginas y devolver una lista de números de página"""
        pages = set()
        
        if not page_input.strip():
            return list(range(1, total_pages + 1))
        
        try:
            parts = page_input.split(',')
            
            for part in parts:
                part = part.strip()
                
                if '-' in part:
                    start, end = part.split('-')
                    start = int(start.strip())
                    end = int(end.strip())
                    
                    if start < 1 or end > total_pages or start > end:
                        raise ValueError(f"Rango inválido: {part}")
                    
                    pages.update(range(start, end + 1))
                else:
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
            self.file_label.config(text=f"Archivo: {filename}")
            self.extract_btn.config(state="normal")
            self.status_label.config(text=f"Archivo seleccionado: {filename}")
            
            self.total_pages = self.get_total_pages(file_path)
            if self.total_pages > 0:
                self.total_pages_label.config(text=f"(Total: {self.total_pages} páginas)")
            else:
                self.total_pages_label.config(text="(No se pudo determinar el número de páginas)")
    
    def extract_text(self):
        """Extraer texto del PDF seleccionado usando el método configurado"""
        if not self.selected_file:
            messagebox.showwarning("Advertencia", "Por favor selecciona un archivo PDF primero.")
            return

        # Determinar qué páginas extraer
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

        # Mostrar progreso
        self.progress.start(10)
        method = self.extraction_method.get()
        
        if selected_pages:
            page_text = f"páginas: {', '.join(map(str, selected_pages[:5]))}{'...' if len(selected_pages) > 5 else ''}"
        else:
            page_text = "todas las páginas"
        
        self.status_label.config(text=f"Extrayendo con método {method} - {page_text}")
        self.root.update()

        try:
            # Seleccionar método de extracción
            if method == "direct":
                self.extracted_text = self.extract_text_from_pdf(self.selected_file, selected_pages)
            elif method == "ocr":
                self.extracted_text = self.enhanced_ocr_extraction(self.selected_file, selected_pages)
            elif method == "ocr_enhanced":
                self.extracted_text = self.enhanced_ocr_extraction(self.selected_file, selected_pages)
            else:  # hybrid
                self.extracted_text = self.hybrid_extraction(self.selected_file, selected_pages)
            
            # Post-procesar si hay mejoras activadas
            if self.enhance_formulas.get():
                self.extracted_text = self.post_process_ocr_text(self.extracted_text)
            
            # Extraer fórmulas
            self.extracted_formulas = self.extract_formulas(self.extracted_text)
            
            # Mostrar resultados
            self.display_results()
            
            # Actualizar estado
            pages_processed = len(selected_pages) if selected_pages else self.total_pages
            formulas_found = len(self.extracted_formulas)
            
            self.status_label.config(text=f"Completado: {pages_processed} páginas, {formulas_found} fórmulas detectadas")
            
            # Habilitar botones de guardado
            if self.extracted_text.strip():
                self.save_btn.config(state="normal")
            if self.extracted_formulas:
                self.save_formulas_btn.config(state="normal")

        except Exception as e:
            messagebox.showerror("Error", f"Error al extraer texto:\n{str(e)}")
            self.status_label.config(text="Error en la extracción")

        finally:
            self.progress.stop()
    
    def display_results(self):
        """Mostrar resultados en las pestañas correspondientes"""
        # Mostrar texto completo
        self.text_area.delete(1.0, tk.END)
        if self.extracted_text.strip():
            self.text_area.insert(1.0, self.extracted_text)
        else:
            self.text_area.insert(1.0, "No se pudo extraer texto del PDF.\n\n"
                                       "Sugerencias:\n"
                                       "1. Prueba el método 'Híbrido'\n"
                                       "2. Activa las opciones de mejora de fórmulas\n"
                                       "3. Usa 'OCR + Preprocesamiento' para PDFs escaneados\n"
                                       "4. Verifica que el PDF contenga texto o imágenes legibles")
        
        # Mostrar fórmulas detectadas
        self.formula_area.delete(1.0, tk.END)
        if self.extracted_formulas:
            formula_text = "FÓRMULAS MATEMÁTICAS DETECTADAS:\n"
            formula_text += "="*50 + "\n\n"
            
            for i, formula in enumerate(self.extracted_formulas, 1):
                formula_text += f"Fórmula {i}:\n"
                formula_text += f"{formula}\n\n"
                formula_text += "-"*30 + "\n\n"
            
            # Agregar estadísticas
            formula_text += f"\nESTADÍSTICAS:\n"
            formula_text += f"Total de fórmulas encontradas: {len(self.extracted_formulas)}\n"
            formula_text += f"Promedio de caracteres por fórmula: {sum(len(f) for f in self.extracted_formulas) // len(self.extracted_formulas)}\n"
            
            # Contar símbolos matemáticos
            all_formulas_text = " ".join(self.extracted_formulas)
            math_symbols = ['=', '+', '-', '×', '÷', '^', '√', '∫', '∑', 'π', 'α', 'β', 'γ']
            symbol_counts = {symbol: all_formulas_text.count(symbol) for symbol in math_symbols if symbol in all_formulas_text}
            
            if symbol_counts:
                formula_text += f"\nSímbolos matemáticos encontrados:\n"
                for symbol, count in symbol_counts.items():
                    formula_text += f"  {symbol}: {count} veces\n"
            
            self.formula_area.insert(1.0, formula_text)
        else:
            self.formula_area.insert(1.0, "No se detectaron fórmulas matemáticas.\n\n"
                                          "Para mejorar la detección:\n"
                                          "1. Activa 'Mejorar reconocimiento de fórmulas matemáticas'\n"
                                          "2. Usa el motor OCR 3 (Beta - Mejor para fórmulas)\n"
                                          "3. Asegúrate de que el PDF contenga fórmulas matemáticas\n"
                                          "4. Prueba con el método 'OCR + Preprocesamiento'")
    
    def save_text(self):
        """Guardar el texto extraído en un archivo TXT"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para guardar.")
            return
        
        if self.selected_file:
            pdf_name = Path(self.selected_file).stem
            default_name = f"{pdf_name}_texto_completo.txt"
        else:
            default_name = "texto_extraido.txt"
        
        file_path = filedialog.asksaveasfilename(
            title="Guardar texto completo",
            defaultextension=".txt",
            initialvalue=default_name,
            filetypes=[("Archivos de texto", "*.txt"), ("Todos los archivos", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(self.extracted_text)
                messagebox.showinfo("Éxito", f"Texto guardado en:\n{file_path}")
                self.status_label.config(text=f"Guardado: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar archivo:\n{str(e)}")
    
    def save_formulas(self):
        """Guardar las fórmulas detectadas en un archivo separado"""
        if not self.extracted_formulas:
            messagebox.showwarning("Advertencia", "No hay fórmulas para guardar.")
            return
        
        if self.selected_file:
            pdf_name = Path(self.selected_file).stem
            default_name = f"{pdf_name}_formulas.txt"
        else:
            default_name = "formulas_extraidas.txt"
        
        file_path = filedialog.asksaveasfilename(
            title="Guardar fórmulas matemáticas",
            defaultextension=".txt",
            initialvalue=default_name,
            filetypes=[
                ("Archivos de texto", "*.txt"), 
                ("JSON", "*.json"),
                ("Todos los archivos", "*.*")
            ]
        )
        
        if file_path:
            try:
                if file_path.endswith('.json'):
                    # Guardar como JSON estructurado
                    formula_data = {
                        "metadata": {
                            "source_file": os.path.basename(self.selected_file) if self.selected_file else "unknown",
                            "extraction_method": self.extraction_method.get(),
                            "total_formulas": len(self.extracted_formulas),
                            "extraction_date": str(Path(file_path).stat().st_mtime)
                        },
                        "formulas": [
                            {
                                "id": i+1,
                                "content": formula,
                                "length": len(formula)
                            }
                            for i, formula in enumerate(self.extracted_formulas)
                        ]
                    }
                    
                    with open(file_path, 'w', encoding='utf-8') as file:
                        json.dump(formula_data, file, indent=2, ensure_ascii=False)
                else:
                    # Guardar como texto plano
                    with open(file_path, 'w', encoding='utf-8') as file:
                        file.write("FÓRMULAS MATEMÁTICAS EXTRAÍDAS\n")
                        file.write("="*50 + "\n\n")
                        
                        for i, formula in enumerate(self.extracted_formulas, 1):
                            file.write(f"Fórmula {i}:\n")
                            file.write(f"{formula}\n\n")
                            file.write("-"*30 + "\n\n")
                
                messagebox.showinfo("Éxito", f"Fórmulas guardadas en:\n{file_path}")
                self.status_label.config(text=f"Fórmulas guardadas: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar fórmulas:\n{str(e)}")
    
    def clear_text(self):
        """Limpiar todas las áreas y resetear la aplicación"""
        self.text_area.delete(1.0, tk.END)
        self.formula_area.delete(1.0, tk.END)
        self.extracted_text = ""
        self.extracted_formulas = []
        self.selected_file = None
        self.total_pages = 0
        self.file_label.config(text="Ningún archivo seleccionado")
        self.total_pages_label.config(text="")
        self.pages_entry.delete(0, tk.END)
        self.all_pages.set(True)
        self.pages_entry.config(state="disabled")
        self.extract_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.save_formulas_btn.config(state="disabled")
        self.status_label.config(text="Listo")

def main():
    """Función principal con verificación de dependencias"""
    # Verificar dependencias básicas
    missing_deps = []
    
    try:
        import PyPDF2
    except ImportError:
        missing_deps.append("PyPDF2")
    
    try:
        import requests
    except ImportError:
        missing_deps.append("requests")
    
    try:
        import PIL
    except ImportError:
        missing_deps.append("Pillow")
    
    try:
        import fitz
    except ImportError:
        missing_deps.append("PyMuPDF")
    
    if missing_deps:
        print("Error: Faltan las siguientes dependencias:")
        for dep in missing_deps:
            print(f"  - {dep}")
        print("\nInstala las dependencias con:")
        print("pip install PyPDF2 requests Pillow PyMuPDF")
        return
    
    # Crear y ejecutar aplicación
    root = tk.Tk()
    app = PDFFormulaExtractor(root)
    
    # Configurar cierre limpio
    def on_closing():
        try:
            root.quit()
            root.destroy()
        except:
            pass
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("\nAplicación cerrada por el usuario")
    except Exception as e:
        print(f"Error inesperado: {e}")

if __name__ == "__main__":
    main()