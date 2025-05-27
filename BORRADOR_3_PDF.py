import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import PyPDF2
import requests
import base64
import os
from pathlib import Path
import tempfile
import re
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
import fitz  # PyMuPDF para extraer imágenes de fórmulas
from PIL import Image, ImageTk
import io
import threading

class FormulaProcessor:
    """Clase para procesar y detectar fórmulas matemáticas"""
    
    def __init__(self):
        # Patrones para detectar fórmulas matemáticas
        self.formula_patterns = [
            r'[A-Za-z]\s*=\s*[^=\n]+',  # Variable = expresión
            r'\([^)]*[+\-*/×÷][^)]*\)',  # Expresiones entre paréntesis
            r'[A-Za-z]+\s*[+\-*/×÷]\s*[A-Za-z0-9()]+',  # Operaciones básicas
            r'\d+\s*[×*]\s*\([^)]+\)',  # Multiplicaciones con paréntesis
            r'[A-Za-z]\s*:\s*[^A-Z\n]{10,}',  # Definiciones de variables
            r'√\w+|∑|∫|∂|α|β|γ|δ|π|θ|λ|μ|σ|Σ|Π',  # Símbolos matemáticos
            r'^[A-Za-z]\s*=\s*.+?=.+?=.+',  # NUEVO: Ecuaciones largas con varios signos igual
            r'^[A-Za-z]\s*=\s*[\d\w\s\+\-\*/\.,\(\)]+\/[\d\w\.,\(\)]+',  # NUEVO: Fracciones largas
        ]
    
    def detect_formulas(self, text):
        """Detectar fórmulas en el texto"""
        formulas = []
        lines = text.split('\n')
        
        for i, line in enumerate(lines):
            line_clean = line.strip()
            if not line_clean:
                continue
                
            # Verificar si la línea contiene una fórmula
            for pattern in self.formula_patterns:
                if re.search(pattern, line_clean):
                    formulas.append({
                        'line_number': i + 1,
                        'content': line_clean,
                        'original_line': line,
                        'page': self.extract_page_from_context(lines, i)
                    })
                    break
        
        return formulas
    
    def extract_page_from_context(self, lines, current_line):
        """Extraer número de página del contexto"""
        # Buscar hacia atrás para encontrar el marcador de página
        for i in range(current_line, max(0, current_line - 10), -1):
            if '--- Página' in lines[i]:
                match = re.search(r'Página (\d+)', lines[i])
                if match:
                    return int(match.group(1))
        return 1
    
    def simplify_formula(self, formula_text):
        """Simplificar fórmula a una línea legible"""
        # Limpiar espacios extra
        simplified = re.sub(r'\s+', ' ', formula_text.strip())
        
        # Reemplazar símbolos especiales con equivalentes ASCII
        replacements = {
            '×': '*',
            '÷': '/',
            '−': '-',
            '±': '+/-',
            '≤': '<=',
            '≥': '>=',
            '≠': '!=',
            '≡': '===',
            '√': 'sqrt',
            'π': 'pi',
            'α': 'alpha',
            'β': 'beta',
            'γ': 'gamma',
            'δ': 'delta',
            'θ': 'theta',
            'λ': 'lambda',
            'μ': 'mu',
            'σ': 'sigma'
        }
        
        for symbol, replacement in replacements.items():
            simplified = simplified.replace(symbol, replacement)
        
        return simplified
    
    def format_formula_latex(self, formula_text):
        """Convertir fórmula a formato LaTeX-like para mejor visualización"""
        # Esta función podría expandirse para generar LaTeX real
        latex_like = formula_text
        
        # Detectar fracciones simples y formatearlas
        fraction_pattern = r'(\w+)\s*/\s*(\w+)'
        latex_like = re.sub(fraction_pattern, r'(\1)/(\2)', latex_like)
        
        # Detectar exponentes simples
        exponent_pattern = r'(\w+)\^(\w+)'
        latex_like = re.sub(exponent_pattern, r'\1^{\2}', latex_like)
        
        return latex_like

class PDFTextExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Texto PDF con Export a Word")
        self.root.geometry("1100x800")
        
        # Variables
        self.extracted_text = ""
        self.processed_formulas = []
        self.formula_processor = FormulaProcessor()
        self.ocr_api_key = "K83967071688957"
        
        # Variables para preview
        self.preview_window = None
        self.word_content = None
        
        self.setup_ui()
    
    def setup_ui(self):
        # Crear notebook para pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Pestaña principal
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="Extracción de Texto")
        
        # Pestaña de configuración de fórmulas
        self.formula_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.formula_frame, text="Configuración de Fórmulas")
        
        # Pestaña de limpieza de texto
        self.clean_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.clean_frame, text="Limpiar Texto")
        self.setup_clean_tab()
        
        self.setup_main_tab()
        self.setup_formula_tab()
    
    def setup_main_tab(self):
        # Frame principal con padding
        main_container = ttk.Frame(self.main_frame, padding="10")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(main_container, text="Extractor PDF con Export a Word", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Frame para selección de archivo
        file_frame = ttk.Frame(main_container)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.select_btn = ttk.Button(file_frame, text="Seleccionar PDF", 
                                    command=self.select_file)
        self.select_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="Ningún archivo seleccionado")
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Frame para opciones OCR
        ocr_frame = ttk.LabelFrame(main_container, text="Opciones OCR", padding="10")
        ocr_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Primera fila de opciones OCR
        ocr_row1 = ttk.Frame(ocr_frame)
        ocr_row1.pack(fill=tk.X)
        
        self.use_ocr = tk.BooleanVar()
        self.ocr_checkbox = ttk.Checkbutton(ocr_row1, text="Usar OCR (PDFs escaneados)", 
                                           variable=self.use_ocr)
        self.ocr_checkbox.pack(side=tk.LEFT)
        
        self.improve_formulas = tk.BooleanVar(value=True)
        self.formulas_checkbox = ttk.Checkbutton(ocr_row1, text="Procesar fórmulas matemáticas", 
                                                variable=self.improve_formulas)
        self.formulas_checkbox.pack(side=tk.LEFT, padx=(20, 0))
        
        # Segunda fila - Motor OCR
        ocr_row2 = ttk.Frame(ocr_frame)
        ocr_row2.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(ocr_row2, text="Motor OCR:").pack(side=tk.LEFT)
        self.ocr_engine = tk.StringVar(value="2")
        ocr_combo = ttk.Combobox(ocr_row2, textvariable=self.ocr_engine, 
                                values=["1 (Básico)", "2 (Avanzado)", "3 (Beta)"], 
                                width=15, state="readonly")
        ocr_combo.pack(side=tk.LEFT, padx=(5, 0))
        
        # Frame para selección de páginas
        page_frame = ttk.LabelFrame(main_container, text="Selección de Páginas", padding="10")
        page_frame.pack(fill=tk.X, pady=(0, 10))
        
        page_row1 = ttk.Frame(page_frame)
        page_row1.pack(fill=tk.X)
        
        self.all_pages = tk.BooleanVar(value=True)
        self.all_pages_checkbox = ttk.Checkbutton(page_row1, text="Todas las páginas", 
                                                 variable=self.all_pages,
                                                 command=self.toggle_page_selection)
        self.all_pages_checkbox.pack(side=tk.LEFT)
        
        ttk.Label(page_row1, text="Páginas específicas (ej: 1,3,5-8):").pack(side=tk.LEFT, padx=(20, 5))
        self.pages_entry = ttk.Entry(page_row1, width=20, state="disabled")
        self.pages_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        self.total_pages_label = ttk.Label(page_row1, text="", foreground="gray")
        self.total_pages_label.pack(side=tk.LEFT)
        
        # Área de texto
        text_frame = ttk.LabelFrame(main_container, text="Texto Extraído", padding="5")
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.text_area = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, 
                                                  font=("Consolas", 10))
        self.text_area.pack(fill=tk.BOTH, expand=True)
        
        # Frame para botones
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Botones principales
        self.extract_btn = ttk.Button(button_frame, text="Extraer Texto", 
                                     command=self.extract_text, state="disabled")
        self.extract_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.preview_btn = ttk.Button(button_frame, text="Vista Previa Word", 
                                     command=self.preview_word, state="disabled")
        self.preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.export_btn = ttk.Button(button_frame, text="Exportar a Word", 
                                    command=self.export_to_word, state="disabled")
        self.export_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.save_btn = ttk.Button(button_frame, text="Guardar TXT", 
                                  command=self.save_text, state="disabled")
        self.save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_btn = ttk.Button(button_frame, text="Limpiar", 
                                   command=self.clear_text)
        self.clear_btn.pack(side=tk.LEFT)
        
        # Barra de progreso y estado
        self.progress = ttk.Progressbar(main_container, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=(10, 5))
        
        self.status_label = ttk.Label(main_container, text="Listo")
        self.status_label.pack()
        
        # Variables de estado
        self.selected_file = None
        self.total_pages = 0
    
    def setup_formula_tab(self):
        """Configurar la pestaña de opciones de fórmulas"""
        container = ttk.Frame(self.formula_frame, padding="20")
        container.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title = ttk.Label(container, text="Configuración del Procesamiento de Fórmulas", 
                         font=("Arial", 14, "bold"))
        title.pack(pady=(0, 20))
        
        # Opciones de procesamiento
        self.formula_option = tk.StringVar(value="simplified")
        
        options_frame = ttk.LabelFrame(container, text="Método de Procesamiento", padding="15")
        options_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Radiobutton(options_frame, text="Fórmulas simplificadas en una línea (a+b)/c = 12", 
                       variable=self.formula_option, value="simplified").pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(options_frame, text="Fórmulas con formato matemático mejorado", 
                       variable=self.formula_option, value="formatted").pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(options_frame, text="Insertar imagen recortada de la fórmula original", 
                       variable=self.formula_option, value="image").pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(options_frame, text="Placeholder simple (Aquí hay una fórmula de la página X)", 
                       variable=self.formula_option, value="placeholder").pack(anchor=tk.W, pady=2)
        
        # Configuraciones adicionales
        config_frame = ttk.LabelFrame(container, text="Configuraciones Adicionales", padding="15")
        config_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.detect_variables = tk.BooleanVar(value=True)
        ttk.Checkbutton(config_frame, text="Detectar definiciones de variables (A: descripción)", 
                       variable=self.detect_variables).pack(anchor=tk.W, pady=2)
        
        self.group_formulas = tk.BooleanVar(value=True)
        ttk.Checkbutton(config_frame, text="Agrupar fórmulas relacionadas", 
                       variable=self.group_formulas).pack(anchor=tk.W, pady=2)
        
        self.add_formula_index = tk.BooleanVar(value=True)
        ttk.Checkbutton(config_frame, text="Añadir índice de fórmulas al documento", 
                       variable=self.add_formula_index).pack(anchor=tk.W, pady=2)
        
        # Vista previa de fórmulas detectadas
        preview_frame = ttk.LabelFrame(container, text="Fórmulas Detectadas", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        self.formula_tree = ttk.Treeview(preview_frame, columns=('Page', 'Formula'), show='headings', height=8)
        self.formula_tree.heading('Page', text='Página')
        self.formula_tree.heading('Formula', text='Fórmula')
        self.formula_tree.column('Page', width=80)
        self.formula_tree.column('Formula', width=400)
        
        formula_scroll = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.formula_tree.yview)
        self.formula_tree.configure(yscrollcommand=formula_scroll.set)
        
        self.formula_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        formula_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_clean_tab(self):
        """Configura la pestaña para limpiar texto"""
        container = ttk.Frame(self.clean_frame, padding="20")
        container.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(container, text="Eliminar Palabras o Frases del Texto Extraído", font=("Arial", 14, "bold"))
        title.pack(pady=(0, 15))

        instr = ttk.Label(container, text="Escribe cada palabra o frase a eliminar (una por línea):")
        instr.pack(anchor=tk.W)

        self.clean_text_box = scrolledtext.ScrolledText(container, height=6, font=("Consolas", 10))
        self.clean_text_box.pack(fill=tk.X, pady=(0, 10))

        # Opción para añadir "Página X" al inicio de cada página
        self.add_page_index = tk.BooleanVar(value=True)
        page_index_check = ttk.Checkbutton(container, text='Añadir "Página X" al inicio de cada página',
                                           variable=self.add_page_index)
        page_index_check.pack(anchor=tk.W, pady=(0, 10))

        # Botón para limpiar el texto
        clean_btn = ttk.Button(container, text="Limpiar Texto Extraído", command=self.clean_extracted_text)
        clean_btn.pack(pady=(10, 0))

        # Mensaje de estado
        self.clean_status = ttk.Label(container, text="", foreground="green")
        self.clean_status.pack(pady=(10, 0))

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
    
    def extract_text_from_pdf(self, file_path, selected_pages=None):
        """Extraer texto usando PyPDF2"""
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
            raise Exception(f"Error al leer PDF: {str(e)}")
        
        return text
    
    def extract_text_with_ocr(self, file_path, selected_pages=None):
        """Extraer texto usando OCR Space API"""
        try:
            url = 'https://api.ocr.space/parse/image'

            pdf_to_send = file_path
            temp_file = None
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

            with open(pdf_to_send, 'rb') as file:
                files = {'file': file}
                
                engine = self.ocr_engine.get().split()[0]
                
                data = {
                    'apikey': self.ocr_api_key,
                    'language': 'eng',
                    'detectOrientation': 'true',
                    'scale': 'true',
                    'OCREngine': engine,
                    'filetype': 'PDF',
                    'isTable': 'true',
                }
                
                response = requests.post(url, files=files, data=data)
                result = response.json()

                if result.get('IsErroredOnProcessing'):
                    raise Exception(f"Error en OCR: {result.get('ErrorMessage', 'Error desconocido')}")

                text = ""
                if 'ParsedResults' in result:
                    for i, page_result in enumerate(result['ParsedResults']):
                        if 'ParsedText' in page_result:
                            page_text = page_result['ParsedText']
                            text += f"\n--- Página {selected_pages[i] if selected_pages else i + 1} (OCR) ---\n"
                            text += page_text + "\n"

                return text

        except Exception as e:
            raise Exception(f"Error en OCR: {str(e)}")
        finally:
            if 'temp_file' in locals() and temp_file is not None:
                try:
                    os.unlink(temp_file.name)
                except Exception:
                    pass
    
    def process_formulas_in_text(self, text):
        """Procesar fórmulas en el texto según la configuración"""
        if not self.improve_formulas.get():
            return text
        
        # Detectar fórmulas
        self.processed_formulas = self.formula_processor.detect_formulas(text)
        
        # Actualizar la vista de fórmulas en la interfaz
        self.update_formula_tree()
        
        # Procesar según la opción seleccionada
        processed_text = text
        option = self.formula_option.get()
        
        for formula in self.processed_formulas:
            original = formula['original_line']
            
            if option == "simplified":
                replacement = self.formula_processor.simplify_formula(formula['content'])
                replacement = f"[FÓRMULA: {replacement}]"
            elif option == "formatted":
                replacement = self.formula_processor.format_formula_latex(formula['content'])
                replacement = f"[FÓRMULA FORMATEADA: {replacement}]"
            elif option == "image":
                replacement = f"[IMAGEN DE FÓRMULA DE LA PÁGINA {formula['page']}]"
            else:  # placeholder
                replacement = f"[Aquí hay una fórmula de la página {formula['page']}]"
            
            processed_text = processed_text.replace(original, replacement)
        
        return processed_text
    
    def update_formula_tree(self):
        """Actualizar la vista de fórmulas detectadas"""
        # Limpiar el tree
        for item in self.formula_tree.get_children():
            self.formula_tree.delete(item)
        
        # Añadir fórmulas detectadas
        for formula in self.processed_formulas:
            self.formula_tree.insert('', 'end', values=(
                f"Página {formula['page']}", 
                formula['content'][:100] + "..." if len(formula['content']) > 100 else formula['content']
            ))
    
    def extract_text(self):
        """Extraer texto del PDF seleccionado"""
        if not self.selected_file:
            messagebox.showwarning("Advertencia", "Por favor selecciona un archivo PDF primero.")
            return

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

        self.progress.start(10)
        self.status_label.config(text="Extrayendo texto...")
        self.root.update()

        try:
            if self.use_ocr.get():
                self.extracted_text = self.extract_text_with_ocr(self.selected_file, selected_pages)
                if not self.extracted_text.strip():
                    self.extracted_text = self.extract_text_from_pdf(self.selected_file, selected_pages)
                    messagebox.showinfo("Info", "OCR no encontró texto. Se usó extracción normal.")
            else:
                self.extracted_text = self.extract_text_from_pdf(self.selected_file, selected_pages)

            # Procesar fórmulas
            if self.improve_formulas.get():
                self.extracted_text = self.process_formulas_in_text(self.extracted_text)

            # Mostrar texto
            self.text_area.delete(1.0, tk.END)
            if self.extracted_text.strip():
                self.text_area.insert(1.0, self.extracted_text)
                self.save_btn.config(state="normal")
                self.preview_btn.config(state="normal")
                self.export_btn.config(state="normal")

                pages_processed = len(selected_pages) if selected_pages else self.total_pages
                method = "OCR" if self.use_ocr.get() else "extracción directa"
                formulas_found = len(self.processed_formulas) if hasattr(self, 'processed_formulas') else 0
                self.status_label.config(text=f"Texto extraído con {method} de {pages_processed} página(s). {formulas_found} fórmulas detectadas.")
            else:
                self.text_area.insert(1.0, "No se pudo extraer texto del PDF.\n\n"
                                           "Posibles soluciones:\n"
                                           "1. Activa la opción 'Usar OCR' si es un PDF escaneado\n"
                                           "2. Prueba con diferentes motores OCR\n"
                                           "3. Verifica que el PDF contenga texto seleccionable")
                self.status_label.config(text="No se encontró texto")

        except Exception as e:
            messagebox.showerror("Error", f"Error al extraer texto:\n{str(e)}")
            self.status_label.config(text="Error en la extracción")

        finally:
            self.progress.stop()
    
    def create_word_document(self):
        """Crear documento de Word con el texto procesado"""
        doc = Document()
        
        # Título del documento
        if self.selected_file:
            title = f"Texto extraído de: {os.path.basename(self.selected_file)}"
        else:
            title = "Texto extraído de PDF"
        
        title_paragraph = doc.add_heading(title, level=1)
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Información del documento
        info_paragraph = doc.add_paragraph()
        info_paragraph.add_run(f"Fecha de extracción: {tk.datetime.now().strftime('%d/%m/%Y %H:%M')}\n")
        info_paragraph.add_run(f"Páginas procesadas: {self.total_pages}\n")
        if hasattr(self, 'processed_formulas'):
            info_paragraph.add_run(f"Fórmulas detectadas: {len(self.processed_formulas)}\n")
        info_paragraph.add_run(f"Método de procesamiento de fórmulas: {self.formula_option.get()}")
        
        # Agregar índice de fórmulas si está activado
        if self.add_formula_index.get() and hasattr(self, 'processed_formulas') and self.processed_formulas:
            doc.add_heading("Índice de Fórmulas", level=2)
            for i, formula in enumerate(self.processed_formulas, 1):
                formula_paragraph = doc.add_paragraph(f"{i}. Página {formula['page']}: {formula['content'][:100]}...")
                formula_paragraph.style = 'List Number'
        
        doc.add_page_break()
        
        # Procesar el texto línea por línea
        lines = self.extracted_text.split('\n')
        current_page = 1
        
        for line in lines:
            line_clean = line.strip()
            
            # Detectar marcadores de página
            if '--- Página' in line:
                page_match = re.search(r'Página (\d+)', line)
                if page_match:
                    current_page = int(page_match.group(1))
                    doc.add_heading(f"Página {current_page}", level=2)
                continue
            
            if not line_clean:
                doc.add_paragraph()
                continue
            
            # Verificar si es una fórmula procesada
            if line_clean.startswith('[FÓRMULA'):
                formula_paragraph = doc.add_paragraph()
                formula_run = formula_paragraph.add_run(line_clean)
                formula_run.bold = True
                formula_run.font.size = Pt(12)
                formula_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif line_clean.startswith('[IMAGEN DE FÓRMULA'):
                # Para imágenes de fórmulas, añadir placeholder mejorado
                img_paragraph = doc.add_paragraph()
                img_run = img_paragraph.add_run(f"📐 {line_clean}")
                img_run.italic = True
                img_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif line_clean.startswith('[Aquí hay una fórmula'):
                placeholder_paragraph = doc.add_paragraph()
                placeholder_run = placeholder_paragraph.add_run(f"🔢 {line_clean}")
                placeholder_run.italic = True
                placeholder_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                # Texto normal
                doc.add_paragraph(line_clean)
        
        return doc
    
    def preview_word(self):
        """Mostrar vista previa del documento Word"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para previsualizar.")
            return
        
        # Crear documento
        try:
            self.word_content = self.create_word_document()
            self.show_preview_window()
        except Exception as e:
            messagebox.showerror("Error", f"Error al crear vista previa:\n{str(e)}")
    
    def show_preview_window(self):
        """Mostrar ventana de vista previa"""
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.lift()
            return
        
        self.preview_window = tk.Toplevel(self.root)
        self.preview_window.title("Vista Previa - Documento Word")
        self.preview_window.geometry("800x600")
        
        # Frame principal
        preview_frame = ttk.Frame(self.preview_window, padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Información del documento
        info_frame = ttk.LabelFrame(preview_frame, text="Información del Documento", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        if self.selected_file:
            ttk.Label(info_frame, text=f"Archivo origen: {os.path.basename(self.selected_file)}").pack(anchor=tk.W)
        ttk.Label(info_frame, text=f"Páginas: {self.total_pages}").pack(anchor=tk.W)
        if hasattr(self, 'processed_formulas'):
            ttk.Label(info_frame, text=f"Fórmulas detectadas: {len(self.processed_formulas)}").pack(anchor=tk.W)
        ttk.Label(info_frame, text=f"Procesamiento de fórmulas: {self.get_formula_option_text()}").pack(anchor=tk.W)
        
        # Vista previa del contenido
        content_frame = ttk.LabelFrame(preview_frame, text="Vista Previa del Contenido", padding="5")
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        preview_text = scrolledtext.ScrolledText(content_frame, wrap=tk.WORD, font=("Arial", 10))
        preview_text.pack(fill=tk.BOTH, expand=True)
        
        # Generar vista previa del texto
        preview_content = self.generate_preview_text()
        preview_text.insert(1.0, preview_content)
        preview_text.config(state=tk.DISABLED)
        
        # Botones
        button_frame = ttk.Frame(preview_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Exportar a Word", 
                  command=self.export_from_preview).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Modificar Configuración", 
                  command=self.modify_config).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Cerrar Vista Previa", 
                  command=self.preview_window.destroy).pack(side=tk.RIGHT)
    
    def get_formula_option_text(self):
        """Obtener texto descriptivo de la opción de fórmulas seleccionada"""
        option = self.formula_option.get()
        options_text = {
            "simplified": "Fórmulas simplificadas en una línea",
            "formatted": "Fórmulas con formato matemático mejorado",
            "image": "Insertar imagen recortada de la fórmula",
            "placeholder": "Placeholder simple"
        }
        return options_text.get(option, "No especificado")
    
    def generate_preview_text(self):
        """Generar texto de vista previa"""
        preview = f"DOCUMENTO: Texto extraído de PDF\n"
        preview += "=" * 50 + "\n\n"
        
        if self.selected_file:
            preview += f"Archivo origen: {os.path.basename(self.selected_file)}\n"
        preview += f"Páginas procesadas: {self.total_pages}\n"
        if hasattr(self, 'processed_formulas'):
            preview += f"Fórmulas detectadas: {len(self.processed_formulas)}\n"
        preview += f"Procesamiento de fórmulas: {self.get_formula_option_text()}\n\n"
        
        # Índice de fórmulas si está activado
        if self.add_formula_index.get() and hasattr(self, 'processed_formulas') and self.processed_formulas:
            preview += "ÍNDICE DE FÓRMULAS:\n"
            preview += "-" * 20 + "\n"
            for i, formula in enumerate(self.processed_formulas, 1):
                preview += f"{i}. Página {formula['page']}: {formula['content'][:80]}...\n"
            preview += "\n" + "=" * 50 + "\n\n"
        
        # Mostrar una muestra del contenido (primeras 2000 caracteres)
        preview += "MUESTRA DEL CONTENIDO:\n"
        preview += "-" * 25 + "\n"
        content_sample = self.extracted_text[:2000]
        if len(self.extracted_text) > 2000:
            content_sample += "\n\n... (contenido truncado para vista previa) ..."
        
        preview += content_sample
        
        return preview
    
    def modify_config(self):
        """Cambiar a la pestaña de configuración de fórmulas"""
        self.notebook.select(self.formula_frame)
        if self.preview_window:
            self.preview_window.destroy()
    
    def export_from_preview(self):
        """Exportar a Word desde la vista previa"""
        if self.preview_window:
            self.preview_window.destroy()
        self.export_to_word()
    
    def export_to_word(self):
        """Exportar el texto a un documento Word"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para exportar.")
            return
        
        # Sugerir nombre de archivo
        if self.selected_file:
            pdf_name = Path(self.selected_file).stem
            default_name = f"{pdf_name}_exportado.docx"
        else:
            default_name = "texto_extraido.docx"
        
        file_path = filedialog.asksaveasfilename(
            title="Guardar documento Word",
            defaultextension=".docx",
            initialfile=default_name,  # <-- Cambiado aquí
            filetypes=[("Documentos Word", "*.docx"), ("Todos los archivos", "*.*")]
        )
        
        if file_path:
            try:
                self.progress.start(10)
                self.status_label.config(text="Exportando a Word...")
                self.root.update()
                
                # Crear y guardar documento
                if not hasattr(self, 'word_content') or self.word_content is None:
                    self.word_content = self.create_word_document()
                
                self.word_content.save(file_path)
                
                messagebox.showinfo("Éxito", f"Documento Word guardado en:\n{file_path}")
                self.status_label.config(text=f"Exportado: {os.path.basename(file_path)}")
                
                # Preguntar si abrir el archivo
                if messagebox.askyesno("Abrir Documento", "¿Deseas abrir el documento Word ahora?"):
                    try:
                        os.startfile(file_path)  # Windows
                    except:
                        try:
                            os.system(f"open '{file_path}'")  # macOS
                        except:
                            os.system(f"xdg-open '{file_path}'")  # Linux
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al exportar a Word:\n{str(e)}")
                self.status_label.config(text="Error en la exportación")
            finally:
                self.progress.stop()
    
    def save_text(self):
        """Guardar el texto extraído en un archivo TXT"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para guardar.")
            return
        
        if self.selected_file:
            pdf_name = Path(self.selected_file).stem
            default_name = f"{pdf_name}_texto.txt"
        else:
            default_name = "texto_extraido.txt"
        
        file_path = filedialog.asksaveasfilename(
            title="Guardar texto como",
            defaultextension=".txt",
            initialfile=default_name,  # <-- Cambiado aquí
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
    
    def clear_text(self):
        """Limpiar el área de texto y resetear"""
        self.text_area.delete(1.0, tk.END)
        self.extracted_text = ""
        self.processed_formulas = []
        self.word_content = None
        self.selected_file = None
        self.total_pages = 0
        self.file_label.config(text="Ningún archivo seleccionado")
        self.total_pages_label.config(text="")
        self.pages_entry.delete(0, tk.END)
        self.all_pages.set(True)
        self.pages_entry.config(state="disabled")
        self.extract_btn.config(state="disabled")
        self.preview_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.status_label.config(text="Listo")
        
        # Limpiar vista de fórmulas
        for item in self.formula_tree.get_children():
            self.formula_tree.delete(item)
        
        # Cerrar ventana de vista previa si está abierta
        if self.preview_window and self.preview_window.winfo_exists():
            self.preview_window.destroy()

    def clean_extracted_text(self):
        """Elimina palabras/frases del texto extraído y añade índice de página si se desea"""
        if not self.extracted_text:
            self.clean_status.config(text="No hay texto extraído.", foreground="red")
            return

        # Obtener palabras/frases a eliminar
        to_remove = self.clean_text_box.get(1.0, tk.END).strip().split('\n')
        cleaned_text = self.extracted_text

        for item in to_remove:
            item = item.strip()
            if item:
                cleaned_text = cleaned_text.replace(item, "")

        # Añadir "Página X" al inicio de cada página si está activado
        if self.add_page_index.get():
            # Reemplaza los marcadores de página por "Página X" al inicio de línea
            def add_page_header(match):
                page_num = match.group(1)
                return f"\nPágina {page_num}\n"
            cleaned_text = re.sub(r"\n--- Página (\d+) ---\n", add_page_header, cleaned_text)
        else:
            # Quita los marcadores de página si no se quiere el índice
            cleaned_text = re.sub(r"\n--- Página (\d+) ---\n", "\n", cleaned_text)

        self.extracted_text = cleaned_text
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(1.0, self.extracted_text)
        self.clean_status.config(text="Texto limpiado correctamente.", foreground="green")

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
        from docx import Document
    except ImportError:
        missing_deps.append("python-docx")
    
    try:
        import fitz
    except ImportError:
        missing_deps.append("PyMuPDF")
    
    try:
        from PIL import Image
    except ImportError:
        missing_deps.append("Pillow")
    
    if missing_deps:
        print("Error: Faltan las siguientes dependencias:")
        for dep in missing_deps:
            print(f"  - {dep}")
        print("\nInstala las dependencias con:")
        print("pip install PyPDF2 requests python-docx PyMuPDF Pillow")
        print("\nO instala todas las dependencias con:")
        print("pip install -r requirements.txt")
        
        # Crear archivo requirements.txt
        requirements_content = """PyPDF2>=3.0.0
requests>=2.28.0
python-docx>=0.8.11
PyMuPDF>=1.20.0
Pillow>=9.0.0
tkinter"""
        
        try:
            with open("requirements.txt", "w") as f:
                f.write(requirements_content)
            print("\nSe ha creado el archivo 'requirements.txt' con las dependencias necesarias.")
        except:
            pass
        
        return
    
    # Importar datetime aquí para evitar conflictos
    import datetime
    tk.datetime = datetime.datetime
    
    root = tk.Tk()
    app = PDFTextExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()