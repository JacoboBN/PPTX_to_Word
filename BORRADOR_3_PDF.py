### BASE ###

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import PyPDF2
import requests
import base64
import os
from pathlib import Path
import tempfile
import re

class PDFTextExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Texto PDF Mejorado")
        self.root.geometry("900x700")
        
        # Variable para almacenar el texto extraído
        self.extracted_text = ""
        
        # API Key para OCR Space (opcional - puedes usar la gratuita)
        self.ocr_api_key = "K83967071688957"  # Tu API key
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar expansión
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Extractor de Texto PDF Mejorado", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Botón para seleccionar archivo
        self.select_btn = ttk.Button(main_frame, text="Seleccionar PDF", 
                                    command=self.select_file)
        self.select_btn.grid(row=1, column=0, padx=(0, 10), sticky=tk.W)
        
        # Label para mostrar archivo seleccionado
        self.file_label = ttk.Label(main_frame, text="Ningún archivo seleccionado")
        self.file_label.grid(row=1, column=1, sticky=(tk.W, tk.E))
        
        # Frame para opciones OCR
        ocr_frame = ttk.LabelFrame(main_frame, text="Opciones OCR", padding="5")
        ocr_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        ocr_frame.columnconfigure(1, weight=1)
        
        # Checkbox para OCR
        self.use_ocr = tk.BooleanVar()
        self.ocr_checkbox = ttk.Checkbutton(ocr_frame, text="Usar OCR (para PDFs escaneados)", 
                                           variable=self.use_ocr)
        self.ocr_checkbox.grid(row=0, column=0, sticky=tk.W)
        
        # Checkbox para mejorar fórmulas
        self.improve_formulas = tk.BooleanVar(value=True)
        self.formulas_checkbox = ttk.Checkbutton(ocr_frame, text="Mejorar reconocimiento de fórmulas", 
                                                variable=self.improve_formulas)
        self.formulas_checkbox.grid(row=0, column=1, padx=(20, 0), sticky=tk.W)
        
        # Selector de motor OCR
        ttk.Label(ocr_frame, text="Motor OCR:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.ocr_engine = tk.StringVar(value="2")
        ocr_combo = ttk.Combobox(ocr_frame, textvariable=self.ocr_engine, 
                                values=["1 (Básico)", "2 (Avanzado)", "3 (Beta)"], 
                                width=15, state="readonly")
        ocr_combo.grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=(5, 0))
        
        # Frame para selección de páginas
        page_frame = ttk.Frame(main_frame)
        page_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        page_frame.columnconfigure(1, weight=1)
        
        # Checkbox para todas las páginas
        self.all_pages = tk.BooleanVar(value=True)
        self.all_pages_checkbox = ttk.Checkbutton(page_frame, text="Todas las páginas", 
                                                 variable=self.all_pages,
                                                 command=self.toggle_page_selection)
        self.all_pages_checkbox.grid(row=0, column=0, sticky=tk.W)
        
        # Label y entry para páginas específicas
        self.pages_label = ttk.Label(page_frame, text="Páginas específicas (ej: 1,3,5-8):")
        self.pages_label.grid(row=0, column=1, padx=(20, 5), sticky=tk.W)
        
        self.pages_entry = ttk.Entry(page_frame, width=20, state="disabled")
        self.pages_entry.grid(row=0, column=2, sticky=tk.W)
        
        # Label informativo sobre total de páginas
        self.total_pages_label = ttk.Label(page_frame, text="", foreground="gray")
        self.total_pages_label.grid(row=0, column=3, padx=(10, 0), sticky=tk.W)
        
        # Área de texto para mostrar el contenido
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.text_area = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, 
                                                  width=80, height=25, font=("Consolas", 10))
        self.text_area.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Frame para botones inferiores
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=3, pady=(20, 0))
        
        # Botón para extraer texto
        self.extract_btn = ttk.Button(button_frame, text="Extraer Texto", 
                                     command=self.extract_text, state="disabled")
        self.extract_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón para guardar texto
        self.save_btn = ttk.Button(button_frame, text="Guardar como TXT", 
                                  command=self.save_text, state="disabled")
        self.save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón para guardar como Word
        self.save_word_btn = ttk.Button(button_frame, text="Guardar como Word", 
                                       command=self.save_word, state="disabled")
        self.save_word_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón para limpiar
        self.clear_btn = ttk.Button(button_frame, text="Limpiar", 
                                   command=self.clear_text)
        self.clear_btn.pack(side=tk.LEFT)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Label de estado
        self.status_label = ttk.Label(main_frame, text="Listo")
        self.status_label.grid(row=7, column=0, columnspan=3, pady=(5, 0))
        
        self.selected_file = None
        self.total_pages = 0
    
    def fix_text_spacing(self, text):
        """Corregir problemas de espaciado en texto extraído"""
        # Patrón para detectar palabras concatenadas (minúscula seguida de mayúscula)
        text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
        
        # Patrón para separar números de letras
        text = re.sub(r'([0-9])([A-Za-z])', r'\1 \2', text)
        text = re.sub(r'([A-Za-z])([0-9])', r'\1 \2', text)
        
        # Corregir espacios múltiples
        text = re.sub(r'\s+', ' ', text)
        
        # Mejorar separación de párrafos
        text = re.sub(r'([.!?])\s*([A-Z])', r'\1\n\n\2', text)
        
        # Correcciones específicas comunes
        corrections = {
            ' - ': '-',  # Guiones mal espaciados
            'by - laws': 'by-laws',
            'share - ': 'share-',
            'quota - ': 'quota-',
        }
        
        for wrong, correct in corrections.items():
            text = text.replace(wrong, correct)
        
        return text
    
    def improve_mathematical_formulas(self, text):
        """Mejorar el reconocimiento de fórmulas matemáticas"""
        if not self.improve_formulas.get():
            return text
        
        # Patrones comunes de fórmulas mal reconocidas
        improvements = [
            # Fracciones simples
            (r'(\w+)\s*×\s*\(\s*(\w+)\s*-\s*(\w+)\s*\)\s*r\s*=\s*(\w+)\s*\+\s*(\w+)', 
             r'r = \1 × (\2 - \3) / (\4 + \5)'),
            
            # Patrón específico de tu fórmula
            (r'N\s*×\s*\(\s*B\s*-\s*C\s*\)\s*r\s*=\s*N\s*\+\s*V', 
             r'r = N × (B - C) / (N + V)'),
            
            # Separar variables de definiciones
            (r'([A-Z]):\s*([^A-Z\n]+?)(?=[A-Z]:|\n\n|\Z)', 
             r'\1: \2\n'),
            
            # Mejorar espaciado en fórmulas
            (r'(\w)\s*=\s*(\w)', r'\1 = \2'),
            (r'(\w)\s*\+\s*(\w)', r'\1 + \2'),
            (r'(\w)\s*-\s*(\w)', r'\1 - \2'),
            (r'(\w)\s*×\s*(\w)', r'\1 × \2'),
            (r'(\w)\s*/\s*(\w)', r'\1 / \2'),
            
            # Corregir paréntesis mal espaciados
            (r'\(\s*([^)]+)\s*\)', r'(\1)'),
            
            # Mejorar formato de definiciones de variables
            (r'([A-Z])\s*:\s*([^A-Z\n]+)', r'\1: \2'),
        ]
        
        improved_text = text
        for pattern, replacement in improvements:
            improved_text = re.sub(pattern, replacement, improved_text, flags=re.IGNORECASE | re.MULTILINE)
        
        return improved_text
    
    def post_process_ocr_text(self, text):
        """Post-procesamiento específico para texto OCR"""
        # Correcciones comunes de OCR
        corrections = {
            'CINs': 'CINS',
            'ICATI CADE': 'ICADE',
            'Legal Framework of Companies': 'Legal Framework of Companies',
            '•': '■',  # Cambiar bullets mal reconocidos
        }
        
        processed_text = text
        for wrong, correct in corrections.items():
            processed_text = processed_text.replace(wrong, correct)
        
        # Mejorar formato de fórmulas si está activado
        if self.improve_formulas.get():
            processed_text = self.improve_mathematical_formulas(processed_text)
        
        return processed_text
    
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
            self.file_label.config(text=f"Archivo: {filename}")
            self.extract_btn.config(state="normal")
            self.status_label.config(text=f"Archivo seleccionado: {filename}")
            
            # Obtener y mostrar total de páginas
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
                
                # Si no se especifican páginas, usar todas
                if selected_pages is None:
                    selected_pages = list(range(1, total_pages + 1))
                
                for page_num in selected_pages:
                    if 1 <= page_num <= total_pages:
                        page = pdf_reader.pages[page_num - 1]  # PyPDF2 usa índice 0
                        page_text = page.extract_text()
                        if page_text.strip():  # Solo agregar si hay texto
                            # Aplicar corrección de espaciado
                            page_text = self.fix_text_spacing(page_text)
                            text += f"\n--- Página {page_num} ---\n"
                            text += page_text + "\n"
        
        except Exception as e:
            raise Exception(f"Error al leer PDF: {str(e)}")
        
        return text
    
    def extract_text_with_ocr(self, file_path, selected_pages=None):
        """Extraer texto usando OCR Space API mejorado"""
        try:
            url = 'https://api.ocr.space/parse/image'

            # Si se especifican páginas, crear un PDF temporal solo con esas páginas
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
                
                # Obtener motor OCR seleccionado
                engine = self.ocr_engine.get().split()[0]
                
                data = {
                    'apikey': self.ocr_api_key,
                    'language': 'eng',  # Cambiar a inglés para mejor reconocimiento de fórmulas
                    'detectOrientation': 'true',
                    'scale': 'true',
                    'OCREngine': engine,
                    'filetype': 'PDF',
                    'isTable': 'true',  # Mejorar reconocimiento de estructuras
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
                            # Post-procesar el texto
                            page_text = self.post_process_ocr_text(page_text)
                            text += f"\n--- Página {selected_pages[i] if selected_pages else i + 1} (OCR) ---\n"
                            text += page_text + "\n"

                return text

        except Exception as e:
            raise Exception(f"Error en OCR: {str(e)}")
        finally:
            # Eliminar el archivo temporal si se creó
            if 'temp_file' in locals() and temp_file is not None:
                try:
                    os.unlink(temp_file.name)
                except Exception:
                    pass

    def extract_text(self):
        """Extraer texto del PDF seleccionado"""
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

        # Mostrar barra de progreso
        self.progress.start(10)
        if selected_pages:
            self.status_label.config(text=f"Extrayendo texto de páginas: {', '.join(map(str, selected_pages[:5]))}{'...' if len(selected_pages) > 5 else ''}")
        else:
            self.status_label.config(text="Extrayendo texto de todas las páginas...")
        self.root.update()

        try:
            if self.use_ocr.get():
                # OCR mejorado con post-procesamiento
                self.extracted_text = self.extract_text_with_ocr(self.selected_file, selected_pages)
                if not self.extracted_text.strip():
                    # Si OCR no funciona, intentar método normal
                    self.extracted_text = self.extract_text_from_pdf(self.selected_file, selected_pages)
                    messagebox.showinfo("Info", "OCR no encontró texto. Se usó extracción normal.")
            else:
                # Extracción normal con páginas específicas
                self.extracted_text = self.extract_text_from_pdf(self.selected_file, selected_pages)
                # Aplicar mejoras también al texto normal si está activado
                if self.improve_formulas.get():
                    self.extracted_text = self.post_process_ocr_text(self.extracted_text)

            # Mostrar texto en el área de texto
            self.text_area.delete(1.0, tk.END)
            if self.extracted_text.strip():
                self.text_area.insert(1.0, self.extracted_text)
                self.save_btn.config(state="normal")
                self.save_word_btn.config(state="normal")

                # Mostrar estadísticas
                pages_processed = len(selected_pages) if selected_pages else self.total_pages
                method = "OCR mejorado" if self.use_ocr.get() else "extracción directa"
                self.status_label.config(text=f"Texto extraído con {method} de {pages_processed} página(s)")
            else:
                self.text_area.insert(1.0, "No se pudo extraer texto del PDF.\n\n"
                                           "Posibles soluciones:\n"
                                           "1. Activa la opción 'Usar OCR' si es un PDF escaneado\n"
                                           "2. Prueba con diferentes motores OCR\n"
                                           "3. Verifica que el PDF contenga texto seleccionable\n"
                                           "4. Verifica que las páginas especificadas sean correctas")
                self.status_label.config(text="No se encontró texto en las páginas seleccionadas")

        except Exception as e:
            messagebox.showerror("Error", f"Error al extraer texto:\n{str(e)}")
            self.status_label.config(text="Error en la extracción")

        finally:
            self.progress.stop()
    
    def save_text(self):
        """Guardar el texto extraído en un archivo TXT"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para guardar.")
            return
        
        # Sugerir nombre de archivo basado en el PDF original
        if self.selected_file:
            pdf_name = Path(self.selected_file).stem
            default_name = f"{pdf_name}_texto_mejorado.txt"
        else:
            default_name = "texto_extraido_mejorado.txt"
        
        file_path = filedialog.asksaveasfilename(
            title="Guardar texto como",
            defaultextension=".txt",
            initialfile=default_name,
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
    
    def save_word(self):
        """Guardar el texto extraído en un archivo Word"""
        if not self.extracted_text:
            messagebox.showwarning("Advertencia", "No hay texto para guardar.")
            return
        
        try:
            from docx import Document
        except ImportError:
            messagebox.showerror("Error", "La librería python-docx no está instalada.\n\n"
                                         "Instálala con:\npip install python-docx")
            return
        
        # Sugerir nombre de archivo basado en el PDF original
        if self.selected_file:
            pdf_name = Path(self.selected_file).stem
            default_name = f"{pdf_name}_texto_extraido.docx"
        else:
            default_name = "texto_extraido.docx"
        
        file_path = filedialog.asksaveasfilename(
            title="Guardar como Word",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Documentos Word", "*.docx"), ("Todos los archivos", "*.*")]
        )
        
        if file_path:
            try:
                # Crear documento Word
                doc = Document()
                
                # Agregar título
                if self.selected_file:
                    title = f"Texto extraído de: {os.path.basename(self.selected_file)}"
                    doc.add_heading(title, level=1)
                
                # Dividir texto en párrafos y procesar
                paragraphs = self.extracted_text.split('\n')
                
                for paragraph in paragraphs:
                    paragraph = paragraph.strip()
                    if paragraph:
                        if paragraph.startswith('---') and paragraph.endswith('---'):
                            # Es un encabezado de página
                            doc.add_heading(paragraph, level=2)
                        else:
                            # Es texto normal
                            doc.add_paragraph(paragraph)
                
                # Guardar documento
                doc.save(file_path)
                messagebox.showinfo("Éxito", f"Documento Word guardado en:\n{file_path}")
                self.status_label.config(text=f"Guardado Word: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Error al guardar documento Word:\n{str(e)}")
    
    def clear_text(self):
        """Limpiar el área de texto y resetear"""
        self.text_area.delete(1.0, tk.END)
        self.extracted_text = ""
        self.selected_file = None
        self.total_pages = 0
        self.file_label.config(text="Ningún archivo seleccionado")
        self.total_pages_label.config(text="")
        self.pages_entry.delete(0, tk.END)
        self.all_pages.set(True)
        self.pages_entry.config(state="disabled")
        self.extract_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.save_word_btn.config(state="disabled")
        self.status_label.config(text="Listo")

def main():
    # Verificar dependencias
    try:
        import PyPDF2
        import requests
    except ImportError as e:
        print("Error: Faltan dependencias requeridas.")
        print("Instala las dependencias con:")
        print("pip install PyPDF2 requests")
        return
    
    # Verificar dependencia opcional para Word
    try:
        import docx
    except ImportError:
        print("Nota: Para exportar a Word, instala python-docx:")
        print("pip install python-docx")
    
    root = tk.Tk()
    app = PDFTextExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()