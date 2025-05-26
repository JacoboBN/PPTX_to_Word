import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import PyPDF2
import requests
import base64
import os
from pathlib import Path

class PDFTextExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Extractor de Texto PDF")
        self.root.geometry("800x600")
        
        # Variable para almacenar el texto extraído
        self.extracted_text = ""
        
        # API Key para OCR Space (opcional - puedes usar la gratuita)
        self.ocr_api_key = "K83967071688957"  # API key gratuita básica
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar expansión
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Extractor de Texto PDF", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Botón para seleccionar archivo
        self.select_btn = ttk.Button(main_frame, text="Seleccionar PDF", 
                                    command=self.select_file)
        self.select_btn.grid(row=1, column=0, padx=(0, 10), sticky=tk.W)
        
        # Label para mostrar archivo seleccionado
        self.file_label = ttk.Label(main_frame, text="Ningún archivo seleccionado")
        self.file_label.grid(row=1, column=1, sticky=(tk.W, tk.E))
        
        # Checkbox para OCR
        self.use_ocr = tk.BooleanVar()
        self.ocr_checkbox = ttk.Checkbutton(main_frame, text="Usar OCR (para PDFs escaneados)", 
                                           variable=self.use_ocr)
        self.ocr_checkbox.grid(row=1, column=2, padx=(10, 0))
        
        # Área de texto para mostrar el contenido
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(20, 0))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.text_area = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, 
                                                  width=70, height=20)
        self.text_area.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Frame para botones inferiores
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(20, 0))
        
        # Botón para extraer texto
        self.extract_btn = ttk.Button(button_frame, text="Extraer Texto", 
                                     command=self.extract_text, state="disabled")
        self.extract_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón para guardar texto
        self.save_btn = ttk.Button(button_frame, text="Guardar como TXT", 
                                  command=self.save_text, state="disabled")
        self.save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón para limpiar
        self.clear_btn = ttk.Button(button_frame, text="Limpiar", 
                                   command=self.clear_text)
        self.clear_btn.pack(side=tk.LEFT)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Label de estado
        self.status_label = ttk.Label(main_frame, text="Listo")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=(5, 0))
        
        self.selected_file = None
    
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
    
    def extract_text_from_pdf(self, file_path):
        """Extraer texto usando PyPDF2"""
        text = ""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    page_text = page.extract_text()
                    if page_text.strip():  # Solo agregar si hay texto
                        text += f"\n--- Página {page_num + 1} ---\n"
                        text += page_text + "\n"
        
        except Exception as e:
            raise Exception(f"Error al leer PDF: {str(e)}")
        
        return text
    
    def extract_text_with_ocr(self, file_path):
        """Extraer texto usando OCR Space API"""
        try:
            url = 'https://api.ocr.space/parse/image'
            
            with open(file_path, 'rb') as file:
                files = {'file': file}
                data = {
                    'apikey': self.ocr_api_key,
                    'language': 'spa',  # Español
                    'detectOrientation': 'true',
                    'scale': 'true',
                    'OCREngine': '2',
                    'filetype': 'PDF'
                }
                
                response = requests.post(url, files=files, data=data)
                result = response.json()
                
                if result.get('IsErroredOnProcessing'):
                    raise Exception(f"Error en OCR: {result.get('ErrorMessage', 'Error desconocido')}")
                
                text = ""
                if 'ParsedResults' in result:
                    for i, page_result in enumerate(result['ParsedResults']):
                        if 'ParsedText' in page_result:
                            text += f"\n--- Página {i + 1} (OCR) ---\n"
                            text += page_result['ParsedText'] + "\n"
                
                return text
        
        except Exception as e:
            raise Exception(f"Error en OCR: {str(e)}")
    
    def extract_text(self):
        """Extraer texto del PDF seleccionado"""
        if not self.selected_file:
            messagebox.showwarning("Advertencia", "Por favor selecciona un archivo PDF primero.")
            return
        
        # Mostrar barra de progreso
        self.progress.start(10)
        self.status_label.config(text="Extrayendo texto...")
        self.root.update()
        
        try:
            if self.use_ocr.get():
                # Usar OCR
                self.extracted_text = self.extract_text_with_ocr(self.selected_file)
                if not self.extracted_text.strip():
                    # Si OCR no funciona, intentar método normal
                    self.extracted_text = self.extract_text_from_pdf(self.selected_file)
                    messagebox.showinfo("Info", "OCR no encontró texto. Se usó extracción normal.")
            else:
                # Extracción normal
                self.extracted_text = self.extract_text_from_pdf(self.selected_file)
            
            # Mostrar texto en el área de texto
            self.text_area.delete(1.0, tk.END)
            if self.extracted_text.strip():
                self.text_area.insert(1.0, self.extracted_text)
                self.save_btn.config(state="normal")
                self.status_label.config(text="Texto extraído exitosamente")
            else:
                self.text_area.insert(1.0, "No se pudo extraer texto del PDF.\n\n"
                                           "Posibles soluciones:\n"
                                           "1. Activa la opción 'Usar OCR' si es un PDF escaneado\n"
                                           "2. Verifica que el PDF contenga texto seleccionable")
                self.status_label.config(text="No se encontró texto en el PDF")
        
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
            default_name = f"{pdf_name}_texto.txt"
        else:

            default_name = "texto_extraido.txt"
        
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
    
    def clear_text(self):
        """Limpiar el área de texto y resetear"""
        self.text_area.delete(1.0, tk.END)
        self.extracted_text = ""
        self.selected_file = None
        self.file_label.config(text="Ningún archivo seleccionado")
        self.extract_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
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
    
    root = tk.Tk()
    app = PDFTextExtractor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
