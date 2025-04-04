import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from generate_permissions_excel import create_permissions_excel
import os

class PermissionExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Excel de Permisos")
        self.root.geometry("800x600")
        
        # Estilo
        style = ttk.Style()
        style.configure('TButton', padding=5)
        style.configure('TLabel', padding=5)
        
        # Frame principal
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Instrucciones
        instructions = """Instrucciones:
1. Ingresa los objetos y sus permisos en el formato:

Nombre del Objeto
Checked/Not Checked    Checked/Not Checked    Checked/Not Checked    Checked/Not Checked    Checked/Not Checked    Checked/Not Checked    Checked/Not Checked

2. Usa 'Checked' para permisos activos y 'Not Checked' para inactivos
3. Separa los permisos con tabulaciones o espacios
4. Deja una línea en blanco entre cada objeto"""
        
        # Label de instrucciones
        ttk.Label(main_frame, text=instructions, justify=tk.LEFT).grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        # Área de texto
        self.text_area = scrolledtext.ScrolledText(main_frame, width=80, height=20)
        self.text_area.grid(row=1, column=0, columnspan=2, pady=10)
        
        # Botones
        ttk.Button(main_frame, text="Generar Excel", command=self.generate_excel).grid(row=2, column=0, pady=10)
        ttk.Button(main_frame, text="Limpiar", command=self.clear_text).grid(row=2, column=1, pady=10)
        
        # Ejemplo de datos
        self.example_data = """Accounts
Checked    Checked    Checked    Checked    Not Checked    Not Checked    Not Checked

Ideas
Checked    Checked    Not Checked    Not Checked        Not Checked"""
        
        # Botón para cargar ejemplo
        ttk.Button(main_frame, text="Cargar Ejemplo", command=self.load_example).grid(row=3, column=0, columnspan=2, pady=5)
        
        # Configurar el grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
    
    def generate_excel(self):
        try:
            data = self.text_area.get("1.0", tk.END)
            if not data.strip():
                messagebox.showerror("Error", "Por favor, ingresa algunos datos primero.")
                return
                
            create_permissions_excel(data)
            
            # Mostrar mensaje de éxito con la ubicación del archivo
            file_path = os.path.abspath("permisos.xlsx")
            messagebox.showinfo("Éxito", 
                              f"Excel generado exitosamente!\nUbicación: {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar el Excel: {str(e)}")
    
    def clear_text(self):
        self.text_area.delete("1.0", tk.END)
    
    def load_example(self):
        self.text_area.delete("1.0", tk.END)
        self.text_area.insert("1.0", self.example_data)

def main():
    root = tk.Tk()
    app = PermissionExcelApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 