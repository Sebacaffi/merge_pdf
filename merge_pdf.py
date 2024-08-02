import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, StringVar, PhotoImage
import os
import zipfile
import re
import json
from PyPDF2 import PdfMerger

CONFIG_FILE = ("config.json")


class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unir PDFs")

        self.din_numbers_var = StringVar()
        self.source_path_var = StringVar()
        self.destination_path_var = StringVar()

        self.load_config()

        tk.Label(root, text="Números DIN:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(root, textvariable=self.din_numbers_var, width=50).grid(row=0, column=1, padx=10, pady=10)

        tk.Button(root, text="Unir PDF's", command=self.merge_pdfs).grid(row=3, column=1, padx=10, pady=10)

        # Create menu bar
        menubar = tk.Menu(root)
        root.config(menu=menubar)

        # Create config menu
        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Configuración", menu=config_menu)
        config_menu.add_command(label="Editar rutas", command=self.open_config_window)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def select_source_folder(self):
        folder_selected = filedialog.askdirectory()
        self.source_path_var.set(folder_selected)

    def select_destination_folder(self):
        folder_selected = filedialog.askdirectory()
        self.destination_path_var.set(folder_selected)

    def find_zip_file(self, folder_path, din_number):
        pattern = re.compile(f"carpeta_{din_number}")
        for file in os.listdir(folder_path):
            if file.endswith('.zip') and pattern.match(file):
                return os.path.join(folder_path, file)
        return None

    def merge_pdfs(self):
        din_numbers = self.din_numbers_var.get().split(',')
        source_path = self.source_path_var.get()
        destination_path = self.destination_path_var.get()

        if not din_numbers or not source_path or not destination_path:
            messagebox.showerror("Error", "Todos los campos son obligatorios")
            return

        missing_zips = []

        for din_number in din_numbers:
            din_number = din_number.strip()
            if not din_number:
                continue

            zip_file_path = self.find_zip_file(source_path, din_number)
            if not zip_file_path:
                missing_zips.append(din_number)
                continue

            pdf_merger = PdfMerger()
            ordered_files = ["NOTA DE COBRO", "DIN", "COMPROBANTE PAGO TESORERIA"]
            temp_extract_path = os.path.join(source_path, "temp_extract")

            if not os.path.exists(temp_extract_path):
                os.makedirs(temp_extract_path)

            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(temp_extract_path)

            missing_files = []
            for name in ordered_files:
                found = False
                for file in os.listdir(temp_extract_path):
                    if file.startswith(name) and file.endswith('.pdf'):
                        pdf_merger.append(os.path.join(temp_extract_path, file))
                        found = True
                        break
                if not found:
                    missing_files.append(name)

            if missing_files:
                messagebox.showerror("Error", f"No se encontraron los archivos {', '.join(missing_files)} en el .zip para el número DIN {din_number}")
                continue

            output_path = os.path.join(destination_path, din_number + ".pdf")
            pdf_merger.write(output_path)
            pdf_merger.close()

            for file in os.listdir(temp_extract_path):
                os.remove(os.path.join(temp_extract_path, file))
            os.rmdir(temp_extract_path)

        if missing_zips:
            messagebox.showerror("Error", f"No se encontraron archivos .zip para los números DIN: {', '.join(missing_zips)}")
        else:
            messagebox.showinfo("Éxito", "Los archivos PDF se han unido correctamente")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as config_file:
                config = json.load(config_file)
                self.source_path_var.set(config.get('source_path', ''))
                self.destination_path_var.set(config.get('destination_path', ''))

    def save_config(self):
        config = {
            'source_path': self.source_path_var.get(),
            'destination_path': self.destination_path_var.get()
        }
        with open(CONFIG_FILE, 'w') as config_file:
            json.dump(config, config_file)
        messagebox.showinfo("Éxito", "Configuración guardada correctamente")

    def open_config_window(self):
        config_window = Toplevel(self.root)
        config_window.title("Editar Rutas")

        tk.Label(config_window, text="Ruta de los archivos:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        tk.Entry(config_window, textvariable=self.source_path_var, width=80).grid(row=0, column=1, padx=10, pady=10)
        tk.Button(config_window, text="Seleccionar", command=self.select_source_folder).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(config_window, text="Ruta de destino:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        tk.Entry(config_window, textvariable=self.destination_path_var, width=80).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(config_window, text="Seleccionar", command=self.select_destination_folder).grid(row=1, column=2, padx=10, pady=10)

        tk.Button(config_window, text="Guardar", command=self.save_config).grid(row=2, column=1, padx=10, pady=10)

    def on_closing(self):
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
