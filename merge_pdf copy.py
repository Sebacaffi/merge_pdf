import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, Toplevel
import os
import re
import json
from PyPDF2 import PdfMerger
import win32com.client as win32
import zipfile

CONFIG_FILE = "config.json"

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unir PDFs y Enviar Email")

        self.din_number_var = StringVar()
        self.source_path_var = StringVar()
        self.destination_path_var = StringVar()
        self.email_var = StringVar()

        self.load_config()

        tk.Label(root, text="Número DIN:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(root, textvariable=self.din_number_var, width=50).grid(row=0, column=1, padx=10, pady=10)

        tk.Button(root, text="Unir PDF's y Enviar Email", command=self.merge_pdfs_and_send_email).grid(row=2, column=0, columnspan=2, padx=10, pady=10)

        self.create_menu()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        config_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Configuración", menu=config_menu)
        config_menu.add_command(label="Rutas e Email", command=self.open_config_window)

    def open_config_window(self):
        config_window = Toplevel(self.root)
        config_window.title("Configuración de Carpetas")

        tk.Label(config_window, text="Carpeta Origen:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(config_window, textvariable=self.source_path_var, width=50).grid(row=0, column=1, padx=10, pady=10)
        tk.Button(config_window, text="Seleccionar", command=self.select_source_folder).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(config_window, text="Carpeta Destino:").grid(row=1, column=0, padx=10, pady=10)
        tk.Entry(config_window, textvariable=self.destination_path_var, width=50).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(config_window, text="Seleccionar", command=self.select_destination_folder).grid(row=1, column=2, padx=10, pady=10)

        tk.Label(config_window, text="Email de destino:").grid(row=2, column=0, padx=10, pady=10)
        tk.Entry(config_window, textvariable=self.email_var, width=50).grid(row=2, column=1, padx=10, pady=10)

        tk.Button(config_window, text="Guardar Configuración", command=lambda: self.save_config(config_window)).grid(row=3, column=0, columnspan=3, pady=10)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as file:
                config = json.load(file)
                self.source_path_var.set(config.get('source_path', ''))
                self.destination_path_var.set(config.get('destination_path', ''))
                self.email_var.set(config.get('email', ''))
        else:
            self.source_path_var.set('')
            self.destination_path_var.set('')
            self.email_var.set('')

    def save_config(self, window):
        config = {
            'source_path': self.source_path_var.get(),
            'destination_path': self.destination_path_var.get(),
            'email': self.email_var.get()
        }
        with open(CONFIG_FILE, 'w') as file:
            json.dump(config, file)
        messagebox.showinfo("Éxito", "Configuración guardada correctamente")
        window.destroy()

    def select_source_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.source_path_var.set(folder_selected)

    def select_destination_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.destination_path_var.set(folder_selected)

    def find_zip_file(self, folder_path, din_number):
        pattern = re.compile(f"carpeta_{din_number}")
        for file in os.listdir(folder_path):
            if file.endswith('.zip') and pattern.match(file):
                return os.path.join(folder_path, file)
        return None

    def merge_pdfs_and_send_email(self):
        din_number = self.din_number_var.get().strip()
        source_path = self.source_path_var.get()
        destination_path = self.destination_path_var.get()
        email_address = self.email_var.get().strip()

        if not din_number or not source_path or not destination_path or not email_address:
            messagebox.showwarning("Error", "Por favor, complete todos los campos y seleccione las carpetas.")
            return

        zip_file_path = self.find_zip_file(source_path, din_number)
        if not zip_file_path:
            messagebox.showwarning("Error", f"No se encontró el archivo ZIP para DIN: {din_number}")
            return

        required_files = ["NOTA DE COBRO", "DIN", "COMPROBANTE PAGO TESORERIA"]
        pdf_merger = PdfMerger()
        missing_files = []

        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                for req_file in required_files:
                    matching_files = [f for f in zip_ref.namelist() if req_file in f and f.endswith('.pdf')]
                    if not matching_files:
                        missing_files.append(req_file)
                    else:
                        with zip_ref.open(matching_files[0]) as pdf_file:
                            pdf_merger.append(pdf_file)

            if missing_files:
                messagebox.showwarning("Error", f"Faltan los siguientes archivos en el ZIP: {', '.join(missing_files)}")
                return

            output_file_name = f"{din_number}.pdf"
            output_file_path = os.path.join(destination_path, output_file_name)
            with open(output_file_path, 'wb') as f_out:
                pdf_merger.write(f_out)

            messagebox.showinfo("Éxito", "PDFs unidos correctamente")
            self.send_email([output_file_path], email_address)

        except Exception as e:
            messagebox.showerror("Error", f"Ha ocurrido un error: {str(e)}")

    def send_email(self, file_paths, to_email):
        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = to_email
            mail.Subject = 'PDFs Unidos'
            mail.Body = 'Adjunto se encuentran los PDFs unidos.'
            for file_path in file_paths:
                mail.Attachments.Add(file_path)
            mail.Send()
            messagebox.showinfo("Éxito", "Correo enviado correctamente")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo enviar el correo: {str(e)}")

    def on_closing(self):
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
