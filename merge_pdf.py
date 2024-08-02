import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, Toplevel, ttk
import os
import re
import json
from PyPDF2 import PdfMerger
import win32com.client as win32
import zipfile
import shutil

CONFIG_FILE = "config.json"

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Unir PDFs y Enviar Email")

        self.din_numbers_var = StringVar()
        self.source_path_var = StringVar()
        self.email_var = StringVar()

        self.load_config()

        self.tab_control = ttk.Notebook(root)
        self.tab_merge = ttk.Frame(self.tab_control)
        self.tab_send = ttk.Frame(self.tab_control)
        self.tab_control.add(self.tab_merge, text='Unir PDFs')
        self.tab_control.add(self.tab_send, text='Enviar PDFs')
        self.tab_control.pack(expand=1, fill='both')

        self.setup_merge_tab()
        self.setup_send_tab()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_merge_tab(self):
        tk.Label(self.tab_merge, text="Números DIN (separados por coma):").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(self.tab_merge, textvariable=self.din_numbers_var, width=50).grid(row=0, column=1, padx=10, pady=10)

        tk.Label(self.tab_merge, text="Carpeta Origen:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        tk.Entry(self.tab_merge, textvariable=self.source_path_var, width=50).grid(row=1, column=1, padx=10, pady=10)
        tk.Button(self.tab_merge, text="Seleccionar", command=self.select_source_folder).grid(row=1, column=2, padx=10, pady=10)

        tk.Button(self.tab_merge, text="Unir PDF's", command=self.merge_pdfs).grid(row=2, column=0, columnspan=3, padx=10, pady=10)

    def setup_send_tab(self):
        tk.Label(self.tab_send, text="Email de destino:").grid(row=0, column=0, padx=10, pady=10)
        tk.Entry(self.tab_send, textvariable=self.email_var, width=50).grid(row=0, column=1, padx=10, pady=10)

        self.pdf_listbox = tk.Listbox(self.tab_send, selectmode=tk.MULTIPLE)
        self.pdf_listbox.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

        tk.Button(self.tab_send, text="Actualizar Lista", command=self.update_pdf_list).grid(row=2, column=1, padx=10, pady=10)
        tk.Button(self.tab_send, text="Seleccionar Todos", command=self.select_all_pdfs).grid(row=2, column=0, padx=10, pady=10)
        tk.Button(self.tab_send, text="Enviar PDF's", command=self.send_pdfs).grid(row=3, column=0, columnspan=2, padx=10, pady=10)

        self.tab_send.grid_rowconfigure(1, weight=1)
        self.tab_send.grid_columnconfigure(1, weight=1)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as file:
                config = json.load(file)
                self.source_path_var.set(config.get('source_path', ''))
                self.email_var.set(config.get('email', ''))
        else:
            self.source_path_var.set('')
            self.email_var.set('')

    def save_config(self):
        config = {
            'source_path': self.source_path_var.get(),
            'email': self.email_var.get()
        }
        with open(CONFIG_FILE, 'w') as file:
            json.dump(config, file)

    def select_source_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.source_path_var.set(folder_selected)

    def find_zip_file(self, folder_path, din_number):
        pattern = re.compile(f"carpeta_{din_number}")
        for file in os.listdir(folder_path):
            if file.endswith('.zip') and pattern.match(file):
                return os.path.join(folder_path, file)
        return None

    def merge_pdfs(self):
        din_numbers = [din.strip() for din in self.din_numbers_var.get().split(',')]
        source_path = self.source_path_var.get()

        if not din_numbers or not source_path:
            messagebox.showwarning("Error", "Por favor, complete todos los campos y seleccione las carpetas.")
            return

        destination_folder = os.path.join(source_path, "pdf_pendientes_envio")
        os.makedirs(destination_folder, exist_ok=True)

        for din_number in din_numbers:
            zip_file_path = self.find_zip_file(source_path, din_number)
            if not zip_file_path:
                messagebox.showwarning("Error", f"No se encontró el archivo ZIP para DIN: {din_number}")
                continue

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
                    messagebox.showwarning("Error", f"Faltan los siguientes archivos en el ZIP de {din_number}: {', '.join(missing_files)}")
                    continue

                output_file_name = f"{din_number}.pdf"
                output_file_path = os.path.join(destination_folder, output_file_name)
                with open(output_file_path, 'wb') as f_out:
                    pdf_merger.write(f_out)

            except Exception as e:
                messagebox.showerror("Error", f"Ha ocurrido un error con DIN {din_number}: {str(e)}")

        messagebox.showinfo("Éxito", "PDFs unidos correctamente. Revise la carpeta pdf_pendientes_envio.")
        self.save_config()

    def update_pdf_list(self):
        self.pdf_listbox.delete(0, tk.END)
        source_path = self.source_path_var.get()
        pending_folder = os.path.join(source_path, "pdf_pendientes_envio")
        if os.path.exists(pending_folder):
            for file in os.listdir(pending_folder):
                if file.endswith('.pdf'):
                    self.pdf_listbox.insert(tk.END, file)
        else:
            messagebox.showwarning("Error", "La carpeta pdf_pendientes_envio no existe.")

    def select_all_pdfs(self):
        self.pdf_listbox.select_set(0, tk.END)

    def send_pdfs(self):
        selected_pdfs = [self.pdf_listbox.get(i) for i in self.pdf_listbox.curselection()]
        source_path = self.source_path_var.get()
        pending_folder = os.path.join(source_path, "pdf_pendientes_envio")
        sent_folder = os.path.join(source_path, "pdf_enviados")
        os.makedirs(sent_folder, exist_ok=True)

        if not selected_pdfs:
            messagebox.showwarning("Error", "Seleccione al menos un PDF para enviar.")
            return

        email_address = self.email_var.get().strip()
        if not email_address:
            messagebox.showwarning("Error", "Por favor, ingrese un correo electrónico de destino.")
            return

        try:
            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = email_address
            mail.Subject = 'PDFs Unidos'
            mail.Body = 'Adjunto se encuentran los PDFs unidos.'

            for pdf in selected_pdfs:
                pdf_path = os.path.join(pending_folder, pdf)
                mail.Attachments.Add(pdf_path)
                shutil.move(pdf_path, os.path.join(sent_folder, pdf))

            mail.Send()
            messagebox.showinfo("Éxito", "Correo enviado correctamente")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo enviar el correo: {str(e)}")

        self.update_pdf_list()

    def on_closing(self):
        self.save_config()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
