import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
pdf_path = None

def procesar_pdf(pdf_path, excel_path):
    filas = []

    # Regex para capturar los campos
    regex = re.compile(
        r"^(\d{5})\s+([A-ZÁÉÍÓÚÑ]+),\s+([A-ZÁÉÍÓÚÑ\s]+?)\s+(\d{7,8})\s+([A-Z\s\(\)\-]+?)\s+([A-Z\s\(\)\-]+?\d*)\s+([A-ZÁÉÍÓÚÑ\s\.\/]+?)\s+TITULAR"
    )

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split("\n")
            for line in lines:
                match = regex.match(line)
                if match:
                    legajo = match.group(1)
                    apellido = match.group(2).title()
                    nombres = match.group(3).title()
                    documento = match.group(4)
                    cargo_promfyb = match.group(5).strip().title()
                    cargo = match.group(6).strip().title()
                    departamento = match.group(7).strip().title()

                    fila = {
                        "borrowernumber": legajo,
                        "cardnumber": documento,
                        "surname": apellido,
                        "firstname": nombres,
                        "promfyb": cargo_promfyb,
                        "cargo": cargo,
                        "department": departamento
                    }

                    filas.append(fila)

    columnas = [
        "borrowernumber", "cardnumber", "surname", "firstname",
        "promfyb", "cargo", "department"
    ]

    df = pd.DataFrame(filas, columns=columnas)
    df.to_excel(excel_path, index=False)
    messagebox.showinfo("Éxito", f"Datos exportados a {excel_path} correctamente.")

def seleccionar_pdf():
    global pdf_path
    ruta = filedialog.askopenfilename(
        title="Seleccionar PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if ruta:
        pdf_path = ruta
        menu.entryconfig("Guardar Excel", state="normal")
        pdf_cargado.config(text=f"PDF seleccionado:\n {pdf_path}")        

def guardar_excel():
    if not pdf_path:
        messagebox.showerror("Advertencia", "Debe seleccionar un archivo PDF primero.")
        return
    
    excel_path = filedialog.asksaveasfilename(
            title="Guardar como Excel",
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
    if excel_path:
        procesar_pdf(pdf_path, excel_path)

#Main
root = tk.Tk()
root.title("Filtro PDF a Excel")
root.geometry("600x200")

#Menu
menu_bar = tk.Menu(root)
menu = tk.Menu(menu_bar, tearoff=0)
menu.add_command(label="Seleccionar PDF a filtrar", command=seleccionar_pdf)
menu.add_command(label="Guardar Excel", command=guardar_excel, state="disabled")
menu.add_command(label="Salir", command=root.quit)
menu_bar.add_cascade(label="Opciones", menu=menu)
root.config(menu=menu_bar)

#Label para mostrar el PDF seleccionado
pdf_cargado = tk.Label(root, text="Ningún PDF seleccionado", wraplength=500)
pdf_cargado.pack(pady=40)

root.mainloop()