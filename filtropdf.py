import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import unicodedata
pdf_path = None

def procesar_pdf(pdf_path, excel_path):
    filas = []
    ignorados_path = excel_path.replace(".xlsx", "_ignorados.txt") #para los registros que no cumplen el formato
    ignorados = []

    regex = re.compile(r"""
        ^(\d{5})\s+                         # Legajo
        (.+?),\s+                          # Apellido(s)
        ([A-ZÁÉÍÓÚÑ\s]+?)\s+               # Nombre(s)
        (\d{7,8})\s+                       # Documento
        (.+)$                              # Resto (cargo + depto)
        """, re.VERBOSE)

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            lines = page.extract_text().split("\n")
            for line_num, line in enumerate(lines, start=1):
                
                #Ignorando lineas del pdf que no son relevantes
                if (
                    not line.strip()
                    or line.startswith("U.N.C.")
                    or line.startswith("Listado de")
                    or line.startswith("Legajo Agente")
                    or line.startswith("Reporte emitido")
                    or line.startswith("Facultad de")
                    or line.strip().lower().startswith("página")
                    or "Sistema de Gestión" in line
                ):
                    continue

                match = regex.match(line)
                if match and len(match.groups()) == 5:
                    try:
                        legajo = match.group(1)
                        apellido = match.group(2).strip().title()
                        nombres = match.group(3).strip().title()
                        documento = match.group(4)
                        detalle = match.group(5).strip().title()

                        notas = f"{legajo} // {detalle}"

                        fila = {
                            "cardnumber": documento,
                            "surname": apellido,
                            "firstname": nombres,
                            "borrowernotes": notas
                        }

                        filas.append(fila)
                    except Exception as e:
                        ignorados.append(f"[ERROR] Página {page_num}, línea {line_num}: {line} => {e}")
                else:
                    ignorados.append(f"[IGNORADO] Página {page_num}, línea {line_num}: {line}")

    if not filas:
        messagebox.showwarning("Sin datos", "No se encontraron datos válidos en el PDF.")
        return

    columnas = ["cardnumber", "surname", "firstname", "borrowernotes"]
    df = pd.DataFrame(filas, columns=columnas)

    messagebox.showinfo(
    "Cargar correos",
    "Seleccioná el Excel que contiene los correos electrónicos.\n\nDebe tener:\n"
    "- Una columna llamada 'nombre_completo' con el formato 'Apellido Nombre'\n"
    "- Una columna llamada 'email' con el correo correspondiente."
    )
    #De otro excel agregamos los mails de cada usuario
    archivo_mails = filedialog.askopenfilename(
        title="Seleccionar excel con los correos",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )

    if archivo_mails:
        try:
            
            estado.config(text="Filtrando correos...")
            root.update_idletasks()

            df_mails = pd.read_excel(archivo_mails)
            df_mails.columns = [c.lower() for c in df_mails.columns]

            # Asegurar columnas esperadas
            if "nombre_completo" in df_mails.columns and "email" in df_mails.columns:
                # Normalizar datos
                df["nombre_completo"] = (df["surname"].str.strip() + " " + df["firstname"].str.strip()).apply(normalizar)
                df_mails["nombre_completo"] = df_mails["nombre_completo"].apply(lambda x: normalizar(x.replace(",", "")))

                # Diccionario para búsquedas parciales
                dic_mails = dict(zip(df_mails["nombre_completo"], df_mails["email"]))

                def buscar_mail(nombre_pdf):
                    for nombre_mails, mail in dic_mails.items():
                        if nombre_pdf.startswith(nombre_mails):
                            return mail
                    return ""

                # Aplicar a todos los registros
                df["email"] = df["nombre_completo"].apply(buscar_mail)
                df.drop(columns=["nombre_completo"], inplace=True)
            else:
                messagebox.showwarning("Formato inválido", "El archivo de mails debe tener columnas 'nombre_completo' y 'email'.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo de mails:\n{e}")

    #Guardamos los registros ignorados en un txt
    if ignorados:
        with open(ignorados_path, "w", encoding="utf-8") as f:
            f.write("\n".join(ignorados))

    messagebox.showinfo("Éxito", f"Datos exportados a:\n{excel_path}\n\n"
                        f"{len(ignorados)} Registros ignorados guardados en:\n{ignorados_path}")
    
    df.to_excel(excel_path, index=False)
    estado.config(text="¡Exportación completa!")

def normalizar(texto):
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().lower()
    texto = unicodedata.normalize('NFD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    return texto

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
estado = tk.Label(root, text="", fg="blue")
estado.pack()

root.mainloop()