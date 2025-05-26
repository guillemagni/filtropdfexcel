import pdfplumber
import pandas as pd
import re

# path del PDF
pdf_path = "path.pdf"

filas = []

# Regex para capturar los campos
regex = re.compile(
    """
        Expresión regular: 
        Legajo: (\d{5})
        Apellido y nombre: ([A-ZÁÉÍÓÚÑ]+,\s+[A-ZÁÉÍÓÚÑ\s]+?)
        Documento: (\d{7,8})
        Detalle: ([A-Z\s\(\)-]+?\d*)
        Detalle 2: ([A-ZÁÉÍÓÚÑ\s\.]+?)
    """
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

# Crear DataFrame con columnas ordenadas
columnas = [
    "borrowernumber", "cardnumber", "surname", "firstname",
    "promfyb", "cargo", "department"
]

df = pd.DataFrame(filas, columns=columnas)

# Exportar a Excel
df.to_excel("archivo_filtrado.xlsx", index=False)
print("Archivo nombre.xlsx generado correctamente.")
