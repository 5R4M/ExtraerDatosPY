import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

def extract_and_transfer():
    try:
        # Seleccionar archivo de origen
        origen = filedialog.askopenfilename(title="Seleccionar archivo de origen", filetypes=[("Excel files", "*.xlsx")])
        if not origen:
            return
        
        # Leer el archivo de Excel original
        df = pd.read_excel(origen, sheet_name='Hoja1')

        # Pedir al usuario que ingrese los títulos de las columnas
        columnas = simpledialog.askstring("Entrada", "Ingrese los títulos de las columnas separados por comas (ej. cantidad, descripcion):")
        if not columnas:
            return
        columnas = [col.strip() for col in columnas.split(',')]

        # Verificar si las columnas existen en el DataFrame
        for col in columnas:
            if col not in df.columns:
                messagebox.showerror("Error", f"La columna '{col}' no existe en el archivo.")
                return

        # Extraer datos específicos según los títulos de las columnas
        datos_extraidos = df[columnas].dropna(how='all')

        # Crear una nueva columna combinada
        datos_extraidos['Combinado'] = datos_extraidos.apply(lambda row: ' '.join([f"{col}: {row[col]}" for col in columnas]), axis=1)

        # Crear un nuevo archivo de Excel
        wb = Workbook()
        ws = wb.active

        # Escribir los datos extraídos en el nuevo archivo
        for r_idx, row in enumerate(datos_extraidos.values, 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                # Aplicar formato de borde
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                cell.border = thin_border
                # Aplicar tipo y tamaño de letra
                cell.font = Font(name='Arial', size=12)

        # Guardar el nuevo archivo de Excel
        destino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if destino:
            wb.save(destino)
            messagebox.showinfo("Éxito", "Datos transferidos exitosamente")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Crear la ventana principal
root = tk.Tk()
root.title("Transferencia de Datos de Excel")

# Crear el botón para iniciar la transferencia
btn_transferir = tk.Button(root, text="Transferir Datos", command=extract_and_transfer)
btn_transferir.pack(pady=20)

# Ejecutar la aplicación
root.mainloop()
