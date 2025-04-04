import polars as pl
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import platform
import subprocess
from fpdf import FPDF
import pandas as pd
import unicodedata
import pdfplumber
import logging
from datetime import datetime

# Variable global para almacenar el último archivo seleccionado
ultimo_archivo = None
df_actual = None





def seleccionar():
    global ultimo_archivo, df_actual
    archivo = filedialog.askopenfilename(
        filetypes=[
            ("Archivos CSV", "*.csv"),
            ("Archivos de Excel", "*.xls;*.xlsx"),
            ("Archivos JSON", "*.json"),
            ("Archivos PDF", "*.pdf"),
            ("Todos ", "*.*")
        ]
    )

    if archivo:
        ultimo_archivo = archivo
        actualizar_texto("Cargando archivo...")
        try:
            # Leer el archivo según su extensión
            extension = archivo.lower().split('.')[-1]
            
            if extension == 'csv':
                df_actual = pl.read_csv(archivo)
            elif extension in ('xls', 'xlsx'):
                df_actual = pl.read_excel(archivo)
            elif extension == 'json':
                df_actual = pl.read_json(archivo)
            elif extension == 'pdf':
                excel_path = pdf_a_excel(archivo)
                messagebox.showinfo("Éxito", f"PDF convertido a:\n{excel_path}")
                
            actualizar_texto(f"Archivo cargado: {os.path.basename(archivo)}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo:\n{str(e)}")
            actualizar_texto("Error al cargar archivo")
    else:
        messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")


# Añadir al inicio de las funciones de conversión


def convertir_a_excel():
    if df_actual is None:
        messagebox.showwarning("Advertencia", "Primero selecciona un archivo")
        return
    
    if df_actual.is_empty():
        messagebox.showwarning("Advertencia", "El DataFrame está vacío")
        return

    if len(df_actual) > 10000:
        if not messagebox.askyesno("Confirmación", "El archivo es muy grande (>10k filas). ¿Continuar?"):
            return
    
    
    try:
        carpeta_resultados = os.path.join(os.getcwd(), "archivos_convertidos")
        os.makedirs(carpeta_resultados, exist_ok=True)
        
        nombre_base = os.path.splitext(os.path.basename(ultimo_archivo))[0]
        archivo_excel = os.path.join(carpeta_resultados, f"{nombre_base}.xlsx")
        
        df_actual.write_excel(archivo_excel)
        actualizar_texto(f"✓ Convertido a Excel\n{os.path.basename(archivo_excel)}", "#4CAF50")
        
    except Exception as e:
        messagebox.showerror("Error", f"Fallo al convertir a Excel:\n{str(e)}")

def convertir_a_pdf():
    if df_actual is None:
        messagebox.showwarning("Advertencia", "Primero selecciona un archivo")
        return
    
    try:
        carpeta_resultados = os.path.join(os.getcwd(), "archivos_convertidos")
        os.makedirs(carpeta_resultados, exist_ok=True)
        
        nombre_base = os.path.splitext(os.path.basename(ultimo_archivo))[0]
        archivo_pdf = os.path.join(carpeta_resultados, f"{nombre_base}.pdf")
        
        pdf = FPDF()
        pdf.add_page()
        
        # Agregar fuente Unicode
        pdf.add_font("DejaVu", "", "fonts/DejaVuSans.ttf", uni=True)
        pdf.set_font("DejaVu", size=7)

        
        # Encabezado
        pdf.set_fill_color(79, 129, 189)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(200, 10, txt=f"Reporte: {nombre_base}", ln=1, align='C', fill=True)
        pdf.ln(5)
        
        # Configurar tabla
        # Calcular el ancho máximo entre encabezados y contenido
        col_widths = []
        for col in df_actual.columns:
            max_width = pdf.get_string_width(str(col))
            for val in df_actual[col].to_list():
                val_width = pdf.get_string_width(str(val))
                if val_width > max_width:
                    max_width = val_width
            col_widths.append(max(max_width + 6, 20))  # 20 es el mínimo ancho por columna

        
        # Encabezados
        pdf.set_fill_color(79, 129, 189)
        pdf.set_text_color(255, 255, 255)
        for col, width in zip(df_actual.columns, col_widths):
            pdf.cell(width, 10, str(col), border=1, fill=True)
        pdf.ln()
        
        # Datos
        pdf.set_fill_color(255, 255, 255)
        pdf.set_text_color(0, 0, 0)
        for row in df_actual.iter_rows():
            for value, width in zip(row, col_widths):
                pdf.cell(width, 10, str(value), border=1)
            pdf.ln()

        pdf.output(archivo_pdf)
        actualizar_texto(f"✓ Convertido a PDF\n{os.path.basename(archivo_pdf)}", "#4CAF50")
        
    except Exception as e:
        messagebox.showerror("Error", f"Fallo al convertir a PDF:\n{str(e)}")
        actualizar_texto("Error en conversión", "#F44336")


#actualiza el titulo del label
def actualizar_texto(mensaje, color="white"):
    etiqueta.config(text=mensaje, fg=color)



def abrir_carpeta():
    ruta_carpeta = os.path.join(os.getcwd(), "archivos_convertidos")
    try:
        os.makedirs(ruta_carpeta, exist_ok=True)
        sistema = platform.system()
        if sistema == "Windows":
            os.startfile(ruta_carpeta)
        elif sistema == "Darwin":
            subprocess.run(["open", ruta_carpeta])
        else:
            subprocess.run(["xdg-open", ruta_carpeta])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir carpeta:\n{str(e)}")

def main():
    global etiqueta
   
    ventana = tk.Tk()
    ventana.title("Conversor de Archivos")
    ventana.geometry("700x450")
    ventana.configure(bg="#212121")
    ventana.option_add("*foreground", "white")
    ventana.option_add("*background", "#212121")

    # Etiqueta principal
    etiqueta = tk.Label(ventana, text="Selecciona un archivo para convertir", 
                       font=("Arial", 14), bg="#212121", fg="white")
    etiqueta.pack(pady=20)

    # Botón para seleccionar arcchivo
    boton_seleccionar = tk.Button(ventana, text="Seleccionar Archivo", command=seleccionar,
                                 font=("Arial", 12), bg="#4CAF50", fg="white", padx=15, pady=5)
    boton_seleccionar.pack(pady=10)

    # Frame para botones de conversión
    frame_botones = tk.Frame(ventana, bg="#212121")
    frame_botones.pack(pady=10)
    

    # Botón para convertir a Excel
    boton_excel = tk.Button(frame_botones, text="Convertir a xlsx", command=convertir_a_excel,
                           font=("Arial", 12), bg="#2196F3", fg="white", padx=15, pady=5)
    boton_excel.pack(side=tk.LEFT, padx=10)

    # Botón para convertir a PDF
    boton_pdf = tk.Button(frame_botones, text="Convertir a PDF", command=convertir_a_pdf,
                         font=("Arial", 12), bg="#FF5722", fg="white", padx=15, pady=5)
    boton_pdf.pack(side=tk.LEFT, padx=10)

    # Botón para abrir carpeta
    boton_abrir = tk.Button(ventana, text="Abrir Carpeta de Resultados", command=abrir_carpeta,
                           font=("Arial", 12), bg="#9C27B0", fg="white", padx=15, pady=5)
    boton_abrir.pack(pady=10)

    # Botón para cerrar
    boton_cerrar = tk.Button(ventana, text="Cerrar", command=ventana.quit,
                            font=("Arial", 12), bg="#F44336", fg="white", padx=15, pady=5)
    boton_cerrar.pack(pady=10)

    ventana.mainloop()

if __name__ == "__main__":
    main()