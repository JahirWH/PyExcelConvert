import polars as pl
import tkinter as tk
from tkinter import filedialog, messagebox
import os 
import subprocess
import platform
from fpdf import FPDF
import pandas as pd
import pdfplumber

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
            ("Archivos PDF", "*.pdf")
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

def convertir_a_excel():
    if df_actual is None:
        messagebox.showwarning("Advertencia", "Primero selecciona un archivo")
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
        actualizar_texto("Error en conversión", "#F44336")

def convertir_a_pdf():
    if df_actual is None:
        messagebox.showwarning("Advertencia", "Primero selecciona un archivo")
        return
    
    try:
        carpeta_resultados = os.path.join(os.getcwd(), "archivos_convertidos")
        os.makedirs(carpeta_resultados, exist_ok=True)
        
        nombre_base = os.path.splitext(os.path.basename(ultimo_archivo))[0]
        archivo_pdf = os.path.join(carpeta_resultados, f"{nombre_base}.pdf")
        
        # Convertir a PDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)
        
        # Encabezado
        pdf.cell(200, 10, txt=f"Reporte: {nombre_base}", ln=1, align='C')
        pdf.ln(5)
        
        # Convertir a pandas para facilidad
        df_pandas = df_actual.to_pandas()
        
        # Crear tabla
        with pdf.table() as table:
            # Encabezados
            headers = table.row()
            for col in df_pandas.columns:
                headers.cell(str(col))
            
            # Datos
            for row in df_pandas.itertuples(index=False):
                row_cells = table.row()
                for item in row:
                    row_cells.cell(str(item))
        
        pdf.output(archivo_pdf)
        actualizar_texto(f"✓ Convertido a PDF\n{os.path.basename(archivo_pdf)}", "#4CAF50")
        
    except Exception as e:
        messagebox.showerror("Error", f"Fallo al convertir a PDF:\n{str(e)}")
        actualizar_texto("Error en conversión", "#F44336")


def pdf_a_excel(pdf_path):
    try:
        # Crear carpeta si no existe
        os.makedirs("pdf_convertidos", exist_ok=True)
        
        excel_path = os.path.join("pdf_convertidos", f"{os.path.splitext(os.path.basename(pdf_path))[0]}.xlsx")
        
        with pdfplumber.open(pdf_path) as pdf:
            with pd.ExcelWriter(excel_path) as writer:
                for i, page in enumerate(pdf.pages):
                    # Extraer todas las tablas de la página
                    tables = page.extract_tables()
                    
                    for j, table in enumerate(tables):
                        df = pd.DataFrame(table[1:], columns=table[0])
                        df.to_excel(writer, sheet_name=f"Pag_{i+1}_Tabla_{j+1}", index=False)
        
        return excel_path
    except Exception as e:
        raise Exception(f"Error en conversión: {str(e)}")


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