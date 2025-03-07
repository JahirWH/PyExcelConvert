import polars as pl
import tkinter as tk
from tkinter import filedialog, messagebox

def convertir_csv_a_excel():
    archivo_csv = filedialog.askopenfilename(filetypes=[("Archivos CSV", "*.csv")])
    
    if not archivo_csv:
        return  # Si el usuario no selecciona un archivo, salir

    try:
        # Cargar CSV con Polars
        df = pl.read_csv(archivo_csv)

        # Guardar como Excel
        archivo_excel = archivo_csv.replace(".csv", ".xlsx")
        df.write_excel(archivo_excel)

        messagebox.showinfo("Conversión Exitosa", f"Archivo guardado como:\n{archivo_excel}")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo convertir: {e}")

def main():
    ventana = tk.Tk()
    ventana.title("CSV a Excel - PyExcel")
    ventana.geometry("400x200")

    etiqueta = tk.Label(ventana, text="Selecciona un archivo CSV para convertir a Excel", font=("Arial", 12))
    etiqueta.pack(pady=20)

    boton_convertir = tk.Button(ventana, text="Seleccionar CSV", command=convertir_csv_a_excel)
    boton_convertir.pack(pady=10)

    boton_cerrar = tk.Button(ventana, text="Cerrar", command=ventana.quit)
    boton_cerrar.pack(pady=10)

    ventana.mainloop()

if __name__ == "__main__":
    main()
