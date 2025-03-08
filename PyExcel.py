import polars as pl
import tkinter as tk
from tkinter import filedialog, messagebox

def seleccionar():
    archivo = filedialog.askopenfilename(
        filetypes=[
            ("Archivos CSV", "*.csv"),
            ("Archivos de Excel", "*.xls;*.xlsx"),
            ("Archivos JSON", "*.json"),
            ("Bases de datos SQLite", "*.db;*.sqlite"),
            ("Archivos Parquet", "*.parquet"),
            ("Archivos XML", "*.xml"),
            ("Archivos SQL", "*.sql"),
            ("Archivos de Access", "*.mdb;*.accdb"),
            ("Archivos Feather", "*.feather"),
            ("Archivos HDF5", "*.h5;*.hdf5"),
            ("Archivos Pickle", "*.pkl"),
            ("Todos los archivos", "*.*")
        ]
    )
    if archivo:
        messagebox.showinfo("Archivo seleccionado", f"Has seleccionado: {archivo}")
        convertir(archivo)
    else:
        messagebox.showwarning("No se seleccionó archivo", "No se seleccionó ningún archivo.")

def convertir(archivo):

    try:
        df = pl.read_csv(archivo)
        actualizar_texto()

        archivo_excel = archivo.replace(".csv", ".xlsx")
        df.write_excel(archivo_excel)
        

        messagebox.showinfo("Conversión Exitosa", f"Archivo guardado como:\n{archivo_excel}")
        actualizar_texto()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo convertir: {e}")
        actualizar_texto()

def actualizar_texto():
    texto_actual = etiqueta.cget("text")  # Obtener el texto actual de la etiqueta
    if texto_actual == "Selecciona un archivo para convertir a Excel":
        nuevo_texto = "Conversión en proceso..."
    else:
        nuevo_texto = "Selecciona un archivo para convertir a Excel"
    
    # Actualizar el texto de la etiqueta
    etiqueta.config(text=nuevo_texto)

def main():
    # Crear la ventana principal
    ventana = tk.Tk()
    ventana.title("PyExcel")
    ventana.geometry("700x400")
    ventana.configure(bg="#212121")  # Fondo de la ventana negro
    ventana.option_add("*foreground", "white")  # Color del texto blanco
    ventana.option_add("*background", "#212121")  # Color de fondo de widgets negro

    # Crear la etiqueta para mostrar el texto
    global etiqueta
    etiqueta = tk.Label(ventana, text="Selecciona un archivo para convertir a Excel", font=("Arial", 14), bg="#212121", fg="white")
    etiqueta.pack(pady=20)

    # Botón para seleccionar archivo
    boton_convertir = tk.Button(ventana, text="Seleccionar", command=seleccionar, font=("Arial", 12, "bold"), 
                                bg="#4CAF50", fg="white", relief="raised", padx=10, pady=5)
    boton_convertir.pack(pady=40)



    # Botón para cerrar la ventana
    boton_cerrar = tk.Button(ventana, text="Cerrar", command=ventana.quit, font=("Arial", 13, "bold"), fg="black", bg="#ED5353")
    boton_cerrar.pack(pady=20)

    ventana.mainloop()

if __name__ == "__main__":
    main()
