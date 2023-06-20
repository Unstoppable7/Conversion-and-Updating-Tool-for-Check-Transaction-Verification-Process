import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import sys
import threading

# Crear una clase personalizada para redirigir la salida a la caja de texto
class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, msg):
        self.text_widget.insert(tk.END, msg)
        self.text_widget.see(tk.END)

    def flush(self):
        pass

def abrir_quickbooks_report():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo central", filetypes=[("Archivos XLSX", "*.xlsx")])
    if archivo:
        entrada_quickbooks_report.delete(0, tk.END)
        entrada_quickbooks_report.insert(0, archivo)

def clean(clean_inputs = True):
    caja_texto.delete(1.0, tk.END)
    if clean_inputs:
        entrada_quickbooks_report.delete(0, tk.END)

    #Ocultamos barra de progreso
    toggle_progress_bar(False)

def toggle_progress_bar(band):
    if band:
        #Barra de progreso
        barra_progreso_label.place(relx=0.5, rely=0.5, anchor="center")
        barra_progreso.place(relx=0.5, rely=0.55, anchor="center")
    else:   
        # Actualizar la barra de progreso
        barra_progreso["value"] = 0
        barra_progreso_label.config(text="")
        barra_progreso_label.place_forget()
        barra_progreso.place_forget()
    
    ventana.update_idletasks()

def show_info_Quickbooks_report():
     messagebox.showinfo("Info", "Select the report files extracted from Quickbooks with check transactions\n\nEach file must represent an account, which will be searched in the TD Bank transactions report\n\nThis version uses the report extracted from Quickbooks Reports - Memorized Reports - Check Positive Pay")

def ejecutar_hilo():
    # Crear un objeto Thread y pasarle la función procesar como objetivo
    #hilo_procesar = threading.Thread(target=procesar)

    # Iniciar la ejecución del hilo
    #hilo_procesar.start()
    pass

# Crear la ventana principal
ventana = tk.Tk()

# Configurar el tamaño de la ventana
ventana.geometry("800x600")

# Configurar el mensaje de bienvenida
mensaje_bienvenida = tk.Label(ventana, text="Formatting for TD Bank Transaction Verification", font=("Arial", 24))

subtitulo_quickbooks_report = tk.Label(ventana, text="Quickbooks Transactions Report", font=("Arial", 18))
entrada_quickbooks_report = tk.Entry(ventana, width=40)
boton_quickbooks_report = tk.Button(ventana, text="Select file", command=abrir_quickbooks_report)
boton_info_quickbooks_report = tk.Button(ventana, text="Info", command=show_info_Quickbooks_report)

# Configurar el botón de procesamiento
boton_procesar = tk.Button(ventana, text="Start Process", font=("Arial", 22), command=ejecutar_hilo)
# boton_procesar.grid(row=4, column=0, columnspan=2, pady=20)

# Configurar la caja de texto
caja_texto = tk.Text(ventana, height=10, width=50)
# caja_texto.grid(row=5, column=0, columnspan=2, padx=20)

# Configurar el botón "Clean"
boton_clean = tk.Button(ventana, text="Clean", font=("Arial", 14), command=clean)
# boton_clean.grid(row=6, column=0, columnspan=2, pady=10)

# Crear una barra de progreso
barra_progreso = ttk.Progressbar(ventana, mode="determinate", length=300)
barra_progreso_label = tk.Label(ventana, text="", font=("Arial", 12))

#Place
mensaje_bienvenida.place(relx=0.5, rely=0.1, anchor="center")

# Configurar los subtitulos y las cajas de texto de archivo central
subtitulo_quickbooks_report.place(relx=0.5, rely=0.2, anchor="center")
entrada_quickbooks_report.place(relx=0.5, rely=0.25, anchor="center")
boton_quickbooks_report.place(relx=0.5, rely=0.3, anchor="center")
boton_info_quickbooks_report.place(relx=0.68, rely=0.25, anchor="center")

# Configurar el botón de procesamiento
boton_procesar.place(relx=0.5, rely=0.4, anchor="center")

# Configurar la caja de texto
caja_texto.place(relx=0.5, rely=0.742, anchor="center")

# Configurar el botón "Clean"
boton_clean.place(relx=0.5, rely=0.94, anchor="center")

# Redirigir la salida estándar y la salida de error a la caja de texto
sys.stdout = TextRedirector(caja_texto)
sys.stderr = TextRedirector(caja_texto)

# # Establecer el icono de la ventana
# icono = 'APDC LOGO.ico'  # Reemplaza con la ruta completa del archivo de icono
# ventana.iconbitmap(icono)
ventana.title("APDC Check verification process")

# Ejecutar el bucle principal de la interfaz gráfica
ventana.mainloop()