import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import sys
import threading
import pandas as pd


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

def procesar():
    #Limpiamos la caja de texto de informacion
    clean(clean_inputs=False)
    
    #Variables que manejan la barra de progreso
    #TODO
    format_report = 2

    ##################################### Manejo de entrada de archivos
    archivo_quickbooks_report = entrada_quickbooks_report.get()

    #Tratamos si los archivos estan vacios
    if archivo_quickbooks_report == "":
        messagebox.showerror("Error", "No file selected as Quickbooks report")
        return
    
    try:
        archivo_quickbooks_report_name = archivo_quickbooks_report.rsplit('/', 1)[1]
    except:
        messagebox.showerror("Error", "No file selected as Quickbooks report")
        return
    
    #Mostramos barra de progreso
    toggle_progress_bar(True)

    # Verificacion de formato quickbook report
    try:
        qb_report_df = pd.read_excel(archivo_quickbooks_report, header=None, sheet_name="Sheet1")           

    except:
        messagebox.showerror("Error", "Check Positive Pay Quickbooks report '" + archivo_quickbooks_report_name + "' is not in the correct format. \n\nIt is necessary that the sheet where the report is located has the name 'Sheet1'. \n\nPlease check the Quickbooks report and try again")
        clean(clean_inputs=False)
        return

    try:

        if not ((not pd.isna(qb_report_df.iat[1,1])) and str(qb_report_df.iat[0, 4]) == "Date" and str(qb_report_df.iat[0, 6]) == "Num" and str(qb_report_df.iat[0, 8]) == "Name" and str(qb_report_df.iat[0, 10]) == "Credit"):

            messagebox.showerror("Error", "Check Positive Pay Quickbooks report '" + archivo_quickbooks_report_name + "' is not in the correct format. \n\nPlease check the Quickbooks report and try again")
            
            #Ocultamos barra de progreso y limpiamos caja de texto
            clean(clean_inputs=False)
            return
    except:
        messagebox.showerror("Error", "It is not possible to access the information of the Quickbooks report '" + archivo_quickbooks_report_name + "' \n\nPlease check the Quickbooks report and try again")
        clean(False)
        return
    
    #Extraemos la columna 1 y 10, a su vez eliminamos las filas que tengan vacia la celda en la columna 1
    account_range_limit = qb_report_df.iloc[:,[1,10]].dropna(subset=qb_report_df.columns[1])
    #Extraemos las filas que tengan las celda de la columna 10 vacia
    # account_range_limit = account_range_limit[account_range_limit[10].isna()]

    ###############################Formato

    # Eliminar columnas vacias y extras
    columnas_a_eliminar_qb_report = [0,1,2,3,5,7,9]  # Índices de las columnas a eliminar
    columnas_extras_a_eliminar_qb_report = 4 #A partir del este numero se eliminaran
    qb_report_df = qb_report_df.drop(qb_report_df.columns[columnas_a_eliminar_qb_report], axis=1)
    if qb_report_df.shape[1] > 4:
        qb_report_df = qb_report_df.drop(qb_report_df.columns[columnas_extras_a_eliminar_qb_report:], axis=1)
    
    # print(len(qb_report_df.index) - 1)
    # print(account_range_limit.index[len(account_range_limit) - 1])

    #Comparamos la cantidad de filas que tiene qb_report, resto 1 para que sea acorde con la otra comparacion que es por indices (comienza desde 0), CON la cantidad de filas hasta la ultima fila que contiene datos que necesitamos. De esta manera verificamos si hay mas filas extra que se deben de eliminar
    if (len(qb_report_df.index) - 1) > (account_range_limit.index[len(account_range_limit) - 1]):

        # Eliminamos filas fuera de rango de los datos
        qb_report_df = qb_report_df[:account_range_limit.index[len(account_range_limit) - 1] + 1]

    # Eliminar filas que datos vacios, variable calculada anteriormente
    qb_report_df = qb_report_df.drop(account_range_limit.index, axis=0)

    #Eliminamos la primera fila con los titulos
    qb_report_df = qb_report_df.drop(0, axis=0)

    # Insertar una nueva columna llamada 'tmp' y asignar valores según los rangos, agrega el rango desde uno menos del inicio hasta uno menos del final
    for i in range(int(len(account_range_limit) / 2)):
        index = i + 1
        qb_report_df.loc[account_range_limit.index[i * 2]:(account_range_limit.index[(i * 2) + 1]), 'tmp'] = "439780" + str(account_range_limit.iat[i * 2, 0])[:4]

    #Eliminamos filas que no contengan numero de cheque
    indice_columna_numero_transaccion_qb_report = 1

    # Eliminar las filas con valores nulos en la columna especificada
    qb_report_df = qb_report_df.dropna(subset=[qb_report_df.columns[indice_columna_numero_transaccion_qb_report]])

    # print(account_range_limit.index)
    # print(len(account_range_limit) / 2)
    qb_report_df.to_excel("TEST.xlsx",index=False,header=None)








def ejecutar_hilo():
    # Crear un objeto Thread y pasarle la función procesar como objetivo
    hilo_procesar = threading.Thread(target=procesar)

    # Iniciar la ejecución del hilo
    hilo_procesar.start()

# Crear la ventana principal
ventana = tk.Tk()

# Configurar el tamaño de la ventana
ventana.geometry("1200x600")

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
caja_texto = tk.Text(ventana, height=10, width=100)
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