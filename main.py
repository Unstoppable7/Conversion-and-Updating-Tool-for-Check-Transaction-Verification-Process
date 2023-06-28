import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import sys
import threading
import pandas as pd
import time
from datetime import datetime
from openpyxl import load_workbook

OUTPUT_DIRECTORY = "OUTPUT"

timestamp = str(datetime.now().strftime("%m%d%Y %p%#I"))

# Crear una clase personalizada para redirigir la salida a la caja de texto
class TextRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, msg):
        #Mostramos caja de texto
        self.text_widget.place(relx=0.5, rely=0.78, anchor="center")
        #Se realiza la escritura en la caja de texto
        self.text_widget.insert(tk.END, msg)
        self.text_widget.see(tk.END)

    def flush(self):
        pass

def abrir_quickbooks_report():
    archivo = filedialog.askopenfilename(title="Select Quickbooks report", filetypes=[("Archivos XLSX", "*.xlsx")])
    if archivo:
        entrada_quickbooks_report.delete(0, tk.END)
        entrada_quickbooks_report.insert(0, archivo)

def abrir_historial_de_transacciones_file():
    archivo = filedialog.askopenfilename(title="Select transactions history file", filetypes=[("Archivos XLSX", "*.xlsx")])
    if archivo:
        entrada_historial_de_transacciones_file.delete(0, tk.END)
        entrada_historial_de_transacciones_file.insert(0, archivo)

def clean(clean_inputs = True):
    caja_texto.delete(1.0, tk.END)
    if clean_inputs:
        entrada_quickbooks_report.delete(0, tk.END)

    #Ocultamos barra de progreso
    toggle_progress_bar()

def toggle_progress_bar(band=False, label = ""):
    global execution_in_progress

    if band:
        #Barra de progreso
        barra_progreso_label.config(text=label)
        barra_progreso_label.place(relx=0.5, rely=0.55, anchor="center")
        barra_progreso.place(relx=0.5, rely=0.6, anchor="center")
    else:   
        # Actualizar la barra de progreso
        barra_progreso["value"] = 0
        barra_progreso_label.config(text=label)
        barra_progreso_label.place_forget()
        barra_progreso.place_forget()
        execution_in_progress = False
    
    ventana.update_idletasks()

#TODO
def show_info_Quickbooks_report():
     messagebox.showinfo("Info", "Select the report files extracted from Quickbooks with check transactions\n\nEach file must represent an account, which will be searched in the TD Bank transactions report\n\nThis version uses the report extracted from Quickbooks Reports - Memorized Reports - Check Positive Pay")

def update_progress_bar(value, total_tasks, label = ""):

    #Actualizamos el label de la barra de progreso    
    barra_progreso_label.config(text=label)

     # Calcular el progreso basado en el número de tareas completadas
    progreso_a_sumar = (value) / total_tasks * 100

    # Actualizar la barra de progreso
    barra_progreso["value"] = barra_progreso["value"] + progreso_a_sumar

    ventana.update_idletasks()  # Actualizar la ventana para mostrar el progreso

def process():

    #Limpiamos la caja de texto de informacion
    clean(clean_inputs=False)
    #Ocultamos caja de texto si anteriormente fue mostrada
    caja_texto.place_forget()

    global execution_in_progress
    execution_in_progress = True
    
    #Variables que manejan la barra de progreso
    format_report_total_task = 6

    ##################################### Manejo de entrada de archivos
    input_quickbooks_report = entrada_quickbooks_report.get()
    input_transactions_history = entrada_historial_de_transacciones_file.get()

    #Tratamos si los archivos estan vacios
    if input_quickbooks_report == "":
        messagebox.showerror("Error", "No file selected as Quickbooks report")
        execution_in_progress = False
        return
    if input_transactions_history == "":
        messagebox.showerror("Error", "No file selected as Transactions History File")
        execution_in_progress = False
        return

    try:
        quickbooks_report_name = input_quickbooks_report.rsplit('/', 1)[1]
    except:
        messagebox.showerror("Error", "No file selected as Quickbooks report")
        execution_in_progress = False
        return
    try:
        transactions_history_name = input_transactions_history.rsplit('/', 1)[1]
    except:
        messagebox.showerror("Error", "No file selected as Transactions History File")
        execution_in_progress = False
        return
    
    #Mostramos barra de progreso
    barra_progreso["value"] = 25
    toggle_progress_bar(True, "Reading Files")
    # Verificacion de formato quickbook report
    try:
        qb_report_df = pd.read_excel(input_quickbooks_report, header=None, sheet_name="Sheet1")           

    except:
        messagebox.showerror("Error", "Check Positive Pay Quickbooks report '" + quickbooks_report_name + "' is not in the correct format. \n\nIt is necessary that the sheet where the report is located has the name 'Sheet1'. \n\nPlease check the Quickbooks report and try again")
        clean(clean_inputs=False)
        return

    try:

        if not ((not pd.isna(qb_report_df.iat[1,1])) and str(qb_report_df.iat[0, 4]) == "Date" and str(qb_report_df.iat[0, 6]) == "Num" and str(qb_report_df.iat[0, 8]) == "Name" and str(qb_report_df.iat[0, 10]) == "Credit"):

            messagebox.showerror("Error", "Check Positive Pay Quickbooks report '" + quickbooks_report_name + "' is not in the correct format. \n\nPlease check the Quickbooks report and try again")
            
            #Ocultamos barra de progreso y limpiamos caja de texto
            clean(clean_inputs=False)
            return
    except:
        messagebox.showerror("Error", "It is not possible to access the information of the Quickbooks report '" + quickbooks_report_name + "' \n\nPlease check the Quickbooks report and try again")
        clean(False)
        return
    
    # Verificacion de formato transactions history file
    try:
        #Esto cargará todas las hojas del archivo Excel en un diccionario llamado transactions_history_df, donde las claves del diccionario serán los nombres de las hojas y los valores serán los DataFrames correspondientes
        transactions_history_df_dic = pd.read_excel(input_transactions_history, sheet_name=None)           

    except:
        messagebox.showerror("Error", "'" + transactions_history_name + "' is not in the correct format. \n\nIt is necessary that the sheet where the report is located has the name 'Sheet1'. \n\nPlease check the Quickbooks report and try again")
        clean(clean_inputs=False)
        return
    
    try:
        #Accedemos al archivo con el historial de transacciones
        transactions_history_workbook = load_workbook(filename = input_transactions_history)
    except:
        messagebox.showerror("Error",f"Could not access file {input_transactions_history} \nMake sure you have selected the correct file")

        clean(clean_inputs=False)
        return

    #Acutalizamos barra de progreso
    update_progress_bar(1,format_report_total_task,"Deleting unnecessary data to the report Check Positive Pay")

    #Extraemos la columna 1 y 10, a su vez eliminamos las filas que tengan vacia la celda en la columna 1
    account_range_limit = qb_report_df.iloc[:,[1,10]].dropna(subset=qb_report_df.columns[1])

    ############### Eliminar columnas vacias y extras
    columnas_a_eliminar_qb_report = [0,1,2,3,5,7,9]  # Índices de las columnas a eliminar
    columnas_extras_a_eliminar_qb_report = 4 #A partir del este numero se eliminaran
    try:
        qb_report_df = qb_report_df.drop(qb_report_df.columns[columnas_a_eliminar_qb_report], axis=1)
    except Exception as e:
        messagebox.showerror("Error", f"{e}\n\n There was an error trying to remove unnecessary columns, please make sure the file to process has the required format")

        clean(False)
        return

    if qb_report_df.shape[1] > 4:
        qb_report_df = qb_report_df.drop(qb_report_df.columns[columnas_extras_a_eliminar_qb_report:], axis=1)
    
    #Comparamos la cantidad de filas que tiene qb_report, resto 1 para que sea acorde con la otra comparacion que es por indices (comienza desde 0), CON la cantidad de filas hasta la ultima fila que contiene datos que necesitamos. De esta manera verificamos si hay mas filas extra que se deben de eliminar
    if (len(qb_report_df.index) - 1) > (account_range_limit.index[len(account_range_limit) - 1]):

        # Eliminamos filas fuera de rango de los datos
        qb_report_df = qb_report_df[:account_range_limit.index[len(account_range_limit) - 1] + 1]

    # Eliminar filas que datos vacios, variable calculada anteriormente
    qb_report_df = qb_report_df.drop(account_range_limit.index, axis=0)

    #Eliminamos la primera fila con los titulos
    qb_report_df = qb_report_df.drop(0, axis=0)

    #Acutalizamos barra de progreso
    update_progress_bar(1,format_report_total_task,"Inserting account numbers to the report Check Positive Pay")

    try:
        qb_report_df[qb_report_df.columns[1]] = qb_report_df[qb_report_df.columns[1]].astype(int)
    except Exception as e:

        rsp = messagebox.askyesno("Error", f"{e}\n\nThere was a problem when trying to format the column corresponding to account number \n\nDo you want to continue anyway?")

        if not rsp:
            clean(False)
            return
    
    #Eliminamos filas que no contengan numero de cheque
    indice_columna_numero_transaccion_qb_report = 1

    # Eliminar las filas con valores nulos en la columna especificada
    qb_report_df = qb_report_df.dropna(subset=[qb_report_df.columns[indice_columna_numero_transaccion_qb_report]])

    qb_report_df_dic = {}
    void_transactions = pd.DataFrame()
    ############## Insertar una nueva columna llamada 'tmp' y asignar valores según los rangos, agrega el rango desde uno menos del inicio hasta uno menos del final
    #Ordenar las transacciones de cada cuenta
    for i in range(int(len(account_range_limit) / 2)):

        start_index = account_range_limit.index[i * 2]
        end_index = (account_range_limit.index[(i * 2) + 1])
        account_number = str(account_range_limit.iat[i * 2, 0])[:4]

        qb_report_df.loc[start_index:end_index, 'tmp'] = "439780" + account_number

        #Exportamos cada dataframe hacia un diccionario
        qb_report_df_dic[account_number] = qb_report_df.loc[start_index:end_index, :]
        
        #Reseteamos los indices de las filas
        qb_report_df_dic[account_number] = qb_report_df_dic[account_number].reset_index(drop=True)
        #Reseteamos los indices de las columnas
        qb_report_df_dic[account_number].columns = range(qb_report_df_dic[account_number].shape[1])

    ###############Busqueda y comparacion de transacciones para hallar nuevas transacciones a verificar
    #Acutalizamos barra de progreso
    update_progress_bar(1,format_report_total_task,"Looking for transactions to verify")
    accounts_no_processed = []
    transactions_per_account_to_verify = {}
    for account, transaction_history_df in transactions_history_df_dic.items():
        
        #Verificamos si en nuestro reporte de qb tiene transacciones de la cuenta a procesar
        if account in qb_report_df_dic:
            qb_report_df = qb_report_df_dic[account]
            
            #Reseteamos los indices de las columnas
            transactions_history_df_dic[account].columns = range(transactions_history_df_dic[account].shape[1])

            ###Le damos a la columna a comparar formato entero para que se realice correctamente la comparacion
            transactions_history_df_dic[account][transactions_history_df_dic[account].columns[2]] = transaction_history_df[transaction_history_df.columns[2]].astype(int)
            qb_report_df[qb_report_df.columns[1]] = qb_report_df[qb_report_df.columns[1]].astype(int)

            #Realizamos merge tomando solamente los datos del df de la izq
            transactions_merge = pd.merge(qb_report_df, transaction_history_df, left_on=qb_report_df.columns[1], right_on=transaction_history_df.columns[2], how="left", sort=True, indicator=True)

            #Filtramos el merge para tener solo las transacciones que no tuvieron coincidencia (transacciones nuevas a verificar)
            transactions_merge_filtered = transactions_merge[transactions_merge['_merge'] == 'left_only']
            
            #Puede pasar que el qb report si tenga transacciones pero que ninguna sea nueva
            if transactions_merge_filtered.empty:
                #Guardamos las cuentas que no tienen transacciones nuevas
                accounts_no_processed.append(account)
            else:
                ###Guardamos los dataframes de cada cuenta con sus transacciones a verificar
                ###Aplicamos formato comun para archivo salida y datos a insertar

                #Reseteamos los indices de las filas
                transactions_merge_filtered = transactions_merge_filtered.reset_index(drop=True)
                #Reseteamos los indices de las columnas
                transactions_merge_filtered.columns = range(transactions_merge_filtered.shape[1])
                
                #Eliminamos la columna que tiene el numero de transaccion duplicado
                transactions_merge_filtered = transactions_merge_filtered.drop([0],axis=1)

                #Eliminamos columnas innecesarias que me deja el merge
                index_to_remove_from = 4 #A partir del este numero se eliminaran
                if transactions_merge_filtered.shape[1] > index_to_remove_from:
                    transactions_merge_filtered = transactions_merge_filtered.drop(transactions_merge_filtered.columns[index_to_remove_from + 1:], axis=1)

                ##Formato comun

                ### Movemos la columna de numero de cuenta
                position = 1  # Índice de la posición deseada
                columns = transactions_merge_filtered.columns.tolist()
                column_to_move = columns[4]  # Índice de la columna que deseas mover 
                columns.remove(column_to_move)
                columns.insert(position, column_to_move)
                transactions_merge_filtered = transactions_merge_filtered[columns]

                ### Movemos la columna de monto
                position = 3  # Índice de la posición deseada
                columns = transactions_merge_filtered.columns.tolist()
                column_to_move = columns[4]  # Índice de la columna que deseas mover 
                columns.remove(column_to_move)
                columns.insert(position, column_to_move)
                transactions_merge_filtered = transactions_merge_filtered[columns]

                
                # Extraemos transacciones sin monto (void)
                transaction_amount_column = 3
                void_transactions = pd.concat([void_transactions, transactions_merge_filtered.loc[transactions_merge_filtered[transactions_merge_filtered.columns[transaction_amount_column]].isnull()]])

                #Eliminamos las transacciones en las que el monto sea vacio (void)
                transactions_merge_filtered = transactions_merge_filtered.dropna(subset=[transactions_merge_filtered.columns[transaction_amount_column]])

                # Convertir las columnas a tipo de datos de fecha
                try:
                    transactions_merge_filtered[transactions_merge_filtered.columns[0]] = pd.to_datetime(transactions_merge_filtered[transactions_merge_filtered.columns[0]]).dt.strftime('%m/%d/%Y')
                except Exception as e:
                    rsp = messagebox.askyesno("Error", f"{e}\n\nThere was a problem when trying to format the column corresponding to the date\n\nDo you want to continue anyway?")

                    if not rsp:
                        clean(False)
                        return

                #Agregamos resultado al diccionario
                transactions_per_account_to_verify[account] = transactions_merge_filtered
        else: 
            #Guardamos las cuentas que no tienen transacciones en el qb report
            accounts_no_processed.append(account)

    # Concatenar los DataFrames del diccionario en uno solo
    td_bank_report_df = pd.concat(transactions_per_account_to_verify.values(), ignore_index=True)

    ############################# Format TD BANK REPORT
    #Acutalizamos barra de progreso
    update_progress_bar(1,format_report_total_task,"Formatting TD Bank report")

    ### Insertamos columna con la letra I
    td_bank_report_df.insert(3, '', 'I')

    try:
        #Le damos formato a las fechas del dataframe de las transacciones void
        void_transactions[void_transactions.columns[0]] = pd.to_datetime(void_transactions[void_transactions.columns[0]]).dt.strftime('%m/%d/%Y')

    except Exception as e:
        pass

    void_transactions_name_file = f"VOID TRANSACTIONS {timestamp}.xlsx"

    try_again = True
    while try_again:  
        try:
            #Guardamos las transacciones void en un archivo aparte
            if(not void_transactions.empty):

                void_transactions.to_excel(void_transactions_name_file, header=None, index=False)  
                print(f"Void transactions file has been created\n")

            try_again = False
        except PermissionError:
            
            rsp = messagebox.askretrycancel("Permission error", f"Could not update file '{void_transactions_name_file}' \nIf you have this file open please close it \n\nDo you want to try again?")

            if not rsp:
                clean(False)
                return
        except Exception as e:
            messagebox.showerror("Error", str(e) + "\n\nFailed to export file '" + void_transactions_name_file + "'")
            #Ocultamos barra de progreso y limpiamos caja de texto
            clean(False)
            return
    
    ######## Formato especifico a los datos
    ####Formato tipo de datos
    try:
        td_bank_report_df[td_bank_report_df.columns[1]] = td_bank_report_df[td_bank_report_df.columns[1]].astype(str)
    except Exception as e:

        rsp = messagebox.askyesno("Error", f"{e}\n\nThere was a problem when trying to format the column corresponding to account number \n\nDo you want to continue anyway?")

        if not rsp:
            clean(False)
            return
    try:
        td_bank_report_df[td_bank_report_df.columns[2]] = td_bank_report_df[td_bank_report_df.columns[2]].astype(str)
    except Exception as e:
        rsp = messagebox.askyesno("Error", f"{e}\n\nThere was a problem when trying to format the column corresponding to transaction number \n\nDo you want to continue anyway?")

        if not rsp:
            clean(False)
            return
        
    try:
        td_bank_report_df[td_bank_report_df.columns[4]] = td_bank_report_df[td_bank_report_df.columns[4]].astype(float)
    except Exception as e:
        rsp = messagebox.askyesno("Error", f"{e}\n\nThere was a problem when trying to format the column corresponding to transaction amount \n\nDo you want to continue anyway?")

        if not rsp:
            clean(False)
            return

    #Eliminar simbolos de la columna nombres, exceptuando los espacios
    td_bank_report_df[td_bank_report_df.columns[5]] = td_bank_report_df[td_bank_report_df.columns[5]].replace(r'[^a-zA-Z0-9\s]', '', regex=True)
    #Cortamos los nombres hasta un maximo de 30 caracteres
    td_bank_report_df[td_bank_report_df.columns[5]] = td_bank_report_df[td_bank_report_df.columns[5]].str.slice(0, 30)
    
    #Acutalizamos barra de progreso
    update_progress_bar(1,format_report_total_task,"Saving TD Bank report")
    
    result_file_name_xlsx = f"APDC {timestamp}.xlsx"
    result_file_name_csv = f"APDC {timestamp}.csv"

    retry = True
    while retry:
        try:
            # Crear el objeto ExcelWriter con el formato deseado
            writer = pd.ExcelWriter(result_file_name_xlsx, engine='xlsxwriter', date_format= "mm/dd/yyy")
            td_bank_report_df.to_excel(writer, sheet_name="Sheet1", index=False,header=None)

            retry = False
        except PermissionError:
            retry = messagebox.askretrycancel("Error", f"{PermissionError}\n\nProbably '{result_file_name_xlsx}' is open\n\nIf that is the case please close the file and try again")
            if not retry:
                clean(False)
                return
        except Exception as e:
            messagebox.showerror("Error", f"{e}\n\nThere is a problem with {result_file_name_xlsx}")

            clean(False)
            return

    retry = True
    while retry:
        try:
            # Convert the dataframe to an XlsxWriter Excel object.
            td_bank_report_df.to_csv(result_file_name_csv,index=False,header=None, date_format='%m/%d/%Y', float_format='%.2f')

            retry = False
        except PermissionError:
            retry = messagebox.askretrycancel("Error", f"{PermissionError}\n\nProbably '{result_file_name_csv}' is open\n\nIf that is the case please close the file and try again")
            if not retry:
                clean(False)
                return
        except Exception as e:
            messagebox.showerror("Error", f"{e}\n\nThere is a problem with {result_file_name_csv}")

            clean(False)
            return
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    # Agregar algunos formatos de celda.
    format1 = workbook.add_format({"num_format": "#,##0.00"})

    # Aplicar los formatos a las columnas específicas.
    worksheet.set_column('E:E', None, format1)  

    # Cerrar el escritor de Excel de Pandas y guardar el archivo Excel.
    writer.close()

    #Reiniciamos barra de progreso
    barra_progreso["value"] = 100   
    ventana.update_idletasks()
    time.sleep(0.5)

    ############################ UPDATING TRANSACTIONS HISTORY FILE
    #Reiniciamos barra de progreso
    barra_progreso["value"] = 0   
    #Acutalizamos barra de progreso
    update_progress_bar(1,format_report_total_task,f"Updating {transactions_history_name}")

    for account, df in transactions_per_account_to_verify.items():

        ######Formato especifico
        ### Movemos la columna del monto de la transaccion
        position = 4  # Índice de la posición deseada
        columns = df.columns.tolist()
        column_to_move = columns[3]  # Índice de la columna que deseas mover 
        columns.remove(column_to_move)
        columns.insert(position, column_to_move)
        transactions_per_account_to_verify[account] = df[columns]
        df = df[columns]

        #Agregamos el timestamp
        # Obtener el índice de la última fila con datos
        last_row_index = df.last_valid_index()

        # Obtener el número de columnas en el DataFrame
        num_columns = df.shape[1]

        # Agregar una nueva columna después de la última columna con datos
        df.insert(num_columns, '', '')

        # Asignar el valor al dato en la celda específica
        df.at[last_row_index, df.columns[num_columns]] = timestamp

        try:
            #Accedemos a la hoja de la cuenta a procesar
            transactions_history_workbook_page = transactions_history_workbook[account]
        except KeyError:
            messagebox.showerror("Error", "Page '" + account + "' not found in '" + transactions_history_name + "'\n\nPlease verify that '" + transactions_history_name + "' contains '" + account + "' sheet")
            #Ocultamos barra de progreso y limpiamos caja de texto
            clean(False)
            return  
        except Exception as e:
            messagebox.showerror("Error", f"{e}\n\nPlease check {transactions_history_name} and {account} sheet")
            #Ocultamos barra de progreso y limpiamos caja de texto
            clean(False)
            return
        
        #Pasamos la informacion del dataframe () a una lista
        data_frame_rows = df.values.tolist()

        #Proceso de insercion de los datos al page del workbook
        process_approved = False
        retry = True
        while(not process_approved and retry):

            try:
                for row in data_frame_rows:
                    transactions_history_workbook_page.append(row)


                process_approved = True
                
            except PermissionError:

                retry = messagebox.askretrycancel("Error inserting verified transactions", f"Could not access file {transactions_history_name} \nIt is probably open \n\nDo you want to try again?")

            except Exception as e:
                
                retry = messagebox.askretrycancel("Error", f"{e}\n\nPlease check {transactions_history_name}\n\nDo you want to try again?")
    
        if not retry:
            messagebox.showwarning("Warning", f"Account {account} will not be updated on {transactions_history_name}")
            continue
    
    ################################## SAVE TRANSACTION HISTORY FILE
    process_approved = False
    while(not process_approved and retry):
        
        #Acutalizamos barra de progreso
        update_progress_bar(1,format_report_total_task,f"Saving {transactions_history_name}")
        #Reiniciamos barra de progreso

        time.sleep(1)
        #Reiniciamos barra de progreso
        barra_progreso["value"] = 50   
        ventana.update_idletasks()

        try:
            transactions_history_workbook.save(input_transactions_history)
            
            process_approved = True
        except PermissionError:

            retry = messagebox.askretrycancel("Error saving file", f"Could not access file {transactions_history_name} \nIt is probably open \n\nDo you want to try again?")

        except Exception as e:

            retry = messagebox.askretrycancel("Error", f"{e}\n\nPlease check {transactions_history_name}\n\nDo you want to try again?")

    #Reiniciamos barra de progreso
    barra_progreso["value"] = 100   
    ventana.update_idletasks()

    time.sleep(1)

    #Ocultamos barra de progreso
    toggle_progress_bar(False)

    execution_in_progress = False
    ################################## END 



    #Acutalizamos barra de progreso
    update_progress_bar(1,format_report_total_task,"Ending process")
    time.sleep(0.5)
    toggle_progress_bar()
    messagebox.showinfo("Sucess", "The process has finished successfully")

def run_thread():

    global execution_in_progress

    if execution_in_progress:
        messagebox.showwarning("Warning","The execution is already in process\n\nPlease wait for it to finish processing to start another process")
    else:
        # Crear un objeto Thread y pasarle la función procesar como objetivo
        thread_to_process = threading.Thread(target=process)

        # Iniciar la ejecución del hilo
        thread_to_process.start()

execution_in_progress = False

# Crear la ventana principal
ventana = tk.Tk()

# Configurar el tamaño de la ventana
ventana.geometry("800x500")

# Configurar el mensaje de bienvenida
mensaje_bienvenida = tk.Label(ventana, text="Formatting for TD Bank Transaction Verification", font=("Arial", 20))

subtitulo_quickbooks_report = tk.Label(ventana, text="Check Positive Pay Quickbooks Report", font=("Arial", 14))
entrada_quickbooks_report = tk.Entry(ventana, width=40)
boton_quickbooks_report = tk.Button(ventana, text="Select file", command=abrir_quickbooks_report, font=("Arial", 10))
# boton_info_quickbooks_report = tk.Button(ventana, text="Info", command=show_info_Quickbooks_report)

subtitulo_historial_de_transacciones_file = tk.Label(ventana, text="Transaction History File", font=("Arial", 14))
entrada_historial_de_transacciones_file = tk.Entry(ventana, width=40)
boton_historial_de_transacciones_file = tk.Button(ventana, text="Select file", command=abrir_historial_de_transacciones_file, font=("Arial", 10))

# Configurar el botón de procesamiento
boton_procesar = tk.Button(ventana, text="Start Process", font=("Arial", 18), command=run_thread)
# boton_procesar.grid(row=4, column=0, columnspan=2, pady=20)

# Configurar la caja de texto
caja_texto = tk.Text(ventana, height=7, width=50)
# caja_texto.grid(row=5, column=0, columnspan=2, padx=20)

# Configurar el botón "Clean"
# boton_clean = tk.Button(ventana, text="Clean", font=("Arial", 14), command=clean)
# boton_clean.grid(row=6, column=0, columnspan=2, pady=10)

# Crear una barra de progreso
barra_progreso = ttk.Progressbar(ventana, mode="determinate", length=300)
barra_progreso_label = tk.Label(ventana, text="", font=("Arial", 12))

#Place
mensaje_bienvenida.place(relx=0.5, rely=0.1, anchor="center")

# Configurar los subtitulos y las cajas de texto de archivo central
subtitulo_quickbooks_report.place(relx=0.7, rely=0.22, anchor="center")
entrada_quickbooks_report.place(relx=0.7, rely=0.27, anchor="center")
boton_quickbooks_report.place(relx=0.7, rely=0.33, anchor="center")
# boton_info_quickbooks_report.place(relx=0.68, rely=0.25, anchor="center")

# Configurar los subtitulos y las cajas de texto de archivo de historial de transacciones
subtitulo_historial_de_transacciones_file.place(relx=0.3, rely=0.22, anchor="center")
entrada_historial_de_transacciones_file.place(relx=0.3, rely=0.27, anchor="center")
boton_historial_de_transacciones_file.place(relx=0.3, rely=0.33, anchor="center")

# Configurar el botón de procesamiento
boton_procesar.place(relx=0.5, rely=0.45, anchor="center")

# Configurar la caja de texto
# caja_texto.place(relx=0.5, rely=0.742, anchor="center")

# Configurar el botón "Clean"
# boton_clean.place(relx=0.5, rely=0.94, anchor="center")

# Redirigir la salida estándar y la salida de error a la caja de texto
sys.stdout = TextRedirector(caja_texto)
sys.stderr = TextRedirector(caja_texto)

# # Establecer el icono de la ventana
# icono = 'APDC LOGO.ico'  # Reemplaza con la ruta completa del archivo de icono
# ventana.iconbitmap(icono)
ventana.title("APDC Check verification process")

# Ejecutar el bucle principal de la interfaz gráfica
ventana.mainloop()