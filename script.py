import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
import warnings
from io import StringIO

data = None 

def read_data(file_path):
    global data
    if file_path:
        # Leer xlsx
        input_data = pd.read_excel(file_path)

        # Convertir a csv y almacenar
        csv_data = StringIO()
        input_data.to_csv(csv_data, index=False)
        csv_data.seek(0) 
        data = pd.read_csv(csv_data, skiprows=8)

        # Convertir fechas datetime
        data['Recibido'] = pd.to_datetime(data['Recibido'])

        # Registros totales
        total_records = len(data)

        # Escritos / Proyectos
        presentaciones_count = len(data[data['Tipo'].str.contains('escrito', case=False)])
        proyectos_count = len(data[data['Tipo'].str.contains('proyecto', case=False)])

        # Escrito más antiguo y más nuevo
        oldest_record = min(data['Recibido'])
        oldest_record_formatted = oldest_record.strftime("%d/%m/%Y")
        newest_record = max(data['Recibido'])

        # Cantidad de expedientes con presentaciones
        unique_exptes_count = data['Expte'].nunique()

        # Cantidad de transferencias
        transferencias_count = len(data[data['Título'].str.contains('transferencia', case=False)])

        # Diferencia de días a la fecha
        today_date = datetime.now()
        days_difference = (today_date - oldest_record).days
        today_formatted = today_date.strftime("%d/%m/%Y")

        # Repeticiones en escritos (top 10)
        most_titles = data['Título'].value_counts().head(10).to_dict()
        most_titles_df = pd.DataFrame(list(most_titles.items()), columns=['Suma', 'Cantidad'])

        # Información sumaria
        info_text = (f'\n'
                     f'     {total_records} presentaciones en un total de {unique_exptes_count} causas.\n'
                     f'\n'
                     f'     Escritos: {presentaciones_count} / Proyectos: {proyectos_count}\n'
                     f'\n'
                     f'     Escrito más antiguo {oldest_record_formatted} - {days_difference} días corridos al {today_formatted}\n'
                     f'\n'
                     f'     {transferencias_count} escritos incluyen la palabra "transferencia".\n'
                     f'\n'
                     f'     Escritos más repetidos:\n{most_titles_df.to_string(index=False)}')
        return info_text
    else:
        return "    No se ha seleccionado ningún archivo."

def create_listados(data):
    """
    Estandariza y limpia el listado para imprimir a proveyentes
    Args:
    - data: DataFrame conteniendo el listado en csv a partir de la fila 9
    """
    # fecha
    today_date = datetime.now()

    # Calcula la diferencia en días desde la presentación de cada escrito a la fecha
    data['Recibido'] = pd.to_datetime(data['Recibido'])  
    data['DaysDifference'] = (today_date - data['Recibido']).dt.days
    data['Recibido'] = pd.to_datetime(data['Recibido']).dt.strftime('%d/%m/%y')
    
    # Crear listados
    listados = []
    for index, row in data.iterrows():
        # Agrega la columna de diferencia en días en cada escrito
        fifth_column_string = f"{row['DaysDifference']} días al {today_date.strftime('%d/%m')}"
        listados.append((row['Título'], row['Expte'], row['Recibido'], row['Apellido'], fifth_column_string))

    return listados

def save_listados_to_excel(listados, filename):
    """
    Guarda el listado en xlsx
    
    Args:
    - listados: datos
    - filename: Nombre del archivo de salida
    
    """
    # Crear Workbook y worksheey
    wb = Workbook()
    ws = wb.active
    ws.title = "Listados"

    # escribir data en worksheet
    for listado in listados:
        ws.append(listado)
    
    wb.save(filename)

def select_file():
    file_path = filedialog.askopenfilename()
    info_text = read_data(file_path)
    text_area.delete('1.0', tk.END)
    text_area.insert(tk.END, info_text)

def save_file():
    global data  
    current_datetime = datetime.now()
    datestamp = current_datetime.strftime("%d-%m-%Y_%H-%Mhs")
    filename = f"listado_{datestamp}.xlsx"
    if data is not None:
        listados = create_listados(data)
        save_listados_to_excel(listados, filename)
        text_area.insert(tk.END, f"\n\nArchivo guardado como '{filename}'.")

# Crear ventana Tkinter
root = tk.Tk()
root.title("Listado de escritos")

# Tamaño de ventana
window_width = 700
window_height = 590
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width / 2) - (window_width / 2)
y_coordinate = (screen_height / 2) - (window_height / 2)
root.geometry(f"{window_width}x{window_height}+{int(x_coordinate)}+{int(y_coordinate)}")

# Texto descriptivo
descriptive_label = tk.Label(root, text="Seleccionar el archivo de Forum:")
descriptive_label.pack(pady=5)

# Botón para seleccionar archivo
select_file_button = tk.Button(root, text="Seleccionar Archivo", command=select_file)
select_file_button.pack(pady=10)

# Área de texto para mostrar información
text_area = tk.Text(root, wrap="word")
text_area.pack(fill="both", padx=10, pady=10)

# Botón para guardar archivo
save_file_button = tk.Button(root, text="Exportar listado", command=save_file)
save_file_button.pack(pady=10)

# Filtrar UserWarning de openpyxl.styles.stylesheet module
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

# Ejecutar Tkinter
root.mainloop()