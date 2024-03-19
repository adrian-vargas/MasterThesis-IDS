import os
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

# Directorio que contiene los archivos HTML: obtiene la ruta del directorio del script actual
directory = os.path.dirname(os.path.abspath(__file__))

# Verifica si el directorio existe
if os.path.exists(directory):
    print(f"El directorio '{directory}' existe.")
else:
    print(f"El directorio '{directory}' no existe.")

# Nombre del archivo Excel de salida
output_excel_file = 'all_response_times.xlsx'

# Verifica si el archivo Excel existe y lo borra si es así
if os.path.exists(output_excel_file):
    os.remove(output_excel_file)

# Diccionario que almacena los datos organizados por nombre de archivo
all_questions_data = {}

# Itera sobre cada archivo en el directorio
for filename in os.listdir(directory):
    if filename.endswith('.html'):  
        filepath = os.path.join(directory, filename)
        
        # Lista para almacenar los tiempos de respuesta para este archivo
        response_times = []
        
        # Lee el contenido del archivo HTML
        with open(filepath, 'r', encoding='utf-8') as file:
            html_content = file.read()

        # Analiza el contenido HTML
        soup = BeautifulSoup(html_content, 'html.parser')

        # Encuentra todos los contenedores de preguntas
        question_containers = soup.find_all('div', class_='que')

        # Inicializa una variable para almacenar el tiempo del paso 2 anterior
        previous_step_2_time = None

        # Itera sobre cada contenedor de pregunta
        for container in question_containers:
            # Extrae el número de la pregunta
            question_number = container.find('span', class_='qno').text.strip()

            # Encuentra la tabla asociada con esta pregunta
            general_table = container.find_next('table', class_='generaltable')
            
            if general_table:
                rows = general_table.find_all('tr')
                
                # Encuentra las filas del paso 2
                step_2_rows = [row for row in rows if 'Guardada:' in row.text]
                
                if step_2_rows:
                    # Extrae el tiempo de la fila del paso 2
                    step_2_time_str = step_2_rows[0].find_all('td')[1].text.strip()
                    step_2_time = datetime.strptime(step_2_time_str, '%d/%m/%y, %H:%M:%S')
                    
                    # Calcula la diferencia de tiempo
                    if previous_step_2_time is None:
                        # Para la primera pregunta, encuentra la fila del paso 1 para calcular la diferencia
                        step_1_rows = [row for row in rows if 'Iniciado/a' in row.text]
                        step_1_time_str = step_1_rows[0].find_all('td')[1].text.strip()
                        step_1_time = datetime.strptime(step_1_time_str, '%d/%m/%y, %H:%M:%S')
                        time_difference = (step_2_time - step_1_time).total_seconds()
                    else:
                        # Para las siguientes preguntas, calcula la diferencia con el paso 2 anterior
                        time_difference = (step_2_time - previous_step_2_time).total_seconds()
                    
                    # Actualiza el tiempo del paso 2 anterior
                    previous_step_2_time = step_2_time
                    
                    # Guarda la diferencia de tiempo en la lista correspondiente al archivo
                    response_times.append(time_difference)
            
            # Valida que cada pregunta tenga un tiempo de respuesta, incluso si es cero
            while len(response_times) < int(question_number):
                response_times.append(0)
            
        # Guarda los tiempos de respuesta en el diccionario con el nombre del archivo como clave
        all_questions_data[filename] = response_times

# Crea un DataFrame con los datos de todas las preguntas de todos los archivos
df = pd.DataFrame.from_dict(all_questions_data, orient='index')

# Transpone el DataFrame para que los nombres de archivo sean encabezados de columna y resetea el índice
df = df.transpose().reset_index(drop=True)

# Cambia el nombre de la columna "Question Number" a partir de 1 en lugar de 0
df.columns.name = 'Question Number'
df.index += 1 

# Guarda el DataFrame en un archivo Excel
df.to_excel(output_excel_file, sheet_name='Time Difference in Seconds', index_label='Question Number')
print(f"Todos los tiempos de respuesta han sido guardados en {output_excel_file}")
