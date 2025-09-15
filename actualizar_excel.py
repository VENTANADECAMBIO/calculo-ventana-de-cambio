
import pandas as pd
from datetime import datetime, timedelta

# Cargar el archivo Excel
file_path = "Calculo ventana de cambio.xlsx"
lead_time_df = pd.read_excel(file_path, sheet_name="Lead time", engine="openpyxl")
festivos_df = pd.read_excel(file_path, sheet_name="festivos", engine="openpyxl", header=None)

# Obtener la fecha actual del sistema
fecha_actual = datetime.now()

# Convertir festivos a lista de fechas
festivos = pd.to_datetime(festivos_df[0]).dt.date.tolist()

# Función para calcular fecha de entrega excluyendo fines de semana y festivos
def calcular_fecha_entrega(fecha_inicio, dias_habiles):
    fecha = fecha_inicio
    contador = 0
    while contador < dias_habiles:
        fecha += timedelta(days=1)
        if fecha.weekday() < 5 and fecha.date() not in festivos:
            contador += 1
    return fecha

# Aplicar cálculo de fecha de entrega para cada ciudad
lead_time_df["Fecha de Consulta"] = fecha_actual
lead_time_df["Fecha de Entrega"] = lead_time_df["dias  Habiles"].apply(lambda x: calcular_fecha_entrega(fecha_actual, x))

# Guardar el archivo actualizado
lead_time_df.to_excel(file_path, sheet_name="Lead time", index=False)
