import openpyxl
    
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

# Mapeo de nombres de universidades
UNIVERSITY_SHORTNAMES = {
    'GA - MIT xPRO': 'xPRO',
    'GA - The University of Chicago': 'UCPE',
    'GA - Chicago Booth Executive Education': 'Booth',
    'GA - MIT Professional Education': 'MIT PE',
    'GA - Miami Herbert Business School': 'UMiami',
    'GA - University of Chicago Graham School': 'Graham',
    'GA - University of Pennsylvania': 'UPenn',
    'GA - Stanford Center for Professional Development': 'STF'
}

def ajustar_fecha_fin_semana(fecha):
    """Ajusta la fecha al viernes anterior si cae en sábado o domingo"""
    if fecha.weekday() == 5:  # Sábado
        return fecha - timedelta(days=1)
    elif fecha.weekday() == 6:  # Domingo
        return fecha - timedelta(days=2)
    return fecha

def load_data(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        elif file.name.endswith('.xlsx'):
            df = pd.read_excel(file)
        else:
            st.error('Tipo de archivo no soportado. Sube un archivo CSV o Excel.')
            return pd.DataFrame()

        # Convertir y ajustar fechas
        df['Fecha_de_Envio'] = pd.to_datetime(df['Fecha_de_Envio'], format='%d-%m-%Y', errors='coerce')
        df['Fecha_de_Envio'] = df['Fecha_de_Envio'].apply(ajustar_fecha_fin_semana)
        
        # Convertir nombres de universidades a versiones cortas
        df['Universidad'] = df['Universidad'].map(UNIVERSITY_SHORTNAMES).fillna(df['Universidad'])
        
        # Procesar columna Batch
        if 'Batch' in df.columns:
            try:
                # Intentar parsear como fecha y formatear
                df['Batch'] = pd.to_datetime(df['Batch'], errors='ignore')
                df['Batch'] = df['Batch'].dt.strftime('%B %Y')
            except:
                # Si falla, dejar el valor original
                df['Batch'] = df['Batch'].astype(str)
        else:
            st.warning("La columna 'Batch' no está presente en los datos. Se generará automáticamente.")
            df['Batch'] = df['Fecha_de_Envio'].dt.strftime('%B %Y')
        
        return df
    except Exception as e:
        st.error(f"Error al cargar el archivo: {e}")
        return pd.DataFrame()

def generate_calendar(df):
    if df.empty:
        return pd.DataFrame()

    # Crear rango de fechas
    min_date = df['Fecha_de_Envio'].min().replace(day=1)
    max_date = df['Fecha_de_Envio'].max()
    if max_date.day < 28:
        max_date = max_date + pd.offsets.MonthEnd()
    
    all_dates = pd.date_range(start=min_date, end=max_date, freq='D')
    
    # Crear DataFrame del calendario
    calendar_df = pd.DataFrame(all_dates, columns=['Date'])
    calendar_df['Day'] = calendar_df['Date'].dt.day
    calendar_df['Weekday'] = calendar_df['Date'].dt.weekday
    calendar_df['Week'] = calendar_df['Date'].dt.isocalendar().week
    calendar_df['Month'] = calendar_df['Date'].dt.strftime('%B')
    
    # Combinar con datos de reportes (universidad + batch)
    report_dates = df.groupby('Fecha_de_Envio').apply(
        lambda x: ', '.join([f"{row['Universidad']} ({row['Batch']})" for _, row in x.iterrows()])
    ).reset_index(name='Universidad')
    
    calendar_df = pd.merge(calendar_df, report_dates, left_on='Date', right_on='Fecha_de_Envio', how='left')
    
    # Filtrar solo días laborables (Lunes-Viernes)
    calendar_df = calendar_df[calendar_df['Weekday'] < 5]
    
    # Crear tabla pivote
    pivot_df = calendar_df.pivot_table(index=['Week', 'Month'], 
                                     columns='Weekday', 
                                     values='Universidad', 
                                     aggfunc='first')
    
    # Obtener números de día como strings (sin decimales)
    day_numbers = calendar_df.pivot_table(index=['Week', 'Month'], 
                                        columns='Weekday', 
                                        values='Day', 
                                        aggfunc=lambda x: str(int(x.iloc[0])) if not pd.isna(x.iloc[0]) else np.nan)
    
    # Asegurar que tenemos los 5 días laborables
    for i in range(5):
        if i not in pivot_df.columns:
            pivot_df[i] = np.nan
        if i not in day_numbers.columns:
            day_numbers[i] = np.nan
    
    # Ordenar columnas
    pivot_df = pivot_df.reindex(sorted(pivot_df.columns), axis=1)
    day_numbers = day_numbers.reindex(sorted(day_numbers.columns), axis=1)
    
    # Combinar datos
    result_df = pivot_df.copy()
    for col in result_df.columns:
        result_df[col] = result_df[col].where(pd.notnull(result_df[col]), day_numbers[col])
    
    # Renombrar columnas
    result_df.columns = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
    
    return result_df.reset_index()

def style_calendar(df):
    def contar_universidades(x):
        if isinstance(x, str) and x.replace('.','',1).isdigit():  # Si es número de día
            return 0
        elif isinstance(x, str):
            return len(x.split(','))  # Contar universidades
        return 0
    
    counts = df[['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']].applymap(contar_universidades)
    min_val = counts.min().min()
    max_val = counts.max().max()
    
    def color_celda(val):
        if isinstance(val, str) and ',' in val:  # Múltiples entregas
            count = len(val.split(','))
            intensity = (count - min_val) / (max_val - min_val) if max_val > min_val else 0.5
            return f'background-color: rgba(0, 100, 255, {0.3 + intensity*0.7}); color: white'
        elif isinstance(val, str) and '(' in val:  # Entrega única
            return 'background-color: rgba(0, 100, 255, 0.3); color: white'
        elif isinstance(val, str) and val.replace('.','',1).isdigit():  # Número de día
            return 'background-color: #f5f5f5; color: #333; font-weight: bold'
        return ''
    
    styled_df = df.style.applymap(color_celda, subset=['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'])
    
    return styled_df, counts

# Interfaz de usuario
st.title('Calendario de Envío de Reportes')

archivo_subido = st.file_uploader('Sube tu archivo CSV o Excel', type=['csv', 'xlsx'])

if archivo_subido:
    df = load_data(archivo_subido)

    st.subheader('Datos Cargados (con fechas ajustadas a días laborables)')
    st.write(df)

    if not df.empty:
        calendario_df = generate_calendar(df)

        st.subheader('Calendario de Entregas')
        if not calendario_df.empty:
            calendario_estilizado, conteos = style_calendar(calendario_df)
            st.dataframe(calendario_estilizado, height=800)
            
            st.subheader('Resumen por Día de la Semana')
            st.bar_chart(conteos.sum().rename('Total Entregas'))
        else:
            st.warning('No hay datos para mostrar en el calendario.')
    else:
        st.warning('El archivo no contiene datos válidos.')
