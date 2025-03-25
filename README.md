# reporting-scheduling
A script to visualize Report Delivery Dates, identifying overlaps and adjusting dates that fall on weekends by moving them to Friday.

    import pandas as pd
    import streamlit as st
    
    # Función para obtener universidades y fechas por semana
    def get_universities_for_week(week, df):
        reports_in_week = df[df['Week'] == week]
        universities_in_week = reports_in_week[['report_batch_university', 'Fecha_de_Envio']].drop_duplicates()
        universities_list = [f"{row['report_batch_university']} ({row['Fecha_de_Envio'].strftime('%d-%m-%Y')})" for _, row in universities_in_week.iterrows()]
        return ', '.join(universities_list)
    
    # Función para ajustar las fechas de envío (pasar sábado/domingo al viernes anterior)
    def adjust_weekend_dates(date):
        if date.weekday() == 5:  # Sábado
            return date - pd.Timedelta(days=1)
        elif date.weekday() == 6:  # Domingo
            return date - pd.Timedelta(days=2)
        else:
            return date
    
    # Función para generar el calendario
    def generate_calendar(df):
        # Asegurarse de que la columna 'Fecha_de_Envio' existe
        if 'Fecha_de_Envio' not in df.columns:
            st.error("La columna 'Fecha_de_Envio' no existe en los datos.")
            return pd.DataFrame()
    
        # Crear la columna 'report_batch_university' concatenando 'Report Batch' y 'Universidad', convirtiendo a str
        if 'Report Batch' in df.columns and 'Universidad' in df.columns:
            df['report_batch_university'] = df['Report Batch'].astype(str) + ' - ' + df['Universidad'].astype(str)
        else:
            st.error("Las columnas 'Report Batch' y 'Universidad' no existen en los datos.")
            return pd.DataFrame()
    
        # Convertir 'Fecha_de_Envio' a tipo datetime
        df['Fecha_de_Envio'] = pd.to_datetime(df['Fecha_de_Envio'], errors='coerce')
    
        # Ajustar fechas que caen en sábado o domingo
        df['Fecha_de_Envio'] = df['Fecha_de_Envio'].apply(adjust_weekend_dates)
        
        # Crear la columna 'Week' a partir de 'Fecha_de_Envio'
        df['Week'] = df['Fecha_de_Envio'].dt.strftime('%W - %Y')  # semana del año y año
    
        # Generar todas las semanas del año
        weeks_of_year = pd.date_range(start=f'{df["Fecha_de_Envio"].min().year}-01-01', 
                                      end=f'{df["Fecha_de_Envio"].max().year}-12-31', 
                                      freq='W-MON').strftime('%W - %Y')
    
        # Crear la tabla de calendario con las semanas del año
        calendar_df = pd.DataFrame({
            'Week': weeks_of_year,
            'Monday': [""]*len(weeks_of_year),
            'Tuesday': [""]*len(weeks_of_year),
            'Wednesday': [""]*len(weeks_of_year),
            'Thursday': [""]*len(weeks_of_year),
            'Friday': [""]*len(weeks_of_year),
            'Universities': [""]*len(weeks_of_year)
        })
    
        # Rellenar las columnas con el conteo y fecha de entregas
        for index, week in enumerate(calendar_df['Week']):
            reports_in_week = df[df['Week'] == week]
            if not reports_in_week.empty:
                for _, report in reports_in_week.iterrows():
                    day_name = report['Fecha_de_Envio'].strftime('%A')
                    day_number = report['Fecha_de_Envio'].strftime('%d')
    
                    # Rellenar solo los días de lunes a viernes
                    if day_name in ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']:
                        current_value = calendar_df.at[index, day_name]
                        if current_value:
                            new_value = f"{current_value}, {day_number} ({len(reports_in_week)})"
                        else:
                            new_value = f"{day_number} ({len(reports_in_week)})"
                        calendar_df.at[index, day_name] = new_value
    
                # Llenar el campo de "Universities" con el nombre y fecha de las universidades
                calendar_df.at[index, 'Universities'] = get_universities_for_week(week, df)
        
        return calendar_df
    
    # Función para mostrar la tabla con detalles en una ventana emergente
    def show_details_popup(df):
        if st.button('Ver más detalles'):
            with st.expander("Detalles del calendario"):
                st.write(df)
    
    # Subir archivo Excel
    uploaded_file = st.file_uploader("Sube un archivo Excel", type=["xlsx"])
    
    if uploaded_file is not None:
        # Leer el archivo Excel
        df = pd.read_excel(uploaded_file)
    
        # Asegurarse de que tiene las columnas necesarias
        if 'Fecha_de_Envio' in df.columns and 'Report Batch' in df.columns and 'Universidad' in df.columns:
            # Generar el calendario
            calendar_df = generate_calendar(df)
    
            # Mostrar la tabla con conteos de entregas por semana
            st.write("Calendario de Entregas")
            st.dataframe(calendar_df)
    
            # Mostrar detalles adicionales en ventana emergente
            show_details_popup(calendar_df)
        else:
            st.error("El archivo debe contener las columnas 'Fecha_de_Envio', 'Report Batch' y 'Universidad'.")
    else:
        st.info("Por favor, sube un archivo Excel para generar el calendario.")
