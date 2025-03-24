# reporting-scheduling
A script to visualize Report Delivery Dates, identifying overlaps and adjusting dates that fall on weekends by moving them to Friday.

    #!/usr/bin/env python
    # -*- coding: utf-8 -*-
    
    import streamlit as st
    import pandas as pd
    import numpy as np
    import calendar
    from datetime import datetime, timedelta
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    import matplotlib.colors as mcolors
    
    DATE_FORMAT = '%Y-%m-%d'
    
    def adjust_to_friday(date):
        """
        Adjust dates falling on Saturday or Sunday to the preceding Friday.
        """
        # Convert to datetime if it's a string
        if isinstance(date, str):
            date = datetime.strptime(date, DATE_FORMAT).date()
        
        # Get the day of the week (0 = Monday, 6 = Sunday)
        day_of_week = date.weekday()
        
        # If Saturday (5) or Sunday (6), move back to Friday (4)
        if day_of_week == 5:  # Saturday
            return date - timedelta(days=1)
        elif day_of_week == 6:  # Sunday
            return date - timedelta(days=2)
        
        return date
    
    def get_color_intensity(count, max_count):
        """
        Generate a color with intensity based on the count of reports.
        Uses a blue color palette with increasing intensity.
        """
        if count == 0:
            return 'white'
        
        # Normalize the count to a value between 0 and 1
        normalized_count = count / max_count if max_count > 0 else 0
        
        # Create a color map from light blue to dark blue
        cmap = mcolors.LinearSegmentedColormap.from_list('blue_gradient', ['#E6F2FF', '#00008B'])
        
        # Convert the color to RGBA for use in Plotly
        color = cmap(normalized_count)
        return f'rgba({int(color[0]*255)}, {int(color[1]*255)}, {int(color[2]*255)}, {color[3]})'
    
    def create_calendar_final(df, date_column, title):
        """Crea un calendario completo funcional con todas las semanas."""
        try:
            # Validación de datos
            if df.empty or date_column not in df.columns:
                st.warning("Datos insuficientes para generar el calendario")
                return None, None
                
            df[date_column] = pd.to_datetime(df[date_column], errors='coerce')
            df = df.dropna(subset=[date_column]).copy()
            
            # Adjust dates falling on weekends to Friday
            df[date_column] = df[date_column].apply(adjust_to_friday)
            
            if df.empty:
                st.warning("No hay fechas válidas para visualizar")
                return None, None
    
            # Preparar datos con información de universidades por fecha
            universidad_by_date = df.groupby(pd.to_datetime(df[date_column]).dt.date).apply(
                lambda x: {
                    'count': len(x),
                    'universidades': list(x['Universidad'].unique())
                }
            ).to_dict()
    
            # Encontrar el máximo número de informes en un día para la escala de color
            max_count = max([info['count'] for info in universidad_by_date.values()] or [0])
    
            # Obtener rango de fechas
            min_date, max_date = df[date_column].min(), df[date_column].max()
            start_month = datetime(min_date.year, min_date.month, 1)
            end_month = (datetime(max_date.year + 1, 1, 1) if max_date.month == 12 
                       else datetime(max_date.year, max_date.month + 1, 1)) - timedelta(days=1)
            
            months_to_display = (end_month.year - start_month.year) * 12 + end_month.month - start_month.month + 1
            
            # Crear figura con subplots tipo 'domain' para tablas
            fig = make_subplots(
                rows=months_to_display, 
                cols=1,
                subplot_titles=[f"{calendar.month_name[(start_month.month + i) % 12 or 12]} {start_month.year + (start_month.month + i - 1) // 12}" 
                              for i in range(months_to_display)],
                vertical_spacing=0.05,
                specs=[[{"type": "domain"}] for _ in range(months_to_display)]
            )
            
            current_month = start_month
            
            for month_idx in range(months_to_display):
                year, month = current_month.year, current_month.month
                cal = calendar.monthcalendar(year, month)
                
                # Preparar datos para el mes
                days = []
                fill_colors = []
                line_colors = []
                font_colors = []
                
                for week in cal:
                    for day in week:
                        if day == 0:
                            # Días fuera del mes
                            days.append("")
                            fill_colors.append('rgba(0,0,0,0)')
                            line_colors.append('rgba(0,0,0,0)')
                            font_colors.append('black')
                        else:
                            date_str = f"{year}-{month:02d}-{day:02d}"
                            date_obj = datetime.strptime(date_str, DATE_FORMAT).date()
                            
                            # Obtener información de la fecha
                            date_info = universidad_by_date.get(date_obj, {'count': 0, 'universidades': []})
                            count = date_info['count']
                            
                            # Formatear texto
                            day_text = f"{day}"
                            
                            if count > 0:
                                day_text += f" ({count})"
                                
                            days.append(day_text)
                            
                            # Determinar color de fondo y fuente
                            if count > 0:
                                # Fechas con entregas
                                fill_color = get_color_intensity(count, max_count)
                                fill_colors.append(fill_color)
                                line_colors.append('#00008B')  # Azul oscuro
                                font_colors.append('white')
                            else:
                                # Fechas sin entregas
                                fill_colors.append('#F0F0F0')  # Gris clarito
                                line_colors.append('lightgray')
                                font_colors.append('black')
                
                # Organizar datos por columnas (días de la semana)
                columns = [[], [], [], [], [], [], []]
                colors_cols = [[], [], [], [], [], [], []]
                borders_cols = [[], [], [], [], [], [], []]
                font_cols = [[], [], [], [], [], [], []]
                
                for i in range(len(days)):
                    col_idx = i % 7
                    columns[col_idx].append(days[i])
                    colors_cols[col_idx].append(fill_colors[i])
                    borders_cols[col_idx].append(line_colors[i])
                    font_cols[col_idx].append(font_colors[i])
                
                # Crear tabla para el mes
                table = go.Table(
                    header=dict(
                        values=['Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb', 'Dom'],
                        fill_color='royalblue',
                        font=dict(color='white', size=12),
                        align='center'
                    ),
                    cells=dict(
                        values=columns,
                        fill_color=colors_cols,
                        line_color=borders_cols,
                        font=dict(
                            size=12, 
                            color=font_cols  # Usar colores de fuente diferenciados
                        ),
                        height=30,
                        align='center'
                    )
                )
                
                fig.add_trace(table, row=month_idx+1, col=1)
                
                # Avanzar al próximo mes
                if month == 12:
                    current_month = datetime(year + 1, 1, 1)
                else:
                    current_month = datetime(year, month + 1, 1)
    
            # Configuración final
            fig.update_layout(
                title_text=title,
                height=150 * months_to_display,
                margin=dict(l=20, r=20, t=100, b=20),
                showlegend=False
            )
            
            return fig, universidad_by_date
        
        except Exception as e:
            st.error(f"Error al crear el calendario: {str(e)}")
            return None, None
    
    # Rest of the code remains the same as in the previous version
    def main():
        st.set_page_config(page_title="Calendario Definitivo", layout="wide")
        st.title("📅 Calendario Completo de Entregas")
        
        # Carga de datos
        uploaded_file = st.file_uploader("Sube tu archivo (Excel o CSV)", type=["xlsx", "csv"])
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file) if uploaded_file.name.endswith('.xlsx') else pd.read_csv(uploaded_file)
                st.success("✅ Datos cargados correctamente")
                
                st.subheader("Vista previa de datos")
                st.dataframe(df.head())
                
                date_col = st.selectbox("Selecciona la columna de fecha", df.columns)
                
                if st.button("Generar Calendario", type="primary"):
                    with st.spinner("Creando visualización..."):
                        fig, universidad_by_date = create_calendar_final(df, date_col, "Calendario de Entregas")
                        if fig:
                            # Mostrar calendario
                            st.plotly_chart(fig, use_container_width=True)
                            
                            # Mostrar detalles de universidades por fecha
                            st.subheader("Detalles de Entregas")
                            details_df = pd.DataFrame([
                                {
                                    'Fecha Original': str(date),
                                    'Fecha Ajustada': str(date),
                                    'Número de Entregas': info['count'],
                                    'Universidades': ', '.join(info['universidades'])
                                }
                                for date, info in universidad_by_date.items() if info['count'] > 0
                            ]).sort_values('Fecha Ajustada')
                            
                            st.dataframe(details_df)
            
            except Exception as e:
                st.error(f"Error: {str(e)}")
        
        # Ejemplo con datos
        if st.button("Usar datos de ejemplo"):
            sample_data = {
                'Fecha': [
                    '2023-01-01', '2023-01-07', '2023-01-14', '2023-01-21', '2023-01-28',
                    '2023-02-04', '2023-02-11', '2023-02-18', '2023-02-25',
                    '2023-03-04', '2023-03-11', '2023-03-18', '2023-03-25'
                ],
                'Universidad': [f"U{i%3+1}" for i in range(13)],
                'Batch': [f"B{i+1}" for i in range(13)],
                'Enviado': ['X' if i%2==0 else '' for i in range(13)]
            }
            df = pd.DataFrame(sample_data)
            
            st.subheader("Datos de ejemplo")
            st.dataframe(df)
            
            with st.spinner("Generando calendario..."):
                fig, universidad_by_date = create_calendar_final(df, 'Fecha', "Ejemplo de Calendario")
                if fig:
                    # Mostrar calendario
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Mostrar detalles de universidades por fecha
                    st.subheader("Detalles de Entregas")
                    details_df = pd.DataFrame([
                        {
                            'Fecha Original': str(date),
                            'Fecha Ajustada': str(date),
                            'Número de Entregas': info['count'],
                            'Universidades': ', '.join(info['universidades'])
                        }
                        for date, info in universidad_by_date.items() if info['count'] > 0
                    ]).sort_values('Fecha Ajustada')
                    
                    st.dataframe(details_df)
    
    if __name__ == "__main__":
        main()
