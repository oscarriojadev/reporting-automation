# reporting-merger
        """
               ____                          ____  _         _          ____
              / __ \______________ ______   / __ \(_)___    (_)___ _   / __ \___ _   __
             / / / / ___/ ___/ __  / ___/  / /_/ / / __ \  / / __  /  / / / / _ \ | / /
            / /_/ (__  ) /__/ /_/ / /     / _, _/ / /_/ / / / /_/ /  / /_/ /  __/ |/ /
            \____/____/\___/\__,_/_/     /_/ |_/_/\____/_/ /\__,_/  /_____/\___/|___/
                                                      /___/
        OSCAR RIOJA DEV
        """  
        import os
        import pandas as pd
        import openpyxl
        import traceback  # Para obtener información detallada de los errores
        from datetime import datetime
        
        # Definir la carpeta donde están los archivos
        folder_path = '/mnt/c/Users/oscar.rioja/OneDrive - Global Alumni/Documentos/Finanzas/Consolidation of Historial Revenue Reports/Facturación_2024_Todos los Revenue Reports Histórico/Todos los Revenue Reports Histórico/All Schools'
        
        # Inicializar una lista para almacenar los datos combinados
        combined_data = []
        excel_summary = []  # Para almacenar el resumen de archivos y pestañas procesados
        
        # Lista de nombres exactos de los archivos que se deben procesar
        exact_files = [
            'Booth_May 2020_Revenue Report.xlsx',
            'Booth_September 2020_Revenue Report.xlsx',
            'Booth_February 2021_Revenue Report.xlsx',
            'Booth_May 2021_Revenue Report.xlsx',
            'Booth_September 2021_Revenue Report.xlsx',
            'Booth_May 2022_Revenue Report.xlsx',
            'Booth_June 2022_Revenue Report.xlsx',
            'Booth_September 2022_Revenue Report.xlsx',
            'Booth_October 2022_Revenue Report.xlsx',
            'Booth_November 2022_Revenue Report.xlsx',
            'Booth_February 2023_Revenue Report.xlsx',
            'Booth_April 2023_Revenue Report.xlsx',
            'Booth_September 2023_Revenue Report.xlsx',
            'Booth_November 2023_Revenue Report.xlsx',
            'Booth_February 2024_Revenue Report.xlsx',
            'Booth_April 2024_Revenue Report.xlsx',
            'Booth_September 2024_Revenue Report.xlsx'
        ]
        
        
        # Función para listar todos los archivos en folder_path
        def list_all_files_in_folder(folder_path):
            try:
                files_in_folder = os.listdir(folder_path)
                if files_in_folder:
                    print("Archivos encontrados en la carpeta:")
                    for file in files_in_folder:
                        print(f"- {file}")
                else:
                    print("No se encontraron archivos en la carpeta.")
            except Exception as e:
                print(f"Error al intentar listar archivos en {folder_path}. Motivo: {e}")
        
        # Función para procesar pestañas de cualquier archivo Excel
        def process_excel_file(file_path, file_name):
            try:
                excel_file = pd.ExcelFile(file_path)
        
                # Extraer el mes y año del nombre del archivo
                # Por ejemplo, de 'Booth_February 2021_Revenue Report.xlsx' obtiene 'February 2021'
                month_year = file_name.split('_')[1]
        
                # Recorrer las pestañas del archivo
                for sheet_name in excel_file.sheet_names:
                    print(f"Evaluando la pestaña: {sheet_name} del archivo {file_name}")
        
                    # Excluir las pestañas que contengan 'Summary' o 'Total' en el nombre
                    if 'summary' in sheet_name.lower() or 'total' in sheet_name.lower():
                        continue  # No procesar estas pestañas
        
                    try:
                        data = None  # Inicializar 'data' como None
        
                        # Procesar solo si la pestaña no es 'Summary' ni 'Total'
                        data = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=5, usecols="D:J")
                        data.rename(columns={
                            'Full Name': 'Student',
                            'Email': 'email',
                            'Status': 'Status',
                            'Tuition Fee': 'Tuition Fee',
                            'Discount': 'Discount',
                            'Refund': 'Refund'
                        }, inplace=True)
                        data['Transfer'] = pd.NA
                        data['Course'] = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1).iloc[0, 2]
                        data['Program'] = sheet_name
                        # Añadir la columna Month_Year
                        data['Month_Year'] = month_year
        
                        # Solo agregar a combined_data si data fue definido
                        if data is not None:
                            combined_data.append(data)
        
                        # Guardar el archivo y las pestañas procesadas
                        excel_summary.append({'Archivo': file_name, 'Pestaña': sheet_name})
        
                    except Exception as e:
                        print(f"Error procesando la pestaña {sheet_name} del archivo {file_name}: {e}")
                        print(f"Línea exacta del error: {traceback.format_exc()}")
        
            except Exception as e:
                print(f"Error al abrir el archivo {file_name}: {e}")
                print(f"Línea exacta del error: {traceback.format_exc()}")
        
        
        # Listar archivos en la carpeta
        list_all_files_in_folder(folder_path)
        
        # Procesar cada archivo si está en la lista de exact_files
        for file_name in os.listdir(folder_path):
            if file_name in exact_files:
                print(f"Procesando archivo: {file_name}")
                file_path = os.path.join(folder_path, file_name)
                process_excel_file(file_path, file_name)
        
        # Lista de columnas que se deben eliminar del DataFrame final
        columns_to_remove = [
            'Program Name', 'Type', 'Enrolled Amount', 'Rev. Share', 'Academic Program',
            'Unnamed: 2', 'Unnamed: 3', 'Cohort', 'Rev Share', 'Nombre de la oportunidad',
            'Estado admisión/inscripción', 'Estado pagos', 'Inscrito convocatoria',
            'Universidad', 'Programa académico', 'Nueva convocatoria', 'Name', 'Mail',
            'TOTAL AMOUNT', 'Amount from September', 'Amount From November', 'Amount Due',
            'Language', 'Stage', 'Rev. Share Umiami', 'Certificate', 'Mes inicio certificate',
            'Empresa', 'Pagos EM', 'CoC - SENT'
        ]
        
        # Combinar todos los datos en un solo DataFrame
        if combined_data:
            combined_df = pd.concat(combined_data, ignore_index=True)
        
            # Eliminar las columnas especificadas si existen en el DataFrame
            combined_df = combined_df.drop(columns=[col for col in columns_to_remove if col in combined_df.columns])
        
            # Guardar el resultado en un archivo Excel
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file_path = os.path.join(folder_path, f'Booth_HRR_2020-2024_{timestamp}.xlsx') 
            combined_df.to_excel(output_file_path, index=False)
            print(f"Datos combinados guardados en {output_file_path}")
        else:
            print("No se encontraron datos para combinar.")
        
        # Guardar el resumen de archivos y pestañas procesadas en un archivo Excel
        summary_df = pd.DataFrame(excel_summary)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        summary_output_file_path = os.path.join(folder_path, f'Booth_HRR_2020-2024_ProcessedTabs_{timestamp}.xlsx') 
        summary_df.to_excel(summary_output_file_path, index=False)
        print(f"Resumen guardado en {summary_output_file_path}")
