# reporting-merger
  
    import os
    import pandas as pd
    import traceback
    from datetime import datetime
    
    #Define the folder path where the files are located
    
    folder_path = (
        '/mnt/c/Users/oscar.rioja/OneDrive - Global Alumni/'
        'Documents/Finance/Consolidation of Historical Revenue Reports/'
        'Billing_2024_All Historical Revenue Reports/All Historical Revenue Reports/'
        'All Schools'
    )
    
    #Initialize a list to store combined data
    combined_data = []
    excel_summary = []  # To store the summary of processed files and sheets
    
    #List of exact file names to be processed
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
    
    def list_all_files_in_folder(folder_path):
        """List all files in the specified folder."""
        try:
            files_in_folder = os.listdir(folder_path)
            if files_in_folder:
                print("Files found in the folder:")
                for file in files_in_folder:
                    print(f"- {file}")
            else:
                print("No files found in the folder.")
        except Exception as e:
            print(f"Error listing files in {folder_path}. Reason: {e}")
    
    def process_excel_file(file_path, file_name):
        """Process sheets of any Excel file."""
        try:
            excel_file = pd.ExcelFile(file_path)
    
            #Extract month and year from the file name
            month_year = file_name.split('_')[1]
    
            #Iterate through the sheets in the file
            for sheet_name in excel_file.sheet_names:
                print(f"Evaluating sheet: {sheet_name} in file {file_name}")
    
                #Exclude sheets containing 'Summary' or 'Total' in the name
                if 'summary' in sheet_name.lower() or 'total' in sheet_name.lower():
                    continue  # Skip these sheets
    
                try:
                    data = pd.read_excel(
                        file_path, sheet_name=sheet_name, skiprows=5, usecols="D:J"
                    )
                    data.rename(columns={
                        'Full Name': 'Student',
                        'Email': 'email',
                        'Status': 'Status',
                        'Tuition Fee': 'Tuition Fee',
                        'Discount': 'Discount',
                        'Refund': 'Refund'
                    }, inplace=True)
                    data['Transfer'] = pd.NA
                    data['Course'] = pd.read_excel(
                        file_path, sheet_name=sheet_name, nrows=1
                    ).iloc[0, 2]
                    data['Program'] = sheet_name
                    data['Month_Year'] = month_year
    
                    if data is not None:
                        combined_data.append(data)
    
                    # Save the processed file and sheets
                    excel_summary.append({'File': file_name, 'Sheet': sheet_name})
    
                except Exception as e:
                    print(f"Error processing sheet {sheet_name} in file {file_name}: {e}")
                    print(f"Exact error line: {traceback.format_exc()}")
    
        except Exception as e:
            print(f"Error opening file {file_name}: {e}")
            print(f"Exact error line: {traceback.format_exc()}")
    
    #List files in the folder
    list_all_files_in_folder(folder_path)
    
    #Process each file if it is in the exact_files list
    for file_name in os.listdir(folder_path):
        if file_name in exact_files:
            print(f"Processing file: {file_name}")
            file_path = os.path.join(folder_path, file_name)
            process_excel_file(file_path, file_name)
    
    #List of columns to be removed from the final DataFrame
    columns_to_remove = [
        'Program Name', 'Type', 'Enrolled Amount', 'Rev. Share', 'Academic Program',
        'Unnamed: 2', 'Unnamed: 3', 'Cohort', 'Rev Share', 'Opportunity Name',
        'Admission/Enrollment Status', 'Payment Status', 'Enrolled in Call',
        'University', 'Academic Program', 'New Call', 'Name', 'Mail',
        'TOTAL AMOUNT', 'Amount from September', 'Amount From November', 'Amount Due',
        'Language', 'Stage', 'Rev. Share Umiami', 'Certificate', 'Certificate Start Month',
        'Company', 'EM Payments', 'CoC - SENT'
    ]
    
    #Combine all data into a single DataFrame
    if combined_data:
        combined_df = pd.concat(combined_data, ignore_index=True)
    
        #Remove the specified columns if they exist in the DataFrame
        combined_df = combined_df.drop(
            columns=[col for col in columns_to_remove if col in combined_df.columns]
        )
    
        #Save the result to an Excel file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file_path = os.path.join(
            folder_path, f'Booth_HRR_2020-2024_{timestamp}.xlsx'
        )
        combined_df.to_excel(output_file_path, index=False)
        print(f"Combined data saved to {output_file_path}")
    else:
        print("No data found to combine.")
    
    #Save the summary of processed files and sheets to an Excel file
    summary_df = pd.DataFrame(excel_summary)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    summary_output_file_path = os.path.join(
        folder_path, f'Booth_HRR_2020-2024_ProcessedTabs_{timestamp}.xlsx'
    )
    summary_df.to_excel(summary_output_file_path, index=False)
    print(f"Summary saved to {summary_output_file_path}")
