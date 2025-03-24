# Revenue Report Generator
## Overview

This script generates comprehensive revenue reports for educational programs. It processes participant data from Excel files, creates detailed financial analysis, and outputs formatted reports with course-specific information and summary statistics.

## Features

- Processes participant data from Excel files
- Calculates key financial metrics including:
  - Gross revenue
  - Discounts
  - Total fees collected
  - Refunds
  - Net revenue
  - University share (40%)
- Creates separate sheets for:
  - Overall participant data
  - Course-specific data
  - Status-specific information (Enrolled, Refund, Transfer In/Out)
  - Summary statistics
- Applies professional formatting with Trebuchet MS font and color-coded headers
- Automatically adjusts column widths for readability

## Usage

1. Ensure the input Excel file follows the expected format with the required sheet name
2. Set the input file path variable `file_path` to point to your participant data file
3. Run the script to generate the formatted revenue report

```python
# Example usage with default settings
file_path = '/path/to/Booth_December 2024_Participants List.xlsx'
sheet_name = 'Listado General (DATA)'
```

## Required Input Columns

The input Excel file should contain the following columns:
- Full Name
- Email
- Status
- LIST PRICE - DISCOUNTS (USD)
- Total Discounts USD
- GA Course Code
- Booth Course Code
- Academic Program

## Output

The script generates an Excel file with the following sheets:
- Participants: Complete participant data
- TOTAL: Summary of all financial metrics
- Course-specific sheets: One sheet per course with detailed participant information
- Summary sheets: Financial analysis for each course

## Configuration Options

The university share percentage is set to 40% by default:
```python
total_university_share_percentage = 0.40  # Set the university share percentage
```

## Formatting

The output Excel file includes the following formatting:
- Headers with dark blue background and white text
- All text in Trebuchet MS font
- Bold formatting for summary rows
- Auto-adjusted column widths
- Table formatting with wrap text
- Auto-filter for easy data sorting

## Alternative Configuration

An alternative configuration is included as a comment at the end of the script, which provides more granular control over the report generation process, including:
- Configurable base paths
- University-specific naming patterns
- Additional sheet types and ordering
- Enhanced formatting options

```
   ____                          ____  _         _          ____
  / __ \______________ ______   / __ \(_)___    (_)___ _   / __ \___ _   __
 / / / / ___/ ___/ __  / ___/  / /_/ / / __ \  / / __  /  / / / / _ \ | / /
/ /_/ (__  ) /__/ /_/ / /     / _, _/ / /_/ / / / /_/ /  / /_/ /  __/ |/ /
\____/____/\___/\__,_/_/     /_/ |_/_/\____/_/ /\__,_/  /_____/\___/|___/
                                          /___/
OSCAR RIOJA DEV
Last Modified: 17-Feb-2024
```
      
      import pandas as pd
      import numpy as np
      import os
      from openpyxl import Workbook
      from openpyxl.styles import Font, PatternFill, numbers
      import re
      
      # Define the input file path and sheet name
      file_path = '/mnt/c/Users/oscar.rioja/OneDrive - Global Alumni/Revenue Reports/Booth/Dec 24/Working Files/Booth_December 2024_Participants List.xlsx'
      sheet_name = 'Listado General (DATA)'
      
      # Load the TOTAL PARTICIPANTS table
      df = pd.read_excel(file_path, sheet_name=sheet_name)
      
      # Define the output file path
      output_file_path = os.path.join(os.path.dirname(file_path), 'Booth_December 2024_Revenue Report_OR.xlsx')
      
      # Create the 'Student' field by concatenating 'Name' and 'Last Name'
      df['Student'] = df['Name'] + ' ' + df['Last Name']
      
      # Select and rename the relevant columns to match the output table requirements
      program_df = df[['Full Name', 'Email', 'Status', 'LIST PRICE - DISCOUNTS (USD)', 
                       'Total Discounts USD', 'GA Course Code', 'Booth Course Code', 'Academic Program']].copy()
      
      # Rename the columns to match the new table format
      program_df.columns = ['Student', 'Email', 'Status', 'Tuition Fee', 'Discount', 
                           'Course', 'Booth Course Code', 'Program']
      
      
      # Add the 'Refund' column
      program_df['Refund'] = np.where(program_df['Status'] == 'Refund', -program_df['Tuition Fee'], np.nan)
      
      # Define the total summary variables
      total_enrollments = total_gross_revenue = total_discounts = total_fees_collected = total_refunds = total_net_revenue = 0
      total_transfers = total_transfer_amount = total_active_participants = total_final_fees = total_university_share = 0
      total_number_of_refunds = total_number_of_transfers = 0
      total_university_share_percentage = 0.40  # Set the university share percentage
      
      # Function to sanitize sheet names
      def sanitize_sheet_name(name):
          # Replace invalid characters with an underscore
          return re.sub(r'[:\\/?*\[\]]', '_', name)[:31]
      
      # Create an Excel writer object
      with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
          # Create the "Participants" sheet with the original data
          df.to_excel(writer, sheet_name='Participants', index=False)
          print(f"Saved participants data to 'Participants' sheet.")
      
          # Calculate cumulative totals for the "TOTAL" sheet
          for course in program_df['Course'].unique():
              # Convert the course to a string to avoid float issues
              course = str(course)
      
              # Filter the DataFrame for the current course
              course_table = program_df[program_df['Course'] == course]
      
              # Check if the course_table is empty
              if course_table.empty:
                  print(f"Skipping empty course table for course: {course}")
                  continue
      
              # Calculate summary data for the current course
              course_retail_price = course_table[course_table['Discount'] == 0]['Tuition Fee'].max()
              course_type = "Open"
              number_of_enrollments = len(course_table)
              gross_revenue = (course_table['Tuition Fee'] + course_table['Discount']).sum()
              discounts = course_table['Discount'].sum()
              total_fees_collected_course = gross_revenue - discounts
              number_of_refunds = len(course_table[course_table['Status'] == 'Refund'])
              refunds = course_table[course_table['Status'] == 'Refund']['Tuition Fee'].sum()
              total_participants = number_of_enrollments - number_of_refunds
              net_revenue = total_fees_collected_course - refunds
              number_of_transfers = len(course_table[course_table['Status'].str.startswith('Transfer Out') | course_table['Status'].str.startswith('Transfer out')])
              transfers = course_table[course_table['Status'].str.startswith('Transfer Out') | course_table['Status'].str.startswith('Transfer out')]['Tuition Fee'].sum()
              active_participants = total_participants - number_of_transfers
              final_fees_collected = net_revenue - transfers
              university_share = total_university_share_percentage * final_fees_collected
      
              # Update cumulative totals
              total_enrollments += number_of_enrollments
              total_gross_revenue += gross_revenue
              total_discounts += discounts
              total_fees_collected += total_fees_collected_course
              total_refunds += refunds
              total_transfers += transfers
              total_net_revenue += net_revenue
              total_active_participants += active_participants
              total_final_fees += final_fees_collected
              total_university_share += university_share
              total_number_of_refunds += number_of_refunds
              total_number_of_transfers += number_of_transfers
      
          # Now, create the "TOTAL" summary sheet with cumulative data
          total_summary_data = {
              'Category': [
                  'Number of enrollments', 'Gross revenue', 'Discounts', 'Total fees collected',
                  'Number of refunds', 'Refunds', 'Total Participants', 'Net revenue', 'Number of transfers', 'Transfers', 'Active participants',
                  'Final fees collected', '% University Share', 'University Share'
              ],
              'Value': [
                  total_enrollments, total_gross_revenue, total_discounts, total_fees_collected,
                  total_number_of_refunds, total_refunds, total_enrollments - total_number_of_refunds, total_net_revenue, total_number_of_transfers, total_transfers,
                  total_active_participants, total_final_fees, f"{total_university_share_percentage * 100:.2f}%", total_university_share
              ]
          }
      
          total_summary_df = pd.DataFrame(total_summary_data)
          total_summary_df.to_excel(writer, sheet_name='TOTAL', index=False)
          print(f"Saved total summary sheet.")
      
          # Iterate over each unique course and create a new sheet for each
          for course in program_df['Course'].unique():
              # Convert the course to a string to avoid float issues
              course = str(course)
      
              # Filter the DataFrame for the current course
              course_table = program_df[program_df['Course'] == course]
      
              # Check if the course_table is empty
              if course_table.empty:
                  print(f"Skipping empty course table for course: {course}")
                  continue
      
              # Use the 'Booth Course Code' for the sheet name if available, otherwise use 'Course'
              course_name = str(course_table['Course'].iloc[0]) if not course_table['Course'].empty else course
              course_name = sanitize_sheet_name(course_name)  # Sanitize the sheet name
      
      
              # Calculate summary data for the current course
              course_retail_price = course_table[course_table['Discount'] == 0]['Tuition Fee'].max()
              course_type = "Open"
              number_of_enrollments = len(course_table)
              gross_revenue = (course_table['Tuition Fee'] + course_table['Discount']).sum()
              discounts = course_table['Discount'].sum()
              total_fees_collected_course = gross_revenue - discounts
              number_of_refunds = len(course_table[course_table['Status'] == 'Refund'])
              refunds = course_table[course_table['Status'] == 'Refund']['Tuition Fee'].sum()
              total_participants = number_of_enrollments - number_of_refunds
              net_revenue = total_fees_collected_course - refunds
              number_of_transfers = len(course_table[course_table['Status'].str.startswith('Transfer Out') | course_table['Status'].str.startswith('Transfer out')])
              transfers = course_table[course_table['Status'].str.startswith('Transfer Out') | course_table['Status'].str.startswith('Transfer out')]['Tuition Fee'].sum()
              active_participants = total_participants - number_of_transfers
              final_fees_collected = net_revenue - transfers
              university_share = total_university_share_percentage * final_fees_collected
      
              # Create a summary DataFrame for the current course
              summary_data = {
                  'Cohort': [
                      'Course Name', 'Course Code', 'Course Retail Price', 'Course Type', 'Number of enrollments', 'Gross revenue',
                      'Discounts', 'Total fees collected', 'Number of refunds', 'Refunds',
                      'Total Participants', 'Net revenue', 'Number of transfers', 'Transfers', 'Active participants', 'Final fees collected',
                      '% University Share', 'University Share'
                  ],
                  'Details': [
                      course_table['Program'].iloc[0], course_table['Booth Course Code'].iloc[0], course_retail_price, course_type, number_of_enrollments, gross_revenue,
                      discounts, total_fees_collected_course, number_of_refunds, refunds,
                      total_participants, net_revenue, number_of_transfers, transfers, active_participants, final_fees_collected,
                      f"{total_university_share_percentage * 100:.2f}%", university_share
                  ]
              }
      
              # Convert summary data to a DataFrame
              summary_df = pd.DataFrame(summary_data)
      
              # Create a summary sheet for the course
              summary_sheet_name = f"Summary {sanitize_sheet_name(course)}"  # Sanitize the summary sheet name
              summary_df.to_excel(writer, sheet_name=summary_sheet_name, index=False)
              print(f"Saved summary for: {summary_sheet_name}")
      
              # Save each course table to a separate sheet
              course_table.to_excel(writer, sheet_name=course_name, index=False)
              print(f"Saved course table for: {course_name}")
      
          # Apply number formatting for the summary values (no decimal places where needed)
          workbook = writer.book
      
          # Apply formatting to all sheets
          for sheet_name in writer.sheets:
              sheet = writer.sheets[sheet_name]
      
              # Apply 'Trebuchet MS' font to all cells
              for row in sheet.iter_rows():
                  for cell in row:
                      cell.font = Font(name='Trebuchet MS')  # This line applies the Trebuchet MS font
      
              # Apply bold font, white color, and background color to the header row
              header_fill = PatternFill(start_color="44546A", end_color="44546A", fill_type="solid")
              for cell in sheet[1]:
                  cell.font = Font(bold=True, color="FFFFFF", name='Trebuchet MS')
                  cell.fill = header_fill
      
              # Adjust column width based on the content
              for col in sheet.columns:
                  max_length = 0
                  column = col[0].column_letter  # Get the column name
                  for cell in col:
                      try:
                          if len(str(cell.value)) > max_length:
                              max_length = len(cell.value)
                      except:
                          pass
                  adjusted_width = (max_length + 2)
                  sheet.column_dimensions[column].width = adjusted_width
      
              # Apply table formatting to the data range
              max_row = sheet.max_row
              max_col = sheet.max_column
              data_range = sheet[f'A1:{chr(65+max_col-1)}{max_row}']
              for row in data_range:
                  for cell in row:
                      cell.alignment = cell.alignment.copy(wrap_text=True)
      
              # Apply auto-filter and adjust column width
              table_range = f"A1:{chr(64+max_col)}{max_row}"
              sheet.auto_filter.ref = table_range
      
      print(f"Saved tables, summaries, and Total summary in the Excel file: {output_file_path}")
