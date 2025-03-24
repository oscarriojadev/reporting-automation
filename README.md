# Revenue Report Generator

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
