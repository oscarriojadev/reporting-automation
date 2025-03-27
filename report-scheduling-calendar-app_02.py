import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar

# Set page layout to wide mode - MUST BE FIRST STREAMLIT COMMAND
st.set_page_config(layout="wide")

class UniversityNames:
    @classmethod
    def get_short_name(cls, long_name):
        mapping = {
            'GA - MIT xPRO': 'xPRO',
            'GA - The University of Chicago': 'UCPE',
            'GA - Chicago Booth Executive Education': 'Booth',
            'GA - MIT Professional Education': 'MITPE',
            'GA - Miami Herbert Business School': 'UMiami',
            'GA - University of Chicago Graham School': 'Graham',
            'GA - University of Pennsylvania': 'UPenn',
            'GA - Stanford Center for Professional Development': 'STF',
            'MIT xPRO': 'xPRO',
            'The University of Chicago': 'UCPE',
            'Chicago Booth Executive Education': 'Booth',
            'MIT Professional Education': 'MITPE',
            'Miami Herbert Business School': 'UMiami',
            'University of Chicago Graham School': 'Graham',
            'University of Pennsylvania': 'UPenn',
            'Stanford Center for Professional Development': 'STF'
        }
        return mapping.get(long_name, long_name)

def adjust_weekend_date(date):
    if pd.isna(date):
        return date
    if date.weekday() == 5:  # Saturday
        return date - timedelta(days=1)
    elif date.weekday() == 6:  # Sunday
        return date - timedelta(days=2)
    return date

def load_and_process_data(file):
    try:
        if file.name.endswith('.csv'):
            df = pd.read_csv(file)
        elif file.name.endswith('.xlsx'):
            df = pd.read_excel(file)
        else:
            st.error("Unsupported file format")
            return None
        
        for col in df.columns:
            if 'date' in col.lower() or 'fecha' in col.lower():
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].apply(adjust_weekend_date)
        
        if 'Universidad' in df.columns:
            df['Universidad'] = df['Universidad'].apply(
                lambda x: UniversityNames.get_short_name(x) if pd.notna(x) else x
            )
        
        if 'Batch' in df.columns:
            df['Batch'] = pd.to_datetime(df['Batch'], format='%Y-%m', errors='coerce').dt.strftime('%b %Y')
        
        df['Universidad'] = df['Universidad'].fillna('').astype(str)
        df['Batch'] = df['Batch'].fillna('').astype(str)
        
        # Create combined University-Batch column
        df['UniversityBatch'] = df['Universidad'] + ' (' + df['Batch'] + ')'
        
        return df
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return None

def eliminar_duplicados_universidad_batch(df):
    return df.drop_duplicates(subset=['Universidad', 'Batch'])

def generate_calendar(df):
    if 'Fecha_de_Envio' not in df.columns:
        st.error("DataFrame doesn't contain 'Fecha_de_Envio' column")
        return pd.DataFrame()
    
    try:
        min_date = df['Fecha_de_Envio'].min()
        max_date = df['Fecha_de_Envio'].max()
        all_dates = pd.date_range(start=min_date, end=max_date)

        calendar_df = pd.DataFrame({
            'Date': all_dates,
            'Week': all_dates.strftime('%U'),
            'Weekday': all_dates.strftime('%A'),
            'WeekdayNum': all_dates.weekday,
            'Month': all_dates.strftime('%B'),
            'Year': all_dates.strftime('%Y'),
            'Day': all_dates.day
        })

        calendar_df = calendar_df[calendar_df['Weekday'].isin(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'])]
        
        merged_df = pd.merge(
            calendar_df, 
            df[['Fecha_de_Envio', 'UniversityBatch']],
            left_on='Date', 
            right_on='Fecha_de_Envio', 
            how='left'
        )
        
        # Fill empty cells with day number
        merged_df['UniversityBatch'] = merged_df['UniversityBatch'].fillna(merged_df['Day'].astype(str))
        
        # Create a dictionary to store combined week data
        week_data = {}
        
        for _, row in merged_df.iterrows():
            week_key = (row['Year'], row['Week'])
            if week_key not in week_data:
                # Format date as DD-MM-YYYY
                formatted_date = row['Date'].strftime('%d-%m-%Y')
                week_data[week_key] = {
                    'Year': row['Year'],
                    'Week': row['Week'],
                    'Month': row['Month'],
                    'DateWeek': f"{formatted_date} (Week {row['Week']})",
                    'Monday': None,
                    'Tuesday': None,
                    'Wednesday': None,
                    'Thursday': None,
                    'Friday': None
                }
            
            # Assign the value to the correct weekday
            week_data[week_key][row['Weekday']] = row['UniversityBatch']
        
        # Convert the dictionary to a DataFrame
        combined_df = pd.DataFrame.from_dict(week_data, orient='index')
        
        # Reset index and reorder columns
        combined_df = combined_df.reset_index(drop=True)
        combined_df = combined_df[['DateWeek', 'Month', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']]
        
        return combined_df
        
    except Exception as e:
        st.error(f"Error generating calendar: {str(e)}")
        return pd.DataFrame()

def style_calendar(df):
    if df.empty:
        return df
    
    styled_df = df.reset_index(drop=True)
    
    def color_delivery(val):
        if pd.isna(val):
            return ''
        if isinstance(val, str) and val.isdigit():
            return 'background-color: white; color: black'
        return 'background-color: #99CCFF; color: black'
    
    weekday_columns = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    available_weekdays = [col for col in weekday_columns if col in styled_df.columns]
    
    styled_df = styled_df.style.applymap(color_delivery, subset=available_weekdays)
    
    return styled_df

def reorder_calendar(df):
    if df.empty:
        return df
    
    try:
        # Extract date from DateWeek for sorting
        df['SortDate'] = pd.to_datetime(df['DateWeek'].str.extract(r'(\d{2}-\d{2}-\d{4})')[0], format='%d-%m-%Y')
        
        # Sort by date
        df = df.sort_values('SortDate')
        
        # Drop temporary column
        df = df.drop(columns=['SortDate'])
        
        return df
    except Exception as e:
        st.error(f"Error reordering calendar: {str(e)}")
        return df

def main():
    st.title("Report Delivery Calendar")
    
    uploaded_file = st.file_uploader("Upload file (CSV or Excel)", type=['csv', 'xlsx'])
    
    if uploaded_file is not None:
        df = load_and_process_data(uploaded_file)
        
        if df is not None:
            df_unique = eliminar_duplicados_universidad_batch(df)
            
            calendar_df = generate_calendar(df_unique)
            if not calendar_df.empty:
                calendar_reordered = reorder_calendar(calendar_df)
                styled_calendar = style_calendar(calendar_reordered)
                
                st.subheader("Delivery Calendar")
                
                # Custom CSS to adjust column widths
                st.markdown(
                    """
                    <style>
                    .dataframe th, .dataframe td {
                        white-space: nowrap;
                        text-align: left;
                        padding: 8px;
                    }
                    .dataframe thead th {
                        position: sticky;
                        top: 0;
                        background-color: white;
                    }
                    .dataframe-container {
                        overflow: auto;
                        width: 100%;
                    }
                    </style>
                    """,
                    unsafe_allow_html=True
                )
                
                # Display the dataframe with adjusted width
                st.dataframe(styled_calendar, height=600, use_container_width=True)

if __name__ == "__main__":
    main()