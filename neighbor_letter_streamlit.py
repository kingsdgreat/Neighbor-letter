# Save this as neighbor_letter_streamlit.py

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

def main():
    st.title("Neighbor Letter Processor")

    # File upload
    input_file = st.file_uploader("Upload Input Excel File", type=['xlsx', 'xls'])

    # APN Value
    apn_value = st.text_input("APN Value")

    # Owner's First Name
    owner_first_name = st.text_input("Owner's First Name (optional)")

    # Owner's Last Name or Company Name
    owner_last_name = st.text_input("Owner's Last Name or Company Name")

    # Output file name
    output_file_name = st.text_input("Output File Name", value="output.xlsx")

    if st.button("Run"):
        if not input_file:
            st.error("Please upload an input file.")
            return
        if not apn_value.strip():
            st.error("APN value is required.")
            return
        if not owner_last_name.strip():
            st.error("Owner's last name or company name is required.")
            return

        # Process the spreadsheet
        try:
            # Read the uploaded Excel file into a DataFrame
            df = pd.read_excel(input_file)

            # Define the mapping from original column names to new column names
            column_mapping = {
                'Owner 1 First Name': 'First Name',
                'Owner 1 Last Name': 'Last Name',
                'Mailing Address': 'Mailing Address',
                'Mailing City': 'City',
                'Mailing State': 'State',
                'Mailing Zip': 'Zip',
                'County': 'Property County'
            }

            # Check if required columns exist
            missing_columns = set(column_mapping.keys()) - set(df.columns)
            if missing_columns:
                st.error(f"The following required columns are missing in the input file: {missing_columns}")
                return

            # Select the columns and rename them
            df_selected = df[list(column_mapping.keys())].rename(columns=column_mapping)

            # Remove entries matching the owner's name
            if owner_first_name.strip():
                # If first name is provided, match both first and last names
                df_selected = df_selected[
                    ~(
                        (df_selected['First Name'].astype(str).str.strip().str.lower() == owner_first_name.strip().lower()) &
                        (df_selected['Last Name'].astype(str).str.strip().str.lower() == owner_last_name.strip().lower())
                    )
                ]
            else:
                # If first name is empty, match only on last name
                df_selected = df_selected[
                    ~(
                        df_selected['Last Name'].astype(str).str.strip().str.lower() == owner_last_name.strip().lower()
                    )
                ]

            # Add 'Type' column with value 'Neighbors'
            df_selected['Type'] = 'Neighbors'

            # Add 'APN' column with the user-provided value
            df_selected['APN'] = apn_value

            # Calculate tomorrow's date and format it
            tomorrow = datetime.now() + timedelta(days=1)
            df_selected['Mail Date'] = tomorrow.strftime('%b %d, %Y')  # Format as 'Sep 26, 2024'

            # Remove duplicate addresses
            df_selected = df_selected.drop_duplicates(subset=['Mailing Address', 'City', 'State', 'Zip'])

            # Define the desired column order
            output_columns = [
                'Type',
                'First Name',
                'Last Name',
                'Mailing Address',
                'City',
                'State',
                'Zip',
                'Property County',
                'APN',
                'Mail Date'
            ]

            # Reorder the columns
            df_selected = df_selected[output_columns]

            # Save the DataFrame to an Excel file in memory
            from io import BytesIO
            output = BytesIO()
            # Use a context manager to handle the ExcelWriter
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_selected.to_excel(writer, index=False)
            processed_data = output.getvalue()

            st.success(f"Selected columns have been processed successfully.")

            # Provide a download button for the output file
            st.download_button(
                label="Download Output File",
                data=processed_data,
                file_name=output_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}") 

if __name__ == "__main__":
    main()
