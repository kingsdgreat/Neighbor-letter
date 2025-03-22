# Save this as neighbor_letter_streamlit.py

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import random
from io import BytesIO

def main():
    st.title("Neighbor Letter Processor")

    # File upload
    input_file = st.file_uploader("Upload Input Excel File", type=['xlsx', 'xls'])

    # APN Value
    apn_value = st.text_input("Property APN Value")

    # Owner's First Name
    owner_first_name = st.text_input("Owner's First Name (optional)")

    # Owner's Last Name or Company Name
    owner_last_name = st.text_input("Owner's Last Name or Company Name")

    # Property GPS Coordinates
    gps_coordinates = st.text_input("Property GPS Coordinates")

    # Property City (used for unique code generation)
    property_city = st.text_input("Property City (for generating unique codes)")

    # Output file name
    output_file_name = st.text_input("Output File Name", value="output.xlsx")

    if st.button("Run"):
        if not input_file:
            st.error("Please upload an input file.")
            return
        if not apn_value.strip():
            st.error("Property APN value is required.")
            return
        if not owner_last_name.strip():
            st.error("Owner's last name or company name is required.")
            return
        if not gps_coordinates.strip():
            st.error("Property GPS Coordinates is required.")
            return
        if not property_city.strip():
            st.error("Property City is required.")
            return

        try:
            df = pd.read_excel(input_file)

            column_mapping = {
                'Owner 1 First Name': 'First Name',
                'Owner 1 Last Name': 'Last Name',
                'Mailing Address': 'Mailing Address',
                'Mailing City': 'City',
                'Mailing State': 'State',
                'Mailing Zip': 'Zip',
                'County': 'Property County',
                'State': 'Property State'
            }

            missing_columns = set(column_mapping.keys()) - set(df.columns)
            if missing_columns:
                st.error(f"The following required columns are missing in the input file: {missing_columns}")
                return

            df_selected = df[list(column_mapping.keys())].rename(columns=column_mapping)

            if owner_first_name.strip():
                df_selected = df_selected[
                    ~(
                        (df_selected['First Name'].astype(str).str.strip().str.lower() == owner_first_name.strip().lower()) &
                        (df_selected['Last Name'].astype(str).str.strip().str.lower() == owner_last_name.strip().lower())
                    )
                ]
            else:
                df_selected = df_selected[
                    ~(
                        df_selected['Last Name'].astype(str).str.strip().str.lower() == owner_last_name.strip().lower()
                    )
                ]

            df_selected['Type'] = 'Neighbors'
            df_selected['APN'] = apn_value
            df_selected['GPS Coordinates'] = gps_coordinates

            tomorrow = datetime.now() + timedelta(days=1)
            df_selected['Mail Date'] = tomorrow.strftime('%b %d, %Y')

            df_selected = df_selected.drop_duplicates(subset=['Mailing Address', 'City', 'State', 'Zip'])

            # Reset index before assigning unique codes
            df_selected.reset_index(drop=True, inplace=True)

            # Unique code generation using Property City
            city_abbr = property_city[:3].upper()
            start_number = random.randint(100, 999)
            df_selected['Unique Code'] = df_selected.index.map(lambda i: f"{city_abbr}{start_number + i}")

            output_columns = [
                'Type',
                'First Name',
                'Last Name',
                'Mailing Address',
                'City',
                'State',
                'Zip',
                'Property County',
                'Property State',
                'APN',
                'GPS Coordinates',
                'Mail Date',
                'Unique Code'
            ]

            df_selected = df_selected[output_columns]

            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_selected.to_excel(writer, index=False)
            processed_data = output.getvalue()

            st.success(f"Selected columns have been processed successfully.")

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
