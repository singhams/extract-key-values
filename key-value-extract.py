import streamlit as st
import pandas as pd
from io import BytesIO

# Function to extract key-value pairs
def extract_key_values(row, key):
    items = row.split(',')
    key_values = [item for item in items if item.startswith(key)]
    return key_values

# Streamlit app title
st.title("Excel Key-Value Extractor and Unpivot Tool")

# App description
st.write("This Streamlit app allows users to upload an Excel file, specify a column containing key-value pairs, and extract specific keys. The app also provides an option to unpivot the extracted key columns.")

# File uploader for Excel input
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

# Input for column name
key_column = st.text_input("Enter the column name containing key-value pairs", value="column_name")

# Input for key
key = st.text_input("Enter the key to extract", value="key_name")

# Button to process the file
if st.button("Process File"):
    if uploaded_file is not None and key_column and key:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)

        # Apply the function to the specified column in the DataFrame
        df[key] = df[key_column].apply(lambda row: extract_key_values(row, key))

        # Expand the key column into separate columns
        key_df = df[key].apply(pd.Series)

        # Rename the columns to key 1, key 2, etc.
        key_df.columns = [f'{key} {i+1}' for i in range(key_df.shape[1])]

        # Concatenate the original DataFrame with the new key columns
        df = pd.concat([df, key_df], axis=1)

        # Drop the intermediate key column
        df.drop(columns=[key], inplace=True)

        # Unpivot the new columns
        new_columns = [col for col in df.columns if col.startswith(key)]

        unpivoted_df = df.melt(id_vars=[col for col in df.columns if col not in new_columns],
                               value_vars=new_columns,
                               var_name=f'{key}_type',
                               value_name=f'{key}_value')

        # Drop rows with NaN values in the unpivoted columns
        unpivoted_df.dropna(subset=[f'{key}_value'], inplace=True)

        # Convert DataFrame to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Flattened', index=False)
            unpivoted_df.to_excel(writer, sheet_name='Unpivoted', index=False)
        output.seek(0)

        # Provide download link for the Excel file
        st.download_button(
            label="Download Processed Excel file",
            data=output,
            file_name="processed_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload an Excel file and provide both the column name and key.")
