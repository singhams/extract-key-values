import streamlit as st
import pandas as pd
from io import BytesIO

# Display the contents of the README.md file
def display_readme():
    try:
        with open("README.md", "r") as f:
            readme_content = f.read()
        st.markdown(readme_content)
    except FileNotFoundError:
        st.error("README.md file not found.")

display_readme()

# Function to extract key-value pairs
def extract_key_values(row, pair_delimiter, key_value_delimiter):
    items = row.split(pair_delimiter)
    key_values = {}
    key_count = {}
    for item in items:
        if key_value_delimiter in item:
            key, value = item.split(key_value_delimiter, 1)
            key = key.strip()
            value = value.strip()
            if key in key_count:
                key_count[key] += 1
            else:
                key_count[key] = 1
            unique_key = f"{key}_{key_count[key]}"
            key_values[unique_key] = value
    return key_values

# Streamlit app
st.header("Process a File")

# File uploader for Excel input
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

# Input for column name
key_column = st.text_input("Enter the column name containing key-value pairs", value="column_name")

# Input for pair delimiter
pair_delimiter = st.text_input("Enter the delimiter for key-value pairs", value=",")

# Input for key-value delimiter
key_value_delimiter = st.text_input("Enter the delimiter for keys and values", value=":")

# Button to process the file
if st.button("Process File"):
    if uploaded_file is not None and key_column and pair_delimiter and key_value_delimiter:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)

        # Apply the function to the specified column in the DataFrame
        key_values_df = df[key_column].apply(lambda row: pd.Series(extract_key_values(row, pair_delimiter, key_value_delimiter)))

        # Concatenate the original DataFrame with the new key columns
        df = pd.concat([df, key_values_df], axis=1)

        # Unpivot the new columns
        new_columns = key_values_df.columns.tolist()

        unpivoted_df = df.melt(id_vars=[col for col in df.columns if col not in new_columns],
                               value_vars=new_columns,
                               var_name='key',
                               value_name='value')

        # Drop rows with NaN values in the unpivoted columns
        unpivoted_df.dropna(subset=['value'], inplace=True)

        # Add a new column with cleaned keys
        unpivoted_df['cleaned_key'] = unpivoted_df['key'].str.replace(r'_\d+$', '', regex=True)

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
        st.error("Please upload an Excel file and provide the column name and delimiters.")
