import streamlit as st
import pandas as pd
from io import BytesIO

# Function to extract key-value pairs
def extract_key_values(row, pair_delimiter, kv_delimiter, key):
    items = row.split(pair_delimiter)
    key_values = []
    for item in items:
        if item.startswith(key + kv_delimiter):
            key_values.append(item.split(kv_delimiter, 1)[1].strip())
    return key_values

# Display the contents of the README.md file
def display_readme():
    try:
        with open("README.md", "r") as f:
            readme_content = f.read()
        st.markdown(readme_content)
    except FileNotFoundError:
        st.error("README.md file not found.")

display_readme()

# File uploader for Excel input
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

# Input for column name
key_column = st.text_input("Enter the column name containing key-value pairs", value="column_name")

# Input for pair delimiter
pair_delimiter = st.text_input("Enter the delimiter used between key-value pairs", value=",")

# Input for key-value delimiter
kv_delimiter = st.text_input("Enter the delimiter used between keys and values", value=":")

# Input for key
key = st.text_input("Enter the key to extract", value="key_name")

# Button to process the file
if st.button("Process File"):
    if uploaded_file is not None and key_column and pair_delimiter and kv_delimiter and key:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)

        # Apply the function to the specified column in the DataFrame
        df[key] = df[key_column].apply(lambda row: extract_key_values(row, pair_delimiter, kv_delimiter, key))

        # Expand the key column into separate columns
        key_df = df[key].apply(pd.Series)

        # Rename the columns to key_n
        key_df.columns = [f'{key}_{i+1}' for i in range(key_df.shape[1])]

        # Concatenate the original DataFrame with the new key columns
        df = pd.concat([df, key_df], axis=1)

        # Drop the intermediate key column
        df.drop(columns=[key], inplace=True)

        # Unpivot the new columns
        new_columns = [col for col in df.columns if col.startswith(key)]

        unpivoted_df = df.melt(id_vars=[col for col in df.columns if col not in new_columns],
                               value_vars=new_columns,
                               var_name='key_n',
                               value_name='value')

        # Add a new column with cleaned keys
        unpivoted_df['key'] = unpivoted_df['key_n'].str.replace(r'_\d+$', '', regex=True)

        # Drop rows with NaN values in the unpivoted columns
        unpivoted_df.dropna(subset=['value'], inplace=True)

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
        st.error("Please upload an Excel file and provide all necessary inputs.")
