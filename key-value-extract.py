import streamlit as st
import pandas as pd
from io import BytesIO

# Function to extract key-value pairs
def extract_key_values(row, delimiter, key=None):
    items = row.split(delimiter)
    if key:
        key_values = [item for item in items if item.startswith(key)]
    else:
        key_values = items
    return key_values

# Streamlit app
st.title("Excel Key-Value Extractor and Unpivot Tool")

# File uploader for Excel input
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

# Input for column name
key_column = st.text_input("Enter the column name containing key-value pairs", value="column_name")

# Input for delimiter
delimiter = st.text_input("Enter the delimiter used in the key-value pairs", value=",")

# Radio button to choose extraction mode
extraction_mode = st.radio("Choose extraction mode", ("Extract all key-value pairs", "Extract specific key-value pairs"))

# Input for specific key (only shown if the user chooses to extract specific key-value pairs)
if extraction_mode == "Extract specific key-value pairs":
    key = st.text_input("Enter the key to extract", value="key_name")
else:
    key = None

# Button to process the file
if st.button("Process File"):
    if uploaded_file is not None and key_column and delimiter:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)

        # Apply the function to the specified column in the DataFrame
        df['extracted'] = df[key_column].apply(lambda row: extract_key_values(row, delimiter, key))

        # Expand the extracted column into separate columns
        extracted_df = df['extracted'].apply(pd.Series)

        # Rename the columns to key 1, key 2, etc.
        extracted_df.columns = [f'{key or "key"} {i+1}' for i in range(extracted_df.shape[1])]

        # Concatenate the original DataFrame with the new extracted columns
        df = pd.concat([df, extracted_df], axis=1)

        # Drop the intermediate extracted column
        df.drop(columns=['extracted'], inplace=True)

        # Unpivot the new columns
        new_columns = [col for col in df.columns if col.startswith(key or "key")]

        unpivoted_df = df.melt(id_vars=[col for col in df.columns if col not in new_columns],
                               value_vars=new_columns,
                               var_name=f'{key or "key"}_type',
                               value_name=f'{key or "key"}_value')

        # Drop rows with NaN values in the unpivoted columns
        unpivoted_df.dropna(subset=[f'{key or "key"}_value'], inplace=True)

        # Convert DataFrame to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            unpivoted_df.to_excel(writer, sheet_name='Flattened', index=False)

        # Provide download link
        st.download_button("Download Excel", data=output.getvalue(), file_name='flattened.xlsx', key="download_excel")
