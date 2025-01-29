if st.button("Process File"):
    if uploaded_file is not None and key_column and pair_delimiter and (kv_delimiter or extraction_mode == "Extract values without keys"):
        # Load the Excel file into a DataFrame
        df = pd.read_excel(uploaded_file)

        if extraction_mode == "Extract all key-value pairs":
            # Apply the function to the specified column in the DataFrame
            df['extracted'] = df[key_column].apply(lambda row: extract_all_key_values(str(row), pair_delimiter, kv_delimiter))

            # Expand the extracted column into separate columns
            extracted_df = df['extracted'].apply(pd.Series)

            # Rename the columns to the actual keys
            extracted_df.columns = [f'{col}' for col in extracted_df.columns]

            # Concatenate the original DataFrame with the new extracted columns
            df = pd.concat([df, extracted_df], axis=1)

            # Drop the intermediate extracted column
            df.drop(columns=['extracted'], inplace=True)

            # Unpivot the new columns
            new_columns = [col for col in df.columns if col in extracted_df.columns]

            unpivoted_df = df.melt(id_vars=[col for col in df.columns if col not in new_columns],
                                   value_vars=new_columns,
                                   var_name='key_n',
                                   value_name='value')

            # Add a new column with cleaned keys
            unpivoted_df['key'] = unpivoted_df['key_n'].str.replace(r'_\d+$', '', regex=True)
            
        elif extraction_mode == "Extract specific key-value pairs" and key:
            # Apply the function to the specified column in the DataFrame
            df[key] = df[key_column].apply(lambda row: extract_specific_key_values(str(row), pair_delimiter, kv_delimiter, key))

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

        elif extraction_mode == "Extract values without keys":
            # Apply the function to the specified column in the DataFrame
            df['extracted'] = df[key_column].apply(lambda row: extract_values(str(row), pair_delimiter))

            # Expand the extracted column into separate columns
            extracted_df = df['extracted'].apply(pd.Series)

            # Rename the columns to value_n
            extracted_df.columns = [f'value_{i+1}' for i in range(extracted_df.shape[1])]

            # Concatenate the original DataFrame with the new extracted columns
            df = pd.concat([df, extracted_df], axis=1)

            # Drop the intermediate extracted column
            df.drop(columns=['extracted'], inplace=True)

            # Unpivot the new columns
            new_columns = [col for col in df.columns if col.startswith('value')]

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
        st.error("Please upload an Excel file and provide the necessary inputs.")
