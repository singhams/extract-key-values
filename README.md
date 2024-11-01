# Extract Key Values
Sometimes you get a metadata export from a system like AEM, which has a generic field for Tags that may contain a bunch of different key:value pairs in a single cell. For example the data in one particular cell might look something like: `industry:healthcare,industry:laboratory,audience:corporate,audience:all,topic:science,topic:math` , etc. etc. Say you only want the audience values - this is the script for you.

- Upload your Excel file. 
- Enter the name of the column your data is in (ex. Tags). 
- Enter the key you are looking for (ex. industry)
- Download the output file.
  
It will have two sheets: 
- Flattened: has new columns for each instance of key:value for that key. It might end up being 2 columns or 20. Depends on your data. 
- Unpivoted: has the new columns unpivoted into a single column, so there is one key:value per row. Note that in this sheet, your records (rows) will be duplicated.
