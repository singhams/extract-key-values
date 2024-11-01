# Excel Key-Value Extractor and Unpivot Tool
Sometimes you get a metadata export from a system like AEM, which has a generic field for Tags that may contain a bunch of different key:value pairs in a single cell. For example, the data in one particular cell might look something like: `industry:healthcare,industry:laboratory,audience:corporate,audience:all,topic:science,topic:math` etc. etc. It can be helpful to parse this out so we can analyze each attribute separately.

### How it works
- Upload an Excel file
- Specify the column name containing key-value pairs
- Download the processed Excel file, which will have two sheets: 
  - Flattened: has new columns for each instance of key:value pair. It might end up being 2 columns or 20. Depends on your data. 
  - Unpivoted: has the new columns unpivoted into a single column, so there is one key:value per row. Note that in this sheet, your records (rows) will be duplicated.
