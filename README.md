# similarpanda

A module to check for differences between pandas Dataframes, and generate a report in Excel format.

This is helpful in a workplace setting, where you might need to check between different versions of an Excel file.

Inputs (to create the object)
* new: DataFrame
* old: DataFrame
* key_column: A reference to a common DataFrame column. If omitted, the data will be matched on the row position.
   
Returns:
* A similarpanda object
* Use similarpanda.output_excel() to generate an Excel report.
    
Note:
* Tabular data must contain unique column names.
* Only checks cell values. Formulae, formatting ets is disregarded.

Includes a helper Jupyer notebook, which supports in loading of tabular data:
* From specified named tables within named Excel files
* From the clipboard
   
![Example output](example.png)
