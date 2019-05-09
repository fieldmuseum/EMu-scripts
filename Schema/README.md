Folder for scripts relating to the Schema.
These files contain scripts that run on the EMu schema.pl file
  1. convertSchemaToExcel.tgz is from Chresty at Axiell


## HELP FILE FOR THE convertSchemaToExcel.pl SCRIPT
The convertSchemaToExcel.pl script can be used to convert the schema.pl file to an Excel document. Running the script involves the following:

1) Place the script in the environment running EMu. 
   For example, the ~emu/client/work/ directory where client is the environment name.
2) Confirm if the ~emu/client/utils/schema/pl file is present where client is the enviornment name.
3) Enter `convertSchemaToExcel.pl [excel_file_name]` to run the script. 
    For example, `convertSchemaToExcel.pl client_schema.xls` OR `convertSchemaToExcel.pl client_schema`
   
#### Troubleshooting:
* You might need to enter `perl convertSchemaToExcel.pl [excel_file_name]` to run the script
* If you get an error like this "You may need to install the Spreadsheet::WriteExcel module"), enter `cpan install Spreadsheet::WriteExcel` OR `sudo cpan install Spreadsheet::WriteExcel`

### NOTE: 
* The script will create the Excel file in the same folder as the script.
* All tables except the ecatalogue table are shared tables, so fields appearing in the resulting Excel file will include custom fields from other implementations.


## Column Names in the Excel file are the following:
1) Table - The backend table name of the table
2) ColumnName - The backend column name of the field
3) DataKind - The kind of the column. Types include the following:
    * dkAtom
    * dkTable
    * dkKey
    * dkNested
    * dkTuple
4) DataType - The data type of the column. Types include the following:
    * Date
    * Text
    * Currency
    * Float
    * Integer
    * Time
    * UserId
    * UserName
    * String
    * Longitude
    * Latitude
5) ItemCount - If the field is a table field the ItemCount field will output the number of rows in the table
6) ItemFields - The number of characters a field can hold per row. If the field is a nested table the numbers will be seperated by a pipe symbol '|'.
7) ItemName - Same as ColumnName.
8) ItemPrompt - Item prompt.
9) Location - Where the field can be found. The value can be either of the following:
    * both
    * server
    * client
10) ItemBase - Base name of a column if it is a table value field.
11) LookupName - The lookup list name associated with the column.
12) RefColumn - The backend column name of a fild existing in the reference table where the data will be pulled from.
13) RefLink - A reference to the reference column.
14) RefKey - The unique field used to distinguish what record is referenced. Will 
	usually be assigned 'irn'.
15) RefPrompt - The column used to reference the label which will be used via the client-side.
16) RefTable - Specifices the reference table if the field is an attachemnt field.
17) RefLocal - The local ref field linked to the attachment field.
18) Reportable - A boolean field indicating if the field is reportable. Will only appear if the field is set to False. The following values are accepted:
    * False
    * True
19) DataLocal - A boolean field to indicate if the field is a local type field in the client. The following values are accepted:
    * False
    * True
20) RefSort - The reference sort fields.
21) RefVirtualTable - The virtual reference table.
22) LookupParent - If the field is part of a lookup hierarchy the column name of the parent field will be assigned here.
23) ItemNest - If the field is a nested table the ItemNest field will be assigned the Index column name.
24) Title - Title information.
25) StringQuery - A boolean field to indicate if String Query is allowed.
26) UseThesaurus - If the fiel uses the Thesaurus button. UseThesaurus will be assigned 1.
27) FormatInput - Accepted input format of data.
28) FormatOutput - Accepted output format of data.
29) RefVirtualLink - The virtual link reference field.
30) NotifyColumns - Notify columns.
