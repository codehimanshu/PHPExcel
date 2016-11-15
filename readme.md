**Excel-SQL convertor**
===================


This tool can be used to import a Microsoft Excel file and put the contents of it into the database and Exporting the contents of the Database to an Excel file.

----------
Prerequisites
-------
> Apache Server
> PHP
> MySQL


Excel to DB
--------
> Create an Excel sheet in Examples folder and name it as file.xlxs
> **Currently, only certain field names are allowed, these are:**
> - title
> - details
> - topicId
> - ques_level
> - ques_type

>Setup your database as:
>- Database Name: questions
>- Relpace the credentials of MySQL in the script.

>Run the script in your browser.
> 
DB to Excel
-----------
> Setup the database configuration in script
> This script also supports only the above specified structure.
> Run the script in your browser and the Excel file will be created.
