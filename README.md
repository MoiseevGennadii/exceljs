# exceljs
This code creates an Excel report using nodeJS, libraries 'axios' and 'exceljs'.
● Data is taken from the public API https://api.publicapis.org/entries
● “colored” header with the names of object properties: “API”, “Description”, etc.
● One report line - one record from the API response.
● Active links (clickable).
● Rows sorted by first column name (“API”).
● Objects with HTTPS with the value false are excluded from the report.
● The output is an xsls file with a report.
