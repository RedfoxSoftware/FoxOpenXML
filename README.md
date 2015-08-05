# FoxOpenXML
Uses OpenXML and the Nuget package ClosedXML to convert a delimited list of strings to Excel.

Adds several `ToExcel` extension methods to an IEnumerable list of strings.  Allows specification of destination file path and file name, delimiter (defaults to tab if not provided) and worksheet name (defaults to 'sheet1' if not provided.

`ToExcel(string filePath, char delimiter, string worksheetName)` 

`ToExcel(string filePath, char delimiter)`

`ToExcel(string filePath, string worksheetName)`

`ToExcel(string filePath)`

Package can be found on Nuget at: https://www.nuget.org/packages/FoxOpenXML
