# FoxOpenXML

Release notes:
1.0.3 -> Add ability to import data into existing worksheet. For now, starts import at cell 1,1. 
1.0.2 -> Add support for multiple xl tabs. 
1.0.1 -> Add overloaded methods. Check if destination directory exists. 
1.0.0 -> Initial release.  More features coming soon.

Uses OpenXML and the Nuget package ClosedXML to convert a delimited list of strings to Excel.

Adds several `ToExcel` extension methods to an IEnumerable list of strings.  Allows specification of destination file path and file name, delimiter (defaults to tab if not provided) and worksheet name (defaults to 'sheet1' if not provided).

`ToExcel<T>(string filePath, char delimiter, string worksheetName)` 

`ToExcel<T>(string filePath, char delimiter)`

`ToExcel<T>(string filePath, string worksheetName)`

`ToExcel<T>(string filePath)`

`ToExcel<T>(this IList<IEnumerable<T>> source, string filePath, char delimiter, List<string> worksheetNames = null)`

`InsertData(this XLWorkbook source, IEnumerable<string> data, string worksheetName, char delimiter = '\t')`

Package can be found on Nuget at: https://www.nuget.org/packages/FoxOpenXML
