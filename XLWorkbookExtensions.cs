using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace FoxOpenXML
{
    public static class XlWorkbookExtensions
    {
        /// <summary>
        /// Import data into an existing XLWorkbook worksheet. Currently imports data starting at cell 1,1.
        /// </summary>
        /// <param name="source">Existing XLWorkbook.</param>
        /// <param name="data">Data to import.  Delimited IEnumerable of strings.</param>
        /// <param name="delimiter">Delimiter to use.</param>
        /// <param name="worksheetName">Worksheet name to import into.  Creates a new sheet if it does not currently exist.</param>
        /// <returns>XLWorkbook (unsaved to disk).</returns>
        public static XLWorkbook InsertData(this XLWorkbook source, IEnumerable<string> data, string worksheetName, char delimiter = '\t')
        {
            if (source == null) throw new NullReferenceException("XLWorkbook cannot be null.");
            if (data == null) throw new NullReferenceException("IEnumerable cannot be null.");
            if (string.IsNullOrEmpty(worksheetName)) throw new NullReferenceException("worksheetName cannot be null.");

            IXLWorksheet worksheet;
            var isWorksheetPresent = source.TryGetWorksheet(worksheetName, out worksheet);
            if (!isWorksheetPresent)
            {
                source.Worksheets.Add(worksheetName);
                isWorksheetPresent = source.TryGetWorksheet(worksheetName, out worksheet);
                if (!isWorksheetPresent) throw new Exception("Specified worksheet is not present and could not be added.");
            }

            if (worksheet == null) throw new Exception("Unknown error occurred retrieving specified worksheet.");

            var dt = data.ToList().CreateDataTable(delimiter);
            worksheet.Cell(1, 1).InsertData(dt.AsEnumerable());

            return source;
        }
    }
}
