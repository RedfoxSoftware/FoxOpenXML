using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FoxOpenXML
{
    public static class EnumerableExtensions
    {
        /// <summary>
        /// Convert list of delimited items to Excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source">Delimited list of items.</param>
        /// <param name="filePath">Full path and file name for workbook.</param>
        /// <param name="delimiter">Delimiter used for list item.</param>
        /// <param name="worksheetName">Default name for the worksheet.</param>
        public static void ToExcel<T>(this IEnumerable<T> source, string filePath, char delimiter, string worksheetName)
        {
            var enumerable = source as IList<T> ?? source.ToList();
            if (!enumerable.Any())
                throw new Exception("No data was detected.");

            var dt = enumerable.CreateDataTable(delimiter);

            var directory = Path.GetDirectoryName(filePath);
            if (directory != null && !Directory.Exists(directory)) Directory.CreateDirectory(directory);

            var workbook = new XLWorkbook();
            workbook.Worksheets.Add(dt, worksheetName);
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Convert list of delimited items to Excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source">Delimited list of items.</param>
        /// <param name="filePath">Full path and file name for workbook.</param>
        /// <param name="delimiter">Delimiter used for list item.</param>
        public static void ToExcel<T>(this IEnumerable<T> source, string filePath, char delimiter)
        {
            source.ToExcel(filePath, delimiter, "Sheet1");
        }

        /// <summary>
        /// Convert list of tab delimited items to Excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source">Delimited list of items.</param>
        /// <param name="filePath">Full path and file name for workbook.</param>
        /// <param name="worksheetName">Default name for the worksheet.</param>
        public static void ToExcel<T>(this IEnumerable<T> source, string filePath, string worksheetName)
        {
            source.ToExcel(filePath, '\t', worksheetName);
        }

        /// <summary>
        /// Convert list of tab delimited items to Excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source">Delimited list of items.</param>
        /// <param name="filePath">Full path and file name for workbook.</param>
        public static void ToExcel<T>(this IEnumerable<T> source, string filePath)
        {
            source.ToExcel(filePath, '\t', "Sheet1");
        }
    }
}
