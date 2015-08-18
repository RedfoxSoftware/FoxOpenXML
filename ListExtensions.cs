using System.IO;
using System.Linq;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;

namespace FoxOpenXML
{
    public static class ListExtensions
    {
        /// <summary>
        /// Create and return an XLWorkbook object.
        /// </summary>
        /// <param name="source">List of tuples. Tuple.Item1 = datatable (data to write to excel). Tuple.Item2 = string (workbook name).</param>
        /// <returns>ClosedXML.XLWorkBook</returns>
        public static XLWorkbook XlWorkbook(this IList<Tuple<DataTable, string>> source)
        {
            var workbook = new XLWorkbook();
            foreach (var t in source)
            {
                workbook.Worksheets.Add(t.Item1, t.Item2);
            }
            return workbook;
        }

        /// <summary>
        /// Returns a datatable created from an IList of items.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source">Delimited list of items.</param>
        /// <param name="delimiter">Delimiter used for list item.</param>
        /// <returns>Datatable</returns>
        public static DataTable CreateDataTable<T>(this IList<T> source, char delimiter)
        {
            var dt = new DataTable();
            var columns = source[0].ToString().Split(delimiter).Count();
            for (var ii = 0; ii < columns; ii++)
            {
                dt.Columns.Add();
            }

            foreach (var objSplit in source.Select(item => item.ToString().Split(delimiter)).Where(objSplit => objSplit != null))
            {
                var cols = objSplit.Count();
                if (cols > columns)
                {
                    var diff = cols - columns;
                    for (var ii = 0; ii < diff; ii++)
                    {
                        dt.Columns.Add();
                    }
                    columns = cols;
                }

                // ReSharper disable once CoVariantArrayConversion
                dt.Rows.Add(objSplit);
            }
            return dt;
        }

        /// <summary>
        /// Convert list of delimited items to Excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="source">IList of IEnumerable delimited items. Each list of IEnumerable will be put into a separate spreadsheet tab.</param>
        /// <param name="filePath">Full path and file name for workbook.</param>
        /// <param name="delimiter">Delimiter used for list item.</param>
        /// <param name="worksheetNames">Default names for the worksheet, with each name going into a unique tab. If null it will default to "sheet1", "sheet2", etc.</param>
        public static void ToExcel<T>(this IList<IEnumerable<T>> source, string filePath, char delimiter, List<string> worksheetNames = null)
        {
            var dtTupleList = new List<Tuple<DataTable, string>>();
            var index = 0;
            foreach (var s in source)
            {
                if (s == null) continue;

                var enumerable = s as IList<T> ?? s.ToList();
                if (!enumerable.Any()) continue;

                var name = String.Format("Sheet{0}{1}", index, 1);
                if (worksheetNames != null && worksheetNames.Any() && worksheetNames.Count >= index + 1)
                {
                    name = worksheetNames[index];
                }

                var tuple = new Tuple<DataTable, string>(enumerable.CreateDataTable(delimiter), name);
                dtTupleList.Add(tuple);

                index++;
            }

            var workbook = dtTupleList.XlWorkbook();

            if (!string.IsNullOrEmpty(filePath))
            {
                var directory = Path.GetDirectoryName(filePath);
                if (directory != null && !Directory.Exists(directory)) Directory.CreateDirectory(directory);

                workbook.SaveAs(filePath);
            }
        }
    }
}
