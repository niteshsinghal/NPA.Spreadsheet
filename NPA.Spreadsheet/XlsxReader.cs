using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Sheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;
using Row = DocumentFormat.OpenXml.Spreadsheet.Row;
namespace NPA.Spreadsheet
{
    internal class XlsxReader : IReader
    {
        private readonly DataFormatter _dataFormatter =
            new DataFormatter();
        private IDictionary<int, NumberingFormat> _customFormatRecords = new Dictionary<int, NumberingFormat>();
        private IList<CellFormat> _xfRecords = new List<CellFormat>();

        public IList<IList<string>> Read(FileInfo inputFile)
        {
            var table = new List<IList<string>>();

            using (var document = SpreadsheetDocument.Open(inputFile.FullName, false))
            {
                var wbPart = document.WorkbookPart;

                var sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet == null)
                    return table; // return an empty table

                ReadStyles(wbPart.WorkbookStylesPart);

                var wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));

                var cells = wsPart.Worksheet.Descendants<Cell>()
                    .OrderBy(ReverseColumnRow);

                // Read rows
                var row = new List<string>();
                var ColumnCount = 0;
                foreach (var cell in cells)
                {
                    int columnCode, rowCode;
                    ColumnRow(cell, out columnCode, out rowCode);
                    while (rowCode > table.Count)
                    {
                        if (ColumnCount == 0) ColumnCount = row.Count;
                        if(row.Count < ColumnCount)
                        {
                            var SkipColumn = ColumnCount - row.Count;
                            for(int col = 0;col<SkipColumn;col++)
                                row.Add("");
                        }
                        table.Add(row);
                        row = new List<string>();
                    }
                    
                    while (row.Count < columnCode)
                        row.Add("");

                    // Discard values without header
                    
                    row.Add(GetCellValue(wbPart, cell));
                }

                table.Add(row);
            }

            return table;
        }

        public IList<IList<string>> ReadFirstRow(FileInfo inputFile)
        {
            var table = new List<IList<string>>();

            using (var document = SpreadsheetDocument.Open(inputFile.FullName, false))
            {
                var wbPart = document.WorkbookPart;

                var sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet == null)
                    return table; // return an empty table

                ReadStyles(wbPart.WorkbookStylesPart);

                var wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));

                var rows = wsPart.Worksheet.Descendants<Row>();
                if (rows.Count() > 0)
                {
                    var cells = rows.FirstOrDefault().Elements<Cell>();
                    //var cells = wsPart.Worksheet.Descendants<Cell>()
                    //.OrderBy(ReverseColumnRow);

                    // Read rows
                    var row = new List<string>();
                    //var ColumnCount = 0;
                    foreach (var cell in cells)
                    {
                        int columnCode, rowCode;
                        ColumnRow(cell, out columnCode, out rowCode);
                        while (rowCode > table.Count)
                        {
                            //if (ColumnCount == 0) ColumnCount = row.Count;
                            //if (row.Count < ColumnCount)
                            //{
                            //    var SkipColumn = ColumnCount - row.Count;
                            //    for (int col = 0; col < SkipColumn; col++)
                            //        row.Add("");
                            //}
                            table.Add(row);
                            return table;
                        }

                        //while (row.Count < columnCode)
                        //    row.Add("");

                        // Discard values without header

                        row.Add(GetCellValue(wbPart, cell));
                    }

                    table.Add(row);
                }
            }

            return table;
        }

        private void ReadStyles(WorkbookStylesPart wsStyles)
        {
            var formats = wsStyles.Stylesheet.CellFormats;
            foreach (var format in formats.Descendants<CellFormat>())
            {
                _xfRecords.Add(format);
            }
            if (wsStyles.Stylesheet != null && wsStyles.Stylesheet.NumberingFormats != null)
            {
                foreach (var format in wsStyles.Stylesheet.NumberingFormats.Descendants<NumberingFormat>())
                {
                    _customFormatRecords.Add((int)format.NumberFormatId.Value, format);
                }
               
            }
        }

        private static void ColumnRow(Cell cell, out int column, out int row)
        {
            string columnCode, rowCode;
            ColumnRow(cell, out columnCode, out rowCode);
            column = columnCode.Aggregate(0, (sum, c) =>
                sum * ('Z' - 'A' + 1) + (char.ToUpper(c) - 'A') + 1) - 1;
            row = Convert.ToInt32(rowCode) - 1;
        }

        private class ColRow : IComparable<ColRow>
        {
            public int Column { get; set; }
            public int Row { get; set; }

            public ColRow(int row, int column)
            {
                Row = row;
                Column = column;
            }

            public int CompareTo(ColRow other)
            {
                if (Row > other.Row)
                    return 1;
                if (Row < other.Row)
                    return -1;
                if (Column > other.Column)
                    return 1;
                if (Column < other.Column)
                    return -1;
                return 0;
            }
        }

        private static ColRow ReverseColumnRow(Cell cell)
        {
            int columnCode, rowCode;
            ColumnRow(cell, out columnCode, out rowCode);
            return new ColRow(rowCode, columnCode);
        }

        private static void ColumnRow(Cell cell, out string column, out string row)
        {
            var m = Regex.Match(cell.CellReference.Value, @"^(?<Column>[A-Z]+)(?<Row>[0-9]+)$", RegexOptions.IgnoreCase);
            column = m.Groups["Column"].Value;
            row = m.Groups["Row"].Value;
        }

        private string GetCellValue(WorkbookPart wbPart, Cell cell)
        {
            var value = cell.InnerText;

            // If the cell represents an integer number, you are done. 
            // For dates, this code returns the serialized value that 
            // represents the date. The code handles strings and 
            // Booleans individually. For shared strings, the code 
            // looks up the corresponding value in the shared string 
            // table. For Booleans, the code converts the value into 
            // the words TRUE or FALSE.
            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:

                        // For shared strings, look up the value in the
                        // shared strings table.
                        var stringTable =
                            wbPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                        // If the shared string table is missing, something 
                        // is wrong. Return the index that is in
                        // the cell. Otherwise, look up the correct text in 
                        // the table.
                        if (stringTable != null)
                        {
                            value =
                                stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                        }
                        break;

                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                        break;
                }
            }

            if (cell.CellFormula != null)
            {
                value = cell.CellValue.Text;
            }

            if (cell.StyleIndex != null)
            {
                var formatIndex = GetFormatIndex(cell);
                if (formatIndex == 58) formatIndex = 14;
                var formatString = GetFormatString(formatIndex);
                
                value = FormatCellContents(value, formatIndex, formatString);
            }

            return value;
        }

        private string GetFormatString(int formatIndex)
        {
            if (formatIndex >= HSSFDataFormat.NumberOfBuiltinBuiltinFormats)
            {
                
                var numFmt = _customFormatRecords[formatIndex];
                if (numFmt == null)
                    throw new ApplicationException("Requested format at index " +
                        formatIndex + ", but it wasn't found");
                return numFmt.FormatCode.InnerText;
            }
            else
            {
                return HSSFDataFormat.GetBuiltinFormat((short)formatIndex);
            }
        }

        private int GetFormatIndex(Cell cell)
        {
            var index = Convert.ToInt32(cell.StyleIndex.InnerText);
            if (index >= _xfRecords.Count)
                throw new ApplicationException("Cell " + cell.CellReference.Value +
                    " uses XF with index " + index + ", but we don't have that");
            if (_xfRecords[index].NumberFormatId == null)
                return 0;
            else
                return Convert.ToInt32(_xfRecords[index].NumberFormatId.InnerText);

            //return Convert.ToInt32(_xfRecords[index].NumberFormatId.InnerText);
        }

        private string FormatCellContents(string value, int formatIndex, string formatString)
        {
            double d = 0;
            
            if ((value.StartsWith("0") && !value.Contains(".")) || value.ToUpper().Contains("E"))
            {
                if (formatIndex == 164)
                    return _dataFormatter.FormatRawCellContents(d, formatIndex, formatString);
                else
                    return value;
            }
            else
            {
                try
                {
                    d = double.Parse(value, NumberStyles.Any, CultureInfo.InvariantCulture);
                }
                catch
                {
                    return value;
                }
                try
                {
                    return _dataFormatter.FormatRawCellContents(d, formatIndex, formatString);
                }
                catch
                {
                    return value;
                }
                
            }
        }
    }
}
