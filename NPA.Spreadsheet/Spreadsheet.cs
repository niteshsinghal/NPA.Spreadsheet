using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPA.Spreadsheet
{
    public static class Spreadsheet
    {
        public static IList<IList<string>> Read(FileInfo inputFile)
        {
            if (!inputFile.Exists)
                throw new ApplicationException("Input file does not exist!");

            IReader reader;
            switch (inputFile.Extension.ToLower())
            {
                case ".xls":
                    reader = new XlsReader();
                    break;
                case ".xlsx":
                    reader = new XlsxReader();
                    break;
                case ".csv":
                    reader = new CsvReader();
                    break;
                default:
                    throw new ApplicationException("Unsupported file format!");
            }

            return reader.Read(inputFile);
        }

        public static IList<IList<string>> ReadHeaders(FileInfo inputFile)
        {
            if (!inputFile.Exists)
                throw new ApplicationException("Input file does not exist!");

            IReader reader;
            switch (inputFile.Extension.ToLower())
            {
                case ".xls":
                    reader = new XlsReader();
                    break;
                case ".xlsx":
                    reader = new XlsxReader();
                    break;
                case ".csv":
                    reader = new CsvReader();
                    break;
                default:
                    throw new ApplicationException("Unsupported file format!");
            }

            return reader.ReadFirstRow(inputFile);
        }

        public static IList<IList<string>> ReadCsv(FileInfo inputFile, char separator = default(char))
        {
            if (!inputFile.Exists)
                throw new ApplicationException("Input file does not exist!");

            if (inputFile.Extension.ToLower() != ".csv")
                throw new ArgumentException("inputFile should be a CSV file");

            return new CsvReader(separator).Read(inputFile);
        }
    }
}
