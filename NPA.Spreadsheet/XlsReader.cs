using System.Collections.Generic;
using System.IO;
using NPOI.POIFS.FileSystem;
using System.Linq;
namespace NPA.Spreadsheet
{
    internal class XlsReader : IReader
    {
        /// <summary>
        /// Read all rows
        /// </summary>
        /// <param name="inputFile"></param>
        /// <returns></returns>
        public IList<IList<string>> Read(FileInfo inputFile)
        {
            var fs = new POIFSFileSystem(inputFile.OpenRead());

            var converter = new Xls2Strings(fs);
            converter.OutputFormulaValues = true;
            converter.Process();
            return converter.Output;
        }
        /// <summary>
        /// Read Header or First Row
        /// </summary>
        /// <param name="inputFile"></param>
        /// <returns></returns>
        public IList<IList<string>> ReadFirstRow(FileInfo inputFile)
        {
            var returndata = Read(inputFile);
            IList<IList<string>> Output = new List<IList<string>>
            {
                returndata.FirstOrDefault()
            };

            return Output;

        }

       /// <summary>
       /// Return specified number of rows only
       /// </summary>
       /// <param name="inputFile"></param>
       /// <param name="numberOfRows"></param>
       /// <returns></returns>
        public IList<IList<string>> ReadFirstNRow(FileInfo inputFile,int numberOfRows)
        {
            var returndata = Read(inputFile);
            return returndata.Take(numberOfRows).ToList();
        }
    }
}