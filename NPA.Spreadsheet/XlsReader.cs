using System.Collections.Generic;
using System.IO;
using NPOI.POIFS.FileSystem;
using System.Linq;
namespace NPA.Spreadsheet
{
    internal class XlsReader : IReader
    {
        public IList<IList<string>> Read(FileInfo inputFile)
        {
            var fs = new POIFSFileSystem(inputFile.OpenRead());

            var converter = new Xls2Strings(fs);
            converter.OutputFormulaValues = true;
            converter.Process();
            return converter.Output;
        }
        public IList<IList<string>> ReadFirstRow(FileInfo inputFile)
        {
            var returndata = Read(inputFile);
            IList<IList<string>> Output = new List<IList<string>>
            {
                returndata.FirstOrDefault()
            };

            return Output;

        }
    }
}