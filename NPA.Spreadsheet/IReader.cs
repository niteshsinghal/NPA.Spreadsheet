using System.Collections.Generic;
using System.IO;

namespace NPA.Spreadsheet
{
    internal interface IReader
    {
        IList<IList<string>> Read(FileInfo inputFile);
        IList<IList<string>> ReadFirstRow(FileInfo inputFile);
    }
}