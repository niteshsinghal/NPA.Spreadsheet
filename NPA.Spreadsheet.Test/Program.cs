using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace NPA.Spreadsheet.Test
{
    class Program
    {

        static void Main(string[] args)
        {
          
            string _filePath = Path.Combine(Environment.CurrentDirectory, @"SampleFiles\File.csv");
            //// .csv , .xls , .xlsx 
            var inputFile = new FileInfo(_filePath);
            var sheet = Spreadsheet.Read(inputFile);
            //var sheet = Spreadsheet.ReadHeaders(inputFile);


            Console.WriteLine(Environment.NewLine + " COUNT " + sheet.Count);

            Console.WriteLine(Environment.NewLine + " COMPLETED ");

            Console.ReadLine();
        }
        
    }
}
