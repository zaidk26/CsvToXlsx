using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvToXlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            string outputFile = args[args.Length - 1];
            Array.Resize(ref args, args.Length - 1);

            using (var package = new ExcelPackage())
            {
                foreach(string csvFileInfo in args)
                {
                    string[] fileInfo = csvFileInfo.Split(':');
                    CreateSheet(fileInfo[0],fileInfo[1], package);
                }                

                FileInfo file = new FileInfo(outputFile+".xlsx");
                package.SaveAs(file);

            }
        }

        private static void CreateSheet(string csvFileLink,string sheetName, ExcelPackage package)
        {
            // Add a new worksheet to the empty workbook
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
            //Add the headers


            using (TextFieldParser parser = new TextFieldParser(csvFileLink))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                int row = 1;
                while (!parser.EndOfData)
                {
                    //Process row
                    string[] fields = parser.ReadFields();
                    int cell = 1;

                    foreach (string field in fields)
                    {
                        worksheet.Cells[row, cell].Value = field;
                        cell++;
                        //Console.WriteLine(field);
                    }
                    // Console.Read();
                    row++;
                }
            }
        }
    }
}
