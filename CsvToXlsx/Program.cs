using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

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
                    string[] fileInfo = csvFileInfo.Split('>');
                    CreateSheet(fileInfo[0],fileInfo[1], package);
                }                

                FileInfo file = new FileInfo(@outputFile);
                
                package.SaveAs(file);

            }
        }

        private static void CreateSheet(string csvFileLink,string sheetName, ExcelPackage package)
        {
            CultureInfo provider = CultureInfo.InvariantCulture;

            // Add a new worksheet to the empty workbook
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
            //Add the headers


            using (TextFieldParser parser = new TextFieldParser(@csvFileLink))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                int row = 1;
                while (!parser.EndOfData)
                {
                    //Process row
                    string[] fields = parser.ReadFields();
                    int cell = 1;

                    if (parser.LineNumber <= 0)
                    {
                        row++;
                    }
                    else
                    {
                        row = Int32.Parse(parser.LineNumber.ToString()) - 1;
                        
                    }
                        
                    foreach (string field in fields)
                    {



                        //format Currency
                        if (Regex.Match(field, @"[0-9]+\.[0-9][0-9]$").Success)
                        {
                            worksheet.Cells[row, cell].Style.Numberformat.Format = "0.00";
                            worksheet.Cells[row, cell].Value = double.Parse(field);
                        }
                        //format Float Number
                        else if (Regex.Match(field, @"[0-9]+\.[0-9][0-9][0-9]+$").Success)
                        {
                            string[] parts = field.Split('.');
                            string points =  "";
                            for (int i = 0; i < parts[1].Length; i++)
                            {
                                points += "0";
                            }
                            worksheet.Cells[row, cell].Style.Numberformat.Format = "0."+points;
                            worksheet.Cells[row, cell].Value = double.Parse(field);
                        }
                        //format Interger
                        else if (Regex.Match(field, @"^\d+$").Success)
                        {
                            worksheet.Cells[row, cell].Style.Numberformat.Format = "0";
                            worksheet.Cells[row, cell].Value = Int64.Parse(field);
                        }
                        //format Date
                        else if (Regex.Match(field, @"^\d\d\/\d\d\/\d\d\d\d$").Success)
                        {

                            worksheet.Cells[row, cell].Style.Numberformat.Format = "dd-mm-yyyy";
                            worksheet.Cells[row, cell].Value = DateTime.ParseExact(field, "dd/mm/yyyy",provider);
                        }
                        //its text
                        else
                        {                            
                            worksheet.Cells[row, cell].Value = field;
                        }
                                                       

                        cell++;

                    }
                    
                    row++;
                }
            }

            worksheet.Cells.AutoFitColumns();
        }
    }
}
