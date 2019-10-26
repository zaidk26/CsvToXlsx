﻿using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CsvToXlsx
{
    class Program
    {
        private static bool autofit = true;
        static void Main(string[] args)
        {
            string outputFile = args[args.Length - 1];
            Array.Resize(ref args, args.Length - 1);

            using (var package = new ExcelPackage())
            {
                foreach(string csvFileInfo in args)
                {
                    string[] fileInfo = csvFileInfo.Split('>');
                    CreateSheet(fileInfo[0],fileInfo[1],fileInfo[2],package);
                }                

                FileInfo file = new FileInfo(outputFile);
                
                package.SaveAs(file);

            }
        }

        /// <summary>
        /// 
        /// Create Sheet
        /// 
        /// </summary>
        /// <param name="csvFileLink"></param>
        /// <param name="sheetName"></param>
        /// <param name="package"></param>
        private static void CreateSheet(string csvFileLink,string sheetName,string dateFormat, ExcelPackage package)
        {
            CultureInfo provider = CultureInfo.InvariantCulture;

            // Add a new worksheet to the empty workbook
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
            //Add the headers
           

            CsvFileParser(csvFileLink, dateFormat, provider, worksheet);
            if (autofit)
            {
                worksheet.Cells.AutoFitColumns();
            }
            
        }

        /// <summary>
        /// 
        /// Parse file and add data to sheet
        /// 
        /// </summary>
        /// <param name="csvFileLink"></param>
        /// <param name="provider"></param>
        /// <param name="worksheet"></param>
        private static void CsvFileParser(string csvFileLink,string dateFormat, CultureInfo provider, ExcelWorksheet worksheet)
        {
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

                    if (parser.LineNumber <= 0)
                    {
                        row++;
                    }
                    else
                    {
                        row = Int32.Parse(parser.LineNumber.ToString()) - 1;
                    }

                    //Add data to sheet
                    foreach (string field in fields)
                    {
                        //check if has styling
                        if (field.Contains("#$#"))
                        {
                            string[] fieldData = Regex.Split(field,@"#\$#"); 
                            
                            FormatAndInsertValue(provider, worksheet, row, cell, fieldData[0], dateFormat);
                            
                            StyleCell(worksheet, row, cell, fieldData.Where((v, i) => i != 0).ToArray());
                        }
                        else
                        {
                            FormatAndInsertValue(provider, worksheet, row, cell, field, dateFormat);
                        }
                        
                        cell++;
                    }

                    row++;
                }
            }
        }

        /// <summary>
        /// Auto format the value and insert into field
        /// 
        /// </summary>
        /// <param name="provider"></param>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <param name="field"></param>
        private static void FormatAndInsertValue(CultureInfo provider, ExcelWorksheet worksheet, int row, int cell, string field,string dateFormat)
        {
            //format Currency
            if (Regex.Match(field, @"^[0-9]+\.[0-9][0-9]$").Success)
            {
                worksheet.Cells[row, cell].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[row, cell].Value = double.Parse(field);
            }
            //format Currency Negative
            else if (Regex.Match(field, @"^\([0-9]+\.[0-9][0-9]\)$").Success)
            {
                worksheet.Cells[row, cell].Style.Numberformat.Format = "#,##0.00";
                worksheet.Cells[row, cell].Value = double.Parse(field.Replace("(", "-").Replace(")", ""));
            }
            //format Float Number
            else if (Regex.Match(field, @"^[0-9]+\.[0-9][0-9][0-9]+$").Success)
            {
                string[] parts = field.Split('.');
                string points = "";
                for (int i = 0; i < parts[1].Length; i++)
                {
                    points += "0";
                }
                worksheet.Cells[row, cell].Style.Numberformat.Format = "#,##0." + points;
                worksheet.Cells[row, cell].Value = double.Parse(field);
            }
            //format Float Number Negative
            else if (Regex.Match(field, @"^\([0-9]+\.[0-9][0-9][0-9]+\)$").Success)
            {
                string[] parts = field.Split('.');
                string points = "";
                for (int i = 0; i < parts[1].Replace(")","").Length; i++)
                {
                    points += "0";
                }
                worksheet.Cells[row, cell].Style.Numberformat.Format = "#,##0." + points;
                worksheet.Cells[row, cell].Value = double.Parse(field.Replace("(", "-").Replace(")", ""));
            }
            //format Interger
            else if (Regex.Match(field, @"^[1-9]([0-9]+)?$").Success && field.Length < 16)
            {
                worksheet.Cells[row, cell].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[row, cell].Value = Int64.Parse(field);
            }
            //format Negative Interger
            else if (Regex.Match(field, @"^\([1-9]([0-9]+)?\)$").Success && field.Length < 16)
            {
                worksheet.Cells[row, cell].Style.Numberformat.Format = "#,##0";
                worksheet.Cells[row, cell].Value = Int64.Parse(field.Replace("(","-").Replace(")",""));
            }
            //format Date
            else if (Regex.Match(field, @"^\d\d\/\d\d\/\d\d\d\d$").Success)
            {

                worksheet.Cells[row, cell].Style.Numberformat.Format = dateFormat;
                worksheet.Cells[row, cell].Value = DateTime.ParseExact(field, "dd/mm/yyyy", provider);
            }
            //its text
            else
            {
                worksheet.Cells[row, cell].Value = field;
            }
        }


        // <summary>
        /// Auto format the value and insert into field
        /// 
        /// </summary>
        /// <param name="provider"></param>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <param name="field"></param>
        private static void StyleCell(ExcelWorksheet worksheet, int row, int cell, string[] styles)
        {
            foreach (string style in styles)
            {
                string[] statement = style.Split('=');
                string property = statement[0];
                string value = statement[1];

                switch (property)
                {
                    case "column-width":
                        autofit = false;
                        worksheet.Column(cell).Width = double.Parse(value);
                        break;

                    case "font-bold":
                        worksheet.Cells[row, cell].Style.Font.Bold = true;
                        break;

                    case "font-italic":
                        worksheet.Cells[row, cell].Style.Font.Italic = true;
                        break;

                    case "font-underline":
                        worksheet.Cells[row, cell].Style.Font.UnderLine = true;
                        break;

                    case "font-color":
                        worksheet.Cells[row, cell].Style.Font.UnderLine = true;
                        break;

                }
            }
        }
    }
}
