using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace minjust_parser.Core.Services
{
    public static class Excel
    {
        public static List<string> Read()
        {
            List<string> data = new List<string>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("Test.xlsx", false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                int count = 0;

                var rer = sheetData.Elements<Row>();
                
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        /*Console.WriteLine(count++);*/
                        if (c.CellValue.Text.Length==0)
                        {
                            break;
                        }
                        Console.WriteLine(c.CellValue.Text);
                        text = c.CellValue.Text;
                        data.Add(text);
                    }
                }
            }

            return data;
        }
    }
}