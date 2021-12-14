using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace minjust_parser.Helpers
{
    public static class ExcelHelper
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
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c.CellValue.Text;
                        
                        data.Add(text);
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }

            return data;
        }
    }
}