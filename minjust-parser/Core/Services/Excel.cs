using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using minjust_parser.Models;

namespace minjust_parser.Core.Services
{
    public static class Excel
    {
        public static List<string> Read(Config config)
        {
            List<string> data = new List<string>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(config.FilePathInput, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;

                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if (c.CellValue.Text.Length == 0)
                        {
                            break;
                        }

                        Console.WriteLine($"{c.CellValue.Text} добавлен в очередь для парсинга.");
                        text = c.CellValue.Text;
                        data.Add(text);
                    }
                }
            }

            return data;
        }
        public static void Write(List<PersonData> pd, string filePath, long count, string idNumber)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                // Insert the text into the SharedStringTablePart.
                
                
                // Insert a new worksheet.
                WorksheetPart worksheetPart=spreadSheet.WorkbookPart.Workbook.WorkbookPart.WorksheetParts.First();
                //InsertWorksheet(spreadSheet.WorkbookPart);
                
                // Insert cell A1 into the new worksheet.
                int index;
                
                uint rowIndex = Convert.ToUInt32(count);

                string symbs="ABCDEFGHIJKL";

                for (int i = 0; i < 12; i++)
                {
                    Cell cell=InsertCellInWorksheet(symbs[i].ToString(), rowIndex, worksheetPart);
                    index = InsertSharedStringItem(pd[i].value, shareStringPart);
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    worksheetPart.Worksheet.Save();
                }
                
              
                
                
            }
        }
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
        public static void WriteStartPattern(string filePath)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                // Insert the text into the SharedStringTablePart.
                
                
                // Insert a new worksheet.
                WorksheetPart worksheetPart=spreadSheet.WorkbookPart.Workbook.WorkbookPart.WorksheetParts.First();
                //InsertWorksheet(spreadSheet.WorkbookPart);
                
                // Insert cell A1 into the new worksheet.
                int index;
                
                Cell cell1 = InsertCellInWorksheet("A", 1, worksheetPart);
                Cell cell2 = InsertCellInWorksheet("B", 1, worksheetPart);
                Cell cell3 = InsertCellInWorksheet("C", 1, worksheetPart);
                Cell cell4 = InsertCellInWorksheet("D", 1, worksheetPart);
                Cell cell5 = InsertCellInWorksheet("E", 1, worksheetPart);
                
                // Set the value of cell A1.
                index = InsertSharedStringItem("ИНН", shareStringPart);
                cell1.CellValue = new CellValue(index.ToString());
                index = InsertSharedStringItem("Прізвище, ім'я, по батькові", shareStringPart);
                cell2.CellValue = new CellValue(index.ToString());
                index = InsertSharedStringItem("Місцезнаходження", shareStringPart);
                cell3.CellValue = new CellValue(index.ToString());
                index = InsertSharedStringItem("Види діяльності", shareStringPart);
                cell4.CellValue = new CellValue(index.ToString());
                index = InsertSharedStringItem("Інформація для здійснення зв'язку", shareStringPart);
                cell5.CellValue = new CellValue(index.ToString());
                
                cell1.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell2.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell3.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell4.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell5.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    
                
                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }
        }
    }
}