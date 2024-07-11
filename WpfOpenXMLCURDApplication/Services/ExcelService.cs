// Services/ExcelService.cs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.Linq;
using WpfOpenXMLCURDApplication.Services;

namespace WpfOpenXMLCURDApplication.Services
{

    public class ExcelService : IExcelService
    {
        // Creates a new Excel file with a single sheet
        public void CreateExcelFile(string filePath)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                workbookPart.Workbook.Save();
            }
        }

        // Reads data from an existing Excel file into a DataTable
        public DataTable ReadExcelFile(string filePath)
        {
            DataTable dt = new DataTable();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                int rowIndex= 0;
                int colIndex = 0;
                foreach (Row r in sheetData.Elements<Row>())
                {
                    DataRow tempRow = dt.NewRow();
                    colIndex = 0;
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if(dt.Columns.Count <= colIndex)
                            dt.Columns.Add(colIndex.ToString(), typeof(String));
                        tempRow[rowIndex] = GetValue(spreadsheetDocument, c);
                        colIndex++;
                    }
                    dt.Rows.Add(tempRow);
                    rowIndex++;
                }
            }

            return dt;
        }

        // Updates a specific cell in the Excel sheet
        public void UpdateCell(string filePath, string sheetName, string addressName, string value)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                if (sheet == null)
                {
                    throw new ArgumentException("Invalid sheet name.");
                }

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                Cell cell = GetCell(worksheetPart.Worksheet, addressName);
                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);

                worksheetPart.Worksheet.Save();
            }
        }

        // Deletes a row from the Excel sheet
        public void DeleteRow(string filePath, string sheetName, uint rowIndex)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                if (sheet == null)
                {
                    throw new ArgumentException("Invalid sheet name.");
                }

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                Row row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();

                if (row != null)
                {
                    sheetData.RemoveChild(row);
                    worksheetPart.Worksheet.Save();
                }
            }
        }

        // Retrieves a cell from the worksheet based on its address
        private Cell GetCell(Worksheet worksheet, string addressName)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            //string columnName = GetColumnName(addressName);
            uint rowIndex = GetRowIndex(addressName);

            Row row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();

            if (row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            Cell cell = row.Elements<Cell>().Where(c => c.CellReference.Value == addressName).FirstOrDefault();

            if (cell == null)
            {
                cell = new Cell() { CellReference = addressName };
                row.Append(cell);
            }

            return cell;
        }

        // Extracts the column name from a cell reference (e.g., "A1" -> "A")
        private string GetColumnName(string cellReference)
        {
            return new string(cellReference.Where(c => Char.IsLetter(c)).ToArray());
        }

        // Extracts the row index from a cell reference (e.g., "A1" -> 1)
        private uint GetRowIndex(string cellReference)
        {
            return 1;
                //uint.Parse(new string(cellReference.Where(c => Char.IsDigit(c)).ToArray()));
        }

        // Gets the value of a cell, handling shared strings if necessary
        private string GetValue(SpreadsheetDocument doc, Cell cell)
        {
            SharedStringTablePart stringTablePart = doc.WorkbookPart.SharedStringTablePart;
            if (cell.CellValue == null) return "";

            string value = cell.CellValue.InnerXml;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            return value;
        }
    }
}

