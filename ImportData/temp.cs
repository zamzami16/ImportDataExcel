using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace ArrayToExcel
{
    public static class EXcel2DataSet
    {
        public static DataTable MyExcelData(string filepath, bool ColumnHeader = true, bool _Isemptyheader = false)
        {
            DataTable dt = new DataTable();

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filepath, false))
            {

                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.ElementAt(0).Id.Value; //sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                int rowCount = sheetData.Descendants<Row>().Count();
                if (rowCount == 0)
                {
                    return dt;
                }
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                var charcolumn = 'A';
                foreach (Cell cell in rows.ElementAt(0))
                {
                    if (GetCellValue(spreadSheetDocument, cell).ToString() != "" || _Isemptyheader)
                    {
                        if (ColumnHeader)
                            dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                        else
                        {
                            dt.Columns.Add(charcolumn.ToString());
                            charcolumn++;
                        }
                    }
                }
                foreach (Row row in rows) //this will also include your header row...
                {
                    DataRow tempRow = dt.NewRow();
                    int columnIndex = 0;
                    foreach (Cell cell in row.Descendants<Cell>())
                    {
                        // Gets the column index of the cell with data
                        int cellColumnIndex = (int)GetColumnIndexFromName(GetColumnName(cell.CellReference));
                        cellColumnIndex--; //zero based index
                        if (columnIndex < cellColumnIndex)
                        {
                            do
                            {
                                //tempRow[columnIndex] = ""; //Insert blank data here;
                                columnIndex++;
                            }
                            while (columnIndex < cellColumnIndex);
                        }
                        if (columnIndex < dt.Columns.Count)
                        {
                            tempRow[columnIndex] = GetCellValue(spreadSheetDocument, cell);
                            columnIndex++;
                        }
                        //tempRow[columnIndex] = GetCellValue(spreadSheetDocument, cell);
                        //columnIndex++;
                    }
                    dt.Rows.Add(tempRow);
                }
            }
            if (ColumnHeader)
                dt.Rows.RemoveAt(0);
            return dt;
        }

        private static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);
            return match.Value;
        }

        private static int? GetColumnIndexFromName(string columnNameOrCellReference)
        {
            int columnIndex = 0;
            int factor = 1;
            for (int pos = columnNameOrCellReference.Length - 1; pos >= 0; pos--) // R to L
            {
                if (Char.IsLetter(columnNameOrCellReference[pos])) // for letters (columnName)
                {
                    columnIndex += factor * ((columnNameOrCellReference[pos] - 'A') + 1);
                    factor *= 26;
                }
            }
            return columnIndex;

        }

        private static string GetCellValue(SpreadsheetDocument document, DocumentFormat.OpenXml.Spreadsheet.Cell cell)
        {
            DateTime ReleaseDate = new DateTime(1899, 12, 30);
            TimeSpan offset = new(days: 0, hours: 0, minutes: 0, seconds: 0, milliseconds: 0);
            ReleaseDate.Add(offset);
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            object value = string.Empty;
            DocumentFormat.OpenXml.Spreadsheet.CellFormats cellFormats = (DocumentFormat.OpenXml.Spreadsheet.CellFormats)document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;

            string format = string.Empty; 
            uint formatid = 0;

            if (cell.DataType == null)
            {
                DocumentFormat.OpenXml.Spreadsheet.CellFormat cf = new CellFormat();
                if (cell.StyleIndex == null)
                {
                    cf = cellFormats.Descendants<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(0);
                }
                else
                {
                    cf = cellFormats.Descendants<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ElementAt<DocumentFormat.OpenXml.Spreadsheet.CellFormat>(Convert.ToInt32(cell.StyleIndex.Value));
                }

                formatid = cf.NumberFormatId;

                if (cell != null && cell.InnerText.Length > 0)
                {
                    value = cell.CellValue.Text;
                    if (formatid > 13 && formatid <= 22)
                    {
                        var dateFormat = GetDateTimeFormat(formatid);
                        string val = value.ToString().Replace(".", ",");
                        return DateTime.FromOADate(double.Parse(val)).ToString();
                    }

                }
                else
                {

                    value = cell.InnerText;
                }
            }

            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                        return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(cell.CellValue.Text)].InnerText;
                    case CellValues.Boolean:
                        return cell.CellValue.Text == "1" ? "true" : "false";
                    case CellValues.Date:
                        {
                            DateTime answer = ReleaseDate.AddDays(Convert.ToDouble(cell.CellValue.Text));
                            return answer.ToShortDateString();
                        }
                    case CellValues.Number:
                        return Convert.ToDecimal(cell.CellValue.Text).ToString();
                    default:
                        if (cell.CellValue != null)
                            return cell.CellValue.Text;
                        return string.Empty;
                }
            }

            return value.ToString();
        }

        private  static string GetDateTimeFormat(uint numberFormatId)
        {
            return DateFormatDictionary.ContainsKey(numberFormatId) ? DateFormatDictionary[numberFormatId] : string.Empty;
        }

        //// https://msdn.microsoft.com/en-GB/library/documentformat.openxml.spreadsheet.numberingformat(v=office.14).aspx
        private readonly static Dictionary<uint, string> DateFormatDictionary = new Dictionary<uint, string>()
        {
            [14] = "dd/MM/yyyy",
            [15] = "d-MMM-yy",
            [16] = "d-MMM",
            [17] = "MMM-yy",
            [18] = "h:mm AM/PM",
            [19] = "h:mm:ss AM/PM",
            [20] = "h:mm",
            [21] = "h:mm:ss",
            [22] = "M/d/yy h:mm",
            [30] = "M/d/yy",
            [34] = "yyyy-MM-dd",
            [45] = "mm:ss",
            [46] = "[h]:mm:ss",
            [47] = "mmss.0",
            [51] = "MM-dd",
            [52] = "yyyy-MM-dd",
            [53] = "yyyy-MM-dd",
            [55] = "yyyy-MM-dd",
            [56] = "yyyy-MM-dd",
            [58] = "MM-dd",
            [165] = "M/d/yy",
            [166] = "dd MMMM yyyy",
            [167] = "dd/MM/yyyy",
            [168] = "dd/MM/yy",
            [169] = "d.M.yy",
            [170] = "yyyy-MM-dd",
            [171] = "dd MMMM yyyy",
            [172] = "d MMMM yyyy",
            [173] = "M/d",
            [174] = "M/d/yy",
            [175] = "MM/dd/yy",
            [176] = "d-MMM",
            [177] = "d-MMM-yy",
            [178] = "dd-MMM-yy",
            [179] = "MMM-yy",
            [180] = "MMMM-yy",
            [181] = "MMMM d, yyyy",
            [182] = "M/d/yy hh:mm t",
            [183] = "M/d/y HH:mm",
            [184] = "MMM",
            [185] = "MMM-dd",
            [186] = "M/d/yyyy",
            [187] = "d-MMM-yyyy"
        };
    }
}
/*
private void Test(string filename)
{
DataTable dt = new DataTable();
 
using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filename, false))
{
 
WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
string relationshipId = sheets.First().Id.Value;
WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
Worksheet workSheet = worksheetPart.Worksheet;
SheetData sheetData = workSheet.GetFirstChild<SheetData>();
IEnumerable<Row> rows = sheetData.Descendants<Row>();
 
foreach (Cell cell in rows.ElementAt(0))
{
dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
}
 
foreach (Row row in rows) //this will also include your header row...
{
DataRow tempRow = dt.NewRow();
 
for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
{
tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
}
 
dt.Rows.Add(tempRow);
}
 
}
dt.Rows.RemoveAt(0); //...so i'm taking it out here.
 
}
public static string GetCellValue(SpreadsheetDocument document, Cell cell)
{
SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
string value = cell.CellValue.InnerXml;
DateTime ReleaseDate = new DateTime(1899, 12, 30);
if (cell.DataType != null)
{
switch (cell.DataType.Value)
{
case CellValues.SharedString:
return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(cell.CellValue.Text)].InnerText;
case CellValues.Boolean:
return cell.CellValue.Text == "1" ? "true" : "false";
case CellValues.Date:
{
DateTime answer = ReleaseDate.AddDays(Convert.ToDouble(cell.CellValue.Text));
return answer.ToShortDateString();
}
case CellValues.Number:
return Convert.ToDecimal(cell.CellValue.Text).ToString();
default:
if (cell.CellValue != null)
return cell.CellValue.Text;
return string.Empty;
}
}
 
if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
{
return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
}
else
{
DateTime answer = ReleaseDate.AddDays(Convert.ToDouble(cell.CellValue.Text));
return answer.ToShortDateString();
}
}
 
--------------------or-----------------------
 
private DataTable ReadExcelFile(string filename)
{
// Initialize an instance of DataTable
DataTable dt = new DataTable();
try
{
// Use SpreadSheetDocument class of Open XML SDK to open excel file
using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, false))
{
// Get Workbook Part of Spread Sheet Document
WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
// Get all sheets in spread sheet document
IEnumerable<Sheet> sheetcollection = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
// Get relationship Id
string relationshipId = sheetcollection.First().Id.Value;
// Get sheet1 Part of Spread Sheet Document
WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
// Get Data in Excel file
SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
IEnumerable<Row> rowcollection = sheetData.Descendants<Row>();
if (rowcollection.Count() == 0)
{
return dt;
}
// Add columns
foreach (Cell cell in rowcollection.ElementAt(0))
{
dt.Columns.Add(GetValueOfCell(spreadsheetDocument, cell));
}
// Add rows into DataTable
foreach (Row row in rowcollection)
{
DataRow temprow = dt.NewRow();
int columnIndex = 0;
foreach (Cell cell in row.Descendants<Cell>())
{
// Get Cell Column Index
int cellColumnIndex = (int)GetColumnIndex(GetColumnName(cell.CellReference));
if (columnIndex < cellColumnIndex)
{
do
{
temprow[columnIndex] = string.Empty;
columnIndex++;
}
while (columnIndex < cellColumnIndex);
}
temprow[columnIndex] = GetValueOfCell(spreadsheetDocument, cell);
columnIndex++;
}
// Add the row to DataTable
// the rows include header row
dt.Rows.Add(temprow);
}
}
// Here remove header row
dt.Rows.RemoveAt(0);
return dt;
}
catch (IOException ex)
{
throw new IOException(ex.Message);
}
}
 
public static string GetColumnName(string cellReference)
{
// Create a regular expression to match the column name portion of the cell name.
Regex regex = new Regex("[A-Za-z]+");
Match match = regex.Match(cellReference);
return match.Value;
}
 
public static int? GetColumnIndex(string columnNameOrCellReference)
{
int columnIndex = 0;
int factor = 1;
for (int pos = columnNameOrCellReference.Length - 1; pos >= 0; pos--) // R to L
{
if (Char.IsLetter(columnNameOrCellReference[pos])) // for letters (columnName)
{
columnIndex += factor * ((columnNameOrCellReference[pos] - 'A') + 1);
factor *= 26;
}
}
return columnIndex;
 
}
/// <summary>
/// Get Value of Cell
/// </summary>
/// <param name="spreadsheetdocument">SpreadSheet Document Object</param>
/// <param name="cell">Cell Object</param>
/// <returns>The Value in Cell</returns>
private static string GetValueOfCell(SpreadsheetDocument spreadsheetdocument, Cell cell)
{
// Get value in Cell
SharedStringTablePart sharedString = spreadsheetdocument.WorkbookPart.SharedStringTablePart;
if (cell.CellValue == null)
{
return string.Empty;
}
string cellValue = cell.CellValue.InnerText;
 
// The condition that the Cell DataType is SharedString
if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
{
return sharedString.SharedStringTable.ChildElements[int.Parse(cellValue)].InnerText;
}
else
{
return cellValue;
}
}*/