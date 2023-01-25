// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using Table = DocumentFormat.OpenXml.Spreadsheet.Table;

Console.WriteLine("Hello, World!");


var fileName = "ExcelTemplate.xlsx";
//var tableName = "Table1";

using (SpreadsheetDocument spreadsheetDocument =
        SpreadsheetDocument.Open(fileName, true))
{

    var worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.FirstOrDefault();
    var sheets = spreadsheetDocument.WorkbookPart.Workbook.Elements<Sheets>().FirstOrDefault();
       
    var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == "Sheet1")
        .FirstOrDefault();

    var colMin = 1;
    var colMax = 8;
    var rowMin = 16;
    var rowMax = 18;


    TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>("rId" + (worksheetPart.TableDefinitionParts.Count() + 50));
    int tableNo = worksheetPart.TableDefinitionParts.Count();

    string reference = ((char)(64 + colMin)).ToString() + rowMin + ":" + ((char)(64 + colMax)).ToString() + rowMax;

    Table table = new Table() { Id = (UInt32)tableNo, Name = "Table" + tableNo, DisplayName = "Table" + tableNo, Reference = reference, TotalsRowShown = false };
    AutoFilter autoFilter = new AutoFilter() { Reference = reference };

    TableColumns tableColumns = new TableColumns() { Count = (UInt32)(colMax - colMin + 1) };
    for (int i = 0; i < (colMax - colMin + 1); i++)
    {
        tableColumns.Append(new TableColumn() { Id = (UInt32)(colMin + i), Name = "Column" + i }); //changed i+1 -> colMin + i
                                                                                                   //Add cell values (shared string)
    }

    TableStyleInfo tableStyleInfo = new TableStyleInfo() { Name = "TableStyleLight1", ShowFirstColumn = false, ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false };

    table.Append(autoFilter);
    table.Append(tableColumns);
    table.Append(tableStyleInfo);

    tableDefinitionPart.Table = table;

    TableParts tableParts = (TableParts)worksheetPart.Worksheet.ChildElements.Where(ce => ce is TableParts).FirstOrDefault(); // Add table parts only once
    if (tableParts is null)
    {
        tableParts = new TableParts();
        tableParts.Count = (UInt32)0;
        worksheetPart.Worksheet.Append(tableParts);
    }

    spreadsheetDocument.Save();
    Console.WriteLine("Press any key to continue..");

    Console.ReadKey();
    
}



//DataTable dtTable = new DataTable();
//List<string> rowList = new List<string>();
//ISheet sheet;

//XSSFWorkbook workbook;

//using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
//{
//    workbook = new XSSFWorkbook(stream);
//    stream.Close();
//    stream.Dispose();
//}

//using (var ofile = new FileStream(filename, FileMode.Open, FileAccess.Write))
//{
//    sheet = workbook.GetSheetAt(0);

//    var tbl = workbook.GetTable("Table1");

//    // if table exists
//    // remove all row exclude all row until last row of the sheet
//    // add new data rows from the 2nd row of the table
//    if (tbl != null)
//    {
//        for (var rowIndex = tbl.StartCellReference.Row + 1; rowIndex < tbl.EndCellReference.Row + 1; rowIndex++)
//        {
//            var row = sheet.GetRow(rowIndex);
//            sheet.RemoveRow(row);
//        }
//    }

//    workbook.Write(ofile);

//}



////    // if table not exits
////    // find the last row of the sheet
////    // create a heading row "Approval History"
////    // creeate column heading row
////    // create data rows

////    IRow headerRow = sheet.GetRow(0);

////    int cellCount = headerRow.LastCellNum;

////    for (int j = 0; j < cellCount; j++)
////    {
////        ICell cell = headerRow.GetCell(j);

////        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
////        {
////            dtTable.Columns.Add(cell.ToString());
////        }
////    }

////    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
////    {
////        IRow row = sheet.GetRow(i);

////        if (row == null) continue;

////        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

////        for (int j = row.FirstCellNum; j < cellCount; j++)
////        {
////            if (row.GetCell(j) != null)
////            {
////                if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
////                {
////                    rowList.Add(row.GetCell(j).ToString());
////                }
////            }
////        }

////        if (rowList.Count > 0)
////            dtTable.Rows.Add(rowList.ToArray());

////        rowList.Clear();
////    }
////}

////List<UserDetails> persons = new List<UserDetails>()
////            {
////                new UserDetails() {ID="1001", Name="ABCD", City ="City1", Country="USA"},
////                new UserDetails() {ID="1002", Name="PQRS", City ="City2", Country="INDIA"},
////                new UserDetails() {ID="1003", Name="XYZZ", City ="City3", Country="CHINA"},
////                new UserDetails() {ID="1004", Name="LMNO", City ="City4", Country="UK"},
////           };

////// Lets converts our object data to Datatable for a simplified logic.
////// Datatable is most easy way to deal with complex datatypes for easy reading and formatting.

////DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(persons), (typeof(DataTable)));
////var memoryStream = new MemoryStream();

////using (var fs = new FileStream("Book12.xlsx", FileMode.Create, FileAccess.Write))
////{
////    IWorkbook workbook = new XSSFWorkbook();
////    ISheet excelSheet = workbook.CreateSheet("Sheet1");

////    List<String> columns = new List<string>();
////    IRow row = excelSheet.CreateRow(0);
////    int columnIndex = 0;

////    foreach (System.Data.DataColumn column in table.Columns)
////    {
////        columns.Add(column.ColumnName);
////        row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
////        columnIndex++;
////    }

////    int rowIndex = 1;
////    foreach (DataRow dsrow in table.Rows)
////    {
////        row = excelSheet.CreateRow(rowIndex);
////        int cellIndex = 0;
////        foreach (String col in columns)
////        {
////            row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
////            cellIndex++;
////        }

////        rowIndex++;
////    }
////    workbook.Write(fs);
////}




//class UserDetails
//{
//    public string ID { get; set; }
//    public string Name { get; set; }
//    public string City { get; set; }
//    public string Country { get; set; }
//}