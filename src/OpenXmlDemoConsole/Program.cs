// See https://aka.ms/new-console-template for more information
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Data.Common;

Console.WriteLine("Hello, World!");


//using (var doc = SpreadsheetDocument.Open("ExcelTemplate.xlsx", true))
//{
//    //Read the first Sheets 
//        Sheet sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
//        Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheet.Id.Value) as WorksheetPart).Worksheet;
//        IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
//        int counter = 0;
//        foreach (Row row in rows)
//        {
//            counter = counter + 1;
//            //Read the first row as header
//            if (counter == 1)
//            {
//                var j = 1;
//                foreach (Cell cell in row.Descendants<Cell>())
//                {
//                    var colunmName = firstRowIsHeader ? GetCellValue(doc, cell) : "Field" + j++;
//                    Console.WriteLine(colunmName);
//                    Headers.Add(colunmName);
//                    dt.Columns.Add(colunmName);
//                }
//            }
//            else
//            {
//                dt.Rows.Add();
//                int i = 0;
//                foreach (Cell cell in row.Descendants<Cell>())
//                {
//                    dt.Rows[dt.Rows.Count - 1][i] = GetCellValue(doc, cell);
//                    i++;
//                }
//            }
//        }

//}



