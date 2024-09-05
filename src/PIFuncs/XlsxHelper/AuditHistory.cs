using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace XlsxHelper
{
    public class AuditHistory
    {
        public void Append(string xlsxFilename, string[] headers, List<string[]> data, string tagLabel = "Approval History")
        {
            if (xlsxFilename == null || Path.GetExtension(xlsxFilename) != ".xlsx")
                throw new ArgumentNullException("Invalid output filename");

            if (data == null)
                throw new ArgumentNullException("Invalid data specified");

            using (var workbook = new XLWorkbook(xlsxFilename))
            {
                foreach (var sheet in workbook.Worksheets)
                {
                    int index = FindIndex(sheet, tagLabel);

                    if (index != -1)
                    {
                        DeleteRows(sheet, index);
                    }

                    int startingRow = sheet.LastRowUsed().RowNumber() + 3;
                    int startingColumnIndex = 1;
                    int headerRowIndex = startingRow + 1;
                    int startingRowIndex = headerRowIndex + 1;

                    if (headers == null || headers.Length == 0)
                    {
                        headers = new string[] { "Level in Route", "Role/Designation", "Name of the Approver", "Date of Approval" };
                    }

                    AddTitle(sheet, startingRow); //append title

                    for (int head = 0; head < headers.Length; head++) //append headers
                    {
                        CreateHeader(sheet, headerRowIndex, startingColumnIndex + head, headers[head]);
                    }

                    var row = startingRowIndex;
                    var col = startingColumnIndex;

                    foreach (var r in data)
                    {
                        foreach (var v in r)
                        {
                            if (col > headers.Length)
                                break;

                            CreateRow(sheet, row, col, v);
                            col++;
                        }
                        col = startingColumnIndex;
                        row++;
                    }
                }

                workbook.Save();
            }
        }

        private void AddTitle(IXLWorksheet worksheet, int row)
        {
            string Row = row.ToString();
            var title = worksheet.Range($"A{Row}:D{Row}");
            title.Merge();
            title.Value = "Approval History";
            title.Style.Font.Bold = true;
            title.Style.Font.FontSize = 12;
            title.Style.Font.FontName = "Times New Roman";
            title.Style.Fill.BackgroundColor = XLColor.LightGray;
            title.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            title.Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            title.Style.Border.TopBorder = XLBorderStyleValues.Thick;
            title.Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            title.Style.Border.RightBorder = XLBorderStyleValues.Thick;
            title.Style.Border.InsideBorder = XLBorderStyleValues.Thick;
        }

        private void CreateHeader(IXLWorksheet worksheet, int row, int column, string CellValue)
        {
            var cell = worksheet.Cell(row, column);
            cell.Value = CellValue;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 12;
            cell.Style.Font.FontName = "Times New Roman";
            cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thick;
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Thick;
        }

        private void CreateRow(IXLWorksheet worksheet, int row, int column, string CellValue)
        {
            var cell = worksheet.Cell(row, column);
            cell.Value = CellValue;
            cell.Style.Font.FontSize = 12;
            cell.Style.Font.FontName = "Times New Roman";
            cell.Style.Border.BottomBorder = XLBorderStyleValues.Thick;
            cell.Style.Border.TopBorder = XLBorderStyleValues.Thick;
            cell.Style.Border.LeftBorder = XLBorderStyleValues.Thick;
            cell.Style.Border.RightBorder = XLBorderStyleValues.Thick;
        }

        private int FindIndex(IXLWorksheet worksheet, string TextToFind)
        {
            string valueToFind = TextToFind;
            var rows = worksheet.RowsUsed();
            int rowIndex = -1;
            bool found = false;

            foreach (var row in rows)
            {
                foreach (var cell in row.CellsUsed())
                {
                    if (cell.Value.ToString() == valueToFind)
                    {
                        rowIndex = cell.Address.RowNumber;
                        found = true;
                        break;
                    }
                }
                if (found) { break; }
            }

            return rowIndex;
        }

        private void DeleteRows(IXLWorksheet worksheet, int rowIndex)
        {
            int rowToDelete = rowIndex;
            var rows = worksheet.RowsUsed();
            foreach (var row in rows)
            {
                // Check if the row number is greater than or equal to the row to delete
                if (row.RowNumber() >= rowToDelete)
                {
                    row.Delete();
                }
            }
        }
    }
}
