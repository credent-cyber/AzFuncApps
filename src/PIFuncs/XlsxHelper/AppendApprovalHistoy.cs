using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // for .xlsx
using NPOI.HSSF.UserModel; // for .xls
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;

namespace PIFunc.XlsxHelper
{
    public class AppendApprovalHistory
    {
        public void Append(string xlsxFilename, string[] headers, List<string[]> data, string tagLabel = "Approval History")
        {
            if (xlsxFilename == null || (Path.GetExtension(xlsxFilename) != ".xlsx" && Path.GetExtension(xlsxFilename) != ".xls"))
                throw new ArgumentNullException("Invalid output filename (Pass xlsx or xls file)");

            if (data == null)
                throw new ArgumentNullException("Invalid data specified");

            IWorkbook workbook;
            using (var fs = new FileStream(xlsxFilename, FileMode.Open, FileAccess.Read))
            {
                if (Path.GetExtension(xlsxFilename) == ".xlsx")
                {
                    workbook = new XSSFWorkbook(fs); // .xlsx
                }
                else
                {
                    workbook = new HSSFWorkbook(fs); // .xls
                }
            }

            foreach (ISheet sheet in workbook)
            {
                int titleRowIndex = FindIndex(sheet, tagLabel); // Get "Approval History" title row index
                if (titleRowIndex != -1)
                {
                    int startRow = titleRowIndex; // Start deleting from the title row itself
                    int endRow = sheet.LastRowNum; // Set end row to the last row of the table

                    DeleteRows(sheet, startRow); // Delete "Approval History" row and existing data rows in the table range
                    UnmergeApprovalHistoryRange(sheet, startRow, endRow); // Unmerge cells in the range including the title row
                }

                // Continue with your code to add new "Approval History" data
                int startingRow = sheet.LastRowNum + 3;
                int startingColumnIndex = 0;
                int headerRowIndex = startingRow + 0;
                int startingRowIndex = headerRowIndex + 1;

                if (headers == null || headers.Length == 0)
                {
                    headers = new string[] { "Level in Route", "Role/Designation", "Name of the Approver", "Date of Approval" };
                }

                AddTitle(sheet, startingRow); // append title

                for (int head = 0; head < headers.Length; head++) // append headers
                {
                    CreateHeader(sheet, headerRowIndex, startingColumnIndex + head, headers[head]);
                }

                int row = startingRowIndex;

                foreach (var rowData in data)
                {
                    int col = startingColumnIndex;
                    foreach (var cellValue in rowData)
                    {
                        if (col >= headers.Length)
                            break;

                        CreateRow(sheet, row, col, cellValue);
                        col++;
                    }
                    row++;
                }
            }

            // Save the workbook
            using (var fs = new FileStream(xlsxFilename, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        private void AddTitle(ISheet sheet, int row)
        {
            int firstRow = row - 1;
            int lastRow = row - 1;
            int firstCol = 0;
            int lastCol = 3;

            var titleRow = sheet.GetRow(firstRow) ?? sheet.CreateRow(firstRow);

            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var mergedRegion = sheet.GetMergedRegion(i);
                if (mergedRegion.FirstRow == firstRow && mergedRegion.LastRow == lastRow &&
                    mergedRegion.FirstColumn <= lastCol && mergedRegion.LastColumn >= firstCol)
                {
                    sheet.RemoveMergedRegion(i);
                    break;
                }
            }

            var titleRange = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            sheet.AddMergedRegion(titleRange);

            var titleCell = titleRow.GetCell(firstCol) ?? titleRow.CreateCell(firstCol);
            titleCell.SetCellValue("Approval History");

            var titleStyle = sheet.Workbook.CreateCellStyle();
            titleStyle.Alignment = HorizontalAlignment.Center;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;
            titleStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            titleStyle.FillPattern = FillPattern.SolidForeground;

            var titleFont = sheet.Workbook.CreateFont();
            titleFont.IsBold = true;
            titleFont.FontHeightInPoints = 12;
            titleFont.FontName = "Times New Roman";
            titleStyle.SetFont(titleFont);

            titleCell.CellStyle = titleStyle;
        }

        private void CreateHeader(ISheet sheet, int rowIndex, int columnIndex, string cellValue)
        {
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.CreateCell(columnIndex);
            cell.SetCellValue(cellValue);

            ICellStyle style = sheet.Workbook.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.LightBlue.Index;
            style.FillPattern = FillPattern.SolidForeground;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;

            IFont font = sheet.Workbook.CreateFont();
            font.IsBold = true;
            font.FontHeightInPoints = 12;
            font.FontName = "Times New Roman";
            style.SetFont(font);

            cell.CellStyle = style;
        }

        private void CreateRow(ISheet sheet, int rowIndex, int columnIndex, string cellValue)
        {
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            var cell = row.CreateCell(columnIndex);
            cell.SetCellValue(cellValue);

            ICellStyle style = sheet.Workbook.CreateCellStyle();
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;

            IFont font = sheet.Workbook.CreateFont();
            font.FontHeightInPoints = 12;
            font.FontName = "Times New Roman";
            style.SetFont(font);

            cell.CellStyle = style;
        }

        private int FindIndex(ISheet sheet, string textToFind)
        {
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row != null)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        if (cell.ToString() == textToFind)
                        {
                            return rowIndex;
                        }
                    }
                }
            }
            return -1;
        }

        private void DeleteRows(ISheet sheet, int startRow)
        {
            int lastRow = sheet.LastRowNum;
            for (int i = startRow; i <= lastRow; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    sheet.RemoveRow(row);
                }
            }
        }

        private void UnmergeApprovalHistoryRange(ISheet sheet, int startRow, int endRow)
        {
            for (int i = sheet.NumMergedRegions - 1; i >= 0; i--)
            {
                var mergedRegion = sheet.GetMergedRegion(i);

                bool withinTableRange = mergedRegion.FirstRow >= startRow && mergedRegion.LastRow <= endRow;

                if (withinTableRange)
                {
                    sheet.RemoveMergedRegion(i);
                }
            }
        }
    }
}
